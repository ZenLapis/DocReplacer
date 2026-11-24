"""Combine every file in a folder into one PDF, fast.

Office files are converted by several hidden Word/Excel instances in
parallel (PowerPoint is single-instance, so it runs on its own thread),
images and existing PDFs are added directly, and everything is merged with
PyMuPDF. Each source file gets a bookmark in the final PDF.
"""

import atexit
import hashlib
import io
import queue
import os
import re
import shutil
import tempfile
import threading
import time
import unicodedata
import zipfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import fitz  # PyMuPDF

WORD_EXTS = {".docx", ".doc", ".docm", ".dotx", ".dot", ".rtf", ".odt", ".txt", ".xml"}
EXCEL_EXTS = {".xlsx", ".xls", ".xlsm", ".xlsb", ".ods", ".csv"}
PPT_EXTS = {".pptx", ".ppt", ".pptm", ".ppsx", ".pps", ".odp"}
IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff", ".webp"}
PDF_EXTS = {".pdf"}
SUPPORTED_EXTS = WORD_EXTS | EXCEL_EXTS | PPT_EXTS | IMAGE_EXTS | PDF_EXTS

A4 = fitz.paper_rect("a4")
MAX_WORD_WORKERS = 6
MAX_EXCEL_WORKERS = 2


def _natural_key(text):
    text = unicodedata.normalize("NFC", text).casefold()
    return [int(part) if part.isdigit() else part for part in re.split(r"(\d+)", text)]


def collect_files(folder):
    """Return (supported, skipped) file lists in natural path order."""
    folder = Path(folder)
    supported, skipped = [], []
    for root_dir, dirs, files in os.walk(folder):
        dirs[:] = [d for d in dirs if not d.startswith(("_backup_", ".replacerx-"))]
        for filename in files:
            if filename.startswith("~$") or filename.lower() in ("desktop.ini", "thumbs.db"):
                continue
            path = Path(root_dir) / filename
            (supported if path.suffix.lower() in SUPPORTED_EXTS else skipped).append(path)
    key = lambda p: [_natural_key(part) for part in p.relative_to(folder).parts]
    supported.sort(key=key)
    skipped.sort(key=key)
    return supported, skipped


def next_pdf_path(folder):
    folder = Path(folder).resolve()
    candidate = folder.parent / f"{folder.name}.pdf"
    counter = 2
    while candidate.exists():
        candidate = folder.parent / f"{folder.name} ({counter}).pdf"
        counter += 1
    return candidate


# --------------------------------------------------------------------------
# Office conversion (COM)
# --------------------------------------------------------------------------

class _OfficeWorker:
    """One hidden Office instance living on one thread."""

    def __init__(self, kind):
        self.kind = kind
        self.app = None

    def _start(self):
        import win32com.client

        progid = {
            "word": "Word.Application",
            "excel": "Excel.Application",
            "powerpoint": "PowerPoint.Application",
        }[self.kind]
        self.app = win32com.client.DispatchEx(progid)
        if self.kind == "word":
            self.app.Visible = False
            self.app.DisplayAlerts = 0
            self.app.ScreenUpdating = False
            try:
                self.app.Options.UpdateLinksAtOpen = False
                self.app.Options.CheckGrammarAsYouType = False
                self.app.Options.CheckSpellingAsYouType = False
            except Exception:
                pass
        elif self.kind == "excel":
            self.app.Visible = False
            self.app.DisplayAlerts = False
            self.app.ScreenUpdating = False
            self.app.AskToUpdateLinks = False
        try:
            self.app.AutomationSecurity = 3  # disable macros
        except Exception:
            pass

    def _alive(self):
        try:
            _ = self.app.Name
            return True
        except Exception:
            return False

    def convert(self, src, dst):
        if self.app is None or not self._alive():
            self.quit()
            self._start()
        src, dst = str(src), str(dst)
        if self.kind == "word":
            doc = self.app.Documents.Open(
                src, ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False,
                PasswordDocument="", Visible=False, NoEncodingDialog=True,
            )
            try:
                # 17 = wdExportFormatPDF, 0 = print quality, 0 = whole doc
                doc.ExportAsFixedFormat(
                    dst, 17, OpenAfterExport=False, OptimizeFor=0, Range=0,
                    Item=0, IncludeDocProps=False, KeepIRM=True,
                    CreateBookmarks=0, DocStructureTags=False,
                    BitmapMissingFonts=True, UseISO19005_1=False,
                )
            finally:
                doc.Close(0)
        elif self.kind == "excel":
            wb = self.app.Workbooks.Open(src, UpdateLinks=0, ReadOnly=True, Password="")
            try:
                wb.ExportAsFixedFormat(0, dst, 0, False, False)  # whole workbook
            finally:
                wb.Close(False)
        else:
            pres = self.app.Presentations.Open(src, True, False, False)  # RO, untitled, no window
            try:
                pres.SaveAs(dst, 32)  # ppSaveAsPDF
            finally:
                pres.Close()

    def quit(self):
        if self.app is not None:
            try:
                self.app.Quit()
            except Exception:
                pass
            self.app = None


IDLE_SECONDS = 600
CACHE_DIR = Path(os.getenv("LOCALAPPDATA", tempfile.gettempdir())) / "DocReplacer" / "pdf-cache"
CACHE_LIMIT = 800 * 1024 * 1024


class _OfficePool:
    """Long-lived Office instances for one app kind, shared by every PDF job.

    Starting Word/Excel/PowerPoint costs seconds; keeping warm instances
    between jobs (and between hotels in a batch) removes that cost. Idle
    instances quit after IDLE_SECONDS, and all of them quit at exit.
    """

    def __init__(self, kind, size):
        self.kind, self.size = kind, size
        self.jobs = queue.Queue()
        self.lock = threading.Lock()
        self.threads = []

    def submit(self, jobs, n_workers, on_done):
        remaining = [len(jobs)]
        finished = threading.Event()
        count_lock = threading.Lock()

        def done(index, error):
            try:
                on_done(index, error)
            finally:
                with count_lock:
                    remaining[0] -= 1
                    if remaining[0] == 0:
                        finished.set()

        with self.lock:
            for index, src, dst in jobs:
                self.jobs.put((index, src, dst, done))
            self.threads = [t for t in self.threads if t.is_alive()]
            while len(self.threads) < min(n_workers, self.size):
                thread = threading.Thread(target=self._worker, daemon=True)
                thread.start()
                self.threads.append(thread)
        finished.wait()

    def _worker(self):
        import pythoncom

        pythoncom.CoInitialize()
        office = _OfficeWorker(self.kind)
        try:
            try:
                office._start()  # warm up before the first job arrives
            except Exception:
                office.quit()
            while True:
                try:
                    job = self.jobs.get(timeout=IDLE_SECONDS)
                except queue.Empty:
                    with self.lock:
                        if self.jobs.empty():
                            self.threads = [t for t in self.threads if t is not threading.current_thread()]
                            return
                    continue
                if job is None:
                    return
                index, src, dst, done = job
                try:
                    office.convert(src, dst)
                    error = None
                except Exception as exc:
                    error = _com_error_text(exc)
                    office.quit()  # a failed document can leave Office unusable
                done(index, error)
        finally:
            office.quit()
            pythoncom.CoUninitialize()

    def warm(self, n_workers):
        with self.lock:
            self.threads = [t for t in self.threads if t.is_alive()]
            while len(self.threads) < min(n_workers, self.size):
                thread = threading.Thread(target=self._worker, daemon=True)
                thread.start()
                self.threads.append(thread)

    def shutdown(self):
        with self.lock:
            threads = [t for t in self.threads if t.is_alive()]
            for _ in threads:
                self.jobs.put(None)
        for thread in threads:
            thread.join(timeout=15)


_POOLS = {}
_POOLS_LOCK = threading.Lock()


def _pool(kind):
    sizes = {"word": MAX_WORD_WORKERS, "excel": MAX_EXCEL_WORKERS, "powerpoint": 1}
    with _POOLS_LOCK:
        if kind not in _POOLS:
            _POOLS[kind] = _OfficePool(kind, sizes[kind])
        return _POOLS[kind]


@atexit.register
def shutdown_office():
    for pool in list(_POOLS.values()):
        pool.shutdown()


def prewarm_office():
    """Start hidden Word instances in the background so the next PDF starts instantly."""
    try:
        _pool("word").warm(max(1, min(MAX_WORD_WORKERS, (os.cpu_count() or 4) // 2)))
    except Exception:
        pass


def _run_office_pool(kind, jobs, n_workers, on_done):
    """Convert jobs [(index, src, dst)] on the shared warm Office pool."""
    if jobs:
        _pool(kind).submit(jobs, n_workers, on_done)


def _cache_key(path):
    digest = hashlib.sha1()
    digest.update(unicodedata.normalize("NFC", path.name).casefold().encode())
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _cache_store(key, pdf):
    try:
        CACHE_DIR.mkdir(parents=True, exist_ok=True)
        tmp = CACHE_DIR / f"{key}.{threading.get_ident()}.tmp"
        shutil.copyfile(pdf, tmp)
        os.replace(tmp, CACHE_DIR / f"{key}.pdf")
    except OSError:
        pass


def _cache_prune():
    try:
        entries = sorted(CACHE_DIR.glob("*.pdf"), key=lambda p: p.stat().st_atime)
        total = sum(p.stat().st_size for p in entries)
        while entries and total > CACHE_LIMIT:
            victim = entries.pop(0)
            total -= victim.stat().st_size
            victim.unlink(missing_ok=True)
    except OSError:
        pass


def _com_error_text(exc):
    args = getattr(exc, "args", ())
    if len(args) >= 3 and isinstance(args[2], tuple) and len(args[2]) >= 3 and args[2][2]:
        return str(args[2][2]).strip()
    return str(exc)


# --------------------------------------------------------------------------
# Merge
# --------------------------------------------------------------------------

def _image_page(out, path):
    from PIL import Image, ImageOps

    with Image.open(path) as img:
        img.seek(0)
        img = ImageOps.exif_transpose(img)
        if img.mode not in ("RGB", "L"):
            img = img.convert("RGB")
        width, height = img.size
        buf = io.BytesIO()
        if path.suffix.lower() in (".jpg", ".jpeg") and img.mode in ("RGB", "L"):
            img.save(buf, "JPEG", quality=92)
        else:
            img.save(buf, "PNG")
    page_rect = A4 if height >= width else fitz.Rect(0, 0, A4.height, A4.width)
    page = out.new_page(width=page_rect.width, height=page_rect.height)
    margin = 18
    page.insert_image(
        fitz.Rect(margin, margin, page_rect.width - margin, page_rect.height - margin),
        stream=buf.getvalue(), keep_proportion=True,
    )


def combine_folder_to_pdf(folder, output=None, log=None, progress=None):
    """Combine every supported file under *folder* into one PDF.

    log(text, tag) and progress(text) are optional callbacks.
    Returns dict(output, added, failed, skipped, seconds).
    """
    log = log or (lambda text, tag="info": None)
    progress = progress or (lambda text: None)
    started = time.perf_counter()
    folder = Path(folder).resolve()
    if not folder.is_dir():
        raise NotADirectoryError(f"Klasör bulunamadı: {folder}")

    files, skipped = collect_files(folder)
    if not files:
        raise FileNotFoundError("Klasörde PDF'e çevrilebilecek dosya bulunamadı.")
    for path in skipped:
        log(f"– Atlandı (desteklenmiyor): {path.relative_to(folder)}", "warn")

    output = Path(output) if output else next_pdf_path(folder)
    work_dir = Path(tempfile.mkdtemp(prefix="replacerx-pdf-"))
    total = len(files)
    errors = {}
    done_count = [0]
    count_lock = threading.Lock()

    def on_done(index, error):
        if error:
            errors[index] = error
        elif cache_keys.get(index):
            _cache_store(cache_keys[index], converted[index])
        with count_lock:
            done_count[0] += 1
            current = done_count[0]
        try:
            progress(f"PDF'e çevriliyor: {current}/{office_total}")
        except Exception:
            pass  # UI updates must never break conversion

    jobs = {"word": [], "excel": [], "powerpoint": []}
    converted, cache_keys = {}, {}
    for index, path in enumerate(files):
        ext = path.suffix.lower()
        kind = "word" if ext in WORD_EXTS else "excel" if ext in EXCEL_EXTS else "powerpoint" if ext in PPT_EXTS else None
        if kind:
            try:
                key = _cache_key(path)
            except OSError:
                key = None
            cached = CACHE_DIR / f"{key}.pdf" if key else None
            if cached and cached.is_file():
                converted[index] = cached
                try:
                    os.utime(cached)
                except OSError:
                    pass
            else:
                dst = work_dir / f"{index:05d}.pdf"
                converted[index] = dst
                cache_keys[index] = key
                jobs[kind].append((index, path, dst))
    office_total = sum(len(v) for v in jobs.values())
    reused = len(converted) - office_total

    try:
        if office_total or reused:
            cpu = os.cpu_count() or 4
            log(
                f"{total} dosya bulundu ({office_total} Office belgesi paralel çevriliyor"
                + (f", {reused} önbellekten" if reused else "") + ")...",
                "info",
            )
            with ThreadPoolExecutor(max_workers=3) as pool:
                futures = [
                    pool.submit(_run_office_pool, "word", jobs["word"],
                                max(1, min(MAX_WORD_WORKERS, cpu // 2, len(jobs["word"]))), on_done),
                    pool.submit(_run_office_pool, "excel", jobs["excel"],
                                max(1, min(MAX_EXCEL_WORKERS, len(jobs["excel"]))), on_done),
                    pool.submit(_run_office_pool, "powerpoint", jobs["powerpoint"], 1, on_done),
                ]
                for future in futures:
                    future.result()
        else:
            log(f"{total} dosya bulundu.", "info")

        progress("PDF birleştiriliyor...")
        out = fitz.open()
        toc = []
        added = 0
        for index, path in enumerate(files):
            rel = str(path.relative_to(folder))
            if index in errors:
                log(f"✗ {rel}: {errors[index]}", "err")
                continue
            ext = path.suffix.lower()
            start_page = out.page_count
            try:
                if ext in IMAGE_EXTS:
                    _image_page(out, path)
                else:
                    src_pdf = path if ext in PDF_EXTS else converted[index]
                    with fitz.open(src_pdf) as part:
                        if part.needs_pass and not part.authenticate(""):
                            raise ValueError("şifreli PDF")
                        out.insert_pdf(part)
            except Exception as exc:
                errors[index] = str(exc)
                log(f"✗ {rel}: {exc}", "err")
                continue
            if out.page_count > start_page:
                toc.append([1, path.stem, start_page + 1])
                added += 1

        if added == 0:
            raise RuntimeError("Hiçbir dosya PDF'e eklenemedi.")
        out.set_toc(toc)
        progress("PDF kaydediliyor...")
        tmp_output = output.with_name(output.name + ".part")
        out.save(str(tmp_output), garbage=3, deflate=True)
        page_count = out.page_count
        out.close()
        os.replace(tmp_output, output)
    finally:
        shutil.rmtree(work_dir, ignore_errors=True)
        _cache_prune()

    seconds = time.perf_counter() - started
    return {
        "output": output,
        "added": added,
        "failed": len(errors),
        "skipped": len(skipped),
        "pages": page_count,
        "seconds": seconds,
    }


def combine_to_pdf(source, output=None, log=None, progress=None):
    """Like combine_folder_to_pdf, but *source* may also be a .zip file.

    A ZIP is extracted to a temp folder; the PDF is written next to the ZIP.
    """
    source = Path(source).resolve()
    if source.is_dir():
        return combine_folder_to_pdf(source, output, log, progress)
    if source.suffix.lower() != ".zip" or not source.is_file():
        raise ValueError(f"Klasör veya .zip bekleniyordu: {source}")
    if output is None:
        output = source.with_suffix(".pdf")
        counter = 2
        while output.exists():
            output = source.with_name(f"{source.stem} ({counter}).pdf")
            counter += 1
    extract_dir = Path(tempfile.mkdtemp(prefix="replacerx-zip-"))
    try:
        if progress:
            progress("ZIP açılıyor...")
        with zipfile.ZipFile(source) as archive:
            for member in archive.infolist():
                target = (extract_dir / member.filename).resolve()
                if not target.is_relative_to(extract_dir.resolve()):
                    raise ValueError(f"ZIP içinde güvensiz yol: {member.filename}")
            archive.extractall(extract_dir)
        entries = [p for p in extract_dir.iterdir() if not p.name.startswith("__MACOSX")]
        folder = entries[0] if len(entries) == 1 and entries[0].is_dir() else extract_dir
        return combine_folder_to_pdf(folder, output, log, progress)
    finally:
        shutil.rmtree(extract_dir, ignore_errors=True)
