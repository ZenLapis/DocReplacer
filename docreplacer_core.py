"""DocReplacer engine: FR.A4.1 detection, safe copy, name/logo replacement, ZIP."""
import json
import os
import re
import shutil
import sys
import unicodedata
import uuid
import zipfile
from pathlib import Path

from docx import Document

APP_NAME = "DocReplacer"
APP_VERSION = "2.0"
SETTINGS_DIR = Path(os.getenv("LOCALAPPDATA", Path.home())) / APP_NAME
SETTINGS_FILE = SETTINGS_DIR / "settings.json"
LEGACY_SETTINGS_FILE = Path(os.getenv("LOCALAPPDATA", Path.home())) / "DocReplace" / "settings.json"
IMAGE_SUFFIXES = {".png", ".jpg", ".jpeg", ".bmp"}

FR_A41_RE = re.compile(r"^FR\.A4\.1_", re.IGNORECASE)
NAME_PATTERNS = (
    re.compile(
        r"^\s*1\s*[.)-]?\s*(.+?)(?:[’'])?(?:deki|daki|teki|taki)\s+"
        r"Sürdürülebilirlik\b",
        re.IGNORECASE,
    ),
    re.compile(
        r"^\s*1\s*[.)-]?\s*(.+?)(?:[’'])?(?:deki|daki|teki|taki)\b",
        re.IGNORECASE,
    ),
)
INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def load_app_settings():
    for path in (SETTINGS_FILE, LEGACY_SETTINGS_FILE):
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return data
        except (OSError, ValueError, TypeError):
            continue
    return {}


def save_app_settings(data):
    temp = SETTINGS_FILE.with_suffix(".tmp")
    try:
        SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
        temp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(temp, SETTINGS_FILE)
    except OSError:
        try:
            temp.unlink(missing_ok=True)
        except OSError:
            pass


def normalized(text):
    return unicodedata.normalize("NFC", text)


def same_name(left, right):
    return normalized(left).strip().casefold() == normalized(right).strip().casefold()


def resolve_dropped_path(value):
    path = Path(value).expanduser().resolve()
    if path.is_file():
        return path.parent
    return path


def find_fr_a41(folder):
    folder = Path(folder)
    candidates = []
    for root_dir, dirs, files in os.walk(folder):
        dirs[:] = [d for d in dirs if not d.startswith("_backup_")]
        root_path = Path(root_dir)
        for filename in files:
            display_name = normalized(filename)
            if (
                FR_A41_RE.match(display_name)
                and display_name.lower().endswith(".docx")
                and not filename.startswith("~$")
            ):
                candidate = root_path / filename
                depth = len(candidate.relative_to(folder).parts)
                candidates.append((depth, normalized(str(candidate)).casefold(), candidate))
    if not candidates:
        raise FileNotFoundError("FR.A4.1_ ile başlayan .docx dosyası bulunamadı.")
    candidates.sort(key=lambda item: (item[0], item[1]))
    return candidates[0][2]


def iter_table_paragraphs(tables):
    for table in tables:
        for row in table.rows:
            for cell in row.cells:
                yield from cell.paragraphs
                yield from iter_table_paragraphs(cell.tables)


def extract_business_name(fr_path):
    doc = Document(str(fr_path))
    paragraphs = list(doc.paragraphs)
    paragraphs.extend(iter_table_paragraphs(doc.tables))
    for paragraph in paragraphs:
        text = normalized(" ".join(paragraph.text.split()))
        for pattern in NAME_PATTERNS:
            match = pattern.search(text)
            if match:
                name = match.group(1).strip(" -–—:;,.\t")
                if name:
                    return name
    raise ValueError(
        "FR.A4.1 bulundu, ancak '1. [İşletme adı]’deki ...' satırından işletme adı okunamadı."
    )


def inspect_folder(folder):
    folder = Path(folder).resolve()
    if not folder.is_dir():
        raise NotADirectoryError(f"Klasör bulunamadı: {folder}")
    fr_path = find_fr_a41(folder)
    business_name = extract_business_name(fr_path)
    target = folder_rename_target(folder, business_name)
    return {
        "folder": folder,
        "fr_path": fr_path,
        "business_name": business_name,
        "target": target,
        "matches": target is not None and target == folder,
    }


def clean_folder_component(name):
    cleaned = INVALID_FILENAME_CHARS.sub("-", normalized(name)).strip().rstrip(".")
    if not cleaned:
        raise ValueError("FR.A4.1 içindeki işletme adı geçerli bir klasör adı oluşturmuyor.")
    return cleaned


def folder_rename_target(folder, business_name):
    folder = Path(folder).resolve()
    parts = folder.name.rsplit(" - ", 2)
    if len(parts) != 3 or not re.fullmatch(r"\d{4}", parts[2].strip()):
        return None
    desired_name = f"{parts[0]} - {clean_folder_component(business_name)} - {parts[2]}"
    if same_name(folder.name, desired_name):
        return folder
    return folder.with_name(desired_name)


def sync_folder_name(folder):
    details = inspect_folder(folder)
    target = details["target"]
    if target is None:
        raise ValueError(
            "Klasör adı '[başlık] - [işletme] - [4 haneli yıl]' biçiminde değil; güvenli şekilde ayrıştırılamadı."
        )
    if target == details["folder"]:
        return details, False
    if target.exists():
        raise FileExistsError(f"Hedef klasör zaten var: {target}")
    details["folder"].rename(target)
    details["target"] = target
    return details, True


def case_replace(text, old_name, new_name):
    def replacement(match):
        return matched_replacement(match.group(0), new_name)

    return re.sub(re.escape(old_name), replacement, text, flags=re.IGNORECASE)


def matched_replacement(found, new_name):
    if found.isupper():
        return new_name.upper()
    if found.islower():
        return new_name.lower()
    return new_name


def replace_paragraph(paragraph, old_name, new_name):
    if not old_name or not paragraph.runs:
        return False
    full_text = "".join(run.text for run in paragraph.runs)
    matches = list(re.finditer(re.escape(old_name), full_text, re.IGNORECASE))
    if not matches:
        return False

    spans = []
    cursor = 0
    for index, run in enumerate(paragraph.runs):
        end = cursor + len(run.text)
        spans.append((cursor, end, index))
        cursor = end

    def locate(position):
        for start, end, index in spans:
            if start <= position < end:
                return index, position - start
        raise IndexError(f"Metin konumu bir Word çalışmasına eşlenemedi: {position}")

    # Work right-to-left so replacing a longer/shorter name never invalidates
    # the offsets of matches that still need to be processed.
    for match in reversed(matches):
        start_run, start_offset = locate(match.start())
        end_run, end_offset = locate(match.end() - 1)
        end_offset += 1
        replacement = matched_replacement(match.group(0), new_name)
        if start_run == end_run:
            text = paragraph.runs[start_run].text
            paragraph.runs[start_run].text = text[:start_offset] + replacement + text[end_offset:]
            continue
        first_text = paragraph.runs[start_run].text
        last_text = paragraph.runs[end_run].text
        paragraph.runs[start_run].text = first_text[:start_offset] + replacement
        for index in range(start_run + 1, end_run):
            paragraph.runs[index].text = ""
        paragraph.runs[end_run].text = last_text[end_offset:]
    return True


def replace_tables(tables, old_name, new_name):
    changed = False
    for table in tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    changed = replace_paragraph(paragraph, old_name, new_name) or changed
                changed = replace_tables(cell.tables, old_name, new_name) or changed
    return changed


def replace_docx(file_path, old_name, new_name, logo_path=None):
    doc = Document(str(file_path))
    name_changed = False
    logo_changed = False

    for paragraph in doc.paragraphs:
        name_changed = replace_paragraph(paragraph, old_name, new_name) or name_changed
    name_changed = replace_tables(doc.tables, old_name, new_name) or name_changed

    seen_parts = set()
    for section in doc.sections:
        for container in (section.header, section.footer):
            for paragraph in container.paragraphs:
                name_changed = replace_paragraph(paragraph, old_name, new_name) or name_changed
            name_changed = replace_tables(container.tables, old_name, new_name) or name_changed

        if logo_path:
            header_part = section.header.part
            if id(header_part) in seen_parts:
                continue
            seen_parts.add(id(header_part))
            for relationship in header_part.rels.values():
                if "image" in relationship.reltype:
                    with open(logo_path, "rb") as logo_file:
                        relationship.target_part._blob = logo_file.read()
                    logo_changed = True

    if name_changed or logo_changed:
        doc.save(str(file_path))

    new_path = Path(file_path)
    new_filename = case_replace(new_path.name, old_name, new_name)
    renamed = new_filename != new_path.name
    if renamed:
        target = new_path.with_name(new_filename)
        if target.exists():
            raise FileExistsError(f"Dosya yeniden adlandırılamadı; hedef zaten var: {target}")
        new_path.rename(target)
        new_path = target

    return {
        "name": name_changed,
        "logo": logo_changed,
        "renamed": renamed,
        "path": new_path,
    }


def collect_docx(folder):
    paths = []
    for root_dir, dirs, files in os.walk(folder):
        dirs[:] = [d for d in dirs if not d.startswith("_backup_")]
        for filename in files:
            if filename.lower().endswith(".docx") and not filename.startswith("~$"):
                paths.append(Path(root_dir) / filename)
    paths.sort(key=lambda path: normalized(str(path.relative_to(folder))).casefold())
    return paths


def next_zip_path(folder):
    folder = Path(folder).resolve()
    candidate = folder.parent / f"{folder.name}.zip"
    counter = 2
    while candidate.exists():
        candidate = folder.parent / f"{folder.name} ({counter}).zip"
        counter += 1
    return candidate


def create_finished_zip(folder):
    folder = Path(folder).resolve()
    if not folder.is_dir():
        raise NotADirectoryError(f"ZIP oluşturulacak klasör bulunamadı: {folder}")
    zip_path = next_zip_path(folder)
    try:
        with zipfile.ZipFile(
            zip_path, mode="x", compression=zipfile.ZIP_DEFLATED, compresslevel=6
        ) as archive:
            archive.write(folder, folder.name)
            for root_dir, dirs, files in os.walk(folder):
                root_path = Path(root_dir)
                for directory in dirs:
                    directory_path = root_path / directory
                    archive.write(directory_path, directory_path.relative_to(folder.parent))
                for filename in files:
                    file_path = root_path / filename
                    archive.write(file_path, file_path.relative_to(folder.parent))
        return zip_path
    except Exception:
        if zip_path.exists():
            zip_path.unlink()
        raise


def copy_target_for(source_folder, new_name):
    source_folder = Path(source_folder).resolve()
    target = folder_rename_target(source_folder, new_name)
    if target is None:
        raise ValueError(
            "Kaynak klasör adı '[başlık] - [işletme] - [4 haneli yıl]' biçiminde değil."
        )
    if target == source_folder:
        raise ValueError("Yeni isim kaynak klasördeki işletme adıyla aynı.")
    if target.exists():
        raise FileExistsError(f"Yeni kopyanın hedef klasörü zaten var: {target}")
    return target


def make_source_copy(source_folder, new_name):
    source_folder = Path(source_folder).resolve()
    if not collect_docx(source_folder):
        raise FileNotFoundError("Kaynak klasörde hazırlanabilecek .docx belgesi bulunamadı.")
    target = copy_target_for(source_folder, new_name)
    work_folder = source_folder.parent / f".docflow-copy-{uuid.uuid4().hex}"
    try:
        shutil.copytree(
            source_folder,
            work_folder,
            ignore=shutil.ignore_patterns(
                "_backup_*", ".replacerx-*", ".docflow-*", "__pycache__", "~$*"
            ),
        )
        work_folder.rename(target)
        return target
    except Exception:
        if work_folder.exists():
            shutil.rmtree(work_folder)
        raise


def prepare_hotel_copy(source_folder, old_name, new_name, logo=None, log=None, progress=None):
    """Copy source -> new hotel folder, replace name/logo in every .docx, verify, ZIP.

    Returns dict(ok, folder, zip, changed, unchanged, failed).
    """
    log = log or (lambda text, tag="info": None)
    progress = progress or (lambda text: None)
    source_folder = Path(source_folder).resolve()
    progress("Kaynak klasör kopyalanıyor...")
    finished_folder = make_source_copy(source_folder, new_name)
    log(f"Kaynak korundu: {source_folder.name}", "info")
    log(f"Yeni kopya: {finished_folder.name}\n", "ok")

    files = collect_docx(finished_folder)
    if not files:
        raise FileNotFoundError("Klasörde .docx dosyası bulunamadı.")
    log(f"{len(files)} Word belgesi bulundu. İşlem başlıyor...\n", "info")
    changed = unchanged = failed = 0
    for index, file_path in enumerate(files, 1):
        progress(f"İşleniyor: {index}/{len(files)}")
        try:
            result = replace_docx(file_path, old_name, new_name, logo)
            if result["name"] or result["logo"] or result["renamed"]:
                changed += 1
                log(f"✓ {file_path.name}", "ok")
            else:
                unchanged += 1
        except Exception as exc:
            failed += 1
            log(f"✗ {file_path.name}: {exc}", "err")

    if failed == 0:
        try:
            details = inspect_folder(finished_folder)
            if not same_name(details["business_name"], new_name):
                raise ValueError(
                    f"FR.A4.1 hâlâ '{details['business_name']}' adını içeriyor; "
                    f"beklenen ad '{new_name}'."
                )
            log("\nFR.A4.1 yeni işletme adıyla doğrulandı.", "ok")
        except Exception as exc:
            failed += 1
            log(f"\nYeni kopya doğrulanamadı: {exc}", "err")

    zip_path = None
    if failed == 0:
        progress("ZIP oluşturuluyor...")
        try:
            zip_path = create_finished_zip(finished_folder)
            log(f"\nZIP oluşturuldu:\n{zip_path.name}", "ok")
        except Exception as exc:
            log(f"\nZIP oluşturulamadı: {exc}", "warn")
    else:
        log("\nHatalı belge bulunduğu için tamamlanmış ZIP oluşturulmadı.", "warn")

    log(f"\nTamamlandı: {changed} değişti, {unchanged} değişmedi, {failed} hata.", "summary")
    return {
        "ok": failed == 0,
        "folder": finished_folder,
        "zip": zip_path,
        "changed": changed,
        "unchanged": unchanged,
        "failed": failed,
    }


def match_key(text):
    return re.sub(r"[\W_]+", "", normalized(text).casefold())


def name_from_logo(path):
    name = re.sub(r"[_-]+", " ", Path(path).stem)
    return " ".join(name.split()).strip()
