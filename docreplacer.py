"""DocReplacer desktop app (WebView2 UI on top of docreplacer_core)."""
import base64
import ctypes
import io
import json
import os
import subprocess
import sys
import threading
import time
import uuid
from pathlib import Path

import webview
from webview.dom import DOMEventHandler
from PIL import Image, ImageOps

import docreplacer_core as core
from pdf_combine import prewarm_office, combine_to_pdf

RESOURCE_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
UI_FILE = RESOURCE_DIR / "ui" / "docreplacer.html"
IMAGE_TYPES = ("Görseller (*.png;*.jpg;*.jpeg;*.bmp)",)


def load_settings():
    for path in (core.SETTINGS_FILE, core.LEGACY_SETTINGS_FILE):
        try:
            data = json.loads(path.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return data
        except (OSError, ValueError):
            continue
    return {}


def save_settings(data):
    try:
        core.SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
        tmp = core.SETTINGS_FILE.with_suffix(".tmp")
        tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(tmp, core.SETTINGS_FILE)
    except OSError:
        pass


def thumbnail(path, size=96):
    try:
        with Image.open(path) as img:
            img = ImageOps.exif_transpose(img).convert("RGBA")
            img.thumbnail((size, size), Image.Resampling.LANCZOS)
            buf = io.BytesIO()
            img.save(buf, "PNG")
        return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()
    except (OSError, ValueError):
        return ""


APP_TITLE = "DocReplacer"


class Api:
    def __init__(self):
        self._window = None
        self.settings = load_settings()
        self.busy = False

    # ---- plumbing -------------------------------------------------------
    def _emit(self, event, payload=None):
        if self._window:
            self._window.evaluate_js(f"window.app && app.emit({json.dumps(event)}, {json.dumps(payload, ensure_ascii=False, default=str)})")

    def _log(self, text, tag="info"):
        for line in str(text).strip("\n").split("\n"):
            if line.strip():
                self._emit("log", {"text": line, "tag": tag, "t": time.strftime("%H:%M:%S")})

    def _progress(self, text):
        self._emit("progress", text)

    def _run(self, job, *args):
        if self.busy:
            return {"error": "Başka bir işlem sürüyor."}
        self.busy = True
        self._emit("busy", True)

        def wrapper():
            try:
                job(*args)
            except Exception as exc:  # surfaced to UI
                self._log(f"Hata: {exc}", "err")
                self._emit("toast", {"kind": "err", "title": "İşlem başarısız", "text": str(exc)})
                notify(self._window, "İşlem başarısız", str(exc), ok=False)
            finally:
                self.busy = False
                self._progress("")
                self._emit("busy", False)

        threading.Thread(target=wrapper, daemon=True).start()
        return {"ok": True}

    # ---- state ----------------------------------------------------------
    def get_state(self):
        s = self.settings
        folder = s.get("folder", "")
        return {
            "version": core.APP_VERSION,
            "folder": folder if folder and Path(folder).is_dir() else "",
            "logo": s.get("logo", "") if s.get("logo") and Path(s["logo"]).is_file() else "",
            "create_pdf": s.get("create_pdf", True),
            "batch_pdf": s.get("batch_pdf", True),
        }

    def save(self, patch):
        self.settings.update(patch or {})
        save_settings(self.settings)

    # ---- dialogs --------------------------------------------------------
    def pick_folder(self, title="Klasör seçin"):
        res = self._window.create_file_dialog(webview.FOLDER_DIALOG)
        return res[0] if res else ""

    def pick_logo(self):
        res = self._window.create_file_dialog(webview.OPEN_DIALOG, file_types=IMAGE_TYPES)
        return res[0] if res else ""

    def pick_logos(self):
        res = self._window.create_file_dialog(webview.OPEN_DIALOG, allow_multiple=True, file_types=IMAGE_TYPES)
        return self.describe_logos(list(res or []))

    def pick_zips(self):
        res = self._window.create_file_dialog(webview.OPEN_DIALOG, allow_multiple=True, file_types=("ZIP arşivleri (*.zip)",))
        return list(res or [])

    def open_path(self, path, reveal=False):
        if not path or not Path(path).exists():
            return False
        if reveal:
            subprocess.Popen(["explorer", "/select,", str(Path(path))])
        else:
            os.startfile(str(path))
        return True

    # ---- single folder --------------------------------------------------
    def classify(self, paths):
        out = []
        for p in paths:
            path = Path(p)
            if path.is_dir():
                out.append({"path": str(path), "type": "folder"})
            elif path.suffix.lower() in core.IMAGE_SUFFIXES:
                out.append({"path": str(path), "type": "image"})
            elif path.suffix.lower() == ".zip":
                out.append({"path": str(path), "type": "zip"})
            elif path.is_file():
                out.append({"path": str(path.parent), "type": "folder"})
        return out

    def inspect(self, folder):
        prewarm_office()
        try:
            d = core.inspect_folder(folder)
        except Exception as exc:
            return {"error": str(exc)}
        target = d["target"]
        state = "unparsed" if target is None else ("match" if target == d["folder"] else "differs")
        return {
            "folder": str(d["folder"]),
            "folder_name": d["folder"].name,
            "fr": d["fr_path"].name,
            "name": d["business_name"],
            "state": state,
            "target_name": target.name if target else "",
            "docx": len(core.collect_docx(d["folder"])),
        }

    def preview_target(self, folder, new_name):
        try:
            t = core.folder_rename_target(folder, new_name) if new_name.strip() else None
            return t.name if t else ""
        except Exception:
            return ""

    def sync_folder(self, folder):
        try:
            d, renamed = core.sync_folder_name(folder)
            self._log(f"Klasör yeniden adlandırıldı → {d['target'].name}", "ok")
            return {"folder": str(d["target"]), "renamed": renamed}
        except Exception as exc:
            return {"error": str(exc)}

    def logo_preview(self, path, size=720):
        return thumbnail(path, size)

    def prepare_single(self, v):
        def job():
            self._log(f"Hazırlanıyor: {v['old']} → {v['new']}", "summary")
            result = core.prepare_hotel_copy(v["folder"], v["old"], v["new"], v.get("logo") or None,
                                             log=self._log, progress=self._progress)
            pdf = None
            if result["ok"] and v.get("pdf"):
                self._progress("Birleşik PDF hazırlanıyor…")
                pdf = combine_to_pdf(result["folder"], log=self._log, progress=self._progress)
                self._log(f"PDF hazır: {pdf['output'].name} ({pdf['pages']} sayfa)", "ok")
            notify(self._window, "Klasör hazır", f"{v['new']} için hazırlama tamamlandı.")
            self._emit("single_done", {
                "ok": result["ok"], "folder": str(result["folder"]),
                "folder_name": Path(result["folder"]).name,
                "zip": str(result["zip"]) if result["zip"] else "",
                "pdf": str(pdf["output"]) if pdf else "", "pages": pdf["pages"] if pdf else 0,
                "changed": result["changed"], "unchanged": result["unchanged"], "failed": result["failed"],
                "new": v["new"],
            })
        return self._run(job)

    # ---- batch ----------------------------------------------------------
    def describe_logos(self, paths):
        out = []
        for item in paths:
            p = Path(item)
            files = sorted(x for x in p.iterdir() if x.suffix.lower() in core.IMAGE_SUFFIXES) if p.is_dir() else (
                [p] if p.suffix.lower() in core.IMAGE_SUFFIXES else [])
            for f in files:
                out.append({"path": str(f), "file": f.name, "guess": core.name_from_logo(f),
                            "key": core.match_key(core.name_from_logo(f)), "thumb": thumbnail(f)})
        return out

    def key(self, text):
        return core.match_key(text)

    def match_logo_folder(self, names):
        folder = self.pick_folder()
        if not folder:
            return None
        images = [p for p in Path(folder).iterdir() if p.suffix.lower() in core.IMAGE_SUFFIXES]
        result = {}
        for name in names:
            key = core.match_key(name)
            exact = [p for p in images if core.match_key(p.stem) == key]
            loose = [p for p in images if key and len(core.match_key(p.stem)) >= 3 and (
                key in core.match_key(p.stem) or core.match_key(p.stem) in key)]
            pick = exact[0] if exact else (loose[0] if len(loose) == 1 else None)
            if pick:
                result[name] = {"path": str(pick), "file": pick.name, "thumb": thumbnail(pick)}
        return result

    def run_batch(self, source, rows, pdf_too):
        try:
            d = core.inspect_folder(source)
        except Exception as exc:
            return {"error": f"Kaynak klasör okunamadı: {exc}"}
        old = d["business_name"]
        seen = set()
        for r in rows:
            name = r["name"].strip()
            k = core.normalized(name).casefold()
            if not name:
                return {"error": "Listede adı boş bir satır var."}
            if k in seen:
                return {"error": f"Listede aynı ad iki kez var: {name}"}
            if core.same_name(name, old):
                return {"error": f"“{name}” kaynak işletmeyle aynı."}
            if r.get("logo") and not Path(r["logo"]).is_file():
                return {"error": f"Logo bulunamadı: {r['logo']}"}
            seen.add(k)

        def job():
            from concurrent.futures import ThreadPoolExecutor
            ok = [0]
            pdf_pool = ThreadPoolExecutor(max_workers=1)  # PDF of hotel N overlaps copy of hotel N+1
            pending = []

            def make_pdf(r, folder, prefix):
                try:
                    combine_to_pdf(folder, log=self._log, progress=lambda t: self._progress(prefix + t))
                    ok[0] += 1
                    self._emit("row", {"id": r["id"], "status": "ok", "label": "Klasör + ZIP + PDF", "folder": str(folder)})
                except Exception as exc:
                    self._log(f"{r['name']} PDF: {exc}", "err")
                    self._emit("row", {"id": r["id"], "status": "err", "label": "PDF hatası", "folder": str(folder)})

            self._log(f"Toplu hazırlama · kaynak: {d['folder'].name} · mevcut ad: {old}", "summary")
            for i, r in enumerate(rows, 1):
                prefix = f"{i}/{len(rows)} · {r['name']} · "
                self._emit("row", {"id": r["id"], "status": "run", "label": "Hazırlanıyor"})
                try:
                    res = core.prepare_hotel_copy(
                        d["folder"], old, r["name"], r.get("logo") or None,
                        log=lambda t, tag="info": None if t.startswith("✓") else self._log(t, tag),
                        progress=lambda t: self._progress(prefix + t))
                    if not res["ok"]:
                        self._emit("row", {"id": r["id"], "status": "err", "label": f"{res['failed']} hata",
                                           "folder": str(res["folder"])})
                        continue
                    if pdf_too:
                        self._emit("row", {"id": r["id"], "status": "run", "label": "PDF sırada"})
                        pending.append(pdf_pool.submit(make_pdf, r, res["folder"], prefix))
                    else:
                        ok[0] += 1
                        self._emit("row", {"id": r["id"], "status": "ok", "label": "Klasör + ZIP", "folder": str(res["folder"])})
                except Exception as exc:
                    self._log(f"{r['name']}: {exc}", "err")
                    self._emit("row", {"id": r["id"], "status": "err", "label": "Hata"})
            for future in pending:
                future.result()
            pdf_pool.shutdown()
            self._log(f"Toplu işlem tamamlandı: {ok[0]}/{len(rows)} başarılı.", "summary")
            self._emit("batch_done", {"ok": ok[0], "total": len(rows)})
            notify(self._window, "Toplu işlem bitti", f"{ok[0]}/{len(rows)} işletme hazırlandı.", ok[0] == len(rows))
        return self._run(job)

    # ---- pdf ------------------------------------------------------------
    def run_pdfs(self, items):
        def job():
            for it in items:
                self._emit("pdfrow", {"id": it["id"], "status": "run", "label": "Birleştiriliyor"})
                try:
                    res = combine_to_pdf(it["path"], log=self._log, progress=self._progress)
                    self._log(f"PDF oluşturuldu: {res['output'].name} · {res['pages']} sayfa · {res['seconds']:.1f} sn", "ok")
                    self._emit("pdfrow", {"id": it["id"], "status": "ok", "label": f"{res['pages']} sayfa",
                                          "output": str(res["output"])})
                except Exception as exc:
                    self._log(f"PDF hatası ({Path(it['path']).name}): {exc}", "err")
                    self._emit("pdfrow", {"id": it["id"], "status": "err", "label": "Hata"})
            notify(self._window, "PDF hazır", f"{len(items)} PDF birleştirme tamamlandı.")
        return self._run(job)


def notify(window, title, text, ok=True):
    """Finish sound plus a Windows notification, so a long job can run in the background."""
    try:
        import winsound
        winsound.PlaySound("SystemAsterisk" if ok else "SystemHand",
                           winsound.SND_ALIAS | winsound.SND_ASYNC)
    except Exception:
        pass
    try:
        import clr  # noqa: F401
        from System import Action
        from System.Drawing import Icon
        from System.Windows.Forms import NotifyIcon, ToolTipIcon

        def show():
            tray = NotifyIcon()
            tray.Icon = Icon(str(RESOURCE_DIR / "ui" / "app.ico"))
            tray.Text = APP_TITLE
            tray.Visible = True
            tray.ShowBalloonTip(6000, title, text, ToolTipIcon.Info if ok else ToolTipIcon.Warning)
            threading.Timer(12, lambda: window.native.Invoke(Action(tray.Dispose))).start()

        window.native.Invoke(Action(show))
    except Exception:
        pass


def set_window_icon(window):
    try:
        import clr  # noqa: F401  (pythonnet, loaded by pywebview)
        from System import Action
        from System.Drawing import Icon
        icon = Icon(str(RESOURCE_DIR / "ui" / "app.ico"))
        window.native.Invoke(Action(lambda: setattr(window.native, "Icon", icon)))
    except Exception:
        pass


def style_titlebar(window):
    set_window_icon(window)
    """Dark, colour-matched native title bar on Windows 11."""
    try:
        hwnd = window.native.Handle.ToInt32()
        dwm = ctypes.windll.dwmapi
        on = ctypes.c_int(1)
        dwm.DwmSetWindowAttribute(hwnd, 20, ctypes.byref(on), 4)
        color = ctypes.c_int(0x00110F0E)  # COLORREF 0x00BBGGRR -> #0E0F11
        dwm.DwmSetWindowAttribute(hwnd, 35, ctypes.byref(color), 4)
    except Exception:
        pass


def main():
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("ZenLapis.DocReplacer")
    except Exception:
        pass
    api = Api()
    geo = api.settings.get("window", {})
    window = webview.create_window(
        "DocReplacer", url=UI_FILE.as_uri(), js_api=api,
        width=geo.get("w", 1240), height=geo.get("h", 800), min_size=(980, 660),
        background_color="#0E0F11", text_select=False,
    )
    api._window = window

    def on_drop(event):
        files = (event.get("dataTransfer") or {}).get("files") or []
        paths = [f.get("pywebviewFullPath") for f in files if f.get("pywebviewFullPath")]
        if paths:
            window.evaluate_js(f"app.onDrop({json.dumps(api.classify(paths), ensure_ascii=False)})")

    def on_loaded():
        window.dom.document.events.drop += DOMEventHandler(on_drop, True, True)

    def on_resized(w, h):
        api.settings["window"] = {"w": w, "h": h}

    def on_closing():
        save_settings(api.settings)

    window.events.loaded += on_loaded
    window.events.shown += lambda: style_titlebar(window)
    window.events.resized += on_resized
    window.events.closing += on_closing
    webview.start(gui="edgechromium", private_mode=False,
                  storage_path=str(core.SETTINGS_DIR / "webview"))


if __name__ == "__main__":
    main()
