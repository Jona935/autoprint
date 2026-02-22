#!/usr/bin/env python3
"""
AutoPrint GUI - Impresion automatica de PDFs
Aplicacion de bandeja del sistema para imprimir PDFs automaticamente.
"""

import os
import sys
import time
import json
import shutil
import queue as _queue
import threading
import winreg
import win32print
import subprocess
from pathlib import Path
from datetime import date, datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import pystray
from PIL import Image, ImageDraw, ImageTk

# ===== CONSTANTES =====
APP_NAME     = "AutoPrint"
APP_VERSION  = "1.2"
CONFIG_DIR   = Path(os.environ.get("APPDATA", Path.home())) / "AutoPrint"
CONFIG_FILE  = CONFIG_DIR / "config.json"
HISTORY_FILE = CONFIG_DIR / "history.log"
PENDING_FILE  = CONFIG_DIR / "pending.json"
LASTSEEN_FILE = CONFIG_DIR / "last_seen.json"  # {folder: iso_ts} ultima vez activo por carpeta
STARTUP_KEY   = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Run"

# Paleta de colores
C_BG      = "#1a1a2e"
C_SURFACE = "#16213e"
C_CARD    = "#0f3460"
C_ACCENT  = "#e94560"
C_SUCCESS = "#10b981"
C_DANGER  = "#ef4444"
C_WARNING = "#f59e0b"
C_TEXT    = "#e2e8f0"
C_MUTED   = "#94a3b8"
C_INPUT   = "#1e2a45"
C_BORDER  = "#2d3748"


# ===== CONFIGURACION PERSISTENTE =====

class Config:
    DEFAULTS = {
        "rules":          [],       # lista de reglas {name, folder, printer, archive_enabled, archive_folder}
        "wait_seconds":   3,
        "autostart":      False,
        "active":         False,
        "printed_today":  0,
        "printed_total":  0,
        "last_date":      "",
        "widget_x":       -1,
        "widget_y":       -1,
        "widget_visible": False,
        "notify_detect":   True,
        "notify_print":    True,
        "notify_error":    True,
        "last_file":       "",
        "schedule_enabled": False,
        "schedule_start":   "08:00",
        "schedule_end":     "18:00",
    }

    def get(self, k, default=None):
        return self._data.get(k, self.DEFAULTS.get(k, default))

    def __init__(self):
        self._data = dict(self.DEFAULTS)
        self.load()
        self._check_daily_reset()

    def _check_daily_reset(self):
        today = date.today().isoformat()
        if self._data.get("last_date") != today:
            self._data["printed_today"] = 0
            self._data["last_date"]     = today
            self.save()

    def load(self):
        try:
            if CONFIG_FILE.exists():
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self._data.update(data)
                # Migracion v1.x -> v1.2: convertir carpeta/impresora unica a regla
                if not self._data.get("rules"):
                    folder  = data.get("folder", "")
                    printer = data.get("printer", "")
                    if folder and printer:
                        self._data["rules"] = [{
                            "name":            "Regla 1",
                            "folder":          folder,
                            "printer":         printer,
                            "archive_enabled": data.get("archive_enabled", False),
                            "archive_folder":  data.get("archive_folder", ""),
                        }]
        except Exception:
            pass

    def save(self):
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(self._data, f, indent=2, ensure_ascii=False)

    def __getitem__(self, k):
        return self._data.get(k, self.DEFAULTS.get(k))

    def __setitem__(self, k, v):
        self._data[k] = v
        self.save()


# ===== COLA DE IMPRESION (serializa jobs para evitar conflictos) =====

class PrintQueue:
    """
    Un solo hilo worker procesa los trabajos de impresion de uno en uno.
    Evita que dos Acrobats arranquen simultaneamente contra la misma impresora.
    """
    _ACROBAT_TIMEOUT = 45   # segundos max esperando que Acrobat cierre
    _GAP_BETWEEN     = 2    # segundos de pausa entre trabajos consecutivos

    def __init__(self, log_fn):
        self._q        = _queue.Queue()
        self._log      = log_fn
        self._running  = True
        self._thread   = threading.Thread(target=self._run, daemon=True,
                                          name="PrintQueueWorker")
        self._thread.start()

    @property
    def pending(self):
        return self._q.qsize()

    def submit(self, job: dict):
        """
        job = {
            "acrobat":  str,          # ruta a Acrobat.exe
            "path":     str,          # PDF a imprimir
            "printer":  str,          # nombre de impresora
            "on_done":  callable,     # fn(status:str) -> None
        }
        """
        self._q.put(job)

    def _run(self):
        last_finished = 0.0
        while self._running:
            try:
                job = self._q.get(timeout=1)
            except _queue.Empty:
                continue

            # Pausa minima entre trabajos consecutivos
            gap = self._GAP_BETWEEN - (time.time() - last_finished)
            if gap > 0:
                time.sleep(gap)

            name    = Path(job["path"]).name
            pending = self._q.qsize()
            extra   = f"  ({pending} en cola)" if pending else ""
            self._log(f"[Cola] Imprimiendo: {name} -> {job['printer']}{extra}")

            status = "OK"
            try:
                proc = subprocess.Popen(
                    [job["acrobat"], "/t", job["path"], job["printer"]]
                )
                try:
                    proc.wait(timeout=self._ACROBAT_TIMEOUT)
                except subprocess.TimeoutExpired:
                    proc.kill()
                    self._log(f"[Cola] AVISO: Acrobat cerrado por timeout — "
                              f"trabajo ya enviado al spooler")
            except Exception as e:
                status = f"ERROR: {e}"
                self._log(f"[Cola] ERROR: {e}")

            last_finished = time.time()

            if job.get("on_done"):
                try:
                    job["on_done"](status)
                except Exception:
                    pass

            self._q.task_done()


# ===== MANEJADOR DE ARCHIVOS PDF =====

class PDFHandler(FileSystemEventHandler):
    def __init__(self, printer, acrobat_path, wait_seconds, log_fn,
                 archive_enabled=False, archive_folder="",
                 on_detected_fn=None, on_printed_fn=None,
                 rule_name="", print_queue=None,
                 schedule_fn=None, on_pending_fn=None):
        super().__init__()
        self.printer         = printer
        self.acrobat_path    = acrobat_path
        self.wait_seconds    = wait_seconds
        self.log_fn          = log_fn
        self.archive_enabled = archive_enabled
        self.archive_folder  = archive_folder
        self.on_detected_fn  = on_detected_fn
        self.on_printed_fn   = on_printed_fn
        self.rule_name       = rule_name
        self.print_queue     = print_queue
        self.schedule_fn     = schedule_fn    # () -> bool: estamos dentro del horario?
        self.on_pending_fn   = on_pending_fn  # (job_dict) -> None: guardar para despues
        self._printed        = set()

    def _log(self, msg):
        prefix = f"[{self.rule_name}] " if self.rule_name else ""
        self.log_fn(f"{prefix}{msg}")

    def on_created(self, event):
        if event.is_directory:
            return
        path = event.src_path
        if not path.lower().endswith(".pdf"):
            return

        name = Path(path).name
        self._log(f"PDF detectado: {name}")
        if self.on_detected_fn:
            self.on_detected_fn(name, self.rule_name)

        time.sleep(self.wait_seconds)

        if path not in self._printed:
            self._printed.add(path)   # marcar YA para evitar doble envio

            # Verificar si estamos dentro del horario
            in_schedule = self.schedule_fn() if self.schedule_fn else True

            if not in_schedule:
                # Guardar para imprimir cuando inicie el horario
                self._log(f"Fuera de horario — guardado para despues: {name}")
                if self.on_pending_fn:
                    self.on_pending_fn({
                        "path":            path,
                        "printer":         self.printer,
                        "acrobat":         self.acrobat_path,
                        "rule_name":       self.rule_name,
                        "archive_enabled": self.archive_enabled,
                        "archive_folder":  self.archive_folder,
                        "detected_at":     datetime.now().isoformat(timespec="seconds"),
                    })
                return

            self._submit_to_queue(path, name)

    def _submit_to_queue(self, path, name=None):
        if name is None:
            name = Path(path).name

        pending = self.print_queue.pending if self.print_queue else 0
        if pending > 0:
            self._log(f"Encolando: {name}  ({pending} trabajo(s) antes)")
        else:
            self._log(f"Enviando a cola: {name}")

        def _on_done(status, _name=name, _path=path):
            if status == "OK":
                self._log(f"OK — Impreso en: {self.printer}")
                if self.on_printed_fn:
                    self.on_printed_fn(_name, self.printer, "OK", self.rule_name)
                if self.archive_enabled and self.archive_folder:
                    threading.Thread(
                        target=self._move_to_archive,
                        args=(_path,), daemon=True
                    ).start()
            else:
                self._log(f"ERROR al imprimir: {status}")
                if self.on_printed_fn:
                    self.on_printed_fn(_name, self.printer, status, self.rule_name)

        if self.print_queue:
            self.print_queue.submit({
                "acrobat": self.acrobat_path,
                "path":    path,
                "printer": self.printer,
                "on_done": _on_done,
            })
        else:
            try:
                subprocess.Popen([self.acrobat_path, "/t", path, self.printer])
                _on_done("OK")
            except Exception as e:
                _on_done(f"ERROR: {e}")

    def _move_to_archive(self, src_path):
        time.sleep(8)
        src  = Path(src_path)
        dest = Path(self.archive_folder) / src.name

        if dest.exists():
            ts   = time.strftime("%Y%m%d_%H%M%S")
            dest = Path(self.archive_folder) / f"{src.stem}_{ts}{src.suffix}"

        copied = False
        for intento in range(1, 7):
            try:
                if not src.exists():
                    return
                shutil.copy2(str(src), str(dest))
                copied = True
                break
            except Exception as e:
                self._log(f"Copiando... intento {intento}/6 ({e})")
                time.sleep(5)

        if not copied:
            self._log(f"ERROR: No se pudo copiar '{src.name}' al archivo local")
            return

        self._log(f"Copiado a local: {dest.name}")

        for intento in range(1, 11):
            try:
                src.unlink()
                self._log(f"Eliminado del Drive: {src.name}")
                self._log(f"Archivado en: {self.archive_folder}")
                return
            except PermissionError:
                time.sleep(6)
            except FileNotFoundError:
                self._log(f"Archivado en: {self.archive_folder}")
                return
            except Exception as e:
                self._log(f"ERROR eliminando del Drive (intento {intento}/10): {e}")
                time.sleep(6)

        self._log(f"AVISO: '{src.name}' copiado pero no eliminado del Drive.")


# ===== WIDGET FLOTANTE GLASSMORPHISM =====

class FloatingWidget:
    """Widget de escritorio con efecto glassmorphism via DWM de Windows."""

    _TKEY = "#010203"   # color "llave" usado para transparencia

    def __init__(self, app):
        self.app     = app
        self.win     = None
        self._drag_x = 0
        self._drag_y = 0

    def show(self):
        if self.win and self.win.winfo_exists():
            self.win.deiconify()
            self._pin_to_desktop()
            return

        win = tk.Toplevel()
        self.win = win
        win.overrideredirect(True)
        win.wm_attributes("-topmost", False)
        win.configure(bg=self._TKEY)
        win.resizable(False, False)

        sw = win.winfo_screenwidth()
        sh = win.winfo_screenheight()
        x = self.app.config["widget_x"]
        y = self.app.config["widget_y"]
        if x < 0 or y < 0:
            x = sw - 280
            y = 80
        x = max(0, min(x, sw - 260))
        y = max(0, min(y, sh - 160))
        win.geometry(f"260x160+{x}+{y}")

        win.update_idletasks()
        self._apply_glass()
        self._build(win)

        win.after(150, self._pin_to_desktop)
        win.after(2000, self._keep_on_desktop)

    def _apply_glass(self):
        """Activa el efecto Acrylic/Blur de Windows 10/11 via DWM."""
        import ctypes

        class ACCENT_POLICY(ctypes.Structure):
            _fields_ = [
                ("AccentState",   ctypes.c_uint),
                ("AccentFlags",   ctypes.c_uint),
                ("GradientColor", ctypes.c_uint),
                ("AnimationId",   ctypes.c_uint),
            ]

        class WCAD(ctypes.Structure):
            _fields_ = [
                ("Attribute",  ctypes.c_uint),
                ("Data",       ctypes.c_void_p),
                ("SizeOfData", ctypes.c_size_t),
            ]

        try:
            accent = ACCENT_POLICY()
            accent.AccentState   = 4            # ACCENT_ENABLE_ACRYLICBLURBEHIND
            accent.AccentFlags   = 2
            accent.GradientColor = 0x18101828   # azul oscuro, ~10% opacidad

            data = WCAD()
            data.Attribute  = 19                # WCA_ACCENT_POLICY
            data.Data       = ctypes.cast(ctypes.pointer(accent), ctypes.c_void_p)
            data.SizeOfData = ctypes.sizeof(accent)

            hwnd = int(self.win.wm_frame(), 16)
            ctypes.windll.user32.SetWindowCompositionAttribute(
                hwnd, ctypes.byref(data)
            )
            self.win.wm_attributes("-transparentcolor", self._TKEY)
        except Exception:
            self.win.wm_attributes("-alpha", 0.80)

    def _pin_to_desktop(self):
        try:
            import ctypes
            hwnd = int(self.win.wm_frame(), 16)
            ctypes.windll.user32.SetWindowPos(
                hwnd, 1, 0, 0, 0, 0,
                0x0002 | 0x0001 | 0x0010
            )
        except Exception:
            pass

    def _keep_on_desktop(self):
        if self.win and self.win.winfo_exists() and self.win.winfo_viewable():
            self._pin_to_desktop()
            self.win.after(2000, self._keep_on_desktop)

    def hide(self):
        if self.win and self.win.winfo_exists():
            self.win.withdraw()
        self.app.config["widget_visible"] = False

    def _build(self, win):
        G     = self._TKEY
        _DRAG = "#0d0d1a"

        # Barra de arrastre con fondo visible
        bar = tk.Frame(win, bg=_DRAG, pady=4)
        bar.pack(fill="x")
        self._bind_drag(bar)

        lbl_title = tk.Label(bar, text="AutoPrint",
                             font=("Segoe UI", 8, "bold"),
                             bg=_DRAG, fg="#cccccc")
        lbl_title.pack(side="left", padx=10)
        self._bind_drag(lbl_title)

        btn_close = tk.Button(bar, text="X",
                              font=("Segoe UI", 8), bg=_DRAG, fg="#aaaaaa",
                              activebackground=_DRAG, activeforeground="#ef4444",
                              relief="flat", cursor="hand2", bd=0,
                              command=self.hide)
        btn_close.pack(side="right", padx=8)

        # Estado
        self._lbl_status = tk.Label(win, text="Detenido",
                                    font=("Segoe UI", 8, "bold"),
                                    bg=G, fg=C_DANGER)
        self._lbl_status.pack(anchor="w", padx=14, pady=(4, 2))

        # Contadores grandes
        cnt = tk.Frame(win, bg=G)
        cnt.pack(fill="x", padx=14, pady=(2, 6))

        col_hoy = tk.Frame(cnt, bg=G)
        col_hoy.pack(side="left", expand=True)
        self._lbl_hoy = tk.Label(col_hoy, text="0",
                                  font=("Segoe UI", 42, "bold"),
                                  bg=G, fg="#10b981")
        self._lbl_hoy.pack()
        self._bind_drag(self._lbl_hoy)
        lbl_hoy_sub = tk.Label(col_hoy, text="HOY",
                               font=("Segoe UI", 7, "bold"),
                               bg=G, fg="#aaaaaa")
        lbl_hoy_sub.pack()
        self._bind_drag(lbl_hoy_sub)

        tk.Frame(cnt, bg="#555555", width=1).pack(
            side="left", fill="y", padx=10, pady=6)

        col_tot = tk.Frame(cnt, bg=G)
        col_tot.pack(side="left", expand=True)
        self._lbl_total = tk.Label(col_tot, text="0",
                                    font=("Segoe UI", 42, "bold"),
                                    bg=G, fg="#e2e8f0")
        self._lbl_total.pack()
        self._bind_drag(self._lbl_total)
        lbl_tot_sub = tk.Label(col_tot, text="TOTAL",
                               font=("Segoe UI", 7, "bold"),
                               bg=G, fg="#aaaaaa")
        lbl_tot_sub.pack()
        self._bind_drag(lbl_tot_sub)

        tk.Frame(win, bg="#555555", height=1).pack(fill="x", padx=14, pady=(0, 4))

        self._lbl_last = tk.Label(win, text="Sin actividad",
                                   font=("Segoe UI", 7),
                                   bg=G, fg="#94a3b8",
                                   wraplength=240, justify="left")
        self._lbl_last.pack(anchor="w", padx=14, pady=(0, 8))

        self.refresh()

    def _bind_drag(self, widget):
        widget.bind("<ButtonPress-1>",   self._drag_start)
        widget.bind("<B1-Motion>",       self._drag_move)
        widget.bind("<ButtonRelease-1>", self._drag_end)

    def _drag_start(self, e):
        self._drag_x = e.x_root - self.win.winfo_x()
        self._drag_y = e.y_root - self.win.winfo_y()

    def _drag_move(self, e):
        self.win.geometry(f"+{e.x_root - self._drag_x}+{e.y_root - self._drag_y}")

    def _drag_end(self, e):
        self.app.config["widget_x"] = self.win.winfo_x()
        self.app.config["widget_y"] = self.win.winfo_y()

    def refresh(self):
        if not self.win or not self.win.winfo_exists():
            return
        today = self.app.config["printed_today"]
        total = self.app.config["printed_total"]
        last  = self.app.config["last_file"] or "Sin actividad"
        st    = "Activo" if self.app.is_watching else "Detenido"
        c_st  = C_SUCCESS if self.app.is_watching else C_DANGER
        try:
            self._lbl_status.config(text=st, fg=c_st)
            self._lbl_hoy.config(text=str(today))
            self._lbl_total.config(text=str(total))
            self._lbl_last.config(text=last)
        except Exception:
            pass

    def is_visible(self):
        return bool(self.win and self.win.winfo_exists()
                    and self.win.winfo_viewable())


# ===== DIALOGO DE REGLA =====

class RuleDialog:
    """Dialogo para agregar o editar una regla de vigilancia."""

    def __init__(self, parent, app, rule=None, on_save=None):
        self.app     = app
        self.on_save = on_save
        self.result  = None

        self.win = tk.Toplevel(parent)
        self.win.title("Nueva regla" if rule is None else "Editar regla")
        self.win.geometry("520x460")
        self.win.resizable(False, False)
        self.win.configure(bg=C_BG)
        self.win.grab_set()
        self.win.transient(parent)

        rule = rule or {}
        self._v_name            = tk.StringVar(value=rule.get("name", ""))
        self._v_folder          = tk.StringVar(value=rule.get("folder", ""))
        self._v_printer         = tk.StringVar(value=rule.get("printer", ""))
        self._v_archive_enabled = tk.BooleanVar(value=rule.get("archive_enabled", False))
        self._v_archive_folder  = tk.StringVar(value=rule.get("archive_folder", ""))

        self._build()

    def _build(self):
        win = self.win
        pad = dict(padx=20)

        tk.Label(win, text="Configurar regla",
                 font=("Segoe UI", 13, "bold"),
                 bg=C_BG, fg=C_TEXT).pack(pady=(18, 14), **pad, anchor="w")

        # Nombre
        self._field(win, "Nombre (opcional)")
        tk.Entry(win, textvariable=self._v_name,
                 font=("Segoe UI", 10), bg=C_INPUT, fg=C_TEXT,
                 relief="flat", insertbackground=C_TEXT).pack(
            fill="x", **pad, ipady=6, pady=(0, 10))

        # Carpeta
        self._field(win, "Carpeta a vigilar")
        row_f = tk.Frame(win, bg=C_BG)
        row_f.pack(fill="x", **pad, pady=(0, 10))
        tk.Entry(row_f, textvariable=self._v_folder,
                 font=("Segoe UI", 10), bg=C_INPUT, fg=C_TEXT,
                 relief="flat", insertbackground=C_TEXT).pack(
            side="left", fill="x", expand=True, ipady=6, padx=(0, 6))
        self._btn(row_f, "Examinar", self._browse_folder).pack(side="left", padx=(0, 4))
        if self.app.gdrive_path:
            tk.Button(row_f, text="Drive",
                      font=("Segoe UI", 9, "bold"),
                      bg="#1967d2", fg="white", relief="flat", cursor="hand2",
                      padx=8, pady=4,
                      command=lambda: self._v_folder.set(self.app.gdrive_path)
                      ).pack(side="left", padx=(0, 4))
        if self.app.onedrive_path:
            tk.Button(row_f, text="OneDrive",
                      font=("Segoe UI", 9, "bold"),
                      bg="#0078d4", fg="white", relief="flat", cursor="hand2",
                      padx=8, pady=4,
                      command=lambda: self._v_folder.set(self.app.onedrive_path)
                      ).pack(side="left")

        # Impresora
        self._field(win, "Impresora")
        row_p = tk.Frame(win, bg=C_BG)
        row_p.pack(fill="x", **pad, pady=(0, 10))
        printers = self._get_printers()
        self._cb_printer = ttk.Combobox(row_p, textvariable=self._v_printer,
                                         values=printers, state="readonly",
                                         font=("Segoe UI", 10))
        self._cb_printer.pack(side="left", fill="x", expand=True)
        if not self._v_printer.get() and printers:
            self._v_printer.set(printers[0])
        self._btn(row_p, "Actualizar", self._refresh_printers).pack(side="left", padx=(6, 0))

        # Carpeta de archivo
        chk_frame = tk.Frame(win, bg=C_BG)
        chk_frame.pack(fill="x", **pad, pady=(4, 0))
        tk.Checkbutton(chk_frame,
                       text="Mover PDFs a carpeta local despues de imprimir",
                       variable=self._v_archive_enabled,
                       font=("Segoe UI", 9),
                       bg=C_BG, fg=C_TEXT, selectcolor=C_SURFACE,
                       activebackground=C_BG, activeforeground=C_TEXT,
                       command=self._on_archive_toggle).pack(side="left")

        self._arch_row = tk.Frame(win, bg=C_BG)
        self._arch_row.pack(fill="x", **pad, pady=(4, 10))
        self._arch_entry = tk.Entry(self._arch_row,
                                    textvariable=self._v_archive_folder,
                                    font=("Segoe UI", 10), bg=C_INPUT, fg=C_TEXT,
                                    relief="flat", insertbackground=C_TEXT)
        self._arch_entry.pack(side="left", fill="x", expand=True, ipady=6, padx=(0, 6))
        self._arch_btn = self._btn(self._arch_row, "Examinar", self._browse_archive)
        self._arch_btn.pack(side="left", padx=(0, 4))
        self._arch_new = tk.Button(self._arch_row, text="+ Nueva",
                                   font=("Segoe UI", 9, "bold"),
                                   bg=C_ACCENT, fg="white",
                                   relief="flat", cursor="hand2", padx=8, pady=4,
                                   command=self._create_archive_folder)
        self._arch_new.pack(side="left")
        self._on_archive_toggle()

        # Botones
        tk.Frame(win, bg=C_BORDER, height=1).pack(fill="x", pady=(8, 0))
        btn_row = tk.Frame(win, bg=C_BG, pady=12)
        btn_row.pack(fill="x", **pad)
        tk.Button(btn_row, text="Guardar",
                  font=("Segoe UI", 10, "bold"),
                  bg=C_SUCCESS, fg="white", relief="flat", cursor="hand2",
                  padx=16, pady=8, command=self._save).pack(side="left", padx=(0, 8))
        tk.Button(btn_row, text="Cancelar",
                  font=("Segoe UI", 10), bg=C_CARD, fg=C_TEXT,
                  relief="flat", cursor="hand2", padx=16, pady=8,
                  command=self.win.destroy).pack(side="left")

    def _field(self, parent, text):
        tk.Label(parent, text=text,
                 font=("Segoe UI", 9, "bold"),
                 bg=C_BG, fg=C_MUTED).pack(anchor="w", padx=20, pady=(0, 3))

    def _btn(self, parent, text, cmd):
        return tk.Button(parent, text=text,
                         font=("Segoe UI", 9), bg=C_CARD, fg=C_TEXT,
                         relief="flat", cursor="hand2", padx=8, pady=4, command=cmd)

    def _get_printers(self):
        try:
            return [p[2] for p in win32print.EnumPrinters(2)]
        except Exception:
            return []

    def _refresh_printers(self):
        new = self._get_printers()
        self._cb_printer["values"] = new
        if new and not self._v_printer.get():
            self._v_printer.set(new[0])

    def _browse_folder(self):
        d = filedialog.askdirectory(title="Carpeta a vigilar",
                                    initialdir=self._v_folder.get() or str(Path.home()),
                                    parent=self.win)
        if d:
            self._v_folder.set(d)

    def _browse_archive(self):
        d = filedialog.askdirectory(title="Carpeta de archivo local",
                                    initialdir=self._v_archive_folder.get() or str(Path.home()),
                                    parent=self.win)
        if d:
            self._v_archive_folder.set(d)

    def _on_archive_toggle(self):
        enabled = self._v_archive_enabled.get()
        state   = "normal" if enabled else "disabled"
        bg      = C_INPUT if enabled else C_BG
        self._arch_entry.config(state=state, bg=bg)
        self._arch_btn.config(state=state)
        self._arch_new.config(state=state)

    def _create_archive_folder(self):
        dialog = tk.Toplevel(self.win)
        dialog.title("Crear carpeta")
        dialog.geometry("380x150")
        dialog.resizable(False, False)
        dialog.configure(bg=C_BG)
        dialog.grab_set()
        dialog.transient(self.win)

        tk.Label(dialog, text="Nombre de la nueva carpeta:",
                 font=("Segoe UI", 10), bg=C_BG, fg=C_TEXT).pack(
            pady=(16, 6), padx=20, anchor="w")
        name_var = tk.StringVar(value="PDFs Impresos")
        entry = tk.Entry(dialog, textvariable=name_var,
                         font=("Segoe UI", 11), bg=C_INPUT, fg=C_TEXT,
                         relief="flat", insertbackground=C_TEXT)
        entry.pack(fill="x", padx=20, ipady=7)
        entry.select_range(0, "end")
        entry.focus_set()

        def do_create():
            name = name_var.get().strip()
            if not name:
                return
            base = filedialog.askdirectory(title="Donde crear la carpeta",
                                            initialdir=str(Path.home()),
                                            parent=dialog)
            if not base:
                return
            new_path = Path(base) / name
            try:
                new_path.mkdir(parents=True, exist_ok=True)
                self._v_archive_folder.set(str(new_path))
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("Error", str(e), parent=dialog)

        tk.Button(dialog, text="Crear",
                  font=("Segoe UI", 10, "bold"),
                  bg=C_SUCCESS, fg="white", relief="flat", cursor="hand2",
                  padx=14, pady=7, command=do_create).pack(pady=12)
        entry.bind("<Return>", lambda e: do_create())

    def _save(self):
        folder  = self._v_folder.get().strip()
        printer = self._v_printer.get().strip()

        if not folder:
            messagebox.showwarning("Falta carpeta",
                                   "Selecciona una carpeta a vigilar.", parent=self.win)
            return
        if not os.path.exists(folder):
            messagebox.showwarning("Carpeta no existe",
                                   f"La carpeta no existe:\n{folder}", parent=self.win)
            return
        if not printer:
            messagebox.showwarning("Falta impresora",
                                   "Selecciona una impresora.", parent=self.win)
            return

        archive_enabled = self._v_archive_enabled.get()
        archive_folder  = self._v_archive_folder.get().strip()
        if archive_enabled and archive_folder and not os.path.exists(archive_folder):
            messagebox.showwarning("Carpeta no existe",
                                   f"La carpeta de archivo no existe:\n{archive_folder}",
                                   parent=self.win)
            return

        name = self._v_name.get().strip() or folder.split("/")[-1].split("\\")[-1]

        self.result = {
            "name":            name,
            "folder":          folder,
            "printer":         printer,
            "archive_enabled": archive_enabled,
            "archive_folder":  archive_folder,
        }
        if self.on_save:
            self.on_save(self.result)
        self.win.destroy()


# ===== HELPER DE ARCHIVO (mover PDF a carpeta local) =====

class _ArchiveHelper:
    """Wrapper minimo para reutilizar la logica de archivado sin un PDFHandler completo."""
    def __init__(self, log_fn, rule_name=""):
        self.log_fn    = log_fn
        self.rule_name = rule_name

    def _log(self, msg):
        prefix = f"[{self.rule_name}] " if self.rule_name else ""
        self.log_fn(f"{prefix}{msg}")

    def move(self, src_path, archive_folder):
        time.sleep(8)
        src  = Path(src_path)
        dest = Path(archive_folder) / src.name
        if dest.exists():
            ts   = time.strftime("%Y%m%d_%H%M%S")
            dest = Path(archive_folder) / f"{src.stem}_{ts}{src.suffix}"
        copied = False
        for i in range(1, 7):
            try:
                if not src.exists():
                    return
                shutil.copy2(str(src), str(dest))
                copied = True
                break
            except Exception as e:
                self._log(f"Copiando... intento {i}/6 ({e})")
                time.sleep(5)
        if not copied:
            self._log(f"ERROR: No se pudo copiar '{src.name}'")
            return
        self._log(f"Copiado a local: {dest.name}")
        for i in range(1, 11):
            try:
                src.unlink()
                self._log(f"Eliminado del Drive: {src.name}")
                return
            except (PermissionError, Exception):
                time.sleep(6)


# ===== APLICACION PRINCIPAL =====

class AutoPrintApp:
    def __init__(self):
        self.config      = Config()
        self._observers  = []          # lista de Observer activos
        self.is_watching = False
        self.tray_icon   = None
        self.root        = None
        self._log_entries = []
        self.widget      = None

        self._status_lbl  = None
        self._toggle_btn  = None
        self._log_widget  = None
        self._lbl_hoy     = None
        self._lbl_total   = None
        self._rules_frame = None      # frame donde se muestran las tarjetas de reglas

        self._v_wait            = None
        self._v_autostart       = None
        self._v_notify_detect   = None
        self._v_notify_print    = None

        self.acrobat_path    = self._find_acrobat()
        self.gdrive_path     = self._find_gdrive()
        self.onedrive_path   = self._find_onedrive()
        self._notify_times    = {}    # clave -> timestamp ultimo envio (anti-spam)
        self._pending_detect  = {}    # rule_name -> archivos pendientes de agrupar
        self._pending_lock    = threading.Lock()
        self._schedule_was_in = False  # estado anterior del horario
        self._monitor_running = False
        self._lbl_pending     = None   # label de pendientes en la UI
        self._v_schedule_enabled = None
        self._v_schedule_start_h = None
        self._v_schedule_start_m = None
        self._v_schedule_end_h   = None
        self._v_schedule_end_m   = None
        # Cola global de impresion — un solo hilo serializa todos los trabajos
        self._print_queue     = PrintQueue(log_fn=self._log)

    # ------------------------------------------------------------------
    # Horario de impresion
    # ------------------------------------------------------------------

    def _is_in_schedule(self):
        """True si el horario esta desactivado o si la hora actual cae dentro del rango."""
        if not self.config["schedule_enabled"]:
            return True
        try:
            now   = datetime.now().time()
            s_str = self.config["schedule_start"]
            e_str = self.config["schedule_end"]
            sh, sm = map(int, s_str.split(":"))
            eh, em = map(int, e_str.split(":"))
            start = now.replace(hour=sh, minute=sm, second=0, microsecond=0)
            end   = now.replace(hour=eh, minute=em, second=0, microsecond=0)
            # Soporte para horario que cruza medianoche (e.g. 22:00 - 06:00)
            if start <= end:
                return start <= now.replace(second=0, microsecond=0) <= end
            else:
                t = now.replace(second=0, microsecond=0)
                return t >= start or t <= end
        except Exception:
            return True

    def _add_pending_job(self, job: dict):
        """Guarda un trabajo pendiente (fuera de horario) en disco y en memoria."""
        with self._pending_lock:
            jobs = self._load_pending_raw()
            jobs.append(job)
            self._save_pending_raw(jobs)
            total = len(jobs)

        sched_start = self.config["schedule_start"] if self.config["schedule_enabled"] else ""
        hora_txt    = f" — se imprimira a las {sched_start}" if sched_start else ""
        self._notify(
            f"Fuera de horario — [{job.get('rule_name','')}]",
            f"{Path(job['path']).name}{hora_txt}",
            cooldown_key=f"pending_{job.get('rule_name','')}",
            cooldown_secs=10,
        )
        if self.root:
            self.root.after(0, self._refresh_pending_label)

    def _load_pending_raw(self) -> list:
        try:
            if PENDING_FILE.exists():
                with open(PENDING_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception:
            pass
        return []

    def _save_pending_raw(self, jobs: list):
        try:
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            with open(PENDING_FILE, "w", encoding="utf-8") as f:
                json.dump(jobs, f, indent=2, ensure_ascii=False)
        except Exception:
            pass

    def pending_count(self) -> int:
        with self._pending_lock:
            return len(self._load_pending_raw())

    def flush_pending_jobs(self, force=False):
        """
        Imprime todos los archivos pendientes.
        force=True omite la verificacion de horario (boton manual).
        """
        if not force and not self._is_in_schedule():
            return
        with self._pending_lock:
            jobs = self._load_pending_raw()
            self._save_pending_raw([])   # limpiar antes de procesar

        if not jobs:
            return

        self._log(f"─── Iniciando impresion de {len(jobs)} archivo(s) pendiente(s) ───")
        for job in jobs:
            path = job.get("path", "")
            if not path or not os.path.exists(path):
                self._log(f"AVISO: archivo ya no existe, omitido: {Path(path).name if path else '?'}")
                continue

            rule_name       = job.get("rule_name", "")
            archive_enabled = job.get("archive_enabled", False)
            archive_folder  = job.get("archive_folder", "")
            detected_at     = job.get("detected_at", "")
            name            = Path(path).name
            extra           = f" (detectado {detected_at})" if detected_at else ""
            self._log(f"[{rule_name}] Imprimiendo pendiente: {name}{extra}")

            def _on_done(status, _job=job, _name=name):
                if status == "OK":
                    self._log(f"[{_job.get('rule_name','')}] OK — Impreso en: {_job['printer']}")
                    self._on_printed(_name, _job["printer"], "OK", _job.get("rule_name",""))
                    if _job.get("archive_enabled") and _job.get("archive_folder"):
                        # Crear un handler temporal solo para archivar
                        h = PDFHandler.__new__(PDFHandler)
                        h.archive_folder  = _job["archive_folder"]
                        h.archive_enabled = True
                        h.log_fn          = self._log
                        h.rule_name       = _job.get("rule_name", "")
                        h._log            = h._log if hasattr(h, '_log') else lambda m: self._log(m)
                        threading.Thread(target=h._move_to_archive,
                                         args=(_job["path"],), daemon=True).start()
                else:
                    self._log(f"[{_job.get('rule_name','')}] ERROR: {status}")

            self._print_queue.submit({
                "acrobat": job.get("acrobat", self.acrobat_path or ""),
                "path":    path,
                "printer": job.get("printer", ""),
                "on_done": _on_done,
            })

        if self.root:
            self.root.after(0, self._refresh_pending_label)

    def _handle_pending_on_startup(self):
        """Llama al arrancar: imprime pendientes si estamos en horario o notifica si no."""
        n = self.pending_count()
        if n == 0:
            return

        if self._is_in_schedule():
            self._log(f"Arranque: {n} archivo(s) pendiente(s) — imprimiendo ahora")
            self._notify(
                "AutoPrint — Archivos pendientes",
                f"Imprimiendo {n} archivo(s) guardados del periodo anterior",
                cooldown_key="startup_flush",
            )
            threading.Thread(target=self.flush_pending_jobs, daemon=True,
                             name="FlushOnStart").start()
        else:
            sched_start = self.config["schedule_start"] if self.config["schedule_enabled"] else ""
            hora_txt    = f" a las {sched_start}" if sched_start else ""
            self._log(f"Arranque fuera de horario: {n} archivo(s) pendiente(s) — esperando{hora_txt}")
            self._notify(
                "AutoPrint — Fuera de horario",
                f"{n} archivo(s) pendiente(s){hora_txt}",
                cooldown_key="startup_pending",
            )

    # ------------------------------------------------------------------
    # Escaneo de arranque — archivos que llegaron con la app apagada
    # ------------------------------------------------------------------

    def _load_last_seen(self) -> dict:
        try:
            if LASTSEEN_FILE.exists():
                with open(LASTSEEN_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception:
            pass
        return {}

    def _save_last_seen(self, data: dict):
        try:
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            with open(LASTSEEN_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
        except Exception:
            pass

    def _touch_last_seen(self):
        """Actualiza el timestamp 'ahora' para todas las carpetas vigiladas activas."""
        rules = self.config.get("rules", [])
        data  = self._load_last_seen()
        now   = datetime.now().isoformat(timespec="seconds")
        for rule in rules:
            folder = rule.get("folder", "")
            if folder:
                data[folder] = now
        self._save_last_seen(data)

    def _scan_missed_files(self):
        """
        Al arrancar, escanea cada carpeta vigilada buscando PDFs cuya fecha de
        modificacion sea posterior a la ultima vez que la app estuvo activa.
        Los encontrados se imprimen (si hay horario) o se encolan como pendientes.
        """
        rules = self.config.get("rules", [])
        if not rules or not self.acrobat_path:
            return

        last_seen  = self._load_last_seen()
        now_str    = datetime.now().isoformat(timespec="seconds")
        found_any  = 0

        for rule in rules:
            folder  = rule.get("folder", "")
            printer = rule.get("printer", "")
            if not folder or not printer or not os.path.exists(folder):
                continue

            rule_name = rule.get("name") or Path(folder).name

            # Timestamp de la ultima vez que vigilamos esta carpeta
            last_str = last_seen.get(folder, "")
            if last_str:
                try:
                    last_ts = datetime.fromisoformat(last_str).timestamp()
                except Exception:
                    last_ts = 0.0
            else:
                # Primera vez que vemos esta carpeta — no imprimir todo lo que hay
                self._log(f"[{rule_name}] Primera ejecucion en esta carpeta, omitiendo archivos existentes")
                last_seen[folder] = now_str
                continue

            # Buscar PDFs mas nuevos que la ultima vez activo
            try:
                nuevos = sorted(
                    [f for f in Path(folder).iterdir()
                     if f.suffix.lower() == ".pdf" and f.stat().st_mtime > last_ts],
                    key=lambda f: f.stat().st_mtime
                )
            except Exception as e:
                self._log(f"[{rule_name}] Error escaneando carpeta: {e}")
                continue

            if not nuevos:
                last_seen[folder] = now_str
                continue

            self._log(f"[{rule_name}] {len(nuevos)} PDF(s) llegaron mientras la app estaba cerrada")
            found_any += len(nuevos)

            for pdf in nuevos:
                path      = str(pdf)
                name      = pdf.name
                mtime_str = datetime.fromtimestamp(pdf.stat().st_mtime).strftime("%d/%m %H:%M")
                self._log(f"[{rule_name}] PDF perdido: {name} (llegó {mtime_str})")

                if self._is_in_schedule():
                    # Imprimir ahora via cola
                    def _on_done(status, _rule=rule, _name=name, _path=path):
                        r_name = _rule.get("name", "")
                        if status == "OK":
                            self._log(f"[{r_name}] OK — Impreso en: {_rule['printer']}")
                            self._on_printed(_name, _rule["printer"], "OK", r_name)
                            if _rule.get("archive_enabled") and _rule.get("archive_folder"):
                                h = _ArchiveHelper(self._log, r_name)
                                threading.Thread(target=h.move,
                                                 args=(_path, _rule["archive_folder"]),
                                                 daemon=True).start()
                        else:
                            self._log(f"[{r_name}] ERROR: {status}")
                            self._on_printed(_name, _rule["printer"], status, r_name)

                    self._print_queue.submit({
                        "acrobat": self.acrobat_path,
                        "path":    path,
                        "printer": printer,
                        "on_done": _on_done,
                    })
                else:
                    self._add_pending_job({
                        "path":            path,
                        "printer":         printer,
                        "acrobat":         self.acrobat_path,
                        "rule_name":       rule_name,
                        "archive_enabled": rule.get("archive_enabled", False),
                        "archive_folder":  rule.get("archive_folder", ""),
                        "detected_at":     mtime_str,
                    })

            last_seen[folder] = now_str

        self._save_last_seen(last_seen)

        if found_any > 0 and self._is_in_schedule():
            self._notify(
                "AutoPrint — Archivos perdidos",
                f"{found_any} PDF(s) llegaron mientras la app estaba cerrada — imprimiendo",
                cooldown_key="missed_found",
            )

    def _schedule_monitor(self):
        """Hilo daemon que detecta la transicion fuera->dentro del horario."""
        self._schedule_was_in = self._is_in_schedule()
        while self._monitor_running:
            time.sleep(30)
            now_in = self._is_in_schedule()
            # Transicion fuera -> dentro del horario
            if not self._schedule_was_in and now_in:
                n = self.pending_count()
                if n > 0:
                    self._log(f"Horario iniciado — imprimiendo {n} archivo(s) pendiente(s)")
                    self._notify(
                        "AutoPrint — Horario iniciado",
                        f"Imprimiendo {n} archivo(s) pendiente(s)",
                        cooldown_key="schedule_start",
                    )
                else:
                    self._log("Horario iniciado")
                threading.Thread(target=self.flush_pending_jobs,
                                 daemon=True, name="FlushPending").start()
            # Transicion dentro -> fuera del horario
            elif self._schedule_was_in and not now_in:
                sched_end = self.config["schedule_start"] if self.config["schedule_enabled"] else ""
                self._log(f"Horario finalizado — nuevos PDFs se guardaran para despues")
            self._schedule_was_in = now_in
            if self.root:
                self.root.after(0, self._refresh_pending_label)

    def _refresh_pending_label(self):
        """Actualiza el label de 'N pendientes' en la UI."""
        if not self._lbl_pending:
            return
        n = self.pending_count()
        if n == 0:
            self._lbl_pending.config(text="Sin archivos pendientes", fg=C_MUTED)
        else:
            in_sch = self._is_in_schedule()
            state  = "se imprimiran ahora" if in_sch else "se imprimiran al inicio del horario"
            self._lbl_pending.config(
                text=f"{n} archivo(s) pendiente(s) — {state}", fg=C_WARNING)

    def _save_schedule_from_ui(self):
        if self._v_schedule_enabled is None:
            return
        enabled = self._v_schedule_enabled.get()
        try:
            sh = int(self._v_schedule_start_h.get())
            sm = int(self._v_schedule_start_m.get())
            eh = int(self._v_schedule_end_h.get())
            em = int(self._v_schedule_end_m.get())
            sh = max(0, min(23, sh)); sm = max(0, min(59, sm))
            eh = max(0, min(23, eh)); em = max(0, min(59, em))
        except (ValueError, AttributeError):
            return
        self.config["schedule_enabled"] = enabled
        self.config["schedule_start"]   = f"{sh:02d}:{sm:02d}"
        self.config["schedule_end"]     = f"{eh:02d}:{em:02d}"

    # ------------------------------------------------------------------
    # Deteccion de software
    # ------------------------------------------------------------------

    def _find_acrobat(self):
        for p in [
            r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
            r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
            r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
            r"C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
        ]:
            if os.path.exists(p):
                return p
        return None

    def _find_gdrive(self):
        for hkey, subkey in [
            (winreg.HKEY_CURRENT_USER,  r"SOFTWARE\Google\DriveFS"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Google\DriveFS"),
        ]:
            try:
                key = winreg.OpenKey(hkey, subkey)
                for val in ("MountPoint", "Path", "RootPath"):
                    try:
                        v, _ = winreg.QueryValueEx(key, val)
                        if v and Path(v).exists():
                            winreg.CloseKey(key)
                            return str(v)
                    except FileNotFoundError:
                        pass
                winreg.CloseKey(key)
            except FileNotFoundError:
                pass

        user = os.environ.get("USERNAME", "")
        for p in [
            Path(f"C:/Users/{user}/Google Drive"),
            Path(f"C:/Users/{user}/Mi unidad"),
            Path(f"C:/Users/{user}/My Drive"),
            Path("G:/My Drive"), Path("G:/"),
        ]:
            if p.exists() and p.is_dir():
                return str(p)
        return None

    def _find_onedrive(self):
        od = os.environ.get("OneDrive", "")
        if od and Path(od).exists():
            return od
        user = os.environ.get("USERNAME", "")
        for p in [
            Path(f"C:/Users/{user}/OneDrive"),
            Path(f"C:/Users/{user}/OneDrive - Personal"),
        ]:
            if p.exists() and p.is_dir():
                return str(p)
        return None

    # ------------------------------------------------------------------
    # Icono de bandeja
    # ------------------------------------------------------------------

    def _make_icon(self, active=False):
        size = 64
        img  = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        d    = ImageDraw.Draw(img)
        bg   = (16, 185, 129) if active else (75, 85, 99)
        d.ellipse([2, 2, size-2, size-2], fill=bg)
        d.rectangle([14, 26, 50, 42], fill="white")
        d.rectangle([20, 16, 44, 28], fill="white")
        d.rectangle([20, 36, 44, 52], fill=(220, 220, 220))
        d.line([24, 41, 40, 41], fill=(150, 150, 150), width=1)
        d.line([24, 46, 36, 46], fill=(150, 150, 150), width=1)
        return img

    def _update_tray_icon(self):
        if self.tray_icon:
            self.tray_icon.icon = self._make_icon(self.is_watching)

    # ------------------------------------------------------------------
    # Notificaciones
    # ------------------------------------------------------------------

    def _notify(self, title, msg, cooldown_key=None, cooldown_secs=4):
        """Muestra notificacion de bandeja. cooldown_key evita spam del mismo tipo."""
        if cooldown_key:
            now = time.time()
            if now - self._notify_times.get(cooldown_key, 0) < cooldown_secs:
                return
            self._notify_times[cooldown_key] = now

        def _do():
            try:
                if self.tray_icon:
                    self.tray_icon.notify(msg, title)
            except Exception:
                pass
        threading.Thread(target=_do, daemon=True).start()

    def _on_detected(self, filename, rule_name=""):
        if self.config["notify_detect"]:
            loc = f" — {rule_name}" if rule_name else ""
            self._notify(
                f"PDF detectado{loc}",
                filename,
                cooldown_key=f"detect_{rule_name}",
                cooldown_secs=3,
            )

    def _on_printed(self, filename, printer, status, rule_name=""):
        if status == "OK":
            self.config["printed_today"] = self.config["printed_today"] + 1
            self.config["printed_total"] = self.config["printed_total"] + 1
            self.config["last_file"]     = filename
            if self.config["notify_print"]:
                loc = f" — {rule_name}" if rule_name else ""
                self._notify(
                    f"Impreso{loc}",
                    f"{filename}\n{printer}",
                    cooldown_key=f"print_{rule_name}",
                    cooldown_secs=2,
                )
        elif self.config.get("notify_error", True):
            self._notify(
                "Error al imprimir",
                f"{filename}\n{status}",
                cooldown_key=f"error_{rule_name}",
                cooldown_secs=10,
            )
        self._save_history(filename, printer, status, rule_name)
        if self.root:
            self.root.after(0, self._update_counter_ui)
        if self.widget:
            try:
                self.widget.win.after(0, self.widget.refresh)
            except Exception:
                pass

    def _save_history(self, filename, printer, status, rule_name=""):
        try:
            CONFIG_DIR.mkdir(parents=True, exist_ok=True)
            ts   = time.strftime("%Y-%m-%d %H:%M:%S")
            rule = f" | {rule_name}" if rule_name else ""
            line = f"{ts} | {filename} | {printer}{rule} | {status}\n"
            with open(HISTORY_FILE, "a", encoding="utf-8") as f:
                f.write(line)
        except Exception:
            pass

    def _update_counter_ui(self):
        today = self.config["printed_today"]
        total = self.config["printed_total"]
        try:
            if self._lbl_hoy:
                self._lbl_hoy.config(text=str(today))
            if self._lbl_total:
                self._lbl_total.config(text=str(total))
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Widget flotante
    # ------------------------------------------------------------------

    def _toggle_widget(self, icon=None, item=None):
        if self.root:
            self.root.after(0, self._do_toggle_widget)

    def _do_toggle_widget(self):
        if not self.widget:
            self.widget = FloatingWidget(self)
        if self.widget.is_visible():
            self.widget.hide()
        else:
            self.widget.show()
            self.config["widget_visible"] = True

    # ------------------------------------------------------------------
    # Bandeja del sistema
    # ------------------------------------------------------------------

    def _setup_tray(self):
        menu = pystray.Menu(
            pystray.MenuItem("Abrir configuracion", self._tray_show, default=True),
            pystray.MenuItem(
                lambda _: "Ocultar widget" if (self.widget and self.widget.is_visible())
                          else "Mostrar widget",
                self._toggle_widget,
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem(
                lambda _: "Detener monitoreo" if self.is_watching else "Iniciar monitoreo",
                self._toggle_watching,
            ),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Apagar", self._quit_app),
        )
        self.tray_icon = pystray.Icon(
            APP_NAME, self._make_icon(self.is_watching), APP_NAME, menu
        )

    def _tray_show(self, icon=None, item=None):
        if self.root:
            self.root.after(0, self._do_show_window)

    def _do_show_window(self):
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    # ------------------------------------------------------------------
    # Construccion de la ventana principal
    # ------------------------------------------------------------------

    def _build_window(self):
        root = tk.Tk()
        self.root = root

        root.title(f"{APP_NAME} v{APP_VERSION} — Impresion automatica de PDFs")
        root.geometry("580x720")
        root.minsize(500, 540)
        root.resizable(True, True)
        root.configure(bg=C_BG)
        root.protocol("WM_DELETE_WINDOW", self._hide_window)

        try:
            ico = self._make_icon(self.is_watching).resize((32, 32))
            self._tk_icon = ImageTk.PhotoImage(ico)
            root.iconphoto(True, self._tk_icon)
        except Exception:
            pass

        self._v_wait           = tk.IntVar(value=self.config["wait_seconds"])
        self._v_autostart      = tk.BooleanVar(value=self.config["autostart"])
        self._v_notify_detect  = tk.BooleanVar(value=self.config["notify_detect"])
        self._v_notify_print   = tk.BooleanVar(value=self.config["notify_print"])
        self._v_notify_error   = tk.BooleanVar(value=self.config.get("notify_error", True))

        # Horario
        s_start = self.config["schedule_start"]
        s_end   = self.config["schedule_end"]
        self._v_schedule_enabled = tk.BooleanVar(value=self.config["schedule_enabled"])
        self._v_schedule_start_h = tk.StringVar(value=s_start.split(":")[0])
        self._v_schedule_start_m = tk.StringVar(value=s_start.split(":")[1])
        self._v_schedule_end_h   = tk.StringVar(value=s_end.split(":")[0])
        self._v_schedule_end_m   = tk.StringVar(value=s_end.split(":")[1])

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("TCombobox",
            fieldbackground=C_INPUT, background=C_CARD,
            foreground=C_TEXT, selectforeground=C_TEXT,
            selectbackground=C_ACCENT)
        style.map("TCombobox", fieldbackground=[("readonly", C_INPUT)])
        style.configure("Vertical.TScrollbar",
            background=C_CARD, troughcolor=C_BG,
            arrowcolor=C_MUTED, borderwidth=0)

        self._build_ui()

        if self.config["widget_visible"]:
            root.after(500, self._do_toggle_widget)

    def _build_ui(self):
        root = self.root

        # ── Header ──────────────────────────────────────────────────────
        hdr = tk.Frame(root, bg=C_CARD)
        hdr.pack(fill="x", side="top")

        try:
            logo_img         = self._make_icon(self.is_watching).resize((48, 48), Image.LANCZOS)
            self._logo_photo = ImageTk.PhotoImage(logo_img)
            self._logo_lbl   = tk.Label(hdr, image=self._logo_photo, bg=C_CARD)
            self._logo_lbl.pack(side="left", padx=(16, 8), pady=12)
        except Exception:
            self._logo_photo = None
            self._logo_lbl   = None

        hdr_text = tk.Frame(hdr, bg=C_CARD)
        hdr_text.pack(side="left", pady=12)
        tk.Label(hdr_text, text="AutoPrint",
                 font=("Segoe UI", 18, "bold"), bg=C_CARD, fg=C_TEXT).pack(anchor="w")
        tk.Label(hdr_text, text=f"v{APP_VERSION} — Impresion automatica de PDFs",
                 font=("Segoe UI", 8), bg=C_CARD, fg=C_MUTED).pack(anchor="w")

        # Contadores
        cnt_frame = tk.Frame(hdr, bg=C_CARD)
        cnt_frame.pack(side="right", padx=16, pady=8)

        hoy_frame = tk.Frame(cnt_frame, bg=C_CARD)
        hoy_frame.pack(side="left", padx=(0, 12))
        self._lbl_hoy = tk.Label(
            hoy_frame, text=str(self.config["printed_today"]),
            font=("Segoe UI", 20, "bold"), bg=C_CARD, fg=C_SUCCESS)
        self._lbl_hoy.pack()
        tk.Label(hoy_frame, text="hoy", font=("Segoe UI", 7),
                 bg=C_CARD, fg=C_MUTED).pack()

        total_frame = tk.Frame(cnt_frame, bg=C_CARD)
        total_frame.pack(side="left")
        self._lbl_total = tk.Label(
            total_frame, text=str(self.config["printed_total"]),
            font=("Segoe UI", 20, "bold"), bg=C_CARD, fg=C_TEXT)
        self._lbl_total.pack()
        tk.Label(total_frame, text="total", font=("Segoe UI", 7),
                 bg=C_CARD, fg=C_MUTED).pack()

        self._status_lbl = tk.Label(
            hdr, text="Detenido",
            font=("Segoe UI", 9, "bold"), bg=C_CARD, fg=C_DANGER)
        self._status_lbl.pack(side="right", padx=(0, 8))

        # ── Botones fijos al fondo ─────────────────────────────────────
        btn_row = tk.Frame(root, bg=C_BG, pady=10)
        btn_row.pack(side="bottom", fill="x", padx=16)

        self._toggle_btn = tk.Button(
            btn_row,
            text="Iniciar" if not self.is_watching else "Detener",
            font=("Segoe UI", 11, "bold"),
            bg=C_SUCCESS if not self.is_watching else C_DANGER,
            fg="white", relief="flat", cursor="hand2",
            pady=10, padx=18, command=self._toggle_ui)
        self._toggle_btn.pack(side="left", fill="x", expand=True, padx=(0, 5))

        tk.Button(btn_row, text="Widget",
                  font=("Segoe UI", 10), bg=C_CARD, fg=C_TEXT,
                  relief="flat", cursor="hand2", pady=10, padx=10,
                  command=self._do_toggle_widget).pack(side="left", padx=(0, 5))

        tk.Button(btn_row, text="Historial",
                  font=("Segoe UI", 10), bg=C_CARD, fg=C_TEXT,
                  relief="flat", cursor="hand2", pady=10, padx=10,
                  command=self._open_history).pack(side="left", padx=(0, 5))

        tk.Button(btn_row, text="Ocultar",
                  font=("Segoe UI", 10), bg=C_CARD, fg=C_TEXT,
                  relief="flat", cursor="hand2", pady=10, padx=10,
                  command=self._hide_window).pack(side="left", padx=(0, 5))

        tk.Button(btn_row, text="Apagar",
                  font=("Segoe UI", 10, "bold"), bg="#7f1d1d", fg="white",
                  relief="flat", cursor="hand2", pady=10, padx=10,
                  command=self._confirm_quit).pack(side="left")

        tk.Frame(root, bg=C_BORDER, height=1).pack(fill="x", side="bottom")

        # ── Canvas scrollable ──────────────────────────────────────────
        wrap  = tk.Frame(root, bg=C_BG)
        wrap.pack(side="top", fill="both", expand=True)

        canvas    = tk.Canvas(wrap, bg=C_BG, highlightthickness=0, bd=0)
        scrollbar = ttk.Scrollbar(wrap, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        inner    = tk.Frame(canvas, bg=C_BG)
        inner_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfig(inner_id, width=e.width))
        inner.bind("<Configure>",
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind_all("<MouseWheel>",
                        lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))

        body = tk.Frame(inner, bg=C_BG, padx=16, pady=10)
        body.pack(fill="both", expand=True)

        self._section(body, "Reglas de vigilancia",   self._sec_rules)
        self._section(body, "Horario de impresion",  self._sec_schedule)
        self._section(body, "Configuracion",          self._sec_config)
        self._section(body, "Notificaciones",         self._sec_notifications)
        self._section(body, "Sistema",                self._sec_system)

        # ── Log ───────────────────────────────────────────────────────
        log_card = tk.Frame(body, bg=C_SURFACE, padx=12, pady=10)
        log_card.pack(fill="x", pady=(0, 6))

        log_hdr = tk.Frame(log_card, bg=C_SURFACE)
        log_hdr.pack(fill="x")
        tk.Label(log_hdr, text="Registro de actividad",
                 font=("Segoe UI", 9, "bold"), bg=C_SURFACE, fg=C_MUTED).pack(side="left")
        tk.Button(log_hdr, text="Exportar",
                  font=("Segoe UI", 8), bg=C_SURFACE, fg=C_MUTED,
                  relief="flat", cursor="hand2",
                  command=self._export_log).pack(side="right", padx=(4, 0))
        tk.Button(log_hdr, text="Limpiar",
                  font=("Segoe UI", 8), bg=C_SURFACE, fg=C_MUTED,
                  relief="flat", cursor="hand2",
                  command=self._clear_log).pack(side="right")

        log_scroll = tk.Scrollbar(log_card)
        log_scroll.pack(side="right", fill="y")

        self._log_widget = tk.Text(
            log_card, height=7,
            bg="#0a0a1a", fg=C_MUTED,
            font=("Consolas", 9),
            state="disabled", relief="flat", bd=0,
            insertbackground=C_TEXT,
            yscrollcommand=log_scroll.set)
        self._log_widget.pack(fill="x", pady=(4, 0))
        log_scroll.config(command=self._log_widget.yview)

        # Tags de color por tipo de mensaje
        self._log_widget.tag_config("ok",      foreground=C_SUCCESS)
        self._log_widget.tag_config("error",   foreground=C_DANGER)
        self._log_widget.tag_config("warn",    foreground=C_WARNING)
        self._log_widget.tag_config("detect",  foreground="#60a5fa")
        self._log_widget.tag_config("start",   foreground="#a78bfa")
        self._log_widget.tag_config("stop",    foreground=C_MUTED)
        self._log_widget.tag_config("archive", foreground="#22d3ee")
        self._log_widget.tag_config("sep",     foreground="#2d3748")
        self._log_widget.tag_config("info",    foreground=C_MUTED)

        for entry in self._log_entries[-30:]:
            self._append_log_widget(entry)

        self._refresh_status_ui()

    # ------------------------------------------------------------------
    # Secciones
    # ------------------------------------------------------------------

    def _section(self, parent, title, fn):
        card = tk.Frame(parent, bg=C_SURFACE, padx=14, pady=10)
        card.pack(fill="x", pady=(0, 8))
        tk.Label(card, text=title, font=("Segoe UI", 10, "bold"),
                 bg=C_SURFACE, fg=C_TEXT).pack(anchor="w", pady=(0, 7))
        fn(card)

    def _sec_rules(self, parent):
        # Info de drives detectados
        info_row = tk.Frame(parent, bg=C_SURFACE)
        info_row.pack(fill="x", pady=(0, 8))

        drives = []
        if self.gdrive_path:
            drives.append(f"Google Drive: {self.gdrive_path}")
        if self.onedrive_path:
            drives.append(f"OneDrive: {self.onedrive_path}")
        if drives:
            tk.Label(info_row, text="  ".join(drives),
                     font=("Segoe UI", 8), bg=C_SURFACE, fg=C_SUCCESS).pack(side="left")
        else:
            tk.Label(info_row, text="No se detecto Google Drive ni OneDrive",
                     font=("Segoe UI", 8), bg=C_SURFACE, fg=C_MUTED).pack(side="left")

        if not self.acrobat_path:
            tk.Label(parent, text="Adobe Acrobat no encontrado — necesario para imprimir",
                     font=("Segoe UI", 8), bg=C_SURFACE, fg=C_DANGER).pack(anchor="w", pady=(0, 6))

        # Frame de tarjetas de reglas
        self._rules_frame = tk.Frame(parent, bg=C_SURFACE)
        self._rules_frame.pack(fill="x")
        self._render_rules()

        # Boton agregar
        tk.Button(parent, text="+ Agregar regla",
                  font=("Segoe UI", 9, "bold"),
                  bg=C_CARD, fg=C_TEXT,
                  relief="flat", cursor="hand2", padx=10, pady=6,
                  command=self._add_rule).pack(anchor="w", pady=(8, 0))

    def _render_rules(self):
        """Redibuja todas las tarjetas de reglas."""
        if not self._rules_frame:
            return
        for w in self._rules_frame.winfo_children():
            w.destroy()

        rules = self.config["rules"]
        if not rules:
            tk.Label(self._rules_frame,
                     text="Sin reglas — agrega al menos una para empezar",
                     font=("Segoe UI", 9), bg=C_SURFACE, fg=C_MUTED).pack(
                anchor="w", pady=4)
            return

        for i, rule in enumerate(rules):
            self._rule_card(self._rules_frame, i, rule)

    def _rule_card(self, parent, idx, rule):
        """Tarjeta visual para una regla."""
        card = tk.Frame(parent, bg=C_CARD, padx=10, pady=8)
        card.pack(fill="x", pady=(0, 5))

        # Fila principal: nombre + botones
        top = tk.Frame(card, bg=C_CARD)
        top.pack(fill="x")

        name = rule.get("name") or rule.get("folder", "Regla")
        tk.Label(top, text=f"  {name}",
                 font=("Segoe UI", 10, "bold"),
                 bg=C_CARD, fg=C_TEXT).pack(side="left")

        tk.Button(top, text="Editar",
                  font=("Segoe UI", 8), bg=C_SURFACE, fg=C_TEXT,
                  relief="flat", cursor="hand2", padx=6, pady=2,
                  command=lambda i=idx: self._edit_rule(i)).pack(side="right", padx=(4, 0))
        tk.Button(top, text="X",
                  font=("Segoe UI", 8), bg=C_SURFACE, fg=C_DANGER,
                  relief="flat", cursor="hand2", padx=6, pady=2,
                  command=lambda i=idx: self._delete_rule(i)).pack(side="right")

        # Detalles
        folder  = rule.get("folder", "")
        printer = rule.get("printer", "")
        archive = rule.get("archive_folder", "") if rule.get("archive_enabled") else ""

        folder_short = (folder[:52] + "...") if len(folder) > 55 else folder
        tk.Label(card, text=f"Carpeta: {folder_short}",
                 font=("Segoe UI", 8), bg=C_CARD, fg=C_MUTED).pack(anchor="w", pady=(3, 0))
        tk.Label(card, text=f"Impresora: {printer}",
                 font=("Segoe UI", 8), bg=C_CARD, fg=C_MUTED).pack(anchor="w")
        if archive:
            archive_short = (archive[:52] + "...") if len(archive) > 55 else archive
            tk.Label(card, text=f"Archivo: {archive_short}",
                     font=("Segoe UI", 8), bg=C_CARD, fg=C_SUCCESS).pack(anchor="w")

    def _add_rule(self):
        RuleDialog(self.root, self, rule=None, on_save=self._on_rule_saved)

    def _edit_rule(self, idx):
        rules = self.config["rules"]
        if idx < len(rules):
            RuleDialog(self.root, self, rule=rules[idx],
                       on_save=lambda r, i=idx: self._on_rule_edited(i, r))

    def _on_rule_saved(self, rule):
        rules = list(self.config["rules"])
        rules.append(rule)
        self.config["rules"] = rules
        self._render_rules()

    def _on_rule_edited(self, idx, rule):
        rules = list(self.config["rules"])
        rules[idx] = rule
        self.config["rules"] = rules
        self._render_rules()

    def _delete_rule(self, idx):
        rules = list(self.config["rules"])
        name  = rules[idx].get("name", f"Regla {idx+1}")
        if messagebox.askyesno("Eliminar regla",
                               f"Eliminar la regla '{name}'?",
                               parent=self.root):
            rules.pop(idx)
            self.config["rules"] = rules
            self._render_rules()

    def _sec_schedule(self, parent):
        def _time_spinbox(frame, var, from_, to_):
            return tk.Spinbox(frame, textvariable=var, from_=from_, to=to_,
                              width=3, font=("Segoe UI", 11, "bold"),
                              bg=C_INPUT, fg=C_TEXT, relief="flat",
                              buttonbackground=C_CARD, format="%02.0f",
                              command=self._save_schedule_from_ui)

        # Toggle "siempre" / "con horario"
        top = tk.Frame(parent, bg=C_SURFACE)
        top.pack(fill="x", pady=(0, 8))
        tk.Radiobutton(top, text="Siempre activo",
                       variable=self._v_schedule_enabled, value=False,
                       font=("Segoe UI", 9), bg=C_SURFACE, fg=C_TEXT,
                       selectcolor=C_BG, activebackground=C_SURFACE,
                       command=self._on_schedule_toggle).pack(side="left", padx=(0, 20))
        tk.Radiobutton(top, text="Solo en este horario:",
                       variable=self._v_schedule_enabled, value=True,
                       font=("Segoe UI", 9), bg=C_SURFACE, fg=C_TEXT,
                       selectcolor=C_BG, activebackground=C_SURFACE,
                       command=self._on_schedule_toggle).pack(side="left")

        # Fila de tiempos
        self._sched_row = tk.Frame(parent, bg=C_SURFACE)
        self._sched_row.pack(fill="x", pady=(0, 6))

        tk.Label(self._sched_row, text="De",
                 font=("Segoe UI", 9), bg=C_SURFACE, fg=C_MUTED).pack(side="left", padx=(0, 6))
        _time_spinbox(self._sched_row, self._v_schedule_start_h, 0, 23).pack(side="left")
        tk.Label(self._sched_row, text=":",
                 font=("Segoe UI", 11, "bold"), bg=C_SURFACE, fg=C_TEXT).pack(side="left")
        _time_spinbox(self._sched_row, self._v_schedule_start_m, 0, 59).pack(side="left", padx=(0, 14))

        tk.Label(self._sched_row, text="a",
                 font=("Segoe UI", 9), bg=C_SURFACE, fg=C_MUTED).pack(side="left", padx=(0, 6))
        _time_spinbox(self._sched_row, self._v_schedule_end_h, 0, 23).pack(side="left")
        tk.Label(self._sched_row, text=":",
                 font=("Segoe UI", 11, "bold"), bg=C_SURFACE, fg=C_TEXT).pack(side="left")
        _time_spinbox(self._sched_row, self._v_schedule_end_m, 0, 59).pack(side="left", padx=(0, 16))

        tk.Button(self._sched_row, text="Guardar",
                  font=("Segoe UI", 9), bg=C_CARD, fg=C_TEXT,
                  relief="flat", cursor="hand2", padx=8, pady=3,
                  command=self._save_schedule_from_ui).pack(side="left")

        tk.Label(parent,
                 text="Los PDFs que lleguen fuera del horario se guardan y se imprimen al inicio del siguiente periodo",
                 font=("Segoe UI", 8), bg=C_SURFACE, fg=C_MUTED,
                 wraplength=500, justify="left").pack(anchor="w", pady=(0, 4))

        # Aviso si autostart esta desactivado
        if not self.config.get("autostart", False):
            tk.Label(parent,
                     text="Consejo: activa 'Iniciar con Windows' (seccion Sistema) para que la app "
                          "arranque sola tras un reinicio y no pierda archivos pendientes",
                     font=("Segoe UI", 8), bg=C_SURFACE, fg=C_WARNING,
                     wraplength=500, justify="left").pack(anchor="w", pady=(0, 8))
        else:
            tk.Frame(parent, bg=C_SURFACE, height=4).pack()

        # Estado de pendientes
        pend_row = tk.Frame(parent, bg=C_SURFACE)
        pend_row.pack(fill="x")

        self._lbl_pending = tk.Label(pend_row, text="Sin archivos pendientes",
                                     font=("Segoe UI", 8, "bold"),
                                     bg=C_SURFACE, fg=C_MUTED)
        self._lbl_pending.pack(side="left")

        tk.Button(pend_row, text="Imprimir pendientes ahora",
                  font=("Segoe UI", 8), bg=C_CARD, fg=C_TEXT,
                  relief="flat", cursor="hand2", padx=8, pady=3,
                  command=lambda: threading.Thread(
                      target=lambda: self.flush_pending_jobs(force=True),
                      daemon=True).start()
                  ).pack(side="left", padx=(10, 0))

        self._on_schedule_toggle()
        self._refresh_pending_label()

    def _on_schedule_toggle(self):
        enabled = self._v_schedule_enabled.get()
        state   = "normal" if enabled else "disabled"
        for child in self._sched_row.winfo_children():
            try:
                child.config(state=state)
            except Exception:
                pass
        self._save_schedule_from_ui()

    def _sec_notifications(self, parent):
        options = [
            ("Notificar cuando se detecta un PDF nuevo",  self._v_notify_detect, "notify_detect"),
            ("Notificar cuando se envia a imprimir",      self._v_notify_print,  "notify_print"),
            ("Notificar si ocurre un error al imprimir",  self._v_notify_error,  "notify_error"),
        ]
        for i, (text, var, key) in enumerate(options):
            tk.Checkbutton(parent, text=text,
                           variable=var, font=("Segoe UI", 9),
                           bg=C_SURFACE, fg=C_TEXT, selectcolor=C_BG,
                           activebackground=C_SURFACE, activeforeground=C_TEXT,
                           command=lambda k=key, v=var: self.config.__setitem__(k, v.get())
                           ).pack(anchor="w", pady=(0 if i == 0 else 3, 0))

        tk.Label(parent,
                 text="Las notificaciones tienen control de frecuencia para evitar spam",
                 font=("Segoe UI", 8), bg=C_SURFACE, fg=C_MUTED
                 ).pack(anchor="w", pady=(7, 0))

    def _sec_config(self, parent):
        row = tk.Frame(parent, bg=C_SURFACE)
        row.pack(fill="x")
        tk.Label(row, text="Espera antes de imprimir (seg):",
                 font=("Segoe UI", 9), bg=C_SURFACE, fg=C_MUTED).pack(side="left")
        tk.Spinbox(row, from_=1, to=60, textvariable=self._v_wait, width=5,
                   font=("Segoe UI", 10), bg=C_INPUT, fg=C_TEXT, relief="flat",
                   buttonbackground=C_CARD).pack(side="left", padx=(10, 0), ipady=3)

        if self.acrobat_path:
            txt, color = f"Adobe Acrobat: {Path(self.acrobat_path).name}", C_SUCCESS
        else:
            txt, color = "Adobe Acrobat NO encontrado", C_DANGER
        tk.Label(parent, text=txt, font=("Segoe UI", 8),
                 bg=C_SURFACE, fg=color).pack(anchor="w", pady=(7, 0))

    def _sec_system(self, parent):
        tk.Checkbutton(parent, text="Iniciar automaticamente con Windows",
                       variable=self._v_autostart, font=("Segoe UI", 10),
                       bg=C_SURFACE, fg=C_TEXT, selectcolor=C_BG,
                       activebackground=C_SURFACE, activeforeground=C_TEXT,
                       command=self._toggle_autostart).pack(side="left")

    # ------------------------------------------------------------------
    # Historial
    # ------------------------------------------------------------------

    def _open_history(self):
        win = tk.Toplevel(self.root)
        win.title("Historial de impresiones")
        win.geometry("640x420")
        win.configure(bg=C_BG)
        win.transient(self.root)

        tk.Label(win, text="Historial de impresiones",
                 font=("Segoe UI", 12, "bold"),
                 bg=C_BG, fg=C_TEXT).pack(pady=(14, 6), padx=16, anchor="w")

        stats = tk.Frame(win, bg=C_SURFACE, padx=14, pady=10)
        stats.pack(fill="x", padx=16, pady=(0, 10))
        for label, value, color in [
            ("Impresos hoy",  self.config["printed_today"], C_SUCCESS),
            ("Total impreso", self.config["printed_total"], C_TEXT),
        ]:
            col = tk.Frame(stats, bg=C_SURFACE)
            col.pack(side="left", padx=(0, 30))
            tk.Label(col, text=str(value), font=("Segoe UI", 24, "bold"),
                     bg=C_SURFACE, fg=color).pack()
            tk.Label(col, text=label, font=("Segoe UI", 8),
                     bg=C_SURFACE, fg=C_MUTED).pack()

        txt_frame = tk.Frame(win, bg=C_BG)
        txt_frame.pack(fill="both", expand=True, padx=16, pady=(0, 10))
        sb = tk.Scrollbar(txt_frame)
        sb.pack(side="right", fill="y")
        txt = tk.Text(txt_frame, font=("Consolas", 9),
                      bg="#0a0a1a", fg=C_MUTED, relief="flat",
                      state="disabled", yscrollcommand=sb.set)
        txt.pack(fill="both", expand=True)
        sb.config(command=txt.yview)

        try:
            if HISTORY_FILE.exists():
                with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                    lines = f.readlines()
                txt.config(state="normal")
                txt.insert("end", f"{'FECHA':<22} {'ARCHIVO':<35} {'IMPRESORA':<25} ESTADO\n")
                txt.insert("end", "-" * 95 + "\n")
                for line in reversed(lines[-200:]):
                    txt.insert("end", line)
                txt.config(state="disabled")
            else:
                txt.config(state="normal")
                txt.insert("end", "Sin historial todavia.")
                txt.config(state="disabled")
        except Exception:
            pass

        btn_row = tk.Frame(win, bg=C_BG)
        btn_row.pack(fill="x", padx=16, pady=(0, 12))

        def reset_counters():
            if messagebox.askyesno("Reiniciar contadores",
                                   "Reiniciar contadores de hoy y total a cero?",
                                   parent=win):
                self.config["printed_today"] = 0
                self.config["printed_total"] = 0
                self._update_counter_ui()
                win.destroy()

        tk.Button(btn_row, text="Reiniciar contadores",
                  font=("Segoe UI", 9), bg="#7f1d1d", fg="white",
                  relief="flat", cursor="hand2", padx=10, pady=6,
                  command=reset_counters).pack(side="left")
        tk.Button(btn_row, text="Cerrar",
                  font=("Segoe UI", 9), bg=C_CARD, fg=C_TEXT,
                  relief="flat", cursor="hand2", padx=10, pady=6,
                  command=win.destroy).pack(side="right")

    # ------------------------------------------------------------------
    # Helpers UI
    # ------------------------------------------------------------------

    def _hide_window(self):
        if self.root:
            self.root.withdraw()

    def _clear_log(self):
        try:
            self._log_widget.config(state="normal")
            self._log_widget.delete("1.0", "end")
            self._log_widget.config(state="disabled")
            self._log_entries.clear()
        except Exception:
            pass

    def _export_log(self):
        if not self._log_entries:
            messagebox.showinfo("Exportar log", "No hay entradas en el registro.", parent=self.root)
            return
        path = filedialog.asksaveasfilename(
            title="Guardar registro",
            defaultextension=".txt",
            filetypes=[("Texto", "*.txt"), ("Todos", "*.*")],
            initialfile=f"autoprint_log_{time.strftime('%Y%m%d_%H%M%S')}.txt",
            parent=self.root,
        )
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                f.write("\n".join(self._log_entries))
            self._log(f"Log exportado: {path}")
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=self.root)

    def _refresh_status_ui(self):
        if not self.root:
            return
        if self.is_watching:
            if self._status_lbl:
                self._status_lbl.config(text="Activo", fg=C_SUCCESS)
            if self._toggle_btn:
                self._toggle_btn.config(text="Detener", bg=C_DANGER)
        else:
            if self._status_lbl:
                self._status_lbl.config(text="Detenido", fg=C_DANGER)
            if self._toggle_btn:
                self._toggle_btn.config(text="Iniciar", bg=C_SUCCESS)
        try:
            if self._logo_lbl and self._logo_photo:
                new_img = self._make_icon(self.is_watching).resize((48, 48), Image.LANCZOS)
                self._logo_photo.paste(new_img)
                self._logo_lbl.config(image=self._logo_photo)
        except Exception:
            pass
        self._update_counter_ui()
        if self.widget:
            try:
                self.widget.win.after(0, self.widget.refresh)
            except Exception:
                pass

    def _log_tag_for(self, entry):
        low = entry.lower()
        if "error" in low:                                      return "error"
        if "aviso" in low:                                      return "warn"
        if "ok" in low or "impreso" in low:                     return "ok"
        if "copiado" in low or "archivado" in low or "eliminado" in low: return "archive"
        if "detectado" in low or "enviando" in low:             return "detect"
        if "iniciado" in low or "vigilando" in low:             return "start"
        if "detenido" in low:                                   return "stop"
        if entry.strip().startswith("─"):                       return "sep"
        return "info"

    def _append_log_widget(self, entry):
        if not self._log_widget:
            return
        try:
            self._log_widget.config(state="normal")
            tag = self._log_tag_for(entry)
            self._log_widget.insert("end", entry + "\n", tag)
            self._log_widget.see("end")
            self._log_widget.config(state="disabled")
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Log
    # ------------------------------------------------------------------

    def _log(self, msg):
        ts    = time.strftime("%H:%M:%S")
        entry = f"[{ts}] {msg}"
        self._log_entries.append(entry)
        if len(self._log_entries) > 300:
            self._log_entries = self._log_entries[-300:]
        print(entry)
        if self.root:
            self.root.after(0, self._append_log_widget, entry)

    # ------------------------------------------------------------------
    # Monitoreo multi-carpeta
    # ------------------------------------------------------------------

    def _save_from_ui(self):
        if self._v_wait:       self.config["wait_seconds"] = self._v_wait.get()
        if self._v_autostart:  self.config["autostart"]    = self._v_autostart.get()

    def _toggle_ui(self):
        self._save_from_ui()
        self._toggle_watching()
        self._refresh_status_ui()

    def _toggle_watching(self, icon=None, item=None):
        if self.is_watching:
            self._stop_watching()
        else:
            self._start_watching()

    def _start_watching(self):
        rules = self.config["rules"]
        if not rules:
            self._show_error("No hay reglas configuradas.\nAgrega al menos una regla.")
            return False
        if not self.acrobat_path:
            self._show_error("Adobe Acrobat Reader no encontrado.")
            return False

        started = 0
        for rule in rules:
            folder  = rule.get("folder", "")
            printer = rule.get("printer", "")
            if not folder or not printer:
                continue
            if not os.path.exists(folder):
                self._log(f"AVISO: carpeta no existe, regla omitida: {folder}")
                continue

            archive_enabled = rule.get("archive_enabled", False)
            archive_folder  = rule.get("archive_folder", "")
            if archive_enabled and archive_folder and not os.path.exists(archive_folder):
                self._log(f"AVISO: carpeta de archivo no existe, archivo desactivado: {archive_folder}")
                archive_enabled = False

            rule_name = rule.get("name") or folder.split("/")[-1].split("\\")[-1]
            handler = PDFHandler(
                printer         = printer,
                acrobat_path    = self.acrobat_path,
                wait_seconds    = self.config["wait_seconds"],
                log_fn          = self._log,
                archive_enabled = archive_enabled,
                archive_folder  = archive_folder,
                on_detected_fn  = self._on_detected,
                on_printed_fn   = self._on_printed,
                rule_name       = rule_name,
                print_queue     = self._print_queue,
                schedule_fn     = self._is_in_schedule,
                on_pending_fn   = self._add_pending_job,
            )
            obs = Observer()
            obs.schedule(handler, folder, recursive=False)
            obs.start()
            self._observers.append(obs)
            self._log(f"[{rule_name}] Vigilando: {folder} -> {printer}")
            started += 1

        if started == 0:
            self._show_error("Ninguna regla valida para iniciar.")
            return False

        self.is_watching      = True
        self.config["active"] = True
        self._touch_last_seen()   # marcar ahora como "ultima vez activo"
        self._log(f"─── Sesion iniciada  {time.strftime('%H:%M:%S')} — {started} regla(s) activa(s) ───")
        self._update_tray_icon()
        return True

    def _stop_watching(self):
        for obs in self._observers:
            try:
                obs.stop()
                obs.join()
            except Exception:
                pass
        self._observers.clear()
        self.is_watching      = False
        self.config["active"] = False
        self._touch_last_seen()   # guardar "la ultima vez activo" para el escaneo de arranque
        self._log(f"─── Sesion detenida  {time.strftime('%H:%M:%S')} ───")
        self._update_tray_icon()

    def _show_error(self, msg):
        if self.root and self.root.winfo_exists():
            messagebox.showerror("AutoPrint", msg, parent=self.root)
        else:
            print(f"[ERROR] {msg}")

    # ------------------------------------------------------------------
    # Inicio automatico con Windows
    # ------------------------------------------------------------------

    def _toggle_autostart(self):
        enabled = self._v_autostart.get()
        if not self._set_autostart(enabled):
            messagebox.showerror("Error",
                "No se pudo modificar el inicio automatico.\n"
                "Intenta como administrador.", parent=self.root)
            self._v_autostart.set(not enabled)
        else:
            self.config["autostart"] = enabled

    def _set_autostart(self, enable):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, STARTUP_KEY, 0,
                                 winreg.KEY_SET_VALUE)
            if enable:
                cmd = (f'"{sys.executable}"' if getattr(sys, "frozen", False)
                       else f'"{sys.executable}" "{os.path.abspath(__file__)}"')
                winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, cmd)
            else:
                try:
                    winreg.DeleteValue(key, APP_NAME)
                except FileNotFoundError:
                    pass
            winreg.CloseKey(key)
            return True
        except Exception as e:
            print(f"Error autostart: {e}")
            return False

    # ------------------------------------------------------------------
    # Salir
    # ------------------------------------------------------------------

    def _confirm_quit(self):
        if messagebox.askyesno("Apagar AutoPrint",
                               "Apagar AutoPrint completamente?\n\n"
                               "El monitoreo se detendra.",
                               parent=self.root, icon="warning"):
            self._quit_app()

    def _quit_app(self, icon=None, item=None):
        self._monitor_running = False
        self._stop_watching()
        if self.tray_icon:
            self.tray_icon.stop()
        if self.root:
            self.root.after(0, self._do_quit)

    def _do_quit(self):
        try:
            self.root.destroy()
        except Exception:
            pass
        sys.exit(0)

    # ------------------------------------------------------------------
    # Inicio
    # ------------------------------------------------------------------

    def run(self):
        import ctypes
        mutex = ctypes.windll.kernel32.CreateMutexW(None, False, "AutoPrintAppMutex_v1")
        if ctypes.windll.kernel32.GetLastError() == 183:
            messagebox.showinfo(APP_NAME,
                "AutoPrint ya esta en ejecucion.\nRevisa el icono en la bandeja del sistema.")
            sys.exit(0)

        self._setup_tray()

        if self.config["active"] and self.config["rules"]:
            self._start_watching()

        # Iniciar tray PRIMERO para que las notificaciones de arranque lleguen
        threading.Thread(target=self.tray_icon.run, daemon=True, name="TrayThread").start()
        time.sleep(0.8)  # dar tiempo al tray para registrarse en el sistema

        # Hilo monitor de horario
        self._monitor_running = True
        threading.Thread(target=self._schedule_monitor, daemon=True,
                         name="ScheduleMonitor").start()

        # Manejar archivos pendientes al arrancar
        self._handle_pending_on_startup()

        # Escanear carpetas por PDFs que llegaron con la app apagada
        threading.Thread(target=self._scan_missed_files, daemon=True,
                         name="StartupScan").start()

        self._build_window()
        self.root.mainloop()


# ===== PUNTO DE ENTRADA =====

if __name__ == "__main__":
    app = AutoPrintApp()
    app.run()
