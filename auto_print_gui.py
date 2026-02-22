#!/usr/bin/env python3
"""
AutoPrint GUI - Impresión automática de PDFs
Aplicación de bandeja del sistema para imprimir PDFs automáticamente.
"""

import os
import sys
import time
import json
import shutil
import threading
import winreg
import win32print
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import pystray
from PIL import Image, ImageDraw, ImageTk

# ===== CONSTANTES =====
APP_NAME    = "AutoPrint"
APP_VERSION = "1.0"
CONFIG_DIR  = Path(os.environ.get("APPDATA", Path.home())) / "AutoPrint"
CONFIG_FILE = CONFIG_DIR / "config.json"
STARTUP_KEY = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Run"

# Paleta de colores
C_BG       = "#1a1a2e"
C_SURFACE  = "#16213e"
C_CARD     = "#0f3460"
C_ACCENT   = "#e94560"
C_SUCCESS  = "#10b981"
C_DANGER   = "#ef4444"
C_WARNING  = "#f59e0b"
C_TEXT     = "#e2e8f0"
C_MUTED    = "#94a3b8"
C_INPUT    = "#1e2a45"
C_BORDER   = "#2d3748"


# ===== CONFIGURACIÓN PERSISTENTE =====

class Config:
    DEFAULTS = {
        "printer":         "",
        "folder":          "",
        "wait_seconds":    3,
        "autostart":       False,
        "active":          False,
        "archive_enabled": False,
        "archive_folder":  "",
    }

    def __init__(self):
        self._data = dict(self.DEFAULTS)
        self.load()

    def load(self):
        try:
            if CONFIG_FILE.exists():
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    self._data.update(json.load(f))
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


# ===== MANEJADOR DE ARCHIVOS PDF =====

class PDFHandler(FileSystemEventHandler):
    def __init__(self, printer, acrobat_path, wait_seconds, log_fn,
                 archive_enabled=False, archive_folder=""):
        super().__init__()
        self.printer         = printer
        self.acrobat_path    = acrobat_path
        self.wait_seconds    = wait_seconds
        self.log_fn          = log_fn
        self.archive_enabled = archive_enabled
        self.archive_folder  = archive_folder
        self._printed        = set()

    def on_created(self, event):
        if event.is_directory:
            return
        path = event.src_path
        if not path.lower().endswith(".pdf"):
            return

        name = Path(path).name
        self.log_fn(f"Nuevo PDF detectado: {name}")
        time.sleep(self.wait_seconds)

        if path not in self._printed:
            self.log_fn(f"Enviando a imprimir: {name}")
            try:
                subprocess.Popen([self.acrobat_path, "/t", path, self.printer])
                self._printed.add(path)
                self.log_fn(f"OK — Imprimiendo en: {self.printer}")

                # Mover a carpeta de archivo si está activado
                if self.archive_enabled and self.archive_folder:
                    threading.Thread(
                        target=self._move_to_archive,
                        args=(path,),
                        daemon=True
                    ).start()

            except Exception as e:
                self.log_fn(f"ERROR al imprimir: {e}")

    def _move_to_archive(self, src_path):
        """Copia el archivo a la carpeta local y luego lo elimina del Drive."""
        time.sleep(8)  # Esperar a que Acrobat termine de leer el PDF

        src  = Path(src_path)
        dest = Path(self.archive_folder) / src.name

        if dest.exists():
            ts   = time.strftime("%Y%m%d_%H%M%S")
            dest = Path(self.archive_folder) / f"{src.stem}_{ts}{src.suffix}"

        # ── Paso 1: copiar al destino local ──────────────────────────────
        copied = False
        for intento in range(1, 7):
            try:
                if not src.exists():
                    self.log_fn(f"Archivo ya no existe en origen: {src.name}")
                    return
                shutil.copy2(str(src), str(dest))
                copied = True
                break
            except Exception as e:
                self.log_fn(f"Copiando... intento {intento}/6 ({e})")
                time.sleep(5)

        if not copied:
            self.log_fn(f"ERROR: No se pudo copiar '{src.name}' al archivo local")
            return

        self.log_fn(f"Copiado a local: {dest.name}")

        # ── Paso 2: eliminar del Drive (reintentos por bloqueo de Drive) ─
        for intento in range(1, 11):
            try:
                src.unlink()
                self.log_fn(f"Eliminado del Drive: {src.name}")
                self.log_fn(f"Archivado en: {self.archive_folder}")
                return
            except PermissionError:
                # Drive aún tiene el archivo bloqueado mientras sincroniza
                time.sleep(6)
            except FileNotFoundError:
                # Ya fue eliminado de otra forma
                self.log_fn(f"Archivado en: {self.archive_folder}")
                return
            except Exception as e:
                self.log_fn(f"ERROR eliminando del Drive (intento {intento}/10): {e}")
                time.sleep(6)

        self.log_fn(
            f"AVISO: '{src.name}' fue copiado a local pero NO pudo eliminarse del Drive.\n"
            f"  Puedes borrarlo manualmente de: {src.parent}"
        )


# ===== APLICACIÓN PRINCIPAL =====

class AutoPrintApp:
    def __init__(self):
        self.config       = Config()
        self.observer     = None
        self.is_watching  = False
        self.tray_icon    = None
        self.root         = None
        self._log_entries = []

        # Referencias a widgets de estado
        self._status_lbl  = None
        self._toggle_btn  = None
        self._log_widget  = None

        # Variables tk
        self._v_printer         = None
        self._v_folder          = None
        self._v_wait            = None
        self._v_autostart       = None
        self._v_archive_enabled = None
        self._v_archive_folder  = None

        self.acrobat_path = self._find_acrobat()
        self.gdrive_path  = self._find_gdrive()

    # ------------------------------------------------------------------
    # Detección de software
    # ------------------------------------------------------------------

    def _find_acrobat(self):
        candidates = [
            r"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
            r"C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
            r"C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe",
            r"C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe",
        ]
        for p in candidates:
            if os.path.exists(p):
                return p
        return None

    def _find_gdrive(self):
        # Buscar en registro
        reg_checks = [
            (winreg.HKEY_CURRENT_USER,  r"SOFTWARE\Google\DriveFS"),
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Google\DriveFS"),
        ]
        for hkey, subkey in reg_checks:
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

        # Rutas comunes
        user = os.environ.get("USERNAME", "")
        for p in [
            Path(f"C:/Users/{user}/Google Drive"),
            Path(f"C:/Users/{user}/Mi unidad"),
            Path(f"C:/Users/{user}/My Drive"),
            Path("G:/My Drive"),
            Path("G:/"),
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

        bg = (16, 185, 129) if active else (75, 85, 99)
        d.ellipse([2, 2, size - 2, size - 2], fill=bg)

        # Cuerpo impresora
        d.rectangle([14, 26, 50, 42], fill="white")
        # Bandeja entrada
        d.rectangle([20, 16, 44, 28], fill="white")
        # Papel salida
        d.rectangle([20, 36, 44, 52], fill=(220, 220, 220))
        d.line([24, 41, 40, 41], fill=(150, 150, 150), width=1)
        d.line([24, 46, 36, 46], fill=(150, 150, 150), width=1)
        return img

    def _update_tray_icon(self):
        if self.tray_icon:
            self.tray_icon.icon = self._make_icon(self.is_watching)

    # ------------------------------------------------------------------
    # Bandeja del sistema
    # ------------------------------------------------------------------

    def _setup_tray(self):
        menu = pystray.Menu(
            pystray.MenuItem("Abrir configuracion", self._tray_show, default=True),
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
        """Llamado desde el hilo del tray — delega al hilo principal."""
        if self.root:
            self.root.after(0, self._do_show_window)

    def _do_show_window(self):
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    # ------------------------------------------------------------------
    # Construcción de la ventana
    # ------------------------------------------------------------------

    def _build_window(self):
        root = tk.Tk()
        self.root = root

        root.title(f"{APP_NAME} — Impresion automatica de PDFs")
        root.geometry("560x680")
        root.minsize(480, 500)
        root.resizable(True, True)
        root.configure(bg=C_BG)
        root.protocol("WM_DELETE_WINDOW", self._hide_window)

        # Icono de la ventana
        try:
            ico = self._make_icon(self.is_watching).resize((32, 32))
            self._tk_icon = ImageTk.PhotoImage(ico)
            root.iconphoto(True, self._tk_icon)
        except Exception:
            pass

        # Variables tk
        self._v_printer         = tk.StringVar(value=self.config["printer"])
        self._v_folder          = tk.StringVar(value=self.config["folder"])
        self._v_wait            = tk.IntVar(value=self.config["wait_seconds"])
        self._v_autostart       = tk.BooleanVar(value=self.config["autostart"])
        self._v_archive_enabled = tk.BooleanVar(value=self.config["archive_enabled"])
        self._v_archive_folder  = tk.StringVar(value=self.config["archive_folder"])

        # Estilo ttk
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

    def _build_ui(self):
        root = self.root

        # ── Header fijo ──────────────────────────────────────────────────
        hdr = tk.Frame(root, bg=C_CARD)
        hdr.pack(fill="x", side="top")

        # Logo (icono como imagen)
        try:
            logo_img = self._make_icon(self.is_watching).resize((48, 48), Image.LANCZOS)
            self._logo_photo = ImageTk.PhotoImage(logo_img)
            self._logo_lbl   = tk.Label(hdr, image=self._logo_photo, bg=C_CARD)
            self._logo_lbl.pack(side="left", padx=(16, 8), pady=12)
        except Exception:
            self._logo_photo = None
            self._logo_lbl   = None

        # Textos del header
        hdr_text = tk.Frame(hdr, bg=C_CARD)
        hdr_text.pack(side="left", pady=12)
        tk.Label(
            hdr_text, text="AutoPrint",
            font=("Segoe UI", 18, "bold"),
            bg=C_CARD, fg=C_TEXT
        ).pack(anchor="w")
        tk.Label(
            hdr_text, text="Impresion automatica de PDFs",
            font=("Segoe UI", 8),
            bg=C_CARD, fg=C_MUTED
        ).pack(anchor="w")

        # Badge de estado (derecha)
        self._status_lbl = tk.Label(
            hdr, text="● Detenido",
            font=("Segoe UI", 10, "bold"),
            bg=C_CARD, fg=C_DANGER
        )
        self._status_lbl.pack(side="right", padx=20)

        # ── Botones fijos en el fondo ────────────────────────────────────
        btn_row = tk.Frame(root, bg=C_BG, pady=12)
        btn_row.pack(side="bottom", fill="x", padx=16)

        self._toggle_btn = tk.Button(
            btn_row,
            text="▶  Iniciar" if not self.is_watching else "■  Detener",
            font=("Segoe UI", 11, "bold"),
            bg=C_SUCCESS if not self.is_watching else C_DANGER,
            fg="white", relief="flat", cursor="hand2",
            pady=10, padx=18,
            command=self._toggle_ui
        )
        self._toggle_btn.pack(side="left", fill="x", expand=True, padx=(0, 6))

        tk.Button(
            btn_row, text="⬇  Ocultar",
            font=("Segoe UI", 10),
            bg=C_CARD, fg=C_TEXT,
            relief="flat", cursor="hand2",
            pady=10, padx=12,
            command=self._hide_window
        ).pack(side="left", padx=(0, 6))

        tk.Button(
            btn_row, text="⏻  Apagar",
            font=("Segoe UI", 10, "bold"),
            bg="#7f1d1d", fg="white",
            relief="flat", cursor="hand2",
            pady=10, padx=12,
            command=self._confirm_quit
        ).pack(side="left")

        # ── Separador ────────────────────────────────────────────────────
        tk.Frame(root, bg=C_BORDER, height=1).pack(fill="x", side="bottom")

        # ── Canvas scrollable para el cuerpo ─────────────────────────────
        scroll_wrap = tk.Frame(root, bg=C_BG)
        scroll_wrap.pack(side="top", fill="both", expand=True)

        canvas = tk.Canvas(scroll_wrap, bg=C_BG, highlightthickness=0, bd=0)
        scrollbar = ttk.Scrollbar(scroll_wrap, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        # Frame interior que vive dentro del canvas
        inner = tk.Frame(canvas, bg=C_BG)
        inner_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        # Hacer que el inner frame use todo el ancho del canvas
        def _on_canvas_resize(event):
            canvas.itemconfig(inner_id, width=event.width)

        def _on_inner_resize(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        canvas.bind("<Configure>", _on_canvas_resize)
        inner.bind("<Configure>", _on_inner_resize)

        # Scroll con rueda del mouse
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        # ── Secciones dentro del inner ────────────────────────────────────
        body = tk.Frame(inner, bg=C_BG, padx=16, pady=12)
        body.pack(fill="both", expand=True)

        self._section(body, "Impresora", self._sec_printer)
        self._section(body, "Carpeta a vigilar (Drive)", self._sec_folder)
        self._section(body, "Carpeta de archivo local", self._sec_archive)
        self._section(body, "Configuracion", self._sec_config)
        self._section(body, "Sistema", self._sec_system)

        # ── Log ──────────────────────────────────────────────────────────
        log_card = tk.Frame(body, bg=C_SURFACE, padx=12, pady=10)
        log_card.pack(fill="x", pady=(0, 6))

        tk.Label(
            log_card, text="Registro de actividad",
            font=("Segoe UI", 9, "bold"),
            bg=C_SURFACE, fg=C_MUTED
        ).pack(anchor="w")

        log_scroll = tk.Scrollbar(log_card)
        log_scroll.pack(side="right", fill="y")

        self._log_widget = tk.Text(
            log_card, height=5,
            bg="#0a0a1a", fg=C_MUTED,
            font=("Consolas", 9),
            state="disabled", relief="flat", bd=0,
            insertbackground=C_TEXT,
            yscrollcommand=log_scroll.set
        )
        self._log_widget.pack(fill="x", pady=(4, 0))
        log_scroll.config(command=self._log_widget.yview)

        for entry in self._log_entries[-30:]:
            self._append_log_widget(entry)

        self._refresh_status_ui()

    # ------------------------------------------------------------------
    # Constructores de secciones
    # ------------------------------------------------------------------

    def _section(self, parent, title, content_fn):
        card = tk.Frame(parent, bg=C_SURFACE, padx=14, pady=12)
        card.pack(fill="x", pady=(0, 10))
        tk.Label(
            card, text=title,
            font=("Segoe UI", 10, "bold"),
            bg=C_SURFACE, fg=C_TEXT
        ).pack(anchor="w", pady=(0, 8))
        content_fn(card)

    def _sec_printer(self, parent):
        row = tk.Frame(parent, bg=C_SURFACE)
        row.pack(fill="x")

        printers = self._get_printers()
        cb = ttk.Combobox(
            row, textvariable=self._v_printer,
            values=printers, state="readonly",
            font=("Segoe UI", 10)
        )
        cb.pack(side="left", fill="x", expand=True)

        if not self._v_printer.get() and printers:
            self._v_printer.set(printers[0])

        def refresh():
            new = self._get_printers()
            cb["values"] = new
            if new and not self._v_printer.get():
                self._v_printer.set(new[0])

        self._styled_btn(row, "↺", refresh).pack(side="left", padx=(6, 0))

    def _sec_folder(self, parent):
        row = tk.Frame(parent, bg=C_SURFACE)
        row.pack(fill="x")

        tk.Entry(
            row, textvariable=self._v_folder,
            font=("Segoe UI", 10),
            bg=C_INPUT, fg=C_TEXT, relief="flat",
            insertbackground=C_TEXT
        ).pack(side="left", fill="x", expand=True, ipady=6, padx=(0, 6))

        self._styled_btn(row, "Examinar", self._browse_folder).pack(side="left", padx=(0, 4))

        if self.gdrive_path:
            tk.Button(
                row, text="Drive",
                font=("Segoe UI", 9, "bold"),
                bg="#1967d2", fg="white",
                relief="flat", cursor="hand2",
                padx=8, pady=4,
                command=lambda: self._v_folder.set(self.gdrive_path)
            ).pack(side="left")

        # Estado Google Drive
        info_row = tk.Frame(parent, bg=C_SURFACE)
        info_row.pack(fill="x", pady=(6, 0))
        if self.gdrive_path:
            txt   = f"Google Drive detectado: {self.gdrive_path}"
            color = C_SUCCESS
        else:
            txt   = "Google Drive no detectado en este equipo"
            color = C_MUTED
        tk.Label(info_row, text=txt, font=("Segoe UI", 8),
                 bg=C_SURFACE, fg=color).pack(anchor="w")

    def _sec_archive(self, parent):
        # Checkbox activar/desactivar
        top = tk.Frame(parent, bg=C_SURFACE)
        top.pack(fill="x", pady=(0, 8))

        tk.Checkbutton(
            top,
            text="Mover PDFs a carpeta local despues de imprimir",
            variable=self._v_archive_enabled,
            font=("Segoe UI", 9),
            bg=C_SURFACE, fg=C_TEXT,
            selectcolor=C_BG,
            activebackground=C_SURFACE, activeforeground=C_TEXT,
            command=self._on_archive_toggle
        ).pack(side="left")

        # Fila de ruta
        self._archive_row = tk.Frame(parent, bg=C_SURFACE)
        self._archive_row.pack(fill="x")

        self._archive_entry = tk.Entry(
            self._archive_row, textvariable=self._v_archive_folder,
            font=("Segoe UI", 10),
            bg=C_INPUT, fg=C_TEXT, relief="flat",
            insertbackground=C_TEXT,
            state="normal" if self._v_archive_enabled.get() else "disabled"
        )
        self._archive_entry.pack(side="left", fill="x", expand=True, ipady=6, padx=(0, 6))

        self._archive_browse_btn = self._styled_btn(
            self._archive_row, "Examinar", self._browse_archive
        )
        self._archive_browse_btn.pack(side="left", padx=(0, 4))

        self._archive_new_btn = tk.Button(
            self._archive_row, text="+ Nueva carpeta",
            font=("Segoe UI", 9, "bold"),
            bg=C_ACCENT, fg="white",
            relief="flat", cursor="hand2",
            padx=8, pady=4,
            command=self._create_archive_folder
        )
        self._archive_new_btn.pack(side="left")

        # Nota informativa
        self._archive_info = tk.Label(
            parent,
            text="Los PDFs se moveran automaticamente 8 seg. despues de imprimir",
            font=("Segoe UI", 8),
            bg=C_SURFACE, fg=C_MUTED
        )
        self._archive_info.pack(anchor="w", pady=(6, 0))

        # Estado inicial de los controles
        self._on_archive_toggle()

    def _on_archive_toggle(self):
        enabled = self._v_archive_enabled.get()
        state   = "normal" if enabled else "disabled"
        bg      = C_INPUT if enabled else C_SURFACE
        try:
            self._archive_entry.config(state=state, bg=bg)
            self._archive_browse_btn.config(state=state)
            self._archive_new_btn.config(state=state)
        except Exception:
            pass

    def _browse_archive(self):
        initial = self._v_archive_folder.get() or str(Path.home())
        folder  = filedialog.askdirectory(
            title="Seleccionar carpeta de archivo local",
            initialdir=initial
        )
        if folder:
            self._v_archive_folder.set(folder)

    def _create_archive_folder(self):
        """Abre un diálogo para escribir el nombre de la nueva carpeta."""
        dialog = tk.Toplevel(self.root)
        dialog.title("Crear nueva carpeta")
        dialog.geometry("420x180")
        dialog.resizable(False, False)
        dialog.configure(bg=C_BG)
        dialog.grab_set()
        dialog.transient(self.root)

        tk.Label(
            dialog, text="Crear carpeta de archivo",
            font=("Segoe UI", 12, "bold"),
            bg=C_BG, fg=C_TEXT
        ).pack(pady=(16, 4))

        tk.Label(
            dialog, text="Selecciona la ubicacion y nombre de la nueva carpeta:",
            font=("Segoe UI", 9), bg=C_BG, fg=C_MUTED
        ).pack()

        name_var = tk.StringVar(value="PDFs Impresos")
        entry = tk.Entry(
            dialog, textvariable=name_var,
            font=("Segoe UI", 11),
            bg=C_INPUT, fg=C_TEXT, relief="flat",
            insertbackground=C_TEXT
        )
        entry.pack(fill="x", padx=30, pady=10, ipady=7)
        entry.select_range(0, "end")
        entry.focus_set()

        def do_create():
            name = name_var.get().strip()
            if not name:
                return
            base = filedialog.askdirectory(
                title="Donde crear la carpeta",
                initialdir=str(Path.home()),
                parent=dialog
            )
            if not base:
                return
            new_path = Path(base) / name
            try:
                new_path.mkdir(parents=True, exist_ok=True)
                self._v_archive_folder.set(str(new_path))
                self._log(f"Carpeta creada: {new_path}")
                dialog.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"No se pudo crear la carpeta:\n{e}", parent=dialog)

        tk.Button(
            dialog, text="Crear carpeta",
            font=("Segoe UI", 10, "bold"),
            bg=C_SUCCESS, fg="white",
            relief="flat", cursor="hand2",
            padx=16, pady=8,
            command=do_create
        ).pack(pady=(0, 10))

        entry.bind("<Return>", lambda e: do_create())

    def _sec_config(self, parent):
        row = tk.Frame(parent, bg=C_SURFACE)
        row.pack(fill="x")

        tk.Label(
            row, text="Espera antes de imprimir (segundos):",
            font=("Segoe UI", 9), bg=C_SURFACE, fg=C_MUTED
        ).pack(side="left")

        tk.Spinbox(
            row, from_=1, to=60,
            textvariable=self._v_wait, width=5,
            font=("Segoe UI", 10),
            bg=C_INPUT, fg=C_TEXT, relief="flat",
            buttonbackground=C_CARD
        ).pack(side="left", padx=(10, 0), ipady=3)

        # Estado Acrobat
        info = tk.Frame(parent, bg=C_SURFACE)
        info.pack(fill="x", pady=(8, 0))
        if self.acrobat_path:
            name  = Path(self.acrobat_path).name
            txt   = f"Adobe Acrobat encontrado: {name}"
            color = C_SUCCESS
        else:
            txt   = "Adobe Acrobat NO encontrado — necesario para imprimir PDFs"
            color = C_DANGER
        tk.Label(info, text=txt, font=("Segoe UI", 8),
                 bg=C_SURFACE, fg=color).pack(anchor="w")

    def _sec_system(self, parent):
        row = tk.Frame(parent, bg=C_SURFACE)
        row.pack(fill="x")

        tk.Checkbutton(
            row,
            text="Iniciar automaticamente con Windows",
            variable=self._v_autostart,
            font=("Segoe UI", 10),
            bg=C_SURFACE, fg=C_TEXT,
            selectcolor=C_BG,
            activebackground=C_SURFACE, activeforeground=C_TEXT,
            command=self._toggle_autostart
        ).pack(side="left")

    # ------------------------------------------------------------------
    # Helpers UI
    # ------------------------------------------------------------------

    def _styled_btn(self, parent, text, cmd):
        return tk.Button(
            parent, text=text,
            font=("Segoe UI", 9),
            bg=C_CARD, fg=C_TEXT,
            relief="flat", cursor="hand2",
            padx=8, pady=4,
            command=cmd
        )

    def _hide_window(self):
        if self.root:
            self.root.withdraw()

    def _browse_folder(self):
        initial = self._v_folder.get() or str(Path.home())
        folder  = filedialog.askdirectory(
            title="Seleccionar carpeta a vigilar",
            initialdir=initial
        )
        if folder:
            self._v_folder.set(folder)

    def _get_printers(self):
        try:
            return [p[2] for p in win32print.EnumPrinters(2)]
        except Exception:
            return []

    def _refresh_status_ui(self):
        if not self.root:
            return
        if self.is_watching:
            if self._status_lbl:
                self._status_lbl.config(text="● Activo", fg=C_SUCCESS)
            if self._toggle_btn:
                self._toggle_btn.config(text="■  Detener", bg=C_DANGER)
        else:
            if self._status_lbl:
                self._status_lbl.config(text="● Detenido", fg=C_DANGER)
            if self._toggle_btn:
                self._toggle_btn.config(text="▶  Iniciar", bg=C_SUCCESS)
        # Actualizar logo
        try:
            if self._logo_lbl and self._logo_photo:
                new_img = self._make_icon(self.is_watching).resize((48, 48), Image.LANCZOS)
                self._logo_photo.paste(new_img)
                self._logo_lbl.config(image=self._logo_photo)
        except Exception:
            pass

    def _append_log_widget(self, entry):
        if not self._log_widget:
            return
        try:
            self._log_widget.config(state="normal")
            self._log_widget.insert("end", entry + "\n")
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
        if len(self._log_entries) > 200:
            self._log_entries = self._log_entries[-200:]
        print(entry)
        if self.root:
            self.root.after(0, self._append_log_widget, entry)

    # ------------------------------------------------------------------
    # Monitoreo
    # ------------------------------------------------------------------

    def _save_from_ui(self):
        if self._v_printer:
            self.config["printer"]         = self._v_printer.get()
        if self._v_folder:
            self.config["folder"]          = self._v_folder.get()
        if self._v_wait:
            self.config["wait_seconds"]    = self._v_wait.get()
        if self._v_autostart:
            self.config["autostart"]       = self._v_autostart.get()
        if self._v_archive_enabled:
            self.config["archive_enabled"] = self._v_archive_enabled.get()
        if self._v_archive_folder:
            self.config["archive_folder"]  = self._v_archive_folder.get()

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
        folder  = self.config["folder"]
        printer = self.config["printer"]

        if not folder:
            self._show_error("Selecciona una carpeta para vigilar.")
            return False
        if not os.path.exists(folder):
            self._show_error(f"La carpeta no existe:\n{folder}")
            return False
        if not printer:
            self._show_error("Selecciona una impresora.")
            return False
        if not self.acrobat_path:
            self._show_error(
                "Adobe Acrobat Reader no encontrado.\n"
                "Instala Adobe Acrobat Reader para poder imprimir PDFs."
            )
            return False

        archive_enabled = self.config["archive_enabled"]
        archive_folder  = self.config["archive_folder"]

        if archive_enabled and archive_folder and not os.path.exists(archive_folder):
            self._show_error(
                f"La carpeta de archivo no existe:\n{archive_folder}\n\n"
                "Crea la carpeta o elige otra en la seccion 'Carpeta de archivo local'."
            )
            return False

        handler = PDFHandler(
            printer         = printer,
            acrobat_path    = self.acrobat_path,
            wait_seconds    = self.config["wait_seconds"],
            log_fn          = self._log,
            archive_enabled = archive_enabled,
            archive_folder  = archive_folder,
        )
        self.observer = Observer()
        self.observer.schedule(handler, folder, recursive=False)
        self.observer.start()

        self.is_watching       = True
        self.config["active"]  = True

        self._log(f"Iniciado — Impresora: {printer}")
        self._log(f"Vigilando: {folder}")
        self._update_tray_icon()
        return True

    def _stop_watching(self):
        if self.observer:
            self.observer.stop()
            self.observer.join()
            self.observer = None
        self.is_watching      = False
        self.config["active"] = False
        self._log("Monitoreo detenido")
        self._update_tray_icon()

    def _show_error(self, msg):
        if self.root and self.root.winfo_exists():
            messagebox.showerror("AutoPrint — Error", msg, parent=self.root)
        else:
            print(f"[ERROR] {msg}")

    # ------------------------------------------------------------------
    # Inicio automático con Windows
    # ------------------------------------------------------------------

    def _toggle_autostart(self):
        enabled = self._v_autostart.get()
        ok      = self._set_autostart(enabled)
        if not ok:
            messagebox.showerror(
                "Error",
                "No se pudo modificar el inicio automatico.\n"
                "Intenta ejecutar el programa como administrador.",
                parent=self.root
            )
            self._v_autostart.set(not enabled)
        else:
            self.config["autostart"] = enabled

    def _set_autostart(self, enable: bool) -> bool:
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER, STARTUP_KEY, 0,
                winreg.KEY_SET_VALUE
            )
            if enable:
                if getattr(sys, "frozen", False):
                    cmd = f'"{sys.executable}"'
                else:
                    cmd = f'"{sys.executable}" "{os.path.abspath(__file__)}"'
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
        resp = messagebox.askyesno(
            "Apagar AutoPrint",
            "¿Deseas apagar AutoPrint completamente?\n\n"
            "El monitoreo se detendrá y el icono desaparecerá de la bandeja.",
            parent=self.root,
            icon="warning"
        )
        if resp:
            self._quit_app()

    def _quit_app(self, icon=None, item=None):
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
        # Prevenir múltiples instancias
        import ctypes
        mutex = ctypes.windll.kernel32.CreateMutexW(None, False, "AutoPrintAppMutex_v1")
        if ctypes.windll.kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
            # Traer la ventana existente al frente
            hwnd = ctypes.windll.user32.FindWindowW(None, None)
            messagebox.showinfo(
                APP_NAME,
                "AutoPrint ya está en ejecucion.\nRevisa el icono en la bandeja del sistema."
            )
            sys.exit(0)

        # Configurar bandeja
        self._setup_tray()

        # Reanudar si estaba activo al cerrar
        if self.config["active"] and self.config["folder"] and self.config["printer"]:
            self._start_watching()

        # Iniciar tray en hilo daemon
        tray_thread = threading.Thread(
            target=self.tray_icon.run,
            daemon=True, name="TrayThread"
        )
        tray_thread.start()

        # Ventana principal en hilo principal (bloquea hasta cerrar)
        self._build_window()
        self.root.mainloop()


# ===== PUNTO DE ENTRADA =====

if __name__ == "__main__":
    app = AutoPrintApp()
    app.run()
