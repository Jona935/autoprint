# 🖨️ AutoPrint

**Impresión automática de PDFs desde Google Drive para Windows**

AutoPrint vigila una carpeta (Google Drive u otra) y envía automáticamente cualquier PDF nuevo a la impresora que elijas. Corre en segundo plano desde la bandeja del sistema.

---

## ✨ Funcionalidades

- 🖨️ **Impresión automática** — detecta PDFs nuevos y los imprime al instante
- ☁️ **Google Drive** — detecta la carpeta automáticamente
- 💾 **Archivo local** — mueve los PDFs impresos a una carpeta local para liberar espacio en Drive
- 🔔 **Bandeja del sistema** — corre en segundo plano sin interrumpir tu trabajo
- 💾 **Guarda configuración** — recuerda impresora, carpeta y ajustes entre sesiones
- 🚀 **Inicio con Windows** — opción para arrancar automáticamente al encender el equipo
- ⏻ **Control total** — inicia, detén o apaga la app desde la ventana o el icono de bandeja

---

## 📥 Descarga

Ve a [**Releases**](../../releases) y descarga el instalador de la última versión.

| Versión | Descarga | Fecha |
|---------|----------|-------|
| v1.0    | [AutoPrint_Setup_v1.0.exe](../../releases/tag/v1.0) | 2026-02-21 |

---

## 🚀 Instalación

1. Descarga `AutoPrint_Setup_v1.0.exe` desde [Releases](../../releases)
2. Ejecuta el instalador (doble click)
3. Sigue el asistente (elige si quieres icono en escritorio e inicio con Windows)
4. ¡Listo! La app se abre automáticamente al terminar

**Requisitos:**
- Windows 10 / 11
- Adobe Acrobat Reader instalado

---

## 🖥️ Uso

1. Selecciona tu **impresora** en el desplegable
2. Elige la **carpeta a vigilar** (botón "Drive" para autodetectar Google Drive)
3. Opcional: activa **Carpeta de archivo local** para mover PDFs a tu PC tras imprimir
4. Haz click en **▶ Iniciar**
5. Cierra la ventana con X — la app sigue corriendo en la **bandeja del sistema**

---

## 🛠️ Compilar desde fuente

```bash
pip install pystray Pillow watchdog pywin32 pyinstaller
python -m PyInstaller AutoPrint.spec
```

---

## 📄 Licencia

MIT
