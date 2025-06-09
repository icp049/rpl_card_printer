# 📇 RPL Library Card Printer (Local and Network Versions)

A Windows-only desktop application for printing Regina Public Library barcode cards in single-card or triple keychain format. Built with [CustomTkinter](https://github.com/TomSchimansky/CustomTkinter), the app features a modern UI, barcode preview, printer selection, and direct-to-printer support for high-DPI card output.

---

## 🎯 Features

- ✅ Clean modern UI using CustomTkinter
- 📥 Input 14-digit RPL library numbers (validated)
- 🖨️ Local printer detection via `win32print`
- 🎞️ Barcode preview with number annotation
- 🔘 Two print formats:
  - **Single Card**
  - **Triple Keychain**
- 🖼️ Visual mode selector with images
- ⏳ Animated progress bar during printing
- 📎 Splash screen with loading animation
- 📦 Portable: easily packaged with PyInstaller for deployment

---


---

## 📦 Dependencies

Install via `pip`:

```bash
pip install customtkinter pillow python-barcode pywin32

if you wish to do an executable file. run this command in the root folder

FOR LOCAL VERSION
pyinstaller --onefile --windowed --icon=printer.ico `
  --add-data "printer.ico;." `
  --add-data "snip1.PNG;." `
  --add-data "snip2.PNG;." `
  --add-data "refresh.png;." `
  rpl_card_printer_local.py


FOR NETWORK VERSION
pyinstaller --onefile --windowed --icon=printer.ico `
  --add-data "printer.ico;." `
  --add-data "snip1.PNG;." `
  --add-data "snip2.PNG;." `
  rpl_card_printer_network.py 