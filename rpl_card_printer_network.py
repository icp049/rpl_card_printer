import customtkinter as ctk
import tkinter.messagebox as messagebox
from barcode import get_barcode_class
from barcode.writer import ImageWriter
from PIL import Image, ImageTk, ImageDraw, ImageFont, ImageWin
import win32print
import win32ui
from io import BytesIO
import threading


class BarcodePrinterApp:
    def __init__(self, root):
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.root = root
        self.root.title("Codabar Barcode Printer")
        self.root.geometry("700x750")
        self.root.minsize(600, 600)

        self.input_var = ctk.StringVar()
        self.printer_var = ctk.StringVar()
        self.print_mode = ctk.StringVar(value="single")

        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        main_frame = ctk.CTkFrame(root)
        main_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)
        main_frame.grid_rowconfigure(3, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)

        input_frame = ctk.CTkFrame(main_frame)
        input_frame.grid(row=0, column=0, sticky="ew", pady=10)
        input_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(input_frame, text="Library Account Number:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
        self.entry = ctk.CTkEntry(input_frame, textvariable=self.input_var)
        self.entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        ctk.CTkButton(input_frame, text="Generate Barcode", command=self.generate_barcode).grid(
            row=1, column=0, columnspan=2, pady=10
        )

        self.canvas = ctk.CTkCanvas(main_frame, bg="white", highlightthickness=0)
        self.canvas.grid(row=1, column=0, pady=10, sticky="nsew")
        self.canvas.bind("<Configure>", self.resize_canvas)

        self.create_print_mode_selector(main_frame)

        printer_frame = ctk.CTkFrame(main_frame)
        printer_frame.grid(row=4, column=0, sticky="ew", pady=10)

        ctk.CTkLabel(printer_frame, text="Select Printer:").pack(anchor="center")

        dropdown_row = ctk.CTkFrame(printer_frame, fg_color="transparent")
        dropdown_row.pack(pady=5, fill="x", padx=30)

        try:
            server_name = r"\\printserver"
            printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_NAME, server_name, 2)
            self.printer_map = {
                p["pPrinterName"].split("\\")[-1]: p["pPrinterName"]
                for p in printers if "card printer" in p["pPrinterName"].lower()
            }
            printer_display_names = list(self.printer_map.keys())
        except Exception as e:
            messagebox.showerror("Printer Load Error", f"Could not load printers from printserver:\n{e}")
            self.printer_map = {}
            printer_display_names = []

        self.printer_dropdown = ctk.CTkOptionMenu(
            dropdown_row,
            variable=self.printer_var,
            values=printer_display_names,
            command=self.update_printer_status,
            width=280
        )
        self.printer_dropdown.pack(side="left", padx=(0, 20))

        self.status_icon = ctk.CTkLabel(dropdown_row, text="●", font=("Arial", 18))
        self.status_icon.pack(side="left")

        self.status_text = ctk.CTkLabel(dropdown_row, text="Ready", font=("Arial", 14, "bold"))
        self.status_text.pack(side="left")

        if printer_display_names:
            self.printer_var.set(printer_display_names[0])
            self.update_printer_status(printer_display_names[0])

        ctk.CTkButton(main_frame, text="PRINT CARD", command=self.print_barcode, fg_color="#9622d4", width=200, height=50,
                      font=("Arial", 16, "bold")).grid(row=5, column=0, pady=10)

        self.progress_bar = ctk.CTkProgressBar(main_frame, mode="indeterminate")
        self.progress_bar.grid(row=6, column=0, pady=(0, 20))
        self.progress_bar.grid_remove()

    def generate_barcode(self):
        number = self.input_var.get().strip()
        if not (number.isdigit() and len(number) == 14):
            messagebox.showerror("Input Error", "Please enter exactly 14 digits.")
            return

        try:
            wrapped_number = f"A{number}A"
            Codabar = get_barcode_class('codabar')
            barcode = Codabar(wrapped_number, writer=ImageWriter())
            buffer = BytesIO()
            barcode.write(buffer, options={"write_text": False})
            buffer.seek(0)

            barcode_img = Image.open(buffer).convert("RGB")
            try:
                font = ImageFont.truetype("arial.ttf", 60)
            except:
                font = ImageFont.load_default()

            bbox = font.getbbox(number)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
            barcode_width, barcode_height = barcode_img.size
            total_height = barcode_height + text_height + 50
            combined_img = Image.new("RGB", (barcode_width, total_height), "white")
            combined_img.paste(barcode_img, (0, 0))

            draw = ImageDraw.Draw(combined_img)
            text_x = (barcode_width - text_width) // 2
            draw.text((text_x, barcode_height + 10), number, font=font, fill="black")

            self.image = combined_img
            self.root.after(100, self.update_preview_image)

        except Exception as e:
            messagebox.showerror("Barcode Error", str(e))

    def update_preview_image(self):
        if not hasattr(self, 'image'):
            return
        canvas_width = self.canvas.winfo_width()
        if canvas_width <= 1:
            return
        ratio = canvas_width / self.image.width
        preview_height = int(self.image.height * ratio)
        self.tk_image = ImageTk.PhotoImage(self.image.resize((canvas_width, preview_height)))
        self.canvas.config(height=preview_height)
        self.canvas.delete("all")
        self.canvas.create_image(canvas_width // 2, preview_height // 2, image=self.tk_image)

    def resize_canvas(self, event):
        self.update_preview_image()

    def print_barcode(self):
        self.progress_bar.grid()
        self.progress_bar.start()
        self.root.after(100, self._print_dispatch)

    def _print_dispatch(self):
        if self.print_mode.get() == "single":
            self.print_barcode_single()
        elif self.print_mode.get() == "triple":
            self.print_barcode_triple()
        self.progress_bar.stop()
        self.progress_bar.grid_remove()

    def print_barcode_single(self):
        if not hasattr(self, 'image'):
            messagebox.showerror("Print Error", "Generate the barcode first.")
            return

        display_name = self.printer_var.get()
        printer_name = self.printer_map.get(display_name, display_name)

        dpi = 300
        card_width_px = int(2.125 * dpi)
        card_height_px = int(3.375 * dpi)

        target_width = 600
        target_height = 180
        image_resized = self.image.resize((target_width, target_height), Image.LANCZOS)

        left = (card_width_px - target_width) // 2
        top = 20
        right = left + target_width
        bottom = top + target_height

        hdc = win32ui.CreateDC()
        hdc.CreatePrinterDC(printer_name)
        hdc.StartDoc("Codabar Print - Single")
        hdc.StartPage()

        dib = ImageWin.Dib(image_resized)
        dib.draw(hdc.GetHandleOutput(), (left, top, right, bottom))

        hdc.EndPage()
        hdc.EndDoc()
        hdc.DeleteDC()

        messagebox.showinfo("Print Success", f"Printed to {printer_name} (Single Mode).")

    def print_barcode_triple(self):
        if not hasattr(self, 'image'):
            messagebox.showerror("Print Error", "Generate the barcode first.")
            return

        display_name = self.printer_var.get()
        printer_name = self.printer_map.get(display_name, display_name)

        dpi = 300
        card_width_px = int(2.125 * dpi)
        card_height_px = int(3.375 * dpi)
        zone_height = card_height_px // 3

        barcode_width = 630
        barcode_height = 160
        header_spacing = 15

        left = (card_width_px - barcode_width) // 2
        right = left + barcode_width

        image_resized = self.image.resize((barcode_width, barcode_height), Image.LANCZOS)
        dib = ImageWin.Dib(image_resized)

        hdc = win32ui.CreateDC()
        hdc.CreatePrinterDC(printer_name)
        hdc.StartDoc("Codabar Print - Triple")
        hdc.StartPage()

        font = win32ui.CreateFont({"name": "Arial", "height": 40, "weight": 700})
        hdc.SelectObject(font)

        header_text = "reginalibrary.ca | sasklibraries.ca"
        text_width, text_height = hdc.GetTextExtent(header_text)

        for i in range(3):
            zone_top = i * zone_height
            top = zone_top + (zone_height - barcode_height - text_height - header_spacing) // 2 + text_height + header_spacing
            bottom = top + barcode_height

            hdc.TextOut((card_width_px - text_width) // 2, top - header_spacing - text_height, header_text)
            dib.draw(hdc.GetHandleOutput(), (left, top, right, bottom))

        hdc.EndPage()
        hdc.EndDoc()
        hdc.DeleteDC()

        messagebox.showinfo("Print Success", f"Printed to {printer_name} (Triple Mode).")

    def resource_path(self, relative_path):
        """ Get absolute path to resource for dev and for PyInstaller """
        import sys, os
        if hasattr(sys, '_MEIPASS'):
            return os.path.join(sys._MEIPASS, relative_path)
        return os.path.join(os.path.abspath("."), relative_path)

    def create_print_mode_selector(self, parent):
        self.mode_frame = ctk.CTkFrame(parent)
        self.mode_frame.grid(row=2, column=0, pady=10, sticky="ew")
        self.mode_frame.grid_columnconfigure((0, 1), weight=1)

        button_width = 130
        button_height = 200

        try:
            single_img = Image.open(self.resource_path("snip1.PNG")).resize((button_width, button_height))
        except:
            single_img = Image.new("RGB", (button_width, button_height), "gray")

        try:
            triple_img = Image.open(self.resource_path("snip2.PNG")).resize((button_width, button_height))
        except:
            triple_img = Image.new("RGB", (button_width, button_height), "gray")

        self.mode_images = {
            "Single Card": ctk.CTkImage(light_image=single_img, size=(button_width, button_height)),
            "Triple Keychain": ctk.CTkImage(light_image=triple_img, size=(button_width, button_height)),
        }

        self.mode_buttons = {}

        for i, mode in enumerate(["Single Card", "Triple Keychain"]):
            sub_frame = ctk.CTkFrame(self.mode_frame, fg_color="transparent")
            sub_frame.grid(row=0, column=i, padx=20)

            btn = ctk.CTkButton(
                sub_frame,
                image=self.mode_images[mode],
                text="",
                width=button_width,
                height=button_height,
                fg_color="transparent",
                border_width=2,
                corner_radius=10,
                command=lambda m=mode: self.select_print_mode(m)
            )
            btn.pack()

            label = ctk.CTkLabel(sub_frame, text=mode, font=("Arial", 14))
            label.pack(pady=(5, 0))

            self.mode_buttons[mode] = btn

        self.select_print_mode("Single Card")

    def select_print_mode(self, mode):
        mode_map = {"Single Card": "single", "Triple Keychain": "triple"}
        self.print_mode.set(mode_map.get(mode, "single"))
        for m, btn in self.mode_buttons.items():
            if m == mode:
                btn.configure(border_color="skyblue", border_width=3)
            else:
                btn.configure(border_color="gray", border_width=1)

    def update_printer_status(self, selected_printer=None):
        display_name = selected_printer or self.printer_var.get()
        printer_name = self.printer_map.get(display_name, display_name)

        status_flags = {
            0x00000002: "Paused",
            0x00000004: "Error",
            0x00000008: "Deleting",
            0x00000010: "Paper Jam",
            0x00000020: "Paper Out",
            0x00000040: "Manual Feed",
            0x00000080: "Paper Problem",
            0x00000100: "Offline",
            0x00000200: "IO Active",
            0x00000400: "Busy",
            0x00000800: "Printing",
            0x00001000: "Output Bin Full",
            0x00002000: "Not Available",
            0x00004000: "Waiting",
            0x00008000: "Processing",
            0x00010000: "Initializing",
            0x00020000: "Warming Up",
            0x00040000: "Toner Low",
            0x00080000: "No Toner",
            0x00100000: "Page Punt",
            0x00200000: "User Intervention",
            0x00400000: "Out of Memory",
            0x00800000: "Door Open",
            0x01000000: "Server Unknown",
            0x02000000: "Power Save",
        }

        try:
            hPrinter = win32print.OpenPrinter(printer_name)
            info = win32print.GetPrinter(hPrinter, 2)
            status_code = info['Status']
            win32print.ClosePrinter(hPrinter)

            status_list = [desc for code, desc in status_flags.items() if status_code & code]
            status_text = ", ".join(status_list) if status_list else "Ready"

            if any(word in status_text.lower() for word in ["error", "offline", "jam", "punt"]):
                color = "red"
            elif any(word in status_text.lower() for word in ["paused", "waiting", "processing"]):
                color = "orange"
            else:
                color = "green"

            self.status_icon.configure(text="●", text_color=color)
            self.status_text.configure(text=status_text, text_color=color)

        except Exception:
            self.status_icon.configure(text="●", text_color="gray")
            self.status_text.configure(text="Unknown", text_color="gray")


if __name__ == "__main__":
    root = ctk.CTk()
    app = BarcodePrinterApp(root)
    root.mainloop()