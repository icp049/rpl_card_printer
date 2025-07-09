import ctypes
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

import customtkinter as ctk
import tkinter.messagebox as messagebox
from barcode import get_barcode_class
from barcode.writer import ImageWriter
from PIL import Image, ImageTk, ImageDraw, ImageFont, ImageWin
import win32print
import win32ui
from io import BytesIO
import win32con
import threading
import time
import gc
from datetime import datetime

def resource_path(relative_path):
    """ Get absolute path to resource for dev and for PyInstaller """
    import sys, os
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

class BarcodePrinterApp:
    def __init__(self, root):
        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")
        self.root = root

        self.root.title("RPL Library Card Printer (Test m.1)")
        self.root.geometry("700x760")
        self.root.resizable(False, False)

        self.input_var = ctk.StringVar()
        self.printer_var = ctk.StringVar()
        self.print_mode = ctk.StringVar(value="single")

        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        main_frame = ctk.CTkFrame(root)
        main_frame.grid(row=0, column=0, sticky="nsew")
        main_frame.grid_columnconfigure(0, weight=1)

        main_frame.grid_rowconfigure(0, weight=0)
        main_frame.grid_rowconfigure(1, weight=1)
        main_frame.grid_rowconfigure(2, weight=0)
        main_frame.grid_rowconfigure(3, weight=0)
        main_frame.grid_rowconfigure(4, weight=0)
        main_frame.grid_rowconfigure(5, weight=0)

        input_frame = ctk.CTkFrame(main_frame)
        input_frame.grid(row=0, column=0, sticky="ew", pady=10)
        input_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(input_frame, text="Library Account Number:").grid(row=0, column=0, padx=10, pady=10, sticky="e")

        entry_row = ctk.CTkFrame(input_frame, fg_color="transparent")
        entry_row.grid(row=0, column=1, sticky="ew", padx=(10, 10))
        entry_row.grid_columnconfigure(0, weight=1)

        self.entry = ctk.CTkEntry(entry_row, textvariable=self.input_var, height=40)
        self.entry.grid(row=0, column=0, sticky="ew")

        reset_icon = ctk.CTkImage(light_image=Image.open(resource_path("refresh.png")), size=(30, 30))

        reset_button = ctk.CTkButton(
            entry_row,
            image=reset_icon,
            text="",
            width=30,
            height=30,
            command=self.clear_input,
            fg_color="transparent",
            hover_color="#eeeeee",
            text_color="gray",
            font=("Arial", 30)
        )
        reset_button.grid(row=0, column=1, padx=(5, 0))

        ctk.CTkButton(input_frame, text="Generate Barcode", command=self.generate_barcode, width=200).grid(
            row=1, column=0, columnspan=2, pady=10
        )

        self.canvas = ctk.CTkCanvas(main_frame, bg="white", highlightthickness=0)
        self.canvas.grid(row=1, column=0, sticky="nsew")
        self.canvas.bind("<Configure>", self.resize_canvas)

        self.create_print_mode_selector(main_frame)

        self.printer_var = ctk.StringVar(value="Select Printer")

        printer_frame = ctk.CTkFrame(main_frame)
        printer_frame.grid(row=3, column=0, sticky="ew", pady=10)

        ctk.CTkLabel(printer_frame, text="Printer:").pack(anchor="center")

        dropdown_row = ctk.CTkFrame(printer_frame, fg_color="transparent")
        dropdown_row.pack(pady=5)

        server_name = r"\\printserver"
        try:
            printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_NAME, server_name, 2)
            self.printer_map = {
                p["pPrinterName"].split("\\")[-1]: p["pPrinterName"]
                for p in printers
                if "card" in p["pPrinterName"].lower()  # You can modify the filter as needed
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
            width=280
        )
        self.printer_dropdown.pack()

        print_button = ctk.CTkButton(
            main_frame,
            text="PRINT CARD",
            command=self.print_barcode,
            fg_color="#9622d4",
            width=200,
            height=50,
            font=("Arial", 16, "bold")
        )
        print_button.grid(row=4, column=0, pady=10)

        self.progress_bar = ctk.CTkProgressBar(main_frame, mode="indeterminate")
        self.progress_bar.grid(row=5, column=0, pady=(0, 20))
        self.progress_bar.grid_remove()
        
    def generate_barcode(self):
        number = self.input_var.get().strip()
        if not (number.isdigit() and len(number) == 14):
            messagebox.showerror("Input Error", "Please enter exactly 14 digits.")
            return

        try:
            wrapped_number = f"A{number}A"
            codabar = get_barcode_class('codabar')
            barcode = codabar(wrapped_number, writer=ImageWriter())
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
        confirm = messagebox.askyesno("Confirm Print", "Before proceeding, is there a card in the printer?")
        if not confirm:
            return

        self.progress_bar.grid()
        self.progress_bar.start()
        self.print_failed = False
        self.print_thread = threading.Thread(target=self._print_dispatch, daemon=True)
        self.print_thread.start()
        self.root.after(15000, self.check_print_timeout)

    def _print_dispatch(self):
        try:
            if self.print_mode.get() == "single":
                self.print_barcode_single()
            elif self.print_mode.get() == "triple":
                self.print_barcode_triple()
        except Exception as e:
            error_msg = str(e)
            self.root.after(0, lambda: messagebox.showerror("Print Error", error_msg))
        finally:
            self.root.after(0, self.progress_bar.stop)
            self.root.after(0, self.progress_bar.grid_remove)

    def print_barcode_single(self):
        if not hasattr(self, 'image'):
            self.root.after(0, lambda: messagebox.showerror("Print Error", "Generate the barcode first."))
            return

        printer_name = self.printer_map.get(self.printer_var.get(), self.printer_var.get())
        dpi = 300
        card_width = int(3.375 * dpi)   # 1011
        card_height = int(2.125 * dpi)  # 638

        # Create full-size white background
        card = Image.new("RGB", (card_width, card_height), "white")

        # Resize barcode image and center it
        barcode_resized = self.image.resize((600, 180))
        x = (card_width - 600) // 2
        y = (card_height - 180) // 2
        card.paste(barcode_resized, (x, y))

        card = card.convert("RGB")  # Ensure it's RGB format
        dib = ImageWin.Dib(card)

        try:
            hdc = win32ui.CreateDC()
            hdc.CreatePrinterDC(printer_name)

            hdc.StartDoc("CardPrint")
            hdc.StartPage()
            dib.draw(hdc.GetHandleOutput(), (0, 0, card_width, card_height))  # Full card draw
            hdc.EndPage()
            hdc.EndDoc()
        except Exception as e:
            try:
                hdc.AbortDoc()
            except:
               pass
            error_msg = f"Printing failed:\n{e}"
            self.root.after(0, lambda: self.prompt_retry(error_msg, self.print_barcode))
        finally:
            try:
                hdc.DeleteDC()
            except:
                pass
            gc.collect()
            time.sleep(1.5)
            with open("print_log.txt", "a") as f:
                f.write(f"{datetime.now()} - Printed to {printer_name} (Single Card)\\n")
            self.root.after(0, lambda: self.handle_print_success(f"Printed to {printer_name} (Single Card)."))

            
    def print_barcode_triple(self):
        if not hasattr(self, 'image'):
            self.root.after(0, lambda: messagebox.showerror("Print Error", "Generate the barcode first."))
            return

        printer_name = self.printer_map.get(self.printer_var.get(), self.printer_var.get())
        dpi = 300
        card_width = int(3.375 * dpi)
        card_height = int(2.125 * dpi)

        # Resize the barcode
        barcode_resized = self.image.resize((600, 180))

        hdc = None  # Safe default
        try:
            hdc = win32ui.CreateDC()
            hdc.CreatePrinterDC(printer_name)

            hdc.StartDoc("CardPrint")
            hdc.StartPage()

            # Setup barcode and text
            zone_height = card_height // 3
            dib = ImageWin.Dib(barcode_resized)

            font = win32ui.CreateFont({"name": "Arial", "height": 44, "weight": 700})
            hdc.SelectObject(font)
            header_text = "reginalibrary.ca | sasklibraries.ca"
            text_width, text_height = hdc.GetTextExtent(header_text)

            for i in range(3):
                zone_top = i * zone_height
                top = zone_top + (zone_height - 180 - text_height - 15) // 2 + text_height + 15
                left = (card_width - 600) // 2

                hdc.TextOut((card_width - text_width) // 2, top - 15 - text_height, header_text)
                dib.draw(hdc.GetHandleOutput(), (left, top, left + 600, top + 180))

            hdc.EndPage()
            hdc.EndDoc()

        except Exception as e:
            try:
                if hdc:
                    hdc.AbortDoc()
            except:
                pass
            error_msg = f"Printing failed:\n{e}"
            self.root.after(0, lambda: self.prompt_retry(error_msg, self.print_barcode))
        finally:
            try:
                if hdc:
                    hdc.DeleteDC()
            except:
                pass
            gc.collect()
            time.sleep(1.5)
            with open("print_log.txt", "a") as f:
                f.write(f"{datetime.now()} - Printed to {printer_name} (Triple)\n")
            self.root.after(0, lambda: self.handle_print_success(f"Printed to {printer_name} (Triple Keychain)."))



    def create_print_mode_selector(self, parent):
        self.mode_frame = ctk.CTkFrame(parent)
        self.mode_frame.grid(row=2, column=0, pady=10, sticky="ew")
        self.mode_frame.grid_columnconfigure((0, 1), weight=1)

        button_width = 130
        button_height = 200

        try:
            single_img = Image.open(resource_path("snip1.PNG")).resize((button_width, button_height))
        except:
            single_img = Image.new("RGB", (button_width, button_height), "gray")

        try:
            triple_img = Image.open(resource_path("snip2.PNG")).resize((button_width, button_height))
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

    def clear_input(self):
        self.input_var.set("")
        self.canvas.delete("all")

    def prompt_retry(self, message, retry_function):
        def ask_and_handle():
            self.progress_bar.stop()
            self.progress_bar.grid_remove()
            retry = messagebox.askretrycancel("Print Error", message)
            if retry:
                self.print_barcode()
        self.root.after(0, ask_and_handle)

    def handle_print_success(self, message):
        if getattr(self, "print_failed", False):
            return
        self.progress_bar.stop()
        self.progress_bar.grid_remove()
        messagebox.showinfo("Print Success", message)

    def check_print_timeout(self):
        if self.print_thread.is_alive():
            self.print_failed = True
            self.progress_bar.stop()
            self.progress_bar.grid_remove()
            messagebox.showerror("Print Timeout", "Printer is not responding. Please check the printer and try again.")

if __name__ == "__main__":
    root = ctk.CTk()
    icon_path = resource_path("printer.ico")
    root.iconbitmap(icon_path)
    app = BarcodePrinterApp(root)
    root.mainloop()


