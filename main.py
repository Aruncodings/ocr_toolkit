import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image
import pytesseract
import speech_recognition as sr
from reportlab.pdfgen import canvas
import pyttsx3
import os
import customtkinter as ctk
from docx import Document

pytesseract.pytesseract.tesseract_cmd = 'C:/Program Files/Tesseract-OCR/tesseract.exe'

# Global variable to store OCR text
ocr_text = ""

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configure window
        self.title("IMAGE TO OCR CONVERSION")
        self.geometry(f"{1100}x580")

        self.grid_columnconfigure(1, weight=1)
        self.grid_columnconfigure((2, 3), weight=0)
        self.grid_rowconfigure((0, 1, 2), weight=1)

        # Sidebar setup
        self.sidebar_frame = ctk.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)

        self.logo_label = ctk.CTkLabel(self.sidebar_frame, text="IMAGE TO OCR CONVERSION",
                                       font=ctk.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))

        # Upload image button
        self.upload_image_button = ctk.CTkButton(self.sidebar_frame, text="Upload Image",
                                                 command=self.upload_image)
        self.upload_image_button.grid(row=1, column=0, padx=20, pady=10)

        # Clear result button
        self.clear_result_button = ctk.CTkButton(self.sidebar_frame, text="Clear Result",
                                                 command=self.clear_result)
        self.clear_result_button.grid(row=2, column=0, padx=20, pady=10)

        # Appearance mode settings
        self.appearance_mode_label = ctk.CTkLabel(self.sidebar_frame, text="Appearance Mode:",
                                                  anchor="w")
        self.appearance_mode_label.grid(row=5, column=0, padx=20, pady=(10, 0))
        
        self.appearance_mode_optionemenu = ctk.CTkOptionMenu(self.sidebar_frame,
                                                             values=["Light", "Dark", "System"],
                                                             command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=6, column=0, padx=20, pady=(10, 10))

        # Scaling settings
        self.scaling_label = ctk.CTkLabel(self.sidebar_frame, text="UI Scaling:", anchor="w")
        self.scaling_label.grid(row=7, column=0, padx=20, pady=(10, 0))

        self.scaling_optionemenu = ctk.CTkOptionMenu(self.sidebar_frame,
                                                     values=["80%", "90%", "100%", "110%", "120%"],
                                                     command=self.change_scaling_event)
        self.scaling_optionemenu.grid(row=8, column=0, padx=20, pady=(10, 20))

        # Textbox for OCR results
        self.textbox = ctk.CTkTextbox(self, width=250)
        self.textbox.grid(row=0, column=1, padx=(20, 0), pady=(20, 0), sticky="nsew")

        # Tab view for options
        self.tabview = ctk.CTkTabview(self, width=250)
        self.tabview.grid(row=1, column=1, padx=(20, 0), pady=(20, 0), sticky="nsew")
        self.tabview.grid_rowconfigure(0, weight=1)
        self.tabview.grid_columnconfigure(0, weight=1)
        self.tabview.add("OCR Format")
        self.tabview.add("Voice Conversion")
        self.tabview.add("Voice to Text")

        # OCR format options
        self.radio_var_ocr_format = tk.IntVar(value=0)
        self.radio_button_pdf = ctk.CTkRadioButton(master=self.tabview.tab("OCR Format"),
                                                   variable=self.radio_var_ocr_format, value=1, text="PDF")
        self.radio_button_pdf.grid(row=0, column=0, padx=20, pady=(20, 10))
        
        self.radio_button_txt = ctk.CTkRadioButton(master=self.tabview.tab("OCR Format"),
                                                   variable=self.radio_var_ocr_format, value=2, text="TXT")
        self.radio_button_txt.grid(row=1, column=0, padx=20, pady=(10, 10))
        
        self.radio_button_word = ctk.CTkRadioButton(master=self.tabview.tab("OCR Format"),
                                                    variable=self.radio_var_ocr_format, value=3, text="Word")
        self.radio_button_word.grid(row=2, column=0, padx=20, pady=(10, 10))
        
        self.proceed_button_ocr_format = ctk.CTkButton(master=self.tabview.tab("OCR Format"),
                                                       text="Proceed", command=self.proceed_operation)
        self.proceed_button_ocr_format.grid(row=3, column=0, padx=20, pady=(10, 0))

        # Voice conversion options
        self.proceed_button_voice_conversion = ctk.CTkButton(master=self.tabview.tab("Voice Conversion"),
                                                             text="Proceed", command=self.convert_text_to_speech)
        self.proceed_button_voice_conversion.grid(row=0, column=0, padx=20, pady=(20, 0))

        # Recognize speech button
        self.recognize_speech_button = ctk.CTkButton(master=self.tabview.tab("Voice to Text"),
                                                     text="Recognize Speech", command=self.recognize_speech)
        self.recognize_speech_button.grid(row=0, column=0, padx=20, pady=(20, 0))

    def upload_image(self):
        file_paths = filedialog.askopenfilenames(title="Select Image(s)", filetypes=[("Image files", "*.png;*.jpg;*.jpeg")])
        if file_paths:
            global ocr_text
            ocr_text = ""
            for file_path in file_paths:
                img = Image.open(file_path)
                scanned_text = pytesseract.image_to_string(img)
                ocr_text += scanned_text + "\n"
            self.textbox.delete(1.0, tk.END)
            self.textbox.insert(tk.END, ocr_text)

    def proceed_operation(self):
        operation = self.radio_var_ocr_format.get()
        if operation == 1:
            self.generate_pdf_with_filenames()
        elif operation == 2:
            self.save_txt_file()
        elif operation == 3:
            self.save_word_file()

    def generate_pdf_with_filenames(self):
        global ocr_text
        if not ocr_text:
            messagebox.showerror("Error", "No OCR text available.")
            return

        pdf_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if pdf_path:
            try:
                c = canvas.Canvas(pdf_path)
                lines = ocr_text.split('\n')
                y_position = 800
                for line in lines:
                    c.drawString(10, y_position, line)
                    y_position -= 12
                c.save()
                messagebox.showinfo("Success", f"PDF saved successfully: {pdf_path}")
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred while saving PDF: {str(e)}")

    def save_txt_file(self):
        global ocr_text
        if not ocr_text:
            messagebox.showerror("Error", "No OCR text available.")
            return

        txt_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
        if txt_path:
            try:
                with open(txt_path, 'w') as txt_file:
                    txt_file.write(ocr_text)
                messagebox.showinfo("Success", "Text saved to TXT file successfully.")
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def save_word_file(self):
        global ocr_text
        if not ocr_text:
            messagebox.showerror("Error", "No OCR text available.")
            return

        word_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word files", "*.docx")])
        if word_path:
            try:
                doc = Document()
                doc.add_paragraph(ocr_text)
                doc.save(word_path)
                messagebox.showinfo("Success", f"Word document saved successfully: {word_path}")
            except Exception as e:
                messagebox.showerror("Error", f"An error occurred while saving Word document: {str(e)}")

    def change_appearance_mode_event(self, new_appearance_mode: str):
        ctk.set_appearance_mode(new_appearance_mode)

    def change_scaling_event(self, new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        ctk.set_widget_scaling(new_scaling_float)

    def recognize_speech(self):
        recognizer = sr.Recognizer()
        mic = sr.Microphone()
        self.recognize_speech_button.config(text="Listening...", state="disabled")

        with mic as source:
            recognizer.adjust_for_ambient_noise(source)
            audio = recognizer.listen(source)

        try:
            self.textbox.delete(1.0, tk.END)
            recognized_text = recognizer.recognize_google(audio)
            self.textbox.insert(tk.END, recognized_text)
            global ocr_text
            ocr_text = recognized_text
        except sr.UnknownValueError:
            messagebox.showerror("Error", "Could not understand audio")
        except sr.RequestError as e:
            messagebox.showerror("Error", f"Could not request results; {e}")
        finally:
            self.recognize_speech_button.config(text="Recognize Speech", state="normal")

    def convert_text_to_speech(self):
        global ocr_text
        if not ocr_text:
            messagebox.showerror("Error", "No OCR text available.")
            return

        engine = pyttsx3.init()
        engine.say(ocr_text)
        engine.runAndWait()

    def clear_result(self):
        global ocr_text
        ocr_text = ""
        self.textbox.delete(1.0, tk.END)


if __name__ == "__main__":
    app = App()
    app.mainloop()
