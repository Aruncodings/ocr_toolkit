# 🖼️ Image to OCR Conversion Tool

A powerful desktop application built using Python and `CustomTkinter` that allows users to extract text from images using OCR, convert voice to text, and even read out recognized text with text-to-speech capabilities. Export results to PDF, TXT, or Word formats with ease.

## 🚀 Features

- ✅ Upload and process multiple image files for OCR.
- 📝 Display extracted text in a styled GUI textbox.
- 📄 Export OCR results to:
  - PDF
  - TXT
  - Word (.docx)
- 🗣️ Convert text to speech (TTS).
- 🎤 Voice to text recognition via microphone.
- 🌙 Light/Dark/System theme switching.
- 🔍 UI scaling for better accessibility.

## 🧰 Technologies Used

- `Python`
- `Tkinter` & `CustomTkinter`
- `Pillow` (PIL)
- `pytesseract`
- `speech_recognition`
- `pyttsx3`
- `reportlab`
- `python-docx`

## 🖥️ Prerequisites

Make sure you have the following installed:

1. Python 3.7 or higher
2. [Tesseract OCR](https://github.com/tesseract-ocr/tesseract)
   - Set `pytesseract.pytesseract.tesseract_cmd` in your code to the install path.

## 📦 Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/image-to-ocr-conversion.git
   cd image-to-ocr-conversion
