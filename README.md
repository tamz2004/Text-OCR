# Text OCR 

Text OCR  is a simple desktop application built using Python and Tkinter. It allows users to extract text from image files or Excel files and save the extracted text in various formats such as text file, PDF, DOCX, or Excel.

## Features

- Extract text from image files (PNG, JPG, JPEG, GIF) using pytesseract OCR.
- Extract text from Excel files (XLSX, XLS) using pandas library.
- Save extracted text in formats: text file, PDF, DOCX, or Excel.
- Simple and user-friendly GUI built with Tkinter.

## Requirements

- Python 3.x
- Tkinter (included in standard Python library)
- pytesseract
- pandas
- opencv-python
- fpdf
- python-docx
- openpyxl

## Installation

1. Clone the repository:

   ```bash
   git clone https://github.com/tamz2004/text-ocr.git
   
1. Install the required dependencies:

    ```bash
    pip install pytesseract pandas opencv-python fpdf python-docx openpyxl

2. Run the File

  ```bash
   python ocr.py
```
## Usage

1. Launch the application by running text_ocr_tool.py.
2. Click on the "Browse" button to select an image file or an Excel file.
3. The application will extract text from the selected file and display it in the text area.
4. Choose the desired format (text, PDF, DOCX, or Excel) from the save options dialog.
5. Optionally, choose the output directory for saving the file.
6. Click the "Save" button to save the extracted text in the selected format.    


## Contribution
Contributions are welcome! If you find any bugs or have suggestions for improvements, please open an issue or submit a pull request.

## License
This project is licensed under the MIT License - see the LICENSE
