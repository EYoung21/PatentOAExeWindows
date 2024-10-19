# OA Processor

## Overview
The OA Processor is a Python application designed to extract information from patent office actions (OAs) and generate summaries, while also providing functionality to download related patent PDFs and summarize their content using OpenAI's API.

## Prerequisites
Before running the application, ensure you have the following installed:

- Python (3.x)
- PyMuPDF: Install with `pip install PyMuPDF`
- Pillow: Install with `pip install Pillow`
- pytesseract: Install with `pip install pytesseract`
- Selenium: Install with `pip install selenium`
- python-dotenv: Install with `pip install python-dotenv`
- python-docx: Install with `pip install python-docx`
- Requests: Install with `pip install requests`
- OpenAI: Install with `pip install openai`

## Tesseract Installation
You need to have Tesseract OCR installed. Download it from [here](https://github.com/tesseract-ocr/tesseract) and ensure it is located in the default folder: C:\Program Files\Tesseract-OCR\tesseract.exe


## How to Build the Executable
If you haven't already built the executable, follow these steps:

1. Download Tesseract (as outlined above).
2. Open a terminal (Command Prompt or PowerShell).
3. Navigate to the directory containing your code.
4. Run the command:
    ```bash
    pyinstaller oa_processor.spec
    ```
5. Right-click on the `dist` folder that is created.
6. Select "Reveal in File Explorer".
7. Open the `dist` folder, copy the generated `.exe` file, and paste it wherever you want to run it. Double-click the `.exe` to run the application.

## Usage
Ensure your environment variables are set correctly, particularly `OPENAI_API_KEY`, in a `.env` file.  
Place the PDFs you want to process in the working directory.  
Run the executable, and it will process each PDF file, extracting necessary data and generating summaries.  
The processed data will be saved in a newly created directory named `DataCSV`, `Reference Summaries & PDFs`.

## Features
- Extracts application ID, reference number, due date, examiner name, and examiner phone number from patent OAs.
- Downloads related patent PDFs from Google Patents.
- Uses OpenAI API to generate summaries of the patent technologies.
- Outputs results to a CSV file and generates individual summary documents in `.docx` format.

## Contributing
Contributions are welcome! Please feel free to submit a pull request or create an issue if you find a bug or have a feature request.

## License
This project is licensed under the GNU General Public License - see the `LICENSE` file for details.
