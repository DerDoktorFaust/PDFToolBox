# PDFToolBox

PDFToolBox allows the following functions:

1. Merge PDFs
2. OCR PDFs
3. Extract All Text
4. Extract Annotations

This comes with a GUI running on PyQT6 (PyQt5 version is in the source directory).

You can drag and drop files into the files widget.

Under available OCR languages, if you select nothing, it will default to English

You need the following installed on your computer (I recommend using brew to install them):
1. ocrmypdf (brew install ocrmypdf)
2. tesseract (brew install tesseract)

The app runs, as noted above, on tesseract. If you need additional languages, add them through tesseract. 

App was tested on MacOS Ventura. 

Creating a distributable app seems impossible at this time. Using Pyinstaller, a .app cannot be run because this app uses subprocesses.
That is, to get the most speed out of tesseract, this app does not use pytesseract, but tesseract natively by running the command
on your computer. It uses commands from the above progams, ocrmypdf and tesseract, and runs them as if they were in a terminal.
Thus, Pyinstaller does not like this. I am currently searching for a workaround, but the standard methods have not worked
(i.e. using instead of subprocess.run, using subprocess.Popen; using subprocess.PIPE, etc).
