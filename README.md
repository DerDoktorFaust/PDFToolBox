# PDFToolBox

PDFToolBox allows the following functions:

1. Merge PDFs
2. OCR PDFs
3. Extract All Text
4. Extract Annotations
5. Convert Word documents to PDF

This comes with a GUI running on PyQT6 (PyQt5 version is in the source directory).

You can drag and drop files into the files widget.

Under available OCR languages, if you select nothing, it will default to English. Commented out code contains a subprocess method to get all of the user's installed languages and list them. The uncommented code just contains the languages I use. Feel free to reverse the comments in your own application, though I recommend just inputting your own language codes because otherwise the list gets really long.

You need the following installed on your computer:
1. tesseract 
2. Microsoft Word (if you want to convert .docx and .doc files to PDF)

The app runs, as noted above, on tesseract. If you need additional languages, add them through tesseract. 

App was tested on MacOS Ventura. 

Creating a distributable app seems impossible at this time. OCRmyPDF uses Pikepdf, which Pyinstaller cannot presently package. The way around this is to use subprocesses with the user having installed on their own system both tesseract and ocrmypdf. This release contains the code for subprocesses, but they are commented out. A user could use that, if they like.

If you do want an "app" version of this, you will need to run p2app, but use the -A flag. Py2app will also not package this app due to an issue either with p2app or the Pillow package. Nevertheless, if you use the -A flag and keep the python script files forever in the same location, you can use it like a regular app on your system.

##Update Log

#Version 1.1.1
-Added garbage collection of converted PDF files when merging is used (i.e. a Word document gets converted to PDF when its part of the merging process; this new PDF is now deleted); when simply converting a Word doc to PDF, the PDF is not deleted

#Version 1.1
-Added conversion of Word files to PDF files with a button
-Added ability to add Word files and automatically convert to PDF when merging

#Version 1.0.2
-Changed method of dropping files into drag-and-drop widget to prevent errors on Windows and Linux systems

#Version 1.0.1
-Added in non-subprocess calls to ocrmypdf
-Removed subprocess.run calls for getting installed languages

Version 1.0
-Initial Commit
