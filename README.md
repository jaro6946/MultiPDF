Welcome to the README, like the wiki, but without pictures!

MultiPDF Version 1.0

This Python project allows you to fill PDFs that have fillable fields using an excel spreadsheet, a function know as mail merging. If you need to fill 20 PDFs with different information from a spreadsheet, this module should work with some stipulations.  It will create the the individual pdfs and one combined pdf for easy printing.

Limitations: You must open your completed PDFs with Mac Preview and save them there to have a PDF that will work with Adobe. You can also open them in Google Chrome and print them, though you will not be able to save them and work on them in Adobe.  The combined pdf that is created will only work with Mac Preview.

Adobe products will not work to open these files for reasons talked about here https://github.com/pmaupin/pdfrw/issues/84. If you are getting errors regarding encryption, you may have to use https://smallpdf.com/unlock-pdf to unlock the PDF before you use this program.


Usage

First step, make csv template:

python3 MultiPDF.py ExamplePDF.pdf

It makes a CSV template that can be imported into excel that you can fill with your information. Export or save the spreadsheet as a CSV


Next step, Edit your CSV with the info you want to put in your pdfs and save


Final step, make your pdfs:
python3 MultiPDF.py PDF_Fields.csv ExamplePDF.pdf
