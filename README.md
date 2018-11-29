MultiPDF
Version 1.0


These functions allows you to fill PDFs that have fillable fields using an excel spreadsheet.  This program utilizes the PDFRW library and this blog post got me started https://bostata.com/post/how_to_populate_fillable_pdfs_with_python/ 

If you need to fill 20 PDFs with different information, this module should work with some stipulations

Limitations: You must open your completed PDFs with Mac Preview and save them there to have a PDF that will work with Adobe.  You can also open them in Google Chrome and print them, though you will not be able to save them and work on them in Adobe.  

Adobe products will not work to open these files for reasons talked about here https://github.com/pmaupin/pdfrw/issues/84. If you are getting errors regarding encryption, you may have to use https://smallpdf.com/unlock-pdf to unlock the PDF before you use this program.


Methods:
get_field_keys(input_pdf_path) 
This makes a CSV template that can be imported into excel that you can fill with your information.  Export the spreadsheet as a CSV and then used as an argument in the write_fillable_pdf method.    

write_fillable_pdfs(input_CSV, input_pdf_template)
This method will write your PDF files as specifide in the input_CSV
