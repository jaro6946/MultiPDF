#MultiPDF
#Version 1.0


#These functions allows you to fill PDFs that have fillable fields using an excel spreadsheet.  
#
# If you need to fill 20 PDFs with different information, this module should work 
#   with some stipulations

#Limitations: You must open your completed PDFs with Mac Preview.  Adobe products will
#             not work.  
#             If you are getting errors regarding encryption, you may have to use
#             https://smallpdf.com/unlock-pdf to unlock the PDF


#Methods:
#get_field_keys(input_pdf_path) 
#   This makes a CSV template that can be imported into
#   excel that you can fill with your information.  This spread sheet will then be exported
#   as a CSV and then used as an argument in the write_fillable_pdf method.    

#write_fillable_pdfs(input_CSV, input_pdf_template)
#   This method will write your PDF files as specifde in the input_CSV


#! /usr/bin/python

import os
import pdfrw

ANNOT_KEY = '/Annots'
ANNOT_FIELD_KEY = '/T'
ANNOT_VAL_KEY = '/V'
ANNOT_RECT_KEY = '/Rect'
SUBTYPE_KEY = '/Subtype'
WIDGET_SUBTYPE_KEY = '/Widget'


def get_field_keys(input_pdf_path):
    #Loads PDF and exports formatted CSV template of PDF fillables

    template_pdf = pdfrw.PdfReader(input_pdf_path)
    annotations = template_pdf.pages[0][ANNOT_KEY]
    
    dict_fields_and_values={}

    for annotation in annotations:
        if annotation[SUBTYPE_KEY] == WIDGET_SUBTYPE_KEY:
            if annotation[ANNOT_FIELD_KEY]:
                key = annotation[ANNOT_FIELD_KEY][1:-1]
                if annotation[ANNOT_VAL_KEY]:
                    item = annotation[ANNOT_VAL_KEY][1:-1]
                else:
                    item=None
                dict_fields_and_values.update({key:item})
                   

    with open('PDF_Fields.csv', 'w', newline='') as csvfile:
        list_fields_and_values=list(dict_fields_and_values.items())
        csvfile.write('Name ending files: ,project, \n,Folder name: ,folder,\n,Folder Path: ,Current Directory,\n')
        csvfile.write('Name begining files:,PDF Name 1, PDF Name 2, PDF Name 3, Etc,\n')
        for fields_values in list_fields_and_values:
            fields_values=list(fields_values)
            if fields_values[1] is None:
                fields_values[1]=''
            csvfile.write('%s,%s \n' %(fields_values[0],fields_values[1]))
            

def write_fillable_pdfs(input_CSV, input_pdf_template):
    #Loads CSV template and creates pdfs


    #Import CSV file with form information and create a python list with that information
    with open(input_CSV) as CSV_Template:
        lines=CSV_Template.readlines()
        cell=[]
        row=[]
        table=[]
        for line in lines:
            for characters in line:

                if characters==',':
                    cell=''.join(cell)
                    if cell != '':
                        row.append(cell)
                    cell=[]
                elif characters=='\n':
                    table.append(row)
                    row=[]
                    cell=[]
                else:
                    cell.append(characters)

    file_ammount=len(table[3])-1
    
    #make list element equal sized based on the number of files requested to be created.  Fill in empty cells with ''
    for rows in table:
        while len(rows)<file_ammount+1:
            rows.append('')
        for cells in rows:
            if cells is None:
                cells=''
    
    
    

    #Repeats file path, folder name, project name in table
    #also checks for blank entries in file names
    #Changes Current directory text to file path

    for i, rows in enumerate(table):
        for j, cells in enumerate(rows[1:]):
            j=j+1
            
            if rows[j]=='Current Directory':
                rows[j]=os.getcwd()
            
            try:
                if rows[j+1]=='':
                    
                    
                    if rows[j+1]!=rows[j] and i<3:
                        rows[j+1]=rows[j]

                    elif i<4:
                        print('error, there is a blank where there shouldnt be')

                    else:
                        break
            except:
                continue


  

    

    #put pdf names into list
    PDF_Names=[]
    
    for i in range(len(table[1][1:])): 
        PDF_Names.append(table[1][i+1]+'_'+table[0][i+1])



    working_directory=os.getcwd()+'/'
    
    folder=['','']
    
    #Lets make some PDFs
    for i in range(file_ammount):
        
        #make directory for PDFs if needed
        folder[0]=table[1][i+1]
        
        if folder[0]!=folder[1] and not folder[0] in os.listdir(table[2][i+1]):
            os.mkdir(working_directory+folder[0])
        folder[1]=folder[0]

        #create PDF file paths
        destination_folder=working_directory+folder[0]+'/'
        file_name=table[3][i+1]
        name_ending_each_file=table[0][i+1]
        PDF_file_path=destination_folder+'/'+file_name+name_ending_each_file
    
        #load up pdf template
        template_pdf = pdfrw.PdfReader(input_pdf_template)
        annotations = template_pdf.pages[0][ANNOT_KEY]
        
        #create dictionary of form information
        data_table=[]
        for rows in table[4:]:
            data_table.append([rows[0],rows[i+1]])
        data_dict=dict(data_table)

        #Edit PDF template and make PDF
        for annotation in annotations:
            if annotation[SUBTYPE_KEY] == WIDGET_SUBTYPE_KEY:
                if annotation[ANNOT_FIELD_KEY]:
                    key = annotation[ANNOT_FIELD_KEY][1:-1]
                    if key in data_dict.keys():
                        annotation.update(
                            pdfrw.PdfDict(V='{}'.format(data_dict[key]))
                        )
                       
                        annotation.update(
                        pdfrw.PdfDict(AP='{}'.format({'/N': (144, 0)}))
                        )

                        
                        annotation.update(
                        pdfrw.PdfDict(DA='{}'.format('/Helv 0 Tf 0 g'))
                        )

        pdfrw.PdfWriter().write(PDF_file_path+'.pdf', template_pdf)




if __name__ == '__main__':
    #get_field_keys(INVOICE_TEMPLATE_PATH)
    #write_fillable_pdf(INVOICE_TEMPLATE_PATH, INVOICE_OUTPUT_PATH, data_dict)
    #write_fillable_pdf('PDF_Fields.csv','DA-31-unlocked.pdf')
    print('lol')