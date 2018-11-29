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
#get_field_keys(input_pdf_template) 
#   This makes a CSV template that can be imported into
#   excel that you can fill with your information.  This spread sheet will then be exported
#   as a CSV and then used as an argument in the write_fillable_pdf method.    

#write_fillable_pdfs(input_CSV, input_pdf_template)
#   This method will write your PDF files as specifde in the input_CSV


#! /usr/bin/python

import os
import pdfrw
import json

ANNOT_KEY = '/Annots'
ANNOT_FIELD_KEY = '/T'
ANNOT_VAL_KEY = '/V'
ANNOT_RECT_KEY = '/Rect'
SUBTYPE_KEY = '/Subtype'
WIDGET_SUBTYPE_KEY = '/Widget'
ANNOT_APEARANCE_KEY='/AP'



def makeCsv(input_pdf_template):
    #Loads PDF and exports formatted CSV template of PDF fillables

    template_pdf = pdfrw.PdfReader(input_pdf_template)
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
        csvfile.write('Name ending files: ,project, \n Folder name: ,folder,\n Folder Path: ,Current Directory,\n')
        csvfile.write('Name begining files:,PDF Name 1, PDF Name 2, PDF Name 3, Etc,\n')
        for fields_values in list_fields_and_values:
            fields_values=list(fields_values)
            if fields_values[1] is None:
                fields_values[1]=''
            csvfile.write('%s,%s \n' %(fields_values[0],fields_values[1]))

    pdfrw.PdfWriter().write(os.getcwd()+'/'+'PDF_Template.pdf', input_pdf_template)

    print('\n\n\nI made you a CSV file called PDF_Fields.csv in your current directory.  If you are just checking out the', 
        ' program, feel free to just edit the CSV file for the next step.  If you want to make a lot of PDFs, it will be easier to import',
        ' the CSV to excel and add your form information.  You will then need to export it to a CSV to import it into the makePdfs function\n\n\n'
    )
            

def makePdfs(input_CSV, input_pdf_template):
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
    
    #load up pdf template
    template_pdf = pdfrw.PdfReader(input_pdf_template)
    annotations = template_pdf.pages[0][ANNOT_KEY]

    #Get font and font size
    flag=False
    for annotation in annotations:
        if annotation[SUBTYPE_KEY] == WIDGET_SUBTYPE_KEY:
            if annotation[ANNOT_FIELD_KEY]:
                key = annotation[ANNOT_FIELD_KEY][1:-1]
                print('\n\n',key)
            if annotation[ANNOT_APEARANCE_KEY] and annotation[ANNOT_APEARANCE_KEY] is not None:
                example_AP_dictionary=pdfrw.PdfDict({'/AP':annotation[ANNOT_APEARANCE_KEY]})
                
                print(example_AP_dictionary['/AP']['/N'], '\n\n')
                #with open("data_file.json", "w") as write_file:
                    #json.dump(example_AP_dictionary, write_file)
                #with open('data_file.json','r') as data:
                    #example_AP_dictionary=pdfrw.PdfDict(json.load(data))
                
                flag=True
                break
    if flag==False:
        raise Exception('You must fill one of the PDF fields with text so we can get the right font and such')

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
    
        
        
        #create dictionary of form keys and items information
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
                        pdfrw.PdfDict(example_AP_dictionary)
                        )

                        
                        

        pdfrw.PdfWriter().write(PDF_file_path+'.pdf', template_pdf)

def printDic(inputPdf):

    #load up pdf template
    template_pdf = pdfrw.PdfReader(inputPdf)
    annotations = template_pdf.pages[0][ANNOT_KEY]
    print('-------------------------------------------------------')

    for annotation in annotations:
        try:
            '''print('\n',annotation['/T'])
            print(annotation['/AP'])
            print(example_AP_dictionary['/AP'])    
            print(annotation['/AP']['/N'],'\n') '''       
            
            
        except:
            
            continue
        

    annotation=annotations[1]
    example_AP_dictionary=pdfrw.PdfDict(annotation[ANNOT_APEARANCE_KEY])
    print(type(example_AP_dictionary))

    #example_AP_dictionary= {'/AP':example_AP_dictionary}
    
    example_AP_dictionary=pdfrw.PdfDict(example_AP_dictionary)
    
    print(type(example_AP_dictionary))
    print(template_pdf.BasePdfName)


    print(example_AP_dictionary.keys())

    print(example_AP_dictionary['/AP']['/N'])
    with open("data_file.json", "w") as write_file:
                json.dump(example_AP_dictionary, write_file)
            #with open('data_file.json','r') as data:
                #example_AP_dictionary=pdfrw.PdfDict(json.load(data))

if __name__ == "__main__":
    #makeCsv('Jacob_DA31.pdf')
    #makeCsv('DA-31-unlocked.pdf')
    #makePdfs('PDF_Fields.csv','Jacob_DA31.pdf')
    #makePdfs('PDF_Fields.csv','DA-31-unlocked.pdf')
    #makePdfs('PDF_Fields.csv','PDF_Template.pdf')
    printDic('DA-31-unlocked.pdf')
    
    
    pass

