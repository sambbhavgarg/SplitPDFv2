from PyPDF2 import PdfFileWriter, PdfFileReader
import pandas as pd
import os

class SplitPDF:
    def __init__(self, config): 
        #path to excel data file
        self.excel_file_path = os.path.join(config['data_dir'], config['excel_data_file_name'])
        
        #path to all certificates file created from Mail Merge in MS Word
        self.all_certs_file_path = os.path.join(config['program_name_dir'], config['all_certs_mail_merge_pdf'])
        print(self.all_certs_file_path)
        
        #path to the directory where all individual certificates will be stored
        self.out_path = os.path.join(config['program_name_dir'], config['out_dir_name_for_split_pdfs'])
        
        #loading excel data file into a pandas df object
        self.df = pd.read_excel(self.excel_file_path, sheet_name=config['sheet_num'])
        
        #adding a column where the student name is suffixed to create unique names, in order to tackle conflict while doing
        # - Gmass
        self.df['attachment'] = self.df.iloc[:, 0].str.strip() + f"{config['individual_file_suffix_1']}{config['individual_file_suffix_2']}"
        
        print("Printing Results...")
        display(self.df)
                
        self.run, self.create_sheet_with_attachments_column = config['run'], config['create_sheet_with_attachments_column']

    def splitPDF(self):
        if self.run:
            #change -1 if the individual file names refer to some other column
            
            #iterator object for storing file names with suffixes
            iter_names = iter((self.df.iloc[:, -1]).values.tolist()) 
            
            #PDF reader object
            inputpdf = PdfFileReader(open(self.all_certs_file_path, "rb"))
            
            #create out directory folder if it doesnt exist            
            if not os.path.exists(self.out_path):
                os.mkdir(self.out_path)
            
            #template pdf name
            output_pdf = '{}.pdf'
            
            #iteratively get pages from inputpdf and write to new pdf with personalized file name
            for i in range(inputpdf.numPages):
            #     if i%2 == 0:
                output = PdfFileWriter()
                output.addPage(inputpdf.getPage(i))
            #     else:
            #     output.addPage(inputpdf.getPage(i))
            #     with open(os.path.join(out_path, output_pdf.format(i)), "wb") as outputStream:
            #         output.write(outputStream)
                output_file = os.path.join(self.out_path, output_pdf.format(next(iter_names)))
                with open(output_file, 'wb') as outputStream:
                    output.write(outputStream)
                    
        if self.create_sheet_with_attachments_column:
            self.df.to_excel(os.path.join(config['program_name_dir'], 'data-with-attachment.xlsx'), index=False)