import PyPDF2
from pathlib import Path
import os
from datetime import datetime
import time



class src_info():
    '''
    This class will hold all the meta data information user had key-in
    '''
    def __init__(self):
        self.remove_duplicate=False
        self.operation=''
        self.target_dir=''
        self.file_count=int()
        self.source_location=None
        self.file_name=''
        self.page_count=0

        self.cache_src=None
        self.file_info={ "operation"   :     "", 
                         "target_dir"  :     "",
                         "file"        :     []
                        }
    
    def add_entry(self,each_file):
        '''
        This function is used to construct a place holder for source pdf informations
        '''
        if(len(self.file_info['file'])>0 and len(each_file)>0 and self.remove_duplicate):
            for each_entry in self.file_info['file']:
                if (each_file['file_loc']        == each_entry ['file_loc'] and
                    each_file['file_name']       == each_entry ['file_name'] and
                    each_file['page_count']      == each_entry ['page_count'] and
                    each_file['abs_path']        == each_entry ['abs_path'] and
                    each_file['from_page_no']    == each_entry ['from_page_no'] and
                    each_file['to_page_no']      == each_entry ['to_page_no']):
                        return False
                
                    
        self.file_info["file"].append(each_file) 
        return True


    
    def get_operation(self):
        '''
        This function is to get the choice of operation user has entered
        '''
        
        choice=''
        operation=''
        choice=input('Please select appropriate action [1]-Split; [2]-Merge; [3]-Slice-Merge  : ')
        print('')

        while choice not in ['1','2','3']:
            print('Invalid choice. Retrying.....')
            choice=input('Please select appropriate action [1]. Split; [2]. Merge; [3].Slice-Merge  : ')
            print('')
        
        if choice=='1':
            operation='Split'
        elif choice=='2':
            operation='Merge'
        elif choice=='3':
            operation='Slice-Merge'
        return operation
        

    def get_file_meta_data(self):
        '''
        This function holds the meta information of source pdf file
        '''
        while True:
            
            try:
                if self.cache_src==None:
                    self.source_location=input('File Location (ex: C:\\Users ): ')
                    self.cache_src=self.source_location
                else:
                    self.source_location=input('File Location (ex: C:\\Users, or hit enter for previous location ): ')
                    if self.source_location==None or self.source_location=='':
                        self.source_location=self.cache_src


                self.file_name=input('File Name    : ')
                
                if '.pdf' not in self.file_name:
                    self.file_name=self.file_name+".pdf"
                
                source_file_path=Path(self.source_location+"\\"+self.file_name)
                
                with open(source_file_path,mode='rb') as src_file:

                    src_pdf_reader=PyPDF2.PdfFileReader(src_file)
                    self.page_count=src_pdf_reader.numPages
                    print(f"{self.file_name} has total {self.page_count} pages")
                    break
            except (FileNotFoundError) as e:
                print(f"{e} Retrying......")
                continue
            
        page_to_split=[0,0]
        if self. operation in ('Split','Slice-Merge')  :        
            
            while page_to_split[0] not in range(1,self.page_count+1):
                page_to_split[0]=input("From which page you want to split: ")
                if page_to_split[0].isdigit()==True:
                    page_to_split[0]=int(page_to_split[0])
                else:
                    print(f"Sorry invalid page no. From page should be within 1 to {self.page_count} . Lets retry!")
                    continue

                if page_to_split[0] not in range(1,self.page_count+1):
                    print(f"Sorry invalid page no. From page should be within 1 to {self.page_count} . Lets retry!")
            
            while page_to_split[1] not in range(int(page_to_split[0]),self.page_count+1):
                page_to_split[1]=input("Upto which page you want to split: ")

                if page_to_split[1].isdigit()==True:
                   page_to_split[1]=int(page_to_split[1]) 
                else:
                   print(f"Sorry invalid page no. To page should be within {page_to_split[0]} to {self.page_count} . Lets retry!")
                   continue 
                if page_to_split[1] not in range(int(page_to_split[0]),self.page_count+1):
                    print(f"Sorry invalid page no. To page should be within {page_to_split[0]} to {self.page_count} . Lets retry!")
            
        else:
            page_to_split=[1,self.page_count]
            # print(f"Merged page count added: {page_to_split[0]}/{page_to_split[1]}")
        page_to_split.sort()
        return[self.source_location,self.file_name,self.page_count,str(source_file_path),page_to_split[0],page_to_split[1]]
        
        
    def get_file_info(self):
        '''
        This function is used to get directory information
        '''
        self.operation=self.get_operation()
        self.file_info['operation'] = self.operation
        
        self.target_dir=input("Please provide target directory to download output PDF , else hit enter : ")
        if self.target_dir==None or self.target_dir=='':
            self.target_dir=str(os.getcwd())
            self.file_info['target_dir'] = str(os.getcwd())
        else:
            self.file_info['target_dir'] = self.target_dir
        while (not os.path.isdir(self.file_info['target_dir'])):
            print("Sorry, Target directory doesnot exits.")
            self.target_dir=input("Please provide target directory to download output PDF , else hit enter : ")
            if self.target_dir==None or self.target_dir=='':
                self.target_dir=str(os.getcwd())
                self.file_info['target_dir'] = str(os.getcwd())
            else:
                self.file_info['target_dir'] = self.target_dir
        print('')
        print(f"Selected Output directory                                               : {self.target_dir}")
        print("\n")
        if self. operation in ('Merge','Slice-Merge')  :
            self.file_count=input("Number of file(s) involved: ")
            while True:
                if (self.file_count.isdigit()==False):
                    print('Invalid Number of files. It should be postive digit and more than 0.')
                    self.file_count=input("Number of files: ")
                else:
                    break
            
            self.file_count=int(self.file_count)
            
            if self.file_count<2:
                print(f"Sorry, minimum 2 file required for {self. operation}. Please retry with correct input ")
                return False
        elif self.operation=='Split':
            self.file_count=1
        

        for i in range(self.file_count):
            if self.file_info['operation']=='Split':
                print("*"*82)
                print("*"*35+" File "+self.file_info['operation']+" "+ "*"*35)    
                print("*"*82)
                print('\n'*2)
                print("+"*20)
                print("File information: ")
                print("+"*20)
                print('\n'*1)
                each_file={}
                file_meta_data=[]
                file_meta_data=self.get_file_meta_data()
                each_file['file_loc']          = file_meta_data[0] 
                each_file['file_name']         = file_meta_data[1] 
                each_file['page_count']        = file_meta_data[2] 
                each_file['abs_path']          = file_meta_data[3]
                each_file['from_page_no']      = file_meta_data[4]
                each_file['to_page_no']        = file_meta_data[5]
            elif self.file_info['operation'] in ('Merge') :
                if i==0:
                    print("*"*90)
                    print("*"*35+" File "+self.file_info['operation']+" "+ "*"*35)    
                    print("*"*90)
                    
                print('\n'*2)
                print("+"*20)
                print(f"{i+1}th File information: ")
                print("+"*20)
                print('\n'*2)
                each_file={}
                file_meta_data=[]
                file_meta_data=self.get_file_meta_data()
                each_file['file_loc']          = file_meta_data[0] 
                each_file['file_name']         = file_meta_data[1] 
                each_file['page_count']        = file_meta_data[2] 
                each_file['abs_path']          = file_meta_data[3]
                each_file['from_page_no']      = file_meta_data[4]
                each_file['to_page_no']        = file_meta_data[5]
            elif self.file_info['operation'] in ('Slice-Merge') :
                if i==0:
                    print("*"*90)
                    print("*"*35+" File "+self.file_info['operation']+" "+ "*"*35)    
                    print("*"*90)
                    
                print('\n'*2)
                print("+"*20)
                print(f"{i+1}th File information: ")
                print("+"*20)
                print('\n'*1)
                each_file={}
                file_meta_data=[]
                file_meta_data=self.get_file_meta_data()
                each_file['file_loc']          = file_meta_data[0] 
                each_file['file_name']         = file_meta_data[1] 
                each_file['page_count']        = file_meta_data[2] 
                each_file['abs_path']          = file_meta_data[3]
                each_file['from_page_no']      = file_meta_data[4]
                each_file['to_page_no']        = file_meta_data[5]
            
            # self.file_info["file"].append(each_file)  
            add_file_entry=self.add_entry(each_file)
            while(not add_file_entry):
                print("Duplicate information found. Please re-enter....")
                each_file={}
                file_meta_data=[]
                file_meta_data=self.get_file_meta_data()
                each_file['file_loc']          = file_meta_data[0] 
                each_file['file_name']         = file_meta_data[1] 
                each_file['page_count']        = file_meta_data[2] 
                each_file['abs_path']          = file_meta_data[3]
                each_file['from_page_no']      = file_meta_data[4]
                each_file['to_page_no']        = file_meta_data[5]
                add_file_entry=self.add_entry(each_file)
        
        return True
            
                
                
            
            
    
    
    

class play_pdf(src_info):
    '''
    This class is used mainly for doing the operation , such as Split , Merge or Slice-Merge
    '''
    def __init__(self):
          src_info.__init__(self)
          self.meta_derived = src_info.get_file_info(self)
          
    
    def make_pdf(self):
        '''
        This function does the pdf operation
        '''
        if not self.meta_derived:
            return
        target_file_name=None
        output_file_path=None
        tmp=None
        pdf_writer=PyPDF2.PdfFileWriter()
        for each_entry in self.file_info['file']:
            if target_file_name==None or target_file_name=='':
                target_file_name=str.replace(each_entry ['file_name'],'.pdf','')
                if self.operation=='Split':
                    target_file_name=target_file_name+"["+str(each_entry ['from_page_no'])+"_"+str.replace(str(each_entry ['to_page_no']),'.pdf','')+"]"+"_"+str(datetime.now().microsecond)+".pdf"
                else:
                    target_file_name="Merged_"+str.replace(each_entry ['file_name'],'.pdf','')+"_"+str(datetime.now().microsecond)+".pdf"
                


            source_file_path=Path(each_entry ['abs_path'])
            
            with open(source_file_path,mode='rb') as src_file:
                pdf_reader=PyPDF2.PdfFileReader(src_file)
                
                
                for page_num in range(each_entry ['from_page_no'],each_entry ['to_page_no']+1):
                    pdf_writer.addPage(pdf_reader.getPage(int(page_num)-1))
                        
                
                with open(Path(self.target_dir+"\\"+target_file_name),"wb") as pdf_output:
                    output_file_path=Path(self.target_dir+"\\"+target_file_name)
                    pdf_writer.write(pdf_output)  
        
        print("\n"*5)
        print("*"*110)
        print("*"*35+" Congratulation Operation Completed !!!" + "*"*35)
        print("*"*110)
        print(f"New file Dowloaded At: {output_file_path}")
        print("")
        print("NOTE: This window will be close automatically in next 15 seconds.")
        print('\n'*5)

        time.sleep(10)



def welcome():
    '''
    This function used to construct User Guide.
    '''
    no_spcs=123
    
    print("\n"*1)
    # print("\t\t"+"="*123)
    print(f"\t\t"+"="*50+" Welcome to PDF Utilty "+"="*52)
    print(f" "*15+" |User Guide: "+" "*111+"|")
    print("\t\t"+"|"+"-"*no_spcs+"|")
    print(" "*16+"| Operation - Split      :  By this operation you can split the document.                                                   |")
    print(" "*16+"| Operation - Merge      :  By this operation you can merge n number of document.                                           |")
    print(" "*16+"| Operation - Slice-Merge:  By this operation you can merge the splitted document on the fly.                               |") 
    print(" "*16+"|                           Example, if you have 3 files and you want to make pdf using first page of each file.            |")
    print(" "*16+"|                           You should choose this option                                                                   |")
    print("\t\t"+"|"+" "*no_spcs+"|")
    print("\t\t"+"|"+"-"*no_spcs+"|")
    print(f" "*15+" |Expected Input Parameter value: "+" "*91+"|")
    print("\t\t"+"|"+"-"*no_spcs+"|")
    print("\t\t"+"|"+" "*no_spcs+"|")
    print(" "*16+"| 01) Action [Options- [1]-Split; [2]-Merge; [3]-Slice-Merge:] : Expected value- 1 or 2 or 3 ,i.e. Positive digit.          |")
    print("\t\t"+"|"+" "*no_spcs+"|")
    print(" "*16+"| 02) Number of file(s) Involved : Expected value- Positive digit, for Slice-Merge and Merge it should be greater than 1    |")
    print("\t\t"+"|"+" "*no_spcs+"|")
    print(" "*16+"| 03) File Location              : Expected value- Any directory location where the PDF file is placed, ex: C://Users       |")
    print("\t\t"+"|"+" "*no_spcs+"|")
    print(" "*16+"| 04) File Name                  : Expected value- Source PDF file name with or without extension,case sensitive,           |")
    print(" "*16+"|                                  ex: <TEST.pdf> or <TEST>                                                                 |")
    print(" "*16+"| 05) From Page Split            : Expected value- Any Positive digit within source PDF page limit.                         |") 
    print(" "*16+"|                                  This Output PDF will START from this page.                                               |")
    print("\t\t"+"|"+" "*no_spcs+"|")
    print(" "*16+"| 06) To Page Split              : Expected value- Any Positive digit within source PDF page limit. Should be greater       |") 
    print(" "*16+"|                                  than <From Page Split>. This Output PDF will END from this page.                         |")
    print("\t\t"+"|"+" "*no_spcs+"|")
    print(" "*16+"| 07) Output File Location       : Expected value- Any directory location where you want to place the output PDF file.      |")
    print(" "*16+"|                                  ex: C:Users . Default is current working directory or script running directory           |")
    print("\t\t"+"|"+" "*no_spcs+"|")
    print("\t\t"+"="*(no_spcs+2))
    print(f" "*15+" | Created by Sushovan Basak. Linkedin: https://www.linkedin.com/in/sushovan-basak-5ab2294b/ "+" "*32+"|")
    print("\t\t"+"="*(no_spcs+2))
    print("\n"*1)
if __name__ == "__main__":
    welcome()
    pdf=play_pdf()
    pdf.make_pdf()
    
    
    
