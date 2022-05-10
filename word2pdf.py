from tkinter import *
from tkinter.ttk import *
from tkinter.filedialog import askopenfile 
from tkinter import filedialog
import os
# Compile with :   
# c:\Python310\python.exe -m PyInstaller --noconsole --onefile "c:\Users\NSET_3420-03.DESKTOP-JSM47UL\Desktop\word2pdf.py"

import aspose.words as aw

class WordToPdf:
    def __init__(self):
        self.ws = Tk()
        self.ws.title('Word to PDF Converter - sendmail4ram@gmail.com')
        self.ws.geometry('900x120') 
        
        self.doclbl = Label(self.ws,text="Choose Word File")
        self.doclbl.grid(row=0, column=0, padx=10)
        self.docpathTxt = Text(self.ws, height = 1,width = 80)
        self.docpathTxt.grid(row=0, column=1)
        
        self.docbtn = Button(self.ws, text='Browse', command=self.open_file)
        self.docbtn.grid(row=0, column=2)


        self.pdflbl = Label(self.ws, text="Choose output PDF file")
        self.pdflbl.grid(row=1, column=0, padx=10)
        self.pdfpathTxt = Text(self.ws, height = 1,width = 80)
        self.pdfpathTxt.grid(row=1, column=1)
        self.pdfpathTxt.config(state=DISABLED)
        self.pdfbtn = Button(self.ws, text='Browse', command=self.choose_pdf)
        self.pdfbtn.grid(row=1, column=2)

        self.result_var = StringVar(value="")
        self.result = Label(self.ws,textvariable=self.result_var)
        self.result.grid(row=3, columnspan=3, padx=10)

        self.convertbtn = Button(self.ws, text='Convert to PDF', command=self.convert2pdf)
        self.convertbtn.grid(row=4, columnspan=3, pady=10)
    
        self.ws.mainloop()

    def open_file(self):
        self.doc_fname = askopenfile(mode='r', filetypes=[
            ('Word Document (*.docx', '*.docx'), 
            ('Word 97-2003 Document (*.doc', '*.doc'),
            ("all files","*.*")]).name
        self.docpathTxt.config(state='normal')
        self.docpathTxt.delete("1.0",END)
        self.docpathTxt.insert("1.0", self.doc_fname)
        self.docpathTxt.config(state='disabled')
        return  str(self.doc_fname)
    def choose_pdf(self):
        self.pdf_fname = filedialog.asksaveasfilename(defaultextension=".pdf",initialdir='/', title='Save File', filetypes = (("pdf files","*.pdf"),("all files","*.*")))    
        self.pdfpathTxt.config(state='normal')
        self.pdfpathTxt.delete("1.0",END)
        self.pdfpathTxt.insert("1.0", self.pdf_fname)
        self.pdfpathTxt.config(state='disabled')
        return str(self.pdf_fname)
    
    def convert2pdf(self):
        if self.doc_fname and self.pdf_fname:
            if os.path.isfile(self.doc_fname):
                self.result_var.set('Converting')
                try:
                    # Load word document
                    doc = aw.Document(self.doc_fname)
                    # Save as PDF
                    doc.save(self.pdf_fname)
                    # convert(self.doc_fname, self.pdf_fname)
                    self.result_var.set('CONVERTED - SUCCESS!!')
                except IOError as e:
                    self.result_var.set('FAILED - Something went wrrong!! \nPlease make sure that the word file is not opened in another program. Close the another program and try again'.format(e))
                    raise e
                except Exception as e:
                    self.result_var.set('FAILED - Something went wrrong!! \nPlease make sure that the word file of type .doc or .docx extension and the file exists (redable)'.format(e))
                    raise e
            else:
                print("Error in {}".format(self.doc_fname))
                raise Exception(self.doc_fname)

        
    def get_doc_fname(self):
        return self.docpathTxt.get("1.0",END)
    def get_pdf_fname(self):
        return self.pdfpathTxt.get("1.0",END)

w = WordToPdf()

# convert('C:/Users/NSET_3420-03.DESKTOP-JSM47UL/Desktop/pdfdoc.docx', 'C:/Users/NSET_3420-03.DESKTOP-JSM47UL/Desktop/s.pdf')