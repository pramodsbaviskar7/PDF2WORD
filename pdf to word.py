#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from tkinter import *
from tkinter import filedialog
from tkinter.filedialog import askopenfilename,asksaveasfile
from PyPDF2 import PdfFileReader
#opening the file with .pdf extension only
def openFile(): 
              
    file = askopenfilename(defaultextension=".pdf", 
                                          filetypes=[("Pdf files","*.pdf")])
    if file == "":  
        file = None
    else:
        fileEntry.delete(0,END)
        fileEntry.config(fg="blue")
        fileEntry.insert(0,file)
#reading the pdf content
def convert():
    try:
        pdf = fileEntry.get()
        pdfFile = open(pdf, 'rb')
        # creating a pdf reader object
        pdfReader = PdfFileReader(pdfFile) 

        # creating a page object 
        pageObj = pdfReader.getPage(0) 
      
        # extracting text from page 
        extractedText= pageObj.extractText()
        readPdf.delete(1.0,END)
        readPdf.insert(INSERT,extractedText)

        # closing the pdf file object 
        pdfFile.close()
    except FileNotFoundError:
        fileEntry.delete(0,END)
        fileEntry.config(fg="red")
        fileEntry.insert(0,"Please select a pdf file first")
    except:
        pass
#converting it to.docx
    
def save2word():
    text = str(readPdf.get(1.0,END))
    wordfile = asksaveasfile(mode='w',defaultextension=".doc", 
                                          filetypes=[("word file","*.doc"),
                                                     ("text file","*.txt"),
                                                     ("Python file","*.py")])
    
    if wordfile is None:
        return
    wordfile.write(text)
    wordfile.close()
    print("saved")
    fileEntry.delete(0,END)
    fileEntry.insert(0,"pdf Extracted and Saved...")



    
#Designing frontend view
root = Tk()
root.geometry("800x350")
root.config(bg="light grey")
root.title("PDF to word ")
root.resizable(0,0)
try:
    root.wm_iconbitmap("icon.ico")
except:
    print('icon file is not available')
    pass
file= ""
defaultText = "\n\n\n\n\t\t Your extracted text will apear here.\n \t\t     you can modify that text too."

appName = Label(root,text="PDF to WORD Converter by Pramod ",font=('arial',20,'bold'),
                bg="light grey",fg='grey')
appName.place(x=150,y=5)
#Select pdf file
labelFile = Label(root,text="Select file",font=('Calibri',13,'bold'))
labelFile.place(x=30,y=50)
fileEntry = Entry(root,font=('calibri',12),width=40)
fileEntry.pack(ipadx=200,pady=50,padx=150)

openFileButton = Button(root,text=" Open ",font=('arial',12,'bold'),width=30,
                    bg="grey",fg='white',command=openFile)
openFileButton.place(x=150,y=80)

convert2Text = Button(root,text="Read",font=('arial',8,'bold'),
                   bg="grey",fg='WHITE',width=20,command=convert)
convert2Text.place(x=250,y=120)

readPdf = Text(root,font=('calibri',12),fg='light grey',bg='black',width=60,height=60,bd=10)
readPdf.pack(padx=20,ipadx=20,pady=20,ipady=20)
readPdf.insert(INSERT,defaultText)

save2Word = Button(root,text="Save as word",font=('arial',10,'bold'),
                   bg="grey",fg='WHITE',command=save2word)
save2Word.place(x=255,y=320)


if __name__ == "__main__":
    root.mainloop()


# In[ ]:





# In[ ]:




