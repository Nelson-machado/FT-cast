from tkinter import *
from tkinter.filedialog import askopenfilename
from PIL import Image, ImageTk
import comtypes.client
import datetime
import time

machado_root = Tk()

# width x Height
machado_root.geometry("555x555")

# Title of the window.
machado_root.title("machado creations")

# titlebar icon
windowPhoto = PhotoImage(file="img/icon.jpg")
machado_root.iconphoto(False, windowPhoto)

# Min width, height
machado_root.minsize(555, 555)
# machado_root.configure(background='white')
# Max width, height
machado_root.maxsize(555, 555)
chosenFileVar = StringVar()
changedFileVar = StringVar()
convertBtnVar = StringVar()

def getFile():
    chosenFileVar.set('')
    
    Filename = askopenfilename() # showing an "Open" dialog box and return the path to the selected file.
    if Filename is not None and not (Filename == ""):
        if srcSelBtn.visible:
            srcSelBtn.pack_forget()

        chosenlabel1.pack(fill="both")
        chosenFileVar.set(Filename)
        chosenlabel2.pack(fill='y', padx=15, pady=5)
        convertBtn.pack(fill='x', padx=15, pady=5)
        if Filename.split('.')[1] == 'docx' or Filename.split('.')[1] == 'doc':
            convertBtnVar.set('To PDF')
        elif Filename.split('.')[1] == 'pdf':
            convertBtnVar.set('To docx')
        else:
            convertBtnVar.set('Try Converting')

        clrFileBtn.pack(fill='x', padx=15, pady=5)


def clearSelect():
    chosenFileVar.set('')
    convertBtn.pack_forget()

    if srcSelBtn.visible:
        srcSelBtn.pack(fill="both")
    clrFileBtn.pack_forget()
    chosenlabel2.pack_forget()
    chosenlabel1.pack_forget()

    if chosenlabel3.visible:
        chosenlabel3.pack_forget()
        changedFileVar.set('')
        chosenlabel4.pack_forget()

def convertSelected():
    in_file = StringVar()
    myfileOp = StringVar()
    myfiledate = StringVar()
    out_file = StringVar()

    in_file = chosenFileVar.get()
    print('in_file', in_file)
    # in_file F:/README_LOCATOR.doc

    myfiledate = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
    print('myfiledate', myfiledate)
    # myfiledate 20200913175554

    # Logic for converting the doc/docx files to pdf format.
    if in_file.lower().endswith('.docx') or in_file.lower().endswith('.doc'):
        wdFormatPDF = 17
        if in_file.lower().endswith('.docx'):
            myfileOp = in_file.replace('.docx', f'_{myfiledate}.pdf')
            print('docx-myfileOp-pdf', myfileOp)
        else:
            myfileOp = in_file.replace('.doc', f'_{myfiledate}.pdf')
            print('doc-myfileOp-pdf', myfileOp)

        out_file = myfileOp
        # absolute path is needed
        # out_file F:/README_LOCATOR_20200913175554.pdf
        print('out_file', out_file)

        # Creating COM object
        word = comtypes.client.CreateObject('Word.Application')
        word.Visible = 0
        time.sleep(3)

        # convert doc(x) file to pdf file.
        doc=word.Documents.Open(in_file) # open docx file 
        doc.SaveAs(out_file, FileFormat=wdFormatPDF) # conversion
        doc.Close() # close doc(x) file 
        # word.Visible = False
        word.Quit()

    elif in_file.lower().endswith('.pdf'):
        wdFormatPDF = 16
        myfileOp = in_file.replace('.pdf', f'_{myfiledate}.docx')
        print('pdf-myfileOp-docx', myfileOp)
        out_file = myfileOp
        # out_file F:/README_LOCATOR_20200913175554.pdf
        print('out_file', out_file)

        word = comtypes.client.CreateObject('Word.Application')
        word.Visible = 0
        time.sleep(3)   

        doc=word.Documents.Open(in_file)
        doc.SaveAs(out_file, FileFormat=wdFormatPDF)
        doc.Close() 
        word.Visible = False
        word.Quit()

    elif in_file.split('.')[1] not in ['docx', 'doc', 'pdf']:
        out_file = f"Sorry!! This file's extension is unknown to the converter for the moment."


    chosenlabel3.visible = True
    chosenlabel3.pack(fill="both")

    changedFileVar.set(out_file)

    chosenlabel4.visible = True
    chosenlabel4.pack(fill="both")

    convertBtn.pack_forget()

        
# Think of frames like div - that seggregates the sections.
frame1 = Frame(machado_root, bg="blue", relief=SUNKEN)
frame1.pack(side=TOP, fill="x")

# first Lable
titleLable = Label(frame1, text="Welcome to File converter", bg="red",  fg="white", padx=10, pady=10, font="comicsansms 19 bold", borderwidth=3, relief=SUNKEN )
titleLable.pack()

# side = top, bottom, left, right
#titleLable.pack(side="bottom", anchor="nw")

# adding image in a lable
""" photoImg = PhotoImage(file="converticon.png")
imgLable = Label(image=photoImg)
imgLable.pack() """

frame2 = Frame(machado_root, bg="white", relief=SUNKEN)
frame2.pack(fill="both", expand="Yes")

# to support jpg format using ImageTK of Pillow
# pip install pillow
image = Image.open("img/converter.jpg")
image = image.resize((1000, 800), Image.ANTIALIAS)
photo = ImageTk.PhotoImage(image)

imgLable = Label(frame2, image=photo)
# imgLable.pack(anchor="center")
imgLable.pack(fill="both", expand=YES)

frame3 = Frame(imgLable,  relief=SUNKEN)
frame3.pack(side=BOTTOM, fill="both")

srcSelBtn = Button(frame3, bg="grey", text="Select File", font="comicsansms 14 bold", padx=10, pady=10, command=getFile, borderwidth=4)
srcSelBtn.visible = True
srcSelBtn.pack(fill="both")

chosenlabel1 = Label(frame3, text="Source File:",padx=10, pady=10, font="comicsansms 14 bold")
chosenlabel2 = Label(frame3, textvariable=chosenFileVar, bg="yellow", padx=10, pady=10)

convertBtn = Button(frame3, textvariable=convertBtnVar,  bg="green", font="comicsansms 12 bold",padx=10, pady=10, command=convertSelected, borderwidth=4)
clrFileBtn = Button(frame3, text="Clear selection", font="comicsansms 12 bold", bg="red", padx=10, pady=10, command=clearSelect, borderwidth=4)

chosenlabel3 = Label(frame3, text="File available at:", padx=10, pady=10, font="comicsansms 14 bold")
chosenlabel3.visible = True
chosenlabel4 = Label(frame3, textvariable=changedFileVar, bg="yellow", padx=10, pady=10)
chosenlabel4.visible = True

machado_root.mainloop()