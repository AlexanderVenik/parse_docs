from tkinter import *
from tkPDFViewer import tkPDFViewer as pdf
from tkinter import ttk
from tkinter import filedialog
import win32com.client
import os

root = Tk() 
word = win32com.client.Dispatch("Word.Application")

mainmenu = Menu(root) 
root.config(menu=mainmenu) 
 
filemenu = Menu(mainmenu, tearoff=0)
def openFiles():
    vals = filedialog.askopenfilenames(filetypes =[("Все документы Word", ".docx .docm .dotx .dotm .doc .dot .htm .html .rtf .mht .mhtml .xml .odt .pdf"), ("All files", "*.*")])
    for path in vals:
        name = path[path.rfind("/")+1:]
        left.append((name, path.replace("/","\\")))
    pass
def openDir():
    print(os.getcwd())
    pass
def _exit():
    v1.test_delete()
    pass
filemenu.add_command(label="Открыть файлы", command=openFiles)
filemenu.add_command(label="Открыть папку", command=openDir)
filemenu.add_command(label="Выход", command=_exit)

mainmenu.add_cascade(label="Файл", menu=filemenu)


class Table:
    def __init__(self, master, cols, names, onDblClic):
        self.frame = Frame(root)
        self.treeView = ttk.Treeview(self.frame)
        self.treeView['columns'] = cols

        self.treeView.column("#0", width=0,  stretch=NO)
        self.treeView.heading("#0",text="",anchor=CENTER)
        for col, name in zip(cols, names):
            self.treeView.column(col,anchor=CENTER, width=80)
            self.treeView.heading(col,text=name,anchor=CENTER)
        self.treeView.bind("<Double-1>", onDblClic)
        self.counter=0
    def append(self, vals):
        self.treeView.insert(parent='',index='end',iid=self.counter,text='', values=vals)
        self.counter+=1
    def pack(self):
        self.treeView.pack(fill=Y, side=LEFT, expand=True)
        self.frame.pack(fill=Y, side=LEFT, expand=True)
def convert(path):
    temp = os.getcwd() + "\\" + "temp" + "\\"
    print(path)
    doc = word.Documents.Open(path)
    wdFormatPDF = 17
    newpath = temp + path[path.rfind("\\")+1:].split(".")[0] + ".pdf"
    print(newpath)
    doc.SaveAs(newpath, FileFormat=wdFormatPDF)
    doc.Close()
    return newpath
def onDocumentSelect(e):
    item = left.treeView.identify('item',e.x,e.y)
    path = left.treeView.item(item)["values"][1]
    pathToPdf = convert(path)
    v1.test_delete()
    v1.update_(pathToPdf)

left = Table (root, ('document', 'path'), ("Документ", "Путь"), onDocumentSelect)
left.pack()

v1 = pdf.ShowPdf()
v2 = v1.pdf_view(root, pdf_location = r"C:\Users\Paul\Documents\OUT2.pdf", load="before")
v2.pack(side=LEFT, fill=BOTH, expand=False)


def onErrorSelect(event):
    pass

right = Table(root, ('document', 'page', 'line', 'char', 'desc'), ("Документ", "P", "L", "C", "Описание"), onErrorSelect)
right.pack()

root.mainloop()