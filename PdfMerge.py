# Little programm to merge multiple PDFs into one single file
# Author: Felix Rutsch
# Date: 6.10.2022

import os
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from PyPDF2 import PdfFileReader, PdfFileWriter


class DragDropListbox(tk.Listbox):

    def __init__(self, master, **kw):
        kw['selectmode'] = tk.SINGLE  # only one item selectable at a time
        tk.Listbox.__init__(self, master.root, kw)
        self.bind('<Button-1>', self.setCurrent)  # left mouse click
        self.bind('<B1-Motion>', self.shiftSelection)  # moving while holding left mouse button
        self.bind('<Delete>', self.deletefromlist)  # delete key pressed
        self.curIndex = None
        self.masterapp = master

    def setCurrent(self, event):
        self.curIndex = self.nearest(event.y)

    def shiftSelection(self, event):  # perform swap operation in listbox and pdflist
        i = self.nearest(event.y)
        if i < self.curIndex:
            x = self.get(i)
            self.delete(i)
            self.insert(i + 1, x)

            pdf = self.masterapp.pdflist[i]
            del self.masterapp.pdflist[i]
            self.masterapp.pdflist.insert(i + 1, pdf)

            self.curIndex = i

        elif i > self.curIndex:
            x = self.get(i)
            self.delete(i)
            self.insert(i - 1, x)

            pdf = self.masterapp.pdflist[i]
            del self.masterapp.pdflist[i]
            self.masterapp.pdflist.insert(i - 1, pdf)

            self.curIndex = i

    def deletefromlist(self, _):
        if self.curIndex is not None:
            self.masterapp.delete_pdf(self.curIndex)
            self.curIndex = None


class PdfMergeApp:
    def __init__(self):

        self.pdflist = []  # initialize list of pdfs

        # ----- set up GUI -----

        self.root = tk.Tk()  # initialize tkinter window
        self.root.title('Pdf Merge')
        script_dir = os.path.dirname(__file__)
        try:
            self.root.iconbitmap(script_dir + '/PdfMerge_icon.ico')
        except tk.TclError:  # error ocurring when .ico file is not found
            print("WARNING: icon file not found.")
        pass  # proceed, app can still be executed without the icon

        # initialize rows and colums, so that they adjust width and height when window is resized
        self.root.columnconfigure([0, 1], weight=1, minsize=150)
        self.root.rowconfigure([0, 2], weight=1, minsize=100)

        self.btn_add = tk.Button(master=self.root, text="Add PDF(s)", width=20, command=self.add_pdfs)
        self.btn_add.grid(row=0, column=0, padx=5, pady=5)

        self.btn_merge = tk.Button(master=self.root, text="Merge and Save as...", width=20, command=self.merge)
        self.btn_merge.grid(row=0, column=1, padx=5, pady=5)

        self.lbl_listheader = tk.Label(self.root, text='Selected files to be merged:', relief='ridge', justify='left')
        self.lbl_listheader.grid(row=1, column=0, columnspan=2, sticky="NSEW")  # NSEW centers it

        self.update_pdflist()

        # main loop to keep window open
        self.root.mainloop()

    def update_pdflist(self):

        if self.pdflist:  # at least one pdf file is selected

            self.pdflistbox = DragDropListbox(self, height=1)
            self.pdflistbox.grid(row=2, column=0, columnspan=2, padx=5, pady=5, sticky="NSEW")

            for file in self.pdflist:
                # file is a string containing the directory -> cut at last slash to get filename only
                i = len(file) - 1
                while (file[i] != '/') and (i >= 0):  # iterate backwards through string until a slash is found
                    i -= 1
                filename = file[i + 1:]

                self.pdflistbox.insert('end', filename)

            newItemNumber = len(self.pdflist)
            if newItemNumber <= 15:  # maximum displayed length
                self.pdflistbox.config(height=newItemNumber)
            else:
                self.pdflistbox.config(height=15)

            self.btn_merge.config(state='normal')  # activate merge button

        else:  # pdflist is empty

            self.pdflistbox = tk.Label(self.root, text='No files are currently selected. \nClick \'Add PDF(s)\' to select files. \nHint: use Ctrl or Shift to select multiple files at once.', justify='left')
            self.pdflistbox.grid(row=2, column=0, columnspan=2, sticky="NSEW")  # NSEW centers it

            self.btn_merge.config(state='disabled')  # deactivate merge button

    # ----- button functions -----
    def add_pdfs(self):

        files = filedialog.askopenfilenames(parent=self.root, title='Select .pdf files', filetypes=(("PDF", "*.pdf"), ("all files", "*.*")))

        if files:  # check if any files were selected, otherwise files is empty

            for file in files:
                if file[len(file) - 4:] != '.pdf':  # check if selected files are all of type .pdf
                    messagebox.showerror("Error: Wrong File Format", "At least one of the selected files has a wrong file format. Only pdf files are allowed.")
                    return

            for file in files:
                self.pdflist.append(file)

            self.update_pdflist()

    def delete_pdf(self, index):
        del self.pdflist[index]
        self.update_pdflist()

    def merge(self):
        filesave = filedialog.asksaveasfilename(parent=self.root, title='Save as...', defaultextension='.pdf', filetypes=(("PDF", "*.pdf"), ("all files", "*.*")))

        if filesave:
            pdf_writer = PdfFileWriter()
            for file in self.pdflist:
                pdf_reader = PdfFileReader(file)
                for page in range(pdf_reader.getNumPages()):
                    pdf_writer.addPage(pdf_reader.getPage(page))

            with open(filesave, 'wb') as out:
                pdf_writer.write(out)

            print(f"successfully merged {len(self.pdflist)} files.")
            # self.root.destroy()  # close window as merge is complete


PdfMergeApp()  # execute app
