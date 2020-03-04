import os
from docx import Document
import fitz
from tkfilebrowser import askopendirname, askopenfilenames, asksaveasfilename
try:
    import tkinter as tk
    from tkinter import ttk
    from tkinter import filedialog
except ImportError:
    import Tkinter as tk
    import ttk
    import tkFileDialog as filedialog
import PDFProcessor
import DOCXProcessor


class Reader:
    '''
        read PDF file `file_path` with PyMuPDF to get the layout data, including text, image and 
        the associated properties, e.g. boundary box, font, size, image width, height, then parse
        it with consideration for sentence completence, DOCX generation structure.
    '''

    def __init__(self, file_path):
        self._doc = fitz.open(file_path)

    def __getitem__(self, index):
        if isinstance(index, slice):
            stop = index.stop if not index.stop is None else self._doc.pageCount
            res = [self._doc[i] for i in range(stop)]
            return res[index]
        else:
            return self._doc[index]

    @staticmethod
    def layout(page):
        '''raw layout of PDF page'''
        layout = page.getText('dict')

        # remove blocks exceeds page region: negtive bbox
        layout['blocks'] = list(filter(
            lambda block: all(x>0 for x in block['bbox']),
            layout['blocks']))

        # reading order: from top to bottom, from left to right
        layout['blocks'].sort(
            key=lambda block: (block['bbox'][1],
                block['bbox'][0]))

        return layout

    @staticmethod
    def parse(page, debug=False):
        '''precessed layout'''
        raw = Reader.layout(page)
        if debug:
            return PDFProcessor.layout_debug(raw)
        else:
            return PDFProcessor.layout(raw)


class Writer:
    '''
        generate .docx file with python-docx based on page layout data.
    '''

    def __init__(self):
        self._doc = Document()

    def make_page(self, layout):
        '''generate page'''
        DOCXProcessor.make_page(self._doc, layout)

    def save(self, filename='res.docx'):
        '''save docx file'''
        self._doc.save(filename)


pdf_file='';
docx_file='';

docx = Writer()
def c_open_file():
    rep = askopenfilenames(parent=root, initialdir='/', initialfile='tmp',
                           filetypes=[("All files", "*")])
    pdf_file = rep[0];
    pdf = Reader(pdf_file)
    for page in pdf:
        layout = pdf.parse(page, True)
        docx.make_page(layout)
    docx.save(rep[0].replace('.pdf' ,'')+'.docx')
    print(rep[0])



def c_save():
    rep = asksaveasfilename(parent=root, defaultext=".docx", initialdir='/tmp', initialfile='sample.docx',
                            filetypes=[("All files", "*")])

    docx.save(rep[0])
    print(rep[0])



if __name__ == '__main__':


    root = tk.Tk()

    style = ttk.Style(root)
    style.theme_use("clam")
    root.configure(bg=style.lookup('TFrame', 'background'))
   # ttk.Label(root, text='Default dialogs').grid(row=0, column=0, padx=4, pady=4, sticky='ew')
    ttk.Label(root, text='Welcome to Document Convertor').grid(row=0, column=1, padx=4, pady=4, sticky='ew')
    #ttk.Button(root, text="Open files", command=c_open_file_old).grid(row=1, column=0, padx=4, pady=4, sticky='ew')
    #ttk.Button(root, text="Open folder", command=c_open_dir_old).grid(row=2, column=0, padx=4, pady=4, sticky='ew')
    #ttk.Button(root, text="Save file", command=c_save_old).grid(row=3, column=0, padx=4, pady=4, sticky='ew')
    ttk.Button(root, text="Convert files", command=c_open_file).grid(row=1, column=1, padx=4, pady=4, sticky='ew')
    #ttk.Button(root, text="Open folder", command=c_open_dir).grid(row=2, column=1, padx=4, pady=4, sticky='ew')
    #ttk.Button(root, text="Save file", command=c_save).grid(row=3, column=1, padx=4, pady=4, sticky='ew')
    #ttk.Button(root, text="Open paths", command=c_path).grid(row=4, column=1, padx=4, pady=4, sticky='ew')
    root.mainloop()

    #pdf_file = os.path.join(currentDirectory, 'pdf.pdf')
    #docx_file = os.path.join(currentDirectory, 'demo.docx')


