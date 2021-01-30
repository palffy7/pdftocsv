import os
import pandas as pd
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdftypes import resolve1
from tkinter import *
from tkinter import filedialog


def create_csv(folder_path):
    pdf_files = [f for f in os.listdir(folder_path) if f.endswith('.pdf')]
    encoding = 'iso-8859-1'
    first = True
    k = 0
    files_count = len(pdf_files)
    for pdf_file_name in pdf_files:
        columns_pdf = []
        values = []
        file = folder_path + '/' + pdf_file_name
        with open(file, 'rb') as pdf_file:
            parser = PDFParser(pdf_file)
            doc = PDFDocument(parser)
            fields = resolve1(doc.catalog['AcroForm'])['Fields']
            for i in fields:
                field = resolve1(i)
                # try because of diameter sign
                try:
                    name = str(field.get('T'), encoding)
                except:
                    name = str(field.get('T')[:-8], encoding)
                opt = field.get('Opt')
                sel = field.get('V')
                # check if options are available and comparison
                if opt != None:
                    if not isinstance(type(opt), list):
                        opt = resolve1(opt)
                    for e in opt:
                        # Field has no 2 array list
                        if name == 'Beobachter':
                            if e == sel:
                                value = e
                        elif e[0] == sel:
                            value = e[1]
                else:
                    value = sel

                # just bytes can be decoded
                if isinstance(value, bytes):
                    try:
                        value = str(value, encoding)
                    except:
                        value = value
                elif str(value)[0] == r"/":
                    value = str(value)[2:-1]
                else:
                    value = str(value)

                columns_pdf.append(name)
                values.append(value)

            if first:
                columns_init = columns_pdf.copy()
                columns_init.append('file')
                df = pd.DataFrame(columns=columns_init)
                first = False
            df_pdf = pd.DataFrame([values], columns=columns_pdf)
            filename = [pdf_file_name]
            df_pdf['file'] = filename
            df = df.append(df_pdf)
            k += 1
            text_count.set(str(k) + ' von ' + str(files_count))
            root.update()
    df = df.replace({'None': '-'})
    df = df.fillna('-')
    first_col = df.pop('file')
    df.insert(0, 'file', first_col)
    df.to_csv(folder_path + '.csv', index=False)
    root.destroy()


def get_folder_path():
    folder_selected = filedialog.askdirectory(initial=os.getcwd())
    if not folder_selected == '':
        new_text = 'csv file is built. \n This can take a moment.'
        text.set(new_text)
        root.update()
        create_csv(folder_selected)


def main():
    global root, text, text_count
    cwd = os.getcwd()
    root = Tk()
    root.geometry('400x300')
    root.title('PTC_Exporter')

    if os.path.isfile(os.path.join(cwd, 'Logo.png')):
        widget = Label(root, compound='top')
        widget.logo_png = PhotoImage(file=os.path.join(cwd, 'Logo.png'))
        widget['image'] = widget.logo_png
        widget.pack()

    btn = Button(root, text='Choose folder', command=get_folder_path).pack()

    text = StringVar()
    text.set('Choose folder with button')
    label = Label(root, textvariable=text)
    label.pack()

    text_count = StringVar()
    text_count.set('')
    label_count = Label(root, textvariable=text_count)
    label_count.pack()

    info_text = StringVar()
    info_text.set('Info: This tool exports data from pdf forms in a csv file. \n'
                  'Click on the "Choose folder" and choose a folder which contains \n'
                  'the pdf form files. You can find the generated csv file afterwards \n'
                  'in the folder where the chosen folder is named after the folder name. \n'
                  'The csv file can afterwards be imported in an excel file for further use.')

    label_info = Label(root, textvariable=info_text)
    label_info.place(relx=0.5, rely=0.8, anchor='center')
    label_info.pack()

    root.mainloop()


if __name__ == '__main__':
    main()
