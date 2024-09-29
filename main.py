import os
import tkinter as tk
from tkinter import filedialog

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT

UTM_SOURCE = '?utm_source='
MEDIUM = '&utm_medium=blogs&utm_campaign='
FIN = '&utm_content=article'


def open_file():
    ready['text'] = ''
    selected_file["text"] = ''
    file = filedialog.askopenfile(mode='r', filetypes=[('Документ Word', '*.docx')])
    if file:
        global filepath
        global name
        global ext
        filepath = os.path.abspath(file.name)
        filename = os.path.basename(file.name)
        name, ext = os.path.splitext(os.path.basename(filepath))
        selected_file["text"] = filename


def submit():
    campaign_name = campaign.get()
    if not campaign_name:
        ready['text'] = 'Введите название кампании!'
    else:
        with open('workfiles/platforms.txt', 'r') as f:
            platforms = f.readlines()
        for platform in platforms:
            try:
                document = Document(filepath)
                rels = document.part.rels
                for rel in rels:
                    if rels[rel].reltype == RT.HYPERLINK:
                        old_url = rels[rel]._target
                        new_url = old_url + UTM_SOURCE + platform + MEDIUM + campaign_name + FIN
                        rels[rel]._target = new_url
                out_file = "workfiles/{name}_{uid}{ext}".format(name=name, uid=platform.rstrip('\n'), ext=ext)
                document.save(out_file)
                ready['text'] = 'Готово!'
            except NameError:
                ready['text'] = 'Укажите файл!'


root = tk.Tk()
root.title('Расстановка UTM-меток в файле')
root.geometry('400x250')

import_button = tk.Button(root, text='Импорт .docx', command=open_file)
import_button.pack(padx=6, pady=6)

selected = tk.Label(text='Выбранный файл:')
selected.pack(padx=3, pady=3)
selected_file = tk.Label()
selected_file.pack(padx=6, pady=6)

label = tk.Label(text='Название кампании:')
label.pack(padx=6, pady=6)

campaign = tk.Entry()
campaign.pack(padx=6, pady=3)

submit_button = tk.Button(root, text='Применить', command=submit)
submit_button.pack(padx=6, pady=10)

ready = tk.Label()
ready.pack(padx=6, pady=6)

root.mainloop()
