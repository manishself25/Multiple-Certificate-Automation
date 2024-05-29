#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import tkinter
import numpy as np
import pandas as pd
from tkinter import ttk
from tkinter import *
from tkinter import filedialog  
from docxtpl import DocxTemplate
from tkcalendar import Calendar
from datetime import datetime
from docx2pdf import convert

def generat_certificate(names,application_no,cname):
    
    doc = DocxTemplate("edS Certificate.docx")
    
    now = datetime.now()
    current_date = now.date()
    formatted_date = current_date.strftime('%d %B %Y')
    
    doc.render({"name" :names.title(),
                "application_no" : application_no,
            "course_name" :cname,
            "date" : formatted_date})

    doc_name = names +" " +application_no+ ".docx"
    doc.save(doc_name)
    pdf_name = names +" " +application_no+ ".docx"
    convert(pdf_name)
    print(f"{name} Certificate Generate Successful ")
    
def browseFiles():
    global df
    filename = filedialog.askopenfilename(initialdir="/",
                                          title="Select a File",
                                          filetypes=(("Excel files", "*.xlsx"),
                                                     ("CSV files", "*.csv")))
    label_file_explorer.configure(text="File Opened: " + filename)
    if filename.endswith('.csv'):
        df = pd.read_csv(filename)
    elif filename.endswith('.xlsx'):
        df = pd.read_excel(filename)

def exit_app():
    window.destroy()

window = tkinter.Tk()
window.title('File Explorer')
#window.geometry("500x500")
window.config(background="white")

label_file_explorer = Label(window, text="File Explorer using Tkinter", width=50, height=4, fg="blue")
label_file_explorer.grid(column=0, row=1)

button_explore = Button(window, text="Browse Files", command=browseFiles)
button_explore.grid(column=0, row=2,padx = 50,pady = 10)

button_exit = Button(window, text="Exit", command=exit_app)
button_exit.grid(column=0, row=3,padx = 50,pady = 10)

window.mainloop()

try:
    print(df.columns)
except NameError:
    print("No file loaded or 'df' is not defined.")

for i in range(len(df)):
    application_no = df["Application_no"].iloc[i]
    names = df["Name"].iloc[i]
    cname = df["Course_name"].iloc[i]
    generat_certificate(names,application_no,cname)

