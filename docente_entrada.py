from tkinter import *
import ttkbootstrap
from ttkbootstrap.dialogs import Messagebox
import openpyxl
import datetime
import pandas
import numpy

#Main window configuration:
window = ttkbootstrap.Window(themename='superhero')
window.title('Ingreso datos docente')
window.geometry('370x390')
window.resizable()

#My functions:
def label_config(text_label, font_label, x_label, y_label):
    """
    This function creates ttkbootstrap labels considering the text,
    font and the position of the label as arguments.
    
    >>>lable_config('Hello, I am a label', ('Helvetica', 28), 10, 10)
    """
    ttkbootstrap.Label(window, text=text_label, font=font_label).place(x=x_label, y=y_label)

#-----------------------------------------------------------------------------------------------------------

def load_data():
    """
    This function lets the user to load information about a teacher into an Excel file,
    setting a teacher's code automatically.
    
    No arguments need for this function.
    """
    file = openpyxl.load_workbook('/home/bruce/Escritorio/Python/GUI/BD_SCHOOL_2023.xlsx')
    sheet = file['docente']
    
    first_value = sheet.cell(column=1, row=sheet.max_row).value
    counter = 0
    if type(first_value) == type(''):
        counter = 1
        sheet.cell(column=1, row=sheet.max_row + 1, value=1)
    else:
        counter = first_value + 1
        sheet.cell(column=1, row=sheet.max_row + 1, value=counter)
    
    data_name = entry_name.get()
    data_last_name = entry_last_name.get()
    data_gender = combobox_gender.get()
    data_birthday = entry_birthday.get()
    data_phone = entry_phone.get()
    data_email = entry_email.get()
    data_faculty = combobox_code_faculty.get()
    date = datetime.datetime.now()
    
    data_date = date.strftime('%d/%m/%Y %H:%M:%S')
    if data_name=='' or data_last_name=='' or data_gender=='' or data_birthday=='' or data_phone=='' or data_email=='' or data_faculty=='':
        Messagebox.show_error('Falta información por ingresar', title='Error!')
    elif '/' not in data_birthday:
        Messagebox.show_error('Ingrese la fecha de nacimiento en formato 00/00/0000', title='Error!')
    elif len(data_phone) != 9:
        Messagebox.show_error('Ingrese el teléfono en formato 0000-0000', title='Error!')
    elif '@' not in data_email:
        Messagebox.show_error('Ingrese el correo electrónico en formato micorreo@mail.com', title='Error!')
    else:
        sheet.cell(column=2, row=sheet.max_row, value='DOC0000' + str(counter))
        sheet.cell(column=3, row=sheet.max_row, value=data_name)
        sheet.cell(column=4, row=sheet.max_row, value=data_last_name)
        sheet.cell(column=5, row=sheet.max_row, value=data_gender)
        sheet.cell(column=6, row=sheet.max_row, value=data_birthday)
        sheet.cell(column=7, row=sheet.max_row, value=data_phone)
        sheet.cell(column=8, row=sheet.max_row, value=data_email)
        sheet.cell(column=9, row=sheet.max_row, value=data_date)
        sheet.cell(column=10, row=sheet.max_row, value=data_faculty)
        entry_name.delete(0, END)
        entry_last_name.delete(0, END)
        combobox_gender.delete(0, END)
        entry_birthday.delete(0, END)
        entry_phone.delete(0, END)
        entry_email.delete(0, END)
        combobox_code_faculty.delete(0, END)
        file.save('/home/bruce/Escritorio/Python/GUI/BD_SCHOOL_2023.xlsx')
        Messagebox.show_info('Información cargada correctamente', title='Felicidades')

def faculty_check():
    
    """
    
    This functions helps the user to know what is the name of the faculty asociated with a code.
    
    No arguments taken.
    
    """
    
    faculty_name_nparray = (faculty_sheet[faculty_sheet['CODIGO_FACULTAD']==combobox_code_faculty.get()]['NOMBRE_FACULTAD']).values
    faculty_name_str = numpy.array_str(faculty_name_nparray)
    faculty_name_str_final = faculty_name_str[2:len(faculty_name_str)-2]
    if faculty_name_str_final != '':
        Messagebox.show_info('El código corresponde a: ' + faculty_name_str_final, title='Información facultad')
    else:
        Messagebox.show_error('Ingrese el código de facultad.', title='Error!')
        
#-----------------------------------------------------------------------------------------------------------

#Name field:
label_config('Nombre', ('Samanata', 11), 10, 20)
entry_name = ttkbootstrap.Entry(window, bootstyle='primary')
entry_name.place(x=100, y=20)

#Last name field:
label_config('Apellido', ('Samanata', 11), 10, 60)
entry_last_name = ttkbootstrap.Entry(window, bootstyle='primary')
entry_last_name.place(x=100, y=60)

#Gender field:
label_config('Genero', ('Samanata', 11), 10, 100)
gender = ['FEMENINO', 'MASCULINO']
combobox_gender = ttkbootstrap.Combobox(window, values=gender, bootstyle='primary')
combobox_gender.place(x=100, y=100)

#Birthday field:
label_config('Fecha Nac.', ('Samanata', 11), 10, 140)
entry_birthday = ttkbootstrap.Entry(window, bootstyle='primary')
entry_birthday.place(x=100, y=140)

#Phone field:
label_config('Teléfono', ('Samanata', 11), 10, 180)
entry_phone = ttkbootstrap.Entry(window, bootstyle='primary')
entry_phone.place(x=100, y=180)

#Email field:
label_config('Email', ('Samanata', 11), 10, 220)
entry_email = ttkbootstrap.Entry(window, bootstyle='primary')
entry_email.place(x=100, y=220)

#Faculty code:
label_config('Código facultad', ('Samanata', 11), 10, 260)
faculty_sheet = pandas.read_excel('/home/bruce/Escritorio/Python/GUI/BD_SCHOOL_2023.xlsx', sheet_name='facultad')
code_faculty = faculty_sheet['CODIGO_FACULTAD']
code_faculty_list = list(code_faculty)
combobox_code_faculty = ttkbootstrap.Combobox(window, values=code_faculty_list, bootstyle='primary')
combobox_code_faculty.place(x=140, y=260)

#Load button:
load_button = ttkbootstrap.Button(text='Subir', command=load_data, bootstyle='success').place(x=90, y=325)

#Faculty button:
load_button = ttkbootstrap.Button(text='Facultad', command=faculty_check, bootstyle='warning').place(x=190, y=325)

window.mainloop()