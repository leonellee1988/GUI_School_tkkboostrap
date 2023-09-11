import tkinter
from tkinter import *
import ttkbootstrap
from ttkbootstrap.dialogs import Messagebox
import pandas

#Main window configuration:
window = ttkbootstrap.Window(themename='minty')
window.title('Creación de sección-materia')
window.geometry('525x650')
window.resizable()

#My functions:
def label_config(main, text_label, font_label, x_label, y_label):
    """
    This function creates ttkbootstrap labels considering the text,
    font and the position of the label as arguments.
    
    >>>lable_config('Hello, I am a label', ('Helvetica', 28), 10, 10)
    """
    ttkbootstrap.Label(main, text=text_label, font=font_label).place(x=x_label, y=y_label)
    
def load_data_dataset(sheet):
    """
    The function load the data from an Excel's sheet, the user is able to choose the sheet.
    """
    try:
        data_sheet = pandas.read_excel('/home/bruce/Escritorio/Python/GUI/BD_SCHOOL_2023.xlsx', sheet_name=sheet)
        return data_sheet
    except FileNotFoundError:
        print('Error. No se ha encontrado el archivo')
        return None

def show_faculty(event):
    """
    This function allows the user:
    
    (a)Reads the name of the faculty according to the code entered.
    (b)Resets comboboxes view when the user selects other code.
    (c) Activates 'get_data_combobox_career' function.
    """
    code_faculty_combobox = faculty_combobox.get()
    
    career_combobox.set('')     #when the user choose other faculty code, the career's view combobox reset
    subject_combobox.set('')    #when the user choose other faculty code, the subjects's view combobox reset
    
    if faculty_sheet is not None:
        faculty_selected = faculty_sheet[faculty_sheet['CODIGO_FACULTAD'] == code_faculty_combobox]     #creates a dataframe
        if not faculty_selected.empty:      
            faculty_name = faculty_selected['NOMBRE_FACULTAD'].iloc[0] #gets the name of the faculty, in the dataframe, related to the code
            faculty_label.config(text=faculty_name) #displays the name of the faculty in the label
            career_label.config(text='')        #resets the information of the labels
            bachelor_label.config(text='')      #resets the information of the labels
            subject_label.config(text='')       #resets the information of the labels
            get_data_combobox_career(code_faculty_combobox)
        else:
            faculty_label.config(text='Facultad no encontrada')
    else:
        faculty_label.config(text='Datos de facultad no disponible')
        
def get_data_combobox_career(faculty_code):
    """
    The function allows the user:
    
    (a) Setups comboboxes value into a empty array, if the faculty code did not match with any data.
    (b) If the faculty code matches with a data, creates and assings a value list to the combobox.
    """
    career_selected = career_sheet[career_sheet['CODIGO_FACULTAD'] == faculty_code]     #creates a dataframe
    if career_selected.empty:
        career_combobox['values'] = []      #setups career combobox value when faculty code didn't match with any data
        subject_combobox['values'] = []     #setups subject combobox values when faculty code didn't match with any data
    else:
        career_list = career_selected['CODIGO_CARRERA'].tolist()    #creats a list with the values of the dataframe
        career_combobox['values'] = career_list                     #assigns the list to the combobox values

def show_career(event):
    """
    This function allows the user:
    
    (a)Reads the name of the career according to the code entered.
    (b)Setups the comboboxes according to the codes.
    (c) Activates 'get_data_combobox_subject' function.
    """
    code_career_combobox = career_combobox.get()
    
    subject_combobox.set('')    #when the user choose other career code, the subjects's view combobox reset
    
    if career_sheet is not None:
        career_selected = career_sheet[career_sheet['CODIGO_CARRERA'] == code_career_combobox]      #creates a dataframe
        if not career_selected.empty:
            career_name = career_selected['NOMBRE_CARRERA'].iloc[0]         #gets the name of the career, in the dataframe, related to the code
            bachelor_name = career_selected['GRADO_CARRERA'].iloc[0]        #gets the name of the bachelor career, in the dataframe, related to the code
            career_label.config(text=career_name)      #displays the name of the career in the label
            bachelor_label.config(text=bachelor_name)   #displays the name of the bachelor career in the label
            subject_label.config(text='')                                   #resets the information of the labels
            get_data_combobox_subject(code_career_combobox)
        else:
            career_label.config(text='Carrera no encontrada')
    else:
        career_label.config(text='Datos de carrera no disponible')
        print('empty')

def get_data_combobox_subject(career_code):
    """
    The function allows the user:
    
    (a)Creates and assings a value list to the combobox.
    """
    subject_selected = subject_sheet[subject_sheet['CODIGO_CARRERA'] == career_code]    #creates a dataframe
    subject_list = subject_selected['CODIGO_MATERIA'].tolist()          #creats a list with the values of the dataframe
    subject_combobox['values'] = subject_list                           #assigns the list to the combobox values

def show_subject(event):
    """
    This function allows the user:
    
    (a)Reads the name of the subject according to the code entered.
    """
    code_subject_combobox = subject_combobox.get()
    
    if subject_sheet is not None:
        subject_selected = subject_sheet[subject_sheet['CODIGO_MATERIA'] == code_subject_combobox]      #creates a dataframe
        if not subject_selected.empty:
            subject_name = subject_selected['NOMBRE_MATERIA'].iloc[0]           #gets the name of the subject, in the dataframe, related to the code
            subject_label.config(text=subject_name)     #resets information of the label   
        else:
            subject_label.config(text='Materia no encontrada')
    else:
        subject_label.config(text='Datos de materia no disponible')

def load_student():
    
    row_number = 1
    code_career = career_combobox.get()
    
    if not code_career:
        Messagebox.show_error('Seleccione una carrera', title='Error!')
        return

    student_selected = student_sheet[student_sheet['CODIGO_CARRERA'] == code_career]
    
    if not student_selected.empty:
        student_table.delete(*student_table.get_children())
        for i, row in student_selected.iterrows():
            student_code = row['CODIGO_ALUMNO']
            student_name = row['NOMBRE_ALUMNO']
            student_last_name = row['APELLIDO_ALUMNO']
            student_mail = row['EMAIL_ALUMNO']
            student_table.insert('', 'end', values=(row_number, student_code, student_name, student_last_name, student_mail))
            row_number += 1
    else:
        Messagebox.show_info('No se encontraron estudiantes para esta carrera', title='Atención!')

#-----------------------------------------------------------------------------------------------------------

#Main frames

frame_entry = tkinter.LabelFrame(window, text='Ingreso de información')
frame_entry.place(x=10, y=10, width=500, height=175)

#Faculty information
label_config(frame_entry, 'Código facultad', 'Helvetica', 10, 20)
faculty_sheet = load_data_dataset('facultad')
code_faculty = faculty_sheet['CODIGO_FACULTAD']
code_faculty_list = list(code_faculty)
faculty_combobox = ttkbootstrap.Combobox(frame_entry, values=code_faculty_list)
faculty_combobox.place(x=150, y=20)
faculty_combobox.bind('<<ComboboxSelected>>', show_faculty)

#Career information
label_config(frame_entry, 'Código carrera', 'Helvetica', 10, 60)
career_sheet = load_data_dataset('carrera')
career_combobox = ttkbootstrap.Combobox(frame_entry, values=[])
career_combobox.place(x=150, y=60)
career_combobox.bind('<<ComboboxSelected>>', show_career)

#Subject information
label_config(frame_entry, 'Código materia', 'Helvetica', 10, 100)
subject_sheet = load_data_dataset('materia')
subject_combobox = ttkbootstrap.Combobox(frame_entry, values=[])
subject_combobox.place(x=150, y=100)
subject_combobox.bind('<<ComboboxSelected>>', show_subject)

#Data display
faculty_label_1 = ttkbootstrap.Label(window, text='Facultad:', font=('Helvetica'))
faculty_label_1.place(x=10, y=200)
faculty_label = ttkbootstrap.Label(window, text='', font=('Helvetica'))
faculty_label.place(x=100, y=200)

career_label_1 = ttkbootstrap.Label(window, text='Carrera:', font=('Helvetica'))
career_label_1.place(x=10, y=225)
career_label = ttkbootstrap.Label(window, text='', font=('Helvetica'))
career_label.place(x=100, y=225)

bachelor_label_1 = ttkbootstrap.Label(window, text='Grado:', font=('Helvetica'))
bachelor_label_1.place(x=10, y=250)
bachelor_label = ttkbootstrap.Label(window, text='', font=('Helvetica'))
bachelor_label.place(x=100, y=250)

subject_label_1 = ttkbootstrap.Label(window, text='Asignatura:', font=('Helvetica'))
subject_label_1.place(x=10, y=275)
subject_label = ttkbootstrap.Label(window, text='', font=('Helvetica'))
subject_label.place(x=100, y=275)

#Load student data
student_sheet = load_data_dataset('alumno')
button_load = ttkbootstrap.Button(text='Cargar', command=load_student)
button_load.place(x=225, y=325)

style = ttkbootstrap.Style()
style.configure('Treeview.Heading', background='#78c2ad', foreground='white')
student_table = ttkbootstrap.Treeview(window, columns=('No', 'Codigo', 'Nombre', 'Apellido', 'Email'), show='headings')
student_table.heading('No', text='No')
student_table.heading('Codigo', text='Codigo')
student_table.heading('Nombre', text='Nombre')
student_table.heading('Apellido', text='Apellido')
student_table.heading('Email', text='Email')
student_table.column('No', width=50)
student_table.column('Codigo', width=100)
student_table.column('Nombre', width=100)
student_table.column('Apellido', width=100)
student_table.column('Email', width=150)
student_table.place(x=10, y=375, height=250)

window.mainloop()