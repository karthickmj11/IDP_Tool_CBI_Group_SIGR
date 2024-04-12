from msilib.schema import CheckBox
from sre_parse import State
import tkinter as tk
from tkinter import ttk
from tkinter import FLAT, GROOVE, RAISED, RIDGE, SUNKEN, filedialog
from tracemalloc import Statistic
import openpyxl
import tkinter as tk
from tkinter import filedialog
import openpyxl
import All_in_One_Master
import Group_Master
import CBI_Master
import threading as th
import Group
loc = ""
#global run
com_sheets = ['TC','Subroute','TD','ESP','Point','Signal','Overlap','PSAPR','Cycle','MBL','Route','Switch','ASCV','Interlocking','Berth','TP']
def browse_file():
    global loc
    status_entry.delete(0, tk.END)
    filename = filedialog.askopenfilename()
    file_name_entry.delete(0, tk.END)
    file_name_entry.insert(0, filename)
    loc = str(filename)
    #print(file_name_entry)
    
    try:
        w = openpyxl.load_workbook(loc)
    except openpyxl.utils.exceptions.InvalidFileException:
        status_entry.delete(0, tk.END)
        status_entry.insert(1, "Error: The selected file is not a valid Excel file.")
        return 0
    sheet = [str(i)[12:len(str(i))-2] for i in w.worksheets]
    #print(type(sheet[0]),sheet)
    if(filename !=""):
        if(len(sheet)>=len(com_sheets) and set(sheet)>=set(com_sheets)):
            browse_label.config(state ='normal')
            browse_entry.config(state ='normal')
            op_button.config(state ='normal')
        else:
              status_entry.delete(0, tk.END)
              status_entry.insert(1, "These sheets are missing or naming may be wrong"+str(list(set(com_sheets)-set(sheet))))
def Group_Mas(loc,op_l):
    Group_Master.run(loc,op_l) 
    
def CBI_Mas(loc,op_l):
    CBI_Master.run(loc,op_l) 
 
def browse_folder():
    global op_l
    status_entry.delete(0, tk.END)
    filename = filedialog.askdirectory()
    browse_entry.delete(0, tk.END)
    browse_entry.insert(0, filename)
    if(filename !=""):
        for checkbox in checkboxes:
            checkbox.config(state='normal')
        select_all_button.config(state='normal')
        deselect_all_button.config(state='normal')
        run_backend_button.config(state='normal')
        op_l = str(filename)
    else:
        status_entry.delete(0, tk.END)
        status_entry.insert(1, "Please Select a folder to continue")

def select_all():
    for checkbox in checkboxes:
        checkbox.select()
    run_backend_button.config(state='normal')

def deselect_all():
    for checkbox in checkboxes:
        checkbox.deselect()
    run_backend_button.config(state='disable')
def Run():
    progress["value"] += 0
    l= []
    for i in checkbox_vars:
        l.append(i.get())
    if(l.count(0)==3):
        status_entry.delete(0, tk.END)
        status_entry.insert(1, "Please select atleast one checkbox")
        return 1
    status_entry.insert(1, "Group Files Started....")
    root.update()
    t = th.Thread(target=Group_Mas,args=(loc,op_l))
    t.start()
    #Group_Master.run(loc,op_l)
    progress["value"] += (1/2)*100
    status_entry.delete(0, tk.END)
    status_entry.insert(1, "CBI Files Started....")
    root.update()
    t1 = th.Thread(target = CBI_Mas,args=(loc,op_l))
    t1.start()
    #CBI_Master.run(loc,op_l)
    #print(All_in_One_Master.run(loc,op_l,l))
    status_entry.delete(0, tk.END)
    progress["value"] += (1/2)*100
    status_entry.insert(1, "All Completed!!!!")
    

checkbox_vars = []
root = tk.Tk()
root.geometry("900x400")
root.configure(bg='#ADD8E6')
root.title("U200 IDP")
root.iconbitmap(r"C:\Users\491497\source\repos\All_in_One\Als.ico")

# create a PhotoImage object from an image file
image = tk.PhotoImage(file=r"C:\Users\491497\source\repos\All_in_One\Als.png")
re_im = image.subsample(4, 4)
# create a Label widget with the image
label = tk.Label(root, image=re_im, bg='#ADD8E6')
label.place(x=840, y=240)

file_name_label = tk.Label(root, text="File name:", bg='#ADD8E6', font=("Helvetica", 16),fg='#000')
file_name_label.grid(row=0, column=0, padx=10, pady=10)

file_name_entry = tk.Entry(root, font=("Helvetica", 16),bg='#fff', fg='#000', relief='solid',bd=1,width=40)
file_name_entry.grid(row=0, column=1, padx=10, pady=10)

browse_label = tk.Label(root, text="Open Folder:", bg='#ADD8E6', font=("Helvetica", 16),fg='#000',state = "disabled")
browse_label.grid(row=1, column=0, padx=10, pady=10)

browse_entry = tk.Entry(root, font=("Helvetica", 16),bg='#fff', fg='#000', relief='solid',bd=1,width=40,state = "disabled")
browse_entry.grid(row=1, column=1, padx=10, pady=10)


browse_button = tk.Button(root, text="Select Input", command=browse_file, font=("Helvetica", 16), bg='#87CEFA',fg='#000',relief=RIDGE)
browse_button.grid(row=0, column=2, padx=10, pady=10)

op_button = tk.Button(root, text="Select Output", command=browse_folder, font=("Helvetica", 16), bg='#87CEFA',fg='#000',relief=RIDGE,state = "disabled")
op_button.grid(row=1, column=2, padx=10, pady=10)
names = ['Group','CBI','Sigrule']
checkboxes = []
for i in range(3):
    checkbox = tk.Checkbutton(root, text=names[i], font=("Helvetica", 14),bg='#ADD8E6',fg='#000',relief=FLAT)
    checkbox.grid(row=2, column=i, padx=10, pady=10)
    checkboxes.append(checkbox)
    checkbox.config(state='disable')

for i in range(3):
        checkbox_vars.append(tk.IntVar())
        
for i, checkbox in enumerate(checkboxes):
        checkbox.config(variable=checkbox_vars[i])
        
select_all_button = tk.Button(root, text="Select All", command=select_all, font=("Helvetica", 14), bg='#87CEFA',fg='#000',relief=RIDGE)
select_all_button.grid(row=6, column=0, padx=10, pady=10)
select_all_button.config(state='disable')

deselect_all_button = tk.Button(root, text="Deselect All", command=deselect_all, font=("Helvetica", 14), bg='#87CEFA',fg='#000',relief=RIDGE)
deselect_all_button.grid(row=6, column=1, padx=10, pady=10)
deselect_all_button.config(state='disable')

run_backend_button = tk.Button(root, text="Generate Output",command = Run ,font=("Helvetica", 14), bg='#87CEFA',fg='#000',relief=RIDGE)
run_backend_button.grid(row=6, column=2, padx=10, pady=10)
run_backend_button.config(state='disable')

status_entry = tk.Entry(root, font=("Helvetica", 16),bg='#fff', fg='#000', relief='solid',bd=1,width=40)
#status_entry.config(height=5)
status_entry.grid(row=7, column=1, padx=10, pady=10)

progress = ttk.Progressbar(root, orient="horizontal", length=200, mode="determinate")
progress.grid(row=11, column=1, padx=10, pady=10)
root.mainloop()

