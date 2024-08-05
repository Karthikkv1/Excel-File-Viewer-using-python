from tkinter import *
from tkinter import messagebox
from tkinter import ttk,filedialog
import numpy
import pandas as pd

root=Tk()
root.title("Excel Datasheet Viewer")
root.geometry('1100x400+200+200')

def Open():
    filename = filedialog.askopenfilename(title="Open a File",filetype=(("xlxs files",".xlsx"),("All Files",".")))

    if filename:
        try:
            filename = r"{}".format(filename)
            df = pd.read_excel(filename)

        except:
            messagebox.showerror("Error","You can't access this file!")  

    # now we have to clear previous data to enter new data
    tree.delete(*tree.get_children())

    # Datasheet heading
    tree['column'] = list(df.columns)      
    tree['show'] = "headings"

    #heading title
    for col in tree['column']:
        tree.heading(col,text=col) 

    # Enter data     
    df_rows =df.to_numpy().tolist()
    for row in df_rows:
        tree.insert("","end",values=row)  

# Icon image
image_icon=PhotoImage(file="logo.png")
root.iconphoto(False,image_icon)

# Frame
frame=Frame(root)
frame.pack(pady=25)

# TreeView
tree = ttk.Treeview(frame)
tree.pack()

# Button
button = Button(root,text='Open',width=60,height=2,font=30,fg="white",bg="#0078d7",command=Open)
button.pack(padx=10,pady=20)


root.mainloop()