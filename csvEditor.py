"""
Application for managing CSV files.
Features:
-Import CSV file and display the data in a table
-Edit the data in the table and export the new file

"""
import csv
import tkinter as tk
from tkinter import filedialog as fd
import webbrowser

#Open a file selection box
def chooseFile():
    global filename
    filetypes = (
        ('CSV files', '*.csv'),
        ('All files', '*.*')
    )
    filename = fd.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=filetypes)
    if len(filename) > 0: importFileData(filename)

#Read the imported CSV file and render the data into the datatable
def importFileData(filename):
    global tableCells,header
    
    if len(tableCells) == 0:
        #Create the table canvas
        canvas.create_window(0, 0, window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=vscrollbar.set,xscrollcommand=hscrollbar.set)

        with open(filename, newline='') as csvfile:
            errorLabel.config(text="",foreground="red") #Clear error message
            spamreader = csv.reader(csvfile, delimiter=',', quotechar='|')
            errorLabel.config(text="",foreground="red")
            #Write headers
            header = next(spamreader)
            for i in range(len(header)):
                tableHeader = tk.Entry(scrollable_frame)
                tableHeader.grid(row=2, column=i)
                tableHeader.insert(0,header[i])
                tableHeader.config(state='disabled', disabledforeground='black')
            #Write csv data
            row_count = 0
            column_count = len(header)
            for row in spamreader:
                tableCells.append([])
                for j in range(len(row)):
                    tableCells[row_count].append([])
                    cell = tk.Entry(scrollable_frame)
                    cell.insert(tk.END,row[j])
                    cell.grid(row=row_count+3,column=j)
                    tableCells[row_count][j] = cell
                row_count = row_count+1
    else: 
        errorLabel.config(text="Clear the loaded file first!")
        tk.messagebox.showwarning(message="Clear the loaded file first!")

def exportFileData(filename):
    global header

    if len(canvas.find_all()) > 0:
        new_tableCells = []
        #Update the new_tableCells array with the current values on the input boxes of the cells
        for i in range(len(tableCells)):
            new_tableCells.append([])
            for j in range(len(tableCells[i])):
                new_tableCells[i].append([])
                new_tableCells[i][j] = tableCells[i][j].get()

        filetypes = (
        ('CSV files', '*.csv'),
        ('All files', '*.*')
        )
        filename = fd.asksaveasfile(mode='w',filetypes=filetypes)
        print(filename.name)
        saveFile = open(filename.name,'w',newline='')
        writer = csv.writer(saveFile)
        if len(header)>0:
            writer.writerow(header)
        for i in range(len(new_tableCells)):
            writer.writerow(new_tableCells[i])
        tk.messagebox.showinfo("Information","File exported successfully!")
    else: errorLabel.config(text="You haven't imported a file or created a table!",foreground="red")

#Create a window the inputing the value for the new table size
def newTableInput():
    global sizeSelection

    if len(canvas.find_all()) == 0:
        #Elements for the input information
        sizeSelection = tk.Toplevel(master)
        sizeSelection.minsize(200,100)
        label1 = tk.Label(sizeSelection,text="Select the table size")
        label1.grid(row=0,column=0,columnspan=2)
        labelWidth = tk.Label(sizeSelection,text="Width: ")
        labelWidth.grid(row=1,column=0)
        widthInput = tk.Entry(sizeSelection)
        widthInput.grid(row=1,column=1,padx=5,pady=5)
        labelHeight = tk.Label(sizeSelection,text="Height: ")
        labelHeight.grid(row=2,column=0)
        heightInput = tk.Entry(sizeSelection)
        heightInput.grid(row=2,column=1)
        confirmBtn = tk.Button(sizeSelection,text="Ok",command=lambda: newTableRender(heightInput.get(),widthInput.get()))
        confirmBtn.grid(row=3,column=0,columnspan=2)
    else: errorLabel.config(text="You have a table opened. Clear it first",foreground="red")

#Render a new table with the specified size
def newTableRender(height,width):
    global tableCells,header
    
    try:
        tableCells = []
        header = []
        #Create the table canvas
        canvas.create_window(0, 0, window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=vscrollbar.set,xscrollcommand=hscrollbar.set)
        #Render the cells
        for i in range(int(height)):
            tableCells.append([])
            for j in range(int(width)):
                tableCells[i].append([])
                cell = tk.Entry(scrollable_frame)
                cell.grid(row=i+3,column=j)
                tableCells[i][j] = cell
        sizeSelection.destroy()
    except ValueError: tk.messagebox.showwarning(message="Please enter a valid number")

def clearTable():
    global tableCells,header

    canvas.delete("all")
    for widgets in scrollable_frame.winfo_children():
        widgets.destroy()
    tableCells = []
    header = []
    errorLabel.config(text="Table cleared!",foreground="black")

def callback(url):
    webbrowser.open_new(url)

#Main program
filename = ""
header = []
tableCells = []
master = tk.Tk()
master.resizable(True,True)
master.title("CSV Editor")

#Menu bar
menubar = tk.Menu(master)
filemenu = tk.Menu(menubar, tearoff=0)
filemenu.add_command(label="New",command=newTableInput)
filemenu.add_command(label="Import",command=chooseFile)
filemenu.add_command(label="Export",command=lambda:exportFileData(filename))
menubar.add_cascade(label="File", menu=filemenu)
menubar.add_command(label="Clear",command=clearTable)
menubar.add_command(label="Exit",command=master.quit)
master.config(menu=menubar)

topLabel = tk.Label(master,text="Made by: David Santos",foreground="blue",cursor="hand2")
topLabel.grid(row=0,column=0)
topLabel.bind("<Button-1>", lambda e: callback("https://github.com/odavidsons"))

#Render container and canvas for the datatable
container = tk.Frame(master)
canvas = tk.Canvas(container, width=1000, height=600)
vscrollbar = tk.Scrollbar(master, orient="vertical", command=canvas.yview)
hscrollbar = tk.Scrollbar(master, orient="horizontal", command=canvas.xview)
scrollable_frame = tk.Frame(canvas)

#Bind the scrollregion to the frame
scrollable_frame.bind("<Configure>",lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

#Top labels/buttons
errorLabel = tk.Label(master,text="")
errorLabel.grid(row=1,column=0)

container.grid()
canvas.grid()

#Set position of scrollbars
vscrollbar.grid(row=2, column=1, sticky="ns")
hscrollbar.grid(row=3, column=0, sticky="ew")

master.mainloop()