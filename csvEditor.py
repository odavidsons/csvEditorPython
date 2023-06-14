"""
Application for managing CSV files.
Features:
-Import CSV file and display the data in a table
-Edit the data in the table and export the new file

"""
import csv
import tkinter as tk
from tkinter import filedialog as fd

#Open a file selection box
def chooseFile():
    global filename
    filetypes = (
        ('CSV files', '*.csv'),
        ('All files', '*.*')
    )
    filename = fd.askopenfilename(
        title='Open a file',
        initialdir='/home/dsantos/Documentos',
        filetypes=filetypes)

    importFileData(filename)

#Read the imported CSV file and render the data into the datatable
def importFileData(filename):
    global tableCells,header
    
    if len(tableCells) == 0:
        #Render the table canvas
        canvas.create_window(0, 0, window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=vscrollbar.set)
        canvas.configure(xscrollcommand=hscrollbar.set)

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
    else: errorLabel.config(text="Clear the loaded file first!")

def exportFileData(filename):
    if filename != "":
        new_tableCells = []
        #Update the new_tableCells array with the current values on the input boxes of the cells
        canvas_cells = canvas.find_all
        for i in range(len(tableCells)):
            new_tableCells.append([])
            for j in range(len(tableCells[i])):
                new_tableCells[i].append([])
                new_tableCells[i][j] = tableCells[i][j].get()
        saveFile = open(filename,'w',newline='')
        writer = csv.writer(saveFile)
        writer.writerow(header)
        for i in range(len(new_tableCells)):
            writer.writerow(new_tableCells[i])
        print(new_tableCells)
        tk.messagebox.showinfo("Information","File exported successfully!")
    else: errorLabel.config(text="You haven't imported a file yet!",foreground="red")

def clearTable():
    global tableCells

    canvas.delete('all')
    tableCells = []
    errorLabel.config(text="Table cleared!")

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
filemenu.add_command(label="Import",command=chooseFile)
filemenu.add_command(label="Export",command=lambda:exportFileData(filename))
menubar.add_cascade(label="File", menu=filemenu)
menubar.add_command(label="Clear",command=clearTable)
menubar.add_command(label="Exit",command=master.quit)
master.config(menu=menubar)

topLabel = tk.Label(master,text="Made by David Santos")
topLabel.grid(row=0,column=0)

#Render container and canvas for the datatable
container = tk.Frame(master)
canvas = tk.Canvas(container, width=1000, height=600)
vscrollbar = tk.Scrollbar(container, orient="vertical", command=canvas.yview)
hscrollbar = tk.Scrollbar(container, orient="horizontal", command=canvas.xview)
scrollable_frame = tk.Frame(canvas)

#Bind the scrollregion to the frame
scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )

#Top labels/buttons
errorLabel = tk.Label(master,text="")
errorLabel.grid(row=1,column=0)

container.grid()
canvas.grid()

#Set position of scrollbars
vscrollbar.grid(row=2, column=1, sticky="ns")
hscrollbar.grid(row=3, column=0, sticky="ew")

master.mainloop()