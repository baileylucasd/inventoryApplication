# Import the required library
import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import filedialog as fd
from tkinter.messagebox import showinfo
import pandas as pd



root = tk.Tk()
root.title('Inventory Comparison Tool')
root.resizable(False, False)
root.geometry('700x700')


def select_file():
    filetypes = (
        ('Excel Files', '*.xlsx'),
        ('All files', '*.*')
    )

    filename = fd.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=filetypes)

    showinfo(
        title='Process Ran',
        message=filename
    )

    df = pd.read_excel(filename)

    recorded = df["Recorded Tag"]
    expected = df["Asset tag"]

    common_values = [element for element in recorded.values if element in expected.values]
    missing_expected_values = [element for element in expected.values if element not in recorded.values]
    unexpected_values = [element for element in recorded.values if element not in expected.values]
    unexpected_values = [element for element in unexpected_values if str(element) != 'nan']
    score = "{:.2f}".format(len(common_values)/len(expected)*100)

    lines = ['Expected assets not found during inventory: ', "Total missing: {}".format(len(missing_expected_values)), "\n", 'Unexpected assets found during inventory: ', "Total unexpected: {}".format(len(unexpected_values)), "Accuracy Score: ", score+"%\n"]

    with open('results.txt', 'w') as f:
        f.write(lines[0]+'\n')
        f.write('\n'.join(missing_expected_values))
        f.write('\n\n'+lines[1])
        f.write('\n\n'+lines[2])
        f.write('\n\n'+lines[3]+'\n')
        f.write('\n'.join(unexpected_values))
        f.write('\n\n'+lines[4])
        f.write('\n\n'+lines[5])
        f.write('\n'+lines[6])
    
instructions = ttk.Label(root, text="After downloading the expected asset excel sheet from SNOW for a specific location, copy and paste all assets recorded during inventory into a column of the downloaded sheet and name that column 'recorded tag', then select the sheet with the button below. Results will be output to 'results.txt' in the same folder as the selected sheet, rename to *nameOfBuilding*.txt for each respective building.",
                         font=("Arial", 18), wraplength=600)
instructions.pack(pady=100)

open_button = ttk.Button(
    root,
    text='Open File',
    command=select_file
)

open_button.pack(expand=True)

root.mainloop()