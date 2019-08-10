import tkinter as tk
import pandas as pd

fields = ('Excel location', 'Excel sheet name', 'Header row', 'Columns to include', 'Output CSV name', 'CSV Delimiter')

entries = {}

#https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.DataFrame.to_csv.html
#https://pandas.pydata.org/pandas-docs/stable/reference/api/pandas.read_excel.html
def convert():
    fileName = entries['Excel location'].get().strip()
    sheetName = entries['Excel sheet name'].get().strip()
    colsToUse = entries['Columns to include'].get().strip()
    startRow = entries['Header row'].get().strip()
    startRow = int(startRow) - 1

    df = pd.read_excel(open(fileName, 'rb'), sheet_name=sheetName, header=startRow, usecols=colsToUse)
    df.dropna(how='all', inplace=True)
    df.dropna(axis=1, how='all', inplace=True)

    csvFileName = entries['Output CSV name'].get().strip()
    df.to_csv(csvFileName, sep = entries['CSV Delimiter'].get().strip() ,header=True )
    print("SUCCESS: Conversion successful")
        
def makeform(root, fields):
    for field in fields:
        print(field)
        row = tk.Frame(root)
        lab = tk.Label(row, width=22, text=field+": ", anchor='w')
        ent = tk.Entry(row)
        ent.insert(0, "")
        row.pack(side=tk.TOP, 
                 fill=tk.X, 
                 padx=5, 
                 pady=5)
        lab.pack(side=tk.LEFT)
        ent.pack(side=tk.RIGHT, 
                 expand=tk.YES, 
                 fill=tk.X)
        entries[field] = ent
    return entries

if __name__ == '__main__':
    root = tk.Tk()
    root.title("Excel to csv converter")
    ents = makeform(root, fields)

    b1 = tk.Button(root, text='Convert',
           command=convert)
    b1.pack(side=tk.LEFT, padx=5, pady=5)
    
    b3 = tk.Button(root, text='Quit', command=root.quit)
    b3.pack(side=tk.LEFT, padx=5, pady=5)
    root.mainloop()