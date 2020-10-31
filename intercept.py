import xlrd
import ntpath
import pyautogui
from pyscreeze import ImageNotFoundException
from tkinter import *
from tkinter import filedialog

def inside_catia(data):
    # check if catia document is opened
    try:
        axis = pyautogui.locateCenterOnScreen('images/catia.PNG')
        print(axis)
    except ImageNotFoundException:
        print("Catia not opened.")  

def excel_file():
    filepath = filedialog.askopenfilename(initialdir="/Desktop",title='Select File',filetypes=[("Excel files", ".xlsx .xls .csv")])
    filename = ntpath.basename(filepath)
    if filename.endswith(('.xlsx', '.xls', '.csv')):
        # open excel and sheet
        workbook = xlrd.open_workbook(filepath)
        sheet = workbook.sheet_by_index(0)

        # Stores Position and Partcode value pairs from all rows.
        all_data = []

        # loop over rows and columns to extract Position and Partcode pairs
        for row in range(sheet.nrows):
            current_row = []
            for col in range(sheet.ncols):
                if col == 1 or col == 3:
                    data = sheet.cell_value(row, col)
                    # useful when searching in CC
                    new_data = data.replace('/', "_")
                    current_row.append(new_data)
            all_data.append(current_row)

        # inside_catia(all_data)
        print(all_data)
        return filename

    else:
        print(f'{filename} is not excel file.')
        return filename
def main():
    upload_label = Label(frame,text="Upload Excel file.",font=("Helvetica", 15)).place(x=10,y=10)
    upload_button = Button(frame,text="Upload",font=("Helvetica", 15),command = excel_file).place(width=180,height=50,x=10,y=50)
            

if __name__ == "__main__":
    app = Tk()
    app.wm_iconbitmap('images/intercept.ico')
    app.wm_title('INTERCEPT')
    canvas = Canvas(app,height=120, width=600).pack()
    frame = Frame(app).place(relwidth=1,relheight=1)
    main()
    app.mainloop()
    
    
