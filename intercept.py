import xlrd
import time
import ntpath
import pyautogui
from pyscreeze import ImageNotFoundException
from tkinter import *
from tkinter import ttk,filedialog,messagebox

class FileView:

    def __init__(self,parent):
        self.parent = parent
        self.frame = Frame(parent).place(relwidth=1,relheight=1)
        self.upload_label = Label(self.frame,text="Upload Excel file.",font=("Helvetica", 13))
        self.upload_label.place(x=10,y=10)
        self.upload_button = Button(self.frame,text="Upload",font=("Helvetica", 13),command = self.destroy_widget)
        self.upload_button.place(width=180,height=50,x=10,y=50)

    def inside_catia(self,data):
        # check if catia document is opened
        try:
            axis = pyautogui.locateCenterOnScreen('images/catia.PNG')
            print(axis)
        except ImageNotFoundException:
            print("Catia not opened.")  

    def destroy_widget(self):
        self.upload_label.destroy()
        self.upload_button.destroy()
        self.excel_file()

    def excel_file(self):
        filepath = filedialog.askopenfilename(initialdir="/Desktop",title='Select File',filetypes=[("Excel files", ".xlsx .xls .csv")])
        if not filepath:
            FileView(app)
        else:
            filename = ntpath.basename(filepath)
            if filename.lower().endswith(('.xlsx', '.xls', '.csv')):
                upload_progress = Label(self.frame,text="File uploading",font=("Helvetica", 13))
                row_progress = ttk.Progressbar(self.frame,orient= HORIZONTAL,length=300,mode="determinate")
                # open excel and sheet
                workbook = xlrd.open_workbook(filepath)
                sheet = workbook.sheet_by_index(0)

                # Stores Position and Partcode value pairs from all rows.
                all_data = []

                # loop over rows and columns to extract Position and Partcode pairs
                for row in range(sheet.nrows):
                    current_row = []
                    upload_progress.pack(side='left',padx=10)
                    row_progress.pack(side='left')
                    row_progress['value'] += 100 / sheet.nrows
                    self.parent.update_idletasks()
                    time.sleep(0.01)
                    for col in range(sheet.ncols):
                        if col == 1 or col == 3:
                            data = sheet.cell_value(row, col)
                            # useful when searching in Catia Composer
                            new_data = data.replace('/', "_")
                            current_row.append(new_data)
                    all_data.append(current_row)
                if not len(all_data):
                    messagebox.showerror('Error!',f"{filename} is empty.")
                    self.excel_file()
                else:        
                    upload_progress.config(text='File uploaded successfully.', fg='green')
                    row_progress.destroy()
                    print(all_data)
                    return all_data

            else:
                messagebox.showerror('Error!',f"{filename} is not an Excel file.")
                self.excel_file()

def main():
    global app
    app = Tk()
    app.wm_iconbitmap('images/intercept.ico')
    app.wm_title('INTERCEPT')
    app.geometry("600x120")
    application = FileView(app)
    app.mainloop()

if __name__ == "__main__":
    main()
    
    
