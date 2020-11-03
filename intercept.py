import xlrd
import time
import ntpath
import pyautogui
from pyscreeze import ImageNotFoundException
from tkinter import *
from tkinter import ttk, filedialog, messagebox


class FileView:

    def __init__(self, parent):
        self.parent = parent
        self.frame = Frame(parent).place(relwidth=1, relheight=1)
        self.upload_label = Label(
            self.frame, text="Upload Excel file.", font=("Helvetica", 13))
        self.upload_label.place(x=10, y=10)
        self.upload_button = Button(self.frame, text="Upload", font=(
            "Helvetica", 13), command=self.destroy_widget)
        self.upload_button.place(width=180, height=50, x=10, y=50)

    def inside_catia(self, data):
        # check if catia document is opened
        try:
            axis = pyautogui.locateCenterOnScreen('images/catia.PNG')
            print(axis)
        except ImageNotFoundException:
            print("Catia not opened.")

    def cancelOperation(self):
        if messagebox.askokcancel("Cancel", "This operation will be cancelled."):
            FileView(app)

    def toCatia(self):
        notice.destroy()
        notice_button.destroy()
        self.parent.geometry("650x200")
        self.parent.resizable(False, False)
        new_notice_label = Label(
            self.frame, text="Program running...", fg="green", font=("Helvetica", 13))
        new_notice_label.place(x=10, y=10)
        new_notice_message = Message(
            self.frame, width=600, font=("Helvetica", 13), text="- Don't use the mouse/keyboard during this operation.\n- You can cancel anytime.\n")
        new_notice_message.place(x=10, y=50)
        new_notice_button = Button(
            self.frame, text='Cancel', font=("Helvetica", 13), command=self.cancelOperation)
        new_notice_button.place(x=500, y=150)
        new_notice_progressText = Label(
            self.frame, text="Status", font=("Helvetica", 13))
        new_notice_progressText.place(x=10, y=120)
        new_notice_progressBar = ttk.Progressbar(
            self.frame, orient=HORIZONTAL, length=300, mode="determinate")
        new_notice_progressBar.place(x=100, y=125)

    def notice(self):
        global notice
        global notice_button
        self.parent.geometry("600x200")
        self.parent.resizable(False, False)
        notice = Message(
            self.frame, width=600, font=("Helvetica", 13), text="IMPORTANT NOTICE.\n\n- Make sure the required 3D is open in Catia Composer.\n- Make sure no other apps block Catia Composer's screen.\n- Click 'OK' to continue.")
        notice.place(x=10, y=10)
        notice_button = Button(self.frame, text='Ok', font=(
            "Helvetica", 13), command=self.toCatia)
        notice_button.place(x=450, y=150)

    def destroy_widget(self):
        self.upload_label.destroy()
        self.upload_button.destroy()
        self.excel_file()

    def excel_file(self):
        filepath = filedialog.askopenfilename(
            initialdir="/Desktop", title='Select File', filetypes=[("Excel files", ".xlsx .xls .csv")])
        if not filepath:
            FileView(app)
        else:
            filename = ntpath.basename(filepath)
            if filename.lower().endswith(('.xlsx', '.xls', '.csv')):
                upload_progress = Label(
                    self.frame, text="File uploading", font=("Helvetica", 13))
                row_progress = ttk.Progressbar(
                    self.frame, orient=HORIZONTAL, length=300, mode="determinate")
                # open excel and sheet
                workbook = xlrd.open_workbook(filepath)
                sheet = workbook.sheet_by_index(0)

                # Stores Position and Partcode value pairs from all rows.
                all_data = []

                # loop over rows and columns to extract Position and Partcode pairs
                for row in range(sheet.nrows):
                    current_row = []
                    upload_progress.pack(side='left', padx=10)
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
                    messagebox.showerror('Error!', f"{filename} is empty.")
                    self.excel_file()
                else:
                    row_progress.destroy()
                    upload_progress.destroy()
                    messagebox.showinfo(
                        'Success', 'FILE UPLOADED\n\nClick "OK" to continue.')
                    self.notice()

            else:
                messagebox.showerror(
                    'Error!', f"{filename} is not an Excel file.")
                self.excel_file()

    def onClosing():
        if messagebox.askokcancel("Quit", "Do you really wish to quit?"):
            app.destroy()


def main():
    global app
    app = Tk()
    app.resizable(False, False)
    app.wm_iconbitmap('images/intercept.ico')
    app.wm_title('INTERCEPT')
    app.geometry("600x120")
    application = FileView(app)
    app.wm_protocol("WM_DELETE_WINDOW", FileView.onClosing)
    app.mainloop()


if __name__ == "__main__":
    main()
