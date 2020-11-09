import xlrd
import time
import ntpath
import pyautogui
from tkinter import *
from tkinter import ttk, filedialog, messagebox
import pyperclip as pc


class App:

    def __init__(self, parent):
        self.parent = parent
        self.frame = Frame(parent).place(relwidth=1, relheight=1)
        self.upload_label = Label(
            self.frame, text="Upload Excel file.", font=("Helvetica", 13))
        self.upload_label.place(x=10, y=10)
        self.upload_button = Button(self.frame, text="Upload", font=(
            "Helvetica", 13), command=self.destroy_widget)
        self.upload_button.place(width=180, height=50, x=10, y=60)
    # def searchPartcode(self):
    #     time.sleep(1.5)

    def retryCatia(self):
        time.sleep(1)
        thumb = pyautogui.locateCenterOnScreen('images/thumb.png')
        thumbnail = pyautogui.locateCenterOnScreen('images/thumbnail.png')
        if thumb or thumbnail is not None:
            pyautogui.moveTo(thumb[0], thumb[1])
            pyautogui.moveTo(thumbnail[0], thumbnail[1])
            pyautogui.moveRel(0, 100)
            pyautogui.click()
            self.catiaOpen()
        else:
            if messagebox.askretrycancel("Failed", 'Either Catia/3D is blocked or not open.\n Place this program\'s window near top center.'):
                # will go to find search
                searchy = pyautogui.locateCenterOnScreen('images/searchy.png')
                if searchy is not None:
                    pyautogui.moveTo(searchy[0], searchy[1])
                    pyautogui.click()
                    self.catiaOpen()
                else:
                    pyautogui.moveTo(x, y/8)
                    pyautogui.moveRel(50,0)
                    pyautogui.click()
                    self.catiaOpen()
            else:
                new_notice_button.destroy()
                new_notice_label.destroy()
                new_notice_message.destroy()
                self.notice()

    def catiaOpen(self):
        time.sleep(2)
        catia = pyautogui.locateCenterOnScreen('images/catia.png')
        if catia is not None:
            # check if catia document is opened.
            thumb = pyautogui.locateCenterOnScreen('images/thumb.png')
            thumbnail = pyautogui.locateCenterOnScreen('images/thumbnail.png')
            if thumb or thumbnail is not None:
                # will go to find search
                searchy = pyautogui.locateCenterOnScreen('images/searchy.png')
                if searchy is not None:
                    pyautogui.moveTo(searchy[0], searchy[1])
                    pyautogui.click()
                    pyautogui.moveRel(100, 0)
                    pyautogui.click()
                    pyautogui.hotkey('ctrl', 'a')
                    pyautogui.press('backspace')
                    # self.searchPartcode()
                else:
                    pyautogui.moveTo(catia[0], catia[1])
                    pyautogui.click()
                    pyautogui.moveRel(200, 0)
                    pyautogui.click()
                    pyautogui.hotkey('ctrl', 'f')
                    pyautogui.hotkey('ctrl', 'a')
                    pyautogui.press('backspace')
                    # self.searchPartcode()
            else:
                if messagebox.askretrycancel("Failed!", '3D is either blocked or not opened.\n Check it before retrying.'):
                    # will go to find search
                    searchy = pyautogui.locateCenterOnScreen(
                        'images/searchy.png')
                    if searchy is not None:
                        pyautogui.moveTo(searchy[0], searchy[1])
                        pyautogui.click()
                        self.catiaOpen()
                    else:
                        pyautogui.moveTo(catia[0], catia[1])
                        pyautogui.click()
                        pyautogui.moveRel(200, 0)
                        pyautogui.click()
                        self.catiaOpen()
                else:
                    new_notice_button.destroy()
                    new_notice_label.destroy()
                    new_notice_message.destroy()
                    self.notice()
        else:
            self.retryCatia()

    def cancelOperation(self):
        if messagebox.askokcancel("Cancel", "This operation will be cancelled."):
            self.parent.geometry("600x120")
            App(app)

    def toCatia(self):
        global new_notice_button
        global new_notice_label
        global new_notice_message
        notice.destroy()
        notice_button.destroy()
        self.parent.geometry("650x200")
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
        self.parent.update_idletasks()
        self.catiaOpen()

    def notice(self):
        global notice
        global notice_button
        upload_progress.destroy()
        self.parent.geometry("650x250")
        notice = Message(
            self.frame, width=600, font=("Helvetica", 13), text="IMPORTANT NOTICE.\n\n- Make sure the required 3D is open in Catia Composer.\n- Make sure no other apps block Catia Composer's screen.\n- Remember to place this window near the top of the screen.\n- Click 'OK' to continue.")
        notice.place(x=10, y=10)
        notice_button = Button(self.frame, text='Ok', font=(
            "Helvetica", 13), command=self.toCatia)
        notice_button.place(x=450, y=200)

    def destroy_widget(self):
        self.upload_label.destroy()
        self.upload_button.destroy()
        file_label = Label(self.frame, anchor='w', width=600, height=120,
                           text="Select an excel file from your computer.", font=("Helvetica", 13))
        file_label.pack(side='left', padx=10)
        self.excel_file(file_label)

    def excel_file(self, label):
        filepath = filedialog.askopenfilename(
            initialdir="/Desktop", title='Select File', filetypes=[("Excel files", ".xlsx .xls .csv")])
        if not filepath:
            label.destroy()
            App(app)
        else:
            label.destroy()
            filename = ntpath.basename(filepath)
            if filename.lower().endswith(('.xlsx', '.xls', '.csv')):
                global upload_progress
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
                    empty_label = Label(
                        self.frame, text="This file is empty.", fg='red', font=("Helvetica", 13))
                    empty_label.pack(side='left', padx=10)
                    response = messagebox.showerror(
                        'Error!', f"{filename} is empty.")
                    if response == 'ok':
                        empty_label.destroy()
                        different_file_label = Label(
                            self.frame, text="Try a different excel file.", font=("Helvetica", 13))
                        different_file_label.pack(side='left', padx=10)
                        self.excel_file(different_file_label)
                else:
                    row_progress.destroy()
                    upload_progress.config(
                        text="File uploaded successfully.", fg='green')
                    messagebox.showinfo(
                        'Success', 'FILE UPLOADED\n\nClick "OK" to continue.')
                    self.notice()

            else:
                not_excel_label = Label(
                    self.frame, text="Upload Excel files only.", fg='red', font=("Helvetica", 13))
                not_excel_label.pack(side='left', padx=10)
                not_excel = messagebox.showerror(
                    'Error!', f"{filename} is not an Excel file.")
                if not_excel == 'ok':
                    not_excel_label.destroy()
                    different_file_label = Label(
                        self.frame, text="Try a different excel file.", font=("Helvetica", 13))
                    different_file_label.pack(side='left', padx=10)
                    self.excel_file(different_file_label)

    def onClosing():
        if messagebox.askokcancel("Quit", "Do you really wish to quit?"):
            app.destroy()


def main():
    global app
    global x
    global y
    app = Tk()
    # Place window slightly above the center of screen on start up
    width_window = 600
    height_window = 120
    screen_width = app.winfo_screenwidth()  # width of the screen
    screen_height = app.winfo_screenheight()  # height of the screen
    x_center = (screen_width/2)-(width_window/2)
    y_center = (screen_height/2)-(height_window/2)
    x = x_center
    y = (y_center/8)
    app.resizable(False, False)
    app.wm_iconbitmap('images/intercept.ico')
    app.wm_title('INTERCEPT')
    app.geometry('%dx%d+%d+%d' % (width_window, height_window, x, y))
    application = App(app)
    app.wm_protocol("WM_DELETE_WINDOW", App.onClosing)
    app.mainloop()


if __name__ == "__main__":
    main()
