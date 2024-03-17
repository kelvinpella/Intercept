import xlrd
import time
import ntpath
import pyautogui
import pyperclip
from tkinter import *
from tkinter import font
from tkinter import ttk, filedialog, messagebox


class App:

    def __init__(self, parent):
        self.setup_ui(parent)

    def setup_ui(self, parent):
        self.frame = Frame(parent)
        self.frame.place(relwidth=1, relheight=1)
        self.upload_label = Label(self.frame, text="Upload Excel file.", font=("Helvetica", 13))
        self.upload_label.place(x=10, y=10)
        self.upload_button = Button(self.frame, text="Upload", font=("Helvetica", 13), command=self.select_excel_file)
        self.upload_button.place(width=180, height=50, x=10, y=60)

    def select_excel_file(self):
        filepath = filedialog.askopenfilename(initialdir="/Desktop", title='Select File',
                                              filetypes=[("Excel files", ".xlsx .xls .csv")])
        if filepath:
            self.process_excel_file(filepath)
        else:
            messagebox.showerror("Error", "No file selected.")

    def process_excel_file(self, filepath):
        try:
            workbook = xlrd.open_workbook(filepath)
            sheet = workbook.sheet_by_index(0)
            all_data = self.extract_data(sheet)
            if all_data:
                self.run_analysis(all_data)
            else:
                messagebox.showerror("Error", "The selected file is empty.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process the file: {e}")

    def extract_data(self, sheet):
        all_data = []
        for row_idx in range(sheet.nrows):
            row_data = [sheet.cell_value(row_idx, col_idx) for col_idx in (1, 3)]
            all_data.append(row_data)
        return all_data

    def run_analysis(self, all_data):
        # Implement the logic for analyzing the extracted data
        pass

    # Additional methods for analysis and error handling...


def main():
    root = Tk()
    root.title("INTERCEPT")
    root.geometry("650x500")
    app = App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
