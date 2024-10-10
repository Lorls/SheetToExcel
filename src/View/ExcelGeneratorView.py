import tkinter as tk
from tkinter import ttk

class ExcelGeneratorView:
    def __init__(self, controller):
        self.controller = controller

        self.root = tk.Tk()
        self.root.title("Excel Generator")
        self.root.geometry("300x200")

        self.week_var = tk.StringVar(self.root)
        self.week_label = ttk.Label(self.root, text="Select Waeek")
        self.week_label.pack()

        self.week_menu = ttk.Combobox(self.root, textvariable=self.week_var)
        self.week_menu.pack()

        self.generate_button = ttk.Button(self.root, text="Generate Excel", command=self.generate_excel)
        self.generate_button.pack()

        self.status_label = ttk.Label(self.root, text="")
        self.status_label.pack()

    def populate_weeks(self, weeks):
        """ Populate the dropdown menu with week options """
        self.week_menu['values'] = weeks

    def generate_excel(self):
        selected_week = self.week_var.get()
        if selected_week:
            self.controller.generate_excel_for_selected_week(selected_week)
            self.status_label.config(text=f"Generated Excel for Week {selected_week}")
        else:
            self.status_label.config(text="Please select a week")

    def run(self):
        self.root.mainloop()
