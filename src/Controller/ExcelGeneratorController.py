from src.Model.DataModel import DataModel
from src.View.ExcelGeneratorView import ExcelGeneratorView

class ExcelGeneratorController:
    def __init__(self):
        self.model = DataModel("/Users/antoine/PycharmProjects/SheetToExcel/Ressource/credentials.json", "/Users/antoine/PycharmProjects/SheetToExcel/Ressource/EXEMPLE.xlsx")
        self.view = ExcelGeneratorView(self)
        self.model.authenticate()
        self.model.fetch_data("SUIVI COMMANDE X2")
        self.view.populate_weeks(list(self.model.data_by_week.keys()))

    def generate_excel_for_selected_week(self, week):
        self.model.generate_excel_for_week(week)

    def run(self):
        self.view.run()
