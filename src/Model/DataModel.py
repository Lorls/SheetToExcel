import gspread
from google.oauth2.service_account import Credentials
from collections import defaultdict
import re
import json
import openpyxl
import shutil
import os

class DataModel:
    def __init__(self, credentials_path, template_path):
        self.credentials_path = credentials_path
        self.template_path = template_path
        self.excel_dir = "/Users/antoine/PycharmProjects/SheetToExcel/Excel"
        self.data_by_week = defaultdict(list)
        self.data_by_week_and_urgent = defaultdict(lambda: {'urgent': [], 'non_urgent': []})
        self.client = None

        os.makedirs(self.excel_dir, exist_ok=True)

    def authenticate(self):
        scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = Credentials.from_service_account_file(self.credentials_path, scopes=scopes)
        self.client = gspread.authorize(creds)

    def fetch_data(self, spreadsheet_name):
        spreadsheet = self.client.open(spreadsheet_name)
        worksheets = spreadsheet.worksheets()

        for worksheet in worksheets:
            title = worksheet.title
            if re.match(r'^\d{4}-\d{4}$', title):
                print(f"Checking table: {title}")
                headers = worksheet.row_values(1)

                headers = [header.strip() for header in headers]

                if 'FINI (AUTOMATIQUE)' in headers:
                    print(f"Processing table: {title}")
                    self.process_worksheet(worksheet)
                else:
                    print(f"Skipping table: {title} because 'FINI (AUTOMATIQUE)' column is missing.")

    def process_worksheet(self, worksheet):
        data = worksheet.get_all_values()
        headers = data[0]
        rows = data[1:]

        headers = [header.strip() for header in headers]

        if 'FINI (AUTOMATIQUE)' not in headers:
            print(f"Skipping sheet '{worksheet.title}' because 'FINI (AUTOMATIQUE)' column is missing.")
            return

        fini_column_index = headers.index('FINI (AUTOMATIQUE)')

        filtered_data = []
        for row in rows:
            if row[0] and row[1]:
                entry = {headers[i]: row[i] for i in range(len(headers))}

                column_fini_value = row[fini_column_index]
                if column_fini_value != "OUI":
                    filtered_data.append(entry)

        for entry in filtered_data:
            week = entry.get('SEMAINE PROD')
            atelier_column = entry.get('ATELIER / SPORTSWEAR / SOCKS / BIDONS / TONNELLE')

            if atelier_column and 'ATELIER' in atelier_column:
                if week:
                    self.data_by_week[week].append(entry)
                    if entry.get('COMMANDE URGENTE') == 'OUI':
                        self.data_by_week_and_urgent[week]['urgent'].append(entry)
                    else:
                        self.data_by_week_and_urgent[week]['non_urgent'].append(entry)

    def delete_old_files(self):
        for file in os.listdir(self.data_dir):
            if file.startswith('output_week_'):
                os.remove(os.path.join(self.data_dir, file))
                print(f"Deleted file: {file}")

        for file in os.listdir(self.excel_dir):
            if file.startswith('ORDERS PLANNING WEEKS '):
                os.remove(os.path.join(self.excel_dir, file))
                print(f"Deleted file: {file}")

    def save_json_files(self):
        for week, entries in self.data_by_week.items():
            json_output_week = json.dumps(entries, indent=4, ensure_ascii=False)
            filename = os.path.join(self.data_dir, f'output_week_{week}.json')
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(json_output_week)
            print(f"Created file: {filename}")

        for week, urgency_data in self.data_by_week_and_urgent.items():
            json_output_week_urgent = json.dumps(urgency_data['urgent'], indent=4, ensure_ascii=False)
            urgent_filename = os.path.join(self.data_dir, f'output_week_{week}_urgent.json')
            with open(urgent_filename, 'w', encoding='utf-8') as f:
                f.write(json_output_week_urgent)
            print(f"Created file: {urgent_filename}")

            json_output_week_non_urgent = json.dumps(urgency_data['non_urgent'], indent=4, ensure_ascii=False)
            non_urgent_filename = os.path.join(self.data_dir, f'output_week_{week}_non_urgent.json')
            with open(non_urgent_filename, 'w', encoding='utf-8') as f:
                f.write(json_output_week_non_urgent)
            print(f"Created file: {non_urgent_filename}")

    def generate_excel_for_week(self, week):
        output_path = os.path.join(self.excel_dir, f"ORDERS PLANNING WEEKS {week} 2024.xlsx")
        shutil.copy(self.template_path, output_path)

        workbook = openpyxl.load_workbook(output_path)
        sheet = workbook.active

        sheet["D1"] = f"{week}"
        sheet["D2"] = f"{int(week) + 2}"

        non_urgent_start_row = 4
        non_urgent_count = len(self.data_by_week_and_urgent[week]['non_urgent'])
        last_non_urgent_row = non_urgent_start_row + non_urgent_count - 1

        urgent_label_row = last_non_urgent_row + 2
        urgent_start_row = urgent_label_row + 1

        red_fill = openpyxl.styles.PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

        week_data = self.data_by_week_and_urgent[week]

        for i, entry in enumerate(week_data['non_urgent']):
            client_number = entry.get('NUMERO DOSSIER', '')
            order_name = entry.get('CLUB', '')
            quantity = entry.get('QUANTITE', '')

            sheet[f"B{non_urgent_start_row + i}"] = client_number
            sheet[f"C{non_urgent_start_row + i}"] = order_name
            sheet[f"D{non_urgent_start_row + i}"] = quantity

        sheet[f"C{urgent_label_row}"] = "URGENT"
        sheet[f"B{urgent_label_row}"] = ""  # Empty
        sheet[f"D{urgent_label_row}"] = ""  # Empty

        sheet[f"B{urgent_label_row}"].fill = red_fill
        sheet[f"C{urgent_label_row}"].fill = red_fill
        sheet[f"D{urgent_label_row}"].fill = red_fill

        for i, entry in enumerate(week_data['urgent']):
            client_number = entry.get('NUMERO DOSSIER', '')
            order_name = entry.get('CLUB', '')
            quantity = entry.get('QUANTITE', '')

            sheet[f"B{urgent_start_row + i}"] = client_number
            sheet[f"C{urgent_start_row + i}"] = order_name
            sheet[f"D{urgent_start_row + i}"] = quantity

        workbook.save(output_path)
        print(f"Created Excel file: {output_path}")

    def generate_all_excels(self):
        for week in self.data_by_week_and_urgent:
            if week.isnumeric():
                self.generate_excel_for_week(week)
