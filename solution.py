import json
import csv
import os
from openpyxl import Workbook

def main():
    
    try:
        with open('config.json', 'r', encoding='utf-8') as f:
            config = json.load(f)
    except FileNotFoundError:
        print("config.json не найден")
        config = {}
    except json.JSONDecodeError:
        print("config.json содержит некорректный JSON")
        config = {}
    
    workbook = Workbook()
    
    if "Sheet" in workbook.sheetnames:
        del workbook["Sheet"]
    
    
    for sheet_name, csv_path in config.items():
        ws = workbook.create_sheet(title=sheet_name)
        
        try:
            with open(csv_path, 'r', encoding='utf-8') as file:
                reader = csv.reader(file)
                for row_idx, row in enumerate(reader, start=1):
                   
                    if row and any(cell.strip() for cell in row):
                        for col_idx, value in enumerate(row, start=1):
                            try:
                                if '.' in value:
                                    value = float(value)
                                else:
                                    value = int(value)
                            except (ValueError, TypeError):
                                pass
                            ws.cell(row=row_idx, column=col_idx, value=value)
                    
                    
        except FileNotFoundError:
            print(f"Файл {csv_path} не найден")
        except Exception as e:
            print(f"Ошибка при чтении {csv_path}: {e}")
    
    workbook.save('annual_report.xlsx')
    print(f"annual_report.xlsx сохранен с {len(workbook.sheetnames)} листами")

main()