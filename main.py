import os
import openpyxl
from docxtpl import DocxTemplate
from datetime import date
import json

# Load data from Excel
path = "data.xlsx"
output_path = "output"
workbook = openpyxl.load_workbook(path)
sheet = workbook["below"]
input_file = open("input.json")
input_data = json.load(input_file)

list_values = list(sheet.values)

# Generate docs
doc = DocxTemplate("letter.docx")
today = date.today().strftime("%d/%m/%Y")

if not os.path.exists(output_path):
    os.makedirs(output_path)

for value_tuple in list_values[1:]:
    if(value_tuple[1]):
        doc.render({"rollno": value_tuple[0],
                "percentage": value_tuple[1],
                "name": value_tuple[2],
                "date" : today,
                "from_date": input_data["from_date"],
                "to_date": input_data["to_date"],
                "year": input_data["year"],
                "department": input_data["department"],
                "section": input_data["section"]})
    
    doc_name = "letter" + value_tuple[0] + ".docx"
    doc.save(output_path + "/" + doc_name)
print("Letters Generated Successfully")