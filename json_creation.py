import json
from openpyxl import load_workbook

wb = load_workbook(filename = "pellet-dock.xlsx")
sheet = wb.active

pellet_vessels = {}

for row in sheet.iter_rows(min_row=2,
                           min_col=1,
                           values_only=True):
    vessel_name = row[0]
    vessel = {
        "Vessel Name": row[3],
        "Vessel Number": row[0],
        "BL Date": row[1],
        "Discharge Date": row[2],
        "Pellet Grade": row[4],
        "BL Quantity": row[5],
        "Destination Draft Survey": row[6],
        "Deviation From Original Weight": row[7],
        "Deviation Code ": row[8],
        "Pellet PO#": row[9],
        "Shipment PO#": row[10],
        "Purchase Price": row[11]
    }

    pellet_vessels[vessel_name] = vessel

json_object = json.dumps(pellet_vessels, indent = 4, sort_keys=True, default=str)

with open("sample.json", "w") as outfile:
        outfile.write(json_object)
