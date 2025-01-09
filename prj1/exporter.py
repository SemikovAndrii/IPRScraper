from typing import Dict, Any

import openpyxl


def save_to_excel(parsed_data: Dict[str, Any], filename: str = "registry_data.xlsx") -> None:
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Registry Data"

    headers = [
        "Registry Number", "Registration Date", "Expiration Date", "IPR Type",
        "Protection Document Number", "Protection Document Date", "Protection Document Expiration",
        "Protection Document Name", "Goods", "Owner", "Authorized Persons"
    ]
    sheet.append(headers)

    for value in parsed_data.values():
        authorized_persons = "; ".join([
            f"{person['Name']} ({person['Phone']}, {person['Email']})"
            for person in value["Persons"].get("Authorized Persons", [])
        ])
        row = [
            value.get("Registry Number"),
            value.get("Registration Date"),
            value.get("Expiration Date"),
            value.get("IPR Type"),
            value["Protection Document"].get("Number"),
            value["Protection Document"].get("Date"),
            value["Protection Document"].get("Expiration"),
            value["Protection Document"].get("Name"),
            "; ".join(value.get("Goods", [])),
            value["Persons"].get("Owner"),
            authorized_persons
        ]
        sheet.append(row)

    workbook.save(filename)
    print(f"Data saved to {filename}")
