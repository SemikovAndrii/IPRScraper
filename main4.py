import re
import openpyxl

# Функция для парсинга данных объекта с использованием регулярных выражений
def parse_entry(entry):
    parsed_data = {}

    # Раздел "General Information"
    parsed_data["Registry Number"] = re.search(r'Номер у реєстрі\n(\d+)', entry).group(1)
    parsed_data["Registration Date"] = re.search(r'Дата реєстрації\n([\d.]+)', entry).group(1)
    parsed_data["Expiration Date"] = re.search(r'Термін реєстрації\n([\d.]+)', entry).group(1)
    parsed_data["IPR Type"] = re.search(r'Тип ОПІВ\n(.+?)\n', entry).group(1)

    # Раздел "Protection Document"
    protection_doc = {}
    protection_doc["Number"] = re.search(r'Охоронний документ\nНомер\n(\d+)', entry).group(1)
    protection_doc["Date"] = re.search(r'Охоронний документ.*?Дата\n([\d.]+)', entry, re.S).group(1)
    protection_doc["Expiration"] = re.search(r'Закінчення строку\n([\d.]+)', entry).group(1)
    protection_doc["Name"] = re.search(r'Назва\n(.+?)\n', entry).group(1)
    parsed_data["Protection Document"] = protection_doc

    # Раздел "Goods"
    goods_match = re.search(r'Товари\n(.+?)\nОсоби', entry, re.S)
    if goods_match:
        parsed_data["Goods"] = [item.strip() for item in goods_match.group(1).split(';')]

    # Раздел "Persons"
    persons = {}
    persons["Owner"] = re.search(r'Правовласник\n(.+?)\n', entry).group(1)
    authorized_match = re.search(r'Уповноважена особа\n(.+?)\((.+?), (.+?)\)', entry)
    if authorized_match:
        persons["Authorized Person"] = {
            "Name": authorized_match.group(1),
            "Phone": authorized_match.group(2),
            "Email": authorized_match.group(3),
        }
    parsed_data["Persons"] = persons

    return parsed_data

# Словарь данных
data = {
    '1056': 'ОПІВ\nclose\nЗагальна інформація\nНомер у реєстрі\n1056\nДата реєстрації\n01.10.2024\nТермін реєстрації\n01.10.2025\nТип ОПІВ\nТорговельна марка, зареєстрована в Україні (НТМ)\nОхоронний документ\nНомер\n274960\nДата\n27.04.2020\nЗакінчення строку\n31.05.2028\nНазва\nСвідоцтво України на торговельну марку "VizEll VE"\nТовари\nпряжа бавовняна (крім швейних ниток), розфасована для роздрібної торгівлі з вмістом бавовни ; пряжа бавовняна (крім швейних ниток), розфасована для роздрібної торгівлі інші; пряжа з вовни або тонкого волосу тварин, розфасована для роздрібної торгівлі інша; пряжа із штучних штапельних волокон (крім швейних ниток), розфасована для роздрібної торгівлі\nОсоби\nПравовласник\nВізна Олена Миколаївна\nУповноважена особа\nЩербакова Ольга (380971048999, scherbakova.olga@gmail.com)\nЗакрити',
    # Добавьте другие записи
}

# Обработка всех записей в словаре
parsed_data = {key: parse_entry(value) for key, value in data.items()}

# Сохранение в Excel
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = "Registry Data"

# Заголовки
headers = [
    "Registry Number", "Registration Date", "Expiration Date", "IPR Type",
    "Protection Document Number", "Protection Document Date", "Protection Document Expiration",
    "Protection Document Name", "Goods", "Owner", "Authorized Person Name", "Authorized Person Phone",
    "Authorized Person Email"
]
sheet.append(headers)

# Заполнение данных
for key, value in parsed_data.items():
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
        value["Persons"].get("Authorized Person", {}).get("Name"),
        value["Persons"].get("Authorized Person", {}).get("Phone"),
        value["Persons"].get("Authorized Person", {}).get("Email"),
    ]
    sheet.append(row)

# Сохранение файла
workbook.save("registry_data.xlsx")
print("Data saved to registry_data.xlsx")

