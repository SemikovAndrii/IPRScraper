import re
from typing import Dict, List, Any


def parse_entry(entry: str) -> Dict[str, Any]:
    parsed_data: Dict[str, Any] = {}

    parsed_data["Registry Number"] = re.search(r'Номер у реєстрі\n(\d+)', entry).group(1)
    parsed_data["Registration Date"] = re.search(r'Дата реєстрації\n([\d.]+)', entry).group(1)
    parsed_data["Expiration Date"] = re.search(r'Термін реєстрації\n([\d.]+)', entry).group(1)
    parsed_data["IPR Type"] = re.search(r'Тип ОПІВ\n(.+?)\n', entry).group(1)

    protection_doc: Dict[str, str] = {}
    protection_doc["Number"] = re.search(r'Охоронний документ\nНомер\n(\d+)', entry).group(1)
    protection_doc["Date"] = re.search(r'Охоронний документ.*?Дата\n([\d.]+)', entry, re.S).group(1)
    protection_doc["Expiration"] = re.search(r'Закінчення строку\n([\d.]+)', entry).group(1)
    protection_doc["Name"] = re.search(r'Назва\n(.+?)\n', entry).group(1)
    parsed_data["Protection Document"] = protection_doc

    goods_match = re.search(r'Товари\n(.+?)\nОсоби', entry, re.S)
    if goods_match:
        parsed_data["Goods"] = [item.strip() for item in goods_match.group(1).split(';')]

    persons: Dict[str, Any] = {}
    persons["Owner"] = re.search(r'Правовласник\n(.+?)\n', entry).group(1)

    authorized_section = re.search(r'Уповноважена особа\n(.+)', entry, re.S)
    if authorized_section:
        authorized_people_raw = authorized_section.group(1).split(';')
        authorized_people: List[Dict[str, str]] = []
        for person in authorized_people_raw:
            match = re.search(r'(.+?)\((.+?), (.+?)\)', person.strip())
            if match:
                authorized_people.append({
                    "Name": match.group(1).strip(),
                    "Phone": match.group(2).strip(),
                    "Email": match.group(3).strip(),
                })
        persons["Authorized Persons"] = authorized_people

    parsed_data["Persons"] = persons

    return parsed_data


def parse_all(raw_data: Dict[str, str]) -> Dict[str, Any]:
    return {key: parse_entry(value) for key, value in raw_data.items()}
