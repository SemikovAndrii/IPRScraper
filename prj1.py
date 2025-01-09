import logging
import pprint
import re
import time
from datetime import datetime, timedelta

import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler("script_log.log", mode='w', encoding='utf-8'),
        logging.StreamHandler()
    ]
)

logging.info("Запуск скрипта")

# Настройка WebDriver
driver = webdriver.Chrome()  # Убедитесь, что путь к ChromeDriver указан в PATH
driver.get("https://cabinet.customs.gov.ua/ipr/reg/overview")
logging.info("Открыта страница: https://cabinet.customs.gov.ua/ipr/reg/overview")

# Словарь для сохранения данных
data_dict = {}

try:
    # Ожидание загрузки страницы
    logging.info("Ожидание загрузки страницы...")
    wait = WebDriverWait(driver, 20)
    wait.until(
        EC.presence_of_element_located((By.XPATH, "//label[contains(normalize-space(string(.)), 'Зареєстрований з')]")))
    logging.info("Страница загружена.")

    # Пауза перед вводом данных
    logging.info("Пауза перед вводом данных...")
    time.sleep(2)

    # Вычисление дат
    yesterday = (datetime.now() - timedelta(days=100)).strftime("%d.%m.%Y")
    today = datetime.now().strftime("%d.%m.%Y")
    logging.info(f"Вычисленные даты: c {yesterday} по {today}")

    # Поле "Зареєстрований з"
    logging.info("Ввод даты 'Зареєстрований з'...")
    registered_from_label = driver.find_element(By.XPATH,
                                                "//label[contains(normalize-space(string(.)), 'Зареєстрований з')]")
    registered_from_input = registered_from_label.find_element(By.XPATH, "..//input")
    registered_from_input.clear()
    registered_from_input.send_keys(yesterday)

    # Поле "Зареєстрований по"
    logging.info("Ввод даты 'Зареєстрований по'...")
    registered_to_label = driver.find_element(By.XPATH,
                                              "//label[contains(normalize-space(string(.)), 'Зареєстрований по')]")
    registered_to_input = registered_to_label.find_element(By.XPATH, "..//input")
    registered_to_input.clear()
    registered_to_input.send_keys(today)

    # Пауза перед нажатием кнопки поиска
    logging.info("Пауза перед нажатием кнопки поиска...")
    time.sleep(2)

    # Нажатие кнопки "Знайти"
    logging.info("Нажатие кнопки 'Знайти'...")
    search_button = driver.find_element(By.XPATH, "//span[contains(text(), 'Знайти')]/ancestor::button")
    search_button.click()

    # Ожидание загрузки таблицы или сообщения
    logging.info("Ожидание загрузки таблицы или сообщения...")
    time.sleep(3)  # Дополнительная пауза для динамической загрузки контента

    try:
        # Проверка на наличие сообщения "Нічого не знайдено"
        no_data_message = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//span[contains(@class,'rz-growl-title') and contains(text(),'Нічого не знайдено')]")))
        logging.warning(f"Результат поиска: {no_data_message.text}")
        logging.info("Корректное завершение работы. Данных не найдено.")
    except:
        # Извлечение данных из таблицы
        logging.info("Начало извлечения данных из таблицы...")
        table_body = wait.until(
            EC.presence_of_element_located((By.XPATH, "//tbody[contains(@class,'rz-datatable-data')]")))
        rows = table_body.find_elements(By.TAG_NAME, "tr")
        logging.info(f"Найдено строк в таблице: {len(rows)}")

        # Обработка каждой строки
        for i, row in enumerate(rows, start=1):  # Нумерация строк начинается с 1
            try:
                # Поиск и клик по номеру объекта
                cell = row.find_element(By.XPATH,
                                        ".//td[contains(@class, 'text-center')]//a[contains(@class, 'rz-link')]")
                object_number = cell.text
                logging.info(f"Обработка объекта № {object_number}")
                cell.click()

                # Ожидание появления блока с деталями
                modal = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "customdialog-modal")))
                logging.info(f"Детали объекта № {object_number} извлечены.")

                # Сохранение данных
                data_dict[object_number] = modal.text

                # Клик по кнопке "Закрити"
                close_button = modal.find_element(By.XPATH,
                                                  ".//span[contains(@class, 'rz-button-text') and text()='Закрити']")
                close_button.click()
                logging.info(f"Блок с деталями объекта № {object_number} закрыт.")

                # Пауза для стабильности
                time.sleep(1)
            except Exception as e:
                logging.error(f"Ошибка при обработке строки {i}: {e}", exc_info=True)

except Exception as e:
    logging.error(f"Произошла ошибка: {e}", exc_info=True)

finally:
    logging.info("Закрытие браузера.")
    driver.quit()
    logging.info("Скрипт завершён.")

    # Вывод собранных данных
    logging.info("Собранные данные:")
    print(data_dict)


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

    # Persons
    persons = {}
    persons["Owner"] = re.search(r'Правовласник\n(.+?)\n', entry).group(1)

    # Authorized persons
    authorized_section = re.search(r'Уповноважена особа\n(.+)', entry, re.S)
    if authorized_section:
        authorized_people_raw = authorized_section.group(1).split(';')
        authorized_people = []
        for person in authorized_people_raw:
            match = re.search(r'(.+?)\((.+?), (.+?)\)', person.strip())
            if match:
                authorized_people.append({"Name": match.group(1).strip(),
                    "Phone": match.group(2).strip(),
                    "Email": match.group(3).strip()})
        persons["Authorized Persons"] = authorized_people

    parsed_data["Persons"] = persons

    return parsed_data


# Парсинг собранных данных
parsed_data = {key: parse_entry(value) for key, value in data_dict.items()}
pprint.pprint(parsed_data)

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
