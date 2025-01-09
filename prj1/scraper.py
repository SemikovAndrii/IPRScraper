
import logging
import time
from datetime import datetime, timedelta
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def scrape_data():
    logging.info("Запуск браузера...")
    driver = webdriver.Chrome()
    driver.get("https://cabinet.customs.gov.ua/ipr/reg/overview")
    logging.info("Открыта страница: https://cabinet.customs.gov.ua/ipr/reg/overview")

    raw_data = {}
    try:
        wait = WebDriverWait(driver, 20)
        logging.info("Ожидание загрузки страницы...")

        # Проверяем наличие поля "Зареєстрований з"
        registered_from_label = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//label[contains(normalize-space(string(.)), 'Зареєстрований з')]")
        ))
        logging.info("Элемент 'Зареєстрований з' найден.")

        # Пауза перед вводом данных
        logging.info("Пауза перед вводом данных...")
        time.sleep(2)

        # Вычисление дат
        yesterday = (datetime.now() - timedelta(days=100)).strftime("%d.%m.%Y")
        today = datetime.now().strftime("%d.%m.%Y")
        logging.info(f"Ввод дат: с {yesterday} по {today}")

        # Поля для ввода дат
        registered_from_input = registered_from_label.find_element(By.XPATH, "..//input")
        registered_to_label = driver.find_element(By.XPATH,
            "//label[contains(normalize-space(string(.)), 'Зареєстрований по')]"
        )
        registered_to_input = registered_to_label.find_element(By.XPATH, "..//input")

        registered_from_input.clear()
        registered_from_input.send_keys(yesterday)
        registered_to_input.clear()
        registered_to_input.send_keys(today)

        # Нажатие кнопки поиска
        time.sleep(2)
        logging.info("Нажатие кнопки 'Знайти'...")
        search_button = driver.find_element(By.XPATH, "//span[contains(text(), 'Знайти')]/ancestor::button")
        search_button.click()

        # Ожидание загрузки таблицы
        logging.info("Ожидание загрузки таблицы...")
        table_body = wait.until(
            EC.presence_of_element_located((By.XPATH, "//tbody[contains(@class,'rz-datatable-data')]"))
        )
        rows = table_body.find_elements(By.TAG_NAME, "tr")
        logging.info(f"Найдено строк в таблице: {len(rows)}")

        # Обработка каждой строки
        for i, row in enumerate(rows, start=1):
            try:
                cell = row.find_element(By.XPATH, ".//td[contains(@class, 'text-center')]//a[contains(@class, 'rz-link')]")
                object_number = cell.text
                logging.info(f"Обработка объекта № {object_number}")
                cell.click()

                modal = wait.until(EC.presence_of_element_located((By.CLASS_NAME, "customdialog-modal")))
                raw_data[object_number] = modal.text

                close_button = modal.find_element(By.XPATH, ".//span[contains(text(), 'Закрити')]")
                close_button.click()
                time.sleep(1)
            except Exception as e:
                logging.error(f"Ошибка при обработке строки {i}: {e}", exc_info=True)
    except Exception as e:
        logging.error(f"Ошибка во время выполнения: {e}", exc_info=True)
    finally:
        driver.quit()
        logging.info("Браузер закрыт.")
    return raw_data
