
import logging
from scraper import scrape_data
from parser import parse_all
from exporter import save_to_excel

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

def main():
    logging.info("Запуск процесса...")
    raw_data = scrape_data()
    parsed_data = parse_all(raw_data)
    save_to_excel(parsed_data)
    logging.info("Процесс завершён.")

if __name__ == "__main__":
    main()
