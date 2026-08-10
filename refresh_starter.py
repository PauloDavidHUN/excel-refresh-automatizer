import logging
import os
import config
from refresh_helper import refresh_excel

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(BASE_DIR, "refresh.log"), encoding='utf-8'),
        logging.StreamHandler()
    ]
)

if __name__ == "__main__":
    logging.info("=" * 50)
    logging.info("Excel frissítő folyamat indul...")
    
    for nev, utvonal in config.EXCEL_FAJLOK.items():
        logging.info(f"Feldolgozás alatt: {nev}")
        try:
            refresh_excel(utvonal)
        except Exception as e:
            logging.error(f"Kritikus hiba a(z) {nev} fájl frissítésekor: {e}")
            
    logging.info("Minden feladat befejeződött.")
    logging.info("=" * 50)