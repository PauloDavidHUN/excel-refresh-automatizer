import win32com.client
import time
import os
import logging

# Script saját mappájának meghatározása
BASE_DIR = os.path.dirname(os.path.abspath(__file__))


def wait_for_box_sync(file_path, timeout=120, check_interval=2, stable_checks=3):
    """
    Megvárja amíg a Box szinkronizáció befejeződik.
    
    Args:
        file_path: A fájl elérési útja
        timeout: Maximum várakozási idő másodpercben
        check_interval: Ellenőrzések között ennyi másodpercet vár
        stable_checks: Ennyiszer kell egyforma méretnek lennie = "kész"
    """
    logging.info(f"Box szinkron ellenorzese: {file_path}")
    
    start_time = time.time()   #  VALÓDI idő mérés
    last_size = -1
    stable_count = 0           #  Hányszor volt már egyforma a méret
    
    while True:
        elapsed = time.time() - start_time  # ← Pontos eltelt idő
        
        # Timeout ellenőrzés
        if elapsed >= timeout:
            logging.error(f"Timeout! Szinkron nem fejeződött be {timeout}s alatt.")
            return False
        
        # Fájl létezik-e?
        if not os.path.exists(file_path):
            logging.debug(f"Fajl meg nem letezik... ({elapsed:.0f}s)")
            stable_count = 0  # Reset - ha eltűnt, kezdjük elölről
            time.sleep(check_interval)
            continue
        
        current_size = os.path.getsize(file_path)
        
        # Méret változott?
        if current_size != last_size:
            logging.info(f"Meg szinkronizal... ({elapsed:.0f}s) | Meret: {current_size/1024:.1f} KB")
            last_size = current_size
            stable_count = 0  # Reset - változott, nem kész
        else:
            # Méret ugyanaz mint előző körben
            if current_size > 0:
                stable_count += 1
                logging.debug(f"Stabil meret: {stable_count}/{stable_checks}")
                
                if stable_count >= stable_checks:
                    logging.info(f"Fajl kesz! Meret: {current_size/1024:.1f} KB | Ido: {elapsed:.0f}s")
                    return True
        
        time.sleep(check_interval)



#----------------------------------------------------------------------------------------------------------------
def _var_kapcsolatok_kesz(wb, timeout=120, check_interval=2):
    """
    Megvárja amíg az összes háttérlekérdezés befejeződik.
    Az aláhúzás jelzi hogy ez 'belső' függvény, csak ebben a fájlban használjuk.
    """
    logging.info("Varakozas a hatterlekerdezesekre...")
    start_time = time.time()

    while True:
        elapsed = time.time() - start_time
        if elapsed >= timeout:
            logging.warning(f"Kapcsolat timeout {timeout}s utan.")
            break

        # Ellenőrzi van-e még futó háttérlekérdezés
        van_futó = False
        for conn in wb.Connections:
            try:
                if conn.OLEDBConnection.Refreshing:
                    van_futó = True
                    break
            except AttributeError:
                pass  # Nem OLEDB kapcsolat, kihagyjuk

        if not van_futó:
            logging.info(f"Minden kapcsolat kesz. ({elapsed:.0f}s)")
            break

        logging.debug(f"Meg fut hatterlekerdezés... ({elapsed:.0f}s)")
        time.sleep(check_interval)
#----------------------------------------------------------------------------------------------------------------



def refresh_excel(file_path):
    logging.info(f"refresh_excel indul: {file_path}")
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AskToUpdateLinks = False
    wb = None  # ← fontos: hogy a finally tudja hogy megnyílt-e

    try:
        # Box szinkron ellenőrzés
        if not wait_for_box_sync(file_path):
            logging.error("Box szinkron sikertelen, leallas.")
            return

        # Excel megnyitás
        logging.info(f"Megnyitas: {file_path}")
        wb = excel.Workbooks.Open(file_path, UpdateLinks=0)

        # Kapcsolatok frissítése
        logging.info(f"Kapcsolatok szama: {wb.Connections.Count}")
        for conn in wb.Connections:
            try:
                logging.info(f"  Frissites: {conn.Name}")
                try:
                    conn.OLEDBConnection.BackgroundQuery = False
                except AttributeError:
                    logging.debug(f"  BackgroundQuery nem ertheto el: {conn.Name}")
                conn.Refresh()
                logging.info(f"  Kesz: {conn.Name}")
            except Exception as e:
                logging.error(f"  Kapcsolat hiba ({conn.Name}): {e}", exc_info=True)

        # Várakozás hogy a háttérlekérdezések befejezzenek
        _var_kapcsolatok_kesz(wb)

        # Pivot táblák frissítése
        pivot_count = 0
        pivot_hiba = 0
        for sheet in wb.Sheets:
            for pivot in sheet.PivotTables():
                try:
                    pivot.RefreshTable()
                    pivot_count += 1
                except Exception as e:
                    pivot_hiba += 1
                    logging.warning(f"  Pivot hiba ({sheet.Name} / {pivot.Name}): {e}")

        logging.info(f"Pivotok frissitve: {pivot_count} db, hiba: {pivot_hiba} db")

        wb.Save()
        wb.Close()
        wb = None  # ← jelzi hogy már be van zárva
        logging.info("Kesz! Mentve es bezarva.")

    except Exception as e:
        logging.error(f"Váratlan hiba: {e}", exc_info=True)

    finally:
        # Ha hiba volt és a wb még nyitva van, zárjuk be
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
                logging.warning("Workbook hiba utan bezarva (nem mentve).")
            except:
                pass
        excel.Quit()
        logging.info("Excel leallitva.")