from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import time, os, shutil

SPOOL_DIR = r"C:\Windows\System32\spool\PRINTERS"
DEST_DIR = r"C:\Stellux\raw_bonnen"

class SpoolWatcher(FileSystemEventHandler):
    def on_created(self, event):
        if event.is_directory:
            return

        if event.src_path.endswith(('.SPL', '.SHD')):
            basename = os.path.basename(event.src_path)

            # tijdstempel zonder : (Windows verboden teken)
            ts = time.strftime("%Y-%m-%d_%H-%M-%S")
            dest = os.path.join(DEST_DIR, f"{ts}_{basename}")

            # probeer meerdere keren te kopiëren (bestanden zijn soms nog in gebruik)
            for attempt in range(5):
                try:
                    if os.path.exists(event.src_path):
                        shutil.copy2(event.src_path, dest)
                        print(f"📄 Nieuwe spoolfile gekopieerd naar: {dest}")
                        break
                    else:
                        print(f"[WACHT] Bestand bestaat nog niet: {event.src_path}")
                except Exception as e:
                    print(f"[FOUT] Poging {attempt+1}: kon bestand niet kopiëren: {event.src_path} -> {e}")
                time.sleep(0.3)  # wacht even en probeer opnieuw

if __name__ == "__main__":
    os.makedirs(DEST_DIR, exist_ok=True)
    observer = Observer()
    event_handler = SpoolWatcher()
    observer.schedule(event_handler, SPOOL_DIR, recursive=False)
    observer.start()
    print(f"[INFO] Monitor gestart op {SPOOL_DIR}")

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
