#!/usr/bin/env python3
"""
Stellux - Windows Print Monitor
Monitort de Windows print spooler en kopieert automatisch alle nieuwe bonnen.

INSTALLATIE (eenmalig):
    pip install watchdog

GEBRUIK:
    python windows_monitor.py

VEREISTEN:
    - Run als Administrator
    - Printer moet printen naar Windows spooler
"""

from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import time
import os
import shutil

# Configuratie
SPOOL_DIR = r"C:\Windows\System32\spool\PRINTERS"
DEST_DIR = r"C:\Stellux\raw_bonnen"

class SpoolWatcher(FileSystemEventHandler):
    """Monitort de print spooler en kopieert nieuwe bestanden."""
    
    def on_created(self, event):
        if event.is_directory:
            return

        if event.src_path.endswith(('.SPL', '.SHD')):
            basename = os.path.basename(event.src_path)
            
            # Tijdstempel zonder : (Windows verboden teken)
            ts = time.strftime("%Y-%m-%d_%H-%M-%S")
            dest = os.path.join(DEST_DIR, f"{ts}_{basename}")
            
            # Probeer meerdere keren te kopiëren (bestanden zijn soms nog in gebruik)
            for attempt in range(5):
                try:
                    if os.path.exists(event.src_path):
                        shutil.copy2(event.src_path, dest)
                        
                        # Check grootte
                        size = os.path.getsize(dest)
                        
                        # Toon met kleurcodes (werkt in moderne Windows Terminal)
                        if basename.endswith('.SPL'):
                            if size > 0:
                                print(f"✅ BON DATA: {basename} → {dest} ({size:,} bytes)")
                            else:
                                print(f"⚠️  LEEG: {basename} → {dest} (0 bytes - te laat gekopieerd?)")
                        else:
                            print(f"📋 METADATA: {basename} → {dest} ({size:,} bytes)")
                        
                        break
                    else:
                        print(f"[WACHT] Bestand bestaat nog niet: {event.src_path}")
                except Exception as e:
                    print(f"[FOUT] Poging {attempt+1}: kon bestand niet kopiëren: {basename} -> {e}")
                time.sleep(0.3)  # Wacht even en probeer opnieuw

def main():
    """Start de print monitor."""
    
    print("\n" + "="*80)
    print("🌟 Stellux - Windows Print Monitor")
    print("="*80)
    
    # Check of we in de spooler directory kunnen kijken
    if not os.path.exists(SPOOL_DIR):
        print(f"\n❌ Spooler directory niet gevonden: {SPOOL_DIR}")
        print("\n💡 Controleer of je op Windows bent")
        return
    
    try:
        # Test toegang
        os.listdir(SPOOL_DIR)
    except PermissionError:
        print(f"\n❌ Geen toegang tot spooler directory!")
        print(f"\n💡 Run dit script als Administrator:")
        print(f"   1. Open cmd.exe als Administrator")
        print(f"   2. cd naar deze directory")
        print(f"   3. Run: python windows_monitor.py")
        return
    
    # Maak backup directory aan
    os.makedirs(DEST_DIR, exist_ok=True)
    print(f"\n✓ Monitoring: {SPOOL_DIR}")
    print(f"✓ Backup naar: {DEST_DIR}")
    
    # Start observer
    observer = Observer()
    event_handler = SpoolWatcher()
    observer.schedule(event_handler, SPOOL_DIR, recursive=False)
    observer.start()
    
    print("\n" + "="*80)
    print("🖨️  KLAAR! Print nu een bon - deze wordt automatisch gekopieerd")
    print("="*80)
    print("\n(Druk Ctrl+C om te stoppen)\n")
    
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n\n⏹️  Monitor gestopt")
        observer.stop()
    
    observer.join()
    print("\n✅ Klaar! Kopieer de bestanden nu naar je Mac:\n")
    print(f"   Van: {DEST_DIR}")
    print(f"   Naar: ~/Github/Stellux/raw_bonnen/\n")

if __name__ == "__main__":
    main()

