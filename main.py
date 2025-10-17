import time
import os
import shutil

# --- Configuratie ---
CAPTURED_FILE_PATH = r"C:\Stellux\Virtuele_Printer_File_Port\Captured_Receipt.txt"

# Map waarin de BONNEN met tijdstempel worden opgeslagen (voor archivering)
OUTPUT_ARCHIVE_DIR = r"C:\Stellux\Virtuele_Printer_File_Port\Archived_Receipts"

# De ECHTE poort van je fysieke bonprinter (USB000 uit image_2f48be.jpg)
REAL_PRINTER_PORT = r"\\.\USB000" 
# --------------------

def forward_receipt_to_printer(file_path, printer_port, archive_dir):
    """Leest de gevangen bon, slaat deze op, en stuurt deze door naar de fysieke poort."""
    try:
        # 1. Lees de data in binaire modus ('rb') om ESC/POS-codes te behouden
        with open(file_path, 'rb') as source:
            receipt_data = source.read()

        # Controleer of er daadwerkelijk inhoud is
        if not receipt_data:
            print("Bon gedetecteerd, maar bestand is leeg. Overslaan.")
            return

        print(f"Bon gedetecteerd (Grootte: {len(receipt_data)} bytes).")

        # 2. Archivering (Optioneel maar Aangeraden): Sla de ruwe data op
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        output_filename = f"receipt_{timestamp}.txt"
        full_output_path = os.path.join(archive_dir, output_filename)
        
        with open(full_output_path, 'wb') as out_f: # Opslaan ook in binaire modus
            out_f.write(receipt_data)
        
        print(f"✅ Bon Data Succesvol Opgeslagen (Archief): {full_output_path}")

        # 3. Schrijf de ruwe binaire data naar de fysieke printerpoort
        with open(printer_port, 'wb') as target:
            target.write(receipt_data)
        
        print(f"✅ Bon succesvol doorgestuurd naar {printer_port}. Printen...")

        # 4. Maak het bronbestand leeg (CRUCIAAL)
        with open(file_path, 'wb') as f:
            f.truncate(0)
        
        print("Bronbestand geleegd, klaar voor de volgende bon.")

    except FileNotFoundError:
        print(f"Fout: Bestanden of poort {printer_port} niet gevonden.")
    except PermissionError:
        print("Fout: Geen schrijfrechten. Draai het script als Administrator.")
    except Exception as e:
        print(f"Onverwachte fout bij printen: {e}")

def monitor_file_changes(file_path, printer_port, archive_dir, interval=1):
    """Monitor het bestand op wijzigingen en verwerk de bon bij detectie."""
    
    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    os.makedirs(archive_dir, exist_ok=True)
    
    if not os.path.exists(file_path):
        open(file_path, 'w').close()

    print(f"Start monitoring van {file_path} en doorsturen naar {printer_port}...")
    
    last_file_size = os.path.getsize(file_path)

    while True:
        try:
            current_file_size = os.path.getsize(file_path)

            if current_file_size > last_file_size and current_file_size > 0:
                print("\n--- BON DETECTEERD ---")
                forward_receipt_to_printer(file_path, printer_port, archive_dir)
                last_file_size = 0 # Na het legen is de grootte 0
            else:
                last_file_size = current_file_size
            
            time.sleep(interval)

        except KeyboardInterrupt:
            print("\nMonitoring gestopt door gebruiker.")
            break
        except Exception as e:
            print(f"Fout tijdens monitoring: {e}. Wachten...")
            time.sleep(interval * 5)

if __name__ == "__main__":
    monitor_file_changes(CAPTURED_FILE_PATH, REAL_PRINTER_PORT, OUTPUT_ARCHIVE_DIR)
