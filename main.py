#!/usr/bin/env python3
"""
Stellux - Virtual Printer Forwarder
Monitort het File Port bestand en stuurt data door naar de fysieke printer.
"""

import time
import os
import subprocess
import tempfile

# Import Windows Print API
try:
    import win32print
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False
    print("⚠️  WARNING: pywin32 niet geïnstalleerd. Installeer met: pip install pywin32")
    print("   Fallback naar PRINT commando (minder betrouwbaar)\n")

# --- Configuratie ---
CAPTURED_FILE_PATH = r"C:\Stellux\Virtuele_Printer_File\Captured_Receipt.txt"

# Map waarin de BONNEN met tijdstempel worden opgeslagen (voor archivering)
OUTPUT_ARCHIVE_DIR = r"C:\Stellux\Virtuele_Printer_File\Archived_Receipts"

# De NAAM van je fysieke bonprinter (NIET de poort!)
# Zie in "Printers en Scanners" - gebruik de exacte naam
REAL_PRINTER_NAME = "EPSON TM-T(180dpi) Receipt6"

# --------------------

def forward_to_printer_via_win32(data, printer_name):
    """Stuurt data naar printer via Windows Print API (meest betrouwbaar voor ESC/POS)."""
    if not HAS_WIN32:
        return False
    
    try:
        # Open printer handle
        hPrinter = win32print.OpenPrinter(printer_name)
        
        try:
            # Start print job
            hJob = win32print.StartDocPrinter(hPrinter, 1, ("Stellux Receipt", None, "RAW"))
            
            try:
                # Start page
                win32print.StartPagePrinter(hPrinter)
                
                # Write raw data (ESC/POS codes worden direct verzonden!)
                win32print.WritePrinter(hPrinter, data)
                
                # End page
                win32print.EndPagePrinter(hPrinter)
                
                print(f"✅ Bon succesvol doorgestuurd naar printer '{printer_name}' (via Win32 API)")
                return True
                
            finally:
                # End document
                win32print.EndDocPrinter(hPrinter)
                
        finally:
            # Close printer
            win32print.ClosePrinter(hPrinter)
            
    except Exception as e:
        print(f"❌ Win32 Print API fout: {e}")
        return False

def forward_to_printer_via_copy(data, printer_name):
    """Stuurt data naar printer via Windows COPY commando (meest betrouwbaar)."""
    try:
        # Maak tijdelijk bestand voor de data
        with tempfile.NamedTemporaryFile(mode='wb', delete=False, suffix='.prn') as temp_file:
            temp_file.write(data)
            temp_path = temp_file.name
        
        # Gebruik Windows COPY commando om naar printer te schrijven
        # /B = binary mode (belangrijk voor ESC/POS!)
        cmd = f'copy /B "{temp_path}" "\\\\localhost\\{printer_name}"'
        
        result = subprocess.run(
            cmd,
            shell=True,
            capture_output=True,
            text=True
        )
        
        # Verwijder tijdelijk bestand
        try:
            os.unlink(temp_path)
        except:
            pass
        
        if result.returncode == 0:
            print(f"✅ Bon succesvol doorgestuurd naar printer '{printer_name}'")
            return True
        else:
            print(f"❌ COPY commando fout: {result.stderr}")
            return False
            
    except Exception as e:
        print(f"❌ Fout bij doorsturen naar printer: {e}")
        return False

def forward_to_printer_via_lpr(data, printer_name):
    """Alternatieve methode via Windows LPR (als COPY niet werkt)."""
    try:
        # Maak tijdelijk bestand
        with tempfile.NamedTemporaryFile(mode='wb', delete=False, suffix='.prn') as temp_file:
            temp_file.write(data)
            temp_path = temp_file.name
        
        # Gebruik LPR commando (oudere Windows methode, maar vaak betrouwbaar)
        cmd = f'print /D:"{printer_name}" "{temp_path}"'
        
        result = subprocess.run(
            cmd,
            shell=True,
            capture_output=True,
            text=True
        )
        
        # Kleine wacht tijd voor Windows om te verwerken
        time.sleep(0.5)
        
        # Verwijder tijdelijk bestand
        try:
            os.unlink(temp_path)
        except:
            pass
        
        if result.returncode == 0:
            print(f"✅ Bon succesvol doorgestuurd naar printer '{printer_name}' (via PRINT)")
            return True
        else:
            print(f"⚠️  PRINT commando resultaat: {result.stderr if result.stderr else 'Mogelijk succesvol'}")
            return True  # PRINT command geeft soms foutief een error terug
            
    except Exception as e:
        print(f"❌ Fout bij doorsturen naar printer (LPR): {e}")
        return False

def forward_receipt_to_printer(file_path, printer_name, archive_dir):
    """Leest de gevangen bon, slaat deze op, en stuurt deze door naar de fysieke printer."""
    try:
        # 1. Lees de data in binaire modus ('rb') om ESC/POS-codes te behouden
        with open(file_path, 'rb') as source:
            receipt_data = source.read()

        # Controleer of er daadwerkelijk inhoud is
        if not receipt_data:
            print("Bon gedetecteerd, maar bestand is leeg. Overslaan.")
            return

        print(f"\n📄 Bon gedetecteerd (Grootte: {len(receipt_data)} bytes)")

        # 2. Archivering: Sla de ruwe data op
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        output_filename = f"receipt_{timestamp}.txt"
        full_output_path = os.path.join(archive_dir, output_filename)
        
        with open(full_output_path, 'wb') as out_f:
            out_f.write(receipt_data)
        
        print(f"💾 Bon opgeslagen in archief: {output_filename}")

        # 3. Probeer eerst via Win32 API (BESTE voor ESC/POS!)
        print(f"🖨️  Doorsturen naar printer '{printer_name}'...")
        
        success = False
        
        # Methode 1: Win32 Print API (beste voor ESC/POS)
        if HAS_WIN32:
            success = forward_to_printer_via_win32(receipt_data, printer_name)
        
        # Methode 2: Als Win32 niet beschikbaar of faalt, probeer COPY
        if not success:
            if HAS_WIN32:
                print("⚠️  Win32 methode faalt, probeer COPY commando...")
            success = forward_to_printer_via_copy(receipt_data, printer_name)
        
        # Methode 3: Als COPY niet werkt, probeer PRINT commando
        if not success:
            print("⚠️  COPY methode faalt, probeer PRINT commando...")
            success = forward_to_printer_via_lpr(receipt_data, printer_name)
        
        if not success:
            print("❌ Beide methoden gefaald. Mogelijke oorzaken:")
            print("   1. Printer naam is niet correct (check 'Printers en Scanners')")
            print("   2. Printer is offline of niet gedeeld")
            print("   3. Driver ondersteunt geen directe data input")

        # 4. Maak het bronbestand leeg (CRUCIAAL)
        with open(file_path, 'wb') as f:
            f.truncate(0)
        
        print("✨ Bronbestand geleegd, klaar voor de volgende bon.\n")

    except FileNotFoundError:
        print(f"❌ Fout: Bestand {file_path} niet gevonden.")
    except PermissionError:
        print("❌ Fout: Geen schrijfrechten. Draai het script als Administrator.")
    except Exception as e:
        print(f"❌ Onverwachte fout bij printen: {e}")

def monitor_file_changes(file_path, printer_name, archive_dir, interval=1):
    """Monitor het bestand op wijzigingen en verwerk de bon bij detectie."""
    
    os.makedirs(os.path.dirname(file_path), exist_ok=True)
    os.makedirs(archive_dir, exist_ok=True)
    
    if not os.path.exists(file_path):
        open(file_path, 'w').close()

    print("="*80)
    print("🌟 Stellux - Virtual Printer Forwarder")
    print("="*80)
    print(f"\n📂 Monitoring: {file_path}")
    print(f"🖨️  Doelprinter: {printer_name}")
    print(f"💾 Archief: {archive_dir}")
    print(f"\n⏳ Wachten op bonnen... (Ctrl+C om te stoppen)\n")
    
    last_file_size = os.path.getsize(file_path)

    while True:
        try:
            current_file_size = os.path.getsize(file_path)

            if current_file_size > last_file_size and current_file_size > 0:
                forward_receipt_to_printer(file_path, printer_name, archive_dir)
                last_file_size = 0  # Na het legen is de grootte 0
            else:
                last_file_size = current_file_size
            
            time.sleep(interval)

        except KeyboardInterrupt:
            print("\n\n⏹️  Monitoring gestopt door gebruiker.")
            print("="*80)
            break
        except Exception as e:
            print(f"❌ Fout tijdens monitoring: {e}. Wachten...")
            time.sleep(interval * 5)

if __name__ == "__main__":
    # Check of printer naam correct is ingesteld
    if REAL_PRINTER_NAME == "EPSON TM-T(180dpi) Receipt6":
        print("\n⚠️  BELANGRIJK: Controleer of de printer naam exact overeenkomt!")
        print("   Ga naar 'Printers en Scanners' en kopieer de exacte naam.\n")
    
    monitor_file_changes(CAPTURED_FILE_PATH, REAL_PRINTER_NAME, OUTPUT_ARCHIVE_DIR)
