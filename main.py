"""
Copia foto da iPhone (MTP) → Windows 11
Richiede: pip install pywin32
"""

import os
import json
import datetime
import time
from typing import List, Optional, Tuple
from dataclasses import dataclass
import pythoncom
import win32com.client as win32
import win32clipboard


@dataclass
class CopyResult:
    """Risultato dell'operazione di copia"""
    completate: List[str]
    saltate: List[str]
    errori: List[str]
    tempo_totale: float = 0.0
    file_copiati: int = 0
    velocita_media: float = 0.0


class IPhoneMTPCopier:
    """Gestisce la copia di foto da iPhone tramite MTP"""
    
    def __init__(self, dest_dir: Optional[str] = None):
        self.desktop = os.path.join(os.environ['USERPROFILE'], 'Desktop')
        self.dest_dir = dest_dir or self._get_destination_folder()
        self.log_file = os.path.join(self.dest_dir, "log.txt")
        self.shell = None
        # Statistiche per stima tempo
        self.start_time = None
        self.files_copied_total = 0
        self.copy_times = []  # Lista tempi di copia per calcolare media
    
    def _get_destination_folder(self) -> str:
        """Legge la cartella di destinazione dal file di configurazione"""
        config_path = os.path.join(os.path.dirname(__file__), 'config.json')
        
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
                
            if config.get('use_desktop', True):
                return os.path.join(self.desktop, 'photo_iphone')
            else:
                custom_path = config.get('custom_path', '')
                if custom_path and os.path.exists(os.path.dirname(custom_path)):
                    return custom_path
                else:
                    print(f"⚠️ Invalid custom path: {custom_path}")
                    print("   Using Desktop as fallback.")
                    return os.path.join(self.desktop, 'photo_iphone')
                    
        except FileNotFoundError:
            # Se config.json non esiste, crealo con valori predefiniti
            default_config = {
                "use_desktop": True,
                "custom_path": "C:/Users/nome_utente/Pictures/iPhone"
            }
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, indent=2)
            return os.path.join(self.desktop, 'photo_iphone')
            
        except Exception as e:
            print(f"⚠️ Error reading config: {e}")
            return os.path.join(self.desktop, 'photo_iphone')
        
    def __enter__(self):
        """Inizializza COM e crea directory di destinazione"""
        os.makedirs(self.dest_dir, exist_ok=True)
        pythoncom.CoInitialize()
        self.shell = win32.Dispatch("Shell.Application")
        return self
        
    def __exit__(self, exc_type, exc_val, exc_tb):
        """Pulisce risorse e deinizializza COM"""
        self._clear_clipboard()
        pythoncom.CoUninitialize()

    def log(self, msg: str) -> None:
        """Registra messaggi su console e file"""
        ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{ts}] {msg}"
        print(line)
        with open(self.log_file, "a", encoding="utf-8") as f:
            f.write(line + "\n")


    def _find_iphone(self) -> Optional[object]:
        """Trova il dispositivo iPhone connesso"""
        computer = self.shell.Namespace(17)  # CSIDL_DRIVES
        
        for item in computer.Items():
            name_lower = item.Name.lower()
            if "iphone" in name_lower or "apple" in name_lower:
                return item
        return None

    def _find_folder(self, parent: object, target_name: str) -> Optional[object]:
        """Cerca una sotto-cartella (case-insensitive)"""
        target_lower = target_name.lower()
        
        for sub in parent.GetFolder.Items():
            if sub.IsFolder and sub.Name.lower() == target_lower:
                return sub
        return None

    @staticmethod
    def _clear_clipboard() -> None:
        """Pulisce il clipboard di Windows"""
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            win32clipboard.CloseClipboard()
        except:
            pass

    def _copy_file_mtp(self, file_item: object, dest_path: str) -> bool:
        """Copia un singolo file da MTP"""
        file_start = time.time()
        try:
            dest_folder = self.shell.Namespace(os.path.dirname(dest_path))
            if dest_folder:
                dest_folder.CopyHere(file_item, 16)  # 16 = Respond "Yes to All"
                time.sleep(0.03)
                
                if os.path.exists(dest_path):
                    copy_time = time.time() - file_start
                    self.copy_times.append(copy_time)
                    self.files_copied_total += 1
                    return True
            
            # Fallback: metodo clipboard
            self._clear_clipboard()
            file_item.InvokeVerbEx("copy")
            dest_folder.Self.InvokeVerbEx("paste")
            
            time.sleep(0.1)
            if os.path.exists(dest_path):
                copy_time = time.time() - file_start
                self.copy_times.append(copy_time)
                self.files_copied_total += 1
                return True
            return False
            
        except Exception as e:
            self.log(f"❌ Error during copy: {e}")
            return False

    def _process_roll(self, roll: object, roll_idx: int, total_rolls: int, source_files: List) -> str:
        """Processa una singola roll/cartella"""
        dst_path = os.path.join(self.dest_dir, roll.Name)
        source_count = len(source_files)
        
        # Crea la cartella se non esiste
        if not os.path.exists(dst_path):
            os.makedirs(dst_path)
            self.log(f"📁 [{roll_idx}/{total_rolls}] New folder: {roll.Name} ({source_count} files to copy)")
        else:
            self.log(f"📁 [{roll_idx}/{total_rolls}] Checking folder: {roll.Name} ({source_count} files)")
        
        # Scansiona e copia i file mancanti
        files_copied = 0
        files_skipped = 0
        files_failed = 0
        
        for idx, file_item in enumerate(source_files, 1):
            dest_file = os.path.join(dst_path, file_item.Name)
            
            if os.path.exists(dest_file):
                files_skipped += 1
                self.log(f"  [{idx}/{source_count}] ⏭️  File already exists: {roll.Name}/{file_item.Name}")
                continue
            
            if self._copy_file_mtp(file_item, dest_file):
                files_copied += 1
                self.log(f"  [{idx}/{source_count}] ✅ Copied: {roll.Name}/{file_item.Name}")
            else:
                files_failed += 1
                self.log(f"  [{idx}/{source_count}] ❌ Failed: {roll.Name}/{file_item.Name}")
        
        # Determina lo stato finale
        if files_failed > 0:
            self.log(f"⚠️  [{roll_idx}/{total_rolls}] Completed with errors: {roll.Name} ({files_copied} copied, {files_skipped} skipped, {files_failed} failed)")
            return "errore"
        elif files_copied == 0 and files_skipped == source_count:
            self.log(f"✅ [{roll_idx}/{total_rolls}] Already complete: {roll.Name} ({files_skipped} files)")
            return "saltata"
        else:
            self.log(f"✅ [{roll_idx}/{total_rolls}] Completed: {roll.Name} ({files_copied} copied, {files_skipped} skipped)")
            return "completata"
    
    def _format_time(self, seconds: float) -> str:
        """Formatta secondi in formato leggibile"""
        if seconds < 60:
            return f"{int(seconds)} seconds"
        elif seconds < 3600:
            minutes = int(seconds // 60)
            secs = int(seconds % 60)
            return f"{minutes} minutes and {secs} seconds"
        else:
            hours = int(seconds // 3600)
            minutes = int((seconds % 3600) // 60)
            return f"{hours} hours and {minutes} minutes"
    
    def _calculate_eta(self, rolls_done: int, total_rolls: int, files_per_roll_avg: float) -> str:
        """Calcola tempo stimato rimanente"""
        if not self.copy_times or rolls_done == 0:
            return "Calculating..."
        
        # Velocità media (secondi per file)
        avg_time_per_file = sum(self.copy_times) / len(self.copy_times)
        
        # Stima file rimanenti
        rolls_remaining = total_rolls - rolls_done
        estimated_files_remaining = rolls_remaining * files_per_roll_avg
        
        # Tempo stimato con fattore empirico per scanning cartelle
        estimated_seconds = estimated_files_remaining * avg_time_per_file
        estimated_seconds *= 1.2 # add 20% for folder scanning overhead
        
        # Add "approx" to indicate it's an estimate
        return f"{self._format_time(estimated_seconds)} (approx)"
    
    def copy_photos(self) -> CopyResult:
        """Esegue la copia delle foto dall'iPhone"""
        self.start_time = time.time()
        
        # Trova iPhone
        iphone = self._find_iphone()
        if not iphone:
            self.log("❌ iPhone not found.")
            return CopyResult([], [], [])
        
        # Trova Internal Storage
        internal = self._find_folder(iphone, "Internal Storage")
        if not internal:
            self.log("❌ Internal Storage not found.")
            return CopyResult([], [], [])
        
        # Ottieni tutte le rolls ordinate
        rolls = sorted(
            (item for item in internal.GetFolder.Items() if item.IsFolder),
            key=lambda x: x.Name
        )
        
        total_rolls = len(rolls)
        self.log(f"📲 Starting copy of {total_rolls} folders…")
        result = CopyResult([], [], [])
        
        # Per calcolare media file per cartella
        total_files_in_rolls = 0
        rolls_with_files = 0
        
        for idx, roll in enumerate(rolls, 1):
            try:
                # Stampa prima quale cartella sta per analizzare
                print(f"\n📂 [{idx}/{total_rolls}] Analyzing folder: {roll.Name}")
                
                # Ottieni i file della cartella
                roll_files = [f for f in roll.GetFolder.Items() if not f.IsFolder]
                if roll_files:
                    total_files_in_rolls += len(roll_files)
                    rolls_with_files += 1
                
                status = self._process_roll(roll, idx, total_rolls, roll_files)
                
                if status == "completata":
                    result.completate.append(roll.Name)
                elif status == "saltata":
                    result.saltate.append(roll.Name)
                elif status == "errore":
                    result.errori.append(roll.Name)
                
                # Mostra ETA dopo ogni cartella processata (solo se ci sono stati file copiati)
                if idx < total_rolls and self.files_copied_total > 0 and status != "saltata":
                    avg_files_per_roll = total_files_in_rolls / rolls_with_files if rolls_with_files > 0 else 0
                    avg_files_per_roll *= 1.2 # add 20% because recent folders are statistically larger
                    
                    eta = self._calculate_eta(idx, total_rolls, avg_files_per_roll)
                    self.log(f"⏱️  Estimated time remaining: {eta} [statistical estimate, not based on actual missing files]\n")
                    
            except Exception as e:
                self.log(f"❌ Critical error {roll.Name}: {e}")
                result.errori.append(roll.Name)
        
        # Calcola statistiche finali
        tempo_totale = time.time() - self.start_time
        result.tempo_totale = tempo_totale
        result.file_copiati = self.files_copied_total
        
        if self.files_copied_total > 0:
            result.velocita_media = self.files_copied_total / tempo_totale * 60  # files/minute
        
        self.log(f"\n🎉 FINAL SUMMARY ({total_rolls} total folders):\n"
                f"   ✅ Completed: {len(result.completate)}\n"
                f"   ⏭️  Skipped: {len(result.saltate)}\n"
                f"   ❌ With errors: {len(result.errori)}\n"
                f"\n📊 STATISTICS:\n"
                f"   ⏱️  Total time: {self._format_time(tempo_totale)}\n"
                f"   📁 Files copied: {self.files_copied_total}\n"
                f"   ⚡ Average speed: {result.velocita_media:.1f} files/minute")
        
        return result


if __name__ == "__main__":
    """Entry point principale"""
    with IPhoneMTPCopier() as copier:
        copier.copy_photos()
