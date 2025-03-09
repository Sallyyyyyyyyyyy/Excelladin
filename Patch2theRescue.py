#!/usr/bin/env python
"""
Excelladin Reloaded - Tab Hernoem Patch
Deze patch hernoemt de tabbladen in de GUI.
"""
import os
import sys
import shutil
import re
import datetime
import traceback

def pas_patch_toe():
    """
    Hoofdfunctie van de patch die de tabbladen hernoemt
    """
    # Log initialiseren
    log_entries = []
    log_entries.append(f"=== Excelladin Reloaded Tab Hernoem Patch ===")
    log_entries.append(f"Uitgevoerd op: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    log_entries.append("=" * 50)
    
    try:
        # Pad naar gui.py
        gui_pad = os.path.join('modules', 'gui.py')
        
        # Controleer of het bestand bestaat
        if not os.path.exists(gui_pad):
            log_entries.append(f"FOUT: Bestand {gui_pad} niet gevonden!")
            schrijf_log(log_entries)
            return False
        
        # Maak backup van het bestand
        backup_pad = maak_backup(gui_pad)
        if not backup_pad:
            log_entries.append("FOUT: Kon geen backup maken van gui.py")
            schrijf_log(log_entries)
            return False
        
        log_entries.append(f"Backup gemaakt: {backup_pad}")
        
        # Lees de inhoud van het bestand
        with open(gui_pad, 'r', encoding='utf-8') as bestand:
            inhoud = bestand.read()
        
        # Voer de wijzigingen uit
        nieuwe_inhoud, aantal_wijzigingen = hernoem_tabs(inhoud)
        
        if aantal_wijzigingen == 0:
            log_entries.append("WAARSCHUWING: Geen wijzigingen aangebracht, tabbladen mogelijk al hernoemd")
            schrijf_log(log_entries)
            return False
        
        # Schrijf de nieuwe inhoud terug naar het bestand
        with open(gui_pad, 'w', encoding='utf-8') as bestand:
            bestand.write(nieuwe_inhoud)
        
        log_entries.append(f"Succes: {aantal_wijzigingen} tabbladen hernoemd in {gui_pad}")
        log_entries.append("De volgende tab-labels zijn gewijzigd:")
        log_entries.append("- 'ProductSheet aanmaken' -> '1 ProductSheet Aanmaken'")
        log_entries.append("- 'Sheet Kiezen' -> '2 Sheet Kiezen'")
        log_entries.append("- 'Acties' -> '3 Acties'")
        
        # Probeer patch logging naar logger module te sturen
        try:
            from modules.logger import logger
            logger.logPatch("Tab hernoem patch succesvol uitgevoerd")
        except Exception as e:
            log_entries.append(f"Info: Kon niet loggen naar logger module: {str(e)}")
        
        schrijf_log(log_entries)
        return True
        
    except Exception as e:
        log_entries.append(f"FOUT: Onverwachte fout tijdens uitvoeren patch: {str(e)}")
        log_entries.append(traceback.format_exc())
        schrijf_log(log_entries)
        return False

def maak_backup(bestandspad):
    """
    Maakt een backup van het opgegeven bestand
    
    Args:
        bestandspad (str): Pad naar het bestand
        
    Returns:
        str: Pad naar de backup of None bij fout
    """
    try:
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        bestandsnaam = os.path.basename(bestandspad)
        base_naam, extensie = os.path.splitext(bestandsnaam)
        backup_naam = f"{base_naam}_backup_{timestamp}{extensie}"
        
        # Bepaal de backup directory (naast het originele bestand)
        backup_dir = os.path.dirname(bestandspad)
        backup_pad = os.path.join(backup_dir, backup_naam)
        
        # Kopieer het bestand
        shutil.copy2(bestandspad, backup_pad)
        return backup_pad
    except Exception as e:
        print(f"Fout bij maken backup: {e}")
        return None

def hernoem_tabs(inhoud):
    """
    Hernoemt de tabs in de GUI code
    
    Args:
        inhoud (str): De inhoud van gui.py
        
    Returns:
        tuple: (nieuwe_inhoud, aantal_wijzigingen)
    """
    wijzigingen = 0
    
    # Wijzig productsheet tab label - houd rekening met zowel originele als mogelijke eerdere wijzigingen
    pattern1_orig = r'self\.tabControl\.add\(self\.tabProductSheet, text="ProductSheet aanmaken"\)'
    pattern1_prev = r'self\.tabControl\.add\(self\.tabProductSheet, text="1 ProductSheet"\)'
    replacement1 = r'self.tabControl.add(self.tabProductSheet, text="1 ProductSheet Aanmaken")'
    
    nieuwe_inhoud, aantal = re.subn(pattern1_orig, replacement1, inhoud)
    wijzigingen += aantal
    
    if aantal == 0:  # Als originele patroon niet gevonden, probeer mogelijk eerder gewijzigd patroon
        nieuwe_inhoud, aantal = re.subn(pattern1_prev, replacement1, inhoud if wijzigingen == 0 else nieuwe_inhoud)
        wijzigingen += aantal
    
    # Wijzig sheet kiezen tab label
    pattern2_orig = r'self\.tabControl\.add\(self\.tabSheetKiezen, text="Sheet Kiezen"\)'
    pattern2_prev = r'self\.tabControl\.add\(self\.tabSheetKiezen, text="2 Sheet Kiezen"\)'
    replacement2 = r'self.tabControl.add(self.tabSheetKiezen, text="2 Sheet Kiezen")'
    
    nieuwe_inhoud, aantal = re.subn(pattern2_orig, replacement2, nieuwe_inhoud)
    wijzigingen += aantal
    
    if aantal == 0:  # Als het patroon al is aangepast, toch tellen als 'geen wijziging nodig'
        aantal = len(re.findall(pattern2_prev, nieuwe_inhoud))
        if aantal > 0:
            wijzigingen += 1  # Tel het als een 'match' maar maak geen daadwerkelijke wijziging
    
    # Wijzig acties tab label
    pattern3_orig = r'self\.tabControl\.add\(self\.tabActies, text="Acties"\)'
    pattern3_prev = r'self\.tabControl\.add\(self\.tabActies, text="3 Acties"\)'
    replacement3 = r'self.tabControl.add(self.tabActies, text="3 Acties")'
    
    nieuwe_inhoud, aantal = re.subn(pattern3_orig, replacement3, nieuwe_inhoud)
    wijzigingen += aantal
    
    if aantal == 0:  # Als het patroon al is aangepast, toch tellen als 'geen wijziging nodig'
        aantal = len(re.findall(pattern3_prev, nieuwe_inhoud))
        if aantal > 0:
            wijzigingen += 1  # Tel het als een 'match' maar maak geen daadwerkelijke wijziging
    
    return nieuwe_inhoud, wijzigingen

def schrijf_log(log_entries):
    """
    Schrijft de log naar de console en naar een bestand
    
    Args:
        log_entries (list): Lijst met log-regels
    """
    # Zorg dat logs directory bestaat
    if not os.path.exists('logs'):
        os.makedirs('logs')
    
    # Zorg dat patch_historie subdirectory bestaat
    patch_historie_dir = os.path.join('logs', 'patch_historie')
    if not os.path.exists(patch_historie_dir):
        os.makedirs(patch_historie_dir)
    
    # Maak logbestandsnaam
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    log_bestand = os.path.join(patch_historie_dir, f"patch_geschiedenis{timestamp}.txt")
    
    # Schrijf log naar bestand
    with open(log_bestand, 'w', encoding='utf-8') as bestand:
        for regel in log_entries:
            bestand.write(regel + "\n")
    
    # Toon log op console
    for regel in log_entries:
        print(regel)
    
    print("\nLog is ook opgeslagen in:", log_bestand)

if __name__ == "__main__":
    # Zorg ervoor dat we in de root directory van de applicatie zitten
    # Controleer of modules directory bestaat
    if not os.path.exists('modules'):
        # Probeer naar parent directory te gaan als we in een submap zitten
        script_dir = os.path.dirname(os.path.abspath(__file__))
        parent_dir = os.path.dirname(script_dir)
        
        if os.path.exists(os.path.join(parent_dir, 'modules')):
            os.chdir(parent_dir)
        else:
            # Zoek recursief naar de applicatie root
            for root, dirs, files in os.walk(os.path.dirname(os.path.abspath(os.getcwd()))):
                if 'modules' in dirs and 'main.py' in files:
                    os.chdir(root)
                    print(f"Applicatie root gevonden: {os.getcwd()}")
                    break
    
    # Voer de patch toe
    print("Patch wordt uitgevoerd...\n")
    succes = pas_patch_toe()
    
    if succes:
        print("\nPatch is succesvol uitgevoerd!")
    else:
        print("\nPatch is NIET succesvol uitgevoerd. Zie logbestand voor details.")
    
    input("\nDruk op Enter om af te sluiten...")