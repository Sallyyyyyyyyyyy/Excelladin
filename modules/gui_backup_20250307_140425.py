"""
GUI module voor Excelladin Reloaded
Verantwoordelijk voor alle visuele elementen en interacties
"""
import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import tkinter.font as tkFont

from modules.logger import logger
from modules.settings import instellingen
from modules.excel_handler import excelHandler
from modules.actions import BESCHIKBARE_ACTIES, voerActieUit
from modules.workflow import workflowManager

from assets.theme import KLEUREN, FONTS, STIJLEN, TOOLTIP_STIJL

class Tooltip:
    """
    Tooltip klasse voor het tonen van tooltips bij hover
    """
    def __init__(self, widget, tekst):
        """
        Initialiseer een tooltip
        
        Args:
            widget: Het widget waaraan de tooltip wordt gekoppeld
            tekst (str): De tekst die in de tooltip wordt getoond
        """
        self.widget = widget
        self.tekst = tekst
        self.tooltipWindow = None
        self.widget.bind("<Enter>", self.toon)
        self.widget.bind("<Leave>", self.verberg)
    
    def toon(self, event=None):
        """Toon de tooltip"""
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        
        # Maak tooltipvenster
        self.tooltipWindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        
        label = tk.Label(
            tw, text=self.tekst, justify='left',
            **TOOLTIP_STIJL
        )
        label.pack()
    
    def verberg(self, event=None):
        """Verberg de tooltip"""
        if self.tooltipWindow:
            self.tooltipWindow.destroy()
            self.tooltipWindow = None

class ExcelladinApp:
    """
    Hoofdklasse voor de Excelladin Reloaded applicatie
    """
    def __init__(self, root):
        """
        Initialiseer de applicatie
        
        Args:
            root: Tkinter root window
        """
        self.root = root
        self.root.title("Excelladin Reloaded")
        self.root.geometry("400x800")
        self.root.resizable(False, False)
        
        # Configureer stijlen
        self._configureerStijlen()
        
        # Bouw de GUI
        self._buildGUI()
        
        # Laad laatste bestand indien nodig
        laatsteBestand = instellingen.haalLaatsteBestand()
        if laatsteBestand and os.path.exists(laatsteBestand):
            self.laadExcelBestand(laatsteBestand)
    
    def bevestigAfsluiten(self):
        """Vraag om bevestiging voordat de applicatie wordt afgesloten"""
        if messagebox.askyesno("Afsluiten", "Weet je zeker dat je Excelladin Reloaded wilt afsluiten?"):
            self.root.destroy()
    
    def _configureerStijlen(self):
        """Configureer de stijlen voor Tkinter widgets"""
        # Zorg dat Papyrus beschikbaar is
        beschikbareFonts = tkFont.families()
        titelFont = FONTS["titel"][0] if FONTS["titel"][0] in beschikbareFonts else "Arial"
        
        # Pas lettertype aan indien nodig
        self.titelFont = (titelFont, FONTS["titel"][1], FONTS["titel"][2])
        
        # Configureer stijlen voor ttk widgets
        self.stijl = ttk.Style()
        
        # Maak een speciaal donker thema aan voor ttk widgets
        # Deze instellingen zorgen dat de knoppen daadwerkelijk de donkere kleur gebruiken
        self.stijl.theme_create("ExcelladinThema", parent="alt", 
            settings={
                "TButton": {
                    "configure": {
                        "background": "#0a0d2c",  # Donker marineblauw
                        "foreground": KLEUREN["tekst"],  # Felgeel
                        "font": FONTS["normaal"],
                        "relief": "raised",
                        "borderwidth": 1,
                        "padding": (10, 5)
                    },
                    "map": {
                        "background": [
                            ("active", KLEUREN["button_hover"]),
                            ("disabled", "#555555")
                        ],
                        "foreground": [
                            ("disabled", "#999999")
                        ]
                    }
                },
                "TCheckbutton": {
                    "configure": {
                        "background": KLEUREN["achtergrond"],
                        "foreground": KLEUREN["tekst"],
                        "font": FONTS["normaal"]
                    }
                },
                "TRadiobutton": {
                    "configure": {
                        "background": KLEUREN["achtergrond"],
                        "foreground": KLEUREN["tekst"],
                        "font": FONTS["normaal"]
                    }
                },
                "TNotebook": {
                    "configure": {
                        "background": KLEUREN["achtergrond"],
                        "tabmargins": [2, 5, 2, 0]
                    }
                },
                "TNotebook.Tab": {
                    "configure": {
                        "background": KLEUREN["tabblad_inactief"],
                        "foreground": KLEUREN["tekst"],
                        "font": FONTS["normaal"],
                        "padding": [10, 5]
                    },
                    "map": {
                        "background": [
                            ("selected", KLEUREN["tabblad_actief"])
                        ],
                        "foreground": [
                            ("selected", "#FFFF00")
                        ]
                    }
                }
            }
        )
        
        # Activeer het nieuwe thema
        self.stijl.theme_use("ExcelladinThema")
    
    def _buildGUI(self):
        """Bouw de complete GUI"""
        # Maak hoofdframe
        self.hoofdFrame = tk.Frame(
            self.root,
            background=KLEUREN["achtergrond"]
        )
        self.hoofdFrame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        self._maakHeader()
        
        # Tabbladen
        self._maakTabbladen()
        
        # Statusbalk
        self._maakStatusbalk()
    
    def _maakHeader(self):
        """Maak de header met titel"""
        self.headerFrame = tk.Frame(
            self.hoofdFrame,
            **STIJLEN["header"],
            height=50
        )
        self.headerFrame.pack(fill=tk.X)
        
        # Laad en toon afbeelding (links)
        try:
            from PIL import Image, ImageTk
            import os
            
            img_path = os.path.join('assets', 'alladin.jpg')
            if os.path.exists(img_path):
                # Open de afbeelding en pas de grootte aan
                original_img = Image.open(img_path)
                # Bereken nieuwe grootte met behoud van aspect ratio
                width, height = original_img.size
                new_height = 50
                new_width = int(width * (new_height / height))
                resized_img = original_img.resize((new_width, new_height), Image.LANCZOS)
                
                # Converteer naar PhotoImage voor Tkinter
                tk_img = ImageTk.PhotoImage(resized_img)
                
                # Maak label met afbeelding
                self.logoLabel = tk.Label(
                    self.headerFrame,
                    image=tk_img,
                    background=STIJLEN["header"]["background"]
                )
                self.logoLabel.image = tk_img  # Bewaar referentie
                self.logoLabel.pack(side=tk.LEFT, padx=5)
                
                logger.logInfo("Logo afbeelding succesvol geladen")
            else:
                logger.logWaarschuwing(f"Afbeelding niet gevonden: {img_path}")
                # Fallback naar placeholder als afbeelding niet bestaat
                self.logoPlaceholder = tk.Label(
                    self.headerFrame,
                    text="[Logo]",
                    foreground=STIJLEN["label"]["foreground"],
                    font=STIJLEN["label"]["font"],
                    background=STIJLEN["header"]["background"],
                    width=10,
                    height=2
                )
                self.logoPlaceholder.pack(side=tk.LEFT)
        except Exception as e:
            logger.logFout(f"Fout bij laden logo afbeelding: {e}")
            # Fallback naar placeholder bij fouten
            self.logoPlaceholder = tk.Label(
                self.headerFrame,
                text="[Logo]",
                foreground=STIJLEN["label"]["foreground"],
                font=STIJLEN["label"]["font"],
                background=STIJLEN["header"]["background"],
                width=10,
                height=2
            )
            self.logoPlaceholder.pack(side=tk.LEFT)
        
        # Titel (rechts)
        self.titelLabel = tk.Label(
            self.headerFrame,
            text="Excelladin Reloaded",
            **STIJLEN["titel_label"]
        )
        self.titelLabel.pack(side=tk.RIGHT, padx=10)
    
    def _maakTabbladen(self):
        """Maak de tabbladen"""
        self.tabControl = ttk.Notebook(self.hoofdFrame)
        self.tabControl.pack(fill=tk.BOTH, expand=True)
        
        # Tabblad 1: Sheet Kiezen
        # Tabblad 1: ProductSheet aanmaken
        self.tabProductSheet = tk.Frame(
            self.tabControl,
            background=KLEUREN["achtergrond"]
        )
        self.tabControl.add(self.tabProductSheet, text="1 ProductSheet")
        self._maakProductSheetTab()

        self.tabSheetKiezen = tk.Frame(
            self.tabControl,
            background=KLEUREN["achtergrond"]
        )
        self.tabControl.add(self.tabSheetKiezen, text="2 Sheet Kiezen")
        self._maakSheetKiezenTab()
        
        # Tabblad 2: Acties
        self.tabActies = tk.Frame(
            self.tabControl,
            background=KLEUREN["achtergrond"]
        )
        self.tabControl.add(self.tabActies, text="3 Acties")
        self._maakActiesTab()
    
    def _maakSheetKiezenTab(self):
        """Maak de inhoud van het Sheet Kiezen tabblad"""
        # Hoofdcontainer met padding
        container = tk.Frame(
            self.tabSheetKiezen,
            background=KLEUREN["achtergrond"],
            padx=20,
            pady=20
        )
        container.pack(fill=tk.BOTH, expand=True)
        
        # Label met instructie
        instructieLabel = tk.Label(
            container,
            text="Selecteer een Excel-bestand om te bewerken",
            **STIJLEN["label"],
            pady=10
        )
        instructieLabel.pack(fill=tk.X)
        
        # Bestandsselectie frame
        bestandsFrame = tk.Frame(
            container,
            background=KLEUREN["achtergrond"]
        )
        bestandsFrame.pack(fill=tk.X, pady=10)
        
        # Bestandspad entry
        self.bestandspadVar = tk.StringVar()
        self.bestandspadEntry = tk.Entry(
            bestandsFrame,
            textvariable=self.bestandspadVar,
            **STIJLEN["entry"],
            width=30
        )
        self.bestandspadEntry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Blader knop
                # Vervangen door tk.Button voor betere tekstweergave
        self.bladerButton = tk.Button(
            bestandsFrame,
            text="...",  # Korte tekst die altijd past
            command=self.kiesExcelBestand,
            width=8,  # Kleinere breedte nodig voor kortere tekst
            bg="#000080",  # Donkerblauw
            fg="#FFFF00",  # Fel geel
            font=("Arial", 10, "bold")
        )
        self.bladerButton.pack(side=tk.RIGHT, padx=(10, 0))
        Tooltip(self.bladerButton, "Klik om een Excel-bestand te selecteren")
        
        # Checkbox om bestand te onthouden (tk.Checkbutton in plaats van ttk voor betere werking)
        self.onthoudBestandVar = tk.BooleanVar(value=instellingen.haalOp('Algemeen', 'OnthoudBestand') == 'True')
                # Verbeterde checkbox implementatie
                # Verbeterde checkbox implementatie
        self.onthoudBestandCheck = tk.Checkbutton(
            container,
            text="Onthoud dit bestand voor volgende sessie",
            variable=self.onthoudBestandVar,
            command=self.onthoudBestandUpdate,
            background=KLEUREN["achtergrond"],
            activebackground=KLEUREN["achtergrond"],
            foreground=KLEUREN["tekst"],
            selectcolor="#444444"  # Donkere achtergrond voor betere zichtbaarheid van checkmark
        )
        self.onthoudBestandCheck.pack(fill=tk.X, pady=10)
        Tooltip(self.onthoudBestandCheck, "Als dit is aangevinkt, wordt het bestand automatisch geladen bij het opstarten")
        self.laadButton = ttk.Button(
            container,
            text="Laad Geselecteerd Bestand",
            command=self.laadGeselecteerdBestand
        )
        self.laadButton.pack(fill=tk.X, pady=10)
        Tooltip(self.laadButton, "Klik om het geselecteerde Excel-bestand te laden")

        
        # Laad knop - nu direct na de checkbox
                
        # Frame voor bestandsinformatie
        infoFrame = tk.Frame(
            container,
            background=KLEUREN["achtergrond"],
            relief=tk.GROOVE,
            borderwidth=1,
            padx=10,
            pady=10
        )
        infoFrame.pack(fill=tk.X, pady=10)
        
        # Bestandsinformatie labels
        self.bestandsInfoLabel = tk.Label(
            infoFrame,
            text="Geen bestand geladen",
            **STIJLEN["label"],
            anchor=tk.W
        )
        self.bestandsInfoLabel.pack(fill=tk.X)
        
        self.rijInfoLabel = tk.Label(
            infoFrame,
            text="Rijen: 0",
            **STIJLEN["label"],
            anchor=tk.W
        )
        self.rijInfoLabel.pack(fill=tk.X)
        
        self.kolomInfoLabel = tk.Label(
            infoFrame,
            text="Kolommen: 0",
            **STIJLEN["label"],
            anchor=tk.W
        )
        self.kolomInfoLabel.pack(fill=tk.X)
    
    def _maakActiesTab(self):
        """Maak de inhoud van het Acties tabblad"""
        # Hoofdcontainer met padding
        container = tk.Frame(
            self.tabActies,
            background=KLEUREN["achtergrond"],
            padx=20,
            pady=20
        )
        container.pack(fill=tk.BOTH, expand=True)
        
        # Label met instructie
        instructieLabel = tk.Label(
            container,
            text="Selecteer acties om uit te voeren op het Excel-bestand",
            **STIJLEN["label"],
            pady=10
        )
        instructieLabel.pack(fill=tk.X)
        
        # Bereik selectie frame
        bereikFrame = tk.Frame(
            container,
            background=KLEUREN["achtergrond"],
            padx=10,
            pady=10
        )
        bereikFrame.pack(fill=tk.X)
        
        bereikLabel = tk.Label(
            bereikFrame,
            text="Bereik:",
            **STIJLEN["label"]
        )
        bereikLabel.pack(side=tk.LEFT)
        
        # Bereik opties
        self.bereikVar = tk.StringVar(value="alles")
        
        allesRadio = ttk.Radiobutton(
            bereikFrame,
            text="Alles",
            variable=self.bereikVar,
            value="alles"
        )
        allesRadio.pack(side=tk.LEFT, padx=5)
        Tooltip(allesRadio, "Voer acties uit op alle rijen")
        
        enkelRadio = ttk.Radiobutton(
            bereikFrame,
            text="Rij:",
            variable=self.bereikVar,
            value="enkel"
        )
        enkelRadio.pack(side=tk.LEFT, padx=5)
        Tooltip(enkelRadio, "Voer acties uit op één specifieke rij")
        
        # Rij invoer voor 'enkel'
        self.enkelRijVar = tk.StringVar(value="1")
        enkelRijEntry = tk.Entry(
            bereikFrame,
            textvariable=self.enkelRijVar,
            **STIJLEN["entry"],
            width=5
        )
        enkelRijEntry.pack(side=tk.LEFT)
        
        bereikRadio = ttk.Radiobutton(
            bereikFrame,
            text="Bereik:",
            variable=self.bereikVar,
            value="bereik"
        )
        bereikRadio.pack(side=tk.LEFT, padx=5)
        Tooltip(bereikRadio, "Voer acties uit op een bereik van rijen")
        
        # Bereik invoer
        self.vanRijVar = tk.StringVar(value="1")
        vanRijEntry = tk.Entry(
            bereikFrame,
            textvariable=self.vanRijVar,
            **STIJLEN["entry"],
            width=5
        )
        vanRijEntry.pack(side=tk.LEFT)
        
        totLabel = tk.Label(
            bereikFrame,
            text="tot",
            **STIJLEN["label"]
        )
        totLabel.pack(side=tk.LEFT, padx=2)
        
        self.totRijVar = tk.StringVar(value="10")
        totRijEntry = tk.Entry(
            bereikFrame,
            textvariable=self.totRijVar,
            **STIJLEN["entry"],
            width=5
        )
        totRijEntry.pack(side=tk.LEFT)
        
        # Actielijst label
        actielijstLabel = tk.Label(
            container,
            text="Beschikbare Acties:",
            **STIJLEN["label"],
            anchor=tk.W,
            pady=5
        )
        actielijstLabel.pack(fill=tk.X, pady=(10, 0))
        
        # Scroll container voor acties
        actieScrollFrame = tk.Frame(
            container,
            background=KLEUREN["achtergrond"]
        )
        actieScrollFrame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Scrollbar
        scrollbar = tk.Scrollbar(actieScrollFrame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Canvas voor scrollbare inhoud
        self.actieCanvas = tk.Canvas(
            actieScrollFrame,
            background=KLEUREN["achtergrond"],
            yscrollcommand=scrollbar.set,
            highlightthickness=0
        )
        self.actieCanvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar.config(command=self.actieCanvas.yview)
        
        # Frame voor acties in canvas
        self.actieListFrame = tk.Frame(
            self.actieCanvas,
            background=KLEUREN["achtergrond"]
        )
        
        # Canvas window
        self.actieCanvasWindow = self.actieCanvas.create_window(
            (0, 0),
            window=self.actieListFrame,
            anchor="nw",
            width=350  # Breedte van de canvas minus scrollbar
        )
        
        # Configureer canvas om mee te schalen met frame
        self.actieListFrame.bind("<Configure>", self._configureActieCanvas)
        self.actieCanvas.bind("<Configure>", self._onCanvasResize)
        
        # Voeg acties toe aan de lijst
        self._voegActiesToe()
        
        # Uitvoerknop
        self.uitvoerButton = ttk.Button(
            container,
            text="Voer Geselecteerde Acties Uit",
            command=self.voerGeselecteerdeActiesUit
        )
        self.uitvoerButton.pack(fill=tk.X, pady=10)
        Tooltip(self.uitvoerButton, "Voer alle geselecteerde acties uit in volgorde")
    

    def _maakProductSheetTab(self):
        """Maak de inhoud van het ProductSheet aanmaken tabblad"""
        # Hoofdcontainer met padding
        container = tk.Frame(
            self.tabProductSheet,
            background=KLEUREN["achtergrond"],
            padx=20,
            pady=20
        )
        container.pack(fill=tk.BOTH, expand=True)
        
        # Label met instructie
        instructieLabel = tk.Label(
            container,
            text="Hier kunt u een nieuw ProductSheet aanmaken",
            **STIJLEN["label"],
            pady=10
        )
        instructieLabel.pack(fill=tk.X)
        
        # Placeholder voor toekomstige functionaliteit
        placeholderLabel = tk.Label(
            container,
            text="Functionaliteit nog in ontwikkeling",
            **STIJLEN["label"],
            pady=20
        )
        placeholderLabel.pack(fill=tk.X)
    def _configureActieCanvas(self, event):
        """Pas de canvas grootte aan aan het actieListFrame"""
        # Update het scrollgebied naar het nieuwe formaat van het actieListFrame
        self.actieCanvas.configure(scrollregion=self.actieCanvas.bbox("all"))
    
    def _onCanvasResize(self, event):
        """Pas de breedte van het actieListFrame aan"""
        # Pas de breedte van het window aan aan de canvas
        self.actieCanvas.itemconfig(self.actieCanvasWindow, width=event.width)
    
    def _voegActiesToe(self):
        """Voeg beschikbare acties toe aan de actielijst"""
        # Verwijder bestaande actieframes
        for widget in self.actieListFrame.winfo_children():
            widget.destroy()
        
        # Voeg kolomvullen actie toe voor elke kolom als er een Excel-bestand is geladen
        if excelHandler.isBestandGeopend():
            for kolomNaam in excelHandler.kolomNamen:
                actieFrame = tk.Frame(
                    self.actieListFrame,
                    background=KLEUREN["achtergrond"],
                    padx=5,
                    pady=5,
                    relief=tk.GROOVE,
                    borderwidth=1
                )
                actieFrame.pack(fill=tk.X, pady=2)
                
                # Selectie checkbox
                checkVar = tk.BooleanVar(value=False)
                checkBox = ttk.Checkbutton(
                    actieFrame,
                    text=f"{kolomNaam} vullen",
                    variable=checkVar
                )
                checkBox.pack(side=tk.LEFT)
                
                # Uitvoerknop voor deze actie
                uitvoerBtn = ttk.Button(
                    actieFrame,
                    text="Uitvoeren",
                    command=lambda k=kolomNaam: self.voerKolomVullenActieUit(k)
                )
                uitvoerBtn.pack(side=tk.RIGHT)
                Tooltip(uitvoerBtn, f"Vul kolom '{kolomNaam}' met gecombineerde data uit andere kolommen")
                
                # Bewaar de variabelen en widgets voor later gebruik
                actieFrame.checkVar = checkVar
                actieFrame.uitvoerBtn = uitvoerBtn
                actieFrame.kolomNaam = kolomNaam
                actieFrame.actieType = "kolomVullen"
    
    def _maakStatusbalk(self):
        """Maak de statusbalk onderaan het scherm"""
        self.statusFrame = tk.Frame(
            self.hoofdFrame,
            background=STIJLEN["status"]["background"],
            height=30
        )
        self.statusFrame.pack(fill=tk.X, side=tk.BOTTOM)
        
        # Statuslabel
        self.statusLabel = tk.Label(
            self.statusFrame,
            text="Gereed",
            **STIJLEN["status"]
        )
        self.statusLabel.pack(side=tk.LEFT, padx=10)

        # Afsluitknop
        self.afsluitButton = tk.Button(
            self.statusFrame,
            text="Afsluiten",
            command=self.bevestigAfsluiten,
            bg="#990000",  # Donkerrood
            fg="#FFFFFF",  # Wit
            font=("Arial", 10, "bold")
        )
        self.afsluitButton.pack(side=tk.RIGHT, padx=10)
        Tooltip(self.afsluitButton, "Sluit de applicatie af")
    
    def updateStatus(self, statusTekst):
        """
        Update de statustekst
        
        Args:
            statusTekst (str): Nieuwe statustekst
        """
        self.statusLabel.config(text=statusTekst)
        self.root.update_idletasks()
    
    def toonFoutmelding(self, titel, bericht):
        """
        Toon een foutmelding popup
        
        Args:
            titel (str): Titel van de foutmelding
            bericht (str): Foutmeldingstekst
        """
        messagebox.showerror(titel, bericht)
    
    def toonSuccesmelding(self, titel, bericht):
        """
        Toon een succesmelding popup
        
        Args:
            titel (str): Titel van de succesmelding
            bericht (str): Succesmeldingstekst
        """
        messagebox.showinfo(titel, bericht)
    
    def kiesExcelBestand(self):
        """Toon bestandskiezer om een Excel-bestand te selecteren"""
        bestandspad = filedialog.askopenfilename(
            title="Selecteer Excel-bestand",
            filetypes=[("Excel bestanden", "*.xlsx *.xls")]
        )
        
        if bestandspad:
            self.bestandspadVar.set(bestandspad)
    
    def onthoudBestandUpdate(self):
        """Update de 'onthoud bestand' instelling"""
        onthoud = self.onthoudBestandVar.get()
        instellingen.stelOnthoudBestandIn(onthoud)
        
        # Forceer visuele update van de checkbox
        if onthoud:
            self.onthoudBestandCheck.select()
        else:
            self.onthoudBestandCheck.deselect()
    def laadGeselecteerdBestand(self):
        """Laad het geselecteerde Excel-bestand"""
        bestandspad = self.bestandspadVar.get()
        
        if not bestandspad:
            self.toonFoutmelding("Fout", "Geen bestand geselecteerd")
            return
        
        self.laadExcelBestand(bestandspad)
    
    def laadExcelBestand(self, bestandspad):
        """
        Laad een Excel-bestand en update de UI
        
        Args:
            bestandspad (str): Pad naar het Excel-bestand
        """
        self.updateStatus(f"Bezig met laden van {os.path.basename(bestandspad)}...")
        
        # Laad het bestand
        if excelHandler.openBestand(bestandspad):
            # Update UI
            self.bestandspadVar.set(bestandspad)
            self.bestandsInfoLabel.config(text=f"Bestand: {os.path.basename(bestandspad)}")
            self.rijInfoLabel.config(text=f"Rijen: {excelHandler.haalRijAantal()}")
            self.kolomInfoLabel.config(text=f"Kolommen: {len(excelHandler.kolomNamen)}")
            
            # Sla op als laatste bestand indien nodig
            if self.onthoudBestandVar.get():
                instellingen.stelLaatsteBestandIn(bestandspad)
            
            # Update actielijst met nieuwe kolommen
            self._voegActiesToe()
            
            self.updateStatus("Bestand succesvol geladen")
            self.toonSuccesmelding("Succes", f"Bestand '{os.path.basename(bestandspad)}' succesvol geladen")
        else:
            self.updateStatus("Fout bij laden bestand")
            self.toonFoutmelding("Fout", f"Kon bestand '{bestandspad}' niet laden")
    
    def haalGeselecteerdBereik(self):
        """
        Haal het geselecteerde regelbereik op
        
        Returns:
            tuple: (startRij, eindRij) of None als 'alles' is geselecteerd
        """
        bereikType = self.bereikVar.get()
        
        try:
            if bereikType == "enkel":
                rij = int(self.enkelRijVar.get()) - 1  # Pas aan voor 0-index
                return (rij, rij)
            elif bereikType == "bereik":
                vanRij = int(self.vanRijVar.get()) - 1  # Pas aan voor 0-index
                totRij = int(self.totRijVar.get()) - 1  # Pas aan voor 0-index
                return (vanRij, totRij)
            else:  # "alles"
                return None
        except ValueError:
            self.toonFoutmelding("Fout", "Ongeldige rijwaarden. Gebruik gehele getallen.")
            return None
    
    def voerKolomVullenActieUit(self, kolomNaam):
        """
        Voer een kolomVullen actie uit voor de gegeven kolom
        
        Args:
            kolomNaam (str): Naam van de kolom om te vullen
        """
        if not excelHandler.isBestandGeopend():
            self.toonFoutmelding("Fout", "Geen Excel-bestand geopend")
            return
        
        # Vraag om bronkolommen en formaat
        
        # Toon kolomkeuzedialoog
        bronKolommen = self._toonKolomKeuzeDlg()
        if not bronKolommen:
            return
        
        # Vraag om formaat string
        formaat = simpledialog.askstring(
            "Formaat", 
            "Geef het formaat op. Gebruik {kolomnaam} voor waarden uit kolommen.\nVoorbeeld: '{Voornaam} {Achternaam}'"
        )
        
        if not formaat:
            return
        
        # Maak parameters
        parameters = {
            "doelKolom": kolomNaam,
            "bronKolommen": bronKolommen,
            "formaat": formaat
        }
        
        # Bepaal bereik
        bereik = self.haalGeselecteerdBereik()
        
        # Voer actie uit
        self.updateStatus(f"Bezig met vullen van kolom '{kolomNaam}'...")
        resultaat = voerActieUit("kolomVullen", parameters, bereik)
        
        if resultaat.succes:
            self.updateStatus("Kolom succesvol gevuld")
            self.toonSuccesmelding("Succes", resultaat.bericht)
            # Vraag of gebruiker wil opslaan
            opslaan = messagebox.askyesno(
                "Opslaan",
                "Wil je de wijzigingen opslaan in het Excel-bestand?"
            )
            if opslaan and excelHandler.slaOp():
                self.updateStatus("Wijzigingen opgeslagen")
        else:
            self.updateStatus("Fout bij vullen kolom")
            self.toonFoutmelding("Fout", resultaat.bericht)
    
    def _toonKolomKeuzeDlg(self):
        """
        Toon een dialoog om bronkolommen te selecteren
        
        Returns:
            list: Lijst met geselecteerde kolomnamen of None bij annuleren
        """
        # Maak een nieuw dialoogvenster
        dlg = tk.Toplevel(self.root)
        dlg.title("Selecteer Bronkolommen")
        dlg.geometry("300x400")
        dlg.transient(self.root)
        dlg.grab_set()
        
        # Centereer op het hoofdvenster
        dlg.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - dlg.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dlg.winfo_height()) // 2
        dlg.geometry(f"+{x}+{y}")
        
        # Instructie label
        instructieLabel = tk.Label(
            dlg,
            text="Selecteer de kolommen die je wilt gebruiken als bron",
            **STIJLEN["label"],
            wraplength=280,
            justify=tk.LEFT,
            padx=10,
            pady=10
        )
        instructieLabel.pack(fill=tk.X)
        
        # Frame voor checkboxes
        checkFrame = tk.Frame(dlg, background=KLEUREN["achtergrond"])
        checkFrame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Canvas voor scrolling
        canvas = tk.Canvas(
            checkFrame,
            background=KLEUREN["achtergrond"],
            highlightthickness=0
        )
        scrollbar = tk.Scrollbar(checkFrame, orient="vertical", command=canvas.yview)
        
        scrollFrame = tk.Frame(canvas, background=KLEUREN["achtergrond"])
        
        scrollFrame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollFrame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Voeg checkboxes toe voor elke kolom
        checkVars = {}
        for kolom in excelHandler.kolomNamen:
            var = tk.BooleanVar(value=False)
            checkbox = ttk.Checkbutton(
                scrollFrame,
                text=kolom,
                variable=var
            )
            checkbox.pack(anchor=tk.W, pady=2)
            checkVars[kolom] = var
        
        # Buttons
        buttonFrame = tk.Frame(dlg, background=KLEUREN["achtergrond"])
        buttonFrame.pack(fill=tk.X, padx=10, pady=10)
        
        # Resultaat opslag
        result = {"kolommen": None}
        
        # OK button functie
        def bevestig():
            geselecteerd = [k for k, v in checkVars.items() if v.get()]
            if not geselecteerd:
                messagebox.showwarning("Waarschuwing", "Selecteer ten minste één kolom")
                return
            
            result["kolommen"] = geselecteerd
            dlg.destroy()
        
        # Annuleren button functie
        def annuleer():
            dlg.destroy()
        
        # Buttons
        okButton = ttk.Button(buttonFrame, text="OK", command=bevestig)
        cancelButton = ttk.Button(buttonFrame, text="Annuleren", command=annuleer)
        
        okButton.pack(side=tk.RIGHT, padx=5)
        cancelButton.pack(side=tk.RIGHT, padx=5)
        
        # Wacht tot dialoog sluit
        self.root.wait_window(dlg)
        
        return result["kolommen"]
    
    def voerGeselecteerdeActiesUit(self):
        """Voer alle geselecteerde acties uit in volgorde"""
        if not excelHandler.isBestandGeopend():
            self.toonFoutmelding("Fout", "Geen Excel-bestand geopend")
            return
        
        # Verzamel geselecteerde acties
        geselecteerdeActies = []
        
        for widget in self.actieListFrame.winfo_children():
            if hasattr(widget, 'checkVar') and widget.checkVar.get():
                if widget.actieType == "kolomVullen":
                    geselecteerdeActies.append((widget.actieType, widget.kolomNaam))
        
        if not geselecteerdeActies:
            self.toonFoutmelding("Fout", "Geen acties geselecteerd")
            return
        
        # Maak een tijdelijke workflow
        workflow = workflowManager.maakWorkflow("temp_workflow")
        
        # Doorloop geselecteerde acties en voeg ze toe aan de workflow
        for actieType, kolomNaam in geselecteerdeActies:
            if actieType == "kolomVullen":
                # Vraag om bronkolommen en formaat
                bronKolommen = self._toonKolomKeuzeDlg()
                if not bronKolommen:
                    continue
                
                # Vraag om formaat string
                formaat = simpledialog.askstring(
                    "Formaat", 
                    f"Geef het formaat op voor kolom {kolomNaam}.\nGebruik {{kolomnaam}} voor waarden uit kolommen.\nVoorbeeld: '{{Voornaam}} {{Achternaam}}'"
                )
                
                if not formaat:
                    continue
                
                # Maak parameters
                parameters = {
                    "doelKolom": kolomNaam,
                    "bronKolommen": bronKolommen,
                    "formaat": formaat
                }
                
                # Voeg toe aan workflow
                workflow.voegActieToe(actieType, parameters)
        
        if not workflow.acties:
            self.toonFoutmelding("Info", "Geen acties geconfigureerd")
            return
        
        # Bepaal bereik
        bereik = self.haalGeselecteerdBereik()
        
        # Voer workflow uit met voortgangsupdate
        def updateVoortgang(percentage, actieNaam):
            self.updateStatus(f"Uitvoeren: {actieNaam} ({percentage:.1f}%)")
            self.root.update_idletasks()
        
        self.updateStatus("Bezig met uitvoeren van acties...")
        succes = workflow.voerUit(updateVoortgang, bereik)
        
        # Verwijder tijdelijke workflow
        workflowManager.verwijderWorkflow("temp_workflow")
        
        if succes:
            # Vraag of gebruiker het resultaat wil opslaan
            opslaan = messagebox.askyesno(
                "Opslaan",
                "Acties succesvol uitgevoerd. Wil je de wijzigingen opslaan?"
            )
            
            if opslaan:
                if excelHandler.slaOp():
                    self.updateStatus("Wijzigingen opgeslagen")
                    self.toonSuccesmelding("Succes", "Wijzigingen zijn opgeslagen")
                else:
                    self.updateStatus("Fout bij opslaan")
                    self.toonFoutmelding("Fout", "Kon wijzigingen niet opslaan")
            else:
                self.updateStatus("Wijzigingen niet opgeslagen")
        else:
            self.updateStatus("Fout bij uitvoeren acties")
            self.toonFoutmelding("Fout", "Er is een fout opgetreden bij het uitvoeren van de acties")