import fitz 		# PyMuPDF für PDF-Befehle
import re			# Regex
import tqdm     	# Erweiterte Fortschrittsanzeige
import sys			# Interaktion mit Terminal
import os			# Interaktion mit Betriebssystemen
import shutil       # Einfachere Dateiverschiebung

if os.name == 'nt':
    from colorama import just_fix_windows_console
    just_fix_windows_console()

class formatierung:	# Definiert verschiedene Formatierungsmöglichkeiten
    PURPLE = '\033[95m'
    CYAN = '\033[96m'
    DARKCYAN = '\033[36m'
    BLUE = '\033[94m'
    GREEN = '\033[92m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
    END = '\033[0m'

def dateiNamenRegeln(s): # Vereinheitlicht Namen der Dateiausgabe (optimiert für Interaktionen im Terminal)
    # Entfernt alle nicht-Wörterzeichen (alles abgesehen von Buchstaben und Ziffern)
    s = re.sub(r"[^\w\s]", '', s)
    # Ersetzt alle Leerzeichen mit Unterstrichen
    s = re.sub(r"\s+", '_', s)
    # Ersetzt alle Schriftzeichen mit deren Kleinschrift-Version
    s = s.lower()
    return s

def pdfDateiSucheBuchungsjournal(pdfDatei, suchText):
    doc = fitz.open(pdfDatei)
    buchungsjournalSeiten = []

    with tqdm.tqdm(total=doc.page_count) as bar:
        for seiteJetzt in range(doc.page_count):
            page = doc.load_page(seiteJetzt)
            text = page.get_text()
            if suchText in text:
                buchungsjournalSeiten.append(seiteJetzt)
            bar.update(1)

    doc.close()
    return buchungsjournalSeiten

def pdfDateiPersonaldaten(pdfDatei):
    doc = fitz.open(pdfDatei)
    anfangsdatum, enddatum = None, None
    familienname, vorname, personalnummer, abteilung = None, None, None, None

    for seiteJetzt in range(doc.page_count):
        page = doc.load_page(seiteJetzt)
        text = page.get_text()

        datum = re.search(r'Buchungsjournal (\d{2}).(\d{2}).(\d{4})\s*-\s*(\d{2}).(\d{2}).(\d{4})', text, re.IGNORECASE)
        personaldaten = re.search(r'Name: (.*), (.*) PersonalNr.: (.*) Abteilung: (.*)', text, re.IGNORECASE)

        if datum:
            anfangsdatum = str(datum.group(3)+'-'+datum.group(2))
            enddatum = str(datum.group(6)+'-'+datum.group(5))
        else:
            print(formatierung.BOLD+formatierung.RED+f'ABBRUCH! Kein Datum für {pdfDatei} gefunden.'+formatierung.END)
            exit(1)
            
        if personaldaten:
            familienname = personaldaten.group(1)
            vorname = personaldaten.group(2)
            personalnummer = personaldaten.group(3)
            abteilung = personaldaten.group(4)
            break
        else:
            print(formatierung.BOLD+formatierung.RED+f'ABBRUCH! Keine Personaldaten für {pdfDatei} gefunden.'+formatierung.END)
            exit(1)

    doc.close()
    return anfangsdatum, enddatum, familienname, vorname, personalnummer, abteilung

def pdfDateiAuftrennungNachSeiten(pdfDatei, seiteGesamt, tempVerzeichnis):
    doc = fitz.open(pdfDatei)
    
    if not os.path.exists(tempVerzeichnis):
        os.makedirs(tempVerzeichnis)

    with tqdm.tqdm(total=len(seiteGesamt)) as bar:
        for i, seiteJetzt in enumerate(seiteGesamt):
            seiteStart = seiteJetzt
            seiteEnde = seiteGesamt[i + 1] if i + 1 < len(seiteGesamt) else doc.page_count

            neuDoc = fitz.open()
            for page in range(seiteStart, seiteEnde):
                neuDoc.insert_pdf(doc, from_page=page, to_page=page)

            neuPdfDatei = f'{tempVerzeichnis}/pages_{seiteStart + 1}_to_{seiteEnde}.pdf'
            neuDoc.save(neuPdfDatei)
            neuDoc.close()
            bar.update(1)

    doc.close()

def pdfDateiUmbenennung(tempVerzeichnis):
    pdfDateien = [pdfFile for pdfFile in os.listdir(tempVerzeichnis) if pdfFile.endswith('.pdf')]
    gesamtDateien = len(pdfDateien)

    with tqdm.tqdm(total=gesamtDateien) as bar:
        for pdfDatei in pdfDateien:
            datenExtrahiert = pdfDateiPersonaldaten(os.path.join(tempVerzeichnis, pdfDatei))
            if datenExtrahiert:
                anfangsdatum, enddatum, familienname, vorname, personalnummer, abteilung = datenExtrahiert
                abteilungKurz = ''.join(re.split("[^a-zA-Z]*", abteilung))
                neuPdfName = dateiNamenRegeln(f'Buchungsjournal {anfangsdatum} {familienname} {vorname} {personalnummer}') + '.pdf'
                neuPdfPfad = tempVerzeichnis + '/' + f'{anfangsdatum}' + '/' + dateiNamenRegeln(f'{abteilungKurz}') + '/'
                if not os.path.exists(neuPdfPfad):
                    os.makedirs(neuPdfPfad)
                neuPdfDatei = f'{neuPdfPfad}/{neuPdfName}'
                os.rename(os.path.join(tempVerzeichnis, pdfDatei), neuPdfDatei)
            bar.update(1)


def machTitel(s):
    return formatierung.BOLD+s+formatierung.END

def meldungErfolgreich():
   print(formatierung.BOLD+formatierung.GREEN+'Erfolgreich!'+formatierung.END)

def meldungFehlschlag():
   print(formatierung.BOLD+formatierung.RED+'Fehler!'+formatierung.END)

def ueberschreibenFrage(ziel):
    response = input(formatierung.BOLD+formatierung.YELLOW+f'{os.path.basename(ziel)} existiert bereits im Ausgabeverzeichnis. Überschreiben?'+formatierung.END+' [j/N]\n')
    return response.lower().strip() == 'j'

tempVerzeichnis = '.temp_ubffm_pdf_splitter'
suchText = 'Buchungsjournal'
ausgabeVerzeichnis = 'buchungsjournale'

def trennlinie():
    w, h = shutil.get_terminal_size()
    print("—" * w)

if __name__ == '__main__':
    if len(sys.argv) == 1:
        # Falls keine PDF-Datei spezifiziert wird, nehme alle PDF-Dateien im Skriptordner
        pdfDateien = [file for file in os.listdir() if file.endswith('.pdf')]
        if not pdfDateien:
            print(formatierung.BOLD + formatierung.RED + 'Keine PDF-Dateien gefunden.' + formatierung.END)
            exit(1)
        else:
            print(formatierung.BOLD + formatierung.GREEN + f'Gefundene PDF-Dateien: '+ formatierung.END+f'{", ".join(pdfDateien)}')
    elif len(sys.argv) == 2:
        pdfDateien = [sys.argv[1]]
    else:
        print(formatierung.BOLD + formatierung.RED + 'Bitte geben Sie nur den Dateipfad der PDF-Datei an.' + formatierung.END)
        exit(1)

    for pdfDatei in pdfDateien:
        trennlinie()
        print(machTitel(f'Analysiere PDF-Datei: {pdfDatei}'))
        buchungsjournalSeiten = pdfDateiSucheBuchungsjournal(pdfDatei, suchText)
        if not buchungsjournalSeiten:
            meldungFehlschlag()
            continue

        meldungErfolgreich()

        print(machTitel(f'Extrahiere Buchungsjournale: {pdfDatei}'))
        pdfDateiAuftrennungNachSeiten(pdfDatei, buchungsjournalSeiten, tempVerzeichnis)
        meldungErfolgreich()

        print(machTitel(f'Umbenennung der extrahierten Buchungsjournale: {pdfDatei}'))
        pdfDateiUmbenennung(tempVerzeichnis)
        meldungErfolgreich()

    trennlinie()
    # Verschiebe produzierte Dateien in das Ausgabeverzeichnis
    print(machTitel('Verschiebe Dateien zum Ausgabeverzeichnis'))
    if not os.path.exists(ausgabeVerzeichnis):
        os.makedirs(ausgabeVerzeichnis)
    tempVerzeichnis_items = os.listdir(tempVerzeichnis)

    with tqdm.tqdm(total=len(tempVerzeichnis_items)) as bar:
        for item in tempVerzeichnis_items:
            quelle = os.path.join(tempVerzeichnis, item)
            ziel = os.path.join(ausgabeVerzeichnis, item)
            if os.path.exists(ziel):
                user_response = ueberschreibenFrage(ziel)
                if not user_response:
                    print(f'Überspringe {item}...')
                    bar.update(1)
                    continue
                if os.path.isdir(ziel):
                    shutil.rmtree(ziel)  # Entferne das Zielverzeichnis
                else:
                    os.remove(ziel)  # Entferne die Zieldatei
            if os.path.isdir(quelle):
                shutil.move(quelle, ziel)
            else:
                shutil.copy2(quelle, ziel)
            bar.update(1)

    # Entferne das temporäre Verzeichnis
    if os.path.exists(tempVerzeichnis):
        shutil.rmtree(tempVerzeichnis)
    meldungErfolgreich()
    trennlinie()
    input('Drücken Sie die Eingabetaste, um das Skript zu beenden')
    exit(0)
