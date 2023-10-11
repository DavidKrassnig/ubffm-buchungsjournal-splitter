import fitz 		# PyMuPDF für PDF-Befehle
import re			# Regex
import progressbar	# Erweiterte Fortschrittsanzeige
import sys			# Für 
import os			# 

class formatierung:
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

def dateiNamenRegeln(s):
    # Remove all non-word characters (everything except numbers and letters)
    s = re.sub(r"[^\w\s]", '', s)
    # Replace all runs of whitespace with a single underscore
    s = re.sub(r"\s+", '_', s)
    # Replace all characters with their lowercase counterparts
    s = s.lower()
    return s

def pdfDateiSucheBuchungsjournal(pdfDatei, suchText):
    doc = fitz.open(pdfDatei)
    buchungsjournalSeiten = []

    with progressbar.ProgressBar(max_value=doc.page_count) as bar:
        for seiteJetzt in range(doc.page_count):
            page = doc.load_page(seiteJetzt)
            text = page.get_text()
            if suchText in text:
                buchungsjournalSeiten.append(seiteJetzt)
            bar.update(seiteJetzt)

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
            anfangsdatum = str(datum.group(3)+datum.group(2)+datum.group(1))
            enddatum = str(datum.group(6)+datum.group(5)+datum.group(4))
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

def pdfDateiAuftrennungNachSeiten(pdfDatei, seiteGesamt, ausgabeVerzeichnis):
    doc = fitz.open(pdfDatei)
    
    if not os.path.exists(ausgabeVerzeichnis):
        os.makedirs(ausgabeVerzeichnis)

    with progressbar.ProgressBar(max_value=len(seiteGesamt)) as bar:
        for i, seiteJetzt in enumerate(seiteGesamt):
            seiteStart = seiteJetzt
            seiteEnde = seiteGesamt[i + 1] if i + 1 < len(seiteGesamt) else doc.page_count

            neuDoc = fitz.open()
            for page in range(seiteStart, seiteEnde):
                neuDoc.insert_pdf(doc, from_page=page, to_page=page)

            neuPdfDatei = f'{ausgabeVerzeichnis}/pages_{seiteStart + 1}_to_{seiteEnde}.pdf'
            neuDoc.save(neuPdfDatei)
            neuDoc.close()
            bar.update(i)

    doc.close()

def pdfDateiUmbenennung(ausgabeVerzeichnis):
    with progressbar.ProgressBar(max_value=len(os.listdir(ausgabeVerzeichnis))) as bar:
        for pdfDatei in os.listdir(ausgabeVerzeichnis):
            if pdfDatei.endswith('.pdf'):
                datenExtrahiert = pdfDateiPersonaldaten(os.path.join(ausgabeVerzeichnis, pdfDatei))
                if datenExtrahiert:
                    anfangsdatum, enddatum, familienname, vorname, personalnummer, abteilung = datenExtrahiert
                    neu_pdfDatei = f'{ausgabeVerzeichnis}/'+dateiNamenRegeln(f'Buchungsjournal {anfangsdatum} {enddatum} {familienname} {vorname} {personalnummer}')+'.pdf'
                    os.rename(os.path.join(ausgabeVerzeichnis, pdfDatei), neu_pdfDatei)
                bar.update()

ausgabeVerzeichnis = 'temporary_files'  # Create a directory to save split PDFs

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(formatierung.BOLD+formatierung.RED+'Bitte geben Sie den Dateipfad der PDF-Datei an.'+formatierung.END)
        exit(1)

    pdfDatei = sys.argv[1]
    suchText = 'Buchungsjournal'
    print(formatierung.BOLD+'Analysiere PDF-Datei:'+formatierung.END)
    buchungsjournalSeiten = pdfDateiSucheBuchungsjournal(pdfDatei, suchText)
    if not buchungsjournalSeiten:
        print(formatierung.BOLD+formatierung.RED+'Fehlgeschlagen!'+formatierung.END)
        exit(1)
    else:
        print(formatierung.BOLD+formatierung.GREEN+'Erfolgreich!'+formatierung.END+'\n')

    print(formatierung.BOLD+'Extrahiere Buchungsjournale:'+formatierung.END)
    pdfDateiAuftrennungNachSeiten(pdfDatei, buchungsjournalSeiten, ausgabeVerzeichnis)
    print(formatierung.BOLD+formatierung.GREEN+'Erfolgreich!'+formatierung.END+'\n')

    print(formatierung.BOLD+'Umbenennung der extrahierten Buchungsjournale:'+formatierung.END)
    pdfDateiUmbenennung(ausgabeVerzeichnis)
    print(formatierung.BOLD+formatierung.GREEN+'Erfolgreich!'+formatierung.END)
    exit(0)
