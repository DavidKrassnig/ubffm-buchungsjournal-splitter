import os
import re
import json
import csv
import win32com.client as win32
from datetime import datetime
import tqdm	# Erweiterte Fortschrittsanzeige

verzeichnis_pfad = "buchungsjournale"  # Pfad zum Verzeichnis (anpassen)
csv_datei_pfad = "personal_db.csv"  # Pfad zur CSV-Datei (anpassen)

def datumKonvertierer(yyyymm_date):
    try:
        date_obj = datetime.strptime(yyyymm_date, '%Y%m')
        german_month_names = [
            "Januar", "Februar", "März", "April", "Mai", "Juni",
            "Juli", "August", "September", "Oktober", "November", "Dezember"
        ]
        full_month_name = german_month_names[date_obj.month - 1]  # Subtract 1 to account for 0-based index
        full_date = f"{full_month_name} {date_obj.year}"
        return full_date
    except ValueError:
        return "Ungültiges Datumsformat"

def datumExtrahierer(filename):
    pattern = r'_(\d+)_\w+\.pdf$'
    match = re.search(pattern, filename)
    if match:
        return match.group(1)
    else:
        return None

def personalnummerExtrahierer(dateiname):
    treffer = re.search(r'_(\d+)\.pdf$', dateiname)
    if treffer:
        return treffer.group(1)
    return None

def dictionaryFlacher(originalDict):
    flacheresDict = {}
    
    for key, value in originalDict.items():
        if isinstance(value, dict):
            for subkey, subvalue in value.items():
                flacheresDict[subkey] = subvalue
        else:
            flacheresDict[key] = value
    
    return flacheresDict

def csvLeser(csv_datei):
    daten = {}
    with open(csv_datei, 'r', newline='') as datei:
        csv_leser = csv.reader(datei)
        titelzeile = next(csv_leser)  # Titelzeile überspringen
        for zeile in csv_leser:
            if len(zeile) >= 4:
                personalnummer = zeile[0]
                nachname = zeile[1]
                vorname = zeile[2]
                pdf = zeile[3]
                daten[personalnummer] = {
                    "Nachname": nachname,
                    "Vorname": vorname,
                    "E-Mail": pdf
                }
    return daten

def datenbanks():
    if os.path.exists(csv_datei_pfad) and os.path.isfile(csv_datei_pfad):
        csv_daten = csvLeser(csv_datei_pfad)

        if os.path.exists(verzeichnis_pfad) and os.path.isdir(verzeichnis_pfad):
            datenbank = datenbankErsteller(verzeichnis_pfad, csv_daten)
            csv_daten = None
            return datenbank
        else:
            print(f"Das Verzeichnis '{verzeichnis_pfad}' existiert nicht oder ist kein Verzeichnis.")
            exit(1)
    else:
        print(f"Die CSV-Datei '{csv_datei_pfad}' existiert nicht oder ist keine Datei.")
        exit(1)

def datenbankErsteller(pfad, csv_daten):
    datenbank = {}
    for wurzel, verzeichnisse, dateien in os.walk(pfad):
        aktuell = datenbank
        ordner = wurzel.split(os.path.sep)[1:]
        for ordnername in ordner:
            aktuell = aktuell.setdefault(ordnername, {})
        for datei in dateien:
            personalnummer = None
            if datei.lower().endswith('.pdf'):
                vollerPdfPfad = os.path.abspath(os.path.join(wurzel, datei))
                personalnummer = personalnummerExtrahierer(datei)
                datum = datumKonvertierer(datumExtrahierer(datei))
                if personalnummer in csv_daten:
                    aktuell[datei] = {
                        "Personalnummer": personalnummer,
                        "Nachname": csv_daten[personalnummer]["Nachname"],
                        "Vorname": csv_daten[personalnummer]["Vorname"],
                        "E-Mail": csv_daten[personalnummer]["E-Mail"],
                        "Pfad": vollerPdfPfad,  # Add the full file path
                        "Datum": datum
                    }
                else:
                    aktuell[datei] = {
                        "Personalnummer": personalnummer,
                        "Nachname": None,
                        "Vorname": None,
                        "E-Mail": None,
                        "Pfad": vollerPdfPfad,  # Add the full file path
                        "Datum": datum
                    }

    return datenbank

def auswahlAbteilung(abteilung):
    while True:
        print("\nVerfügbare Abteilungen:")
        subkeys = list(abteilung.keys())
        print("0. Alle Abteilungen")
        for i, key in enumerate(subkeys):
            print(f"{i + 1}. {key}")

        choice = input("Geben Sie die Zahl der auszuwählenden Abteilung ein:  ")
        try:
            choice = int(choice)
            if 0 <= choice <= len(subkeys):
                if choice == 0:
                    return dictionaryFlacher(abteilung)
                else:
                    selected_subkey = subkeys[choice - 1]
                    return abteilung[selected_subkey]
            else:
                print("Keine valide Wahl. Bitte wählen Sie eine valide Nummer.")
        except ValueError:
            print("Keine valide Wahl. Bitte wählen Sie eine valide Nummer.")

def auswahlGesamt(datenbank):
    while True:
        print("\nVerfügbare Zeiträume:")
        main_keys = sorted(list(datenbank.keys()), reverse=True)
        for i, key in enumerate(main_keys):
            print(f"{i + 1}. {key}")

        choice = input("Geben Sie die Zahl des auszuwählenden Zeitraumes ein: ")
        try:
            choice = int(choice)
            if 1 <= choice <= len(main_keys):
                selected_key = main_keys[choice - 1]
                abteilung = datenbank[selected_key]
                selected_abteilung = auswahlAbteilung(abteilung)
                return selected_abteilung
                break  # Exit the loop after selecting a subkey
            else:
                print("Keine valide Wahl. Bitte wählen Sie eine valide Nummer.")
        except ValueError:
            print("Keine valide Wahl. Bitte wählen Sie eine valide Nummer.")

def mailTextErzeuger(vorname, nachname, pdf, path, personalnummer, datum):
    if None in (vorname, nachname, pdf, path):
        return "Fehler!"
    else:
        message = f"Guten Tag {vorname} {nachname},\n\nBitte finden Sie dieser E-Mail beigefügt Ihr Buchungsjournal für {datum}.\n\n\nMit freundlichen Grüßen\n\nIhre Personalabteilung\n(dies ist eine automatisch generierte E-Mail)"
        return message

keineMail = []
eineMail = []

def mailVerschicker(pdfs):
    buchungsjournale_email_account_name = 'Buchungsjournale@ub.uni-frankfurt.de'
    outlook = win32.Dispatch('outlook.application')
    namespace = outlook.GetNamespace('MAPI')
    buchungsjournale_email_account = None

    for account in namespace.Accounts:
        if account.DisplayName == buchungsjournale_email_account_name:
            buchungsjournale_email_account = account
            break
    if buchungsjournale_email_account is None:
        print(f"E-Mail-Account '{buchungsjournale_email_account_name}' wurde nicht gefunden.")
        exit(1)
        
    with tqdm.tqdm(total=len(pdfs)) as bar:
        for pdf in pdfs:
            vorname = pdfs[pdf]['Vorname']
            nachname = pdfs[pdf]['Nachname']
            emailAdresse = pdfs[pdf]['E-Mail']
            anhangPfad = pdfs[pdf]['Pfad']
            personalnummer = pdfs[pdf]['Personalnummer']
            datum = pdfs[pdf]['Datum']
            message = mailTextErzeuger(vorname, nachname, emailAdresse, anhangPfad, personalnummer, datum)
            if message == 'Fehler!':
                keineMail.append(pdf)
                bar.update(1)
            else:
                mail = outlook.CreateItem(0)

                mail._oleobj_.Invoke(*(64209, 0, 8, 0, buchungsjournale_email_account))  # Select the pdf account

                mail.To = emailAdresse
                mail.Subject = f'[Buchungsjournal] {datum}'
                mail.Body = message
                mail.Attachments.Add(anhangPfad)

                mail.Send()
                
                eineMail.append(pdf)
                
                bar.update(1)

def dateiPrint(text):
            print(text)
            output_file.write(text+'\n')

if __name__ == '__main__':
    datenbank = datenbanks()
    auswahl = auswahlGesamt(datenbank)
    mailVerschicker(auswahl)
    anzahlKeineMail = len(keineMail)
    anzahlEineMail = len(eineMail)
    anzahlMail = len(auswahl)
    jetzt = datetime.now().strftime('%Y%m%dT%H%M%S')
    with open(f'logs/log_{jetzt}.txt', "w") as output_file:
        dateiPrint(f'Es wurden {anzahlEineMail}/{anzahlMail} Buchungsjournale erfolgreich verschickt!')
        print(f'Details zu (nicht) verschickten E-Mails stehen in der Log-Datei (log_{jetzt}.txt).')
        output_file.write('--------------------------------------------------\n')
        output_file.write('Folgende Buchungsjournale wurden via E-Mail verschickt:\n')
        for i in eineMail:
            output_file.write(i+'\n')
        output_file.write('--------------------------------------------------\n')
        output_file.write('Folgende Buchungsjournale konnten wegen fehlenden Informationen nicht via E-Mail verschickt werden:\n')
        for i in keineMail:
            output_file.write(i+'\n')

    input('\nDrücken Sie die Eingabetaste, um das Skript zu beenden...')
    exit(0)
