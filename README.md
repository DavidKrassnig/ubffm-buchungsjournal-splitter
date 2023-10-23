# UBFFM Buchungsjournal Splitter
Dies sind zwei Python-Skripte, die ich für die digitale Verteilung von Buchungsjournale an der UBFFM erstellt habe.

## Installation
Um die Skripte einwandfrei laufen lassen zu können, benötigt man Python 3.x und folgende Libraries, die nicht standardmäßig inkludiert sind: PyMuPDF, pywin32, tqdm

Da Github keine leeren Ordner erlaubt, muss eventuell noch ein Ordner "logs" erstellt werden.

## Nutzung
### Splitter
Um eine Buchungsjournal-Sammel-PDF-Datei in die einzelnen Buchungsjournale zu spalten, muss sich die entsprechende PDF-Datei in demselben Ordner wie die Skriptdatei befinden. Danach gibt es zwei Optionen: Entweder man gibt beim Aufruf des Skriptes explizit an, welche PDF-Datei gesplittet werden soll oder man führt das Skript aus, ohne dies zu spezifizieren. Im letzteren Fall werden einfach alle PDF-Dateien, die sich im Ordner befinden gesplittet.

Die gesplitteten Buchungsjournale befinden sich nach der Ausführung des Skripts in dem Unterordner "buchungsjournale" und sind dort jeweils nach Zeitraum und Abteilung sortiert.

#### E-Mail
Nachdem durch den Splitter Buchungsjournale in die entsprechenden Unterordner verteilt worden sind, kann das E-Mail-Skript aufgerufen werden, um diese automatisch an alle entsprechenden Mitarbeiter zu versenden. Damit dies funktioniert müssen folgende Konditionen erfüllt sein:
* Es besteht eine personal_db.csv Datei und sie wurde mit den entsprechenden Informationen gefüllt
* Es besteht existiert mindestens ein Buchungsjournal
* Es muss Outlook auf dem lokalen Rechner installiert sein
* Der spezifizierte E-Mail-Account muss in der lokalen Outlook-Installation eingerichtet sein

Sind diese Bedingungen erfüllt, so kann man über das Auswahl-Menü entscheiden, welcher Zeitraum verschickt werden soll und ob innerhalb des ausgewählten Zeitraumes alle Buchungsjournale oder nur die Buchungsjournale einer bestimmten Abteilung verschickt werden sollen.

Sollte ein Mitarbeiter ein Buchungsjournal erhalten sollen, hat aber keinen entsprechenden Eintrag in personal_db.csv, so wird dies in der produzierten Log-Datei protokolliert und es wird auch direkt nach Abschluss des Skriptes angezeigt, wie viele E-Mails erfolgreich verschickt werden konnten.
