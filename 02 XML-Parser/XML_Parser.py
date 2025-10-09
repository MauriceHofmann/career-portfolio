"""
Datei:      XML_Parser.py
Autor:      Maurice Hofmann
Erstellt:   08.09.2025 
Lizenz:     public

Beschreibung:
    Dieses Skript definiert eine Funktion zum Parsen von XML-Dateien, die Daten aus dem Änderungsmanagement enthalten.
    Diese Funktion extrahiert spezifische Informationen wie Nummer, Baureihe, Ausführungsart und speichert diese in einem Dictionary.

Voraussetzungen:
    - Eine gültige XML-Datei mit Daten im erwarteten Format (Änderungsmanagement-Daten)

Benötigte Bibliotheken:
    - xml.etree.ElementTree: Zum Parsen von XML-Dateien.

Verwendung:
    1.  Stelle sicher, dass die benötigten Bibliotheken installiert sind.
    2.  Passe den `filepath` in deinem Aufruf der Funktion `parse_change_management_xml` an,
        um auf die zu parsende XML-Datei zu verweisen.
    3.  Rufe die Funktion `parse_change_management_xml` mit dem Dateipfad und optional einem Namespace-Dictionary auf.
    4.  Die Funktion gibt ein Dictionary mit den extrahierten Daten zurück.

Exit Codes:
    0: Erfolgreiche Ausführung.
       Die XML-Datei wurde geparst, und die extrahierten Daten wurden erfolgreich auf der Standardausgabe ausgegeben.

    1: Bei der Ausführung ist ein Fehler aufgetreten.
       Genauere Fehlerbedingungen werden in der Standardfehlerausgabe (stderr) ausgegeben. Mögliche Fehler sind:
        - Fehlendes Kommandozeilenargument:
          Das Skript wurde ohne Angabe des XML-Dateipfads ausgeführt.
        - Datei nicht gefunden:
          Die angegebene XML-Datei existiert nicht.
        - XML-Parsing-Fehler:
          Die XML-Datei ist fehlerhaft und kann nicht geparst werden.
        - Unerwarteter Fehler beim XML-Parsing:
          Beim Parsen der XML-Datei ist eine unerwartete Ausnahme aufgetreten.
          Die genaue Fehlermeldung wird in stderr ausgegeben.
        - Fehler bei der Datenextraktion:
          Beim Extrahieren der Daten aus dem geparsten XML-Baum ist ein Fehler aufgetreten.
          Die genaue Fehlermeldung wird in stderr ausgegeben.
    """

########################
# Imports
########################
import sys
import json
import xml.etree.ElementTree as ET


def parse_file(filepath: str) -> ET.Element:
    """
    Parses an XML file and returns the root element.

    This function takes the filepath of an XML file as input, parses the file
    using `xml.etree.ElementTree`, and returns the root element of the
    resulting XML tree.

    Args:
        filepath (str): The path to the XML file.

    Returns:
        The root element of the XML tree.
    """

    return ET.parse(filepath).getroot()



def parse_change_management_xml(root: ET.Element, namespace = {'dai': ''}):
    """
    Extrahiert spezifische Datenfelder aus einer XML-Baumstruktur.

    Diese Funktion analysiert einen XML-Baum (dargestellt durch das `root`-Element)
    und extrahiert Daten aus bestimmten Elementen basierend auf ihren Tag-Namen
    und dem Namespace. Sie ruft Werte für 'Nummer', 'Baureihe',
    'Ausfuehrungsart', 'EinsatzterminZeichnung' und eine verschachtelte Struktur
    unter 'SelliSubtask' ab.

    Args:
        root: Das Root-Element des XML-Baums.
        namespace: Der XML-Namespace, der bei der Suche nach Elementen verwendet werden soll.

    Returns:
        Ein Dictionary, das die extrahierten Daten enthält. Die Schlüssel des
        Dictionarys entsprechen den Tag-Namen (z.B. "Nummer", "Baureihe").

    Raises:
        ExceptionGroup: Wenn während des Parsens Fehler auftreten. Die ExceptionGroup
            enthält eine Liste der einzelnen aufgetretenen Exceptions.
            Mögliche Exceptions sind solche, die ausgelöst werden, wenn ein bestimmtes Element
            nicht in der XML-Struktur gefunden wird.
    """

    extracted_data = {}
    errors = []


    ########################
    # Extraction
    ########################

    # Nummer
    try:
        nummer = root.find(".//dai:Nummer", namespace).text
        extracted_data["Nummer"] = nummer
    except Exception as e:
        errors.append(e)

    # Baureihe
    try:
        baureihe = root.find(".//dai:Baureihe", namespace).text
        extracted_data["Baureihe"] = baureihe
    except Exception as e:
        errors.append(e)

    # Ausführungsart
    try:
        ausfuehrungsart = root.find(".//dai:Ausfuehrungsart", namespace).text
        extracted_data["Ausfuehrungsart"] = ausfuehrungsart
    except Exception:
        errors.append(e)

    if errors:
        raise ExceptionGroup("Error parsing XML content", errors)

    return extracted_data



if __name__ == "__main__":

    if len(sys.argv) < 2:
        print("Usage: python XML_Aender.... <filepath>", file=sys.stderr)
        exit(1)

    f_name = sys.argv[1]

    try:
        root = parse_file(f_name)
    except FileNotFoundError:
        print(f"Fehler: Datei nicht gefunden: {f_name}")
        exit(1)
    except ET.ParseError:
        print(f"Fehler: Fehler beim Parsen der XML-Datei: {f_name}")
        exit(1)
    except Exception as e:
        print(f"Unerwarteter Fehler beim Parsen der XML-Datei: {e}")
        exit(1)


    try:
        result = parse_change_management_xml(root)
    except ExceptionGroup as e:
        print(f"Unerwarteter Fehler beim Parsen der XML-Datei: {e.exceptions}")
        exit(1)
    
    json.dump(result, fp=sys.stdout, indent=4)