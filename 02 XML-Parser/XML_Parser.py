"""
File:        XML_Parser.py
Author:      Maurice Hofmann
Created:     08.09.2025
License:     Public

Description:
    This script defines a function for parsing XML files containing change management data. 
    The function extracts specific information such as number, series, and configuration type, 
    and stores them in a dictionary.

Requirements:
    - A valid XML file containing data in the expected change management format.

Required Libraries:
    - xml.etree.ElementTree: For parsing XML files.

Usage:
    1. Ensure that all required libraries are installed.
    2. Adjust the `filepath` argument in your function call to `parse_change_management_xml` 
       to point to the XML file you want to parse.
    3. Call the function `parse_change_management_xml` with the file path and optionally a namespace dictionary.
    4. The function returns a dictionary containing the extracted data.

Exit Codes:
    0: Successful execution.  
       The XML file was parsed successfully, and the extracted data were printed to the standard output.

    1: An error occurred during execution.  
       Detailed error messages are printed to the standard error output (stderr). Possible causes include:
        - Missing command-line argument:  
          The script was run without specifying the XML file path.
        - File not found:  
          The specified XML file does not exist.
        - XML parsing error:  
          The XML file is malformed and cannot be parsed.
        - Unexpected parsing error:  
          An unexpected exception occurred while parsing the XML file.  
          The exact error message will be printed to stderr.
        - Data extraction error:  
          An error occurred while extracting data from the parsed XML tree.  
          The detailed error message will be printed to stderr.
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
    Extracts specific data fields from an XML tree structure.

    This function analyzes an XML tree (represented by the `root` element)
    and extracts data from certain elements based on their tag names
    and namespace. It retrieves values for 'Nummer', 'Baureihe',
    'Ausfuehrungsart', 'EinsatzterminZeichnung', and a nested structure
    under 'SelliSubtask'.

    Args:
        root: The root element of the XML tree.
        namespace: The XML namespace to be used when searching for elements.

    Returns:
        A dictionary containing the extracted data. The dictionary keys correspond
        to the tag names (e.g., "Nummer", "Baureihe").

    Raises:
        ExceptionGroup: Raised if errors occur during parsing. The ExceptionGroup
            contains a list of individual exceptions that were encountered.
            Possible exceptions include missing elements that could not be found
            in the XML structure.
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