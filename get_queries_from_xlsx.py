import sys
import xml.etree.ElementTree as ET
import zipfile


def get_queries(filename):
    """
    Get External queries from Excel file
    """
    # Unzip connection file
    with zipfile.ZipFile(filename) as zf:
        xml = zf.read("xl/connections.xml")

    root = ET.fromstring(xml)

    for connections in root.findall('connections:connection',
                                     {'connections': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}):
        for connection in connections.findall('.//x15:connection/x15:oledbPr/x15:dbCommand',
                                               {'x15': 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main'}):
            print(connection.attrib.get('text'))
            print()


if __name__ == '__main__': 
    get_queries(sys.argv[1])
