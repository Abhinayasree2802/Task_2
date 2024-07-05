



# import zipfile
# import os
# import xml.etree.ElementTree as ET
# import tkinter as tk
# from tkinter import filedialog

# # Define namespaces
# namespaces = {
#     'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
#     'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
#     'a16': 'http://schemas.microsoft.com/office/drawing/2014/main',
#     'a14': 'http://schemas.microsoft.com/office/drawing/2010/main',
#     's': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
# }

# # Function to parse XML
# def parse_xml(file_path):
#     try:
#         tree = ET.parse(file_path)
#         return tree.getroot()
#     except FileNotFoundError as e:
#         print(f"File not found: {file_path}")
#         raise e
#     except ET.ParseError as e:
#         print(f"Error parsing XML: {e}")
#         raise e

# # Function to parse shared strings XML
# def parse_shared_strings(file_path):
#     try:
#         tree = ET.parse(file_path)
#         root = tree.getroot()
#         strings = {}
#         for idx, si_elem in enumerate(root.findall('s:si', namespaces)):
#             texts = [t_elem.text for t_elem in si_elem.findall('s:t', namespaces) if t_elem.text is not None]
#             text = ''.join(texts).strip()
#             if text:
#                 strings[idx] = text
#         return strings
#     except FileNotFoundError as e:
#         print(f"File not found: {file_path}")
#         raise e
#     except ET.ParseError as e:
#         print(f"Error parsing sharedStrings XML: {e}")
#         raise e

# # Function to convert Excel column letters to numbers
# def col_to_num(col):
#     num = 0
#     for c in col:
#         num = num * 26 + (ord(c.upper()) - ord('A')) + 1
#     return num

# # Function to parse sheet XML and map shared strings to cells
# def parse_sheet(file_path, shared_strings):
#     try:
#         tree = ET.parse(file_path)
#         root = tree.getroot()
#         data = []
#         for row_elem in root.findall('.//s:row', namespaces):
#             row_num = int(row_elem.attrib['r'])
#             for cell_elem in row_elem.findall('s:c', namespaces):
#                 col_ref = cell_elem.attrib['r']
#                 col_num = ''.join([ch for ch in col_ref if ch.isalpha()])
#                 row_num_in_cell = int(''.join([ch for ch in col_ref if ch.isdigit()]))
#                 if row_num != row_num_in_cell:
#                     continue
#                 if 't' in cell_elem.attrib and cell_elem.attrib['t'] == 's':
#                     shared_string_idx = int(cell_elem.find('s:v', namespaces).text)
#                     text = shared_strings.get(shared_string_idx, '')
#                     if text:
#                         data.append({
#                             'from_row': row_num,
#                             'from_col': col_to_num(col_num),
#                             'to_row': row_num,
#                             'to_col': col_to_num(col_num),
#                             'text': text,
#                             'source': 'sharedString'
#                         })
#         return data
#     except FileNotFoundError as e:
#         print(f"File not found: {file_path}")
#         raise e
#     except ET.ParseError as e:
#         print(f"Error parsing sheet XML: {e}")
#         raise e

# # Function to unzip the Excel file and extract XML files
# def unzip_excel(file_path, extract_to):
#     with zipfile.ZipFile(file_path, 'r') as zip_ref:
#         zip_ref.extractall(extract_to)

# def main():
#     # Using tkinter to select files
#     root = tk.Tk()
#     root.withdraw()

#     try:
#         # Select Excel file
#         excel_file = filedialog.askopenfilename(title="Select Excel file")
#         if not excel_file:
#             print("No file selected. Exiting.")
#             return
        
#         extract_to = os.path.join(os.path.dirname(excel_file), 'extracted_files')
        
#         # Unzip the Excel file
#         unzip_excel(excel_file, extract_to)

#         # File paths for drawing1.xml, sharedStrings.xml, and sheet1.xml
#         drawing_xml_file = os.path.join(extract_to, 'xl', 'drawings', 'drawing1.xml')
#         shared_strings_xml_file = os.path.join(extract_to, 'xl', 'sharedStrings.xml')
#         sheet_xml_file = os.path.join(extract_to, 'xl', 'worksheets', 'sheet1.xml')
        
#         # Ask user for output XML file
#         output_file = filedialog.asksaveasfilename(defaultextension=".xml", title="Save output as")
#         if not output_file:
#             print("No output file selected. Exiting.")
#             return

#         # Parse XML files
#         drawing_root = parse_xml(drawing_xml_file)
#         shared_strings = parse_shared_strings(shared_strings_xml_file)
#         sheet_data = parse_sheet(sheet_xml_file, shared_strings)

#         # Extract data
#         # Note: extract_drawing_data function is not fully provided here; ensure it is defined as per your requirements

#         # Combine and sort data
#         combined_data = sheet_data  # Add drawing_data here if needed
#         sorted_data = sorted(combined_data, key=lambda x: (x['from_row'], x['from_col'], x['to_row'], x['to_col']))

#         # Write sorted data to output file
#         # Note: write_sorted_data function is not fully provided here; ensure it is defined as per your requirements

#         print(f"Processing completed successfully. Sorted XML has been written to {output_file}")

#     except Exception as e:
#         print(f"An error occurred: {str(e)}")

#     finally:
#         root.destroy()  # Close the tkinter window after completion or error

# if __name__ == "__main__":
#     main()



# import zipfile
# import os
# import xml.etree.ElementTree as ET
# import tkinter as tk
# from tkinter import filedialog

# # Define namespaces
# namespaces = {
#     'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
#     'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
#     'a16': 'http://schemas.microsoft.com/office/drawing/2014/main',
#     'a14': 'http://schemas.microsoft.com/office/drawing/2010/main',
#     's': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
# }

# # Function to parse XML
# def parse_xml(file_path):
#     try:
#         tree = ET.parse(file_path)
#         return tree.getroot()
#     except FileNotFoundError:
#         print(f"File not found: {file_path}")
#         exit(1)
#     except ET.ParseError as e:
#         print(f"Error parsing XML: {e}")
#         exit(1)

# # Function to parse shared strings XML
# def parse_shared_strings(file_path):
#     try:
#         tree = ET.parse(file_path)
#         root = tree.getroot()
#         strings = {}
#         for idx, si_elem in enumerate(root.findall('s:si', namespaces)):
#             texts = [t_elem.text for t_elem in si_elem.findall('s:t', namespaces) if t_elem.text is not None]
#             text = ''.join(texts).strip()
#             if text:
#                 strings[idx] = text
#         return strings
#     except FileNotFoundError:
#         print(f"File not found: {file_path}")
#         exit(1)
#     except ET.ParseError as e:
#         print(f"Error parsing sharedStrings XML: {e}")
#         exit(1)

# # Function to convert Excel column letters to numbers
# def col_to_num(col):
#     num = 0
#     for c in col:
#         num = num * 26 + (ord(c.upper()) - ord('A')) + 1
#     return num

# # Function to parse sheet XML and map shared strings to cells
# def parse_sheet(file_path, shared_strings):
#     try:
#         tree = ET.parse(file_path)
#         root = tree.getroot()
#         data = []
#         for row_elem in root.findall('.//s:row', namespaces):
#             row_num = int(row_elem.attrib['r'])
#             for cell_elem in row_elem.findall('s:c', namespaces):
#                 col_ref = cell_elem.attrib['r']
#                 col_num = ''.join([ch for ch in col_ref if ch.isalpha()])
#                 row_num_in_cell = int(''.join([ch for ch in col_ref if ch.isdigit()]))
#                 if row_num != row_num_in_cell:
#                     continue
#                 if 't' in cell_elem.attrib and cell_elem.attrib['t'] == 's':
#                     shared_string_idx = int(cell_elem.find('s:v', namespaces).text)
#                     text = shared_strings.get(shared_string_idx, '')
#                     if text:
#                         data.append({
#                             'from_row': row_num,
#                             'from_col': col_to_num(col_num),
#                             'to_row': row_num,
#                             'to_col': col_to_num(col_num),
#                             'text': text,
#                             'source': 'sharedString'
#                         })
#         return data
#     except FileNotFoundError:
#         print(f"File not found: {file_path}")
#         exit(1)
#     except ET.ParseError as e:
#         print(f"Error parsing sheet XML: {e}")
#         exit(1)

# # Function to extract data from drawing XML
# def extract_drawing_data(root):
#     data = []
#     for anchor in root.findall('xdr:twoCellAnchor', namespaces):
#         from_elem = anchor.find('xdr:from', namespaces)
#         to_elem = anchor.find('xdr:to', namespaces)
#         tx_body = anchor.find('.//xdr:txBody', namespaces)

#         if from_elem is not None and to_elem is not None and tx_body is not None:
#             from_row = int(from_elem.find('xdr:row', namespaces).text)
#             from_col = int(from_elem.find('xdr:col', namespaces).text)
#             to_row = int(to_elem.find('xdr:row', namespaces).text)
#             to_col = int(to_elem.find('xdr:col', namespaces).text)

#             paragraphs = []
#             for para in tx_body.findall('.//a:p', namespaces):
#                 texts = [t_elem.text for t_elem in para.findall('.//a:t', namespaces) if t_elem.text is not None]
#                 paragraph_text = ' '.join(texts).strip()
#                 if paragraph_text:
#                     paragraphs.append(paragraph_text)
#             text = '\n'.join(paragraphs).strip()

#             if text:
#                 data.append({
#                     'from_row': from_row,
#                     'from_col': from_col,
#                     'to_row': to_row,
#                     'to_col': to_col,
#                     'text': text,
#                     'source': 'drawing'
#                 })

#     return data

# # Function to write sorted data to output XML
# def write_sorted_data(sorted_data, output_file):
#     new_root = ET.Element("sorted_data")

#     for item in sorted_data:
#         if item['source'] == 'drawing':
#             anchor = ET.SubElement(new_root, "twoCellAnchor")
#             text_elem = ET.SubElement(anchor, "text")
#             text_elem.text = item['text']

#         elif item['source'] == 'sharedString':
#             si_elem = ET.SubElement(new_root, "si")
#             t_elem = ET.SubElement(si_elem, "t")
#             t_elem.text = item['text']

#     new_tree = ET.ElementTree(new_root)
#     new_tree.write(output_file, encoding='utf-8', xml_declaration=True)
#     print(f"Sorted XML has been written to {output_file}")

# # Function to unzip the Excel file and extract XML files
# def unzip_excel(file_path, extract_to):
#     with zipfile.ZipFile(file_path, 'r') as zip_ref:
#         zip_ref.extractall(extract_to)

# # Function to select and process Excel file using tkinter
# def select_and_process_excel():
#     # Using tkinter to select files
#     root = tk.Tk()
#     root.withdraw()

#     # Select Excel file
#     excel_file = filedialog.askopenfilename(title="Select Excel file")
#     if not excel_file:
#         print("No file selected. Exiting.")
#         return

#     # Create a directory to extract files if it doesn't exist
#     extract_to = os.path.join(os.path.dirname(excel_file), 'extracted_files')
#     if not os.path.exists(extract_to):
#         os.makedirs(extract_to)
    
#     # Unzip the Excel file
#     unzip_excel(excel_file, extract_to)

#     # File paths for drawing1.xml, sharedStrings.xml, and sheet1.xml
#     drawing_xml_file = os.path.join(extract_to, 'xl', 'drawings', 'drawing1.xml')
#     shared_strings_xml_file = os.path.join(extract_to, 'xl', 'sharedStrings.xml')
#     sheet_xml_file = os.path.join(extract_to, 'xl', 'worksheets', 'sheet1.xml')
#     output_file = filedialog.asksaveasfilename(defaultextension=".xml", title="Save output as")

#     # Parse XML files
#     drawing_root = parse_xml(drawing_xml_file)
#     shared_strings = parse_shared_strings(shared_strings_xml_file)
#     sheet_data = parse_sheet(sheet_xml_file, shared_strings)

#     # Extract data
#     drawing_data = extract_drawing_data(drawing_root)

#     # Combine and sort data
#     combined_data = drawing_data + sheet_data
#     sorted_data = sorted(combined_data, key=lambda x: (x['from_row'], x['from_col'], x['to_row'], x['to_col']))

#     # Write sorted data to output file
#     write_sorted_data(sorted_data, output_file)

# # Main entry point of the script
# if __name__ == "__main__":
#     select_and_process_excel()


# import zipfile
# import os
# import xml.etree.ElementTree as ET
# import tkinter as tk
# from tkinter import filedialog

# # Define namespaces
# namespaces = {
#     'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
#     'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
#     'a16': 'http://schemas.microsoft.com/office/drawing/2014/main',
#     'a14': 'http://schemas.microsoft.com/office/drawing/2010/main',
#     's': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
# }

# # Function to parse XML
# def parse_xml(file_path):
#     try:
#         tree = ET.parse(file_path)
#         return tree.getroot()
#     except FileNotFoundError:
#         print(f"File not found: {file_path}")
#         exit(1)
#     except ET.ParseError as e:
#         print(f"Error parsing XML: {e}")
#         exit(1)

# # Function to parse shared strings XML
# def parse_shared_strings(file_path):
#     try:
#         tree = ET.parse(file_path)
#         root = tree.getroot()
#         strings = {}
#         for idx, si_elem in enumerate(root.findall('s:si', namespaces)):
#             texts = [t_elem.text for t_elem in si_elem.findall('s:t', namespaces) if t_elem.text is not None]
#             text = ''.join(texts).strip()
#             if text:
#                 strings[idx] = text
#         return strings
#     except FileNotFoundError:
#         print(f"File not found: {file_path}")
#         exit(1)
#     except ET.ParseError as e:
#         print(f"Error parsing sharedStrings XML: {e}")
#         exit(1)

# # Function to convert Excel column letters to numbers
# def col_to_num(col):
#     num = 0
#     for c in col:
#         num = num * 26 + (ord(c.upper()) - ord('A')) + 1
#     return num

# # Function to parse sheet XML and map shared strings to cells
# def parse_sheet(file_path, shared_strings):
#     try:
#         tree = ET.parse(file_path)
#         root = tree.getroot()
#         data = []
#         for row_elem in root.findall('.//s:row', namespaces):
#             row_num = int(row_elem.attrib['r'])
#             for cell_elem in row_elem.findall('s:c', namespaces):
#                 col_ref = cell_elem.attrib['r']
#                 col_num = ''.join([ch for ch in col_ref if ch.isalpha()])
#                 row_num_in_cell = int(''.join([ch for ch in col_ref if ch.isdigit()]))
#                 if row_num != row_num_in_cell:
#                     continue
#                 if 't' in cell_elem.attrib and cell_elem.attrib['t'] == 's':
#                     shared_string_idx = int(cell_elem.find('s:v', namespaces).text)
#                     text = shared_strings.get(shared_string_idx, '')
#                     if text:
#                         data.append({
#                             'from_row': row_num,
#                             'from_col': col_to_num(col_num),
#                             'to_row': row_num,
#                             'to_col': col_to_num(col_num),
#                             'text': text,
#                             'source': 'sharedString'
#                         })
#         return data
#     except FileNotFoundError:
#         print(f"File not found: {file_path}")
#         exit(1)
#     except ET.ParseError as e:
#         print(f"Error parsing sheet XML: {e}")
#         exit(1)

# # Function to extract data from drawing XML
# def extract_drawing_data(root):
#     data = []
#     for anchor in root.findall('xdr:twoCellAnchor', namespaces):
#         from_elem = anchor.find('xdr:from', namespaces)
#         to_elem = anchor.find('xdr:to', namespaces)
#         tx_body = anchor.find('.//xdr:txBody', namespaces)

#         if from_elem is not None and to_elem is not None and tx_body is not None:
#             from_row = int(from_elem.find('xdr:row', namespaces).text)
#             from_col = int(from_elem.find('xdr:col', namespaces).text)
#             to_row = int(to_elem.find('xdr:row', namespaces).text)
#             to_col = int(to_elem.find('xdr:col', namespaces).text)

#             paragraphs = []
#             for para in tx_body.findall('.//a:p', namespaces):
#                 texts = [t_elem.text for t_elem in para.findall('.//a:t', namespaces) if t_elem.text is not None]
#                 paragraph_text = ' '.join(texts).strip()
#                 if paragraph_text:
#                     paragraphs.append(paragraph_text)
#             text = '\n'.join(paragraphs).strip()

#             if text:
#                 data.append({
#                     'from_row': from_row,
#                     'from_col': from_col,
#                     'to_row': to_row,
#                     'to_col': to_col,
#                     'text': text,
#                     'source': 'drawing'
#                 })

#     return data

# # Function to write sorted data to output XML
# def write_sorted_data(sorted_data, output_file):
#     new_root = ET.Element("sorted_data")

#     for item in sorted_data:
#         if item['source'] == 'drawing':
#             anchor = ET.SubElement(new_root, "paragraph")
#             text_elem = ET.SubElement(anchor, "text")
#             text_elem.text = item['text']

#         elif item['source'] == 'sharedString':
#             t_elem = ET.SubElement(new_root, "t")
#             t_elem.text = item['text']

#     new_tree = ET.ElementTree(new_root)
#     new_tree.write(output_file, encoding='utf-8', xml_declaration=True)
#     print(f"Sorted XML has been written to {output_file}")

# # Function to unzip the Excel file and extract XML files
# def unzip_excel(file_path, extract_to):
#     with zipfile.ZipFile(file_path, 'r') as zip_ref:
#         zip_ref.extractall(extract_to)

# # Function to select and process Excel file using tkinter
# def select_and_process_excel():
#     # Using tkinter to select files
#     root = tk.Tk()
#     root.withdraw()

#     # Select Excel file
#     excel_file = filedialog.askopenfilename(title="Select Excel file")
#     if not excel_file:
#         print("No file selected. Exiting.")
#         return

#     # Create a directory to extract files if it doesn't exist
#     extract_to = os.path.join(os.path.dirname(excel_file), 'extracted_files')
#     if not os.path.exists(extract_to):
#         os.makedirs(extract_to)
    
#     # Unzip the Excel file
#     unzip_excel(excel_file, extract_to)

#     # File paths for drawing1.xml, sharedStrings.xml, and sheet1.xml
#     drawing_xml_file = os.path.join(extract_to, 'xl', 'drawings', 'drawing1.xml')
#     shared_strings_xml_file = os.path.join(extract_to, 'xl', 'sharedStrings.xml')
#     sheet_xml_file = os.path.join(extract_to, 'xl', 'worksheets', 'sheet1.xml')
#     output_file = filedialog.asksaveasfilename(defaultextension=".xml", title="Save output as")

#     # Parse XML files
#     drawing_root = parse_xml(drawing_xml_file)
#     shared_strings = parse_shared_strings(shared_strings_xml_file)
#     sheet_data = parse_sheet(sheet_xml_file, shared_strings)

#     # Extract data
#     drawing_data = extract_drawing_data(drawing_root)

#     # Combine and sort data
#     combined_data = drawing_data + sheet_data
#     sorted_data = sorted(combined_data, key=lambda x: (x['from_row'], x['from_col'], x['to_row'], x['to_col']))

#     # Write sorted data to output file
#     write_sorted_data(sorted_data, output_file)

# # Main entry point of the script
# if __name__ == "__main__":
#     select_and_process_excel()




import zipfile
import os
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import filedialog

# Define namespaces
namespaces = {
    'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'a16': 'http://schemas.microsoft.com/office/drawing/2014/main',
    'a14': 'http://schemas.microsoft.com/office/drawing/2010/main',
    's': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
}

# Function to parse XML
def parse_xml(file_path):
    try:
        tree = ET.parse(file_path)
        return tree.getroot()
    except FileNotFoundError:
        print(f"File not found: {file_path}")
        exit(1)
    except ET.ParseError as e:
        print(f"Error parsing XML: {e}")
        exit(1)

# Function to parse shared strings XML
def parse_shared_strings(file_path):
    try:
        tree = ET.parse(file_path)
        root = tree.getroot()
        strings = {}
        for idx, si_elem in enumerate(root.findall('s:si', namespaces)):
            texts = [t_elem.text for t_elem in si_elem.findall('s:t', namespaces) if t_elem.text is not None]
            text = ''.join(texts).strip()
            if text:
                strings[idx] = text
        return strings
    except FileNotFoundError:
        print(f"File not found: {file_path}")
        exit(1)
    except ET.ParseError as e:
        print(f"Error parsing sharedStrings XML: {e}")
        exit(1)

# Function to convert Excel column letters to numbers
def col_to_num(col):
    num = 0
    for c in col:
        num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num

# Function to parse sheet XML and map shared strings to cells
def parse_sheet(file_path, shared_strings):
    try:
        tree = ET.parse(file_path)
        root = tree.getroot()
        data = []
        for row_elem in root.findall('.//s:row', namespaces):
            row_num = int(row_elem.attrib['r'])
            for cell_elem in row_elem.findall('s:c', namespaces):
                col_ref = cell_elem.attrib['r']
                col_num = ''.join([ch for ch in col_ref if ch.isalpha()])
                row_num_in_cell = int(''.join([ch for ch in col_ref if ch.isdigit()]))
                if row_num != row_num_in_cell:
                    continue
                if 't' in cell_elem.attrib and cell_elem.attrib['t'] == 's':
                    shared_string_idx = int(cell_elem.find('s:v', namespaces).text)
                    text = shared_strings.get(shared_string_idx, '')
                    if text:
                        data.append({
                            'from_row': row_num,
                            'from_col': col_to_num(col_num),
                            'to_row': row_num,
                            'to_col': col_to_num(col_num),
                            'text': text,
                            'source': 'sharedString'
                        })
        return data
    except FileNotFoundError:
        print(f"File not found: {file_path}")
        exit(1)
    except ET.ParseError as e:
        print(f"Error parsing sheet XML: {e}")
        exit(1)

# Function to extract data from drawing XML
def extract_drawing_data(root):
    data = []
    for anchor in root.findall('xdr:twoCellAnchor', namespaces):
        from_elem = anchor.find('xdr:from', namespaces)
        to_elem = anchor.find('xdr:to', namespaces)
        tx_body = anchor.find('.//xdr:txBody', namespaces)

        if from_elem is not None and to_elem is not None and tx_body is not None:
            from_row = int(from_elem.find('xdr:row', namespaces).text)
            from_col = int(from_elem.find('xdr:col', namespaces).text)
            to_row = int(to_elem.find('xdr:row', namespaces).text)
            to_col = int(to_elem.find('xdr:col', namespaces).text)

            paragraphs = []
            for para in tx_body.findall('.//a:p', namespaces):
                texts = [t_elem.text for t_elem in para.findall('.//a:t', namespaces) if t_elem.text is not None]
                paragraph_text = ' '.join(texts).strip()
                if paragraph_text:
                    paragraphs.append(paragraph_text)
            text = '\n'.join(paragraphs).strip()

            if text:
                data.append({
                    'from_row': from_row,
                    'from_col': from_col,
                    'to_row': to_row,
                    'to_col': to_col,
                    'text': text,
                    'source': 'drawing'
                })

    return data

# Function to write sorted data to output XML in the new structure
def write_sorted_data(sorted_data, output_file):
    # Create the new structure
    document = ET.Element("Document")

    # Metadata Section
    metadata = ET.SubElement(document, "Metadata")
    title = ET.SubElement(metadata, "Title")
    title.text = "Document Title"  # Modify as needed
    author = ET.SubElement(metadata, "Author")
    author.text = "Author Name"  # Modify as needed
    date = ET.SubElement(metadata, "Date")
    date.text = "Date of Document"  # Modify as needed

    # Content Section
    content = ET.SubElement(document, "Content")

    # Variable to keep track of the current section
    current_section = None

    for item in sorted_data:
        if item['source'] == 'drawing':
            # Assuming drawings indicate a new paragraph
            if current_section is None:
                current_section = ET.SubElement(content, "Section")
            paragraph = ET.SubElement(current_section, "Paragraph")
            sentences = item['text'].split('\n')
            for sentence in sentences:
                sent_elem = ET.SubElement(paragraph, "Sentence")
                sent_elem.text = sentence
        elif item['source'] == 'sharedString':
            # Assuming shared strings indicate a title of a new section
            current_section = ET.SubElement(content, "Section")
            title = ET.SubElement(current_section, "Title")
            title.text = item['text']

    # Write the new XML tree to the output file
    new_tree = ET.ElementTree(document)
    new_tree.write(output_file, encoding='utf-8', xml_declaration=True)
    print(f"Sorted XML has been written to {output_file}")

# Function to unzip the Excel file and extract XML files
def unzip_excel(file_path, extract_to):
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        zip_ref.extractall(extract_to)

# Function to select and process Excel file using tkinter
def select_and_process_excel():
    # Using tkinter to select files
    root = tk.Tk()
    root.withdraw()

    # Select Excel file
    excel_file = filedialog.askopenfilename(title="Select Excel file")
    if not excel_file:
        print("No file selected. Exiting.")
        return

    # Create a directory to extract files if it doesn't exist
    extract_to = os.path.join(os.path.dirname(excel_file), 'extracted_files')
    if not os.path.exists(extract_to):
        os.makedirs(extract_to)
    
    # Unzip the Excel file
    unzip_excel(excel_file, extract_to)

    # File paths for drawing1.xml, sharedStrings.xml, and sheet1.xml
    drawing_xml_file = os.path.join(extract_to, 'xl', 'drawings', 'drawing1.xml')
    shared_strings_xml_file = os.path.join(extract_to, 'xl', 'sharedStrings.xml')
    sheet_xml_file = os.path.join(extract_to, 'xl', 'worksheets', 'sheet1.xml')
    output_file = filedialog.asksaveasfilename(defaultextension=".xml", title="Save output as")

    # Parse XML files
    drawing_root = parse_xml(drawing_xml_file)
    shared_strings = parse_shared_strings(shared_strings_xml_file)
    sheet_data = parse_sheet(sheet_xml_file, shared_strings)

    # Extract data
    drawing_data = extract_drawing_data(drawing_root)

    # Combine and sort data
    combined_data = drawing_data + sheet_data
    sorted_data = sorted(combined_data, key=lambda x: (x['from_row'], x['from_col'], x['to_row'], x['to_col']))

    # Write sorted data to output file
    write_sorted_data(sorted_data, output_file)

# Main entry point of the script
if __name__ == "__main__":
    select_and_process_excel()
