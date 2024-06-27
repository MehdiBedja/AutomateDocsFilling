from docx import Document

def read_exposant_list(file_path):
    # Read the table data from the Word document
    doc = Document(file_path)
    table = doc.tables[0]

    data = []
    keys = None
    for i, row in enumerate(table.rows):
        text = [cell.text for cell in row.cells]
        if i == 0:
            keys = tuple(text)
            continue
        row_data = dict(zip(keys, text))
        data.append(row_data)
    
    return data

def print_exposant_info(exposant_list):
    for person in exposant_list:
        print("Person Information:")
        for key, value in person.items():
            print(f"{key}: {value}")
        print("\n" + "-"*20 + "\n")

def main():
    exposant_list_path = r'C:\Users\DELL\Downloads\EXPOSANT_LISTE_5.docx'
    exposant_list = read_exposant_list(exposant_list_path)
    print_exposant_info(exposant_list)

if __name__ == '__main__':
    main()
