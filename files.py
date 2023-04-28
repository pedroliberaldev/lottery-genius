import xml.etree.ElementTree as Elements


def read_file():
    # Open XLSX file
    opened_file = Elements.parse("results_file/results.xlsx")

    # Get file root
    file_root = opened_file.getroot()

    # Get all
    for child in file_root:
        # Acesse um atributo específico do elemento
        print(child.attrib['Apostas'])


'''
# Itere sobre os elementos filhos da raiz
for child in root:
    # Acesse um atributo específico do elemento
    print(child.attrib['atributo'])

    # Itere sobre os elementos filhos deste elemento
    for subchild in child:
        # Acesse o texto do elemento
        print(subchild.text)
'''
