from ctypes import sizeof
import io
import os

# Imports the Google Cloud client library
from google.cloud import vision

# Instantiates a client
client = vision.ImageAnnotatorClient()

# The name of the image file to annotate
filename = input("Please enter the image name of your receipt: ")
file_name = os.path.abspath('../../' + filename)

# Loads the image into memory
with io.open(file_name, 'rb') as image_file:
    content = image_file.read()

image = vision.Image(content=content)

# Performs label detection on the image file
response = client.label_detection(image=image)
labels = response.label_annotations

def detect_text(path):
    """Detects text in the file."""
    from google.cloud import vision
    import io
    client = vision.ImageAnnotatorClient()

    with io.open(path, 'rb') as image_file:
        content = image_file.read()

    image = vision.Image(content=content)

    response = client.text_detection(image=image)
    texts = response.text_annotations
    if response.error.message:
        raise Exception(
            '{}\nFor more info on error messages, check: '
            'https://cloud.google.com/apis/design/errors'.format(
                response.error.message))
    return texts[0].description


def createArrays(data):
    foundSaleTran = False
    foundAllItems = False
    foundAllPrices = False
    items = []
    prices = []
    for dataPiece in data:
        if "Items in Transaction" in dataPiece:
            foundAllPrices = True
            continue
        if dataPiece == "":
            continue
        if foundSaleTran and not foundAllItems:
            if dataPiece[0] == '$':
                foundAllItems = True
                prices.append(float(dataPiece[1:]))
            elif "@" in dataPiece:
                continue
            else:
                items.append(dataPiece)
        elif dataPiece == "SALE TRANSACTION":
            foundSaleTran = True
        elif foundAllItems and dataPiece[0] == '$' and not foundAllPrices:
            prices.append(float(dataPiece[1:]))
        elif foundAllPrices:
            if dataPiece[0] == '$':
                prices.append(float(dataPiece[1:]))
                break
            if dataPiece == "Balance to pay":
                items.append("Total")
    return (items, prices)
        
def createSpreadsheet(items, prices):
    import xlsxwriter

    workbook = xlsxwriter.Workbook("Precise.xlsx")
    bold = workbook.add_format({'bold': True}) #enabling the bold feature
    worksheet = workbook.add_worksheet()
    worksheet.write('A1', "Items", bold)
    worksheet.write('B1', "Prices", bold)
    worksheet.write('C1', "Did You Buy?", bold)
    worksheet.write_column(1,0, items)
    worksheet.write_column(1, 1, prices)
    purchased = [0] * (len(items)-1) 
    worksheet.write_column(1, 2, purchased)
    worksheet.write(0, 3, "You Pay:", bold)
    for i in range(2,len(items)+1):
        formula = '=B' + str(i) + "*C" + str(i)
        worksheet.write_formula('AA'+str(i), formula)
    formula = '=SUM(AA2:AA' + str(len(items)) + ')'
    worksheet.write_formula('E1', formula)
    workbook.close()


print()
textDesc = detect_text(file_name)
data = textDesc.split("\n")
items, prices = createArrays(data)
createSpreadsheet(items, prices)

print("Open up the spreadsheet \'precise.xlsx\' to see the parsed data!")
