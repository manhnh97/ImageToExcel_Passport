import xlsxwriter
from PIL import Image
from openpyxl import load_workbook
from pathlib import Path

def ReadExcelRoot(pathNameExcelRoot):

    wb = load_workbook(filename = pathNameExcelRoot)
    # WBsheetData = wb['label']
    WBsheetData = wb.active
    
    maxRow = WBsheetData.max_row+1
    start = 2
    while (start <= maxRow):
        # listCount.append(f"{(WBsheetData.cell(row=start, column=1).value)} | {(WBsheetData.cell(row=start, column=2).value)}")
        TextCount = (f"{(WBsheetData.cell(row=start, column=1).value)}|||{(WBsheetData.cell(row=start, column=2).value)}")
        
        yield TextCount
        start += 1

def ResizeImages(pathNameExcelRoot, folderImages, pathNameExcelWrite):
    # Create folder contain Excels safely
    ExportExcelFolder = "ExportExcelFolder"
    Path(ExportExcelFolder).mkdir(parents=True, exist_ok=True)

    # Create an new Excel file and add a worksheet
    workbook = xlsxwriter.Workbook(f"{ExportExcelFolder}/{pathNameExcelWrite}")
    worksheet = workbook.add_worksheet()

    cell_height = 30.0

    # Set Default Row
    worksheet.set_default_row(int(cell_height))
    worksheet.set_column(1,1, 50)

    # Texts in Excel Root
    TextCount = ReadExcelRoot(pathNameExcelRoot)

    numberA = numberB = numberC = 1
    for T in TextCount:
        TextToFile = T.split('|||')
        # print(TextToFile[1])
        try:
            # Parameters Image to Row Excel
            filename = f"{folderImages}/{TextToFile[0]}.png"
            img = Image.open(filename)
            image_width, image_height = img.size
            y_scale = cell_height/image_height
            cellImage = 'B' + str(numberB)
            # Insert Image to Row Excel
            worksheet.insert_image(cellImage, filename, {'x_scale': y_scale, 'y_scale': y_scale})

            # Insert Text to Row Excel
            worksheet.write('A'+str(numberA), TextToFile[0])
            worksheet.write('C'+str(numberC), TextToFile[1])
            
            numberC += 1
            numberA += 1
            numberB += 1

        except FileNotFoundError:
            continue

    workbook.close()

pathNameExcelRoot = "Mazekaz"+".xlsx"
folderImages = '32'
pathNameExcelWrite = "ImageToExcel"+".xlsx"

"""
pathNameExcelRoot = 'images - Copy.xlsx'
folderImages = 'imagesDATA'
pathNameExcelWrite = 'ImagesToExcel.xlsx'
"""

# ===>> Uncomment and Run here <<====
# ResizeImages(pathNameExcelRoot, folderImages, pathNameExcelWrite)






