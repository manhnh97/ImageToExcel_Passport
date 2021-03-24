import xlsxwriter
from PIL import Image
from openpyxl import load_workbook
from pathlib import Path

def ReadExcelRoot(pathNameExcelRoot):

    wb = load_workbook(filename = pathNameExcelRoot)
    # WBsheetData = wb['label']
    WBsheetData = wb.active
    
    # Total Row in Excel
    maxRow = WBsheetData.max_row+1

    # Ignore Fist Row because It is Header/Title
    start = 2
    while (start <= maxRow):
        # listCount.append(f"{(WBsheetData.cell(row=start, column=1).value)} | {(WBsheetData.cell(row=start, column=2).value)}")
        TextCount = (f"{(WBsheetData.cell(row=start, column=1).value)}|||{(WBsheetData.cell(row=start, column=2).value)}")
        
        yield TextCount
        start += 1

def ExportImagesToExcel(pathNameExcelRoot, folderImagesRoot, pathNameExcelWrite):
    # Create folder contain Excels safely
    ExportExcelFolder = "ExportExcelFolder"
    Path(ExportExcelFolder).mkdir(parents=True, exist_ok=True)

    # Create an new Excel file and add a worksheet
    workbook = xlsxwriter.Workbook(f"{ExportExcelFolder}/{pathNameExcelWrite}")
    worksheet = workbook.add_worksheet()

    # Height Image and Cell
    cell_height = 30.0

    # Set Default Row
    worksheet.set_default_row(int(cell_height))
    worksheet.set_column(1,1, 50)

    # Texts in Excel Root
    TextCount = ReadExcelRoot(pathNameExcelRoot)

    numberA = numberB = numberC = 1
    try:
        for T in TextCount:
            TextToFile = T.split('|||')
            # print(TextToFile[1])
            try:
                # Insert Text to Rows a.k.a Column A
                worksheet.write('A'+str(numberA), TextToFile[0])

                # Parameters Image to Row Excel
                filename = f"{folderImagesRoot}/{TextToFile[0]}.png"
                img = Image.open(filename)
                image_width, image_height = img.size
                y_scale = cell_height/image_height
                cellImage = 'B' + str(numberB)
                # Insert Image to Row Excel
                worksheet.insert_image(cellImage, filename, {'x_scale': y_scale, 'y_scale': y_scale})

                # Insert Text to Rows a.k.a Column C
                worksheet.write('C'+str(numberC), TextToFile[1])
                
                numberC += 1
                numberA += 1
                numberB += 1

            except FileNotFoundError:
                continue
    except KeyboardInterrupt:
        print(KeyboardInterrupt)
        workbook.close()
    workbook.close()

pathNameExcelRoot = "Label_CMT_20_03_2021_Real"+".xlsx"
folderImagesRoot = '32'
pathNameExcelWrite = "ImageToExcel"+".xlsx"

"""
pathNameExcelRoot = 'images - Copy.xlsx'
folderImagesRoot = 'imagesDATA'
pathNameExcelWrite = 'ImagesToExcel.xlsx'
"""

# ===>> Uncomment and Run here <<====
ExportImagesToExcel(pathNameExcelRoot, folderImagesRoot, pathNameExcelWrite)






