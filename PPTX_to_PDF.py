from comtypes.client import CreateObject
from comtypes import COMError
import comtypes
import os
import logging
 
module_logger = logging.getLogger("Certificater.PPTX_toPDF")

def init_powerpoint():
    powerpoint = CreateObject('PowerPoint.Application')
    powerpoint.UserControl = 0
    powerpoint.Visible = 1
    return powerpoint

def convertation(powerpoint, inputFileName, outputFileName, formatType = 32):
    if outputFileName[-3:] != 'pdf':
        outputFileName = outputFileName.replace(".pptx","").replace(".ppt","").replace("GENERATED_PPTX","GENERATED_PDF") + ".pdf"
    try:
        deck = powerpoint.Presentations.Open(inputFileName)
        deck.SaveAs(outputFileName, formatType) # formatType = 32 для ppt в pdf
        deck.Close()
    except COMError:
        # Если случилась ошибка и сертификат не создается
        powerpoint.Quit()
        powerpoint = init_powerpoint()
        convertation(powerpoint, inputFileName, outputFileName)

def pptx_to_pdf(file_name, today_date, powerpoint):
    cwd = f"{os.getcwd()}\\GENERATED_PPTX\\{today_date}\\{file_name}.pptx"  # Создаём полный путь до файла
    convertation(powerpoint, cwd, cwd)  # Запуск конвертации