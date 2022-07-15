import glob
import sys
import os
import configparser
import logging
from datetime import datetime
from PdfOperations import PdfOperations
from ExcelOperations import ExcelOperations
from DirectoryOperations import DirectoryOperations

logging.basicConfig(filename="pdfmarkup-" + datetime.now().strftime("%Y%m%d_%H_%M") + " .log", level=logging.INFO)
print("=== Starting Application ===")
logging.info("Starting Application." +  datetime.now().strftime("_%H_%M"))
config_object = configparser.ConfigParser()
config_object.read("config.ini")
docPath = config_object["DocumentConfig"]["DocumentPath"]
# docPath = config_object["DocumentConfig"]["DocumentPath"]
directoryReader = DirectoryOperations(docPath)
latestFolderName = directoryReader.CheckforLatestFolder()
if not latestFolderName:
    print('There is no today folder. Please try again later.')
    os.system('pause')
folders = directoryReader.ScanFolders(latestFolderName)

try:
    docPath = docPath + "\\" + latestFolderName
    directoryReader.CreateOutputDirectory()
    excOperation = ExcelOperations(config_object, docPath) # Excel Reader
    pdfOperation = PdfOperations(logging)# PDF Reader
    excelFile = glob.glob(docPath + "\\*.xlsx")
    if not excelFile:
        print("There is no unknown tag excel file here.")
        os.system('pause')

    for folder in folders:
        print("================================")
        print("Looking on folder " + folder)
        fPath = docPath + "/" + folder
        pdfFiles = glob.glob(fPath + "\\*.pdf")
        if not pdfFiles:
            print("There are no pdf files at " + fPath)
            logging.info("There are no pdf files at " + fPath)
        else:
            doc_tag_list = excOperation.getUnknownTagList(docPath, pdfFiles)
            if not doc_tag_list:
                print("No unknown tags in folder " + folder + " documents.")
                logging.info("No unknown tags in folder " + folder + " documents.")
            else:
                totalPdfFileCount = len(doc_tag_list.keys())
                
                print("Processing " + str(totalPdfFileCount) + " pdf file(s)")
                logging.info("Processing " + str(totalPdfFileCount) + " pdf file(s)")

                # HIGHLIGHT IN PDF KEY = PDF , VALUE = TagList and Doc Owner
                for key, value in doc_tag_list.items():
                    if len(value) > 0:
                        print(key + " document processing")
                        logging.info(key + " document processing")
                        pdfOperation.SearchAndHighlightSpecialCases(fPath, key, value)
                        excOperation.createOutputExcelFile(fPath, pdfOperation.not_found_tags_dict, pdfOperation.found_tags_dict, value)
                        pdfOperation.not_found_tags = []
                        pdfOperation.found_tags = []
                        pdfOperation.not_found_tags_dict = {}
                        pdfOperation.found_tags_dict = {}
                        totalPdfFileCount -= 1
                        if not totalPdfFileCount == 0:
                            print(str(totalPdfFileCount) + " pdf files remaining")
                            logging.info(str(totalPdfFileCount) + " pdf files remaining")
    print("\n")
    print("=== Highlight Completed ===")
    print("Thanks for using pdf highlighter program.")
    logging.info("Highlight Completed" + datetime.now().strftime("_%H_%M"))
    os.system('pause')                  
except (IOError, ValueError, EOFError, PermissionError) as e:
    print('oops! error occurred.')
    print(e)
    logging.error("oops! Error occurred.")
    logging.error(e)
    os.system('pause')
