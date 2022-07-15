import glob
import sys
import os
import configparser
import logging
from datetime import datetime
from PdfOperations import PdfOperations
from ExcelOperations import ExcelOperations
from DirectoryOperations import DirectoryOperations

logging.basicConfig(filename="pdfmarkup-" +
                    datetime.now().strftime("%Y%m%d_%H_%M") + " .log", level=logging.INFO)
print("=== Starting Application ===")
logging.info("Starting Application." + datetime.now().strftime("_%H_%M"))
config_object = configparser.ConfigParser()
config_object.read("config.ini")
docPath = config_object["DocumentConfig"]["DocumentPath"]
directoryReader = DirectoryOperations(docPath)

try:
    directoryReader.CreateOutputDirectory()
    excOperation = ExcelOperations(config_object, docPath)  # Excel Reader
    pdfOperation = PdfOperations(logging)  # PDF Reader
    excelFile = glob.glob(docPath + "\\*.xlsx")
    if not excelFile:
        print("There is no unknown tag excel file here.")
        os.system('pause')
    print("================================")
    print("Looking on folder " + docPath)
    #fPath = docPath + "/" + folder
    pdfFiles = glob.glob(docPath + "\\*.pdf")
    if not pdfFiles:
        print("There are no pdf files at " + docPath)
        logging.info("There are no pdf files at " + docPath)
        os.system('pause')
    total_pdf_file_count = len(pdfFiles)
    print("Looking on " + str(total_pdf_file_count) + " pdf files")
    doc_tag_dict = {}
    tag_list = excOperation.readTagList(docPath)
    total_tag_count = len(tag_list)
    if not tag_list:
        print("No tag list in file " + excelFile)
        logging.info("No tag list in file " + excelFile)
    else:
        for pdf_file in pdfFiles:
            print("Searching " + str(total_tag_count) + " tags in " + pdf_file)
            logging.info("Searching " + str(total_tag_count) +
                         " tags in " + pdf_file)
            pdf_name = os.path.basename(pdf_file)
            found_tag_list = []
            # HIGHLIGHT IN PDF KEY = PDF , VALUE = TagList and Doc Owner
            for tags in tag_list:
                pdfOperation.SearchForTags(
                    docPath, pdf_name, tags, doc_tag_dict, found_tag_list)
    
            doc_tag_dict[pdf_name] = found_tag_list
            total_pdf_file_count -= 1
            if not total_pdf_file_count == 0:
                print(str(total_pdf_file_count) + " pdf files remaining")
                logging.info(str(total_pdf_file_count) + " pdf files remaining")
        excOperation.createDocTag(doc_tag_dict)
        print("\n")
        print("=== Doc Tag Completed ===")
        logging.info("Doc Tag Completed" + datetime.now().strftime("_%H_%M"))
        os.system('pause')
except (IOError, ValueError, EOFError, PermissionError) as e:
    print('oops! error occurred.')
    print(e)
    logging.error("oops! Error occurred.")
    logging.error(e)
    os.system('pause')
