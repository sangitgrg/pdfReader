import fitz
from PyPDF2 import PdfFileWriter, PdfFileReader
import os
import fileinput
import NotHighlighted
import shutil


class PdfOperations:
    not_found_tags = []
    found_tags = []
    ifFoundAny = False
    not_found_tags_dict = {}
    found_tags_dict = {}
    
    #for bookmark
    bm_tag_page = []
    bm_tag_dict = {}

    def __init__(self, logging):
        self.logging = logging

    def CreateDirectory(self, dirPath):
        if not os.path.exists(dirPath):
            os.mkdir(dirPath)

    def BookMarkTags(self, pdf, tag_page_dict, new_pdf_name, docOwner):
        pdf_object = open(pdf, 'rb')
        input1 = PdfFileReader(pdf_object)
        output = PdfFileWriter()
        input_numpages = input1.getNumPages()
        # basically just copy the input file
        for i in range(input_numpages):
            output.addPage(input1.getPage(i))
        # insert page in the output file
        for key, value in tag_page_dict.items():
            if len(value) > 0:
                parent_1 = output.addBookmark(key, value[0])
            for val in value:
                output.addBookmark(str(val+1), val, parent_1)
        outputStream = open("Output\\" + docOwner + "\\" +
                            new_pdf_name + "_withbookmark.pdf", 'wb')
        output.write(outputStream)
        outputStream.close()
        pdf_object.close()
        self.bm_tag_dict = {}

    def SearchAndHighlight(self, filePath, pdf, searchText):
        doc = fitz.open(filePath + "\\" + pdf, filetype="pdf")
        if(doc.pageCount > 1000 or len(searchText) > 200):
            print("======= SKIPPED DOCUMENT ===========")
            print(pdf)
            self.logging.info(
                pdf + " document is skipped due to number of pages greater then 1000 and unknown tags more then 10.")
            print("=====================================")
            return
        for text in searchText:
            for page in doc:
                text_instances = page.searchFor(text.tagNumber)
                if len(text_instances) > 0:
                    self.ifFoundAny = True
                    for foundText in text_instances:   # HIGHLIGHT
                        page.addHighlightAnnot(foundText)
                        self.bm_tag_page.append(page.number)
             # If only not found in any pages
            if self.ifFoundAny == False:
                self.not_found_tags.append(text)
            elif self.ifFoundAny == True:
                self.found_tags.append(text)
                self.bm_tag_dict[text.tagNumber] = self.bm_tag_page
            self.ifFoundAny = False
            self.bm_tag_page = []
        self.CreateDirectory("Output\\ModelNumber")
        if len(self.not_found_tags) > 0:
            self.not_found_tags_dict[pdf] = self.not_found_tags
        if len(self.found_tags) > 0:
            self.found_tags_dict[pdf] = self.found_tags
            doc.save("Output\\" + searchText[0].docOwner +
                     "\\" + pdf, garbage=4, deflate=True, clean=False)
            self.BookMarkTags(
                "Output\\" + searchText[0].docOwner + "\\" + pdf, self.bm_tag_dict, pdf, searchText[0].docOwner)
        doc.close()

    def SearchAndHighlightSpecialCases(self, filePath, pdf, searchText):
        doc = fitz.open(filePath + "\\" + pdf, filetype="pdf")
        for text in searchText:
            for page in doc:
                text_instances = page.searchFor(text.tagNumber)
                if len(text_instances) > 0:
                    self.ifFoundAny = True
                    for foundText in text_instances:   # HIGHLIGHT
                        page.addHighlightAnnot(foundText)
                        self.bm_tag_page.append(page.number)
             # If only not found in any pages
            if self.ifFoundAny == False:
                self.not_found_tags.append(text)
            elif self.ifFoundAny == True:
                self.found_tags.append(text)
                self.bm_tag_dict[text.tagNumber] = self.bm_tag_page
            self.ifFoundAny = False
            self.bm_tag_page = []
        self.CreateDirectory("Output\\ModelNumber")
        if len(self.not_found_tags) > 0:
            self.not_found_tags_dict[pdf] = self.not_found_tags
        if len(self.found_tags) > 0:
            self.found_tags_dict[pdf] = self.found_tags
            doc.save("Output\\ModelNumber" + "\\" + pdf,
                     garbage=4, deflate=True, clean=False)
            self.BookMarkTags("Output\\ModelNumber" + "\\" + pdf,
                              self.bm_tag_dict, pdf, searchText[0].docOwner)
        doc.close()

    def SearchFromBeta(self, tag_no):
        doc = fitz.open(
            "Docs with Issue\\L001-16800-MH-2407-100374-1200_R1.pdf", filetype="pdf")
        for page in doc:
            text_instances = page.searchFor(tag_no)
            if len(text_instances) > 0:
                self.ifFoundAny = True
                for foundText in text_instances:   # HIGHLIGHT
                    page.addHighlightAnnot(foundText)
                    self.bm_tag_page.append(page.number)
        doc.save("Docs with Issue\\Bano.pdf",
                 garbage=4, deflate=True, clean=False)
        doc.close()

    def SearchForTags(self, filePath, pdf_name, tag_no, doc_tag_dict, found_tag_list):
        doc = fitz.open(filePath + "\\" + pdf_name, filetype="pdf")
        for page in doc:
            text_instances = page.searchFor(tag_no)
            if len(text_instances) > 0:
                found_tag_list.append(tag_no)
                break
                # self.ifFoundAny = True
                # for foundText in text_instances:   # HIGHLIGHT
                #     page.addHighlightAnnot(foundText)
                #     self.bm_tag_page.append(page.number)
        #doc.save("Docs with Issue\\Bano.pdf",garbage=4, deflate=True, clean=False)
        doc.close()
