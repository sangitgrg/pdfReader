import os
import shutil
from datetime import datetime

class DirectoryOperations:
    def __init__(self,docPath):
        self.docPath = docPath

    def ScanFolders(self, folderPath):
        actualFolderPath = self.docPath + "\\"+ folderPath
        # actualFolderPath = self.docPath
        dirs = [d for d in os.listdir(actualFolderPath) if os.path.isdir(os.path.join(actualFolderPath, d))]
        return dirs

    def CreateOutputDirectory(self):
            if not os.path.exists("Output"):
                os.mkdir("Output")
            else:
                shutil.rmtree("Output")
                os.mkdir("Output")    

    def CheckforLatestFolder(self):
        # folderName = datetime.now().strftime("%Y%m%d") 
        folderName = "Request"
        return folderName
        for x,subdirs,z in os.walk(self.docPath):
            if folderName in subdirs:
                return folderName
