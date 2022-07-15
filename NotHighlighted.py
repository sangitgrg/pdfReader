class NotHighlightedTags:
    def __init__(self,foundDict, notFoundDict):
        self.foundDict = foundDict
        self.notFoundDict = notFoundDict

class FoundTagOwner:
    def __init__(self, tagNumber, docOwner, docRevision, docTitle, docType):
        self.tagNumber = tagNumber
        self.docOwner = docOwner
        self.docRevision = docRevision
        self.docTitle = docTitle
        self.docType = docType