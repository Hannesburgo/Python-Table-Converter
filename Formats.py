# This class gets all the formats and sizes of the inputted table.
class Formats:
    def __init__(self, dictionary:dict):
        self.formats = dict()
        self.dictionary = dictionary

    def appendFormat(self, formatID):
        if self.dictionary[formatID] not in self.formats:
            self.formats[self.dictionary[formatID]] = list()

    def retrieveFormat(self, formatID):
        try:
            return self.formats[self.dictionary[formatID]]
        except:
            print("[ERROR] Format Unknow - Check if this format was writed correctly or if it exists in the format list")

    def deleteFormat(self, formatID):
        del self.formats[formatID]

    def appendInfo(self, formatID, id, size, color):
        self.formats[self.dictionary[formatID]].append([id, size, self.dictionary[color]])

    def getFormats(self):
        return self.formats