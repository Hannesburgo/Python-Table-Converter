# This class gets all the formats and sizes of the inputted table.
class Formats:
    def __init__(self, dictionary:dict):
        self.formats = dict()
        self.dictionary = dictionary

    def appendFormat(self, stoneFormat):
        try:
            extenso = self.dictionary["Formatos"][stoneFormat]
            if extenso not in self.formats.keys():
                self.formats[extenso] = list()
        except:
            print("[WARNING] Format '"+stoneFormat+"' not recognized in Dictionary - Passing it")
            pass

    def retrieveFormat(self, formatID):
        try:
            return self.formats[self.dictionary[formatID]]
        except:
            print("[ERROR] Format Unknow - Check if this format was writed correctly or if it exists in the format list")

    def deleteFormat(self, formatID):
        del self.formats[formatID]

    def appendInfo(self, stoneID, stoneType, stoneFormat, stoneSize, stoneExtraOne, stoneExtraTwo, isFirstClass): 
        try:
            extenso = self.dictionary["Formatos"][stoneFormat]
            lapidation = None
            # Check if it is zircon first class, if not, Pass
            if isFirstClass:
                if stoneExtraOne == "/I" or stoneType == "ZP":
                    ""
                else:
                    return
            # Check if the extra two is in the extra one because the lapidation is FC
            if stoneExtraOne in self.dictionary["Extras"].keys():
                stoneExtraTwo = stoneExtraOne
                stoneExtraOne = "FC"
            # Check if the extra two and extra one are misplaced between themselves
            if stoneExtraOne in self.dictionary["Extras"].keys() and stoneExtraTwo in self.dictionary["Lapidações"].keys():
                x = stoneExtraOne
                stoneExtraOne = stoneExtraTwo
                stoneExtraTwo = x
            # Check if there is an extra two
            if stoneExtraTwo:
                lapidation = self.dictionary["Lapidações"][stoneExtraOne] + " " + self.dictionary["Extras"][stoneExtraTwo]
            else:
                lapidation = self.dictionary["Lapidações"][stoneExtraOne]
            self.formats[extenso].append(
                [stoneID,
                self.dictionary["Pedras"][stoneType],
                stoneSize,
                lapidation]
            )
        except:
            print("[WARNING] The ID "+str(stoneID)+" has informations that were not recognized in the dictionary - Passing it.")
            print("\nID: "+str(stoneID)+" StoneType: "+stoneType+" stoneSize: "+stoneSize+" ExtraOne: "+stoneExtraOne+" ExtraTwo: "+stoneExtraTwo)
            pass

    def getFormats(self):
        return self.formats