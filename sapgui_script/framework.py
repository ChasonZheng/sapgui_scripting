from win32com import client
import csv

class Runnable:
    def run(self, ses, code, row):
        return

class Transaction:
    
    def __init__(self, tcode, pathToData):
        self.tcode = tcode
        self.path = pathToData
        self.sapGui = client.GetObject("SAPGUI").GetScriptingEngine
        self.ses = self.sapGui.FindById("ses[0]")
        
    def runScript(self, runnable):                
        with open(self.path) as csvfile:                    
            reader = csv.DictReader(csvfile, delimiter = ';')
            for row in reader:
                try:
                    runnable.run(self.ses, self.tcode, row)
                except:
                    print(self.ses.FindById("wnd[0]/sbar").text)
        
    
