from win32com import client
import csv
from datetime import datetime

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
        log = open("log_"+self.tcode+"_"+str(datetime.now()).replace(" ","_").replace(":","_")+".txt", "w")
        with open(self.path) as csvfile:                    
            reader = csv.DictReader(csvfile, delimiter = ';')
            for row in reader:
                try:
                    runnable.run(self.ses, self.tcode, row)
                except:
                    print(self.ses.FindById("wnd[0]/sbar").text)
                    log.write(self.ses.FindById("wnd[0]/sbar").text+"\n")
        log.close()
             
    
