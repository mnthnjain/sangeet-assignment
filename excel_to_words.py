import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
import math
import pprint
import subprocess

from docx import Document
from docx.shared import Inches

from tkinter import Tk
from tkinter.filedialog import askopenfilenames

pp = pprint.PrettyPrinter(indent=2)
filesData = []
sectors = {}


print('\033[92m'+"welcome, always remember that manthan is great"+ '\033[0m')

# reading from the excel file
Tk().withdraw() # we don't want a full GUI, so keep the root window from appearing
filenames = askopenfilenames() # show an "Open" dialog box and return the path to the selected file
print("total files selected: ", len(filenames))

# wrinting into word file
document = Document()

# read all the files
listOfFilesData = []
for filename in filenames:
    listOfFilesData.append(pd.read_excel(filename))   

for fileIndex, filename in  enumerate(filenames):
    # df = pd.read_excel(filename)
    df = listOfFilesData[fileIndex]
    filesData.append(df)

    sectors_in_file = []
    for index, sector in enumerate(df["sector"]):
        if type(sector) == str:
            sectors_in_file.append({"start": index})


    # Industrails – stock selection; investment in Reliance/India (1%), Tata/Russia (2%) and Binani/US (3%)
    foundSectorInList = -1
    for index, sector in enumerate(df["sector"]):
        # print("sector----------------",sectors)
        if type(sector) == str:
            foundSectorInList =  foundSectorInList +1
            endIndexForSector =   df.index.stop - 1 if foundSectorInList == sectors_in_file.__len__()-1 else sectors_in_file[foundSectorInList+1]["start"]-1
            sectorInFileObject = {
                    "start": index,
                    "fileIndex": fileIndex,
                    "end": endIndexForSector
                }
            if sector in sectors:
                sectors[sector].append(sectorInFileObject)
            else:
                sectors[sector] = [sectorInFileObject]
            # sectors[sector]=({"start": index, "name": sector, "fileIndex": fileIndex})
    # for index, sector in enumerate(df["sector"]):
    #     if type(sector) == str:
    #         sectors_in_file.append({"start": index})
    
    # for sector in sectors_in_file:
        
    # for sector in sectors:
    #     sectors[sector][fileIndex]["end"]=sectors[index+1]["start"]-1
        
    # sectors[sector][fileIndex][-1]["end"] = df.index.stop - 1


pp.pprint(sectors)


for index, sector in enumerate(sectors):
    # print(sectors[sector])
    sectorName = sector
    sector =  sectors[sector]
    
    for stockObject in sector:
        
        statement =  sectorName + "- stock: " 
        print(stockObject["fileIndex"])
        statement = statement + listOfFilesData[stockObject["fileIndex"]]["impact"][stockObject["start"]] + "; "
        statement = statement+ "investment in "
        stock_country_pair = []
        
        for row in range(stockObject["start"], stockObject["end"]+1):
            stock_country_pair.append(
                listOfFilesData[stockObject["fileIndex"]]["stocks"][row]+"/" 
                    + listOfFilesData[stockObject["fileIndex"]]["country"][row] + "(" 
                        + str(listOfFilesData[stockObject["fileIndex"]]["return"][row]) +"%)")
        
        str1 =  ", ".join(stock_country_pair[0: -1])
        str2 = " and ".join([str1, stock_country_pair[-1]]) if stock_country_pair.__len__()>1 else " " + stock_country_pair[-1]
        p0 = document.add_paragraph(statement + str2, style='List Bullet')


document.add_page_break()
file_folder=filename.split("/")[0:-1]
filename = filename.split("/").pop()[0:-3] + "docx"
file_folder.append("Manthan is great.docx")
filepath = "/".join(file_folder)

document.save(filepath)

print("successfully saved file as",filepath)
# os.system("start " + filepath)

subprocess.check_call(['xdg-open', filepath])


