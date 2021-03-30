import os
import urllib.request
import zipfile
import pathlib
import glob
import time
from openpyxl import load_workbook

#This var enables verbose
debug = 0

print("Sto scaricando il file!\n")

#call to TIM servers and download zip file
downloadUrl = 'https://www.wholesale.telecomitalia.com/sitepub/SFTP/59_Coperture_Bitstream_NGA_e_VULA/Copertura%20attiva%20e%20pianificata%20FTTCab.zip '
urllib.request.urlretrieve(downloadUrl, 'C:\Desktop\TIM\\2021\down.zip')

if debug == 1: print("File scaricato! Inizio estrazione!\n")

#unzip downloaded file operation
with zipfile.ZipFile('C:\Desktop\TIM\\2021\down.zip', 'r') as zip_ref:
    zip_ref.extractall('C:\Desktop\TIM\\2021')
    time.sleep(2)

if debug == 1: print("File estratto! Cancello lo zip\n")

#delete unusefull zip file
if os.path.exists('C:\Desktop\TIM\\2021\down.zip'):
    os.remove('C:\Desktop\TIM\\2021\\down.zip')
    if debug == 1: print("File cancellato\n")
else:
    if debug == 1: print("Il file non esiste\n")

if debug == 1: print("Apertura excel in corso!\n")

#begin to open file and search
#create a list of all *xlsx files in the folder
listOfFiles = glob.glob('C:\Desktop\TIM\\2021\*.xlsx')
#take the last modified file (the one just downloaded)
lastFile = max(listOfFiles, key=os.path.getctime)
#extract the path
filePath = pathlib.Path(lastFile)
strPath = str(filePath)
print("Ho il percorso!\n" + strPath + "\nCaricamento File in corso! Attendere prego!\n")
file = load_workbook(strPath)
print("Caricamento foglio Excel in corso\n")
#here I select the excel sheet
fttc = file['FTTC']
print("File caricato, inizio ricerca!")

#this function search for the specific cabinet ALADITAB008 for example
for x in range(150000):
    if x > 0:
        #I search for cab name
        if fttc.cell(row=x, column=5).value == "ALADITAB008":
            if fttc.cell(row=x, column=9).value == "Pianificato":
                if fttc.cell(row=x, column=8).value == "100M":
                    print("Il cabinet NON è attivo. Data attivazione prevista: "+str(fttc.cell(row=x, column=11).value) + "\n")
                    print("Velocità di attivazione prevista: " + str(fttc.cell(row=x, column=8).value))
                    break
                elif fttc.cell(row=x, column=8).value == "Upgrade 200M":
                    print("Programmazione upgrade cabinet! Data prevista upgrade 200M: " + str(fttc.cell(row=x, column=11).value) + "\n")
                    break
            elif fttc.cell(row=x, column=9).value == "Attivo":
                print("Il cabinet è ATTIVO! Data attivazione: " + str(fttc.cell(row=x, column=10).value) + "\nContattare Operatore!\n")
                break
            elif fttc.cell(row=x, column=9).value == "Sospeso":
                print("Il cabinet è SOSPESO, attendi\n")
                break
            elif fttc.cell(row=x, column=9).value == "Saturo":
                print("Cabinet SATURO, attendere desaturazione, probabile upgrade 200M\n")
                break
            time.sleep(10)
        else:
            x = x + 1
    #this fix is necessary for a correct row index
    else:
        x = 1

print("Fine del programma, alla prossima!")
