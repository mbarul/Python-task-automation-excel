import glob
import openpyxl
import os.path
import pandas as pd
from pandasPf import saveTrack

saveTrack2 = saveTrack
def MufsNumbers(folder):
    y = -1
    x = 3
    # Open destination file
    excelEmpty = openpyxl.load_workbook(saveTrack)
    sheetEmpty = excelEmpty['Montaż kabli']

    #1FCP number
    names = [os.path.basename(x) for x in glob.glob(f'{folder}/*')]

    #Files of source
    files = glob.glob(f'{folder}/*')
    #print(files)
    lengthOfFiles = len(names)
    print(lengthOfFiles)
    print(names)
    # Open all excel files
    for excel in files:
        excelFile = openpyxl.load_workbook( f'{excel}', data_only=True)
        y += 1
        if y > lengthOfFiles :
           break
        print(y)

        #All sheets
        workSheetsNames = excelFile.get_sheet_names()
        #Open sheets a take values
        for sheets in workSheetsNames:

            if sheets == "Front Page":
                continue

            print(f"{sheets} <- otwarte zakładki z pliku "f"{names[y][:9]} \n")
            x += 1
            # enter names of sheets to cells of destination file
            # 1FCP number
            sheetEmpty['A'f"{x}"].value = names[y][:9]
            # muffs number
            sheetEmpty['B'f"{x}"].value = sheets
            
    # save file
    excelEmpty.save(saveTrack2)
    return





