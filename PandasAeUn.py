import pandas as pd
import glob
import openpyxl
from mufy import saveTrack2


def AerialOrUnderground(folder):
    # Set options for displaying the sheet in the pycharm panel
    pd.set_option('display.max_columns', 10)
    pd.set_option('display.max_rows', 2000)
    pd.set_option('display.width', 500)

    # Files of source
    files = glob.glob(f'{folder}/*')

    # destination file
    destinationExcel = openpyxl.load_workbook(saveTrack2)
    destinationSheet = destinationExcel['Montaż kabli']

    z = 3
    # opening excel files
    for excels in files:
        x = -1
        excelFile = pd.read_excel(f"{excels}", sheet_name=None)

        lengthOfSheets = len(list(excelFile.keys()))
        #print(f"Liczba zakładek w pliku -> {length_of_sheets}")
        g = 3
        for sheet in excelFile:
            x += 1
            nameSheet = list(excelFile.keys())[x]
            if nameSheet == "Front Page":
                continue
            # print(name_sheet)
            sheetExcel1 = pd.read_excel(excels, sheet_name=nameSheet)
            entryCableColumn = pd.Series(sheetExcel1['Entry cable'])
            leavingCableColumn = pd.Series(sheetExcel1['Leaving cable'])

            # ENTRY CABLE COLUMN
            # aerial
            aerialEntryCableColumn = entryCableColumn.str.contains(
                'AERIAL', na=False)
            # underground
            undergroundEntryCableColumn = entryCableColumn.str.contains(
                'UNDERGROUND', na=False)

            # LEAVING CABLE COLUMN
            # aerial
            aerialLeavingCableColumn = leavingCableColumn.str.contains(
                'AERIAL', na=False)
            # underground
            undergroundLeavingCableColumn = leavingCableColumn.str.contains(
                'UNDERGROUND', na=False)

            # ONLY AERIAL
            z += 1
            print(z-3)
            if True in (aerialEntryCableColumn.values & aerialLeavingCableColumn.values):
                destinationSheet['E'f"{z}"].value = 'NIE'

                # if FDP dont put value 'Słup'
                if 'FDP' in sheetExcel1.values:
                    destinationSheet['C'f"{z}"].value = ''
                else:
                    destinationSheet['C'f"{z}"].value = 'Słup'

                print("aerial")
            # ONLY UNDERGROUND
            elif True in (undergroundEntryCableColumn.values & undergroundLeavingCableColumn.values):
                destinationSheet['E'f"{z}"].value = 'TAK'
                print('undeground')

            # CHECKING IN COLUMN: LEAVING CABLE IS CABLE:
            elif any(aerialLeavingCableColumn.values) == False & any(undergroundLeavingCableColumn.values) == False:
                if True in aerialEntryCableColumn.values:
                    destinationSheet['E'f"{z}"].value = 'NIE'

                    # if FDP dont put value 'Słup'
                    if 'FDP' in sheetExcel1.values:
                        destinationSheet['C'f"{z}"].value = ''
                    else:
                        destinationSheet['C'f"{z}"].value = 'Słup'

                    print('from entry aerial')
                else:
                    destinationSheet['E'f"{z}"].value = 'NIE'
                    print('from entry undeground')
            else:
                destinationSheet['E'f"{z}"].value = 'CHECK PLEASE'
                print('check please')
    destinationExcel.save(
        saveTrack2)

    return
