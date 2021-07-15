import pandas as pd
import glob
import openpyxl

saveTrack = r'C:\Users\marek.barul\Desktop\Python pliki\rozliczenia\testowy\1.xlsx'

def ConnectorType(folder, source):
    # Set options for displaying the sheet in the panel
    pd.set_option('display.max_columns', 5)
    pd.set_option('display.max_rows', 20)
    pd.set_option('display.width', 500)
    z = 3
    # Files of source
    
    files = glob.glob(f'{folder}/*')

    # destination file
    destinationExcel = openpyxl.load_workbook(source)
    destinationSheet = destinationExcel['Montaż kabli']

    # opening excel files
    for excels in files:

        x = -1
        excelFile = pd.read_excel(f"{excels}", sheet_name=None)
        lengthOfSheets = len(list(excelFile.keys()))
        print(f"Liczba zakładek w pliku -> {lengthOfSheets}")

        for sheet1 in excelFile:
            x += 1
            nameSheet = list(excelFile.keys())[x]
            print(nameSheet)
            if nameSheet == "Front Page":
                continue
            # print(name_sheet)
            sheetExcel1 = pd.read_excel(excels, sheet_name=nameSheet)
            # print(sheet_excel1)
            z += 1
            #print(z)

            # Column with Splices
            df = pd.DataFrame(sheetExcel1, columns=['Unnamed: 7'])
            # print(df)

            # checking if sheet have a value "PF"
            checkingPF = df.isin(["PF"]).any()
            # print(checking_PF.astype(int))

            value_of_bool_PF = (checkingPF.astype(int) == 1).bool()

            if value_of_bool_PF == 1:
                # print("MID")

                destinationSheet['F'f"{z}"].value = 'MID'
            else:
                # print("CUT")
                destinationSheet['F'f"{z}"].value = "CUT"
    # save file
    destinationExcel.save(saveTrack)
    return
