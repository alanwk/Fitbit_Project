"""
Self-made module for functions used in the Fitbit Project.
"""
import numpy as np

def ObjtoInt(dfname, colname):
    """
    For a column in a dataFrame, change any string values that have commas into integers after removing the comma
    :param dfname: name of the dataFrame being accessed
    :param colname: name of the column in the specified dataFrame being accessed
    :return: string statement explaining the success/failure of the function
    """
    try:
        dfname[colname] = dfname[colname].str.replace(",", "")
        dfname[colname] = dfname[colname].astype('int64')
        return "Objects Changed to Integers Complete"
    except:
        return "Objects Were Not Able to be Changed to Integers for {}".format(colname)

def xFoodNonEntries(df, minThreshold, colname):
    """
    For any values in a specified column lower than the minimum threshold, replace with NaN values.
    :param df: dataFrame being accessed
    :param minThreshold: minimum threshold being used to compare against values. All values lesser than minThreshold will be replaced with NaN
    :param colname: string providing the name of the column being accessed
    :return: cleaned list with NaN values instead of zeroes
    """
    cleanlist = []
    for val in df.T.loc[colname]:
        if int(val) > minThreshold:
            cleanlist.append(val)
        else:
            cleanlist.append(np.NaN)
    return cleanlist

"""
Coding Note: This first function is a target for future revision. The function was made early on in the process and is in need of an update. I am hoping to remove the necessity of the openpyxl module in the near future by only using the csv file provided by Fitbit. Unfortunately, I have been restricted by time and have yet to work on this update. 
"""
def foodClean(filename):
    """
    Locate nutrient and calorie information from food section of exported Fibit file and replace in a transposed position with new headings
    
    :param filename: filename of the excel file in need of food formatting
    :return: String confirming that function produced a new file
    """
    #Excel Workbook Access
    wb = load_workbook(filename)
    ws = wb.active

    #List and Variable Creation
    foodLog = []
    counter = 0
    cellRanges = [] #for testing
    nR = []

    #Search for unorganized food cell values
    for cell in ws['A']:
        try:
            if cell.value == "Daily Totals":
                foodLog.append(cell.coordinate)     #Used to locate nutrient amounts section
            else:
                continue
        except:
            pass

    #Search for desired location of food cell values
    for cell in ws['B']:
        try:
            if cell.value == "Calories In":
                nR.append(cell.coordinate)     #Used to locate start of calories section
            else:
                continue
        except:
            pass

    # Create headers for new columns
    hdRow = nR[0].replace("B", "")
    ws[("C" + hdRow)].value = "Fat (g)"
    ws[("D" + hdRow)].value = "Fiber (g)"
    ws[("E" + hdRow)].value = "Carbs (g)"
    ws[("F" + hdRow)].value = "Sodium (mg)"
    ws[("G" + hdRow)].value = "Protein (g)"

    # Take "Daily Totals" Location as reference, changing to cells of interest, and pasting values in new location
    for entry in foodLog:
        letr = int(entry.replace("A", ""))      #Removing letter from old coordinates to alter position
        nS = int(nR[0].replace("B", "")) + 1    #Determining new starting location
        # cellRanges.append("C" + str(letr + 2) + ":C" + str(letr + 6))  #for testing
        ws[("C" + str(nS + counter))].value = int(ws[("C" + str(letr + 2))].value.replace(" g","").replace(",",""))
        ws[("D" + str(nS + counter))].value = int(ws[("C" + str(letr + 3))].value.replace(" g","").replace(",",""))
        ws[("E" + str(nS + counter))].value = int(ws[("C" + str(letr + 4))].value.replace(" g","").replace(",",""))
        ws[("F" + str(nS + counter))].value = int(ws[("C" + str(letr + 5))].value.replace(" mg","").replace(",",""))
        ws[("G" + str(nS + counter))].value = int(ws[("C" + str(letr + 6))].value.replace(" g","").replace(",",""))
        counter += 1

    # print(cellRanges)
    # print(len(cellRanges))
    wb.save('cleaned_' + filename)
    goodJob = "Cleaned File Created"
    return goodJob

    