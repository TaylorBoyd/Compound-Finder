import csv
import xlwt
from CofA_Functions import *

epsilon = 0.1
libe = "RetentionTimeLibrary.csv"
inputfile = "test.txt"

def file_converter(inputfile):

    # Converts txt file from the GCMS directly into a .csv file to be used by the program.
    # This is done through a temp.csv file which is then later turned in to the file form.
    # Temp file was used to avoid read/write errors when working on the same file.

    #inputfile = (input("Please type the file name you want to use: "))+".txt"
    skip_line = 4
    new_rows = []

    with open(inputfile) as original_file, open("temp.csv", 'w', newline='') as output_file:
        writer = csv.writer(output_file, lineterminator="\n")
        for row in original_file:

            if skip_line > 0 or row == ["[]"]:
                skip_line -= 1
            else:
                if len(row.split()) > 0:
                    new_rows = row.split()
                    writer.writerow(new_rows)
                else:
                    pass

        print("File Successfully Converted!")
    return

def import_library(library):

    # Imports the library from RetentionTimeLibrary.csv
    # Library should only be filled with strings and floats
    # Converts this file to a dict for easy use [float -> str]

    compound_list = {}

    with open(library) as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            if row['Retention Time'] != '':
                fixed = row['Retention Time'].replace('*', "")
                compound_list[float(fixed)] = row['Compound Name']

        print(len(compound_list), "Compounds Loaded!")

    return compound_list\

def worksheet1_list():

    #takes the temp CSV and converts it to a first pass formatting in a list of lists to be used later
    #Rest of the data is placed on worksheet2 list

    with open("temp.csv") as tempfile:
        reader = csv.reader(tempfile, delimiter=",")

        ws1_list = []
        for row in reader:
            current_row = ["", "", "", "", "", "", "", ""]
            list_column = len(row)

            current_row[0] = row[0]
            current_row[3] = round(float(row[1]), 3)
            current_row[4] = int(row[4])

            ws1_list.append(current_row)

        return ws1_list

def compound_percentage(worksheet_list):

    #Outputs a list of the compound percentages to be used later.
    #[[list]] -> [[list]]
    total = 0
    temp_list = worksheet_list

    for i in temp_list:
        total += int(i[4])


    for i in temp_list:
        i[2] = str(round(((int(i[4])/total)*100), 2)) + "%"

    return temp_list

def guess_builder(compounds, worksheet_list):

    #Builds a list of compound guesses in ascending order of distance from retention library
    #{dict} -> [list]

    temp_list = worksheet_list

    for i in temp_list:

        unorganized_rows = {}

        for key in compounds:
            x = (key) - float(i[3])

            if abs(x) <= epsilon:
                unorganized_rows[abs(round(x, 3))] = (compounds[key] + "(" + str(round(x, 3)) + ")")

        organized_row = row_organizer(unorganized_rows)

        max_guess = 4
        guess_index = len(organized_row)

        if guess_index > 0:

            for j in range(guess_index):
                if max_guess == 0:
                    return

                elif j == 0:
                    i[1] = organized_row[j]
                    max_guess -= 1
                else:
                    i[j+4] = organized_row[j]
                    max_guess -= 1
        else:
            i[1] = blank_guesser(compounds, float(i[3]))


    return temp_list


def row_organizer(unorg_row):

    #[dict -> string]
    #Organizes in descending order of difference from Retention Time library values

    organized_row = []
    sorted_keys = []

    for key in unorg_row:
        sorted_keys.append(key)

    sorted_keys.sort()
    for i in sorted_keys:
        organized_row.append(unorg_row[i])

    return organized_row

def blank_guesser(compounds, retention_number):

    #Fills in blanks by picking the closest guess
    #Appends (??) to the front to denote a large guess

    best_guess = ""
    lowest = 2.0

    for key in compounds:
        x = (key) - float(retention_number)

        if abs(x) <= lowest:
            lowest = abs(x)
            best_guess = ("(??)" + compounds[key] + "(" + str(round(x, 3)) + ")")

    return best_guess

def worksheet2_list():

    #takes the temp CSV and converts it to a first pass formatting in a list of lists to be used later
    #Rest of the data is placed on worksheet2 list

    with open("temp.csv") as tempfile:
        reader = csv.reader(tempfile, delimiter=",")
        ws2_list = [["Peak", "Type", "Width", "Start Time", "End Time"]]
        for row in reader:
            temp_list = ["", "", "", "", ""]
            temp_list[0] = row[0]
            temp_list[1] = row[2]
            temp_list[2] = round(float(row[3]), 3)
            temp_list[3] = round(float(row[5]), 3)
            temp_list[4] = round(float(row[6]), 3)
            ws2_list.append(temp_list)

        return ws2_list

def Final_File_Creator(compounds, outputName, generate, lot):

    worksheet1 = worksheet1_list()
    worksheet2 = worksheet2_list()

    worksheet1 = compound_percentage(worksheet1)
    worksheet1 = guess_builder(compounds, worksheet1)

    wb = xlwt.Workbook()
    ws = wb.add_sheet("Compound info")
    ws2 = wb.add_sheet("Additional info")

    Titles = ["Peak", "Guess 1", "Percentage", "Ret Time", "Area", "Guess 2", "Guess 3", "Guess 4"]
    column = 0
    for i in Titles:
        ws.write(0, column, i, xlwt.easyxf("align: horiz center"))
        column += 1

    columns = len(worksheet1[0])
    rows = len(worksheet1)
    for i in range(columns):
        for j in range(rows):
            x = (worksheet1[j][i])
            ws.write(j+1, i, x)


    columns = len(worksheet2[0])
    rows = len(worksheet2)
    for i in range(columns):
        for j in range(rows):
            x = (worksheet2[j][i])
            ws2.write(j, i, x)

    ws.col(1).width = 256*30
    ws.col(2).width = 256*13
    ws.col(4).width = 256 * 13
    ws.col(5).width = 256*30
    ws.col(6).width = 256*30
    ws.col(7).width = 256*30

    # --------------------------------------------------------------
    # Leaves function if not told to generate CofA part
    # --------------------------------------------------------------

    if generate == False:
        wb.save(outputName)
        return

    # --------------------------------------------------------------
    # Certificate of analysis portion
    # --------------------------------------------------------------

    ws3 = wb.add_sheet("Certificate of Analysis")
    cofa = CofA_format_builder()
    cofa = CofA_Static_additions(cofa)
    cofa = CofA_variable_additions(cofa, lot)

    if cofa == False:
        return False


    columns = len(cofa[0])
    rows = len(cofa)
    for i in range(columns):
        for j in range(rows):
            x = (cofa[j][i])
            ws3.write(j, i, x, xlwt.easyxf("align: horiz right", num_format_str="#,##0.00"))

    ws3.col(0).width = 256 * 21
    ws3.col(1).width = 256 * 30
    ws3.col(2).width = 256 * 9

    wb.save(outputName)

    return



#compound_list = import_library(libe)
#file_converter(inputfile)
#Final_File_Creator(compound_list, "Testeroni2.xls")

















