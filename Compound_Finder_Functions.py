import csv
import xlwt
import datetime
from operator import itemgetter

libe = "RetentionTimeLibrary.csv"
inputfile = "test.txt"


def import_gcms(gcms_file):

    with open(gcms_file) as original_file:
        skip_line = 4
        peak_list = []
        for row in original_file:

            if skip_line > 0 or len(row.split()) == 0:
                skip_line -= 1
            else:
                peak_list.append(row.split())

    return peak_list


def import_library(library):

    compound_list = []

    with open(library) as csvfile:
        reader = csv.reader(csvfile, delimiter=",")
        skip_row = 1

        for row in reader:
            temp_row = []

            if skip_row > 0:
                skip_row -= 1
            elif len(row[0]) == 0 or len(row[1]) == 0:
                pass
            else:
                for column in range(3):
                    temp_row.append(row[column].replace("*", ""))

                compound_list.append(temp_row)

    return compound_list


def ws1_list(peak_list):

        worksheet_list = []

        for i in peak_list:

            temp_list = list()
            temp_list.append(int(i[0]))                     # 0 Peak Number
            temp_list.append("")                            # 1 Guess 1
            temp_list.append("")                            # 2 Area Percentage
            temp_list.append(round(float(i[1]), 3))         # 3 Retention Time
            temp_list.append(int(i[4]))                     # 4 Area
            temp_list.append("")                            # 5 Guess 2
            temp_list.append("")                            # 6 Guess 3
            temp_list.append("")                            # 7 Guess 4

            worksheet_list.append(temp_list)

        return worksheet_list


def ws2_list(peak_list):
    worksheet2_list = []

    for i in peak_list:
        temp_list = list()
        temp_list.append(int(i[0]))              # 0 Peak Number
        temp_list.append(i[2])                   # 1 Type
        temp_list.append(round(float(i[3]), 3))  # 2 Width
        temp_list.append(round(float(i[5]), 3))  # 3 Start Time
        temp_list.append(round(float(i[6]), 3))  # 4 End Time

        worksheet2_list.append(temp_list)

    return worksheet2_list


def compound_percentage(worksheet1):

    total = 0

    for i in worksheet1:
        total += int(i[4])

    for i in worksheet1:
        i[2] = str(round(((int(i[4])/total)*100), 2)) + "%"

    return worksheet1


def guess_builder(compound_list, ws1):

    epsilon = 0.1
    for i in ws1:

        unorganized_list = []
        blank_guess = []
        best_guess = 10.0
        ret_time = i[3]
        for compound in compound_list:

            ret_diff = round(float(ret_time) - float(compound[1]), 3)
            if abs(ret_diff) <= epsilon:
                unorganized_list.append(["{}({})".format(compound[0], str(ret_diff)), abs(ret_diff)])
            elif abs(ret_diff) < best_guess:
                best_guess = abs(ret_diff)
                blank_guess = ["(??){}({})".format(compound[0], str(ret_diff)), abs(ret_diff)]

        if len(unorganized_list) == 0:
            unorganized_list.append(blank_guess)

        organized_list = sorted(unorganized_list, key=itemgetter(1))

        try:
            i[1] = organized_list[0][0]
            i[5] = organized_list[1][0]
            i[6] = organized_list[2][0]
            i[7] = organized_list[3][0]
        except IndexError:
            pass

    return ws1


def final_file_creator(worksheet1, worksheet2, generate, cofa, output_name):


    wb = xlwt.Workbook()
    ws = wb.add_sheet("Compound info")
    ws2 = wb.add_sheet("Additional info")

    # --------------------------------------------------------------
    # Generates Worksheet1 - Main information
    # --------------------------------------------------------------

    Titles = ["Peak", "Guess 1", "Percentage", "Ret Time", "Area", "Guess 2", "Guess 3", "Guess 4"]
    column = 0
    for i in Titles:
        ws.write(0, column, i, xlwt.easyxf("align: horiz center; font: bold on; borders: bottom thin"))
        column += 1

    columns = len(worksheet1[0])
    rows = len(worksheet1)
    for i in range(columns):
        for j in range(rows):
            x = (worksheet1[j][i])
            ws.write(j+2, i, x)

    ws.col(1).width = 256 * 30
    ws.col(2).width = 256 * 13
    ws.col(4).width = 256 * 13
    ws.col(5).width = 256 * 30
    ws.col(6).width = 256 * 30
    ws.col(7).width = 256 * 30

    # --------------------------------------------------------------
    # Generates Worksheet2 - Additional Info
    # --------------------------------------------------------------

    Titles = ["Peak", "Type", "Width", "Start Time", "End Time"]
    column = 0
    for i in Titles:
        ws2.write(0, column, i, xlwt.easyxf("align: horiz center; font: bold on; borders: bottom thin"))
        column += 1

    columns = len(worksheet2[0])
    rows = len(worksheet2)
    for i in range(columns):
        for j in range(rows):
            x = (worksheet2[j][i])
            ws2.write(j+2, i, x)

    # --------------------------------------------------------------
    # Leaves function if not told to generate CofA part
    # --------------------------------------------------------------

    if generate == False:
        wb.save(output_name)
        return

    # --------------------------------------------------------------
    # Certificate of analysis portion
    # --------------------------------------------------------------

    ws3 = wb.add_sheet("Certificate of Analysis")

    style_main_string = "font: name Calibri; align: horiz left; font: height 200; borders: bottom thin; borders: top thin; borders: left thin; borders: right thin"
    style_top_string = "font: name Calibri; align: horiz left; font: height 320; font: bold on"
    font0 = xlwt.easyfont("name Calibri, height 220")
    font1 = xlwt.easyfont("bold true, name Calibri, height 220")
    info_style1 = xlwt.easyxf("align: horiz left; font: name Calibri")

    analysis1 = (cofa[1][0][:15], font1)                             # For Analysis formatting in CofA
    analysis2 = (str(datetime.date.today()), font0)
    ws3.write_rich_text(1, 0, (analysis1, analysis2), info_style1)

    batch1 = (cofa[2][0][:8], font1)                                 # For Batch formatting in CofA
    batch2 = (cofa[2][0][8:], font0)
    ws3.write_rich_text(2, 0, (batch1, batch2), info_style1)

    lot1 = (cofa[3][0][:6], font1)                                   # For Lot formatting in CofA
    lot2 = (cofa[3][0][6:], font0)
    ws3.write_rich_text(3, 0, (lot1, lot2), info_style1)

    ws3.write(0, 1, cofa[0][1], xlwt.easyxf(style_top_string))
    ws3.write(1, 1, cofa[1][1], xlwt.easyxf("font: name Calibri; align: horiz left; font: height 200; font: italic on"))

    for i in range(0, 2):
        for j in range(5, 10):
            x = (cofa[j][i])
            ws3.write(j, i, x, xlwt.easyxf(style_main_string, num_format_str="#,##0.00"))

    for i in range(0, 3):
        for j in range(11, 15):
            x = (cofa[j][i])
            ws3.write(j, i, x, xlwt.easyxf(style_main_string, num_format_str="#,##0.00"))

    for i in range(0, 3):
        for j in range(16, 21):
            x = (cofa[j][i])
            ws3.write(j, i, x, xlwt.easyxf(style_main_string, num_format_str="#,##0.00"))

    ws3.col(0).width = 256 * 36
    ws3.col(1).width = 256 * 30
    ws3.col(2).width = 256 * 39

    wb.save(output_name)

    return


def main(generate_cofa, inputfile, library, output_name, cofa):

    peak_list = import_gcms(inputfile)
    compound_list = import_library(library)
    ws1 = ws1_list(peak_list)
    ws2 = ws2_list(peak_list)
    compound_percentage(ws1)
    guess_builder(compound_list, ws1)
    final_file_creator(ws1, ws2, generate_cofa, cofa, output_name)
















