import xlrd
import xlwt

inputfile = "Rumplestilskin.xls"

def CofA_format_builder():

    #Buidls a blank list to eventually be turned in to the C of A format
    #Out puts a blank 2d list to be build

    list_builder = []

    for i in range(21):
        list_builder.append(["", "", ""])

    return list_builder


def CofA_Static_additions(list_builder):

    #Adds all the static labels into list builder to be put into the C of A

    temp_list = list_builder

    temp_list[1][0] = "Analysis Date"
    temp_list[5][0] = "Plant Part"
    temp_list[6][0] = "Cultivation Method"
    temp_list[7][0] = "Extraction Method"
    temp_list[8][0] = "Country of Origin"
    temp_list[9][0] = "Quality"
    temp_list[9][1] = "100% Pure and Natural"
    temp_list[11][0] = "Color"
    temp_list[12][0] = "Odor"
    temp_list[13][0] = "Consistency"
    temp_list[14][0] = "Drops/mL"
    temp_list[16][0] = "Specific Gravity @ 20C"
    temp_list[17][0] = "Refractive Index @ 20C"
    temp_list[18][0] = "Optical Rotation @ 20C"
    temp_list[19][0] = "Viscosity @ 20C"
    temp_list[20][0] = "Solubility"
    temp_list[20][1] = "Soluble in Alcohol and fixed oils"
    temp_list[11][2] = "Conforms"
    temp_list[12][2] = "Conforms"
    temp_list[13][2] = "Conforms"
    temp_list[14][2] = "Conforms"
    temp_list[16][2] = "Conforms"
    temp_list[17][2] = "Conforms"
    temp_list[18][2] = "Conforms"
    temp_list[19][2] = "Conforms"
    temp_list[20][2] = "Conforms"

    return temp_list

def CofA_variable_additions(list_builder, lot):

    book = xlrd.open_workbook(inputfile)
    sh = book.sheet_by_index(0)
    temp_list = list_builder


    for i in range(sh.nrows):


        if sh.cell_value(rowx=i, colx=0) == "END":
            return False

        if sh.cell_value(rowx=i, colx=1).lower() == lot.lower():
            found_row = i
            break

    temp_list[0][1] = sh.cell_value(rowx=found_row, colx=0) #A
    temp_list[1][1] = sh.cell_value(rowx=found_row, colx=4) #E
    temp_list[2][0] = sh.cell_value(rowx=found_row, colx=2) #C
    temp_list[3][0] = sh.cell_value(rowx=found_row, colx=1) #B
    temp_list[5][1] = sh.cell_value(rowx=found_row, colx=6) #G
    temp_list[6][1] = sh.cell_value(rowx=found_row, colx=8) #I
    temp_list[7][1] = sh.cell_value(rowx=found_row, colx=7) #H
    temp_list[8][1] = sh.cell_value(rowx=found_row, colx=5) #F
    temp_list[11][1] = sh.cell_value(rowx=found_row, colx=21) #V
    temp_list[12][1] = sh.cell_value(rowx=found_row, colx=22) #W
    temp_list[13][1] = sh.cell_value(rowx=found_row, colx=23) #X
    temp_list[14][1] = sh.cell_value(rowx=found_row, colx=24) #Y
    temp_list[16][1] = sh.cell_value(rowx=found_row, colx=19) #T
    temp_list[17][1] = sh.cell_value(rowx=found_row, colx=17) #R
    temp_list[18][1] = sh.cell_value(rowx=found_row, colx=18) #S
    temp_list[19][1] = sh.cell_value(rowx=found_row, colx=16) #Q

    return temp_list