#This was the beginning of the reddit bot, it started as a little program which outputted a random pokemon along with their stats.

import openpyxl
import random

Stats = ('Name:',"HP", "ATK", 'DEF', 'SATK', 'SDEF', 'SPD', 'TOTAL')
wb = openpyxl.load_workbook('Pokemon.xlsx')
sheet = wb.get_sheet_by_name('Sheet1')

#generates a random number between the first and last rows in the excel file to be used in later parts of the program
R = (random.randint(2,719))
r = str(R)

#combines the number with "A" collum + the pokemon names
Pokemon_Name = 'A'+r
C = sheet[Pokemon_Name]
B = C.value


print('Dex Number '+ str(R-1))

#Prints random Pokemon + Base Stats
for x in sheet['A'+r : "H"+r]:
    n = 0
    for stat_value in x:
        print(Stats[n], stat_value.value)
        n+=1
