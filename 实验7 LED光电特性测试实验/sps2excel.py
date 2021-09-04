#sps2excel Converter
#This python3 script can help you convert data
#in sps files into the form of xlsx file,
#which can be manipulated by Microsoft Excel later
#By Jialin Chen    chenjl@shangahitech.edu.cn

import xlsxwriter
#Get the name of the sps file and
#read its data to a list named as 'data'
spsname = input('Please input the name of your sps file (without suffix):\n')
spsfile = open(spsname + '.sps','r')
data = spsfile.read()
spsfile.close()
data = data.split('\n')

#Find and choose the data we need in 'data'
for i in range(len(data)):
    if 'BEGIN' in data[i]:
        begin = i + 2
        break

data = data[begin:]

#Write the data into a xlsx file
#named as 'data.xlsx'
workbook = xlsxwriter.Workbook('data.xlsx')
worksheet = workbook.add_worksheet()

for i in range(len(data)):
    row = data[i].split()
    for j in range(len(row)):
        worksheet.write(i,j,float(row[j]))

workbook.close()
