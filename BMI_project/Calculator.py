import numpy as np
import xlrd
import xlwt
import matplotlib.pyplot as plt

## read the raw datas
workbook=xlrd.open_workbook(r'database_3.xlsx')
sheet_BMI=workbook.sheet_by_name('BMI')
sheet_2=workbook.sheet_by_name('2')

## Excel for the result
Result = xlwt.Workbook()
worksheet = Result.add_sheet('Result')

## BMI part
worksheet.write(0, 0,'BMI')

for row in range(1,sheet_BMI.nrows):
    BMI=sheet_BMI.cell_value(row,1)/(sheet_BMI.cell_value(row,0)/100)**2
    worksheet.write(row,0,BMI)

## pyhsical attribute part
worksheet.write(0,1,'pyhthical attribute')
for row in range(1,sheet_2.nrows):
    PA= -110 + 1.34*sheet_2.cell_value(row,0) + 1.54*sheet_2.cell_value(row,1) + 1.20*sheet_2.cell_value(row,2) \
       + 1.11*sheet_2.cell_value(row,3) + 1.15*sheet_2.cell_value(row,4) + 0.177*sheet_2.cell_value(row,5)
    worksheet.write(row,1,PA)

Result.save('Result.xls')

## linear regression
## Age and BMI
## calculate the slope and intercept
workbook_r=xlrd.open_workbook(r'Result.xls')
sheet_BMI_R=workbook_r.sheet_by_name('Result')
sumXY=0
sumX=0
sumY=0
sumXSquaried=0
N=sheet_BMI.nrows
for row in range(1,N):
    sumXY=sumXY + sheet_BMI.cell_value(row,3)*sheet_BMI_R.cell_value(row,0)
    sumX=sumX + sheet_BMI.cell_value(row,3)
    sumY=sumY + sheet_BMI_R.cell_value(row,0)
    sumXSquaried=sumXSquaried + sheet_BMI.cell_value(row,3)**2
slope=(N*sumXY-sumX*sumY)/(N*sumXSquaried-(sumX)**2)
intercept=(sumY-slope*sumX)/N

## plot the scatter map and linear regression line
fig1 = plt.figure(num='Age_BMI', figsize=(10, 5), dpi=75, facecolor='#FFFFFF', edgecolor='#0000FF')
x= sheet_BMI.col_values(3, start_rowx=1, end_rowx=None)
y= sheet_BMI_R.col_values(0, start_rowx=1, end_rowx=None)
p1=plt.scatter(x,y,marker='x',color='b',label='Age and BMI',s=30)
x_line = np.linspace(18,67,3)
y_line = slope*x_line+intercept
plt.plot(x_line, y_line, '-r', label='Age and BMI linear regression')

## weight and physical attributes
## calculate the slope and intercept
sumXY=0
sumX=0
sumY=0
sumXSquaried=0
N=sheet_BMI.nrows
for row in range(1,N):
    sumXY=sumXY + sheet_BMI.cell_value(row,1)*sheet_BMI_R.cell_value(row,1)
    sumX=sumX + sheet_BMI.cell_value(row,1)
    sumY=sumY + sheet_BMI_R.cell_value(row,1)
    sumXSquaried=sumXSquaried + sheet_BMI.cell_value(row,1)**2
slope=(N*sumXY-sumX*sumY)/(N*sumXSquaried-(sumX)**2)
intercept=(sumY-slope*sumX)/N

## plot the scatter map and linear regression line
fig2 = plt.figure(num='Weight_PA', figsize=(10, 5), dpi=75, facecolor='#FFFFFF', edgecolor='#FF0000')
x= sheet_BMI.col_values(1, start_rowx=1, end_rowx=None)
y= sheet_BMI_R.col_values(1, start_rowx=1, end_rowx=None)
p1=plt.scatter(x,y,marker='x',color='b',label='Age and BMI',s=30)
x_line = np.linspace(42,117,3)
y_line = slope*x_line+intercept
plt.plot(x_line, y_line, '-r', label='Age and BMI linear regression')

plt.show()