#Before running this script, ensure that openpyxl is installed. 
#Before running this script, always back up the current tracker Excel file, "Grant Expenditure Tracking, FY2019.xlsx"
#Before running this script, export the data from SAP into an Excel file called "AllFunds.xlsx"

import openpyxl
import datetime

wb = openpyxl.load_workbook('AllFunds.xlsx') #this file was exported from SAP
trackingWorkbook = openpyxl.load_workbook('Grant Expenditure Tracking, FY2019.xlsx') #I have been using this Excel file to monitor certain SAP funds over time. Prior to Python, I had been updating this workbook manually.

trackingWorkbookSheetList = trackingWorkbook.sheetnames
numberWorksheets = len(trackingWorkbookSheetList)
mostRecentTrackingWorksheetName = trackingWorkbookSheetList[numberWorksheets-1] #the most recent worksheet is the last one
trackingSheet = trackingWorkbook[mostRecentTrackingWorksheetName]

sheetMaxRows = trackingSheet.max_row
sheetMaxColumns = trackingSheet.max_column
trackingListCount = 0
trackingList = []

#create and populate a list of the funds that I am tracking
for i in range(5,sheetMaxRows-3,1):
	if trackingSheet.cell(row=i, column=1).value != "Total":
		fund = trackingSheet.cell(row=i, column=1).value
		trackingList.append(fund)
		trackingListCount += 1
		


#Pseudocode:
#Run a for-loop to search through column B to identify "Grand Total for"
#For each "Grand Total for" found:
#	Set variable Fund to the value found in that row and in column G.
#	Set variable ytdActualSpending to the value found in that row and column AB
#	Store these values into a worksheet called "Scrap" found in "Grant Expenditure Tracking, FY 2019.xlsx"

#read fund YTD values from exported SAP file, "AllFunds.xlsx", and write them into a worksheet titled "Scrap" that is found in the tracker Excel file ("Grant Expenditure Tracking, FY2019.xlsx")
sheet = wb['AllFunds']
sheetMaxRows = sheet.max_row
sheetMaxColumns = sheet.max_column
scrapSheet = trackingWorkbook['Scrap']
scrapSheetCount = 1

for i in range(1,sheetMaxRows-2,1):
	currentFundCell = sheet.cell(row=i, column=7)
	if currentFundCell.value != None:
		if currentFundCell.value in trackingList:
			print(str(currentFundCell.value) + " $" + str(sheet.cell(row=i,column=29).value))
			scrapSheetColumnA = 'A' + str(scrapSheetCount)
			scrapSheetColumnB = 'B' + str(scrapSheetCount)
			scrapSheet[scrapSheetColumnA] = currentFundCell.value
			scrapSheet[scrapSheetColumnB] = str(sheet.cell(row=i,column=29).value)
			scrapSheetCount += 1

scrapSheetColumnA= 'A' + str(scrapSheetCount+1)			
scrapSheetColumnB= 'B' + str(scrapSheetCount+1)			
scrapSheet[scrapSheetColumnA] = "Updated:"
scrapSheet[scrapSheetColumnB] = datetime.datetime.now().strftime("%I:%M%p on %B %d, %Y")			

trackingWorkbook.save('Grant Expenditure Tracking, FY2019.xlsx')
		
