import sys
import codecs
import openpyxl


workbook = openpyxl.load_workbook('C:/Users/Nick/Documents/ATO_Docs/ATO_Alumni_Full_Contact_List_Stewart_Howe_Elevate.xlsx')
sheetnames = workbook.get_sheet_names()
ATOSheet = workbook.get_sheet_by_name('ATOME123')
print (sheetnames)
print (ATOSheet.title)

FirstNames = ATOSheet.columns[0]
MiddleInitials = ATOSheet.columns[1]
LastNames = ATOSheet.columns[2]
GraduatingYear = ATOSheet.columns[5]
States = ATOSheet.columns[9]

#User inputs the initials of a state, number of ATO Alumni residing there is returned
def state_tally(state_initials):

	stateCount = 0

	for cells in States:
		if(cells.value == state_initials):
			stateCount += 1

	return stateCount

#It is assumed that the input to this function is 1 Excel sheet, with a phone numbers column
def valid_phone_numbers(sheet):
	valid_count = 0
	for row in range(2, ATOSheet.max_row + 1):
		phone = ATOSheet['O' + str(row)].value
		if(phone != None):
			print (phone)
			valid_count += 1

	print (valid_count)


state = input("What state do you want to get a count of ATO Alumni in? ")
tally = state_tally(state)

print("You chose the state: %s" %state)
print("%d ATO Alumni reside there, according to our data." %tally)

valid_phone_numbers(ATOSheet)




