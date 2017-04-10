#This file will contain all helper functions designed to interact with excel file
import sys
import codecs


#User inputs the initials of a state, function ouputs contact info of Alumni who currently reside there
def info_by_state(state_initials, sheet):

	stateCount = 0

	for row in range(2, sheet.max_row + 1):
		state = sheet['J' + str(row)].value

		if(state == state_initials):

			FirstName = sheet['A' + str(row)].value
			LastName = sheet['C' + str(row)].value
			MiddleInitial = sheet['B' + str(row)].value
			if(MiddleInitial == None):
				MiddleInitial = ""		#Can't print "NoneTypes", if nothing found print empty string

			Phone = sheet['O' + str(row)].value
			if(Phone == None):
				Phone = "#No Phone Info Available#"		#Can't print "NoneTypes", if nothing found print empty string

			#At some point replace this with an output to text file, or Excel spreadsheet?
			print (FirstName + " " + MiddleInitial + " " + LastName + "  " + Phone)
			stateCount += 1

	return stateCount


#It is assumed that the input to this function is 1 Excel sheet, with a phone numbers column
def invalid_phone_numbers(sheet):
	invalid_count = 0
	for row in range(2, sheet.max_row + 1):
		phone = sheet['O' + str(row)].value
		if(phone == None):
			FirstName = sheet['A' + str(row)].value
			LastName = sheet['C' + str(row)].value
			MiddleInitial = sheet['B' + str(row)].value
			if(MiddleInitial == None):
				MiddleInitial = ""		#Can't print "NoneTypes", if nothing found print empty string

			#At some point replace this with an output to text file, or Excel spreadsheet?
			print (FirstName + " " + MiddleInitial + " " + LastName)
			invalid_count += 1

	print ("Number of people without phone number info: %d" %invalid_count)