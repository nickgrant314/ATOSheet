#This file will contain all user interaction
import ATOFileConnect
import ExcelFileHelperFunctions

#Text commonly reused put into functions
def help_text():
	print("If you want Alumni info based on state, type 'State Lookup'")
	print("If you want to retrieve a list of Alumni without current phone info, type 'Invalid Phone Lookup'")
	print("If you would like to exit this experience, type 'I\'m done'")
	print()
	print("If you ever need help remembering function names, type 'Help'")

def user_prompt():
	return input("So what would you like to do? ")



#Returns an Excel Sheet with Alumni data we're concerned with
ATOSheet = ATOFileConnect.file_connect()

#Greet user
help_text()
user_prompt()

#Loop until user tells us they're done
while(userChoice != "I'm done"):

	if(userChoice == "State Lookup"):
		state = input("What state do you want ATO Alumni about? ")
		tally = ExcelFileHelperFunctions.info_by_state(state, ATOSheet)

		print("You chose the state: %s" %state)
		print("%d ATO Alumni reside there, according to our data." %tally)

	elif(userChoice == "Invalid Phone Lookup"):
		ExcelFileHelperFunctions.invalid_phone_numbers(ATOSheet)

	elif(userChoice == "Help"):
		help_text()

	elif(userChoice == "I'm done"):
		break

	else:
		print("You've entered an invalid command Please try again.")
		help_text()

	userChoice = user_prompt()

#End Program
print("Thank you for accessing the ATO Alumni Database. Please come back again some time!")
