# this program was designed to expedite the process of creating my expense report for work
# it takes three files as input; an ical of my appointments, a text file of the address and thier
# corrisponding ids or names, and a text file of the distances betwen each possible combination of
# addresses/locations.  It creates a CSV file with the date of event, distance traveled, starting 
# address, ending address, and comments for while that drive was needed.

# to do:
# 1 complete documentation
# 2 change address_list and distance_list to dict
# 3 add user defined location descriptions
# 4 add user defined vars i.e. path to ical export, date range, output destination
# 5 add gui to facilitate the addition of new locations, execution of the program, and change of date range
# 6 extend google maps support to replace the need of inputing distances if the user decides to do so

import webbrowser
from datetime import date

# enumerated variables for the index of the events in the list
class Event_Index:
	DATE = 0			# date of the event
	TIME = 1			# time of the event
	START_ID = 2		# starting locations id
	START_ADDRESS = 3	# starting locations address
	END_ID = 4			# ending locations id
	END_ADDRESS = 5		# ending locations address
	DISTANCE = 6		# driving distance between the two locations
	COMMENTS = 7		# comments of why the trip was made

def CreateKeyIds():
	file = open('./addresses.txt', 'r')
	id = ""
	key_ids = []
	temp = []
	
	for line in file:
		temp = line.split(" ")
		key_ids.append(temp[0])
	
	return key_ids

# reads the iCal file and parses the needed data (date and start_id)
# requires a iCal file
# returned list will have date, time, and start_id in that order
def CreateEvent_list():
	file = open('/Users/mjonas/Desktop/Completed.ics', 'r')
	summaryline = ""	# line containing the start_id from the file
	date = 0			# date of event as an int
	time = 0			# time of event as an int
	save = True			# set to false if the event has been deleted
	event_list = []		# list to hold the parsed data
	count = 0			# count of correct items saved for each event
	
	key_ids = CreateKeyIds()		# keys to be searched for in the file
	
	# returns the date range, the 16th of last month to the 15th of this month
	(start, end) = FindDateRange()
	
	# loop through the file searching for events that are opaque (not deleted)
	for line in file:
		if "VEVENT" in line:
			for line in file:
				if "TRANSP" in line and not "OPAQUE":
					save = False
					break
				elif "SUMMARY:" in line:
					# remove "SUMMARY:" and the unnecessary white space characters
					line = line.strip('\r\n')
					line = line[8:]
					summaryline = line.lower()
					if not summaryline in key_ids:
						save = False
						break
					else:
						count = count + 1
				elif "DTSTART" in line:
					# parse the date and the time from the line
					date = int(line[33:41])
					time = int(line[42:46])
					if date < start or date > end:
						save = False
						break
					else:
						count = count + 1
				elif "END:VEVENT" in line:
					break
		if save and count == 2:
			event_list.append([date, time, summaryline])
			count = 0
		else:
			save = True
			summaryline = ""
			date = 0
			time = 0
			count = 0

	# sort the list by date and time
	event_list.sort()
	file.close()
	return event_list

# reads the addresses file and creates a list of addresses with their corrisponding ids
# requires a file with the location id and address delineated by a space
# returned list will have id and address in that order
def CreateAddressList():
	file = open('./addresses.txt', 'r')
	id = ""									# holds the id of the address
	address = ""							# holds the address
	index = 0								# index of the first white space in the line from the file
	addressList = []						# list to be returned containing the list of addresses and ids
	
	# loop through the file and parse the id and address from each line
	for line in file:
		# find the index of the white space seperating the id from the address
		line = line.strip('\n')
		index = line.index(" ")
		address = line[index+1:]
		id = line[:index]
		addressList.append([id, address])
		id = ""
		address = ""
	
	file.close()
	return addressList

# reads the location ids with the corrisponding distance between them from a file and creates a list
# requires a file contianing the location id, id, and distance between them delineated by a space
# the returned list will have the id, id, and distance in that order
def CreateDistanceList():
	file = open('./distance.txt', 'r')
	distanceList = []						# list to be returned containing the ids and distances
	# loop through the file and parse the ids and distances
	for line in file:
		line = line.strip('\n')
		distanceList.append(line.split(" "))
	
	file.close()
	return distanceList

# this function will take the list and append the corrisponding address
# requires the list of events and the index of the id of the event in the list for which an address will be appended
# returns the original list with addresses for each location appended
def AssignAddresses(event_list, id):
	addressList = []					# list of id/names and their corrisponding addreses
	addressList = CreateAddressList()
	# loop through the event_list, if the id is not home, look for the matching id in the address list
	for i, event in enumerate(event_list):
		if event[id] != "home" and event[Event_Index.START_ID] != "home":
			for j, address in enumerate(addressList):
				if event[id] == address[0]:
					# append the appropriate address to the list
					event.append(address[1])
					break
	return event_list

# this funciton inserts the home location at the beginning of each day and the end of each day
# requires the event_list
# returns the event_list with the home locations inserted
def InsertHome(event_list):
	# insert home entry at the begining of the list
	event_list.insert(0, [event_list[0][Event_Index.DATE], 0, "home", "home less milage", event_list[0][Event_Index.START_ID], event_list[0][Event_Index.START_ADDRESS]])
	# loop through the list inserting home location at the end and beginning of each day
	for i, event in enumerate(event_list):
		if i!=0 and event_list[i-1][Event_Index.DATE] != event[Event_Index.DATE]:
			# inserting at the beginning of the day
			event_list.insert(i, [event[Event_Index.DATE], 0, "home", "home less milage", event[Event_Index.START_ID], event[Event_Index.START_ADDRESS]])
			# inserting at the end of the day
			event_list[i-1].append("home")
			event_list[i-1].append("Home Less Milage")
	# inserting at the end of the list
	event_list[len(event_list)-1].append("home")
	event_list[len(event_list)-1].append("Home Less Milage")
	return event_list

# this function appends the ending location to each event in the list, the ending locations are determined
# by the locaiton of the next event in the list and its date
# requires the event_list
# returns the event_list with corrisponding ending locations for each event
def AssignEndLocation(event_list):
	# loops through the list, checks if the event contains a "home" id and skips those since they have
	# already been assigned, then assigns the ending locaitons
	for i, event in enumerate(event_list):
		if event[Event_Index.START_ID] != "home" and len(event) == 4:
			event.append(event_list[i+1][Event_Index.START_ID])
	return event_list

# this function appends the distance between the two location in each event
# it requires the event_list and the list of distances between each event
# returns the event_list with with distances between each event appended
def AssignDistance(event_list):
	distanceList = CreateDistanceList()
	for i, event in enumerate(event_list):
		for j, distance in enumerate(distanceList):
			if ((event[Event_Index.START_ID] == distance[0] and event[Event_Index.END_ID] == distance[1]) or (event[Event_Index.END_ID] == distance[0] and event[Event_Index.START_ID] == distance[1])):
				event.append(int(distance[2]))
				break
	return event_list

# this function appends the appropriate comments displaying the need for each entry in the expense report
# it requires the event_list
# returns the event_list with the comments appended
def AssignComments(event_list):
	for i, event in enumerate(event_list):
		endId = event[Event_Index.END_ID]
		startId = event[Event_Index.START_ID]
		# if the end location is home and the start location is not the torrace office
		if endId == "home" and startId != "torrance":
			endId = event[Event_Index.START_ID]
			event.append("Client Appointment(%s)" %(endId[:2]))
		# if the end location is home and the start location is the torrance office
		elif endId == "home" and startId == "torrance":
			event.append("Home")
		# if the end location is the torrance office
		elif endId == "torrance":
			event.append("Torrance office")
		# if the end location is other than the torrance office or home
		else:
			event.append("Client Appointment(%s)" %(endId[:2]))
	return event_list

# this function returns the date range, the start date, date of the 16th of the previous month
# and the end date, the 15th of the current month, in the format YYYYMMDD,
# the dates are returned as integers
def FindDateRange():
	today = date.today()		# current date
	# assign the end date, the 15th of the current month
	if today.month >= 10:
		endDate = "%s%s15" %(today.year, today.month)
	else:
		endDate = "%s0%s15" %(today.year, today.month)

	# assign the start date, the 16th of the previous month
	# if the current month is January
	if today.month == 1:
		startDate = "%s1216" %((today.year-1))
	# if the current month is not January
	elif today.month >= 10:
		startDate = "%s%s16" %(today.year, (today.month-1))
	else:
		startDate = "%s0%s16" %(today.year, (today.month-1))
	start = int(startDate)
	end = int(endDate)
	return (start, end)

# this function takes a integer date in the form of YYYYMMDD and retuns it in the form MM/DD/YYYY
def GetPrintDate(date):
	temp = str(date)
	return "%s/%s/%s" %(temp[4:6], temp[6:], temp[0:4])

# this function takes the event_list and prints it to a CSV file with each event on a new line
# it inserts a blank line every 28 entries to provided easier cuting and pasting into the template
# it requires the event_list
# it prints to the expense.csv file located on the desktop in the format
# date,distance,start_address,end_address,comments
def PrintToCsv(event_list):
	count = 0				# number of lines printed to file, used to insert a new line at line 28
	totalDistance = 0		# total distance, printed and reset each 28 lines printed
	file = open('/Users/mjonas/Desktop/expense.csv', 'w')
	
	for i, event in enumerate(event_list):
		date_of_event = GetPrintDate(event[Event_Index.DATE])
		if int(event[Event_Index.DISTANCE]) > 0:
			file.write("%s,%s,%s,%s,%s\n" %(date_of_event,event[Event_Index.DISTANCE],event[Event_Index.START_ADDRESS],event[Event_Index.END_ADDRESS],event[Event_Index.COMMENTS]))
			totalDistance = totalDistance + event[Event_Index.DISTANCE]
			count += 1
		if count == 28:
			file.write("total miles,%s,reinbursment at .55 per mile,$%s\n" %(totalDistance, (totalDistance*0.55)))
			totalDistance = 0
			count = 0
	
	file.write("total miles,%s,reinbursment at .55 per mile,$%s\n" %(totalDistance, (totalDistance*0.55)))
	file.close()

# this function prints a list of lists to the terminal
# it requires a list of lists
# it prints each list from the list of lists on a seperate line and then prints the number of lists printed
def printList(event_list):
	count = 0
	print
	for item in event_list:
		print item
		count+=1
	print "the number of lists printed: %s\n" %(count)

def openMap(address, destination):
	base_url = 'http://maps.google.com/maps?q='
	url = base_url + address.replace(' ', '+') + '+to+' + destination.replace(' ', '+')
	webbrowser.open_new_tab(url)

# create expense CSV
def createExpense():
	event_list = CreateEvent_list()
	event_list = AssignAddresses(event_list, Event_Index.START_ID)
	event_list = InsertHome(event_list)
	event_list = AssignEndLocation(event_list)
	event_list = AssignAddresses(event_list, Event_Index.END_ID)
	event_list = AssignDistance(event_list)
	event_list = AssignComments(event_list)
	#printList(event_list)
	PrintToCsv(event_list)

# create address list and print
def printIdAndAddress():
	address_list = CreateAddressList()
	print
	print 'Id\t\tAddress'
	for item in address_list:
		print item[0] +'\t\t'+ item[1]

def addIdAndAddress():
	id = ""
	address = ""
	distance = 0
	distance_list = []
	
	#get location name from user
	id = str(raw_input("Enter location's id(name): "))
	
	#get location address from user
	address = str(raw_input("Enter location's address(without punctuation): "))
	
	address_list = CreateAddressList()
	
	#openMap() for each possible destination
	for item in address_list:
		openMap(address, item[1])
	
	#get distances for each possible destination from user in reverse order of how they called to openMap()
	address_list.reverse()
	for item in address_list:
		while True:
			temp = raw_input("Enter distance between " +id+ " and " +item[0]+ ": ")
			try:
				distance = int(temp)
				distance_list.append([id,item[0],distance])
				break
			except ValueError:
				print "Error -  that was not an integer, try again"
	
	#append address file
	file = open('./addresses.txt', 'a')
	file.write("%s %s\n" %(id, address))
	file.close()
	
	#append distances file
	file = open('./distance.txt', 'a')
	for item in distance_list:
		file.write("%s %s %s\n" %(item[0], item[1], item[2]))
	file.close()
	
	print "location: %s %s added." %(id, address)

def deleteIdAndAddress():
	id = ""
	response = ""
	item_in_file = False

	while True:
		#get location name from user
		print "You would like to permanently delete a location."
		id = str(raw_input("Enter location's id(name): "))
		print "Are you sure, you want to delete %s?" %(id)
		response = str(raw_input("Enter yes to continue or no to cancel: "))
		response = response.lower()

		if 'n' in response:
			break
		elif 'y' in response:
			address_list = CreateAddressList()
			for item in address_list:
				if id in item:
					item_in_file = True
					break
			if item_in_file:
				# erase lines from address file
				file = open('./addresses.txt', 'w')
				for item in address_list:
					if not id in item:
						file.write("%s %s\n" %(item[0], item[1]))
				file.close()
				# erase lines from distance file
				distance_list = CreateDistanceList()
				file = open('./distance.txt', 'w')
				for item in distance_list:
					if not id in item:
						file.write("%s %s %s\n" %(item[0], item[1], item[2]))
				file.close()
				print "%s was deleted." %(id)
				break
			else:
				print "%s was not in the list." %(id)
				break
		else:
			print "Error - please enter yes to continue or no to cancel"
		
		id = ""
		response = ""
		item_in_file = False

def menu():
	option = 0
	selection = ""
	while True:
		while True:
			print "\nCreate Expense Menu\n"
			print "1\tCreate expense"
			print "2\tShow id and address"
			print "3\tAdd id and address"
			print "4\tDelete id and address"
			print "5\tExit\n"
			selection = raw_input("Enter selection: ")
			
			#check if selection is an int
			try:
				option = int(selection)
				if option > 0 and option < 6:
					break
				else:
					print "Error - enter an integer(1-5), try again"
			except ValueError:
				print "Error -  that was not an integer, try again"
	
		if option == 1:
			createExpense()
		elif option == 2:
			printIdAndAddress()
		elif option == 3:
			addIdAndAddress()
		elif option == 4:
			deleteIdAndAddress()
		else:
			print "goodbye"
			break
		
		option = 0
		selection = ""

menu()
