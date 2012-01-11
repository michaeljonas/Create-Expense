# this program was designed to expedite the process of creating my expense report for work
# it takes three files as input; an ical of my appointments, a text file of the address and thier
# corrisponding ids or names, and a text file of the distances betwen each possible combination of
# addresses/locations.  It creates a CSV file with the date of event, distance traveled, starting 
# address, ending address, and comments for while that drive was needed.

# to do:
# 1 add gui to facilitate the addition of new locations, execution of the program, and change of date range
# 2 possibly add google maps support to replace the need of inputing distances by hand however this
# may not be practical since they are drive paths and the google maps drive path may not be representive
# of the actual drive path
# 3 add support for other locations, such as "all staff" or other offices


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

# reads the iCal file and parses the needed data (date and start_id)
# requires a iCal file
# returned list will have date, time, and start_id in that order
def CreateEvent_list():
	file = open('/Users/mjonas/Desktop/Completed.ics', 'r')
	summaryline = ""														# line containing the start_id from the file
	dateString = ""															# line containing the date from the file
	date = 0																# date of event as an int
	time = 0																# time of event as an int
	initials = ["cm", "et", "jm", "sb", "dw", "torrance", "ets", "jms"]		# keys to be searched for in the file
	event_list = []															# list to hold the parsed data
	
	# returns the date range, the 16th of last month to the 15th of this month
	(start, end) = FindDateRange()
	
	# loop through the file searching for events that are opaque (not deleted)
	for line in file:
		if "OPAQUE" in line:
			# loop through the file, searching for each line containing "SUMMARY:"
			for line in file:
				if "SUMMARY:" in line:
					summaryline = line
					# remove "SUMMARY:" and the unnecessary white space characters
					summaryline = summaryline.strip('\r\n')
					summaryline = summaryline[8:]
					summaryline = summaryline.lower()
					# check if the summary is contained in the list of keys, if so loop to gain the date of the event
					if summaryline in initials:
						for line in file:
							if "DTSTART" in line:
								dateString = line
								# parse the date and the time from the line
								date = int(dateString[33:41])
								time = int(dateString[42:46])
								break
						# if the date is in the needed range (16th of last month<date<15th of this month)
						# append it to the list
						if date >= start and date <= end:
							event_list.append([date, time, summaryline])
						break
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
def PrintToCsv(aList):
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

event_list = CreateEvent_list()
event_list = AssignAddresses(event_list, Event_Index.START_ID)
event_list = InsertHome(event_list)
event_list = AssignEndLocation(event_list)
event_list = AssignAddresses(event_list, Event_Index.END_ID)
event_list = AssignDistance(event_list)
event_list = AssignComments(event_list)
#printList(event_list)
PrintToCsv(event_list)