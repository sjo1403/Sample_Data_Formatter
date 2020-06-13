from openpyxl import load_workbook

def param_func(paramData):
	
	#this dictionary creates a key-value pair for each row of paramData (key to well)
	key = 0	
	paramRow = {}
	while key < len(paramData):
		paramRow[key] = paramData[key]
		key += 1

	#this dictionary creates a key-value pair for each parameter (key to parameter)
	param = {"wellID" : 0, "rate" : 4, "DTW" : 5, "pH" : 7, "conductivity" : 8, "DO" : 11, "temp" : 12, "redox" : 13}

	index = 0
	pos = 0

	currentWell = paramRow[index][param["wellID"]]
	formerWell = currentWell

	#the following statements populate lowFlowForm template with param info
	for row in paramData:
		if index == 0:
			currentWell = paramRow[index][param["wellID"]]
			formerWell = paramRow[0][param["wellID"]]
			
			newbook = load_workbook(filename="GWS_form_" + currentWell + ".xlsx")
			newSheet = newbook.active
			newSheet['E' + str(pos + 26)] = paramRow[index][param["rate"]]
			newSheet['I' + str(pos + 26)] = paramRow[index][param["temp"]]
			newbook.save(filename="GWS_form_" + currentWell + ".xlsx")

			index += 1
			pos += 1

		elif currentWell != formerWell:
			currentWell = paramRow[index][param["wellID"]]
			pos = 0
			
			newbook = load_workbook(filename="GWS_form_" + currentWell + ".xlsx")
			newSheet = newbook.active
			newSheet['E' + str(pos + 26)] = paramRow[index][param["rate"]]
			newSheet['I' + str(pos + 26)] = paramRow[index][param["temp"]]
			newbook.save(filename="GWS_form_" + currentWell + ".xlsx")

			index += 1
			formerWell = paramRow[index - 1][param["wellID"]]
			pos += 1
		
		else:
			currentWell = paramRow[index][param["wellID"]]
			formerWell = paramRow[index - 1][param["wellID"]]
			
			newbook = load_workbook(filename="GWS_form_" + currentWell + ".xlsx")
			newSheet = newbook.active
			newSheet['E' + str(pos + 26)] = paramRow[index][param["rate"]]
			newSheet['I' + str(pos + 26)] = paramRow[index][param["temp"]]
			newbook.save(filename="GWS_form_" + currentWell + ".xlsx")

			index += 1
			pos += 1