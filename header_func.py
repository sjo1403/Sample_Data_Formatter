from openpyxl import load_workbook
newbook = load_workbook(filename="lowFlowForms.xlsx")
newSheet = newbook.active

def header_func(headerData):

	#this dictionary creates a key-value pair for each row of headerData (key to well)
	key = 0	
	well = {}
	while key < len(headerData):
		well[key] = headerData[key]
		key += 1

	#this dictionary creates a key-value pair for each header-item (key to header-item)
	headers = {"projectNumber" : 0, "client" : 1, "location" : 2, "wellID" : 4, "date" : 3, "sampler" : 5, "weather" : 6,
				"measuringPoint" : 9, "diameter" : 11, "casing" : 12, "DTW" : 14, "DTB" : 15, "equipment" : 21, "purgeBegin" : 26, 
				"WQM" : 27, "purgeEnd" : 29, "sampleTime" : 36, "color" : 37, "odor" : 38}

	index = 0

	#populates lowFlowForm template with header info
	for item in headerData:
		newSheet['D10'] = well[index][headers["projectNumber"]]
		newSheet['D9'] = well[index][headers["client"]]
		newSheet['L10'] = well[index][headers["location"]]
		newSheet['S10'] = well[index][headers["wellID"]]
		newSheet['L11'] = well[index][headers["sampler"]]
		newSheet['L12'] = well[index][headers["sampler"]]
		newSheet['D13'] = well[index][headers["weather"]]
		newSheet['D11'] = well[index][headers["date"]]
		newSheet['E19'] = well[index][headers["diameter"]]
		newSheet['E18'] = well[index][headers["casing"]]
		newSheet['E21'] = well[index][headers["DTW"]]
		newSheet['E20'] = well[index][headers["DTB"]]
		newSheet['E16'] = well[index][headers["WQM"]]
		newSheet['O18'] = well[index][headers["equipment"]]
		newSheet['D12'] = well[index][headers["sampleTime"]]
		newSheet['O21'] = well[index][headers["purgeBegin"]]
		newSheet['T21'] = well[index][headers["purgeEnd"]]
		newSheet['H41'] = well[index][headers["color"]]
		newSheet['M41'] = well[index][headers["odor"]]


		newbook.save(filename="GWS_form_" + well[index][headers["wellID"]] + ".xlsx")

		index += 1