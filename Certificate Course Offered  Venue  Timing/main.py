import os
import platform
import xlsxwriter
import xlrd

global list,list1,list2,list3
list=[]
list1=[]
list2=[]
list3=[]

loc = ("output.xlsx") 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
sheet.cell_value(0, 0) 
for i in range(sheet.nrows): 
    list.append(sheet.cell_value(i, 0))
loc = ("timing.xlsx") 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
sheet.cell_value(0, 0) 
for i in range(sheet.nrows): 
    list1.append(sheet.cell_value(i, 0))
loc = ("venue.xlsx") 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
sheet.cell_value(0, 0) 
for i in range(sheet.nrows): 
    list2.append(sheet.cell_value(i, 0))
loc = ("certificate.xlsx") 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
sheet.cell_value(0, 0) 
for i in range(sheet.nrows): 
    list3.append(sheet.cell_value(i, 0))
while (1):
	print("""
	Enter 1 : add a new Course
	Enter 2 : To view all course
	Enter 3 : To Search by name
	Enter 4 : To see all timing,venue and Certificate
	""")
	try:
		userInput = int(input("Please Select An Above Option: "))
	except ValueError:
		exit("\nHy! That's Not A Number")
	else:
		print("\n")


	if(userInput == 1): 
		
			newStd = input("Enter course name: ")
			if(newStd in list):
				print("\nThis Student Already In The Database")
			else:
				list.append(newStd)
				workbook = xlsxwriter.Workbook("output.xlsx")
				worksheet = workbook.add_worksheet()
				row = 0
				column = 0
				for item in list :
					worksheet.write(row, column, item)
					row += 1
				workbook.close()
				Timing = input("Enter course timing: ")
				list1.append(Timing)
				workbook = xlsxwriter.Workbook("timing.xlsx")
				worksheet = workbook.add_worksheet()
				row = 0
				column = 0
				for item in list1 :
					worksheet.write(row, column, item)
					row += 1
				workbook.close()
				Venue= input("Enter course Venue: ")
				list2.append(Venue)
				workbook = xlsxwriter.Workbook("venue.xlsx")
				worksheet = workbook.add_worksheet()
				row = 0
				column = 0
				for item in list2 :
					worksheet.write(row, column, item)
					row += 1
				workbook.close()
				c= input("Certificate Y or N ")
				list3.append(c)
				workbook = xlsxwriter.Workbook("certificate.xlsx")
				worksheet = workbook.add_worksheet()
				row = 0
				column = 0
				for item in list3 :
					worksheet.write(row, column, item)
					row += 1
				workbook.close()


			runAgn = input("\nwant To Run Again Y/n: ")
			if(runAgn=="n"):
				break
	elif(userInput == 2):
		print("List Course\n")
		for students in list:
			print("=> {}".format(students))
		runAgn = input("\nwant To Run Again Y/n: ")
		if(runAgn=="n"):
			break

	elif(userInput == 3):
                srcStd = input("Enter Student Name To Search: ")
                if(srcStd in list):
                    print("\n=> Record Found Of Student {}".format(srcStd))
                else:
                    print("\n=> No Record Found Of Student {}".format(srcStd))
                runAgn = input("\nwant To Run Again Y/n: ")
                if(runAgn=="n"):
                	break
	elif(userInput == 4):
		for i in list:
			idx = (list.index(i))
			print ("name: {0:>10} timing: {1:>10} venue: {2:>10} certificate: {3:>10} ".format (i,list1[idx],list2[idx],list3[idx]))

		runAgn = input("\nwant To Run Again Y/n: ")
		if(runAgn=="n"):
			break





