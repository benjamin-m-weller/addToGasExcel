import openpyxl
import time

#Open up the excel document
workbook = openpyxl.load_workbook('Gas.xlsx')

#Getting a worksheet refernce
worksheet1=workbook.active

def getInput():

   continueInput="y"

   
   
   while(continueInput!="n"):
      #Going to print out the last line entered
      lastLine=getLastLine()-1
      print("The last line entered in the document is: ")
      print("Date\t\t\t"+"Total\t"+"Gallons\t"+"Mileage")
      print(str(worksheet1["A"+str(lastLine)].value)+"\t"+str(worksheet1["B"+str(lastLine)].value)+"\t"+str(worksheet1["C"+str(lastLine)].value)
            +"\t"+str(worksheet1["D"+str(lastLine)].value))
      
      global totalCost
      totalCost=input("Please enter the total filled: ")
      global gallons
      gallons=input("Please enter the number of gallons you used: ");
      global mileage
      mileage=input("Please enter your mileage: ")
      global dateEntered
      dateEntered=input("Please enter the date you filled up: ")

      continueInput=input("Please enter \"n\" if you want to stop inputing data: ")
   
      #Now going to pass along the duties of putting things into the book for
      #a differnt function
      fillBook()
      
   workbook.save('Gas.xlsx')

def getLastLine():
    currentRow=1
    cellToCheck="A"+str(currentRow)
    
    #Now the loop
    while(worksheet1[cellToCheck].value != None):
       currentRow+=1
       cellToCheck="A"+str(currentRow)
	
    return currentRow

def fillBook():
        #I understand that there is some simple dynamic programming I could do here, but this is really my first swing at python.
	lastLine=getLastLine()
	
	#Here comes some application specific stuff
		
	worksheet1["A"+str(lastLine)]=dateEntered
	worksheet1["B"+str(lastLine)]=totalCost
	worksheet1["C"+str(lastLine)]=gallons
	worksheet1["D"+str(lastLine)]=mileage
	
	#Now for formulas
	worksheet1["E"+str(lastLine)]="=B"+str(lastLine)+"/C"+str(lastLine)
	worksheet1["F"+str(lastLine)]="=D"+str(lastLine)+"/C"+str(lastLine)
	worksheet1["G"+str(lastLine)]="=A"+str(lastLine)+"-A"+str((lastLine-1))
	
	#workbook.save('Gas.xlsx')
	
getInput()

	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
    
