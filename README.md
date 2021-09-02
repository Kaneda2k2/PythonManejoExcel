# PythonManejoExcel
#Solo apuntes
##import openpyxl
##
##wb=openpyxl.load_workbook("e.xlsx") #abrimos la hoja de calculos.
##
##print(type(wb))  #comprobamos que esta todo correcto
#canbiar directorios  # os.getcwd()  os.chdir()

#Getting Sheets from the Workbook

##import openpyxl
##wb=openpyxl.load_workbook("e.xlsx") #abrimos la hoja de calculos.
##print(wb.sheetnames) # Los nombres de las hojas
##sheet=wb["Hoja3"] # Get a sheet from the workbook
##print(sheet) #<Worksheet "Hoja3">
##print(type(sheet))  #vemos el tipo de objeto
##print(sheet.title)#Get a sheet from the workbook.
##anotherSheet=wb.active  # Get the active sheet
##print(anotherSheet) # imprime la hoja que esta en activo


#Getting Cells from the Sheets

##import openpyxl
##wb=openpyxl.load_workbook("e.xlsx") #abrimos como siempre el fichero
##type(wb)
##sheet=wb["Hoja1"] #selecionamos la hoja
##sheet["A1"]# Cogemos la celda de la hoja
##print(sheet["A1"]) # ya tenemos el objeto celda
##print(sheet["A1"].value) # Get the value from the cell
##
##c=sheet["C2"] # Selecionamos otra celda
##print(c.value)  # unprimimos el valor
##
##d=sheet["C1"]  #Selecionamos otra celda
##print(d.value) # imprimimos el valor
##
###Get the row,column ,and value form the cell
##print("Row %s,Column %s is %s" % (c.row, c.column,c.value))#Row 2,Column 3 is 85
##print("Cell %s is %s" % (c.coordinate, c.value)) #Cell C2 is 85
##print(sheet["C1"].value)   #73
##
##
##print(sheet.cell(row=1,column=2))
##print(sheet.cell(row=2,column=1).value)  # da el valor
##
##for i in range(1,8,2): # Go through every other row: # pasar por cada dos filas
##    print(i,sheet.cell(row=i,column=2).value)
##


#####################DETERMINAR EL MAXIMO Y EL MINIMO

##import openpyxl
##wb=openpyxl.load_workbook("e.xlsx")
##sheet=wb["Hoja1"]
##print(sheet.max_row) #  te indica el maximo largo
##print(sheet.max_column) #te indica el maximo horizontal


############Converting Betwen column Letters and Numbers



##import openpyxl
##from openpyxl.utils import get_column_letter,column_index_from_string
##
##print(get_column_letter(1))   #Convierta la columna 1 en letra
##print(get_column_letter(2))   # Convierte la columna 2 en Letra
##print(get_column_letter(44))  #convierte la columna 44 en letr
##print(get_column_letter(999)) #convierte la columna 999 en letra
##
##wb=openpyxl.load_workbook("e.xlsx")
##sheet=wb["Hoja1"]
##print(get_column_letter(sheet.max_column))  #Te da el maximo fila en este caso C
##column_index_from_string("A")  #Get A`s number.
##column_index_from_string("AA") # Get number AA -> same before but invert




#############Gettin Rows and Columns form the Sheets

##import openpyxl
##wb=openpyxl.load_workbook("e.xlsx")
##sheet=wb["Hoja1"]
##tuple(sheet["A1":"C3"]) # Coge todas las casillas de la a1 hasta la c3
##print(tuple(sheet["A1":"C3"]))  ##
##for rowOfCellObjects in sheet["A1":"C3"]:
##    for cellObj in rowOfCellObjects:
##        print(cellObj.coordinate,cellObj.value)
##    print("END OF THE ROW")
##
##

########################Worksheet object rows and columns


##import openpyxl
##
##wb=openpyxl.load_workbook("e.xlsx")
##sheet=wb.active
##
##list(sheet.columns)[1]#Get second column cells.
##
##for cellObj in list(sheet.columns)[1]:
##    print(cellObj.value)
##
##
##import openpyxl,pprint
##print("Opening workbook")
##wb=openpyxl.load_workbook("censuspopdata.xlsx")
##type(wb)
##sheet=wb["Population by Census Tract"]
##countyData={}
##
##
##
###print(sheet.max_row,sheet.max_column)
###TODO: Fill in countyData with each countys population and tracts.
##print("Reading rows")
##
##for row in range(2,sheet.max_row+1):
###Each row in the spreadsheet has data for one census tract.
##    state=sheet["B"+str(row)].value
##    county=sheet["C"+str(row)].value
##    pop=sheet["D"+str(row)].value
##    #[estado][condado][poblacion,vecesqueaparece]
##    #print(str(state)+" "+ str(county)+" "+str(pop))
##    countyData.setdefault(state, {})# Make sure the key for this state exists.
##    countyData[state].setdefault(county, {'tracts': 0, 'pop': 0})# Make sure the key for this county in this state exists.
##    countyData[state][county]['tracts'] += 1 # Each row represents one census tract, so increment by one.
##    countyData[state][county]['pop'] += int(pop) # Increase the county pop by the pop in this census tract.
##
###print(countyData)
##
###TODO : Open a new text file and write the contents of countyData to it.
##
##print("Writing results...")
##resultFile=open("census2010.py","w")
##resultFile.write("allData=" + pprint.pformat(countyData))
##resultFile.close()
##print("Done")


##import os
##import census2010
##print(census2010.allData["AK"]["Anchorage"])
##anchoragePop=census2010.allData["AK"]["Anchorage"]["pop"]
##print("The 2010 population of Anchorage was "+ str(anchoragePop))
##########################################################################################


######################  WRITING EXCEL DOCUMENTS
##import openpyxl





####################################################

##import openpyxl

##wb=openpyxl.Workbook()#Crea un workbook en blanco
##wb.sheetnames # It starts with one sheet  ,empieza con una hoja
##sheet=wb.active
###sheet.title
##sheet.title="Spam Bacon Eggs"  # cambia el titulo
##print(wb.sheetnames)  #



####################ABRUR UN XMLSX  Y HACER BACKUP
##wb = openpyxl.load_workbook('e.xlsx')
##sheet = wb.active
##sheet.title = 'Spam Spam Spam' ## cambia el nombre a la hoja
##wb.save('example_copy.xlsx') # Save the workbook.



###########################CREANDO Y REMOVIENDO HOJAS

##import openpyxl
##wb=openpyxl.Workbook()
####wb.save("example1.xlsx")
##wb.sheetnames
##wb.create_sheet() # Add a new sheet
##wb.sheetnames  # ver los nombres de las hojas
##wb.create_sheet(index=0,title="First Sheet") # index=POSICION,title="nombre"
##print(wb.sheetnames)
##print(wb.create_sheet(index=2,title="Middle Sheet"))
##print(wb.sheetnames)
##del wb["Middle Sheet"]  #borrar la hoja con este nombre
##del wb["Sheet1"]#borrar la hoja con este nombre
##print(wb.sheetnames)
###remember use save()method para cambiar los cambios

#################WRITING VALUES TO CELLS

##import openpyxl
##wb=openpyxl.Workbook()
##sheet=wb["Sheet"]
##sheet["A1"]="hello, world!" # Edit the celss value
##print(sheet["A1"].value)
##


##############Updating a Spreadsheet


##import openpyxl
##
##wb=openpyxl.load_workbook("produceSales.xlsx")
##type(wb)
##sheet=wb["Sheet"] # Nombre de la hoja
##
###the produce types and their updated prices
##
##PRICE_UPDATES={"Garlic":3.07,
##"Celery":1.19,
##"Lemon":1.27}
##
##
##
####Loop through the rows and update the prices.
###STEP 2 CHECK ALL ROWS AND UPDATE INCORRECT PRICES
##
##for rowNum in range(2,sheet.max_row):  #skip the first row
##    produceName=sheet.cell(row=rowNum,column=1).value
##    if produceName in PRICE_UPDATES:
##        sheet.cell(row=rowNum,column=2).value = PRICE_UPDATES[produceName]
##wb.save("updatedProducesSales.xlsx")


##TODO:Loop through the rows and update the prices

##if produceName=="Celery":
##    cellObj=1.19
##if produceName =="Garlic":
##    cellObj=3.07
##if produceName=="Lemon":
##    cellObj=1.27

###########################SETTING THE FONTO STYLE OF CELLS
##import openpyxl
##from openpyxl.styles import Font
##
##wb=openpyxl.Workbook()
##sheet=wb["Sheet"]
##italic24Font=Font(size=24, italic=True) # Creamos la fuente
##sheet["A1"].font=italic24Font # aplicamos la fuente a la A1
##sheet["A1"]="Hello, world"
##wb.save("styles.xlsx")


#########################
##import openpyxl
##from openpyxl.styles import Font
##wb=openpyxl.Workbook()
##sheet=wb["Sheet"]
##fontObj1=Font(name="Times New Roman",bold=True)
##sheet["A1"].font=fontObj1
##sheet["A1"]="Bold Times New Roman"
##
##fontObj2=Font(size=24,italic=True)
##sheet["B3"].font=fontObj2
##sheet["B3"]="24 pt Italic"
##
##wb.save("styles1.xlsx")


##########################################################

##############FORMULAS

#sheet["B9"]="=SUM(B1:B8")    # te da el resultado en la celda B9
##
##import openpyxl
##wb=openpyxl.Workbook()
##sheet=wb.active
##sheet["A1"]=200
##sheet["A2"]=300
##sheet["A3"]="=SUM(A1:A2)" #Establecemos la formula
##wb.save("kipi.xlsx")




#############AJUSTANDO ROWS Y COLUMNAS

##import openpyxl
##wb=openpyxl.Workbook()
##sheet=wb.active    # coge la hoja activa
##sheet["A1"]="Tall row"
##sheet["B2"]="Wide column"
###Set the height and width:
##sheet.row_dimensions[1].height=69 # la altura de la fila
##sheet.column_dimensions["B"].width=20
##wb.save("dimensions.xlsx")






#################MERGIN AND UNMERGING CELLS
##
##import openpyxl
##wb=openpyxl.Workbook()
##sheet=wb.active
##sheet.merge_cells("A1:D3")  # Fusiona estas celdas
##sheet["A1"]="Twelve cells merged together"
##sheet.merge_cells("C5:D5")# Merge these thow cells
##wb.save("merged.xlsx")





##################DESFUSIONAR

##import openpyxl
##wb=openpyxl.load_workbook("merged.xlsx")
##sheet=wb.active
##sheet.unmerge_cells("A1:D3") ## split the cells up
##sheet.unmerge_cells("C5:D5") #
##wb.save("merged.xlsx")


##################Freezing panes. congelando paneles
###To unfreeze all panes.set freeze_panes to None or "A1"
##################DOCUMENTACIOON FREEZ PANES
##sheet.freeze_panes = 'A2' Row 1
##sheet.freeze_panes = 'B1' Column A
##sheet.freeze_panes = 'C1' Columns A and B
##sheet.freeze_panes = 'C2' Row 1 and columns A and B
##sheet.freeze_panes =
##'A1' or sheet.freeze_panes = None
##No frozen panes


##
##import openpyxl
##wb=openpyxl.load_workbook("produceSales.xlsx")
##sheet=wb.active
##sheet.freeze_panes="A2" ### Freeze the rows above A2.
##wb.save("freezeExample.xlsx")
##



################################             GRAFICOS


#Steps
#create reference object form a rectangular selection of cells
#create a series object by passing in the reference object
#create a chart object
#append he series object to the chart object
#add the chart object to the worksheet object ,optionally specifiyin which cell should be the top-left corner of the chart

#openpyxl.chart.Reference() # Funtion with 3 arguments


#the worksheet object containing your chart data

#una tupla of two integers, reprenting the top-left cell of the rectangular selection
#the first integer in the tuple is the row.and the secnd is the column.
#nota (la row empieza en la 1 no en la 0
#una tupla of two integers,representing the botton-right cell of the rectangular selecion of cells containing your chart
#the fist integer in the tuple is the row,and the second is the column
#(1,1),(10,1);(3,2),(6,4);(5,3),(5,3)

##
##import openpyxl
##
##wb=openpyxl.Workbook()
##sheet=wb.active
##for i in range(1,11): #Create some data in column A
##    sheet["A"+str(i)]=i
##
##refObj=openpyxl.chart.Reference(sheet,min_col=1,min_row=1,max_col=1,max_row=10)
##seriesObj=openpyxl.chart.Series(refObj,title="First series")
##
##chartObj=openpyxl.chart.BarChart()
##chartObj.title="My Chart"
##chartObj.append(seriesObj)
##
##sheet.add_chart(chartObj,"C5")
##wb.save("sampleChart.xlsx")
##

########################MULTIPLICAIOON TALBLE MAKER

