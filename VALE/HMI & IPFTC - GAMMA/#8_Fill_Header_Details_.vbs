Set obWCAD = CreateObject("WellCAD.Application")
obWCAD.ShowWindow()


Set obBHDoc = obWCAD.GetBorehole()
Set obHeader = obBHDoc.Header



'read the well name and place it in a variable
		' - Start the loop to check each header item
		NbHeaderItems = obHeader.NbOfItems-1 
		For i = 0 To NbHeaderItems                  
		'Get the name of Each item		
		Item = obHeader.ItemName (i)
		 	If Item = "WELL" Then
		 		strWellName = obHeader.ItemText (Item) 
		 		Exit For
			Else
			End If
		Next



'pass it to the excel sheet for extract and run macro
	Dim objExcel, strExcelPath, objSheet
	
	strExcelPath = "c:\Proc_TV\Templates\FillDeets3.xlsm"
	
	' Open specified spreadsheet and select the first worksheet.
	Set objExcel = CreateObject("Excel.Application")
	objExcel.WorkBooks.Open strExcelPath
	Set objSheet = objExcel.ActiveWorkbook.Worksheets("RunMacro")
	
	'write the well name into the excel document
	objSheet.Cells(1, 1).Value = strWellName
	
	'run the macro to extract the data and reformat it
	objExcel.Application.Run "FillDeets3.xlsm!UseListInColA" 	
	
	' Save and quit.
	objExcel.ActiveWorkbook.Save
	objExcel.ActiveWorkbook.Close
	objExcel.Application.Quit



'Read data from the csv file
	Const ForReading = 1

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objFile = objFSO.OpenTextFile("c:\Proc_TV\Templates\FillDeets.csv", ForReading)
	Dim arrSheet(40,2)
	
	i=0
	Do Until objFile.AtEndOfStream
	    strLine = objFile.ReadLine
	    arrFields = Split(strLine, ",")
	    
	    'write data to an array
		arrSheet(i,0) = arrFields(0)
		arrSheet(i,1) = arrFields(1)
		arrSheet(i,2) = arrFields(2)
		
	'	WScript.Echo arrSheet( i, 0 ) & vbTab & arrSheet( i, 1 ) & vbTab & arrSheet( i, 2 )
 		
 		i=i+1
	Loop

	objFile.Close



'write the stuff out into the header
	'change things in the header

	
'	obHeader.ItemText ":", ""
'	obHeader.ItemText ".", ""
	obHeader.ItemText "COMP", arrSheet( 3, 2 )
	obHeader.ItemText "LOC", arrSheet( 5, 2 )
	obHeader.ItemText "FLD", arrSheet( 6, 2 )
	obHeader.ItemText "STAT", arrSheet( 7, 2 )
	obHeader.ItemText "CNTY", arrSheet( 8, 2 )
	obHeader.ItemText "LOGU", arrSheet( 9, 2 )

	
	obHeader.ItemText "DATE", arrSheet( 11,2 )
	obHeader.ItemText "DRDP", arrSheet( 12,2  )
	obHeader.ItemText "LOTD", arrSheet( 13,2  )
	obHeader.ItemText "RECB", arrSheet( 16,2  )
	
	obHeader.ItemText "PDIP", arrSheet( 18,2  )
	obHeader.ItemText "PAZI", arrSheet( 19,2  )
	obHeader.ItemText "EAST", arrSheet( 20,2  )
	obHeader.ItemText "NRTH", arrSheet( 21,2  )
	obHeader.ItemText "EGL", arrSheet( 22,2  )
	obHeader.ItemText "MAGN", arrSheet( 23,2  )
	
	'obHeader.ItemText "BS#1", arrSheet( 25,2  )
	obHeader.ItemText "RIGN", arrSheet( 25,2  )
	
	'obHeader.ItemText "BS", arrSheet( 26,2  )
	obHeader.ItemText "CASB", arrSheet( 27,2  )
	obHeader.ItemText "CASL", arrSheet( 28,2  )
	obHeader.ItemText "CASX", arrSheet( 27,2  )
	obHeader.ItemText "CASD", arrSheet( 30,2  )