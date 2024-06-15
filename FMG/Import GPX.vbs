
Set obWCAD = CreateObject("WellCAD.Application")
obWCAD.Showwindow()


'Definition of the root directory for templates
'Const RootPath = "c:\Proc_TV\05_FMG_TV_LasPrep"
RootPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)

Set objFSO = CreateObject("Scripting.FileSystemObject")
'look in the folder and loop throuigh all files
For Each objFile In objFSO.GetFolder(RootPath).Files

'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
' looping through files begin


	'if the file has an .las extention then perform the code, otherwise, skip this file
  	If LCase(objFSO.GetExtensionName(objFile.Name)) = "las" Then
		' get the las file name
	    strLASFileName = objFSO.GetFileName(objFile)
	    ' Look in the LAS filename for the first extention dot
	    Dpos = InStr (strLASFileName, ".")
		' create a variable holding just the filename without the .wcl extension
	    strLASFileNameNoExt = Left(strLASfilename, DPos-1)


		' import las file and give it the handle "obBHDoc"
		Set obBHDoc = obWCAD.FileImport(RootPath & "\" & strLASFileName, False, "C:\Proc_TV\05_FMG_TV_LasPrep\templates\AutoBulkLoad.ini","c:\Proc_TV\05_FMG_TV_LasPrep\log.txt")

				
		'=================================
        ' ::: Find out if CALIPER exists in the las data 
        '  - Start the loop to check each log title, based on log index, loops backward through the index values (names)
        
        'make comment
        strComment = "NO CALIPER. "
        
        NbLogs = obBHDoc.NbOfLogs-1        'this gives you the number of logs in the borehole document
        For i = 0 To NbLogs                    'For loop through nblogs number of times
          
          Set obLog = obBHDoc.Log(CInt(CStr(NbLogs - i))) 'ensure an integer is passed
          title = obLog.Name
          
                  ' does CALIPER exist? 
                  If UCase(title) = "CALIPER" Then
                  		'convert caliper to mm from cm
						Set obNewLog = obBHDoc.InsertNewLog (2)
						obNewLog.Name = "CAL_MM"
						obNewLog.Formula = "({CALIPER} * 10)"
						'delete original cal
						obBHDoc.RemoveLog ("CALIPER"),False			
						'change name to cal
						Set obNewLog = obBHDoc.Log("CAL_MM")
						obNewLog.Name = "CAL"
						obNewLog.ScaleHigh = 200
						obNewLog.ScaleLow = 0
                    	strComment = ""
                   		Exit For ' exits the for loop
                  Else
                  
                  End if
        Next
        ' ::: Find out if CALIPER exists in the las data
        '=================================

		'=================================
		' ::: determine which set of azimuth data are present
		boolNSGSANGbExists = False
		boolGSANGbExists = False
		boolAZIDExists = False
		
        '  - Start the loop to check each log title, based on log index, loops backward through the index values (names)
        NbLogs = obBHDoc.NbOfLogs-1        'this gives you the number of logs in the borehole document
        For i = 0 To NbLogs                    'For loop through nblogs number of times
          
          Set obLog = obBHDoc.Log(CInt(CStr(NbLogs - i))) 'ensure an integer is passed
          title = obLog.Name
          
                  ' does Gazimuth exist? set Template to use default gazi
  
                  TemplateFlag = "NOTSET"
				
				Select Case UCase(title)
					Case  "NSGSANGB" 
						boolNSGSANGbExists = True
					Case  "GSANGB" 
						boolGSANGbExists = True
					Case  "AZID" 
						boolAZIDExists = True
				End select
        Next
        ' :::    
        '=================================        
        
        
        '=================================
        'deterine the number of gamma columns
        	boolGAMMAOExists = False
        intGammas = 0
        NbLogs = obBHDoc.NbOfLogs-1        'this gives you the number of logs in the borehole document
        For i = 0 To NbLogs                    'For loop through nblogs number of times
          
			Set obLog = obBHDoc.Log(CInt(CStr(NbLogs - i))) 'ensure an integer is passed
			title = obLog.Name
		      
			' does Gamma exist? need to search based on the first 5 characters only
			If UCase(Left(title,5)) = "GAMMA" Then
				'icrement gammas
				intGammas = intGammas + 1
			End If
			
			If title = "GAMMAO" Then
			   	boolGAMMAOExists = True
			End If
			     
        Next
        'deterine the number of gamma columns
        '=================================
        
        'if gamma = 1 skip
        'if gamma > 1 then pick one
        
		'=================================
        ' ::: 		' pick the correct gamma and delete the other one(S)
        '  - Start the loop to check each log title, based on log index, loops backward through the index values (names)
        If intGammas > 1 Then
        
	        NbLogs = obBHDoc.NbOfLogs-1        'this gives you the number of logs in the borehole document
	        For i = 0 To NbLogs                    'For loop through nblogs number of times
	          
				Set obLog = obBHDoc.Log(CInt(CStr(NbLogs - i))) 'ensure an integer is passed
				title = obLog.Name
			    

				' does Gamma exist? need to search based on the first 5 characters only
				If UCase(Left(title,5)) = "GAMMA" Then
					'read the units
					Unit = obLog.LogUnit
					
					'if the units are not API-GR then delete
					If Unit <> "API-GR" Then
						obBHDoc.RemoveLog (title),FALSE
					Else
						strAPIGammaName = obLog.Name
					End If
										
				Else
				
				End if
	        Next
	        
	        'rename the remaining gamma
	        Set obLog = obBHDoc.Log(strAPIGammaName)
	   		obLog.Name = "GAMMA"	
        
        End If
        ' ::: Find out if GSANGB exists in the las data
        '=================================

		If intGammas = 1 And boolGAMMAOExists Then
		    Set obLog = obBHDoc.Log("GAMMAO")
   			obLog.Name = "GAMMA"	
		End If 

			
		' apply the geophys import template
		obBHDoc.ApplyTemplate  "C:\Proc_TV\05_FMG_TV_LasPrep\templates\GEOPHYSICS IMPORTd2.wdt", false, true, false, false, True


		'=================================
		' ::: Delete non essential columns as defined in PWS_Lookup_DeleteTheseColumnsList_01.ini
		
		'Get current borehole document
		Set obBHole = obWCAD.GetBorehole()
		
		' - Start the loop to check each log title, based on log index, loops backward through the index values (names)
		NbLogs = obBHole.NbOfLogs-1        'this gives you the number of logs in the borehole document
		For i = 0 To NbLogs                    'For loop through nblogs number of times
		  
		  Set obLog = obBHole.Log(CInt(CStr(NbLogs - i))) 'ensure an integer is passed
		  title = obLog.Name
		  
		  'Get access to the lookup file
		  Set FSO = CreateObject("Scripting.FileSystemObject")
		  Set obLookUp = FSO.OpenTextFile("C:\Proc_TV\05_FMG_TV_LasPrep\templates\PWS_Lookup_DeleteTheseColumnsList_01.ini", 1)
		  
		  'Find log in lookup table 
		  LogFound = True
		 
		  Do While Not obLookUp.AtEndOfStream
		    line = obLookUp.ReadLine
		    
		    If Left(line,1) = "[" Then
		    
		      If Ucase(Mid(line,2,Len(line)-1)) = Ucase(title) Then
		        LogFound = True
		        obBHDoc.RemoveLog (title),FALSE
		      Else
		        LogFound = FALSE
		      End If
		      
		    End If
		  Loop
		  'Close ini file
		  obLookUp.Close
		  
		Next
		
		' ::: Delete non essential columns
		'=================================
   	
   	
   		'set page setup to fan fold		
		Set obPage = obBHDoc.Page
		'Set paper mode to fanfold
		obPage.PaperMode = 1
	



		'=================================
		' ::: edit based on gyro or mag deviation, logs and header info
		Set obHeader = obBHDoc.Header
		
		If boolNSGSANGbExists = True Then
				strComment = strComment & "NSGYRO DEVIATION"
				obHeader.ItemText "Comments", strComment
		Else		
			If boolGSANGbExists = True Then
				strComment = strComment & "GYRO DEVIATION"
				obHeader.ItemText "Comments", strComment
			Else
				strComment = strComment & "MAG DEVIATION"
				obHeader.ItemText "Comments", strComment
			End If 'boolGSANGbExists = True	
		End If 'boolNSGSANGbExists = True
		
		
		
		If boolNSGSANGbExists = False Then
				obBHDoc.RemoveLog ("NSGSANG"),FALSE
				obBHDoc.RemoveLog ("NSGSANGB"),FALSE	
		End If 'boolNSGSANGbExists = False
		
		If boolGSANGbExists = False Then
				obBHDoc.RemoveLog ("GSANG"),FALSE
				obBHDoc.RemoveLog ("GSANGB"),FALSE	
		End If 'boolGSANGbExists = False	
		
		    
		' ::: edit based on gyro or mag deviation, logs and header info
		'=================================


		' - change the log colours
'		Set obLog = obBHole.Log("GSANGB")
'		obLog.PenColor  16711680  ' the number for blue
		'color = obLog.PenColor
'		Set obLog = obBHole.Log("GSANG")
'		obLog.PenColor  255  ' the number for red
		
		'change things in the header
		Set obHeader = obBHDoc.Header

		obHeader.ItemText ":", "OPTICAL AND ACOUSTIC IMAGE LOG"
		obHeader.ItemText ".", "ORIENTED TO HIGH SIDE"
		obHeader.ItemText "CNTY", "AUSTRALIA"
		obHeader.ItemText "PD", "M.S.L"
		obHeader.ItemText "MAGN", "1.438"
		
		obHeader.ItemText "FTYP", ""
		obHeader.ItemText "MRS", ""	
		obHeader.ItemText "MTP", ""				
		obHeader.ItemText "MFTP", ""
		obHeader.ItemText "MCRS", ""
		
		obHeader.ItemText "CASD", ""

		'change the bit size
		' - Start the loop to check each header item
		NbHeaderItems = obHeader.NbOfItems-1 
		For i = 0 To NbHeaderItems                  
		'Get the name of Each item		
		Item = obHeader.ItemName (i)
		 	'If Item = "BS" Then
		 		'strBitSize = obHeader.ItemText (Item) 
				'strBitSize = CSng(strBitSize)
				'obHeader.ItemText "BS", strBitSize*10
		 		'Exit For
			'Else
			'End If
		Next


        'rename GAMMA
       
        Set obLog = obBHDoc.Log("GAMMA")
   		obLog.Name = "NG"	
                
        'rename TILD
		       
	Set obLog = obBHDoc.Log("TILD")
	
	obLog.Name = "TILT"	
        
        'rename AZID
	
	Set obLog = obBHDoc.Log("AZID")
	
	obLog.Name = "AZI"
	
	 'rename NSGSANG
			       
	Set obLog = obBHDoc.Log("NSGSANG")
		
	obLog.Name = "TILT NSG"	
	        
	'rename NSGSANGB
		
	Set obLog = obBHDoc.Log("NSGSANGB")
		
	obLog.Name = "AZI NSG"
        


		' save and close
		obBHDoc.SaveAs (RootPath & "\" & strLASFileNameNoExt & "_Gamma Match File" & ".WCL")
		'obWCAD.CloseBorehole False, 0


	Else
	'do nothing
 	End If
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
' looping through files end
Next
' this is the end of the main script, beyond are subfunctions which are called from the main script



'=====================================================================================================================
' getting data from fmpro to wellcad header
'=====================================================================================================================
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
	
	obHeader.ItemText "BS#1", arrSheet( 25,2  )
	obHeader.ItemText "RIGN", arrSheet( 25,2  )
	
	obHeader.ItemText "BS", arrSheet( 26,2  )
	obHeader.ItemText "CASB", arrSheet( 27,2  )
	obHeader.ItemText "CASL", arrSheet( 28,2  )
	obHeader.ItemText "CASX", arrSheet( 29,2  )
	obHeader.ItemText "CASD", arrSheet( 30,2  )






















'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
' Subfunction Library|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'==========================
'this function find the ceiling of the passed variable x
Public Function Ceiling( X ,  Factor )
    ' X is the value you want to round
    ' is the multiple to which you want to round
    Ceiling = (Int(X / Factor) - (X / Factor - Int(X / Factor) > 0)) * Factor
End Function
'==========================
'==========================
' this function is a very clumsy way of interpolating the log ends, but it works
Function InterpolateLogEnd(LogInput, maximumvalNB)
			Set obLog = obBHDoc.Log(LogInput)
				' loop through the log from the bottom up and report the first non null data depth and value
				NbData = obLog.NbOfData
				i=0
				Do
				i = i + 1
				v2 = obLog.DataDepth(i)
				v3 = obLog.Data(i)
				Loop While v3 = -999.25
				v2NB = v2
				v2 = Round ( v2, 1 ) 
				AziComment = ""
				If v2NB <  maximumvalNB Then
					AziComment = "Azimuth and Tilt interpolated beyond " & v2 & " m."
				End If
				' start at this point and loop down filling the nulls with the last value
				Do 
				i=i-1
				obLog.Data (i), v3
				Loop While i > 0
End Function
'==========================
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
' Subfunction Library|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||


