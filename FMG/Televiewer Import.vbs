Set obWCAD = CreateObject("WellCAD.Application")
obWCAD.Showwindow()

strTemplatePath = "c:\Proc_TV\05_FMG_TV_LasPrep\templates\"

'Definition of the root directory
'Const RootPath = "c:\Proc_TV\05_FMG_TV_LasPrep"
RootPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)


'open wellcad doc
' Open the Borehole
'Set obBHDoc = obWCAD.OpenBorehole( RootPath & "\" & WCLFileName )
Set obBHDoc = obWCAD.GetBorehole()

'================================
'::::::::::::::::::::::::::::::::
'import optical images begin
Set objFSO = CreateObject("Scripting.FileSystemObject")
'look in the folder and loop throuigh all files
i=0
boolOBIImagesExists = False
For Each objFile In objFSO.GetFolder(RootPath & "\01 Optical Images").Files

'||||||||||||||||||||||||||||||
' looping through files begin
	'if the file has an .bmp extention then perform the code, otherwise, skip this file
  	If LCase(objFSO.GetExtensionName(objFile.Name)) = "bmp" Then
		' get the bmp file name
	    bmpFileName = objFSO.GetFileName(objFile)
		
		' import bmp file and give it the handle "obBHDoc"
		Set obBHDoc2 = obWCAD.FileImport(RootPath & "\01 Optical Images\" & bmpFileName, True, RootPath & "\BMPImport.ini")
	    boolOBIImagesExists = True		                
	    'move the image To the original doc
   
        Set obLog = obBHDoc2.Log("#1") ' copy from obBHDoc2.Log
        Set obNewLog = obBHDoc.AddLog(obLog) ' paste in obBHDoc (the wellcad file we opened earlier)
        'close obBHdoc2
        obWCAD.CloseBorehole False, 1
        'change name to rgb on first round
        
        If i=0 Then
	        ' change name
	        obNewLog.Name = "RGB"
        Else
        	' merge
        	obBHDoc.MergeLogs "RGB", "#1"
        End if
    
    	i=i+1
        
	Else
	'if las 20 folder does not exist

 	End If
'||||||||||||||||||||||||||
' looping through files end
Next
'import optical images end
'::::::::::::::::::::::::::::::::
'================================

'if obi images exist
If boolOBIImagesExists = True Then
	'Interpolate bad traces
	obBHDoc.CorrectBadTraces "RGB"
	
	'Convert to float 4 all colours
	Set obLog = obBHDoc.ConvertLogTo ("RGB", 12 , False, "c:\Proc_TV\05_FMG_TV_LasPrep\templates\ConvertLogTo.ini")
	
	'Rename to img
	Set obLog = obBHDoc.Log("RGB#1")
	obLog.Name = "IMG"
End If



'================================
'::::::::::::::::::::::::::::::::
'import optical las begin
boolOBIexists = False
'check If las20 folder exists
Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FolderExists(RootPath & "\01 Optical Images\LAS20") Then
	
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		'look in the folder and loop throuigh all files
		i=0
		For Each objFile In objFSO.GetFolder(RootPath & "\01 Optical Images\LAS20").Files
		
		'||||||||||||||||||||||||||
		' looping through files begin
			'if the file has an .las extention then perform the code, otherwise, skip this file
		  	If LCase(objFSO.GetExtensionName(objFile.Name)) = "las" Then
				' get the bmp file name
			    LasFileName = objFSO.GetFileName(objFile)
				
				' import las file and give it the handle "obBHDoc"
			    Set obBHDoc2 = obWCAD.FileImport(RootPath & "\01 Optical Images\LAS20\" & LasFileName, True, RootPath & "\AutoBulkLoad.ini")           
		
			    'move the logs to the original doc
		        Set obLog = obBHDoc2.Log("GAMMA") ' copy from obBHDoc2.Log
		        Set obNewLog = obBHDoc.AddLog(obLog) ' paste in obBHDoc (the wellcad file we opened earlier)
		
		        Set obLog = obBHDoc2.Log("INCL") ' copy from obBHDoc2.Log
		        Set obNewLog = obBHDoc.AddLog(obLog) ' paste in obBHDoc (the wellcad file we opened earlier)
		
		        Set obLog = obBHDoc2.Log("AZ") ' copy from obBHDoc2.Log
		        Set obNewLog = obBHDoc.AddLog(obLog) ' paste in obBHDoc (the wellcad file we opened earlier)
		        
		        Set obLog = obBHDoc2.Log("TMAG") ' copy from obBHDoc2.Log
		        Set obNewLog = obBHDoc.AddLog(obLog) ' paste in obBHDoc (the wellcad file we opened earlier)
		
		        'close obBHdoc2
		        obWCAD.CloseBorehole False, 1
		                
		
		        If i=0 Then
					'rename the three logs
					Set obLog = obBHDoc.Log("GAMMA")
					obLog.Name = "NG OBI"
					Set obLog = obBHDoc.Log("INCL")
					obLog.Name = "TILT OBI"
					Set obLog = obBHDoc.Log("AZ")
					obLog.Name = "AZI OBI"
					Set obLog = obBHDoc.Log("TMAG")
					obLog.Name = "TMAG OBI"
					boolOBIexists = True
					
					'Else
		        	' merge
		        	'obBHDoc.MergeLogs "NG OBI", "GAMMA"
		        	'obBHDoc.MergeLogs "TILT OBI", "INCL"
		        	'obBHDoc.MergeLogs "AZI OBI", "AZ"
		        End if
		    
		    	i=i+1
		        
			Else
			'do nothing
		 	End If
		'|||||||||||||||||||||||||
		' looping through files end
		Next

	Else
	'if las 20 folder does not exist
		Wscript.Echo "No optical LAS file was found."
	End If
'import optical las end
'::::::::::::::::::::::::::::::::
'================================


'================================
'::::::::::::::::::::::::::::::::
'import acoustic images begin
Set objFSO = CreateObject("Scripting.FileSystemObject")
'look in the folder and loop throuigh all files
boolABIImagesExists = False
i=0
For Each objFile In objFSO.GetFolder(RootPath & "\02 Acoustic Images").Files

'||||||||||||||||||||||||||||
' looping through files begin
	'if the file has an .bmp extention then perform the code, otherwise, skip this file
  	If LCase(objFSO.GetExtensionName(objFile.Name)) = "hed" Then
		' get the bmp file name
	    hedFileName = objFSO.GetFileName(objFile)
		
		' import bmp file and give it the handle "obBHDoc"
		Set obBHDoc2 = obWCAD.FileImport(RootPath & "\02 Acoustic Images\" & hedFileName, True)
	    boolABIImagesExists = True		                
	    'move the image To the original doc
   
        Set obLog = obBHDoc2.Log("Travel Time") ' copy from obBHDoc2.Log
        Set obNewLog = obBHDoc.AddLog(obLog) ' paste in obBHDoc (the wellcad file we opened earlier)
        Set obLog = obBHDoc2.Log("Amplitude") ' copy from obBHDoc2.Log
        Set obNewLog = obBHDoc.AddLog(obLog) ' paste in obBHDoc (the wellcad file we opened earlier)
        
        'close obBHdoc2
        obWCAD.CloseBorehole False, 1
        
        'change name to TT and AMP on first round, merge on succesive rounds
        If i=0 Then
			'Change name 
			Set obLog = obBHDoc.Log("Travel Time")
			obLog.Name = "TT"
			Set obLog = obBHDoc.Log("Amplitude")
			obLog.Name = "AMP"
        Else
        	' merge
		    obBHDoc.MergeLogs "TT", "Travel Time"
		    obBHDoc.MergeLogs "AMP", "Amplitude"
        End if
    
    	i=i+1
        
	Else
	'do nothing
 	End If
 	

'||||||||||||||||||||||||||
' looping through files begin
Next
'import acoustic images end
'::::::::::::::::::::::::::::::::
'================================


If boolABIImagesExists = True Then
	'Interpolate bad traces
	obBHDoc.CorrectBadTraces "TT"
	'Interpolate bad traces
	obBHDoc.CorrectBadTraces "AMP"
End If


'================================
'::::::::::::::::::::::::::::::::
'import acoustic las begin
boolABIexists = False
'check If las20 folder exists
Set objFSO = CreateObject("Scripting.FileSystemObject")
	If objFSO.FolderExists(RootPath & "\02 Acoustic Images\LAS20") Then
	
		Set objFSO = CreateObject("Scripting.FileSystemObject")
		'look in the folder and loop throuigh all files
		i=0
		For Each objFile In objFSO.GetFolder(RootPath & "\02 Acoustic Images\LAS20").Files
		
		'||||||||||||||||||||||||||
		' looping through files begin
			'if the file has an .las extention then perform the code, otherwise, skip this file
		  	If LCase(objFSO.GetExtensionName(objFile.Name)) = "las" Then
				' get the bmp file name
			    LasFileName = objFSO.GetFileName(objFile)
				
				' import bmp file and give it the handle "obBHDoc"
			    Set obBHDoc2 = obWCAD.FileImport(RootPath & "\02 Acoustic Images\LAS20\" & LasFileName, True, RootPath & "\AutoBulkLoad.ini")           
		
			    'move the three logs to the original doc
		        Set obLog = obBHDoc2.Log("GAMMA") ' copy from obBHDoc2.Log
		        Set obNewLog = obBHDoc.AddLog(obLog) ' paste in obBHDoc (the wellcad file we opened earlier)
		
		        Set obLog = obBHDoc2.Log("INCL") ' copy from obBHDoc2.Log
		        Set obNewLog = obBHDoc.AddLog(obLog) ' paste in obBHDoc (the wellcad file we opened earlier)
		
		        Set obLog = obBHDoc2.Log("AZ") ' copy from obBHDoc2.Log
		        Set obNewLog = obBHDoc.AddLog(obLog) ' paste in obBHDoc (the wellcad file we opened earlier)

				Set obLog = obBHDoc2.Log("TMAG") ' copy from obBHDoc2.Log
		        Set obNewLog = obBHDoc.AddLog(obLog) ' paste in obBHDoc (the wellcad file we opened earlier)
		
		        'close obBHdoc2
		        obWCAD.CloseBorehole False, 1
		                
		
		        If i=0 Then
					'rename the three logs
					Set obLog = obBHDoc.Log("GAMMA")
					obLog.Name = "NG ABI"
					Set obLog = obBHDoc.Log("INCL")
					obLog.Name = "TILT ABI"
					Set obLog = obBHDoc.Log("AZ")
					obLog.Name = "AZI ABI"
					boolABIexists = True
					Set obLog = obBHDoc.Log("TMAG")
					obLog.Name = "TMAG ABI"
					
					'Else
		        	' merge
		        	'obBHDoc.MergeLogs "NG ABI", "GAMMA"
		        	'obBHDoc.MergeLogs "TILT ABI", "INCL"
		        	'obBHDoc.MergeLogs "AZI ABI", "AZ"
		        End if
		    
		    	i=i+1
		        
			Else
			'do nothing
		 	End If
		'||||||||||||||||||||||||||
		' looping through files end
		Next

	Else
	'if las20 folder does not exist
		Wscript.Echo "No Acoustic Las data found. check folder names and export locations?"
	End If
'import acoustic las end
'::::::::::::::::::::::::::::::::
'================================






'set OTV or ATV devi info to tilt and azi
'This script applies to RTIO naming conventions for logs, we will leave the Tilt and Azi from the OPTV and BHTV tools as they are named in the previous script
'Therefore this set of scripts is unused.

'if otv then use otv
'If boolOBIexists = True Then
	'Set obLog = obBHDoc.Log("TILT OBI")
	'obLog.Name = "TILT"	
	'Set obLog = obBHDoc.Log("AZI OBI")
	'obLog.Name = "AZI"
'Else 'if there is no otv check for atv and rename the ABI tilt and azi
	'If boolABIexists = True Then
		'Set obLog = obBHDoc.Log("TILT ABI")
		'obLog.Name = "TILT"	
		'Set obLog = obBHDoc.Log("AZI ABI")
		'obLog.Name = "AZI"
		
	'Else
		'Wscript.Echo "No Acoustic or Optical azimuth and tilt logs were found?! check folder names and export locations?"
	'End if
'End If



' ==== Apply Template ====

'apply template when there is otv and atv
If boolOBIexists = True And boolABIexists = True Then
	'both
	obBHDoc.ApplyTemplate strTemplatePath & "FMG_Image_OPTV_BHTV_Gamma_TIL_AZI_TMAG_Match.wdt", false, false, false, false, false
End If

'apply template when there is atv only
If boolOBIexists = False And boolABIexists = True Then
	'atv only
	obBHDoc.ApplyTemplate strTemplatePath & "FMG_Image_BHTV_Gamma_TIL_AZI_TMAG_Match.wdt", false, false, false, false, false
End If

'apply template when there is otv only
If boolOBIexists = True And boolABIexists = False Then
	'otv only
	obBHDoc.ApplyTemplate strTemplatePath & "FMG_Image_OPTV_Gamma_TIL_AZI_TMAG_Match", false, false, false, false, false
End if



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
'==========================0, "Select the acoustic data To load ", &H4000, "c:\Proc_TV\05_RTIO_TV_LasPrepv2\RC16BS4B0034\02 Acoustic Images")
Function BrowseForFile(bffRootPath)
     With CreateObject("WScript.Shell")
         Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
         Dim tempFolder : Set tempFolder = fso.GetSpecialFolder(2)
         Dim tempName : tempName = fso.GetTempName() & ".hta"
         Dim path : path = "HKCU\Volatile Environment\MsgResp"
         With tempFolder.CreateTextFile(tempName)
            .Write "<input type=file name=f>" & _
                 "<script>f.click();(new ActiveXObject('WScript.Shell'))" & _
                 ".RegWrite('HKCU\\Volatile Environment\\MsgResp', f.value);" & _
                 "close();</script>"
            .Close
         End With
        .Run tempFolder & "\" & tempName, 1, True
        BrowseForFile = .RegRead(path)
        .RegDelete path
        fso.DeleteFile tempFolder & "\" & tempName
     End With
End Function
 
'MsgBox BrowseForFile


'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
' Subfunction Library|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||
'|||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||||


