

Set obWCAD = CreateObject("WellCAD.Application")
obWCAD.Showwindow()


RootPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)


Set obBHDoc = obWCAD.GetBorehole()


Set objFSO = CreateObject("Scripting.FileSystemObject")
'--------------------------------------------------------------------------------------------------------------------------------------------'



'--------------------------------------------------------------------------------------------------------------------------------------------'


lith_CSV_Exists = False

For Each objFile In objFSO.GetFolder(RootPath & "\Lith").Files
	
  	If LCase(objFSO.GetExtensionName(objFile.Name)) = "csv" Then 
		
	    csvFileName = objFSO.GetFileName(objFile) ' Get the csv file name
		
		Set obBHDoc4 = obWCAD.FileImport(RootPath & "\Lith\" & csvFileName) 
        
        Lith_CSV_Exists = True

        Set obLogGeoLog = obBHDoc4.Log("Lith") 
        Set obNewLogGeoLog = obBHDoc.AddLog(obLogGeoLog)
        obNewLogGeoLog.Name = "Geolog"

        obWCAD.CloseBorehole False, 1 

	Else

 	End If

Next        


'--------------------------------------------------------------------------------------------------------------------------------------------'



'--------------------------------------------------------------------------------------------------------------------------------------------'


Dip_CSV_Exists = False

For Each objFile In objFSO.GetFolder(RootPath & "\Dip").Files
	
  	If LCase(objFSO.GetExtensionName(objFile.Name)) = "csv" Then 
		
	    csvFileName = objFSO.GetFileName(objFile) ' Get the csv file name
		
		Set obBHDoc4 = obWCAD.FileImport(RootPath & "\Dip\" & csvFileName) 
        
        Dip_CSV_Exists = True

        Set obLogDip = obBHDoc4.Log("Dip") 
        Set obNewLogDip = obBHDoc.AddLog(obLogDip)

        obWCAD.CloseBorehole False, 1 

	Else

 	End If

Next        

If Dip_CSV_Exists = True Then 
	
	Set obLogDip = obBHDoc.Log("Dip") 

	obLogDip.Name = "Dip"

End If

'--------------------------------------------------------------------------------------------------------------------------------------------'



'--------------------------------------------------------------------------------------------------------------------------------------------'


Assays_CSV_Exists = False

'For Each objFile In objFSO.GetFolder(RootPath & "\Assays").Files
	
  	'If LCase(objFSO.GetExtensionName(objFile.Name)) = "csv" Then 
		
	    'csvFileName = objFSO.GetFileName(objFile) ' Get the csv file name
		
		'Set obBHDoc5 = obWCAD.FileImport(RootPath & "\Assays\" & csvFileName, RootPath & "\Scripts\Import_Assays.ini") 

        'Assays_CSV_Exists = True

        'Set obLogEST = obBHDoc5.Log("EST") 
        'Set obNewLogEST = obBHDoc.AddLog(obLogEST)

        'Set obLogCU = obBHDoc5.Log("CU") 
        'Set obNewLogCU = obBHDoc.AddLog(obLogCU)

        'Set obLogNI = obBHDoc5.Log("NI") 
        'Set obNewLogNI = obBHDoc.AddLog(obLogNI)

        'Set obLogCO = obBHDoc5.Log("CO") 
        'Set obNewLogCO = obBHDoc.AddLog(obLogCO)

        'Set obLogAS = obBHDoc5.Log("AS") 
        'Set obNewLogAS = obBHDoc.AddLog(obLogAS)
        
        'Set obLogS = obBHDoc5.Log("S") 
        'Set obNewLogS = obBHDoc.AddLog(obLogS)

        'Set obLogFE = obBHDoc5.Log("FE") 
        'Set obNewLogFE = obBHDoc.AddLog(obLogFE)

        'Set obLogPB = obBHDoc5.Log("PB") 
        'Set obNewLogPB = obBHDoc.AddLog(obLogPB)

        'Set obLogZN = obBHDoc5.Log("ZN") 
        'Set obNewLogZN = obBHDoc.AddLog(obLogZN)

        'Set obLogPT = obBHDoc5.Log("PT") 
        'Set obNewLogPT = obBHDoc.AddLog(obLogPT)

        'Set obLogPD = obBHDoc5.Log("PD") 
        'Set obNewLogPD = obBHDoc.AddLog(obLogPD)

        'Set obLogAU = obBHDoc5.Log("AU") 
        'Set obNewLogAU = obBHDoc.AddLog(obLogAU)

        'Set obLogTPM = obBHDoc5.Log("TPM") 
        'Set obNewLogTPM = obBHDoc.AddLog(obLogTPM)

        'Set obLogRH = obBHDoc5.Log("RH") 
        'Set obNewLogRH = obBHDoc.AddLog(obLogRH)

        'Set obLogAG = obBHDoc5.Log("AG") 
        'Set obNewLogAG = obBHDoc.AddLog(obLogAG)

        'Set obLogSG = obBHDoc5.Log("SG") 
        'Set obNewLogSG = obBHDoc.AddLog(obLogSG)

        'obWCAD.CloseBorehole False, 1 
    
	'Else

 	'End If

'Next  



'--------------------------------------------------------------------------------------------------------------------------------------------'



'--------------------------------------------------------------------------------------------------------------------------------------------'
