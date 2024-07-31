'

Set obWCAD = CreateObject("WellCAD.Application")

obWCAD.Showwindow()

Set obBHDoc = obWCAD.NewBorehole

RootPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName) 'Run Script in the same folder as the files

Set objFSO = CreateObject("Scripting.FileSystemObject")


'--------------------------------------------------------------------------------------------------------------------------------------------'



'--------------------------------------------------------------------------------------------------------------------------------------------'

HMI_Exists = False

For Each objFile In objFSO.GetFolder(RootPath & "\HMI").Files
	
  	If LCase(objFSO.GetExtensionName(objFile.Name)) = "tfd" Then
		
	    tfdFileName = objFSO.GetFileName(objFile) 
		
		Set obBHDoc2 = obWCAD.FileImport(RootPath & "\HMI\" & tfdFileName)
        
        HMI_Exists = True

        Set obLog_HMI_GR = obBHDoc2.Log("GR") 
        Set obNewLog_HMI_GR = obBHDoc.AddLog(obLog_HMI_GR) 
        obNewLog_HMI_GR.Name = "HMI GR"

        Set obLogMagSus = obBHDoc2.Log("MagSus") 
        Set obNewLogMagSus = obBHDoc.AddLog(obLogMagSus) 

        Set obLogConductivity = obBHDoc2.Log("Conductivity")
        Set obNewLogConductivity = obBHDoc.AddLog(obLogConductivity) 

        'Set obLogResistivity = obBHDoc2.Log("Resistivity") 
        'Set obNewLogResistivity = obBHDoc.AddLog(obLogResistivity) 

        Set obLogSpeed = obBHDoc2.Log("Speed") 
        Set obNewLogSpeed = obBHDoc.AddLog(obLogSpeed) 
        obNewLogSpeed.Name = "HMI Speed"

        Set obLogTCPU = obBHDoc2.Log("TCPU") 
        Set obNewLogTCPU = obBHDoc.AddLog(obLogTCPU)
        obNewLogTCPU.Name = "HMI TCPU"


        obWCAD.CloseBorehole False, 1 
    
	Else

 	End If

Next

	



IPFTC_Exists = False

For Each objFile In objFSO.GetFolder(RootPath & "\IPFTC").Files
	
  	If LCase(objFSO.GetExtensionName(objFile.Name)) = "tfd" Then 
		
	    tfdFileName = objFSO.GetFileName(objFile) 
		
		Set obBHDoc3 = obWCAD.FileImport(RootPath & "\IPFTC\" & tfdFileName) 
        
        IPFTC_Exists = True

        Set obLogCond = obBHDoc3.Log("Cond") 
        Set obNewLogCond = obBHDoc.AddLog(obLogCond)
        obNewLogCond.Name = "Fluid Cond"

        Set obLogCond25C = obBHDoc3.Log("Cond25C") 
        Set obNewLogCond25C = obBHDoc.AddLog(obLogCond25C)
        obNewLogCond25C.Name = "Fluid Cond 25C"

        'Set obLogIPlin161 = obBHDoc3.Log("IPlin161") 
        'Set obNewLogIPlin161 = obBHDoc.AddLog(obLogIPlin161) 

        'Set obLogIPlin162 = obBHDoc3.Log("IPlin162") 
        'Set obNewLogIPlin162 = obBHDoc.AddLog(obLogIPlin162) 

        'Set obLogIPlin164 = obBHDoc3.Log("IPlin164") 
        'Set obNewLogIPlin164 = obBHDoc.AddLog(obLogIPlin164) 

        Set obLogIPlin166 = obBHDoc3.Log("IPlin166") 
        Set obNewLogIPlin166 = obBHDoc.AddLog(obLogIPlin166) 

        Set obLogIPlin168 = obBHDoc3.Log("IPlin168") 
        Set obNewLogIPlin168 = obBHDoc.AddLog(obLogIPlin168)

        Set obLogIPlin641 = obBHDoc3.Log("IPlin641") 
        Set obNewLogIPlin641 = obBHDoc.AddLog(obLogIPlin641)

        Set obLogIPlin642 = obBHDoc3.Log("IPlin642") 
        Set obNewLogIPlin642 = obBHDoc.AddLog(obLogIPlin642)

        'Set obLogIPlin644 = obBHDoc3.Log("IPlin644") 
        'Set obNewLogIPlin644 = obBHDoc.AddLog(obLogIPlin644)

        Set obLogIPlin646 = obBHDoc3.Log("IPlin646") 
        Set obNewLogIPlin646 = obBHDoc.AddLog(obLogIPlin646)

        'Set obLogIPlin648 = obBHDoc3.Log("IPlin648") 
        'Set obNewLogIPlin648 = obBHDoc.AddLog(obLogIPlin648)

        Set obLogN8 = obBHDoc3.Log("N8") 
        Set obNewLogN8 = obBHDoc.AddLog(obLogN8) 

        Set obLogN16 = obBHDoc3.Log("N16") 
        Set obNewLogN16 = obBHDoc.AddLog(obLogN16)

        Set obLogN32 = obBHDoc3.Log("N32") 
        Set obNewLogN32 = obBHDoc.AddLog(obLogN32)

        Set obLogN64 = obBHDoc3.Log("N64") 
        Set obNewLogN64 = obBHDoc.AddLog(obLogN64)
        
        Set obLogSPR = obBHDoc3.Log("SPR") 
        Set obNewLogSPR = obBHDoc.AddLog(obLogSPR)
        obNewLogSPR.HideLogData = True
        obNewLogSPR.HideLogTitle = True

        Set obLog_FTC_Temp = obBHDoc3.Log("Temp") 
        Set obNewLog_FTC_Temp = obBHDoc.AddLog(obLog_FTC_Temp)
        obNewLog_FTC_Temp.Name = "Fluid Temp"

        Set obLog_IPFTC_Speed= obBHDoc3.Log("Speed") 
        Set obNewLog_IPFTC_Speed = obBHDoc.AddLog(obLog_IPFTC_Speed)
        obNewLog_IPFTC_Speed.Name = "FTC Speed"

        'Set obLogVinj64 = obBHDoc3.Log("Vinj64") 
        'Set obNewLogVinj64 = obBHDoc.AddLog(obLogVinj64)

        'Set obLogVSP = obBHDoc3.Log("VSP") 
        'Set obNewLogVSP = obBHDoc.AddLog(obLogVSP)

        'Set obLogVSPR = obBHDoc3.Log("VSPR") 
        'Set obNewLogVSPR = obBHDoc.AddLog(obLogVSPR)

        
        If IsObject(obBHDoc3.Log("GR") ) Then
            Set obLog_IPFTC_GR = obBHDoc3.Log("GR") 
            Set obNewLog_IPFTC_GR = obBHDoc.AddLog(obLog_IPFTC_GR)
            obNewLog_IPFTC_GR.Name = "IPFTC GR"
        End If

        obWCAD.CloseBorehole False, 1 
    
	Else

 	End If

Next

Set obLogOne = obBHDoc.InsertNewLog(2)
obLogOne.Name = "Adjusted Fluid Cond"
obLogOne.Formula = "{Fluid Cond}/10000"

obBHDoc.RemoveLog "Fluid Cond"

Set obLogTwo = obBHDoc.InsertNewLog(2)
obLogTwo.Name = "Fluid Cond"
obLogTwo.Formula = "If({Adjusted Fluid Cond} < 0, 0, {Adjusted Fluid Cond})"
obLogTwo.LogUnit = "S/m"

obBHDoc.RemoveLog "Adjusted Fluid Cond"



Set obLogThree = obBHDoc.InsertNewLog(2)
obLogThree.Name = "Adjusted Fluid Cond 25C"
obLogThree.Formula = "{Fluid Cond 25C}/10000"

obBHDoc.RemoveLog "Fluid Cond 25C"

Set obLogFour = obBHDoc.InsertNewLog(2)
obLogFour.Name = "Fluid Cond 25C"
obLogFour.Formula = "If({Adjusted Fluid Cond 25C} < 0, 0, {Adjusted Fluid Cond 25C})"
obLogFour.LogUnit = "S/m"

obBHDoc.RemoveLog "Adjusted Fluid Cond 25C"
