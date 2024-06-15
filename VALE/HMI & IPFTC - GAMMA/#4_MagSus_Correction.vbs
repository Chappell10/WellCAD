Set obWCAD = CreateObject("WellCAD.Application")
obWCAD.Showwindow()

Set obBHDoc = obWCAD.GetActiveBorehole()

Set obLog_HMI_TCPU = obBHDoc.Log("HMI TCPU")
obLog_HMI_TCPU.Name = "HMI TCPU"

Set obLogMagSus = obBHDoc.Log("MagSus")
obLogMagSus.Name = "MagSus"

Function CorrectedMagSus(MagSus, HMI_TCPU)
    If IsNumeric(MagSus) And IsNumeric(HMI_TCPU) And (HMI_TCPU>0)And (MagSus>0) Then
        Dim delta: delta = 25 - HMI_TCPU
        If delta >= 0 Then
            CorrectedMagSus = MagSus - delta ^ 1.15
        Else
            delta=HMI_TCPU-25
            CorrectedMagSus = MagSus + delta ^ 1.15
        End If
    Else
        
        CorrectedMagSus = MagSus ' 
    End If
End Function

HMI_TCPU_Data = obLog_HMI_TCPU.DataTable
MagSusData = obLogMagSus.DataTable

Dim CorrectedMagSusData() 
nbRows = obLogMagSus.NbOfData 
Redim CorrectedMagSusData(nbRows, 1)

For i = LBound(MagSusData, 1) + 1 To UBound(MagSusData, 1)
    MagSus = MagSusData(i, 1)
    HMI_TCPU = HMI_TCPU_Data(i, 1)
    CorrectedMagSusData(i, 0) = MagSusData(i, 0) 
    CorrectedMagSusData(i, 1) = CorrectedMagSus(MagSus, HMI_TCPU)
Next

Set obLogCorrectedMagSus = obBHDoc.InsertNewLog(1)
obLogCorrectedMagSus.Name = "MagSus 25C"
obLogCorrectedMagSus.DataTable = CorrectedMagSusData
obLogCorrectedMagSus.LogUnit ="10e-3 SI units"

'obLogMagSus.HideLogTitle = True
'obLogMagSus.HideLogData = True
