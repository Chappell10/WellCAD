Set obWCAD = CreateObject("WellCAD.Application")
obWCAD.Showwindow()

Set obBHDoc = obWCAD.GetActiveBorehole()

Set obLogN8 = obBHDoc.Log("N8")
obLogN8.Name = "N8"

Set obLogN16 = obBHDoc.Log("N16")
obLogN16.Name = "N16"

Set obLogN32 = obBHDoc.Log("N32")
obLogN32.Name = "N32"

Set obLogN64 = obBHDoc.Log("N64")
obLogN64.Name = "N64"

Set obLogCond25C = obBHDoc.Log("Fluid Cond 25C")
obLogCond25C.Name = "Fluid Cond 25C"




Function ResistivityCorrection(d, Ra, RmS, AM, dp) 

    d = d + 3.35 - dp
    'MsgBox "before" & RmS
    Rm=RmS*10000
   ' MsgBox "after" & Rm

    If Ra > 0 And Rm > 0 Then
        If Ra / Rm < 0.08 Then
            Rm = Ra * 0.08
        ElseIf Ra / Rm > 1000 Then
            Rm = Ra * 1000
        End If

        Dim x1, x2, x3, S
        x1 = Log(Ra / Rm)
        x2 = Log(AM / d)
        x3 = x1 * x2

        If Ra / Rm < 0.08 Then x1 = Log(0.08)

        If (Ra / Rm) < 1 Then
            S = (-1.4939863E-1 * x1 + 9.8989832E-1) * x1 + (((-2.5675792E-2 * x2 + 6.7224059E-2) * x2 + 7.6110412E-2) * x2 - 1.9797682E-1) * x2 + x3 * (-1.0115236 - 2.3026418E-2 * x1 * x1 + (x2 * -1.8076212E-1 + 8.5238666E-1) * x2) - 1.2422286E-3
        Else
            S = ((1.5270453E-2 * x1 - 6.5033900E-2) * x1 + 1.2427109) * x1 + (((-3.1720250E-3 * x2 - 2.2673233E-2) * x2 + 1.2836914E-1) * x2 - 5.6806217E-2) * x2 + x3 * (-2.4143625E-1 - 3.3741998E-3 * x1 * x1 + (2.0463816E-3 * x2 + 5.9729697E-2) * x2) - 1.3580321E-1
        End If

        ResistivityCorrection = Rm * (Exp(S))
    Else
        'ResistivityCorrection = Rm incorrect - updated JAM 20240506
        ResistivityCorrection = Ra
    End If
End Function


RaData = obLogN8.DataTable 
RmData = obLogCond25C.DataTable 


Dim Rt8Data() 
nbRows = obLogN8.NbOfData 
Redim Rt8Data(nbRows, 1)


Const d = 3 '
Const dp = 1.87 ' 

AM = 8

For i = LBound(RaData, 1) + 1 To UBound(RaData, 1)
    Ra = RaData(i, 1)
    Rm = RmData(i, 1)
    Rt8Data(i, 0) = RaData(i, 0) 
    Rt8Data(i, 1) = ResistivityCorrection(d, Ra, Rm, AM, dp) 
Next


Set obLogRt8 = obBHDoc.InsertNewLog(1)
obLogRt8.Name = "Rt8"
obLogRt8.DataTable = Rt8Data
obLogRt8.LogUnit = "ohm.m"


obLogN8.HideLogTitle = True
obLogN8.HideLogData = True

RaData = obLogN16.DataTable 

Dim Rt16Data() 
nbRows = obLogN16.NbOfData 
Redim Rt16Data(nbRows, 1)

AM = 16 

For i = LBound(RaData, 1) + 1 To UBound(RaData, 1)
    Ra = RaData(i, 1)
    Rm = RmData(i, 1)
    Rt16Data(i, 0) = RaData(i, 0) 
    Rt16Data(i, 1) = ResistivityCorrection(d, Ra, Rm, AM, dp) 
Next

Set obLogRt16 = obBHDoc.InsertNewLog(1)
obLogRt16.Name = "Rt16"
obLogRt16.DataTable = Rt16Data
obLogRt16.LogUnit = "ohm.m"

obLogN16.HideLogTitle = True
obLogN16.HideLogData = True



RaData = obLogN32.DataTable 

Dim Rt32Data() 
nbRows = obLogN32.NbOfData 
Redim Rt32Data(nbRows, 1)

AM = 32 

For i = LBound(RaData, 1) + 1 To UBound(RaData, 1)
    Ra = RaData(i, 1)
    Rm = RmData(i, 1)
    Rt32Data(i, 0) = RaData(i, 0) 
    Rt32Data(i, 1) = ResistivityCorrection(d, Ra, Rm, AM, dp) 
Next

Set obLogRt32 = obBHDoc.InsertNewLog(1)
obLogRt32.Name = "Rt32"
obLogRt32.DataTable = Rt32Data
obLogRt32.LogUnit = "ohm.m"

obLogN32.HideLogTitle = True
obLogN32.HideLogData = True


RaData = obLogN64.DataTable 

Dim Rt64Data() 
nbRows = obLogN64.NbOfData 
Redim Rt64Data(nbRows, 1)

AM = 64 

For i = LBound(RaData, 1) + 1 To UBound(RaData, 1)
    Ra = RaData(i, 1)
    Rm = RmData(i, 1)
    Rt64Data(i, 0) = RaData(i, 0) 
    Rt64Data(i, 1) = ResistivityCorrection(d, Ra, Rm, AM, dp) 
Next

Set obLogRt64 = obBHDoc.InsertNewLog(1)
obLogRt64.Name = "Rt64"
obLogRt64.DataTable = Rt64Data
obLogRt64.LogUnit = "ohm.m"

obLogN64.HideLogTitle = True
obLogN64.HideLogData = True
