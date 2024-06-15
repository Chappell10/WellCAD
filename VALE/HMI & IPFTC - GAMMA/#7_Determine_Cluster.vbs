Set obWCAD = CreateObject("WellCAD.Application")
obWCAD.Showwindow()

Set obBHDoc = obWCAD.GetActiveBorehole()

Set obLogRt32 = obBHDoc.Log("Rt32")
obLogRt32.Name = "Rt32"

Set obLogIPLIN646 = obBHDoc.Log("IPLIN646")
obLogIPLIN646.Name = "IPLIN646"

Set obLogNaturalGamma = obBHDoc.Log("HMI GR")
obLogNaturalGamma.Name = "HMI GR"

Set obLogCorrectedMagSus = obBHDoc.Log("MagSus 25C")
obLogCorrectedMagSus.Name = "MagSus 25C"

Set obLogTCPU = obBHDoc.Log("HMI TCPU")
obLogTCPU.Name = "HMI TCPU"


Function DetermineCluster(Rt32, IPLIN646, Gamma, Mag, TCPU)


    Dim MinRange()
    ReDim MinRange(29, 3)
    MinRange(0, 0) = 494.33
    MinRange(0, 1) = 1.5205
    MinRange(0, 2) = 0.2143
    MinRange(0, 3) = 0.6292

    MinRange(1, 0) = 1.9348e+05
    MinRange(1, 1) = 2.1451
    MinRange(1, 2) = 3.6285
    MinRange(1, 3) =0.063

    MinRange(2, 0) = 193.16
    MinRange(2, 1) = 1.3872
    MinRange(2, 2) = 0.4868
    MinRange(2, 3) = 0.0009

    MinRange(3, 0) = 37613
    MinRange(3, 1) = 6.4812
    MinRange(3, 2) = 4.2988 
    MinRange(3, 3) = 0.0032

    MinRange(4, 0) = 2.6976      
    MinRange(4, 1) = 20.738       
    MinRange(4, 2) = 293.9
    MinRange(4, 3) = 0.1329

    MinRange(5, 0) = 3048
    MinRange(5, 1) = 0.042
    MinRange(5, 2) = 0.0667
    MinRange(5, 3) = 0.473

    MinRange(6, 0) = 1.5523e+06
    MinRange(6, 1) = 0.1855
    MinRange(6, 2) = 7.3515
    MinRange(6, 3) = 12.187

    MinRange(7, 0) = 0.0096001
    MinRange(7, 1) = 17743
    MinRange(7, 2) = 4.2037
    MinRange(7, 3) = 0.0749

    MinRange(8, 0) = 5.8742
    MinRange(8, 1) = 257.41
    MinRange(8, 2) = 197.49
    MinRange(8, 3) = 0.0016

    MinRange(9, 0) = 745.95
    MinRange(9, 1) = 0.7978
    MinRange(9, 2) = 0.4911
    MinRange(9, 3) = 3.4116

    MinRange(10, 0) = 0.021136
    MinRange(10, 1) = 0.6685
    MinRange(10, 2) = 1.6053
    MinRange(10, 3) = 0.0056

    MinRange(11, 0) = 0.2108
    MinRange(11, 1) = 250.34
    MinRange(11, 2) = 90.653
    MinRange(11, 3) = 0.0071

    MinRange(12, 0) = 320.08
    MinRange(12, 1) = 2.0833
    MinRange(12, 2) = 0.7651
    MinRange(12, 3) = 0.0337

    MinRange(13, 0) = 0.00032269
    MinRange(13, 1) = 2191.4
    MinRange(13, 2) = 1.7224
    MinRange(13, 3) = 0.0403

    MinRange(14, 0) = 205.02
    MinRange(14, 1) = 218.13
    MinRange(14, 2) = 0.0918
    MinRange(14, 3) = 0.0108

    MinRange(15, 0) = 7.1946e+05
    MinRange(15, 1) = 0.0489
    MinRange(15, 2) = 0.5857
    MinRange(15, 3) = 7.9418

    MinRange(16, 0) = 8124.3
    MinRange(16, 1) = 5.2396
    MinRange(16, 2) = 0.0966
    MinRange(16, 3) = 0.1502

    MinRange(17, 0) = 50.888
    MinRange(17, 1) = 11.872
    MinRange(17, 2) = 0.5536
    MinRange(17, 3) = 0.5627

    MinRange(18, 0) = 0.33189
    MinRange(18, 1) = 5.8535
    MinRange(18, 2) = 109.09
    MinRange(15, 3) = 0.0561

    MinRange(19, 0) = 446.68
    MinRange(19, 1) = 2.9297
    MinRange(19, 2) = 155.41
    MinRange(19, 3) = 2.6578

    MinRange(20, 0) = 0.9223
    MinRange(20, 1) = 2.2266e+05
    MinRange(20, 2) = 14.972
    MinRange(20, 3) = 16.828

    MinRange(21, 0) = 0.00024202
    MinRange(21, 1) = 339.54
    MinRange(21, 2) = 0.1874
    MinRange(21, 3) = 0.0014

    MinRange(22, 0) = 0.0015328
    MinRange(22, 1) = 159.85
    MinRange(22, 2) = 0.2787
    MinRange(22, 3) = 0.0136

    MinRange(23, 0) = 0.00016135
    MinRange(23, 1) = 250.18
    MinRange(23, 2) = 0.2431
    MinRange(23, 3) = 0.0053

    MinRange(24, 0) = 0.00080673
    MinRange(24, 1) = 0.28 
    MinRange(24, 2) = 0.874 
    MinRange(24, 3) = 565.2

    MinRange(25, 0) = 4.0618e+06
    MinRange(25, 1) = 5.2915
    MinRange(25, 2) = 0.0432
    MinRange(25, 3) = 0.8056

    MinRange(26, 0) = 8.0673e-05
    MinRange(26, 1) = 433.44
    MinRange(26, 2) = 0.4463
    MinRange(26, 3) = 0.0005

    MinRange(27, 0) = 1052.6
    MinRange(27, 1) = 0.868
    MinRange(27, 2) = 0.2151
    MinRange(27, 3) = 6.4713

    MinRange(28, 0) = 1729.9
    MinRange(28, 1) = 0.0682
    MinRange(28, 2) = 0.1556
    MinRange(28, 3) = 2.9626

    MinRange(29, 0) = 0.00024202
    MinRange(29, 1) = 574
    MinRange(29, 2) = 1.1018
    MinRange(29, 3) = 0.0227



    Dim MaxRange()
    ReDim MaxRange(29, 3)

    MaxRange(0, 0) = 790.35
    MaxRange(0, 1) = 601.91
    MaxRange(0, 2) = 246.89
    MaxRange(0, 3) = 327.95

    MaxRange(1, 0) = 7.1925E+05
    MaxRange(1, 1) = 1.5289E+05
    MaxRange(1, 2) = 780.27
    MaxRange(1, 3) = 1421.9

    MaxRange(2, 0) = 374.83
    MaxRange(2, 1) = 289.51
    MaxRange(2, 2) = 318.63
    MaxRange(2, 3) = 374.38

    MaxRange(3, 0) = 1.9345E+05
    MaxRange(3, 1) = 4564.8
    MaxRange(3, 2) = 1138.4
    MaxRange(3, 3) = 1422

    MaxRange(4, 0) = 398.92
    MaxRange(4, 1) = 540.44
    MaxRange(4, 2) = 2031.6
    MaxRange(4, 3) = 359.05

    MaxRange(5, 0) = 8124.2
    MaxRange(5, 1) = 4725.6
    MaxRange(5, 2) = 1021.7
    MaxRange(5, 3) = 1420.5

    MaxRange(6, 0) = 4.0165E+06
    MaxRange(6, 1) = 67283
    MaxRange(6, 2) = 547.48
    MaxRange(6, 3) = 292.18

    MaxRange(7, 0) = 18482
    MaxRange(7, 1) = 2.0554E+05
    MaxRange(7, 2) = 594.8
    MaxRange(7, 3) = 1597.6

    MaxRange(8, 0) = 272.97
    MaxRange(8, 1) = 622.38
    MaxRange(8, 2) = 562.67
    MaxRange(8, 3) = 133.88

    MaxRange(9, 0) = 1121.7
    MaxRange(9, 1) = 1194.6
    MaxRange(9, 2) = 419.65
    MaxRange(9, 3) = 661.9

    MaxRange(10, 0) = 199.31
    MaxRange(10, 1) = 175.71
    MaxRange(10, 2) = 221.38
    MaxRange(10, 3) = 652.1

    MaxRange(11, 0) = 226.11
    MaxRange(11, 1) = 461.54
    MaxRange(11, 2) = 240.69
    MaxRange(11, 3) = 220.58

    MaxRange(12, 0) = 551.96
    MaxRange(12, 1) = 254.84
    MaxRange(12, 2) = 298.34
    MaxRange(12, 3) = 337.81

    MaxRange(13, 0) = 6538.6
    MaxRange(13, 1) = 17716
    MaxRange(13, 2) = 509.88
    MaxRange(13, 3) = 1602.3

    MaxRange(14, 0) = 1045.5
    MaxRange(14, 1) = 1284.8
    MaxRange(14, 2) = 405.56
    MaxRange(14, 3) = 506.17

    MaxRange(15, 0) = 1.5522E+06
    MaxRange(15, 1) = 6828.8
    MaxRange(15, 2) = 763.4
    MaxRange(15, 3) = 241.02

    MaxRange(16, 0) = 37612
    MaxRange(16, 1) = 983
    MaxRange(16, 2) = 1304.8
    MaxRange(16, 3) = 458.94

    MaxRange(17, 0) = 269.37
    MaxRange(17, 1) = 325.37
    MaxRange(17, 2) = 156.12
    MaxRange(17, 3) = 618.85

    MaxRange(18, 0) = 372.14
    MaxRange(18, 1) = 364.61
    MaxRange(18, 2) = 373.44
    MaxRange(18, 3) = 319.93

    MaxRange(19, 0) = 1133
    MaxRange(19, 1) = 1190.9
    MaxRange(19, 2) = 1399.9
    MaxRange(19, 3) = 144.6

    MaxRange(20, 0) = 2817.4
    MaxRange(20, 1) = 8.8411E+05
    MaxRange(20, 2) = 131.63
    MaxRange(20, 3) = 1422.5

    MaxRange(21, 0) = 206.24
    MaxRange(21, 1) = 482.79
    MaxRange(21, 2) = 127.67
    MaxRange(21, 3) = 690.6

    MaxRange(22, 0) = 183.11
    MaxRange(22, 1) = 314.71
    MaxRange(22, 2) = 157.9
    MaxRange(22, 3) = 656.73

    MaxRange(23, 0) = 211.11
    MaxRange(23, 1) = 384.18
    MaxRange(23, 2) = 135.22
    MaxRange(23, 3) = 522.01

    MaxRange(24, 0) = 1758.6
    MaxRange(24, 1) = 2146
    MaxRange(24, 2) = 481
    MaxRange(24, 3) = 3466.1

    MaxRange(25, 0) = 1.8057E+07
    MaxRange(25, 1) = 123.67
    MaxRange(25, 2) = 428.36
    MaxRange(25, 3) = 273.67

    MaxRange(26, 0) = 253.54
    MaxRange(26, 1) = 619.23
    MaxRange(26, 2) = 279.33
    MaxRange(26, 3) = 724.48

    MaxRange(27, 0) = 1754.7
    MaxRange(27, 1) = 2160
    MaxRange(27, 2) = 663.2
    MaxRange(27, 3) = 1147.2

    MaxRange(28, 0) = 3049.3
    MaxRange(28, 1) = 2578.2
    MaxRange(28, 2) = 866.61
    MaxRange(28, 3) = 1599.1

    MaxRange(29, 0) = 1327
    MaxRange(29, 1) = 2536.9
    MaxRange(29, 2) = 437.22
    MaxRange(29, 3) = 1431


    Dim Cluster
    Cluster = 0


    For c = 0 To 29
        If Rt32 >= MinRange(c, 0) And Rt32 <= MaxRange(c, 0) And _
        IPLIN646 >= MinRange(c, 1) And IPLIN646 <= MaxRange(c, 1) And _
        Gamma >= MinRange(c, 2) And Gamma <= MaxRange(c, 2) And _
        Mag >= MinRange(c, 3) And Mag <= MaxRange(c, 3) Then
            Cluster = c
            Exit For
        End If
    Next 

    DetermineCluster = Cluster
End Function


Rt32_Data = obLogRt32.DataTable
IPLIN646_Data = obLogIPLIN646.DataTable
Gamma_Data = obLogNaturalGamma.DataTable
Mag_Data = obLogCorrectedMagSus.DataTable
TCPU_Data = obLogTCPU.DataTable


Dim PPZonationData() 
nbRows = obLogRt32.NbOfData 
ReDim PPZonationData(nbRows, 2)

fTop = obLogRt32.TopDepth
fBot = obLogRt32.BottomDepth

'MsgBox "Top Depth " & fTop 
'MsgBox "Bot Depth " & fBot

For i = 1 To nbRows

Rt32 = obLogRt32.DataAtDepth(fTop+(i-1)*.1)
IPLIN646 = obLogIPLIN646.DataAtDepth(fTop+(i-1)*.1)
Gamma = obLogNaturalGamma.DataAtDepth(fTop+(i-1)*.1)
Mag = obLogCorrectedMagSus.DataAtDepth(fTop+(i-1)*.1)
TCPU = obLogTCPU.DataAtDepth(fTop+(i-1)*.1)

    'Rt32 = Rt32_Data(i, 1)
    'IPLIN646 = IPLIN646_Data(i, 1)
    'Gamma = Gamma_Data(i, 1)
    'Mag = Mag_Data(i, 1)
    'TCPU = TCPU_Data(i, 1)

   
    PPZonationData(i, 0) = Rt32_Data(i, 0)

    
    PPZonationData(i, 2) = DetermineCluster(Rt32, IPLIN646, Gamma, Mag, TCPU)
    PPZonationData(i, 2) = "Zone " & PPZonationData(i, 2)

    
    If i < nbRows Then
        PPZonationData(i + 1, 1) = Rt32_Data(i, 0)
    End If
Next


PPZonationData(1, 1) = PPZonationData(1, 0) 


Set obLogPPZonation = obBHDoc.InsertNewLog(7)
obLogPPZonation.Name = "PP Zonation"
obLogPPZonation.DataTable = PPZonationData