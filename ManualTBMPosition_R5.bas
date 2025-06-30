Attribute VB_Name = "ManualTBMPosition_R5"
' Topic; Manual TBM Position Program
' Created By; Suben Mukem (SBM) as Survey Engineer.
' Updated; 28/06/2025
'
'Option Base 1
Const Pi As Single = 3.141592654

'----------------General Private Function----------------'

'Convert Degrees to Radian.

Private Function DegtoRad(d)

    DegtoRad = d * (Pi / 180)

 End Function

'Convert Radian to Degrees.

 Private Function RadtoDeg(r)

    RadtoDeg = r * (180 / Pi)

 End Function

'Compute Distance and Azimuth from 2 Points.

Private Function pv_DirecDistAz(EStart, NStart, EEnd, NEnd, DA)

    dE = EEnd - EStart: dN = NEnd - NStart
    Distance = Sqr(dE ^ 2 + dN ^ 2)
    
    If dN <> 0 Then Q = RadtoDeg(Atn(dE / dN))
      If dN = 0 Then
        If dE > 0 Then
          Azi = 90
        ElseIf dE < 0 Then
          Azi = 270
        Else
          Azi = False
      End If
      
    ElseIf dN > 0 Then
      If dE > 0 Then
          Azi = Q
      ElseIf dE < 0 Then
          Azi = 360 + Q
      End If
      
    ElseIf dN < 0 Then
          Azi = 180 + Q
    End If
    
    Select Case UCase$(DA)
      Case "D"
          pv_DirecDistAz = Distance
      Case "A"
          pv_DirecDistAz = Azi
    End Select

End Function 'DirecDistAz

'Convert Degrees to D.MMSS

Private Function pv_DegtoDMSStr3(deg)

        dd = Int(deg)
        mm = Int((deg - Int(deg)) * 60)
        ss = (((deg - Int(deg)) * 60) - Int((deg - Int(deg)) * 60)) * 60
        
        pv_DegtoDMSStr3 = dd + mm / 100 + Round(ss, 2) / 10000

End Function 'DegtoDMSStr3

'Compute Northing and Easting by Local Coordinate (Y, X) , Coordinate of Center and Azimuth.

Function pv_CoorYXtoNE(ECL, NCL, AZCL, Y, X, EN)

    Ei = ECL + Y * Sin(DegtoRad(AZCL)) + X * Sin(DegtoRad(90 + AZCL))
    Ni = NCL + Y * Cos(DegtoRad(AZCL)) + X * Cos(DegtoRad(90 + AZCL))
    
    Select Case UCase$(EN)
     Case "E"
             pv_CoorYXtoNE = Ei
     Case "N"
             pv_CoorYXtoNE = Ni
  End Select
  
End Function 'Coordinate Y,X to N, E

'Compute Local Coordinate (Y, X, L) by Northing and Easting and Azimuth.

Private Function pv_CoorNEtoYXL(ECL, NCL, AZCL, EA, NA, YXL)

    dE = EA - ECL: dN = NA - NCL
    Linear = Sqr(dE ^ 2 + dN ^ 2)
        
    If dN <> 0 Then Q = RadtoDeg(Atn(dE / dN))
      If dN = 0 Then
        If dE > 0 Then
          AZLinear = 90
        ElseIf dE < 0 Then
          AZLinear = 270
        Else
          AZLinear = False
      End If
      
    ElseIf dN > 0 Then
      If dE > 0 Then
          AZLinear = Q
      ElseIf dE < 0 Then
          AZLinear = 360 + Q
      End If
      
    ElseIf dN < 0 Then
          AZLinear = 180 + Q
    End If
    
        Delta = DegtoRad(AZLinear - AZCL)
        Y = Linear * Cos(Delta)
        X = Linear * Sin(Delta)
        
        Select Case UCase$(YXL)
            Case "Y"
                pv_CoorNEtoYXL = Y
            Case "X"
                pv_CoorNEtoYXL = X
            Case "L"
                pv_CoorNEtoYXL = Linear
        End Select
    
End Function 'CoorNEtoYXL

'Compute Hor.Angle, Ver.Angle, Slope Dist from 3 Points.

Private Function DirecSDHAVAby3P(EStn, NStn, ZStn, EBs, NBs, ZBs, EFs, NFs, ZFs, SHV)
    
    'STN to BS
    Az1 = pv_DirecDistAz(EStn, NStn, EBs, NBs, "A")
    HD1 = pv_DirecDistAz(EStn, NStn, EBs, NBs, "D")
    Q1 = RadtoDeg(Atn((ZBs - ZStn) / HD1))
    VA1 = 90 - Q1
    SD1 = Sqr((ZBs - ZStn) ^ 2 + (HD1) ^ 2)
    
    'STN to FS
    Az2 = pv_DirecDistAz(EStn, NStn, EFs, NFs, "A")
    HD2 = pv_DirecDistAz(EStn, NStn, EFs, NFs, "D")
    Q2 = RadtoDeg(Atn((ZFs - ZStn) / HD2))
    VA2 = 90 - Q2
    SD2 = Sqr((ZFs - ZStn) ^ 2 + (HD2) ^ 2)
    
    'Hor. Angle
    If Az2 < Az1 Then
        HA = 360 + (Az2 - Az1)
    Else
        HA = Az2 - Az1
    End If
    
    'Result STN to FS
    Select Case UCase$(SHV)
      Case "S"
          DirecSDHAVAby3P = SD2
      Case "H"
          DirecSDHAVAby3P = pv_DegtoDMSStr3(HA)
      Case "V"
          DirecSDHAVAby3P = pv_DegtoDMSStr3(VA2)
    End Select
    
End Function

'Compute Prism Offset after TBM Pitching and Roll

Private Function PrismOffset(MX, MY, MZ, RearPitch, RearRoll, XYZ)
    
    'Before TBM Pitching
    LinearXZ = Sqr(MX ^ 2 + MZ ^ 2)
    QXZ = RadtoDeg(Atn(MX / MZ)) + RearPitch
    
    'Before TBM Roll
    LinearYZ = Sqr(MY ^ 2 + MZ ^ 2)
    QYZ = RadtoDeg(Atn(MY / MZ)) + RearRoll
    
    'After TBM Pitching and Roll
    AfterX = LinearXZ * Sin(DegtoRad(QXZ))
    AfterY = LinearYZ * Sin(DegtoRad(QYZ))
    AfterZ = LinearYZ * Cos(DegtoRad(QYZ))
    
    'Result After TBM Roll
    Select Case UCase$(XYZ)
      Case "X"
          PrismOffset = AfterX
      Case "Y"
          PrismOffset = AfterY
      Case "Z"
          PrismOffset = AfterZ
    End Select

End Function

'Compute Prism Offset Before TBM Pitching and Roll
Private Function CalibratePrismOffset(AMX, AMY, AMZ, RearPitch, RearRoll, XYZ)

    'After TBM Pitching
    LinearXZ = Sqr(AMX ^ 2 + AMZ ^ 2)
    QXZ = RadtoDeg(Atn(AMX / AMZ)) - RearPitch
    
    'Before TBM Pitching
    BeforeX = LinearXZ * Sin(DegtoRad(QXZ))
    BeforeZP = LinearXZ * Cos(DegtoRad(QXZ))
    
    'After TBM Roll
    LinearYZ = Sqr(AMY ^ 2 + BeforeZP ^ 2)
    QYZ = RadtoDeg(Atn(AMY / BeforeZP)) - RearRoll
    
    'Before TBM Roll
    BeforeY = LinearYZ * Sin(DegtoRad(QYZ))
    BeforeZR = LinearYZ * Cos(DegtoRad(QYZ))
    
    'Result Before TBM Roll
    Select Case UCase$(XYZ)
      Case "X"
          CalibratePrismOffset = BeforeX
      Case "Y"
          CalibratePrismOffset = BeforeY
      Case "Z"
          CalibratePrismOffset = BeforeZR
    End Select

End Function

'Compute TBM Rear Azimuth

Private Function TBMRearAz(TgtE1, TgtN1, AX1, AY1, TgtE2, TgtN2, AX2, AY2)
    
    'Fixed Target to Free Target
    AzTgt12 = pv_DirecDistAz(TgtE1, TgtN1, TgtE2, TgtN2, "A")
    QXY = RadtoDeg(Atn((AY2 - AY1) / (AX2 - AX1)))
    
    'TBM Rear Azimuth
    TBMRearAz = AzTgt12 + QXY

End Function

'Compute Hor.&Ver. Articulation Angle

Private Function ArtAngle(LU, LD, RD, RU, KLU, KLD, KRD, KRU, HL, VL, HV)
    
    dLU = LU - KLU: dLD = LD - KLD
    dRD = RD - KRD: dRU = RU - KRU
    
    'Hor. Articulation Angle
    L1 = ((dLD - dRD) + (dLU - dRU)) / 2
    HorArt = RadtoDeg(Atn(L1 / HL))
    
    'Ver. Articulation Angle
    L2 = ((dLD - dLU) + (dRD - dRU)) / 2
    VerArt = RadtoDeg(Atn(L2 / VL))
    
    Select Case UCase$(HV)
      Case "H"
          ArtAngle = HorArt
      Case "V"
          ArtAngle = VerArt
    End Select

End Function

'Compute pitching
Private Function Pitching(ChStart, ZStrat, ChEnd, ZEnd)

    Pitching = (ZEnd - ZStrat) / (ChEnd - ChStart)

End Function

'Compute Vertical Deviation
Private Function DeviateVt(ChD, ZD, Pitching, ChA, ZA)
    
    ZFind = ZD + Pitching * (ChA - ChD)
    DeviateVt = ZA - ZFind

End Function
'-----------------End Private Function-----------------'

'-----------------Maual TBM Position Computation-----------------'

Sub TBMPosition()
    '1. Index data and Parameter
    'Alignment
    Sheets("Alignment").Visible = True 'Open Worksheets
    Sheets("Alignment").Select
    
    Dim totalDTA As Long
    totalDTA = ThisWorkbook.Sheets("Alignment").Cells(Rows.Count, 3).End(xlUp).Row - 4
    
    Range("B5").Select
    Dim PntDTA() As Variant
    Dim ChDTA() As Variant
    Dim EDTA() As Variant
    Dim NDTA() As Variant
    Dim ZDTA() As Variant
    ReDim PntDTA(totalDTA - 1)
    ReDim ChDTA(totalDTA - 1)
    ReDim EDTA(totalDTA - 1)
    ReDim NDTA(totalDTA - 1)
    ReDim ZDTA(totalDTA - 1)
    
    For d = 0 To totalDTA - 1
    
        PntDTA(d) = ActiveCell.Offset(d, 0)
        ChDTA(d) = ActiveCell.Offset(d, 1)
        EDTA(d) = ActiveCell.Offset(d, 3)
        NDTA(d) = ActiveCell.Offset(d, 2)
        ZDTA(d) = ActiveCell.Offset(d, 4)
        'Debug.Print d, PntDTA(d), ChDTA(d), EDTA(d), NDTA(d), ZDTA(d)
    Next
    
    StartCH = Range("K4")
    Direction = Range("K6")
    If Direction = "Forward" Then
        Excavate_Direc = 1
    ElseIf Direction = "Backward" Then
        Excavate_Direc = -1
    Else
        Excavate_Direc = 1 'Incase forget to input excavation direction.
    End If
    
    Sheets("Alignment").Visible = False 'Close Worksheets
    
    'Target Setting
    Sheets("Target Setting").Visible = True 'Open Worksheets
    Sheets("Target Setting").Select
    Range("C5").Select
    
    Dim totalTgt As Long
    totalTgt = ThisWorkbook.Sheets("Target Setting").Cells(Rows.Count, 3).End(xlUp).Row - 4
    
    Dim TgtName() As Variant
    Dim MX() As Variant
    Dim MY() As Variant
    Dim MZ() As Variant
    ReDim TgtName(totalTgt - 1)
    ReDim MX(totalTgt - 1)
    ReDim MY(totalTgt - 1)
    ReDim MZ(totalTgt - 1)
    
    For t = 0 To totalTgt - 1

        TgtName(t) = ActiveCell.Offset(t, 0)
        MX(t) = ActiveCell.Offset(t, 1)
        MY(t) = ActiveCell.Offset(t, 2)
        MZ(t) = ActiveCell.Offset(t, 3)

    Next
    
    Sheets("Target Setting").Visible = False 'Close Worksheets
    
    'TBM Parameter
    Sheets("TBM Parameter").Visible = True 'Open Worksheets
    Sheets("TBM Parameter").Select
    FrontLength = Range("F4") / 1000
    RearLength = Range("F5") / 1000
    TotalTBMLength = Range("F6") / 1000
    
    HorArtLength = Range("F10")
    VerArtLength = Range("F11")
    KLU = Range("F12")
    KLD = Range("F13")
    KRD = Range("F14")
    KRU = Range("F15")
    
    Sheets("TBM Parameter").Visible = False 'Close Worksheets
    
    'Main Program - Navigation
    Sheets("Main Pro.").Select
    
    Dim TS(0 To 3) As Variant
    TS(0) = Range("G13") 'Name
    TS(1) = Range("G14") 'Northing
    TS(2) = Range("G15") 'Easting
    TS(3) = Range("G16") 'Elevation
    
    Dim BS(0 To 3) As Variant
    BS(0) = Range("I13") 'Name
    BS(1) = Range("I14") 'Northing
    BS(2) = Range("I15") 'Easting
    BS(3) = Range("I16") 'Elevation
    
    Dim TgtA(0 To 3) As Variant
    TgtA(0) = Range("K13") 'Name
    TgtA(1) = Range("K14") 'Northing
    TgtA(2) = Range("K15") 'Easting
    TgtA(3) = Range("K16") 'Elevation
    
    Dim TgtB(0 To 3) As Variant
    TgtB(0) = Range("M13") 'Name
    TgtB(1) = Range("M14") 'Northing
    TgtB(2) = Range("M15") 'Easting
    TgtB(3) = Range("M16") 'Elevation
    
    'Articulation Jack Stoke
    LU = Range("G22")
    LD = Range("I22")
    RD = Range("K22")
    RU = Range("M22")
    
    'Pitching and Roll of TBM Rear
    RearPitch = Range("G28")
    RearRoll = Range("H28")
    
    '2. Compute Back-Sight Target and TBM Target ; Slope distance, Horizontal angle, Vertical angle
    Dim BS_SHV(0 To 2) As Double
    BS_SHV(0) = DirecSDHAVAby3P(TS(2), TS(1), TS(3), BS(2), BS(1), BS(3), BS(2), BS(1), BS(3), "H") 'Hor. Angle
    BS_SHV(1) = DirecSDHAVAby3P(TS(2), TS(1), TS(3), BS(2), BS(1), BS(3), BS(2), BS(1), BS(3), "V") 'Ver. Angle
    BS_SHV(2) = DirecSDHAVAby3P(TS(2), TS(1), TS(3), BS(2), BS(1), BS(3), BS(2), BS(1), BS(3), "S") 'Slope Dist.
    
    Dim TgtA_SHV(0 To 2) As Double
    TgtA_SHV(0) = DirecSDHAVAby3P(TS(2), TS(1), TS(3), BS(2), BS(1), BS(3), TgtA(2), TgtA(1), TgtA(3), "H") 'Hor. Angle
    TgtA_SHV(1) = DirecSDHAVAby3P(TS(2), TS(1), TS(3), BS(2), BS(1), BS(3), TgtA(2), TgtA(1), TgtA(3), "V") 'Ver. Angle
    TgtA_SHV(2) = DirecSDHAVAby3P(TS(2), TS(1), TS(3), BS(2), BS(1), BS(3), TgtA(2), TgtA(1), TgtA(3), "S") 'Slope Dist.
    
    Dim TgtB_SHV(0 To 2) As Double
    TgtB_SHV(0) = DirecSDHAVAby3P(TS(2), TS(1), TS(3), BS(2), BS(1), BS(3), TgtB(2), TgtB(1), TgtB(3), "H") 'Hor. Angle
    TgtB_SHV(1) = DirecSDHAVAby3P(TS(2), TS(1), TS(3), BS(2), BS(1), BS(3), TgtB(2), TgtB(1), TgtB(3), "V") 'Ver. Angle
    TgtB_SHV(2) = DirecSDHAVAby3P(TS(2), TS(1), TS(3), BS(2), BS(1), BS(3), TgtB(2), TgtB(1), TgtB(3), "S") 'Slope Dist.
    
    '3. Compute Articulation Angle
    HorArt = ArtAngle(LU, LD, RD, RU, KLU, KLD, KRD, KRU, HorArtLength, VerArtLength, "H")
    VerArt = ArtAngle(LU, LD, RD, RU, KLU, KLD, KRD, KRU, HorArtLength, VerArtLength, "V")
    
    
    '4. Compute Prism Offset after TBM Pitching and Roll
    Dim TgtNameA() As Variant
    Dim AMX() As Variant
    Dim AMY() As Variant
    Dim AMZ() As Variant
    ReDim TgtNameA(totalTgt - 1)
    ReDim AMX(totalTgt - 1)
    ReDim AMY(totalTgt - 1)
    ReDim AMZ(totalTgt - 1)
    
    For t = 0 To totalTgt - 1

        TgtNameA(t) = TgtName(t)
        AMX(t) = PrismOffset(MX(t), MY(t), MZ(t), RearPitch, RearRoll, "X")
        AMY(t) = PrismOffset(MX(t), MY(t), MZ(t), RearPitch, RearRoll, "Y")
        AMZ(t) = PrismOffset(MX(t), MY(t), MZ(t), RearPitch, RearRoll, "Z")
        'Debug.Print TgtNameA(t), AMX(t), AMY(t), AMZ(t)

    Next
    
    '5. Compute TBM Azimuth of Rear
    'Get index TBM target
    TgtAIndex = Application.Match(TgtA(0), TgtNameA, 0) - 1
    TgtBIndex = Application.Match(TgtB(0), TgtNameA, 0) - 1
    
    TBMRearAzi = TBMRearAz(TgtA(2), TgtA(1), AMX(TgtAIndex), AMY(TgtAIndex), TgtB(2), TgtB(1), AMX(TgtBIndex), AMY(TgtBIndex))
    'Debug.Print TBMRearAzi
    
    '6. Compute Coordinate Center of TBM Tail
    AVG_ArtJackStoke = (Abs((LU - KLU) + (LD - KLD) + (RD - KRD) + (RU - KRU)) / 4) / 1000
    
    ' TBM Length after TBM Pitching
    FrontLengthA = (FrontLength + AVG_ArtJackStoke) * Cos(DegtoRad(Abs(RearPitch)))
    RearLengthA = RearLength * Cos(DegtoRad(Abs(RearPitch)))
    TotalTBMLengthA = FrontLengthA + RearLengthA
    'Debug.Print FrontLengthA, RearLengthA
    
    'Coordinate Center of TBM Tail from Target.A
    dAMX_TgtA = AMX(TgtAIndex) - TotalTBMLengthA
    Dim TBMTail_TgtA(0 To 2) As Double
    TBMTail_TgtA(0) = pv_CoorYXtoNE(TgtA(2), TgtA(1), TBMRearAzi, dAMX_TgtA, -AMY(TgtAIndex), "N") 'Northing
    TBMTail_TgtA(1) = pv_CoorYXtoNE(TgtA(2), TgtA(1), TBMRearAzi, dAMX_TgtA, -AMY(TgtAIndex), "E") 'Easting
    TBMTail_TgtA(2) = (TgtA(3) - AMZ(TgtAIndex)) + dAMX_TgtA * Sin(DegtoRad(RearPitch)) 'Elevation
    'Debug.Print TBMTail_TgtA(0), TBMTail_TgtA(1), TBMTail_TgtA(2)
    
    'Coordinate Center of TBM Tail from Target.B
    dAMX_TgtB = AMX(TgtBIndex) - TotalTBMLengthA
    Dim TBMTail_TgtB(0 To 2) As Double
    TBMTail_TgtB(0) = pv_CoorYXtoNE(TgtB(2), TgtB(1), TBMRearAzi, dAMX_TgtB, -AMY(TgtBIndex), "N") 'Northing
    TBMTail_TgtB(1) = pv_CoorYXtoNE(TgtB(2), TgtB(1), TBMRearAzi, dAMX_TgtB, -AMY(TgtBIndex), "E") 'Easting
    TBMTail_TgtB(2) = (TgtB(3) - AMZ(TgtBIndex)) + dAMX_TgtB * Sin(DegtoRad(RearPitch)) 'Elevation
    'Debug.Print TBMTail_TgtB(0), TBMTail_TgtB(1), TBMTail_TgtB(2)
    
    'Coordinate Center of TBM Tail
    Dim TBMTail(0 To 2) As Double
    TBMTail(0) = (TBMTail_TgtA(0) + TBMTail_TgtB(0)) / 2 'Northing
    TBMTail(1) = (TBMTail_TgtA(1) + TBMTail_TgtB(1)) / 2 'Easting
    TBMTail(2) = (TBMTail_TgtA(2) + TBMTail_TgtB(2)) / 2 'Elevation
    'Debug.Print TBMTail(0), TBMTail(1), TBMTail(2)
    
    '7. Compute Coordinate Center of TBM Articulation
    Dim TBMArt(0 To 2) As Double
    TBMArt(0) = pv_CoorYXtoNE(TBMTail(1), TBMTail(0), TBMRearAzi, RearLengthA, 0, "N") 'Northing
    TBMArt(1) = pv_CoorYXtoNE(TBMTail(1), TBMTail(0), TBMRearAzi, RearLengthA, 0, "E") 'Easting
    TBMArt(2) = TBMTail(2) + RearLengthA * Sin(DegtoRad(RearPitch)) 'Elevation
    'Debug.Print TBMArt(0), TBMArt(1), TBMArt(2)
    
    '8. Compute Coordinate Center of TBM Head
    TBMHeadAzi = TBMRearAzi + HorArt
    FrontPitch = RearPitch + VerArt
    Dim TBMHead(0 To 2) As Double
    TBMHead(0) = pv_CoorYXtoNE(TBMArt(1), TBMArt(0), TBMHeadAzi, FrontLengthA * Cos(DegtoRad(FrontPitch)), 0, "N") 'Northing
    TBMHead(1) = pv_CoorYXtoNE(TBMArt(1), TBMArt(0), TBMHeadAzi, FrontLengthA * Cos(DegtoRad(FrontPitch)), 0, "E") 'Easting
    TBMHead(2) = TBMArt(2) + FrontLengthA * Sin(DegtoRad(FrontPitch)) 'Elevation
    'Debug.Print TBMHead(0), TBMHead(1), TBMHead(2)
    
    '9. Deviation of TBM Center and Chainage
    Dim TBMCenterN(0 To 2) As Double
    Dim TBMCenterE(0 To 2) As Double
    Dim TBMCenterZ(0 To 2) As Double
    TBMCenterN(0) = TBMTail(0): TBMCenterN(1) = TBMArt(0): TBMCenterN(2) = TBMHead(0)
    TBMCenterE(0) = TBMTail(1): TBMCenterE(1) = TBMArt(1): TBMCenterE(2) = TBMHead(1)
    TBMCenterZ(0) = TBMTail(2): TBMCenterZ(1) = TBMArt(2): TBMCenterZ(2) = TBMHead(2)
    'Debug.Print TBMCenterN(0), TBMCenterE(0), TBMCenterZ(0)
    
    Dim ChC(0 To 2) As Double
    Dim OsC(0 To 2) As Double
    Dim VtC(0 To 2) As Double
    Dim AziD(0 To 2) As Double
    
    For c = 0 To 2
    
        Dim Linear() As Variant
        ReDim Linear(totalDTA - 1)

        For d = 0 To totalDTA - 1
            Linear(d) = Sqr((EDTA(d) - TBMCenterE(c)) ^ 2 + (NDTA(d) - TBMCenterN(c)) ^ 2)
        Next
        
        'Find minimum linear from tunnel center to tunnel axis
        minLinear = Application.Min(Linear)
        minIndex = Application.Match(minLinear, Linear, 0) - 1
        'Debug.Print minLinear, minIndex
        
        'Point.B ; Point no., Chainage, Easting, Northing, Elevation
        PntB = PntDTA(minIndex - 1)
        ChB = ChDTA(minIndex - 1)
        EB = EDTA(minIndex - 1)
        NB = NDTA(minIndex - 1)
        ZB = ZDTA(minIndex - 1)
        'Debug.Print PntB, ChB, EB, NB, ZB
        
        'Point.M ; Point no., Chainage, Easting, Northing, Elevation
        PntM = PntDTA(minIndex)
        ChM = ChDTA(minIndex)
        EM = EDTA(minIndex)
        NM = NDTA(minIndex)
        ZM = ZDTA(minIndex)
        'Debug.Print PntM, ChM, EM, NM, ZM
        
        'Point.H ; Point no., Chainage, Easting, Northing, Elevation
        PntH = PntDTA(minIndex + 1)
        ChH = ChDTA(minIndex + 1)
        EH = EDTA(minIndex + 1)
        NH = NDTA(minIndex + 1)
        ZH = ZDTA(minIndex + 1)
        'Debug.Print PntH, ChH, EH, NH, ZH
        
        DistAC = pv_DirecDistAz(EB, NB, TBMCenterE(c), TBMCenterN(c), "D")
        DistHC = pv_DirecDistAz(EH, NH, TBMCenterE(c), TBMCenterN(c), "D")
        'Debug.Print DistAC, DistHC
        
        DistBM = pv_DirecDistAz(EB, NB, EM, NM, "D")
        AzBM = pv_DirecDistAz(EB, NB, EM, NM, "A")
        PitchBM = Pitching(ChB, ZB, ChM, ZM)
        'Debug.Print DistBM, AzBM, PitchBM
    
        DistMH = pv_DirecDistAz(EM, NM, EH, NH, "D")
        AzMH = pv_DirecDistAz(EM, NM, EH, NH, "A")
        PitchMH = Pitching(ChM, ZM, ChH, ZH)
        'Debug.Print DistMH, AzMH, PitchMH
    
        If DistAC < DistHC Then
    
            ChC(c) = ChM + pv_CoorNEtoYXL(EM, NM, AzBM, TBMCenterE(c), TBMCenterN(c), "Y") 'Chainage of tunnel center
            OsC(c) = pv_CoorNEtoYXL(EM, NM, AzBM, TBMCenterE(c), TBMCenterN(c), "X") 'Horizontal deviation of tunnel center
            VtC(c) = DeviateVt(ChM, ZM, PitchBM, ChC(c), TBMCenterZ(c)) 'Vertical deviation of tunnel center
            
            AziD(c) = AzBM 'Design Azimuth
            
        Else
    
            ChC(c) = ChM + pv_CoorNEtoYXL(EM, NM, AzMH, TBMCenterE(c), TBMCenterN(c), "Y") 'Chainage of tunnel center
            OsC(c) = pv_CoorNEtoYXL(EM, NM, AzMH, TBMCenterE(c), TBMCenterN(c), "X") 'Horizontal deviation of tunnel center
            VtC(c) = DeviateVt(ChM, ZM, PitchMH, ChC(c), TBMCenterZ(c)) 'Vertical deviation of tunnel center
    
            AziD(c) = AzMH 'Design Azimuth
    
        End If
        
    Next
    
    '10. TBM Graph
    
    Dim dCh() As Variant
    ReDim dCh(totalDTA - 1)
    
    For k = 0 To totalDTA - 1
        dCh(k) = Abs(ChC(1) - ChDTA(k))
    Next
    
    'Find minimum dCh from artculation center to tunnel axis
    min_dCh = Application.Min(dCh)
    minIndex_dCh = Application.Match(min_dCh, dCh, 0) - 1
    
    'TBM Graph X, Y, Z
    Dim TBMGraphX() As Double
    Dim TBMGraphY() As Double
    Dim TBMGraphZ() As Double
    ReDim TBMGraphX(2)
    ReDim TBMGraphY(2)
    ReDim TBMGraphZ(2)
    
    For k = 0 To 2
        TBMGraphX(k) = pv_CoorNEtoYXL(EDTA(minIndex_dCh), NDTA(minIndex_dCh), AziD(1), TBMCenterE(k), TBMCenterN(k), "Y")
        TBMGraphY(k) = pv_CoorNEtoYXL(EDTA(minIndex_dCh), NDTA(minIndex_dCh), AziD(1), TBMCenterE(k), TBMCenterN(k), "X")
        TBMGraphZ(k) = TBMCenterZ(k) - ZDTA(minIndex_dCh)
    Next
    
    'Data of Tunnel Axis Graph X, Y, Z
    Dim DTAGraphName() As Double
    Dim DTAGraphX() As Double
    Dim DTAGraphY() As Double
    Dim DTAGraphZ() As Double
    ReDim DTAGraphName(40)
    ReDim DTAGraphX(40)
    ReDim DTAGraphY(40)
    ReDim DTAGraphZ(40)
    
    For k = 0 To 40
        DTAGraphName(k) = PntDTA(minIndex_dCh - 20 + k)
        DTAGraphX(k) = pv_CoorNEtoYXL(EDTA(minIndex_dCh), NDTA(minIndex_dCh), AziD(1), EDTA(minIndex_dCh - 20 + k), NDTA(minIndex_dCh - 20 + k), "Y")
        DTAGraphY(k) = pv_CoorNEtoYXL(EDTA(minIndex_dCh), NDTA(minIndex_dCh), AziD(1), EDTA(minIndex_dCh - 20 + k), NDTA(minIndex_dCh - 20 + k), "X")
        DTAGraphZ(k) = ZDTA(minIndex_dCh - 20 + k) - ZDTA(minIndex_dCh)
    Next
    
    '11. Print Result

    'TBM Graph
    Sheets("TBM Graph").Visible = True 'Open
    Sheets("TBM Graph").Select
    Range("C5").Select
    For k = 0 To 2
        ActiveCell.Offset(k, 0).Value = TBMGraphX(k) 'TBM Graph X
        ActiveCell.Offset(k, 1).Value = TBMGraphY(k) * -1 'TBM Graph Y
        ActiveCell.Offset(k, 2).Value = TBMGraphZ(k) 'TBM Graph Z
    Next
    
    Range("B12").Select
    For k = 0 To 40
        ActiveCell.Offset(k, 0).Value = DTAGraphName(k) 'DTA Graph Name
        ActiveCell.Offset(k, 1).Value = DTAGraphX(k) 'DTA Graph X
        ActiveCell.Offset(k, 2).Value = DTAGraphY(k) * -1 'DTA Graph Y
        ActiveCell.Offset(k, 3).Value = DTAGraphZ(k) 'DTA Graph Z
    Next
    Sheets("TBM Graph").Visible = False 'Close
    
    'Back-Sight Target
    Sheets("Main Pro.").Select
    Range("I17").Value = BS_SHV(0) 'Hor. Angle
    Range("I18").Value = BS_SHV(1) 'Ver. Angle
    Range("I19").Value = BS_SHV(2) 'Slope Dist.
    
    'TBM Target A
    Range("K17").Value = TgtA_SHV(0) 'Hor. Angle
    Range("K18").Value = TgtA_SHV(1) 'Ver. Angle
    Range("K19").Value = TgtA_SHV(2) 'Slope Dist.
    
    'TBM Target B
    Range("M17").Value = TgtB_SHV(0) 'Hor. Angle
    Range("M18").Value = TgtB_SHV(1) 'Ver. Angle
    Range("M19").Value = TgtB_SHV(2) 'Slope Dist.
    
    'Articulation Angle
    Range("I28").Value = HorArt
    Range("J28").Value = VerArt
    
    'Coordinate Center of TBM Tail
    Range("G31").Value = TBMTail(0) 'Northing
    Range("G32").Value = TBMTail(1) 'Easting
    Range("G33").Value = TBMTail(2) 'Elevation
    
    'Coordinate Center of TBM Articulation
    Range("I31").Value = TBMArt(0) 'Northing
    Range("I32").Value = TBMArt(1) 'Easting
    Range("I33").Value = TBMArt(2) 'Elevation
    
    'Coordinate Center of TBM Head
    Range("K31").Value = TBMHead(0) 'Northing
    Range("K32").Value = TBMHead(1) 'Easting
    Range("K33").Value = TBMHead(2) 'Elevation
    
    'TBM Azimuth
    Range("G37").Value = TBMRearAzi 'TBM Rear
    Range("K37").Value = TBMHeadAzi 'TBM Head
    Range("J37").Value = TBMHeadAzi - AziD(2) 'Azimuth Deviation
    
    'Deviation of TBM
    'TBM Rear
    Range("G34").Value = ChC(0) 'Chainage
    Range("G35").Value = OsC(0) * Excavate_Direc 'Horizontal deviation
    Range("G36").Value = VtC(0) 'Vertical deviation
    
    Range("T10").Value = OsC(0) * Excavate_Direc * 1000 'Graph Hor.
    Range("T25").Value = VtC(0) * 1000 'Graph Ver.
    
    'TBM Articulation
    Range("I34").Value = ChC(1) 'Chainage
    Range("I35").Value = OsC(1) * Excavate_Direc 'Horizontal deviation
    Range("I36").Value = VtC(1) 'Vertical deviation
    
    Range("V10").Value = OsC(1) * Excavate_Direc * 1000 'Graph Hor.
    Range("V25").Value = VtC(1) * 1000 'Graph Ver.
    
    'TBM Head
    Range("K34").Value = ChC(2) 'Chainage
    Range("K35").Value = OsC(2) * Excavate_Direc 'Horizontal deviation
    Range("K36").Value = VtC(2) 'Vertical deviation
    
    Range("X10").Value = OsC(2) * Excavate_Direc * 1000 'Graph Hor.
    Range("X25").Value = VtC(2) * 1000 'Graph Ver.
    
    'Head Chainage
    Range("K9").Value = ChC(2) 'Chainage
    
    'Tunnel Distance
    Range("K10").Value = ChC(2) - StartCH
    
    'Date
    Range("O9").Value = "=TODAY()"
    
    'Time
    Range("O10").Value = "=NOW()"
    
    Sheets("Main Pro.").Select
    Range("G9").Select
    
    MsgBox "Manual TBM Position Computation Complete!"
    
End Sub

' TBM Rear Pitching and Azimuth Computation
Sub TBMRearPitch()
    Sheets("TBM Rear-Pitch, Roll").Select
    num = Application.Count(Range("C6:C10"))
    
    Range("B5").Select
    Dim PntNo() As Double
    Dim BackN() As Double
    Dim BackE() As Double
    Dim BackZ() As Double
    Dim AheadN() As Double
    Dim AheadE() As Double
    Dim AheadZ() As Double
    ReDim PntNo(num)
    ReDim BackN(num)
    ReDim BackE(num)
    ReDim BackZ(num)
    ReDim AheadN(num)
    ReDim AheadE(num)
    ReDim AheadZ(num)
    For i = 1 To num
        PntNo(i) = ActiveCell.Offset(i, 0)
        BackN(i) = ActiveCell.Offset(i, 1)
        BackE(i) = ActiveCell.Offset(i, 2)
        BackZ(i) = ActiveCell.Offset(i, 3)
        AheadN(i) = ActiveCell.Offset(i, 4)
        AheadE(i) = ActiveCell.Offset(i, 5)
        AheadZ(i) = ActiveCell.Offset(i, 6)
    Next
    
    Dim BADist() As Double
    Dim RearPitch() As Double
    Dim RearAz() As Double
    ReDim BADist(num)
    ReDim RearPitch(num)
    ReDim RearAz(num)
    For i = 1 To num
        BADist(i) = pv_DirecDistAz(BackE(i), BackN(i), AheadE(i), AheadN(i), "D")
        RearPitch(i) = RadtoDeg(Atn((AheadZ(i) - BackZ(i)) / BADist(i)))
        RearAz(i) = pv_DirecDistAz(BackE(i), BackN(i), AheadE(i), AheadN(i), "A")
    Next
    
    'Print Result
    Range("B5").Select
    For i = 1 To num
        ActiveCell.Offset(i, 7).Value = BADist(i)
        ActiveCell.Offset(i, 8).Value = RearPitch(i)
        ActiveCell.Offset(i, 9).Value = RearAz(i)
    Next
    
    Range("J11").Value = Application.Sum(RearPitch) / num
    Range("K11").Value = Application.Sum(RearAz) / num
    
    Range("B6").Select
End Sub

' TBM Rear Roll and Azimuth Computation
Sub TBMRearRoll()
    Sheets("TBM Rear-Pitch, Roll").Select
    num = Application.Count(Range("C17:C21"))
    
    Range("B16").Select
    Dim PntNo() As Double
    Dim LeftN() As Double
    Dim LeftE() As Double
    Dim LeftZ() As Double
    Dim RightN() As Double
    Dim RightE() As Double
    Dim RightZ() As Double
    ReDim PntNo(num)
    ReDim LeftN(num)
    ReDim LeftE(num)
    ReDim LeftZ(num)
    ReDim RightN(num)
    ReDim RightE(num)
    ReDim RightZ(num)
    For i = 1 To num
        PntNo(i) = ActiveCell.Offset(i, 0)
        LeftN(i) = ActiveCell.Offset(i, 1)
        LeftE(i) = ActiveCell.Offset(i, 2)
        LeftZ(i) = ActiveCell.Offset(i, 3)
        RightN(i) = ActiveCell.Offset(i, 4)
        RightE(i) = ActiveCell.Offset(i, 5)
        RightZ(i) = ActiveCell.Offset(i, 6)
    Next
    
    Dim LRDist() As Double
    Dim RearRoll() As Double
    Dim RearAz() As Double
    ReDim LRDist(num)
    ReDim RearRoll(num)
    ReDim RearAz(num)
    For i = 1 To num
        LRDist(i) = pv_DirecDistAz(LeftE(i), LeftN(i), RightE(i), RightN(i), "D")
        RearRoll(i) = RadtoDeg(Atn((LeftZ(i) - RightZ(i)) / LRDist(i)))
        
        RearAzChk = pv_DirecDistAz(LeftE(i), LeftN(i), RightE(i), RightN(i), "A") - 90
        If RearAzChk >= 0 Then
            RearAz(i) = RearAzChk
        Else
            RearAz(i) = 360 + RearAzChk
        End If
    Next
    
    'Print Result
    Range("B16").Select
    For i = 1 To num
        ActiveCell.Offset(i, 7).Value = LRDist(i)
        ActiveCell.Offset(i, 8).Value = RearRoll(i)
        ActiveCell.Offset(i, 9).Value = RearAz(i)
    Next
    
    Range("J22").Value = Application.Sum(RearRoll) / num
    Range("K22").Value = Application.Sum(RearAz) / num
    
    Range("B16").Select
End Sub

' TBM Rear Pitching Computation (Case : Top and Bottom)
Sub TBMRearPitchBT()
    Sheets("TBM Rear-Pitch, Roll").Select
    num = Application.Count(Range("C29:C33"))
    
    Range("B28").Select
    Dim PntNo() As Double
    Dim BottomN() As Double
    Dim BottomE() As Double
    Dim BottomZ() As Double
    Dim TopN() As Double
    Dim TopE() As Double
    Dim TopZ() As Double
    Dim RearAz() As Double
    ReDim PntNo(num)
    ReDim BottomN(num)
    ReDim BottomE(num)
    ReDim BottomZ(num)
    ReDim TopN(num)
    ReDim TopE(num)
    ReDim TopZ(num)
    ReDim RearAz(num)
    For i = 1 To num
        PntNo(i) = ActiveCell.Offset(i, 0)
        BottomN(i) = ActiveCell.Offset(i, 1)
        BottomE(i) = ActiveCell.Offset(i, 2)
        BottomZ(i) = ActiveCell.Offset(i, 3)
        TopN(i) = ActiveCell.Offset(i, 4)
        TopE(i) = ActiveCell.Offset(i, 5)
        TopZ(i) = ActiveCell.Offset(i, 6)
        RearAz(i) = ActiveCell.Offset(i, 7)
    Next
    
    Dim RearPitch() As Double
    ReDim RearPitch(num)
    For i = 1 To num
        diffZ = TopZ(i) - BottomZ(i)
        diffCH = pv_CoorNEtoYXL(TopE(i), TopN(i), RearAz(i), BottomE(i), BottomN(i), "Y")
        RearPitch(i) = RadtoDeg(Atn(diffCH / diffZ))
    Next
    
    'Print Result
    Range("B28").Select
    For i = 1 To num
        ActiveCell.Offset(i, 8).Value = RearPitch(i)
    Next
    
    Range("J34").Value = Application.Sum(RearPitch) / num
    
    Range("B29").Select
End Sub

'TBM Target Compuation
Sub TBMTargetXYZ()
    Sheets("Target Setting").Select
    
    ' TBM Data @Tail
    Dim TBMTail(0 To 2) As Double
    TBMTail(0) = Range("Q5") 'Northing
    TBMTail(1) = Range("Q6") 'Easting
    TBMTail(2) = Range("Q7") 'Elevation
    TBMLength = Range("Q8")
    TBMDirection = Range("Q9")
    TBMPitch = Range("Q10")
    TBMRoll = Range("Q11")
    
    ' TBM Head Position
    Dim TBMHead(0 To 2) As Double
    TBMHead(0) = TBMTail(0) + (TBMLength * Cos(DegtoRad(TBMPitch))) * Cos(DegtoRad(TBMDirection))
    TBMHead(1) = TBMTail(1) + (TBMLength * Cos(DegtoRad(TBMPitch))) * Sin(DegtoRad(TBMDirection))
    TBMHead(2) = TBMTail(2) + (TBMLength * Sin(DegtoRad(TBMPitch)))
    
    'Target Computation @Target Computation @TBM Roll and Pitch = 0 deg.
    Range("T5").Select
    num = Range(Selection, Selection.End(xlDown)).Count
    
    Range("T4").Select
    Dim TgtN() As Double
    Dim TgtE() As Double
    Dim TgtZ() As Double
    Dim KMX() As Double '@Roll and Pitch = 0 deg.
    Dim KMY() As Double '@Roll and Pitch = 0 deg.
    Dim KMZ() As Double '@Roll and Pitch = 0 deg.
    ReDim TgtN(num)
    ReDim TgtE(num)
    ReDim TgtZ(num)
    ReDim KMX(num) '@Roll and Pitch = 0 deg.
    ReDim KMY(num) '@Roll and Pitch = 0 deg.
    ReDim KMZ(num) '@Roll and Pitch = 0 deg.
    For i = 1 To num
        TgtN(i) = ActiveCell.Offset(i, 1)
        TgtE(i) = ActiveCell.Offset(i, 2)
        TgtZ(i) = ActiveCell.Offset(i, 3)

        'TBM Axis @TBM Roll and Pitch (AMX, AMY, AMZ)
        AMX = pv_CoorNEtoYXL(TBMHead(1), TBMHead(0), TBMDirection, TgtE(i), TgtN(i), "Y")
        AMY = pv_CoorNEtoYXL(TBMHead(1), TBMHead(0), TBMDirection, TgtE(i), TgtN(i), "X")
        AMZ = TgtZ(i) - TBMHead(2)

        'TBM Axis @TBM Roll and Pitch = 0 deg.
        KMX(i) = CalibratePrismOffset(AMX * -1, AMY, AMZ, TBMPitch, TBMRoll, "X")
        KMY(i) = CalibratePrismOffset(AMX * -1, AMY, AMZ, TBMPitch, TBMRoll, "Y")
        KMZ(i) = CalibratePrismOffset(AMX * -1, AMY, AMZ, TBMPitch, TBMRoll, "Z")
        'Debug.Print AMX, AMY, AMZ
        'Debug.Print KMX(i), KMY(i), KMZ(i)
    Next
    
    'Print Result
    Range("T4").Select
    For i = 1 To num
        ActiveCell.Offset(i, 4).Value = KMX(i)
        ActiveCell.Offset(i, 5).Value = KMY(i)
        ActiveCell.Offset(i, 6).Value = KMZ(i)
    Next
    Range("T5").Select
    
End Sub


'Clear content
Sub ClearData()
    Sheets("Main Pro.").Select
    
    Range("G9:G10").Select
    Selection.ClearContents
    
    Range("K9:K10").Select
    Selection.ClearContents
    
    Range("O9:O10").Select
    Selection.ClearContents
    
    Range("G13:I13").Select
    Selection.ClearContents
    
    Range("G14:M16").Select
    Selection.ClearContents
    
    Range("G14:M16").Select
    Selection.ClearContents
    
    Range("I17:M19").Select
    Selection.ClearContents
    
    Range("G22:M22").Select
    Selection.ClearContents
    
    Range("G28:J28").Select
    Selection.ClearContents
    
    Range("G31:K36").Select
    Selection.ClearContents
    
    Range("G37").Select
    Selection.ClearContents
    
    Range("J37:K37").Select
    Selection.ClearContents
    
    Range("T10").Select
    Selection.ClearContents
    
    Range("V10").Select
    Selection.ClearContents
    
    Range("X10").Select
    Selection.ClearContents
    
    Range("T25").Select
    Selection.ClearContents
    
    Range("V25").Select
    Selection.ClearContents
    
    Range("X25").Select
    Selection.ClearContents
    
    Range("G9").Select
    
End Sub
'Open Alignment Sheet
Sub OpenAlignmentSheet()
    
    Sheets("Alignment").Visible = True
    Sheets("Alignment").Select
    Range("K4").Select
    
End Sub

'Open Target Setting Sheet
Sub OpenTargetSheet()
    
    Sheets("Target Setting").Visible = True
    Sheets("Target Setting").Select
    Range("C5").Select
    
End Sub

'Open TBM Parameter Sheet
Sub OpenTBMParameterSheet()
    
    Sheets("TBM Parameter").Visible = True
    Sheets("TBM Parameter").Select
    Range("F4").Select
    
End Sub

'Open TBM Graph Sheet
Sub OpenTBMGraphSheet()
    
    Sheets("TBM Graph").Visible = True
    Sheets("TBM Graph").Select
    Range("C5").Select
    
End Sub

'Open TBM Rear-Pitch, Roll Sheet
Sub OpenTBMPitchRollSheet()
    
    Sheets("TBM Rear-Pitch, Roll").Visible = True
    Sheets("TBM Rear-Pitch, Roll").Select
    Range("B6").Select
    
End Sub

'Back to Main Program
Sub BacktoMainPro()

    ActiveSheet.Select
    ActiveWindow.SelectedSheets.Visible = False
    
    Sheets("Main Pro.").Select
    Range("G9").Select
    
End Sub

'Print to PDF
Sub PrintToPDF()
    Filename = "/" & "Manual TBM Position Program"
    FilePath = ActiveWorkbook.Path
        
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
    Filename:=FilePath & Filename, _
    Quality:=xlQualityStandard, _
    IncludeDocProperties:=True, _
    IgnorePrintAreas:=False, _
    OpenAfterPublish:=True
    
End Sub



























