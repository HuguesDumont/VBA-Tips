Attribute VB_Name = "UnitConverter"
Attribute VB_Description = "Divers functions like conversion between roman and arabic numerals or temperature conversion"
Option Explicit

'Function to convert arabic numerals to romans (FR : nombres romains)
Public Function ArabicToRomans(ByVal arabic As Long) As String
Attribute ArabicToRomans.VB_Description = "Function to convert arabic numerals to romans (FR : nombres romains)"
    Const hundreds = ",C,CC,CCC,CD,D,DC,DCC,DCCC,CM"
    Const tenths = ",X,XX,XXX,XL,L,LX,LXX,LXXX,XC"
    Const units = ",I,II,III,IV,V,VI,VII,VIII,IX"
    
    If (arabic = 0) Then
        arabictoroman = CStr(0)
    Else
        If arabic < 0 Then
            ArabicToRomans = "-"
            arabic = -arabic
        End If
        
        ArabicToRomans = ArabicToRomans & String(arabic \ 1000, "M")
        
        arabic = arabic Mod 1000
        ArabicToRomans = ArabicToRomans & Split(hundreds, ",")(arabic \ 100)
        
        arabic = arabic Mod 100
        ArabicToRomans = ArabicToRomans & Split(tenths, ",")(arabic \ 10)
        
        arabic = arabic Mod 10
        ArabicToRomans = ArabicToRomans & Split(units, ",")(arabic)
    End If
End Function

'Function to convert romans (FR : nombres romains) to arabic numerals
'If string isn't a valid roman numeral then return -1
Public Function RomanToArabic(ByVal roman As String) As Long
Attribute RomanToArabic.VB_Description = "Function to convert romans (FR : nombres romains) to arabic numerals\r\nIf string isn't a valid roman numeral then return -1"
    Dim i As Long, unit As Long, oldUnit As Long
    
    oldUnit = 1000
    
    roman = UCase(Trim(roman))
    If (roman = "0") Then
        RomanToArabic = 0
    Else
        For i = 1 To Len(roman)
            Select Case Mid(roman, i, 1)
                Case "I":  unit = 1
                Case "V":  unit = 5
                Case "X":  unit = 10
                Case "L":  unit = 50
                Case "C":  unit = 100
                Case "D":  unit = 500
                Case "M":  unit = 1000
                Case Else: 'invalid roman string because invalid character is detected
                    RomanToArabic = -1
                    Exit Function
            End Select
            If unit > oldUnit Then
                RomanToArabic = RomanToArabic - 2 * oldUnit
            End If
            RomanToArabic = RomanToArabic + unit
            oldUnit = unit
        Next i
    End If
End Function

'Function to convert temperature between units (Kelvin, Celsius, Fahreneit
'To chose source and destination units (default dest = 1 [celsius]) :
'1 : Celsius
'2 : Kelvin
'3 : Fahrenheit
'4 : Rankine
'5 : Reaumur
Public Function TemperatureConversion(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Single
Attribute TemperatureConversion.VB_Description = "'Function to convert temperature between scales (Kelvin, Celsius, Fahreneit\r\nTo chose source and destination scales :\r\n   0 : Kelvin\r\n   1 : Celsius\r\n   2 : Fahrenheit"
    Dim tmp As Double
    If (dest < 1 Or dest > 5) Then
        dest = 1
    End If
    If (src < 1 Or src > 5) Then
        src = 1
    End If
    If (src = dest) Then
        TemperatureConversion = val
    ElseIf (src = 1) Then
        Select Case dest
            Case 2
                TemperatureConversion = val + 273.15
            Case 3
                TemperatureConversion = (val * 1.8) + 32
            Case 4
                TemperatureConversion = (val * 1.8) + 491.67
            Case 5
                TemperatureConversion = val * 0.8
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2
                TemperatureConversion = val - 273.15
            Case 3
                If (val = 32) Then
                    TemperatureConversion = 0
                Else
                    TemperatureConversion = (val - 32) / 1.8
                End If
            Case 4
                If (val = 491.67) Then
                    TemperatureConversion = 0
                Else
                    TemperatureConversion = (val - 491.67) / 1.8
                End If
            Case 5
                TemperatureConversion = val * 1.25
        End Select
    Else
        tmp = TemperatureConversion(val, src, 1)
        TemperatureConversion = TemperatureConversion(tmp, 1, dest)
    End If
End Function

'Function to convert distances between units.
'To chose source and destination units (default destination is meter [m]):
'1 meter m
'2 kilometer km
'3 centimeter cm
'4 millimeter Mm
'5 micrometer um
'6 micron u
'7 nanometer nm
'8 picometer Pm
'9 decimeter dm
'10 nautical league(International)
'11 nautical mile(International)
'12 inch in
'13 yard yd
'14 foot ft
'15 league lea
'16 mile mi
'17 light year ly
'18 exameter Em
'19 petameter Pm
'20 terameter Tm
'21 gigameter Gm
'22 megameter Mm
'23 hectometer hm
'24 dekameter dam
'25 femtometer fm
'26 attometer am
'27 parsec pc
'28 astronomical unit AU
'29 ell
'30 mil
'31 microinch
'32 nautical league(UK)
'33 nautical mile(UK)
'34 mile (roman)
'35 furlong fur
'36 chain ch
'37 rope
'38 rod rd
'39 perch
'40 pole
'41 fathom fath
'42 link li
'43 cubit (UK)
'44 hand
'45 span (cloth)
'46 finger (cloth)
'47 nail (cloth)
'48 reed
'49 ken
'50 caliber cl
'51 centiinch cin
'52 pica
'53 Point
'54 twip
'55 barleycorn
'56 inch (US Survey)
'57 league (statute) lea (US)
'58 mile (statute) mi (US)
'59 foot (US Survey) ft (US)
'60 link (US Survey) li (US)
'61 aln
'62 famn
'63 angstrom a
'64 a.u. of length a.u
'65 x-unit X
'66 fermi F
'67 arpent
'68 roman actus
'69 vara de tarea
'70 vara conuguera
'71 vara castellana
'72 long reed
'73 long cubit
'74 li (Chinese)
'75 zhang (Chinese)
'76 chi (Chinese)
'77 cun (Chinese)
Public Function DistanceConversion(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Double
    Dim tmp As Double
    If (dest < 1 Or dest > 77) Then
        dest = 1
    End If
    If (src < 1 Or src > 77) Then
        src = 1
    End If
    If (dest = src Or val = 0) Then
        DistanceConversion = dest
    ElseIf (src = 1) Then
        Select Case dest
            Case 2
                DistanceConversion = val * 0.001
            Case 3
                DistanceConversion = val * 100
            Case 4
                DistanceConversion = val * 1000
            Case 5
                DistanceConversion = val * 1000000
            Case 6
                DistanceConversion = val * 1000000
            Case 7
                DistanceConversion = val * 1000000000#
            Case 8
                DistanceConversion = val * 1000000000000#
            Case 9
                DistanceConversion = val * 10
            Case 10
                DistanceConversion = val * 0.0001799856
            Case 11
                DistanceConversion = val * 0.0005399568
            Case 12
                DistanceConversion = val * 39.370078740157
            Case 13
                DistanceConversion = val * 1.093613
            Case 14
                DistanceConversion = val * 3.280839
            Case 15
                DistanceConversion = val * 2.0712330174427E-04
            Case 16
                DistanceConversion = val * 0.0006213711
            Case 17
                DistanceConversion = val * 1.056970721911E-16
            Case 18
                DistanceConversion = val * 1E-18
            Case 19
                DistanceConversion = val * 0.000000000000001
            Case 20
                DistanceConversion = val * 0.000000000001
            Case 21
                DistanceConversion = val * 0.000000001
            Case 22
                DistanceConversion = val * 0.000001
            Case 23
                DistanceConversion = val * 0.01
            Case 24
                DistanceConversion = val * 0.1
            Case 25
                DistanceConversion = val * 1E+15
            Case 26
                DistanceConversion = val * 1E+18
            Case 27
                DistanceConversion = val * 3.2407792899604E-17
            Case 28
                DistanceConversion = val * 6.6845871222684E-12
            Case 29
                DistanceConversion = val * 0.874891
            Case 30
                DistanceConversion = val * 39370.07874
            Case 31
                DistanceConversion = val * 39370078.739999
            Case 32
                DistanceConversion = val * 0.0001798706
            Case 33
                DistanceConversion = val * 0.0005396118
            Case 34
                DistanceConversion = val * 0.0006757652
            Case 35
                DistanceConversion = val * 0.0049709695
            Case 36
                DistanceConversion = val * 0.0497097
            Case 37
                DistanceConversion = val * 0.164042
            Case 38
                DistanceConversion = val * 0.198839
            Case 39
                DistanceConversion = val * 0.198839
            Case 40
                DistanceConversion = val * 0.198839
            Case 41
                DistanceConversion = val * 0.546807
            Case 42
                DistanceConversion = val * 4.97097
            Case 43
                DistanceConversion = val * 2.187227
            Case 44
                DistanceConversion = val * 9.84252
            Case 45
                DistanceConversion = val * 4.374453
            Case 46
                DistanceConversion = val * 8.748906
            Case 47
                DistanceConversion = val * 17.497813
            Case 48
                DistanceConversion = val * 0.364538
            Case 49
                DistanceConversion = val * 0.472063
            Case 50
                DistanceConversion = val * 3937.007874
            Case 51
                DistanceConversion = val * 3937.007874
            Case 52
                DistanceConversion = val * 236.220472
            Case 53
                DistanceConversion = val * 2834.645664
            Case 54
                DistanceConversion = val * 56692.91328
            Case 55
                DistanceConversion = val * 118.110236
            Case 56
                DistanceConversion = val * 39.37
            Case 57
                DistanceConversion = val * 0.0002071233
            Case 58
                DistanceConversion = val * 0.0006213699
            Case 59
                DistanceConversion = val * 3.280833
            Case 60
                DistanceConversion = val * 4.970959
            Case 61
                DistanceConversion = val * 1.684132
            Case 62
                DistanceConversion = val * 0.561377
            Case 63
                DistanceConversion = val * 10000000000#
            Case 64
                DistanceConversion = val * 18899990000#
            Case 65
                DistanceConversion = val * 9979996000000#
            Case 66
                DistanceConversion = val * 999999600000000#
            Case 67
                DistanceConversion = val * 0.0170877
            Case 68
                DistanceConversion = val * 0.0281859
            Case 69
                DistanceConversion = val * 0.399129
            Case 70
                DistanceConversion = val * 0.399129
            Case 71
                DistanceConversion = val * 1.197387
            Case 72
                DistanceConversion = val * 0.312461
            Case 73
                DistanceConversion = val * 1.874766
            Case 74
                DistanceConversion = val * 0.0020000004
            Case 75
                DistanceConversion = val * 0.3
            Case 76
                DistanceConversion = val * 3
            Case 77
                DistanceConversion = val * 30
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2
                DistanceConversion = val / 0.001
            Case 3
                DistanceConversion = val / 100
            Case 4
                DistanceConversion = val / 1000
            Case 5
                DistanceConversion = val / 1000000
            Case 6
                DistanceConversion = val / 1000000
            Case 7
                DistanceConversion = val / 1000000000#
            Case 8
                DistanceConversion = val / 1000000000000#
            Case 9
                DistanceConversion = val / 10
            Case 10
                DistanceConversion = val / 0.0001799856
            Case 11
                DistanceConversion = val / 0.0005399568
            Case 12
                DistanceConversion = val / 39.370078740157
            Case 13
                DistanceConversion = val / 1.093613
            Case 14
                DistanceConversion = val / 3.280839
            Case 15
                DistanceConversion = val / 0.0002071237
            Case 16
                DistanceConversion = val / 0.0006213711
            Case 17
                DistanceConversion = val / 1.056970721911E-16
            Case 18
                DistanceConversion = val / 1E-18
            Case 19
                DistanceConversion = val / 0.000000000000001
            Case 20
                DistanceConversion = val / 0.000000000001
            Case 21
                DistanceConversion = val / 0.000000001
            Case 22
                DistanceConversion = val / 0.000001
            Case 23
                DistanceConversion = val / 0.01
            Case 24
                DistanceConversion = val / 0.1
            Case 25
                DistanceConversion = val / 1E+15
            Case 26
                DistanceConversion = val / 1E+18
            Case 27
                DistanceConversion = val / 3.2407792899604E-17
            Case 28
                DistanceConversion = val / 6.6845871222684E-12
            Case 29
                DistanceConversion = val / 0.874891
            Case 30
                DistanceConversion = val / 39370.07874
            Case 31
                DistanceConversion = val / 39370078.739999
            Case 32
                DistanceConversion = val / 0.0001798706
            Case 33
                DistanceConversion = val / 0.0005396118
            Case 34
                DistanceConversion = val / 0.0006757652
            Case 35
                DistanceConversion = val / 0.0049709695
            Case 36
                DistanceConversion = val / 0.0497097
            Case 37
                DistanceConversion = val / 0.164042
            Case 38
                DistanceConversion = val / 0.198839
            Case 39
                DistanceConversion = val / 0.198839
            Case 40
                DistanceConversion = val / 0.198839
            Case 41
                DistanceConversion = val / 0.546807
            Case 42
                DistanceConversion = val / 4.97097
            Case 43
                DistanceConversion = val / 2.187227
            Case 44
                DistanceConversion = val / 9.84252
            Case 45
                DistanceConversion = val / 4.374453
            Case 46
                DistanceConversion = val / 8.748906
            Case 47
                DistanceConversion = val / 17.497813
            Case 48
                DistanceConversion = val / 0.364538
            Case 49
                DistanceConversion = val / 0.472063
            Case 50
                DistanceConversion = val / 3937.007874
            Case 51
                DistanceConversion = val / 3937.007874
            Case 52
                DistanceConversion = val / 236.220472
            Case 53
                DistanceConversion = val / 2834.645664
            Case 54
                DistanceConversion = val / 56692.91328
            Case 55
                DistanceConversion = val / 118.110236
            Case 56
                DistanceConversion = val / 39.37
            Case 57
                DistanceConversion = val / 0.0002071233
            Case 58
                DistanceConversion = val / 0.0006213699
            Case 59
                DistanceConversion = val / 3.280833
            Case 60
                DistanceConversion = val / 4.970959
            Case 61
                DistanceConversion = val / 1.684132
            Case 62
                DistanceConversion = val / 0.561377
            Case 63
                DistanceConversion = val / 10000000000#
            Case 64
                DistanceConversion = val / 18899990000#
            Case 65
                DistanceConversion = val / 9979996000000#
            Case 66
                DistanceConversion = val / 999999600000000#
            Case 67
                DistanceConversion = val / 0.0170877
            Case 68
                DistanceConversion = val / 0.0281859
            Case 69
                DistanceConversion = val / 0.399129
            Case 70
                DistanceConversion = val / 0.399129
            Case 71
                DistanceConversion = val / 1.197387
            Case 72
                DistanceConversion = val / 0.312461
            Case 73
                DistanceConversion = val / 1.874766
            Case 74
                DistanceConversion = val / 0.0020000004
            Case 75
                DistanceConversion = val / 0.3
            Case 76
                DistanceConversion = val / 3
            Case 77
                DistanceConversion = val / 30
        End Select
    Else
        tmp = DistanceConversion(val, src, 1)
        DistanceConversion = DistanceConversion(tmp, 1, dest)
    End If
End Function

'Function to convert acceleration between units.
'To chose source and destination units (default destination is m/s²):
'1 : m/s²
'2 : km/h²
'3 : km/min²
'4 : km/s²
'5 : m/h²
'6 : m/min²
'7 : mm/h²
'8 : mm/min²
'9 : mm/s²
'10 : miles/h²
'11 : miles/min²
'12 : miles/s²
'13 : ft/h²
'14 : ft/min²
'15 : ft/s²
'16 : in/h²
'17 : in/m²
'18 : in/s²
'19 : g
'20 : g(moon)
'21 : gal
Public Function AccelerationConversion(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Double
    Dim tmp As Double
    If (dest < 1 Or dest > 21) Then
        dest = 1
    End If
    
    If (src < 1 Or src > 21) Then
        src = 1
    End If
    
    If (src = dest Or val = 0) Then
        AccelerationConversion = val
    ElseIf (src = 1) Then
        Select Case dest
            Case 2
                AccelerationConversion = val * 12960
            Case 3
                AccelerationConversion = val * 3.6
            Case 4
                AccelerationConversion = val * 0.001
            Case 5
                AccelerationConversion = val * 12960000
            Case 6
                AccelerationConversion = val * 3600
            Case 7
                AccelerationConversion = val * 12960000000#
            Case 8
                AccelerationConversion = val * 3600000
            Case 9
                AccelerationConversion = val * 1000
            Case 10
                AccelerationConversion = val * 8052.9706513958
            Case 11
                AccelerationConversion = val * 2.2369362920544
            Case 12
                AccelerationConversion = val * 6.2137119223733E-04
            Case 13
                AccelerationConversion = val * 42519685.03937
            Case 14
                AccelerationConversion = val * 11811.023622047
            Case 15
                AccelerationConversion = val * 3.2808398950131
            Case 16
                AccelerationConversion = val * 510236220.47244
            Case 17
                AccelerationConversion = val * 141732.28346457
            Case 18
                AccelerationConversion = val * 39.370078740157
            Case 19
                AccelerationConversion = val * 0.10193679918451
            Case 20
                AccelerationConversion = val * 0.61349693251534
            Case 21
                AccelerationConversion = val * 100
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2
                AccelerationConversion = val / 12960
            Case 3
                AccelerationConversion = val / 3.6
            Case 4
                AccelerationConversion = val / 0.001
            Case 5
                AccelerationConversion = val / 12960000
            Case 6
                AccelerationConversion = val / 3600
            Case 7
                AccelerationConversion = val / 12960000000#
            Case 8
                AccelerationConversion = val / 3600000
            Case 9
                AccelerationConversion = val / 1000
            Case 10
                AccelerationConversion = val / 8052.9706513958
            Case 11
                AccelerationConversion = val / 2.2369362920544
            Case 12
                AccelerationConversion = val / 6.2137119223733E-04
            Case 13
                AccelerationConversion = val / 42519685.03937
            Case 14
                AccelerationConversion = val / 11811.023622047
            Case 15
                AccelerationConversion = val / 3.2808398950131
            Case 16
                AccelerationConversion = val / 510236220.47244
            Case 17
                AccelerationConversion = val / 141732.28346457
            Case 18
                AccelerationConversion = val / 39.370078740157
            Case 19
                AccelerationConversion = val / 0.10193679918451
            Case 20
                AccelerationConversion = val / 0.61349693251534
            Case 21
                AccelerationConversion = val / 100
        End Select
    Else
        tmp = AccelerationConversion(val, src, 1)
        AccelerationConversion = AccelerationConversion(tmp, 1, dest)
    End If
End Function

'Function to convert between angle units
'To chose source and destination units (default destination = 1 [degre])
'1 : °   (degre)
'2 : rad (radian)
'3 : '   (arc minute)
'4 : "   (arc second)
'5 : gon (grade)
'6 : mil (angular mil)
Public Function AngleConversion(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Double
    Dim tmp As Double
    If (dest < 1 Or dest > 6) Then
        dest = 1
    End If
    
    If (src < 1 Or src > 6) Then
        src = 1
    End If
    
    If (src = dest Or val = 0) Then
        AngleConversion = val
    ElseIf (src = 1) Then
        Select Case dest
            Case 2
                AngleConversion = val * ((4 * Atn(1)) / 180)
            Case 3
                AngleConversion = val * 60
            Case 4
                AngleConversion = val * 3600
            Case 5
                AngleConversion = val * 200 / 180
            Case 6
                AngleConversion = val * (1000 * (4 * Atn(1))) / 180
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2
                AngleConversion = val * (180 / (4 * Atn(1)))
            Case 3
                AngleConversion = val / 60
            Case 4
                AngleConversion = val / 3600
            Case 5
                AngleConversion = val * 180 / 200
            Case 6
                AngleConversion = val * (180 / (1000 * (4 * Atn(1))))
        End Select
    Else
        tmp = AngleConversion(val, src, 1)
        AngleConversion = AngleConversion(tmp, 1, dest)
    End If
End Function

'Function to convert between bandwith units
'To chose source and destination units (default destination = 1 [bytes])
'1 : B/s (bytes per seconde)
'2 : bps (bits per seconde)
'3 : kB/s (kilobytes per seconde)
'4 : MB/s (mégabytes per seconde)
'5 : GB/s (gigabytes per seconde)
'6 : TB/s (térabytes per seconde)
'7 : PB/s (petabytes per seconde)
'8 : KiB/s (kibibytes per seconde)
'9 : MiB/s (mebibytes per seconde)
'10 : GiB/s (gibibytes per seconde)
'11 : TiB/s (tebibytes per seconde)
'12 : PiB/s (pebibytes per seconde)
'13 : kbps (kilobits per seconde)
'14 : Mbps (mégabits per seconde)
'15 : Gbps (gigabits per seconde)
'16 : Tbps (térabits per seconde)
'17 : Pbps (petabits per seconde)
'18 : Kibps (kibibits per seconde)
'19 : Mibps (mebibits per seconde)
'20 : Gibps (gibibits per seconde)
'21 : Tibps (tebibits per seconde)
'22 : Pibps (pebibits per seconde)
Public Function BandwithConversion(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Double
    Dim tmp As Double
    If (dest < 1 Or dest > 22) Then
        dest = 1
    End If
    
    If (src < 1 Or src > 22) Then
        src = 1
    End If
    
    If (src = dest Or val = 0) Then
        BandwithConversion = val
    ElseIf (src = 1) Then
        Select Case dest
            Case 2
                BandwithConversion = val * 8
            Case 3
                BandwithConversion = val * 0.001
            Case 4
                BandwithConversion = val * 0.000001
            Case 5
                BandwithConversion = val * 0.000000001
            Case 6
                BandwithConversion = val * 0.000000000001
            Case 7
                BandwithConversion = val * 0.000000000000001
            Case 8
                BandwithConversion = val * 0.0009765625
            Case 9
                BandwithConversion = BandwithConversion(val, src, 8) * 0.0009765625
            Case 10
                BandwithConversion = BandwithConversion(val, src, 9) * 0.0009765625
            Case 11
                BandwithConversion = BandwithConversion(val, src, 10) * 0.0009765625
            Case 12
                BandwithConversion = BandwithConversion(val, src, 11) * 0.0009765625
            Case 13
                BandwithConversion = val * 0.008
            Case 14
                BandwithConversion = val * 0.000008
            Case 15
                BandwithConversion = val * 0.000000008
            Case 16
                BandwithConversion = val * 0.000000000008
            Case 17
                BandwithConversion = val * 0.000000000000008
            Case 18
                BandwithConversion = val * 0.0078125
            Case 19
                BandwithConversion = BandwithConversion(val, src, 18) * 0.0009765625
            Case 20
                BandwithConversion = BandwithConversion(val, src, 19) * 0.0009765625
            Case 21
                BandwithConversion = BandwithConversion(val, src, 20) * 0.0009765625
            Case 22
                BandwithConversion = BandwithConversion(val, src, 21) * 0.0009765625
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2
                BandwithConversion = val * 0.125
            Case 3
                BandwithConversion = val * 1000
            Case 4
                BandwithConversion = val * 1000000
            Case 5
                BandwithConversion = val * 1000000000
            Case 6
                BandwithConversion = val * 1000000000000#
            Case 7
                BandwithConversion = val * 1E+15
            Case 8
                BandwithConversion = val * 1024
            Case 9
                BandwithConversion = BandwithConversion(val, 8, dest) * 1024
            Case 10
                BandwithConversion = BandwithConversion(val, 9, dest) * 1024
            Case 11
                BandwithConversion = BandwithConversion(val, 10, dest) * 1024
            Case 12
                BandwithConversion = BandwithConversion(val, 11, dest) * 1024
            Case 13
                BandwithConversion = val * 125
            Case 14
                BandwithConversion = val * 125000
            Case 15
                BandwithConversion = val * 125000000#
            Case 16
                BandwithConversion = val * 125000000000#
            Case 17
                BandwithConversion = val * 125000000000000#
            Case 18
                BandwithConversion = val * 128
            Case 19
                BandwithConversion = BandwithConversion(val, 18, dest) * 1024
            Case 20
                BandwithConversion = BandwithConversion(val, 19, dest) * 1024
            Case 21
                BandwithConversion = BandwithConversion(val, 20, dest) * 1024
            Case 22
                BandwithConversion = BandwithConversion(val, 21, dest) * 1024
        End Select
    Else
        tmp = BandwithConversion(val, src, 1)
        BandwithConversion = BandwithConversion(tmp, 1, dest)
    End If
End Function

'Function to convert between electric load units
'To chose source and destination units (default destination = 1 [ampere-hour])
'1 : Ah  (Amp-hour / Ampere-hour)
'2 : C   (Coulomb / MilliAmp-second / MilliAmpere-second)
'3 : Fd  (Faraday)
'4 : e   (Elementary charge)
'5 : mAh (MilliAmp-hour / MilliAmpere-hour)
'6 : As  (Amp-second / Ampere-second)
'7 : mC  (Millicoulomb)
'8 : µC  (Microcoulomb)
'9 : nC  (Nanocoulomb)
Public Function ElectricLoadConversion(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Double
    Dim tmp As Double
    If (dest < 1 Or dest > 9) Then
        dest = 1
    End If
    
    If (src < 1 Or src > 9) Then
        src = 1
    End If
    
    If (src = dest Or val = 0) Then
        ElectricLoadConversion = val
    ElseIf (src = 1) Then
        Select Case dest
            Case 2
                ElectricLoadConversion = val * 3600
            Case 3
                ElectricLoadConversion = val * 0.037311367755258
            Case 4
                ElectricLoadConversion = val * 2.2469434729634E+22
            Case 5
                ElectricLoadConversion = val * 1000
            Case 6
                ElectricLoadConversion = val * 3600
            Case 7
                ElectricLoadConversion = val * 3600000
            Case 8
                ElectricLoadConversion = val * 3600000000#
            Case 9
                ElectricLoadConversion = val * 3600000000000#
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2
                ElectricLoadConversion = val / 3600
            Case 3
                ElectricLoadConversion = val * 26.801483305556
            Case 4
                ElectricLoadConversion = val * 4.4504902416667E-23
            Case 5
                ElectricLoadConversion = val * 0.001
            Case 6
                ElectricLoadConversion = val / 3600
            Case 7
                ElectricLoadConversion = val / 3600000
            Case 8
                ElectricLoadConversion = val / 3600000000#
            Case 9
                ElectricLoadConversion = val / 3600000000000#
        End Select
    Else
        tmp = ElectricLoadConversion(val, src, 1)
        ElectricLoadConversion = ElectricLoadConversion(tmp, 1, dest)
    End If
End Function

'Function to convert between fuel consumption units
'To chose source and destination units (default destination = 1 [liters/100km])
'1  : L/100km     (Liter per 100 kilometers)
'2  : L/km        (Liter per kilometer)
'3  : gal/100Km   (gallon per 100 kilometers)
'4  : gal/km      (gallon per kilometer)
'5  : km/L        (kilometer per litre)
'6  : km/gal (US) (kilometer by gallon (US))
'7  : km/gal (UK) (kilometer by gallon (UK))
'8  : mpg (US)    (mille per gallon (US))
'9  : mpg (UK     (mille per gallon (UK))
Public Function FuelConsumptionConversion(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Double
    Dim tmp As Double
    If (dest < 1 Or dest > 9) Then
        dest = 1
    End If
    
    If (src < 1 Or src > 9) Then
        src = 1
    End If
    
    If (src = dest Or val = 0) Then
        FuelConsumptionConversion = val
    ElseIf (src = 1) Then
        Select Case dest
            Case 2
                FuelConsumptionConversion = val * 0.01
            Case 3
                FuelConsumptionConversion = val * 0.26417205235815
            Case 4
                FuelConsumptionConversion = val * 2.6417205235815E-03
            Case 5
                FuelConsumptionConversion = val * 100
            Case 6
                FuelConsumptionConversion = val * 378.541178
            Case 7
                FuelConsumptionConversion = val * 454.609188
            Case 8
                FuelConsumptionConversion = val * 235.2145833
            Case 9
                FuelConsumptionConversion = val * 282.4809363
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2
                FuelConsumptionConversion = val * 100
            Case 3
                FuelConsumptionConversion = val * 3.785411784
            Case 4
                FuelConsumptionConversion = val * 378.5411784
            Case 5
                FuelConsumptionConversion = 100 / val
            Case 6
                FuelConsumptionConversion = 378.541178 / val
            Case 7
                FuelConsumptionConversion = 454.609188 / val
            Case 8
                FuelConsumptionConversion = 235.2145833 / val
            Case 9
                FuelConsumptionConversion = 282.4809363 / val
        End Select
    Else
        tmp = FuelConsumptionConversion(val, src, 1)
        FuelConsumptionConversion = FuelConsumptionConversion(tmp, 1, dest)
    End If
End Function

Sub test()
    Debug.Print ElectricLoadConversion(ElectricLoadConversion(3, 4, 3), 3, 4)
End Sub
