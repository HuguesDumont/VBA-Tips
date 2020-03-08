Attribute VB_Name = "UnitConverter"
Attribute VB_Description = "Divers functions like conversion between roman and arabic numerals or temperature conversion"
Option Explicit

'Function to convert between absorbed dose units
'To chose source and destination units (default destination = 1 [gray])
'1  : Gy (gray)
'2  : kGy (kilogray)
'3  : hGy (hectogray)
'4  : daGy (decagray)
'5  : dGy (decigray)
'6  : cGy (centigray)
'7  : mGy (milligray)
'8  : µGy (microgray)
'9  : nGy (nanogray)
'10 : krd (kilorad)
'11 : hrd (hectorad)
'12 : dard (decarad)
'13 : rd (rad)
'14 : drd (decirad)
'15 : crd (centirad)
'16 : mrd (millirad)
'17 : µrd (microrad)
'18 : nrd (nanorad)
'19 : kJ/kg (kilojoule/kilogram)
'20 : hJ/kg (hectojoule/kilogram)
'21 : daJ/kg (decajoule/kilogram)
'22 : J/kg  (joule/kilogram)
'23 : dJ/kg (decijoule/kilogram)
'24 : cJ/kg (centijoule/kilogram)
'25 : mJ/kg (millijoule/kilogram)
'26 : µJ/kg (microjoule/kilogram)
'27 : nJ/kg (nanojoule/kilogram)
Public Function AbsorbedDoseConverter(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Double
    Dim tmp                             As Double
    
    If (dest < 1 Or dest > 27) Then
        dest = 1
    End If
    
    If (src < 1 Or src > 27) Then
        src = 1
    End If
    
    If (src = dest Or val = 0) Then
        AbsorbedDoseConverter = val
    ElseIf (src = 1) Then
        Select Case dest
            Case 2, 19
                AbsorbedDoseConverter = val * 0.001
            Case 10
                AbsorbedDoseConverter = val * 0.1
            Case 5, 12, 23
                AbsorbedDoseConverter = val * 10
            Case 3, 4, 6, 7, 13, 14, 16, 20, 21, 24, 25
                AbsorbedDoseConverter = AbsorbedDoseConverter(val, src, dest - 1) * 10
            Case 26
                AbsorbedDoseConverter = val * 1000000
            Case 8, 9, 17, 18, 27
                AbsorbedDoseConverter = AbsorbedDoseConverter(val, src, dest - 1) * 1000
            Case 11, 22
                AbsorbedDoseConverter = val
            Case 15
                AbsorbedDoseConverter = val * 10000
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2, 19
                AbsorbedDoseConverter = val * 1000
            Case 3, 4, 6, 7, 13, 14, 16, 20, 21, 24, 25
                AbsorbedDoseConverter = AbsorbedDoseConverter(val, src - 1, dest) * 0.1
            Case 10
                AbsorbedDoseConverter = val * 10
            Case 5, 12, 23
                AbsorbedDoseConverter = val * 0.1
            Case 26
                AbsorbedDoseConverter = val * 0.000001
            Case 8, 9, 17, 18, 27
                AbsorbedDoseConverter = AbsorbedDoseConverter(val, src - 1, dest) * 0.001
            Case 11, 22
                AbsorbedDoseConverter = val
            Case 15
                AbsorbedDoseConverter = val * 0.0001
        End Select
    Else
        tmp = AbsorbedDoseConverter(val, src, 1)
        AbsorbedDoseConverter = AbsorbedDoseConverter(tmp, 1, dest)
    End If
End Function

'Function to convert between acceleration units.
'To chose source and destination units (default destination is m/s²):
'1  : m/s² (meter per square seconds)
'2  : km/h² (kilometer per square hour)
'3  : km/min² (kilometer per square minute)
'4  : km/s² (kilometer per square second)
'5  : m/h² (meter per square hour)
'6  : m/min² (meter per square minute)
'7  : mm/h² (millimeter per square hour)
'8  : mm/min² (millimeter per square minute)
'9  : mm/s² (millimeter per square second)
'10 : miles/h² (miles per square hour)
'11 : miles/min² (miles per square minute)
'12 : miles/s² (miles per square second)
'13 : ft/h² (foot per square hour)
'14 : ft/min² (foot per square minute)
'15 : ft/s² (foot per square second)
'16 : in/h² (inch per square hour)
'17 : in/m² (inch per square minute)
'18 : in/s² (inch per square second)
'19 : g (gravitation (earth))
'20 : g(moon) (gravitation (moon))
'21 : Gal (Gal)
Public Function AccelerationConverter(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Double
    Dim tmp                             As Double
    
    If (dest < 1 Or dest > 21) Then
        dest = 1
    End If
    
    If (src < 1 Or src > 21) Then
        src = 1
    End If
    
    If (src = dest Or val = 0) Then
        AccelerationConverter = val
    ElseIf (src = 1) Then
        Select Case dest
            Case 2
                AccelerationConverter = val * 12960
            Case 3
                AccelerationConverter = val * 3.6
            Case 4
                AccelerationConverter = val * 0.001
            Case 5
                AccelerationConverter = val * 12960000
            Case 6
                AccelerationConverter = val * 3600
            Case 7
                AccelerationConverter = val * 12960000000#
            Case 8
                AccelerationConverter = val * 3600000
            Case 9
                AccelerationConverter = val * 1000
            Case 10
                AccelerationConverter = val * 8052.9706513958
            Case 11
                AccelerationConverter = val * 2.2369362920544
            Case 12
                AccelerationConverter = val * 6.2137119223733E-04
            Case 13
                AccelerationConverter = val * 42519685.03937
            Case 14
                AccelerationConverter = val * 11811.023622047
            Case 15
                AccelerationConverter = val * 3.2808398950131
            Case 16
                AccelerationConverter = val * 510236220.47244
            Case 17
                AccelerationConverter = val * 141732.28346457
            Case 18
                AccelerationConverter = val * 39.370078740157
            Case 19
                AccelerationConverter = val * 0.10193679918451
            Case 20
                AccelerationConverter = val * 0.61349693251534
            Case 21
                AccelerationConverter = val * 100
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2
                AccelerationConverter = val / 12960
            Case 3
                AccelerationConverter = val / 3.6
            Case 4
                AccelerationConverter = val / 0.001
            Case 5
                AccelerationConverter = val / 12960000
            Case 6
                AccelerationConverter = val / 3600
            Case 7
                AccelerationConverter = val / 12960000000#
            Case 8
                AccelerationConverter = val / 3600000
            Case 9
                AccelerationConverter = val / 1000
            Case 10
                AccelerationConverter = val / 8052.9706513958
            Case 11
                AccelerationConverter = val / 2.2369362920544
            Case 12
                AccelerationConverter = val / 6.2137119223733E-04
            Case 13
                AccelerationConverter = val / 42519685.03937
            Case 14
                AccelerationConverter = val / 11811.023622047
            Case 15
                AccelerationConverter = val / 3.2808398950131
            Case 16
                AccelerationConverter = val / 510236220.47244
            Case 17
                AccelerationConverter = val / 141732.28346457
            Case 18
                AccelerationConverter = val / 39.370078740157
            Case 19
                AccelerationConverter = val / 0.10193679918451
            Case 20
                AccelerationConverter = val / 0.61349693251534
            Case 21
                AccelerationConverter = val / 100
        End Select
    Else
        tmp = AccelerationConverter(val, src, 1)
        AccelerationConverter = AccelerationConverter(tmp, 1, dest)
    End If
End Function

'Function to convert between angle units
'To chose source and destination units (default destination = 1 [degre])
'1 : °   (degree)
'2 : rad (radian)
'3 : '   (minute of arc)
'4 : "   (second of arc)
'5 : gon ((g) grade)
'6 : mil (NATO) (angular mil)
Public Function AngleConverter(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Double
    Dim tmp                             As Double
    
    If (dest < 1 Or dest > 6) Then
        dest = 1
    End If
    
    If (src < 1 Or src > 6) Then
        src = 1
    End If
    
    If (src = dest Or val = 0) Then
        AngleConverter = val
    ElseIf (src = 1) Then
        Select Case dest
            Case 2
                AngleConverter = val * ((4 * Atn(1)) / 180)
            Case 3
                AngleConverter = val * 60
            Case 4
                AngleConverter = val * 3600
            Case 5
                AngleConverter = val * 200 / 180
            Case 6
                AngleConverter = val * (1000 * (4 * Atn(1))) / 180
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2
                AngleConverter = val * (180 / (4 * Atn(1)))
            Case 3
                AngleConverter = val / 60
            Case 4
                AngleConverter = val / 3600
            Case 5
                AngleConverter = val * 180 / 200
            Case 6
                AngleConverter = val * (180 / (1000 * (4 * Atn(1))))
        End Select
    Else
        tmp = AngleConverter(val, src, 1)
        AngleConverter = AngleConverter(tmp, 1, dest)
    End If
End Function

'Function to convert arabic numerals to romans (FR : nombres romains)
Public Function ArabicToRomans(ByVal arabic As Long) As String
    Const HUNDREDS                      As String = ",C,CC,CCC,CD,D,DC,DCC,DCCC,CM"
    Const TENTHS                        As String = ",X,XX,XXX,XL,L,LX,LXX,LXXX,XC"
    
    Const UNITS                         As String = ",I,II,III,IV,V,VI,VII,VIII,IX"
    
    If (arabic = 0) Then
        ArabicToRomans = CStr(0)
    Else
        If (arabic < 0) Then
            ArabicToRomans = "-"
            arabic = -arabic
        End If
        
        ArabicToRomans = ArabicToRomans & String(arabic \ 1000, "M")
        
        arabic = arabic Mod 1000
        ArabicToRomans = ArabicToRomans & Split(HUNDREDS, ",")(arabic \ 100)
        
        arabic = arabic Mod 100
        ArabicToRomans = ArabicToRomans & Split(TENTHS, ",")(arabic \ 10)
        
        arabic = arabic Mod 10
        ArabicToRomans = ArabicToRomans & Split(UNITS, ",")(arabic)
    End If
End Function

'Function to convert between area units
'To chose source and destination units (default destination = 1 [square meter])
'1  : m2 (square meter)
'2  : km2 (square kilometer)
'3  : hm2 (square hectometer)
'4  : dam2 (square decameter)
'5  : dm2 (square decimeter)
'6  : cm2 (square centimeter)
'7  : mm2 (square millimeter)
'8  : ha (hectare)
'9  : a (are)
'10 : ca (centiare)
'11 : sq mi (square mile)
'12 : ac (acre)
'13 : rood (ro)
'14 : sq yd (square yard)
'15 : sq ft (square foot)
'16 : sq in (square inch)
'17 : lí (lí)
'18 : fen (fen)
'19 : mu (mu)
'20 : shí (shí)
'21 : qing (qing)
Public Function AreaConverter(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Double
    Dim tmp                             As Double
    
    If (dest < 1 Or dest > 21) Then
        dest = 1
    End If
    
    If (src < 1 Or src > 21) Then
        src = 1
    End If
    
    If (src = dest Or val = 0) Then
        AreaConverter = val
    ElseIf (src = 1) Then
        Select Case dest
            Case 2
                AreaConverter = val * 0.000001
            Case 3, 4, 6, 7, 9
                AreaConverter = AreaConverter(val, src, dest - 1) * 100
            Case 5
                AreaConverter = val * 100
            Case 8
                AreaConverter = val * 0.0001
            Case 10
                AreaConverter = val
            Case 11
                AreaConverter = val / 2589988.110336
            Case 12
                AreaConverter = val / 4046.8564224
            Case 13
                AreaConverter = val / 1011.7141056
            Case 14
                AreaConverter = val / 0.83612736
            Case 15
                AreaConverter = val / 0.09290304
            Case 16
                AreaConverter = val / 0.00064516
            Case 17
                AreaConverter = val * 0.15
            Case 18 To 21
                AreaConverter = AreaConverter(val, src, dest - 1) * 0.1
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2
                AreaConverter = val * 1000000
            Case 3, 4, 6, 7, 9
                AreaConverter = AreaConverter(val, src, dest - 1) * 0.01
            Case 5
                AreaConverter = val * 0.01
            Case 8
                AreaConverter = val * 10000
            Case 10
                AreaConverter = val
            Case 11
                AreaConverter = val * 2589988.110336
            Case 12
                AreaConverter = val * 4046.8564224
            Case 13
                AreaConverter = val * 1011.7141056
            Case 14
                AreaConverter = val * 0.83612736
            Case 15
                AreaConverter = val * 0.09290304
            Case 16
                AreaConverter = val * 0.00064516
            Case 17
                AreaConverter = val / 0.15
            Case 18 To 21
                AreaConverter = AreaConverter(val, src - 1, dest) * 10
        End Select
    Else
        tmp = AreaConverter(val, src, 1)
        AreaConverter = AreaConverter(tmp, 1, dest)
    End If
End Function

'Function to convert between bandwith units
'To chose source and destination units (default destination = 1 [bytes])
'1  : B/s (byte per seconde)
'2  : Bps (bit per seconde)
'3  : KB/s (kilobyte per seconde)
'4  : MB/s (megabyte per seconde)
'5  : GB/s (gigabyte per seconde)
'6  : TB/s (terabyte per seconde)
'7  : PB/s (petabyte per seconde)
'8  : KiB/s (kibibyte per seconde)
'9  : MiB/s (mebibyte per seconde)
'10 : GiB/s (gibibyte per seconde)
'11 : TiB/s (tebibyte per seconde)
'12 : PiB/s (pebibyte per seconde)
'13 : Kbps (kilobit per seconde)
'14 : Mbps (megabit per seconde)
'15 : Gbps (gigabit per seconde)
'16 : Tbps (terabit per seconde)
'17 : Pbps (petabit per seconde)
'18 : Kibps (kibibit per seconde)
'19 : Mibps (mebibit per seconde)
'20 : Gibps (gibibit per seconde)
'21 : Tibps (tebibit per seconde)
'22 : Pibps (pebibit per seconde)
Public Function BandwithConverter(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Double
    Dim tmp                             As Double
    
    If (dest < 1 Or dest > 22) Then
        dest = 1
    End If
    
    If (src < 1 Or src > 22) Then
        src = 1
    End If
    
    If (src = dest Or val = 0) Then
        BandwithConverter = val
    ElseIf (src = 1) Then
        Select Case dest
            Case 2
                BandwithConverter = val * 8
            Case 3
                BandwithConverter = val * 0.001
            Case 4 To 7, 14 To 17
                BandwithConverter = BandwithConverter(val, src, dest - 1) / 1000
            Case 8
                BandwithConverter = val / 1024
            Case 9 To 12, 19 To 22
                BandwithConverter = BandwithConverter(val, src, dest - 1) / 1024
            Case 13
                BandwithConverter = val * 0.008
            Case 18
                BandwithConverter = val / 128
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2
                BandwithConverter = val * 0.125
            Case 3
                BandwithConverter = val * 1000
            Case 4 To 7, 14 To 17
                BandwithConverter = BandwithConverter(val, src - 1, dest) * 1000
            Case 8
                BandwithConverter = val * 1024
            Case 9 To 12, 19 To 22
                BandwithConverter = BandwithConverter(val, src - 1, dest) * 1024
            Case 13
                BandwithConverter = val * 125
            Case 18
                BandwithConverter = val * 128
        End Select
    Else
        tmp = BandwithConverter(val, src, 1)
        BandwithConverter = BandwithConverter(tmp, 1, dest)
    End If
End Function

'Function to convert between density units
'To chose source and destination units (default destination = 1 [kilogram per cubic meter])
'1  : kg/m3 (kilogram per cubic meter)
'2  : t/m3 (ton per cubic meter)
'3  : kg/dm3 (kilogram per cubic decimeter)
'4  : g/cm3 (gram per cubic centimeter
'5  : kg/L (kilogram per liter)
'6  : g /mL (gram per milliliter)
'7  : lb/in3 (pound per cubic inch)
'8  : lb/ft3 (pound per cubic feet)
'9  : lb/gal (imperial) (pound per gallon (imperial))
'10 : lb/gal (US) (pound per gallon (US))
Public Function DensityConverter(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Double
    Dim tmp                             As Double
    
    If (dest < 1 Or dest > 10) Then
        dest = 1
    End If
    
    If (src < 1 Or src > 10) Then
        src = 1
    End If
    
    If (src = dest Or val = 0) Then
        DensityConverter = val
    ElseIf (src = 1) Then
        Select Case dest
            Case 2 To 6
                DensityConverter = val * 0.001
            Case 7
                DensityConverter = val / 27679.9
            Case 8
                DensityConverter = val / 16.01846
            Case 9
                DensityConverter = val / 99.77637
            Case 10
                DensityConverter = val / 119.8264
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2 To 6
                DensityConverter = val * 1000
            Case 7
                DensityConverter = val * 27679.9
            Case 8
                DensityConverter = val * 16.01846
            Case 9
                DensityConverter = val * 99.77637
            Case 10
                DensityConverter = val * 119.8264
        End Select
    Else
        tmp = DensityConverter(val, src, 1)
        DensityConverter = DensityConverter(tmp, 1, dest)
    End If
End Function

'Function to convert between distance units.
'To chose source and destination units (default destination is meter [m]):
'1  : m (meter)
'2  : km (kilometer)
'3  : cm (centimeter)
'4  : mm (millimeter)
'5  : µm (micrometer)
'6  : µ (micron)
'7  : nm (nanometer)
'8  : pm (picometer)
'9  : dm (decimeter)
'10 : nautical league (International)
'11 : nautical mile (International)
'12 : in (inch)
'13 : yd (yard)
'14 : ft (foot)
'15 : lea (league)
'16 : mi (mile)
'17 : ly (light year)
'18 : Em (Exameter)
'19 : Pm (Petameter)
'20 : Tm (Terameter)
'21 : Gm (Gigameter)
'22 : Mm (Megameter)
'23 : hm (hectometer)
'24 : dam (decameter)
'25 : fm (femtometer)
'26 : am (Attometer)
'27 : pc (parsec)
'28 : AU (astronomical unit)
'29 : ell
'30 : mil
'31 : microinch
'32 : nautical league (UK)
'33 : nautical mile (UK)
'34 : mile (roman)
'35 : fur (furlong)
'36 : ch (chain)
'37 : rope
'38 : rd (rod)
'39 : perch
'40 : pole
'41 : fath (fathom)
'42 : li (link)
'43 : cubit (UK)
'44 : hand
'45 : span (cloth)
'46 : finger (cloth)
'47 : nail (cloth)
'48 : reed
'49 : ken
'50 : cl (caliber)
'51 : cin (centiinch)
'52 : pica
'53 : point
'54 : twip
'55 : barleycorn
'56 : inch (US Survey)
'57 : lea (US) (league (statute))
'58 : mi (US) (mile (statute))
'59 : ft (US) (foot (US Survey))
'60 : li (Us) (link (US Survey))
'61 : Aln
'62 : Famn
'63 : a (angstrom)
'64 : a.u (a.u. of length)
'65 : X-unit X
'66 : F (fermi)
'67 : arpent
'68 : roman actus
'69 : vara de tarea
'70 : vara conuguera
'71 : vara castellana
'72 : long reed
'73 : long cubit
'74 : li (Chinese)
'75 : zhang (Chinese)
'76 : chi (Chinese)
'77 : cun (Chinese)
Public Function DistanceConverter(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Double
    Dim tmp                             As Double
    
    If (dest < 1 Or dest > 77) Then
        dest = 1
    End If
    
    If (src < 1 Or src > 77) Then
        src = 1
    End If
    
    If (dest = src Or val = 0) Then
        DistanceConverter = dest
    ElseIf (src = 1) Then
        Select Case dest
            Case 2
                DistanceConverter = val * 0.001
            Case 3
                DistanceConverter = val * 100
            Case 4, 24
                DistanceConverter = DistanceConverter(val, src, dest - 1) * 10
            Case 5, 6
                DistanceConverter = val * 1000000
            Case 7, 8, 26
                DistanceConverter = DistanceConverter(val, src, dest - 1) * 1000
            Case 9
                DistanceConverter = val * 10
            Case 10
                DistanceConverter = val * 0.0001799856
            Case 11
                DistanceConverter = val * 0.0005399568
            Case 12
                DistanceConverter = val * 39.370078740157
            Case 13
                DistanceConverter = val * 1.093613
            Case 14
                DistanceConverter = val * 3.280839
            Case 15
                DistanceConverter = val * 2.0712330174427E-04
            Case 16
                DistanceConverter = val * 0.0006213711
            Case 17
                DistanceConverter = val * 1.056970721911E-16
            Case 18 To 22
                DistanceConverter = DistanceConverter(val, src, dest + 1) / 1000
            Case 23
                DistanceConverter = val * 0.01
            Case 25
                DistanceConverter = val * 1E+15
            Case 27
                DistanceConverter = val * 3.2407792899604E-17
            Case 28
                DistanceConverter = val * 6.6845871222684E-12
            Case 29
                DistanceConverter = val * 0.874891
            Case 30
                DistanceConverter = val * 39370.07874
            Case 31
                DistanceConverter = val * 39370078.739999
            Case 32
                DistanceConverter = val * 0.0001798706
            Case 33
                DistanceConverter = val * 0.0005396118
            Case 34
                DistanceConverter = val * 0.0006757652
            Case 35
                DistanceConverter = val * 0.0049709695
            Case 36
                DistanceConverter = val * 0.0497097
            Case 37
                DistanceConverter = val * 0.164042
            Case 38, 39, 40
                DistanceConverter = val * 0.198839
            Case 41
                DistanceConverter = val * 0.546807
            Case 42
                DistanceConverter = val * 4.97097
            Case 43
                DistanceConverter = val * 2.187227
            Case 44
                DistanceConverter = val * 9.84252
            Case 45
                DistanceConverter = val * 4.374453
            Case 46
                DistanceConverter = val * 8.748906
            Case 47
                DistanceConverter = val * 17.497813
            Case 48
                DistanceConverter = val * 0.364538
            Case 49
                DistanceConverter = val * 0.472063
            Case 50, 51
                DistanceConverter = val * 3937.007874
            Case 52
                DistanceConverter = val * 236.220472
            Case 53
                DistanceConverter = val * 2834.645664
            Case 54
                DistanceConverter = val * 56692.91328
            Case 55
                DistanceConverter = val * 118.110236
            Case 56
                DistanceConverter = val * 39.37
            Case 57
                DistanceConverter = val * 0.0002071233
            Case 58
                DistanceConverter = val * 0.0006213699
            Case 59
                DistanceConverter = val * 3.280833
            Case 60
                DistanceConverter = val * 4.970959
            Case 61
                DistanceConverter = val * 1.684132
            Case 62
                DistanceConverter = val * 0.561377
            Case 63
                DistanceConverter = val * 10000000000#
            Case 64
                DistanceConverter = val * 18899990000#
            Case 65
                DistanceConverter = val * 9979996000000#
            Case 66
                DistanceConverter = val * 999999600000000#
            Case 67
                DistanceConverter = val * 0.0170877
            Case 68
                DistanceConverter = val * 0.0281859
            Case 69, 70
                DistanceConverter = val * 0.399129
            Case 71
                DistanceConverter = val * 1.197387
            Case 72
                DistanceConverter = val * 0.312461
            Case 73
                DistanceConverter = val * 1.874766
            Case 74
                DistanceConverter = val * 0.0020000004
            Case 75
                DistanceConverter = val * 0.3
            Case 76, 77
                DistanceConverter = DistanceConverter(val, src, dest - 1) * 10
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2
                DistanceConverter = val * 1000
            Case 3
                DistanceConverter = val * 0.01
            Case 4, 24
                DistanceConverter = DistanceConverter(val, src - 1, dest) * 0.1
            Case 5, 6
                DistanceConverter = val * 0.000001
            Case 7, 8, 26
                DistanceConverter = DistanceConverter(val, src - 1, dest) * 0.001
            Case 9
                DistanceConverter = val * 0.1
            Case 10
                DistanceConverter = val / 0.0001799856
            Case 11
                DistanceConverter = val / 0.0005399568
            Case 12
                DistanceConverter = val / 39.370078740157
            Case 13
                DistanceConverter = val / 1.093613
            Case 14
                DistanceConverter = val / 3.280839
            Case 15
                DistanceConverter = val / 0.0002071237
            Case 16
                DistanceConverter = val / 0.0006213711
            Case 17
                DistanceConverter = val / 1.056970721911E-16
            Case 18 To 22
                DistanceConverter = DistanceConverter(val, src + 1, dest) * 1000
            Case 23
                DistanceConverter = val * 100
            Case 25
                DistanceConverter = val * 0.000000000000001
            Case 27
                DistanceConverter = val / 3.2407792899604E-17
            Case 28
                DistanceConverter = val / 6.6845871222684E-12
            Case 29
                DistanceConverter = val / 0.874891
            Case 30
                DistanceConverter = val / 39370.07874
            Case 31
                DistanceConverter = val / 39370078.739999
            Case 32
                DistanceConverter = val / 0.0001798706
            Case 33
                DistanceConverter = val / 0.0005396118
            Case 34
                DistanceConverter = val / 0.0006757652
            Case 35
                DistanceConverter = val / 0.0049709695
            Case 36
                DistanceConverter = val / 0.0497097
            Case 37
                DistanceConverter = val / 0.164042
            Case 38, 39, 40
                DistanceConverter = val / 0.198839
            Case 41
                DistanceConverter = val / 0.546807
            Case 42
                DistanceConverter = val / 4.97097
            Case 43
                DistanceConverter = val / 2.187227
            Case 44
                DistanceConverter = val / 9.84252
            Case 45
                DistanceConverter = val / 4.374453
            Case 46
                DistanceConverter = val / 8.748906
            Case 47
                DistanceConverter = val / 17.497813
            Case 48
                DistanceConverter = val / 0.364538
            Case 49
                DistanceConverter = val / 0.472063
            Case 50, 51
                DistanceConverter = val / 3937.007874
            Case 52
                DistanceConverter = val / 236.220472
            Case 53
                DistanceConverter = val / 2834.645664
            Case 54
                DistanceConverter = val / 56692.91328
            Case 55
                DistanceConverter = val / 118.110236
            Case 56
                DistanceConverter = val / 39.37
            Case 57
                DistanceConverter = val / 0.0002071233
            Case 58
                DistanceConverter = val / 0.0006213699
            Case 59
                DistanceConverter = val / 3.280833
            Case 60
                DistanceConverter = val / 4.970959
            Case 61
                DistanceConverter = val / 1.684132
            Case 62
                DistanceConverter = val / 0.561377
            Case 63
                DistanceConverter = val / 10000000000#
            Case 64
                DistanceConverter = val / 18899990000#
            Case 65
                DistanceConverter = val / 9979996000000#
            Case 66
                DistanceConverter = val / 999999600000000#
            Case 67
                DistanceConverter = val / 0.0170877
            Case 68
                DistanceConverter = val / 0.0281859
            Case 69, 70
                DistanceConverter = val / 0.399129
            Case 71
                DistanceConverter = val / 1.197387
            Case 72
                DistanceConverter = val / 0.312461
            Case 73
                DistanceConverter = val / 1.874766
            Case 74
                DistanceConverter = val / 0.0020000004
            Case 75
                DistanceConverter = val / 0.3
            Case 76, 77
                DistanceConverter = DistanceConverter(val, src - 1, dest) * 10
        End Select
    Else
        tmp = DistanceConverter(val, src, 1)
        DistanceConverter = DistanceConverter(tmp, 1, dest)
    End If
End Function

'Function to convert between electric load units
'To chose source and destination units (default destination = 1 [ampere-hour])
'1 : Ah  (Amp-hour / Ampere-hour)
'2 : C   (Coulomb / MilliAmp-second / MilliAmpere-second)
'3 : Fd  (Faraday)
'4 : E   (Elementary charge)
'5 : MAh (MilliAmp-hour / MilliAmpere-hour)
'6 : As  (Amp-second / Ampere-second)
'7 : MC  (Millicoulomb)
'8 : µC  (Microcoulomb)
'9 : NC  (Nanocoulomb)
Public Function ElectricChargeConverter(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Double
    Dim tmp                             As Double
    
    If (dest < 1 Or dest > 9) Then
        dest = 1
    End If
    
    If (src < 1 Or src > 9) Then
        src = 1
    End If
    
    If (src = dest Or val = 0) Then
        ElectricChargeConverter = val
    ElseIf (src = 1) Then
        Select Case dest
            Case 2, 6
                ElectricChargeConverter = val * 3600
            Case 3
                ElectricChargeConverter = val * 0.037311367755258
            Case 4
                ElectricChargeConverter = val * 2.2469434729634E+22
            Case 5
                ElectricChargeConverter = val * 1000
            Case 7
                ElectricChargeConverter = val * 3600000
            Case 8, 9
                ElectricChargeConverter = ElectricChargeConverter(val, src, dest - 1) * 1000
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2, 6
                ElectricChargeConverter = val / 3600
            Case 3
                ElectricChargeConverter = val * 26.801483305556
            Case 4
                ElectricChargeConverter = val * 4.4504902416667E-23
            Case 5
                ElectricChargeConverter = val * 0.001
            Case 7
                ElectricChargeConverter = val / 3600000
            Case 8, 9
                ElectricChargeConverter = ElectricChargeConverter(val, src - 1, dest) / 1000
        End Select
    Else
        tmp = ElectricChargeConverter(val, src, 1)
        ElectricChargeConverter = ElectricChargeConverter(tmp, 1, dest)
    End If
End Function

'Function to convert between energy units.
'To chose source and destination units (default destination is 1 : joule [J]):
'1  : J (joule)
'2  : GJ (gigajoule)
'3  : MJ (megajoule)
'4  : kJ (kilojoule)
'5  : kcal (kilocalory)
'6  : cal (calory)
'7  : kWh (Kilowatt-hour)
'8  : MWh (Megawatt-hou)
'9  : Wh (Watt-hour)
'10 : Ws (Watt-second)
'11 : koe (kilo of oil equivalent)
'12 : toe (ton of oil equivalent)
'13 : ktoe (kiloton of oil equivalent)
'14 : Mtoe (megaton of oil equivalent)
'15 : boe (barrel of oil equivalent)
'16 : kboe (thousand barrel of oil equivalent)
'17 : Mboe (million barrel of oil equivalent)
'18 : Gm3 NG (billion cubic meter of natural gas)
'19 : Gft3 NG (billion cubic foot of natural gas)
'20 : Mt LNG (megaton of liquefied natural gas)
'21 : Gt LNG (gigaton of liquefied natural gas)
'22 : kg SKE (kilogram hard coal)
'23 : t SKE (ton hard coal)
'24 : GeV (gigaelectronvolt)
'25 : TeV (tera-electronvolt)
'26 : MeV (mega-electronvolt)
'27 : keV (kilo-electronvolt)
'28 : eV (electronvolt)
'29 : Btu (British termal unit)
'30 : MMBtu (million btu)
'31 : thm (therm)
'32 : quad (quad)
'33 : erg (erg)
'34 : Mt (megaton TNT)
'35 : kT (kiloton TNT)
'36 : ft-lb (Foot-pound)
Public Function EnergyConverter(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Double
    Dim tmp                             As Double
    
    If (dest < 1 Or dest > 36) Then
        dest = 1
    End If
    
    If (src < 1 Or src > 36) Then
        src = 1
    End If
    
    If (dest = src Or val = 0) Then
        EnergyConverter = dest
    ElseIf (src = 1) Then
        Select Case dest
            Case 2
                EnergyConverter = val * 0.000000001
            Case 3, 4, 8, 12 To 14, 16, 17, 21, 25
                EnergyConverter = EnergyConverter(val, src, dest - 1) * 0.001
            Case 5
                EnergyConverter = val / 4186.8
            Case 6
                EnergyConverter = val / 4.1868
            Case 7
                EnergyConverter = val / 3600000
            Case 9
                EnergyConverter = val / 3600
            Case 10
                EnergyConverter = val
            Case 11
                EnergyConverter = val / 41868000
            Case 15
                EnergyConverter = val / 5861520000#
            Case 18
                EnergyConverter = val / 3.76812E+16
            Case 19
                EnergyConverter = val / 1.088568E+15
            Case 20
                EnergyConverter = val / 5.200993789E+16
            Case 22
                EnergyConverter = val / 29307600
            Case 23
                EnergyConverter = val / 29307600000#
            Case 24
                EnergyConverter = val / 1.602176487E-10
            Case 26
                EnergyConverter = val / 1.602176487E-13
            Case 27, 28
                EnergyConverter = EnergyConverter(val, src, dest - 1) * 1000
            Case 29
                EnergyConverter = val / 1055.87
            Case 30
                EnergyConverter = val / 1055870000
            Case 31
                EnergyConverter = val / 105587000
            Case 32
                EnergyConverter = val / 1.05587E+18
            Case 33
                EnergyConverter = val * 10000000
            Case 34
                EnergyConverter = val / 4.184E+15
            Case 35
                EnergyConverter = val / 4184000000000#
            Case 36
                EnergyConverter = val / 1.3558179483314
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2
                EnergyConverter = val * 1000000000
            Case 3, 4, 8, 12 To 14, 16, 17, 21, 25
                EnergyConverter = EnergyConverter(val, src - 1, dest) * 1000
            Case 5
                EnergyConverter = val * 4186.8
            Case 6
                EnergyConverter = val * 4.1868
            Case 7
                EnergyConverter = val * 3600000
            Case 9
                EnergyConverter = val * 3600
            Case 10
                EnergyConverter = val
            Case 11
                EnergyConverter = val * 41868000
            Case 15
                EnergyConverter = val * 5861520000#
            Case 18
                EnergyConverter = val * 3.76812E+16
            Case 19
                EnergyConverter = val * 1.088568E+15
            Case 20
                EnergyConverter = val * 5.200993789E+16
            Case 22
                EnergyConverter = val * 29307600
            Case 23
                EnergyConverter = val * 29307600000#
            Case 24
                EnergyConverter = val * 1.602176487E-10
            Case 26
                EnergyConverter = val * 1.602176487E-13
            Case 27, 28
                EnergyConverter = EnergyConverter(val, src - 1, dest) * 0.001
            Case 29
                EnergyConverter = val * 1055.87
            Case 30
                EnergyConverter = val * 1055870000
            Case 31
                EnergyConverter = val * 105587000
            Case 32
                EnergyConverter = val * 1.05587E+18
            Case 33
                EnergyConverter = val * 0.0000001
            Case 34
                EnergyConverter = val * 4.184E+15
            Case 35
                EnergyConverter = val * 4184000000000#
            Case 36
                EnergyConverter = val * 1.3558179483314
        End Select
    Else
        tmp = EnergyConverter(val, src, 1)
        EnergyConverter = EnergyConverter(tmp, 1, dest)
    End If
End Function

'Function to convert between equivalent dose units
'To chose source and destination units (default destination = 1 [sievert])
'1  : Sv (sievert)
'2  : kSv (kilosievert)
'3  : hSv (hectosievert)
'4  : daSv (decasievert)
'5  : dSv (decisievert)
'6  : cSv (centisievert)
'7  : mSv (millisievert)
'8  : µSv (microsievert)
'9  : nSv (nanosievert)
'10 : krem (kilorem) (roentgen equivalent)
'11 : hrem (hectorem) (roentgen equivalent)
'12 : darem (decarem) (roentgen equivalent)
'13 : rem (rem) (roentgen equivalent)
'14 : drem (decirem) (roentgen equivalent)
'15 : crem (centirem) (roentgen equivalent)
'16 : mrem (millirem) (roentgen equivalent)
'17 : µrem (microrem) (roentgen equivalent)
'18 : nrem (nanorem) (roentgen equivalent)
Public Function EquivalentDoseconverter(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Double
    Dim tmp                             As Double
    
    If (dest < 1 Or dest > 18) Then
        dest = 1
    End If
    
    If (src < 1 Or src > 18) Then
        src = 1
    End If
    
    If (src = dest Or val = 0) Then
        EquivalentDoseconverter = val
    ElseIf (src = 1) Then
        Select Case dest
            Case 2
                EquivalentDoseconverter = val * 0.001
            Case 3, 4, 6, 7, 13, 14, 16
                EquivalentDoseconverter = EquivalentDoseconverter(val, src, dest - 1) * 10
            Case 10
                EquivalentDoseconverter = val * 0.1
            Case 5, 12
                EquivalentDoseconverter = val * 10
            Case 8, 9, 17, 18
                EquivalentDoseconverter = EquivalentDoseconverter(val, src, dest - 1) * 1000
            Case 11
                EquivalentDoseconverter = val
            Case 15
                EquivalentDoseconverter = val * 10000
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2
                EquivalentDoseconverter = val * 1000
            Case 3, 4, 6, 7, 13, 14, 16
                EquivalentDoseconverter = EquivalentDoseconverter(val, src - 1, dest) * 0.1
            Case 10
                EquivalentDoseconverter = val * 10
            Case 5, 12
                EquivalentDoseconverter = val * 0.1
            Case 8, 9, 17, 18
                EquivalentDoseconverter = EquivalentDoseconverter(val, src - 1, dest) * 0.001
            Case 11
                EquivalentDoseconverter = val
            Case 15
                EquivalentDoseconverter = val * 0.0001
        End Select
    Else
        tmp = EquivalentDoseconverter(val, src, 1)
        EquivalentDoseconverter = EquivalentDoseconverter(tmp, 1, dest)
    End If
End Function

'Function to convert between force units
'To chose source and destination units (default destination = 1 [Newton])
'1  : N (Newton)
'2  : MN (meganewton)
'3  : kN (kilonewton)
'4  : hN (hectonewton)
'5  : daN (decanewton)
'6  : dN (decinewton)
'7  : mN (millinewton)
'8  : Mdyn (megadyne)
'9  : kdyn (kilodyne)
'10 : dyn (dyne)
'11 : mdyn (millidyne)
'12 : µdyn (microdyn)
'13 : kp (kilopond)
'14 : lbf (pound)
'15 : pdl (poundal)
Public Function ForceConverter(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Double
    Dim tmp                             As Double
    
    If (dest < 1 Or dest > 15) Then
        dest = 1
    End If
    
    If (src < 1 Or src > 15) Then
        src = 1
    End If
    
    If (src = dest Or val = 0) Then
        ForceConverter = val
    ElseIf (src = 1) Then
        Select Case dest
            Case 2
                ForceConverter = val * 0.000001
            Case 4
                ForceConverter = val * 0.01
            Case 5, 8
                ForceConverter = val * 0.1
            Case 6
                ForceConverter = val * 10
            Case 7
                ForceConverter = val * 1000
            Case 9
                ForceConverter = val * 100
            Case 3, 10 To 12
                ForceConverter = ForceConverter(val, src, dest - 1) * 1000
            Case 13
                ForceConverter = val / 9.80665
            Case 14
                ForceConverter = val / 4.4482216152605
            Case 15
                ForceConverter = val / 0.138255
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2
                ForceConverter = val * 1000000
            Case 4
                ForceConverter = val * 100
            Case 5, 8
                ForceConverter = val * 10
            Case 6
                ForceConverter = val * 0.1
            Case 7
                ForceConverter = val * 0.001
            Case 9
                ForceConverter = val * 0.01
            Case 3, 10 To 12
                ForceConverter = ForceConverter(val, src - 1, dest) * 0.001
            Case 13
                ForceConverter = val * 9.80665
            Case 14
                ForceConverter = val * 4.4482216152605
            Case 15
                ForceConverter = val * 0.138255
        End Select
    Else
        tmp = ForceConverter(val, src, 1)
        ForceConverter = ForceConverter(tmp, 1, dest)
    End If
End Function

'Function to convert between frequency units
'To chose source and destination units (default destination = 1 [hertz])
'1  : Hz (hertz)
'2  : PHz (petahertz)
'3  : THz (terahertz)
'4  : GHz (gigahertz)
'5  : MHz (megahertz)
'6  : kHz (kilohertz)
'7  : hHz (hectohertz)
'8  : daHz (decahertz)
'9  : dHz (decihertz)
'10 : cHz (centihertz)
'11 : mHz (millihertz)
'12 : fresnel (Fresnel)
'13 : cps (cycle per second)
'14 : d(p) (day(period))
'15 : h(p) (hour(period))
'16 : min(p) (minute(period))
'17 : ks(p) (kilosecond(period))
'18 : hs(p) (hectosecond(period))
'19 : das(p) (decasecond(period))
'20 : s(p) (second(period))
'21 : ds(p) (decisecond(period))
'22 : cs(p) (centisecond(period))
'23 : ms(p) (millisecond(period))
'24 : µs(p) (microsecond(period))
'25 : ns(p) (nanosecond(period))
Public Function FrequenceConverter(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Double
    Dim tmp                             As Double
    
    If (dest < 1 Or dest > 25) Then
        dest = 1
    End If
    
    If (src < 1 Or src > 25) Then
        src = 1
    End If
    
    If (src = dest Or val = 0) Then
        FrequenceConverter = val
    ElseIf (src = 1) Then
        Select Case dest
            Case 2
                FrequenceConverter = val * 0.000000000000001
            Case 3 To 6, 25
                FrequenceConverter = FrequenceConverter(val, src, dest - 1) * 1000
            Case 7, 18
                FrequenceConverter = val * 0.01
            Case 9, 21
                FrequenceConverter = val * 10
            Case 8, 10, 11, 19, 22, 23
                FrequenceConverter = FrequenceConverter(val, src, dest - 1) * 10
            Case 12
                FrequenceConverter = val * 0.000000000001
            Case 13, 20
                FrequenceConverter = val
            Case 14
                FrequenceConverter = val / 86400
            Case 15
                FrequenceConverter = val / 3600
            Case 16
                FrequenceConverter = val / 60
            Case 17
                FrequenceConverter = val * 0.001
            Case 24
                FrequenceConverter = val * 1000000
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2
                FrequenceConverter = val * 1E+15
            Case 3 To 6, 25
                FrequenceConverter = FrequenceConverter(val, src - 1, dest) * 0.001
            Case 7, 18
                FrequenceConverter = val * 100
            Case 9, 21
                FrequenceConverter = val * 0.1
            Case 8, 10, 11, 19, 22, 23
                FrequenceConverter = FrequenceConverter(val, src - 1, dest) * 0.1
            Case 12
                FrequenceConverter = val * 1000000000000#
            Case 13, 20
                FrequenceConverter = val
            Case 14
                FrequenceConverter = val * 86400
            Case 15
                FrequenceConverter = val * 3600
            Case 16
                FrequenceConverter = val * 60
            Case 17
                FrequenceConverter = val * 1000
            Case 24
                FrequenceConverter = val * 0.000001
        End Select
    Else
        tmp = FrequenceConverter(val, src, 1)
        FrequenceConverter = FrequenceConverter(tmp, 1, dest)
    End If
End Function

'Function to convert between fuel consumption units
'To chose source and destination units (default destination = 1 [liters/100km])
'1  : l/100km     (liter per 100 kilometers)
'2  : l/km        (liter per kilometer)
'3  : Gal/100Km   (gallon per 100 kilometers)
'4  : Gal/km      (gallon per kilometer)
'5  : Km/l        (kilometer per liter)
'6  : Km/gal (US) (kilometer by gallon (US))
'7  : Km/gal (UK) (kilometer by gallon (UK))
'8  : Mpg (US)    (mille per gallon (US))
'9  : Mpg (UK     (mille per gallon (UK))
Public Function FuelConsumptionConverter(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Double
    Dim tmp                             As Double
    
    If (dest < 1 Or dest > 9) Then
        dest = 1
    End If
    
    If (src < 1 Or src > 9) Then
        src = 1
    End If
    
    If (src = dest Or val = 0) Then
        FuelConsumptionConverter = val
    ElseIf (src = 1) Then
        Select Case dest
            Case 2
                FuelConsumptionConverter = val * 0.01
            Case 3
                FuelConsumptionConverter = val * 0.26417205235815
            Case 4
                FuelConsumptionConverter = val * 2.6417205235815E-03
            Case 5
                FuelConsumptionConverter = val * 100
            Case 6
                FuelConsumptionConverter = val * 378.541178
            Case 7
                FuelConsumptionConverter = val * 454.609188
            Case 8
                FuelConsumptionConverter = val * 235.2145833
            Case 9
                FuelConsumptionConverter = val * 282.4809363
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2
                FuelConsumptionConverter = val * 100
            Case 3
                FuelConsumptionConverter = val * 3.785411784
            Case 4
                FuelConsumptionConverter = val * 378.5411784
            Case 5
                FuelConsumptionConverter = 100 / val
            Case 6
                FuelConsumptionConverter = 378.541178 / val
            Case 7
                FuelConsumptionConverter = 454.609188 / val
            Case 8
                FuelConsumptionConverter = 235.2145833 / val
            Case 9
                FuelConsumptionConverter = 282.4809363 / val
        End Select
    Else
        tmp = FuelConsumptionConverter(val, src, 1)
        FuelConsumptionConverter = FuelConsumptionConverter(tmp, 1, dest)
    End If
End Function

'Function to convert between illuminance units
'To chose source and destination units (default destination = 1 [lux])
'1  : lx (lux)
'2  : lmm2 (lumen per square meter)
'3  : lmcm2 (lument per square centimeter)
'4  : ph (phot)
'5  : nx (nox)
Public Function IlluminanceConverter(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Double
    Dim tmp                             As Double
    
    If (dest < 1 Or dest > 5) Then
        dest = 1
    End If
    
    If (src < 1 Or src > 5) Then
        src = 1
    End If
    
    If (src = dest Or val = 0) Then
        IlluminanceConverter = val
    ElseIf (src = 1) Then
        Select Case dest
            Case 2
                IlluminanceConverter = val
            Case 3, 4
                IlluminanceConverter = val * 0.0001
            Case 5
                IlluminanceConverter = val * 1000
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2
                IlluminanceConverter = val
            Case 3
                IlluminanceConverter = val * 10000
            Case 5
                IlluminanceConverter = val * 0.001
        End Select
    Else
        tmp = IlluminanceConverter(val, src, 1)
        IlluminanceConverter = IlluminanceConverter(tmp, 1, dest)
    End If
End Function

'Function to convert between mass units
'To chose source and destination units (default destination = 1 [kilogram])
'1  : kg (kilogram)
'2  : hg (hectogram)
'3  : dag (decagram)
'4  : g (gram)
'5  : dg (decigram)
'6  : cg (centigram)
'7  : mg (milligram)
'8  : µg (microgram)
'9  : kt (kiloton) (metric)
'10 : ht (hectoton) (metric)
'11 : dat (decaton) (metric)
'12 : t (ton) (metric)
'13 : dt (deciton) (metric)
'14 : ct (centiton) (metric)
'15 : mt (milliton) (metric)
'16 : amu (atomic mass unit)
'17 : carat (metric carat)
'18 : dr (dram)
'19 : gr (grain)
'20 : hundredweight (UK)
'21 : oz (ounce)
'22 : dwt (pennyweight)
'23 : lb (pound)
'24 : quarter
'25 : stone
'26 : ton (long)
'27 : ton (short)
'28 : troy ounce
Public Function MassConverter(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Double
    Dim tmp                             As Double
    
    If (dest < 1 Or dest > 28) Then
        dest = 1
    End If
    
    If (src < 1 Or src > 28) Then
        src = 1
    End If
    
    If (src = dest Or val = 0) Then
        MassConverter = val
    ElseIf (src = 1) Then
        Select Case dest
            Case 2
                MassConverter = val * 10
            Case 3 To 7
                MassConverter = MassConverter(val, src, dest - 1) * 10
            Case 8
                MassConverter = val * 1000000000#
            Case 9
                MassConverter = val * 0.0000001
            Case 10 To 14
                MassConverter = MassConverter(val, src, dest + 1) * 0.1
            Case 15
                MassConverter = val
            Case 16
                MassConverter = val * 6.022136651675E+26
            Case 17
                MassConverter = val * 5000
            Case 18
                MassConverter = val * 564.3833911933
            Case 19
                MassConverter = val * 1000
            Case 20
                MassConverter = val * 0.01968413055222
            Case 21
                MassConverter = val * 35.27396194958
            Case 22
                MassConverter = val * 643.0149313726
            Case 23
                MassConverter = val * 2.204622621849
            Case 24
                MassConverter = val * 0.07873652220889
            Case 25
                MassConverter = val * 0.1574730444178
            Case 26
                MassConverter = val * 9.842065276111E-04
            Case 27
                MassConverter = val * 0.001102311310924
            Case 28
                MassConverter = val * 32.15074656863
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2
                MassConverter = val * 0.1
            Case 3 To 7
                MassConverter = MassConverter(val, src - 1, dest) * 0.1
            Case 8
                MassConverter = val * 0.000000001
            Case 9
                MassConverter = val * 10000000
            Case 10 To 14
                MassConverter = MassConverter(val, src - 1, dest) * 10
            Case 15
                MassConverter = val
            Case 16
                MassConverter = val / 6.022136651675E+26
            Case 17
                MassConverter = val * 0.0002
            Case 19
                MassConverter = val * 0.001
            Case 18
                MassConverter = val / 564.3833911933
            Case 20
                MassConverter = val / 0.01968413055222
            Case 21
                MassConverter = val / 35.27396194958
            Case 22
                MassConverter = val / 643.0149313726
            Case 23
                MassConverter = val / 2.204622621849
            Case 24
                MassConverter = val / 0.07873652220889
            Case 25
                MassConverter = val / 0.1574730444178
            Case 26
                MassConverter = val / 9.842065276111E-04
            Case 27
                MassConverter = val / 0.001102311310924
            Case 28
                MassConverter = val / 32.15074656863
        End Select
    Else
        tmp = MassConverter(val, src, 1)
        MassConverter = MassConverter(tmp, 1, dest)
    End If
End Function

'Function to convert between memory size units
'To chose source and destination units (default destination = 1 [byte])
'1  : B (byte)
'2  : bit (bit)
'3  : KB (kilobyte)
'4  : MB (megabyte)
'5  : GB (gigabyte)
'6  : TB (terabyte)
'7  : PB (petabyte)
'8  : EB (exabyte)
'9  : ZB (zettabyte)
'10 : YB (yottabyte)
'11 : KiB (kibibyte)
'12 : MiB (mebibyte)
'13 : GiB (gibibyte)
'14 : TiB (tebibyte)
'15 : PiB (pebibyte)
'16 : EiB (exbibyte)
'17 : ZiB (zebibyte)
'18 : Yib (yobibyte)
'19 : kbit (kilobit)
'20 : Mbit (megabit)
'21 : Gbit (gigabit)
'22 : Tbit (terabit)
'23 : Pbit (petabit)
'24 : Ebit (exabit)
'25 : Zbit (zetabit)
'26 : Ybit (yottabit)
'27 : kibit (kibibit)
'28 : Mibit (mebibit)
'29 : Gibit (gibibit)
'30 : Tibit (tibibit)
'31 : Pibit (pebibit)
'32 : Eibit (exbibit)
'33 : Zibit (zebibit)
'34 : Yibit (yobibit)
Public Function MemorySizeConverter(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Double
    Dim tmp                             As Double
    
    If (dest < 1 Or dest > 34) Then
        dest = 1
    End If
    
    If (src < 1 Or src > 34) Then
        src = 1
    End If
    
    If (src = dest Or val = 0) Then
        MemorySizeConverter = val
    ElseIf (src = 1) Then
        Select Case dest
            Case 2
                MemorySizeConverter = val * 8
            Case 3
                MemorySizeConverter = val * 0.001
            Case 4 To 10, 20 To 26
                MemorySizeConverter = MemorySizeConverter(val, src, dest - 1) / 1000
            Case 11
                MemorySizeConverter = val / 1024
            Case 12 To 18, 28 To 34
                MemorySizeConverter = MemorySizeConverter(val, src, dest - 1) / 1024
            Case 19
                MemorySizeConverter = val * 0.008
            Case 27
                MemorySizeConverter = val / 128
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2
                MemorySizeConverter = val * 0.125
            Case 3
                MemorySizeConverter = val * 1000
            Case 4 To 10, 20 To 26
                MemorySizeConverter = MemorySizeConverter(val, src, dest - 1) * 1000
            Case 11
                MemorySizeConverter = val * 1024
            Case 12 To 18, 28 To 34
                MemorySizeConverter = MemorySizeConverter(val, src, dest - 1) * 1024
            Case 19
                MemorySizeConverter = val * 125
            Case 27
                MemorySizeConverter = val * 128
        End Select
    Else
        tmp = MemorySizeConverter(val, src, 1)
        MemorySizeConverter = MemorySizeConverter(tmp, 1, dest)
    End If
End Function

'Function to convert between Power units
'To chose source and destination units (default destination = 1 [Watt])
'1  : w (Watt)
'2  : kW (kilowatt)
'3  : mW (milliwatt)
'4  : MW (megawatt)
'5  : GW (gigawatt)
'6  : TW (terawatt)
'7  : PW (petawatt)
'8  : µW (microwatt)
'9  : nW (nanowatt)
'10 : pW (picowatt)
'11 : fW (femtowatt)
'12 : zW (zptowatt)
'13 : hp (metric horsepower)
'14 : bhp (mechanical horsepower)
'15 : refrigerationton (refigeration ton)
Public Function PowerConverter(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Double
    Dim tmp                             As Double
    
    If (dest < 1 Or dest > 15) Then
        dest = 1
    End If
    
    If (src < 1 Or src > 15) Then
        src = 1
    End If
    
    If (src = dest Or val = 0) Then
        PowerConverter = val
    ElseIf (src = 1) Then
        Select Case dest
            Case 2
                PowerConverter = val * 0.001
            Case 3
                PowerConverter = val * 1000
            Case 4
                PowerConverter = val * 0.000001
            Case 5 To 7
                PowerConverter = PowerConverter(val, src, dest - 1) * 0.001
            Case 8
                PowerConverter = val * 1000000
            Case 9 To 12
                PowerConverter = PowerConverter(val, src, dest - 1) * 1000
            Case 13
                PowerConverter = val / 735.39875
            Case 14
                PowerConverter = val / 745.66272
            Case 15
                PowerConverter = val / 3516.8528
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2
                PowerConverter = val * 1000
            Case 3
                PowerConverter = val * 0.001
            Case 4
                PowerConverter = val * 1000000
            Case 5 To 7
                PowerConverter = PowerConverter(val, src, dest - 1) * 1000
            Case 8
                PowerConverter = val * 0.000001
            Case 9 To 12
                PowerConverter = PowerConverter(val, src, dest - 1) * 0.001
            Case 13
                PowerConverter = val * 735.39875
            Case 14
                PowerConverter = val * 745.66272
            Case 15
                PowerConverter = val * 3516.8528
        End Select
    Else
        tmp = PowerConverter(val, src, 1)
        PowerConverter = PowerConverter(tmp, 1, dest)
    End If
End Function

'Function to convert between Pressure units
'To chose source and destination units (default destination = 1 [Pascal])
'1  : Pa (pascal)
'2  : bar (bar)
'3  : hPa (hectopascal)
'4  : kPa (kilopascal)
'5  : MPa (megapascal)
'6  : mbar (millibar)
'7  : at (technical atmosphere)
'8  : atm (physical atmosphere)
'9  : Nm-2 (Newton per square meter)
'10 : psi (pound per square inch)
'11 : Torr (Torr)
'12 : mmHg (millimeter of mercury)
'13 : mmH2O (millimeters water column)
Public Function PressureConverter(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Double
    Dim tmp                             As Double
    
    If (dest < 1 Or dest > 13) Then
        dest = 1
    End If
    
    If (src < 1 Or src > 13) Then
        src = 1
    End If
    
    If (src = dest Or val = 0) Then
        PressureConverter = val
    ElseIf (src = 1) Then
        Select Case dest
            Case 2
                PressureConverter = val * 0.00001
            Case 3, 6
                PressureConverter = val * 0.01
            Case 4, 5
                PressureConverter = PressureConverter(val, src, dest - 1) * 0.1
            Case 7
                PressureConverter = val / 98066.5
            Case 8
                PressureConverter = val / 101325
            Case 9
                PressureConverter = val
            Case 10
                PressureConverter = val / 6894.757293168
            Case 11, 12
                PressureConverter = val / 133.322
            Case 13
                PressureConverter = val / 9.80665
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2
                PressureConverter = val * 100000
            Case 3, 6
                PressureConverter = val * 100
            Case 4, 5
                PressureConverter = PressureConverter(val, src - 1, dest) * 10
            Case 7
                PressureConverter = val * 98066.5
            Case 8
                PressureConverter = val * 101325
            Case 9
                PressureConverter = val
            Case 10
                PressureConverter = val * 6894.757293168
            Case 11, 12
                PressureConverter = val * 133.322
            Case 13
                PressureConverter = val * 9.80665
        End Select
    Else
        tmp = PressureConverter(val, src, 1)
        PressureConverter = PressureConverter(tmp, 1, dest)
    End If
End Function

'Function to convert between radioactive units
'To chose source and destination units (default destination is 1 : Bq [becquerel]):
'1  : Bq (becquerel)
'2  : Ci (curie)
'3  : PBq (petabecquerel)
'4  : TBq (terabecquerel)
'5  : GBq (gigabecquerel)
'6  : MBq (megabecquerel)
'7  : kBq (kilobecquerel)
'8  : hBq (hectobecquerel)
'9  : daBq (decabecquerel)
'10 : dBq (decibecquerel)
'11 : cBq (centibecquerel)
'12 : mBq (millibecquerel)
'13 : kCi (kilocurie)
'14 : hCi (hectocurie)
'15 : daCi (decacurie)
'16 : dCi (decicurie)
'17 : cCi (centicurie)
'18 : mCi (millicurie)
'19 : µCi (microcurie)
'20 : pCi (picocurie)
'21 : Rd (rutherford)
'22 : dpm (disintegrations per minute)
Public Function RadioactiveConverter(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Double
    Dim tmp                             As Double
    
    If (dest < 1 Or dest > 22) Then
        dest = 1
    End If
    
    If (src < 1 Or src > 22) Then
        src = 1
    End If
    
    If (dest = src Or val = 0) Then
        RadioactiveConverter = dest
    ElseIf (src = 1) Then
        Select Case dest
            Case 2
                RadioactiveConverter = val / 37000000000#
            Case 3
                RadioactiveConverter = val * 0.000000000000001
            Case 4 To 7
                RadioactiveConverter = RadioactiveConverter(val, src, dest - 1) * 1000
            Case 8
                RadioactiveConverter = val * 0.01
            Case 9
                RadioactiveConverter = val * 0.1
            Case 10
                RadioactiveConverter = val * 10
            Case 11, 12, 14, 15, 17, 18
                RadioactiveConverter = RadioactiveConverter(val, src, dest - 1) * 10
            Case 13
                RadioactiveConverter = val / 37000000000000#
            Case 16
                RadioactiveConverter = val / 3700000000#
            Case 19
                RadioactiveConverter = val / 37000
            Case 20
                RadioactiveConverter = val / 0.037
            Case 21
                RadioactiveConverter = val / 0.000001
            Case 22
                RadioactiveConverter = val * 60
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2
                RadioactiveConverter = val * 37000000000#
            Case 3
                RadioactiveConverter = val * 1E+15
            Case 4 To 7
                RadioactiveConverter = RadioactiveConverter(val, src, dest - 1) / 1000
            Case 8
                RadioactiveConverter = val * 100
            Case 9
                RadioactiveConverter = val * 10
            Case 10
                RadioactiveConverter = val * 0.1
            Case 11, 12, 14, 15, 17, 18
                RadioactiveConverter = RadioactiveConverter(val, src - 1, dest - 1) * 0.1
            Case 13
                RadioactiveConverter = val * 37000000000000#
            Case 16
                RadioactiveConverter = val * 3700000000#
            Case 19
                RadioactiveConverter = val * 37000
            Case 20
                RadioactiveConverter = val * 0.037
            Case 21
                RadioactiveConverter = val * 1000000
            Case 22
                RadioactiveConverter = val / 60
        End Select
    Else
        tmp = RadioactiveConverter(val, src, 1)
        RadioactiveConverter = RadioactiveConverter(tmp, 1, dest)
    End If
End Function

'Function to convert between speed units
'To chose source and destination units (default destination is 1 : mps [meters per second]):
'1  : mps (meters per second)
'2  : kph (kilometer per hour)
'3  : kps (kilometer per second)
'4  : m/min (meters per minute)
'5  : mps (miles per second)
'6  : mph (miles per hour)
'7  : fps (foot per second)
'8  : ft/min (foot per minute)
'9  : sec/km (second per kilometer)
'10 : sec/hm (second per 100 meters)
'11 : kt (knot)
'12 : seamiles/hour (nautical mile per hour)
'13 : c (celerity/speed of light)
'14 : Ma (Mach speed/speed of sound in the air)
Public Function SpeedConverter(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Double
    Dim tmp                             As Double
    
    If (dest < 1 Or dest > 14) Then
        dest = 1
    End If
    
    If (src < 1 Or src > 14) Then
        src = 1
    End If
    
    If (dest = src Or val = 0) Then
        SpeedConverter = dest
    ElseIf (src = 1) Then
        Select Case dest
            Case 2
                SpeedConverter = val * 3.6
            Case 3
                SpeedConverter = val * 0.001
            Case 4
                SpeedConverter = val * 60
            Case 5
                SpeedConverter = val / 1609.344
            Case 6
                SpeedConverter = val / 0.44704
            Case 7
                SpeedConverter = val * 3.28084
            Case 8
                SpeedConverter = val * 196.8504
            Case 9
                SpeedConverter = val * 1000
            Case 10
                SpeedConverter = val * 100
            Case 11
                SpeedConverter = val / 0.51444444444
            Case 12
                SpeedConverter = val / 0.51444
            Case 13
                SpeedConverter = val / 299792458
            Case 14
                SpeedConverter = val / 340
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2
                SpeedConverter = val / 3.6
            Case 3
                SpeedConverter = val * 1000
            Case 4
                SpeedConverter = val / 60
            Case 5
                SpeedConverter = val * 1609.344
            Case 6
                SpeedConverter = val * 0.44704
            Case 7
                SpeedConverter = val / 3.28084
            Case 8
                SpeedConverter = val / 196.8504
            Case 9
                SpeedConverter = val * 0.001
            Case 10
                SpeedConverter = val * 0.01
            Case 11
                SpeedConverter = val * 0.51444444444
            Case 12
                SpeedConverter = val * 0.51444
            Case 13
                SpeedConverter = val * 299792458
            Case 14
                SpeedConverter = val * 340
        End Select
    Else
        tmp = SpeedConverter(val, src, 1)
        SpeedConverter = SpeedConverter(tmp, 1, dest)
    End If
End Function

'Function to convert romans (FR : nombres romains) to arabic numerals
'If string isn't a valid roman numeral then return -1
Public Function RomanToArabic(ByVal roman As String) As Long
    Dim cpt                             As Long
    Dim unit                            As Long
    Dim oldUnit                         As Long
    
    oldUnit = 1000
    
    roman = UCase(Trim(roman))
    
    If (roman = "0") Then
        RomanToArabic = 0
    Else
        For cpt = 1 To Len(roman)
            Select Case Mid(roman, cpt, 1)
                Case "I"
                    unit = 1
                Case "V"
                    unit = 5
                Case "X"
                    unit = 10
                Case "L"
                    unit = 50
                Case "C"
                    unit = 100
                Case "D"
                    unit = 500
                Case "M"
                    unit = 1000
                Case Else
                    'invalid roman string because invalid character is detected
                    RomanToArabic = -1
                    Exit Function
            End Select
            
            If (unit > oldUnit) Then
                RomanToArabic = RomanToArabic - 2 * oldUnit
            End If
            
            RomanToArabic = RomanToArabic + unit
            oldUnit = unit
        Next cpt
    End If
End Function

'Function to convert between temperature units (Kelvin, Celsius, Fahreneit
'To chose source and destination units (default dest = 1 [celsius]) :
'1 : Celsius
'2 : Kelvin
'3 : Fahrenheit
'4 : Rankine
'5 : Reaumur
Public Function TemperatureConverter(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Single
    Dim tmp                             As Double
    
    If (dest < 1 Or dest > 5) Then
        dest = 1
    End If
    
    If (src < 1 Or src > 5) Then
        src = 1
    End If
    
    If (src = dest) Then
        TemperatureConverter = val
    ElseIf (src = 1) Then
        Select Case dest
            Case 2
                TemperatureConverter = val + 273.15
            Case 3
                TemperatureConverter = (val * 1.8) + 32
            Case 4
                TemperatureConverter = (val * 1.8) + 491.67
            Case 5
                TemperatureConverter = val * 0.8
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2
                TemperatureConverter = val - 273.15
            Case 3
                TemperatureConverter = IIf(val = 32, 0, (val - 32) / 1.8)
            Case 4
                TemperatureConverter = IIf(val = 491.67, 0, (val - 491.67) / 1.8)
            Case 5
                TemperatureConverter = val * 1.25
        End Select
    Else
        tmp = TemperatureConverter(val, src, 1)
        TemperatureConverter = TemperatureConverter(tmp, 1, dest)
    End If
End Function

'Function to convert between volume units
'To chose source and destination units (default destination is 1 : l [liter]):
'1  : l(liter)
'2  : m3 (cubic meter)
'3  : kl (kiloliter)
'4  : hl (hectoliter)
'5  : dal (decaliter)
'6  : dl (deciliter)
'7  : cl (centiliter)
'8  : ml (milliliter)
'9  : km3 (cubic meter)
'10 : hm3 (cubic hectometer)
'11 : dam3 (cubic decameter)
'12 : dm3 (cubic decimeter)
'13 : cm3 (cubic centimeter)
'14 : mm3 (cubic millimeter)
'15 : ft3 (cubic foot)
'16 : in3 (cubic inch)
'17 : gallon (imperial gallon)
'18 : pint (pint)
'19 : gallon (U.S. liquid gallon)
'20 : pint (U.S. liquid pint)
'21 : fl.oz. (fluid ounce)
'22 : tbs (tablespoon)
'23 : tsp (teaspoon)
'24 : bbl (oil barrel)
'25 : imp.bsh. (imperial bushel)
'26 : U.S.bsh (U.S. bushel)
'27 : cup (metric cup)
'28 : imp.cup (imperial cup)
'29 : US.cup (U.S. cup)
Public Function VolumeConverter(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Double
    Dim tmp                             As Double
    
    If (dest < 1 Or dest > 29) Then
        dest = 1
    End If
    
    If (src < 1 Or src > 29) Then
        src = 1
    End If
    
    If (dest = src Or val = 0) Then
        VolumeConverter = dest
    ElseIf (src = 1) Then
        Select Case dest
            Case 2, 3
                VolumeConverter = val * 0.001
            Case 4, 5, 7, 8
                VolumeConverter = VolumeConverter(val, src, dest - 1) * 10
            Case 6
                VolumeConverter = val * 10
            Case 9
                VolumeConverter = val * 0.000000000001
            Case 10, 11, 13, 14
                VolumeConverter = VolumeConverter(val, src, dest - 1) * 1000
            Case 12
                VolumeConverter = val
            Case 15
                VolumeConverter = val / 28.316846592
            Case 16
                VolumeConverter = val / 0.016387064
            Case 17
                VolumeConverter = val / 4.54609
            Case 18
                VolumeConverter = val / 0.56826125
            Case 19
                VolumeConverter = val / 3.785411784
            Case 20
                VolumeConverter = val / 0.473176473
            Case 21
                VolumeConverter = val / 0.03
            Case 22
                VolumeConverter = val / 0.015
            Case 23
                VolumeConverter = val * 200
            Case 24
                VolumeConverter = val / 158.9873
            Case 25
                VolumeConverter = val / 36.368722255958
            Case 26
                VolumeConverter = val / 35.23907016688
            Case 27
                VolumeConverter = val * 4
            Case 28
                VolumeConverter = val * 3.5195077544697
            Case 29
                VolumeConverter = val * 4.2267528377304
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2, 3
                VolumeConverter = val * 1000
            Case 4, 5, 7, 8
                VolumeConverter = VolumeConverter(val, src, dest - 1) * 0.1
            Case 6
                VolumeConverter = val * 0.1
            Case 9
                VolumeConverter = val * 1000000000000#
            Case 10, 11, 13, 14
                VolumeConverter = VolumeConverter(val, src, dest - 1) * 0.001
            Case 12
                VolumeConverter = val
            Case 15
                VolumeConverter = val * 28.316846592
            Case 16
                VolumeConverter = val * 0.016387064
            Case 17
                VolumeConverter = val * 4.54609
            Case 18
                VolumeConverter = val * 0.56826125
            Case 19
                VolumeConverter = val * 3.785411784
            Case 20
                VolumeConverter = val * 0.473176473
            Case 21
                VolumeConverter = val * 0.03
            Case 22
                VolumeConverter = val * 0.015
            Case 23
                VolumeConverter = val * 0.005
            Case 24
                VolumeConverter = val * 158.9873
            Case 25
                VolumeConverter = val * 36.368722255958
            Case 26
                VolumeConverter = val * 35.23907016688
            Case 27
                VolumeConverter = val * 0.25
            Case 28
                VolumeConverter = val / 3.5195077544697
            Case 29
                VolumeConverter = val / 4.2267528377304
        End Select
    Else
        tmp = VolumeConverter(val, src, 1)
        VolumeConverter = VolumeConverter(tmp, 1, dest)
    End If
End Function

'Function to convert between volumetric flow units
'To chose source and destination units (default destination = 1 [liter per second])
'1  : l/s (liter per second)
'2  : m3/year (cubic meter per year)
'3  : m3/h (cubic meter per hour)
'4  : m3/min (cubic meter per hour)
'5  : m3/s (cubic meter per second)
'6  : Ml/day (megaliter per day)
'7  : Ml/s (megaliter per second)
'8  : L/min (liter per minute)
'9  : ft3/year (cubic foot per year)
'10 : ft3/s (cubic foot per second)
'11 : gpd (imperial gallon per day)
'12 : gpm (imperial gallon per minute)
'13 : US gpd (US gallon per day)
'14 : US gpm (US gallon per minute)
Public Function VolumetricFlowConverter(ByVal val As Double, ByVal src As Integer, Optional ByVal dest As Integer = 1) As Double
    Dim tmp                             As Double
    
    If (dest < 1 Or dest > 18) Then
        dest = 1
    End If
    
    If (src < 1 Or src > 18) Then
        src = 1
    End If
    
    If (src = dest Or val = 0) Then
        VolumetricFlowConverter = val
    ElseIf (src = 1) Then
        Select Case dest
            Case 2
                VolumetricFlowConverter = val * 31536
            Case 3
                VolumetricFlowConverter = val * 3.6
            Case 4
                VolumetricFlowConverter = val * 0.06
            Case 5
                VolumetricFlowConverter = val * 0.001
            Case 6
                VolumetricFlowConverter = val * 0.0864
            Case 7
                VolumetricFlowConverter = val * 0.000001
            Case 8
                VolumetricFlowConverter = val * 60
            Case 9
                VolumetricFlowConverter = val * 1113676.621
            Case 10
                VolumetricFlowConverter = val / 28.316846592
            Case 11
                VolumetricFlowConverter = val * 19005.304
            Case 12
                VolumetricFlowConverter = val / 0.0757683211
            Case 13
                VolumetricFlowConverter = val * 22820
            Case 14
                VolumetricFlowConverter = val / 0.06309
        End Select
    ElseIf (dest = 1) Then
        Select Case src
            Case 2
                VolumetricFlowConverter = val / 31536
            Case 3
                VolumetricFlowConverter = val / 3.6
            Case 4
                VolumetricFlowConverter = val / 0.06
            Case 5
                VolumetricFlowConverter = val / 0.001
            Case 6
                VolumetricFlowConverter = val / 0.0864
            Case 7
                VolumetricFlowConverter = val / 0.000001
            Case 8
                VolumetricFlowConverter = val / 60
            Case 9
                VolumetricFlowConverter = val / 1113676.621
            Case 10
                VolumetricFlowConverter = val * 28.316846592
            Case 11
                VolumetricFlowConverter = val / 19005.304
            Case 12
                VolumetricFlowConverter = val * 0.0757683211
            Case 13
                VolumetricFlowConverter = val / 22820
            Case 14
                VolumetricFlowConverter = val * 0.06309
        End Select
    Else
        tmp = VolumetricFlowConverter(val, src, 1)
        VolumetricFlowConverter = VolumetricFlowConverter(tmp, 1, dest)
    End If
End Function
