'**************************************************************************************************************************************************
'V Excelu vloženo na "list2" pojmenovaného "zadani"

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim KeyCells As Range

    ' Tato buňka nebo oblast když se změní, spouští se makro
    Set KeyCells = Sheets("zadani").Range("C5:E32")
    
    If Not Application.Intersect(KeyCells, Range(Target.Address)) Is Nothing Then
        ' Následující makro se změnou buňky nebo oblasti spouští
        Call vypocet
    End If
End Sub

Sub vypocet()
    Call ZobrazeniZadani
    Call zobrazeniObrazku
    
    '---------------------------------------------------------------------------------
    'volba výpočtů podle typu (jestli a které světlíky počítat
    Dim typ As String
    If Worksheets("zadani").Range("C8").Value = "jednokřídlové" Then
        typ = Worksheets("zadani").Range("C9").Value
    ElseIf Worksheets("zadani").Range("C8").Value = "dvoukřídlové" Then
        typ = Worksheets("zadani").Range("C10").Value
    End If
        
        Select Case typ
            Case "1.1L", "1.1P", "2A.1L", "2A.1P"
                Call vypocetDveri
            Case "1.2L", "1.2P", "2A.2L", "2A.2P"
                Call vypocetDveri
                Call vypocetSvetlikuB
            Case "1.3L", "2A.3L"
                Call vypocetDveri
                Call vypocetSvetlikuD
            Case "1.3P", "2A.3P"
                Call vypocetDveri
                Call vypocetSvetlikuC
            Case "1.4L", "1.4P", "2A.4L", "2A.4P"
                Call vypocetDveri
                Call vypocetSvetlikuC
                Call vypocetSvetlikuD
            Case "1.5L", "2A.5L"
                Call vypocetDveri
                Call vypocetSvetlikuB
                Call vypocetSvetlikuD
            Case "1.5P", "2A.5P"
                Call vypocetDveri
                Call vypocetSvetlikuB
                Call vypocetSvetlikuC
            Case "1.6L", "1.6P", "2A.6L", "2A.6P"
                Call vypocetDveri
                Call vypocetSvetlikuB
                Call vypocetSvetlikuC
                Call vypocetSvetlikuD
        End Select
        
    Call vypocetKompletniSestavy
    
End Sub
Sub vypocetDveri()
    'volby výpočtu podle kategorie dveří (Aktiv-Příčkové-Excellent)
    If Worksheets("zadani").Range("kategorie").Value = "AKTIV 77" Then
        If Worksheets("zadani").Range("provedeni").Value = "jednokřídlové" Then
            Call vypocetDveriAktiv_1
        ElseIf Worksheets("zadani").Range("provedeni").Value = "dvoukřídlové" Then
            Call vypocetDveriAktiv_2
        End If
    ElseIf Worksheets("zadani").Range("kategorie").Value = "AKTIV 77 - příčkové" Then
        Call vypocetDveriPrickove
    ElseIf Worksheets("zadani").Range("kategorie").Value = "AKTIV 77 - EXCELLENT" Then
        Call vypocetDveriExcellent
    End If
End Sub

'**************************************************************************************************************************************************
'V Excelu vloženo do modelu "UpravaZobrazeniZadani"

Sub Reset()
    Worksheets("zadani").Range("C5").Value = "AKTIV 77"
    Worksheets("zadani").Range("C6").Value = "Průběžný"
    Worksheets("zadani").Range("C7").Value = "jednostranné"
    Worksheets("zadani").Range("C8").Value = "jednokřídlové"
    Worksheets("zadani").Range("C9").Value = "1.1L"
    Worksheets("zadani").Range("C10").Value = "2A.1L"
    Worksheets("zadani").Range("C11").Value = "B1"
End Sub

Sub ZobrazeniZadani()
    'rozdělení podle Kategorie - Aktiv / příčkové / Excellent
    kategorie = Worksheets("zadani").Range("C5").Value
    aktiv = Worksheets("List3").Range("E1").Value
    prickove = Worksheets("List3").Range("E2").Value
    excellent = Worksheets("List3").Range("E3").Value
    
    If kategorie = aktiv Then
        Worksheets("zadani").Rows("6").Hidden = False
        Worksheets("zadani").Rows("7").Hidden = True
        Worksheets("zadani").Rows("8").Hidden = False
        Worksheets("zadani").Rows("9").Hidden = False
        Worksheets("zadani").Rows("10").Hidden = True
        Worksheets("zadani").Rows("11").Hidden = True
        Worksheets("zadani").Rows("16").Hidden = False
        'jednokřídlové vs. dvoukřídlové
        If Worksheets("zadani").Range("C8").Value = "jednokřídlové" Then
            Worksheets("zadani").Rows("9").Hidden = False
            Worksheets("zadani").Rows("10").Hidden = True
        ElseIf Worksheets("zadani").Range("C8").Value = "dvoukřídlové" Then
            Worksheets("zadani").Rows("9").Hidden = True
            Worksheets("zadani").Rows("10").Hidden = False
        End If
    ElseIf kategorie = prickove Then
        Worksheets("zadani").Rows("6").Hidden = False
        Worksheets("zadani").Rows("7").Hidden = True
        Worksheets("zadani").Rows("8").Hidden = True
        Worksheets("zadani").Rows("9").Hidden = False
        Worksheets("zadani").Rows("10").Hidden = True
        Worksheets("zadani").Rows("11").Hidden = False
        Worksheets("zadani").Rows("16").Hidden = True
        
    ElseIf kategorie = excellent Then
        Worksheets("zadani").Rows("6").Hidden = True
        Worksheets("zadani").Rows("7").Hidden = False
        Worksheets("zadani").Rows("8").Hidden = True
        Worksheets("zadani").Rows("9").Hidden = False
        Worksheets("zadani").Rows("10").Hidden = True
        Worksheets("zadani").Rows("11").Hidden = True
        Worksheets("zadani").Rows("16").Hidden = True
    End If
    
    'rozdělení podle Typu
    If Worksheets("zadani").Range("C8").Value = "jednokřídlové" Then
        Dim typ As String
        typ = Worksheets("zadani").Range("C9").Value
        
        Select Case typ
            Case "1.1L", "1.1P"
                Worksheets("zadani").Rows("18:32").Hidden = True
            Case "1.2L", "1.2P"
                Worksheets("zadani").Rows("18:22").Hidden = False
                Worksheets("zadani").Rows("23:32").Hidden = True
            Case "1.3L"
                Worksheets("zadani").Rows("18:27").Hidden = True
                Worksheets("zadani").Rows("28:32").Hidden = False
            Case "1.3P"
                Worksheets("zadani").Rows("18:22").Hidden = True
                Worksheets("zadani").Rows("23:27").Hidden = False
                Worksheets("zadani").Rows("28:32").Hidden = True
            Case "1.4L", "1.4P"
                Worksheets("zadani").Rows("18:22").Hidden = True
                Worksheets("zadani").Rows("23:32").Hidden = False
            Case "1.5L"
                Worksheets("zadani").Rows("18:22").Hidden = False
                Worksheets("zadani").Rows("23:27").Hidden = True
                Worksheets("zadani").Rows("28:32").Hidden = False
            Case "1.5P"
                Worksheets("zadani").Rows("18:27").Hidden = False
                Worksheets("zadani").Rows("28:32").Hidden = True
            Case "1.6L", "1.6P"
                Worksheets("zadani").Rows("18:32").Hidden = False
        End Select
    Else
        Dim typ_2 As String
        typ_2 = Worksheets("zadani").Range("C10").Value

        Select Case typ_2
            Case "2A.1L", "2A.1P"
                Worksheets("zadani").Rows("18:32").Hidden = True
            Case "2A.2L", "2A.2P"
                Worksheets("zadani").Rows("18:22").Hidden = False
                Worksheets("zadani").Rows("23:32").Hidden = True
            Case "2A.3L"
                Worksheets("zadani").Rows("18:27").Hidden = True
                Worksheets("zadani").Rows("28:32").Hidden = False
            Case "2A.3P"
                Worksheets("zadani").Rows("18:22").Hidden = True
                Worksheets("zadani").Rows("23:27").Hidden = False
                Worksheets("zadani").Rows("28:32").Hidden = True
            Case "2A.4L", "2A.4P"
                Worksheets("zadani").Rows("18:22").Hidden = True
                Worksheets("zadani").Rows("23:32").Hidden = False
            Case "2A.5L"
                Worksheets("zadani").Rows("18:22").Hidden = False
                Worksheets("zadani").Rows("23:27").Hidden = True
                Worksheets("zadani").Rows("28:32").Hidden = False
            Case "2A.5P"
                Worksheets("zadani").Rows("18:27").Hidden = False
                Worksheets("zadani").Rows("28:32").Hidden = True
            Case "2A.6L", "2A.6P"
                Worksheets("zadani").Rows("18:32").Hidden = False
        End Select
    End If
    
    'Výmaz pomocných hodnot na záložce zadani - řádek 100
    Worksheets("zadani").Range("G101:G111").Value = "0"
    Worksheets("zadani").Range("J101:J111").Value = "0"
    Worksheets("zadani").Range("L101:L111").Value = "0"
    Worksheets("zadani").Range("C101:C128").Value = "0"
    'Výmaz vypočtených hodnot
    Worksheets("zadani").Range("G5:G32").Value = ""

End Sub
Sub zobrazeniObrazku()
    'nejdřív vše zneviditelnit
    Worksheets("zadani").Shapes("obr-Aktiv").Visible = False
    Worksheets("zadani").Shapes("obr-Prickove").Visible = False
    Worksheets("zadani").Shapes("obr-Excellent").Visible = False
    '-----------------------------------------------------------
    Worksheets("zadani").Shapes("obr-1.1L").Visible = False
    Worksheets("zadani").Shapes("obr-1.1P").Visible = False
    Worksheets("zadani").Shapes("obr-1.2L").Visible = False
    Worksheets("zadani").Shapes("obr-1.2P").Visible = False
    Worksheets("zadani").Shapes("obr-1.3L").Visible = False
    Worksheets("zadani").Shapes("obr-1.3P").Visible = False
    Worksheets("zadani").Shapes("obr-1.4L").Visible = False
    Worksheets("zadani").Shapes("obr-1.4P").Visible = False
    Worksheets("zadani").Shapes("obr-1.5L").Visible = False
    Worksheets("zadani").Shapes("obr-1.5P").Visible = False
    Worksheets("zadani").Shapes("obr-1.6L").Visible = False
    Worksheets("zadani").Shapes("obr-1.6P").Visible = False
    '-----------------------------------------------------------
    Worksheets("zadani").Shapes("obr-2A.1L").Visible = False
    Worksheets("zadani").Shapes("obr-2A.1P").Visible = False
    Worksheets("zadani").Shapes("obr-2A.2L").Visible = False
    Worksheets("zadani").Shapes("obr-2A.2P").Visible = False
    Worksheets("zadani").Shapes("obr-2A.3L").Visible = False
    Worksheets("zadani").Shapes("obr-2A.3P").Visible = False
    Worksheets("zadani").Shapes("obr-2A.4L").Visible = False
    Worksheets("zadani").Shapes("obr-2A.4P").Visible = False
    Worksheets("zadani").Shapes("obr-2A.5L").Visible = False
    Worksheets("zadani").Shapes("obr-2A.5P").Visible = False
    Worksheets("zadani").Shapes("obr-2A.6L").Visible = False
    Worksheets("zadani").Shapes("obr-2A.6P").Visible = False
    '------------------------------------------------------------
    Worksheets("zadani").Shapes("obr-B1").Visible = False
    Worksheets("zadani").Shapes("obr-B2").Visible = False
    Worksheets("zadani").Shapes("obr-B3").Visible = False
    Worksheets("zadani").Shapes("obr-C1").Visible = False
    Worksheets("zadani").Shapes("obr-C2").Visible = False
    Worksheets("zadani").Shapes("obr-C3").Visible = False
    Worksheets("zadani").Shapes("obr-C4").Visible = False
    Worksheets("zadani").Shapes("obr-D1").Visible = False
    Worksheets("zadani").Shapes("obr-D2").Visible = False
    Worksheets("zadani").Shapes("obr-D3").Visible = False
    Worksheets("zadani").Shapes("obr-H1").Visible = False
    Worksheets("zadani").Shapes("obr-H2").Visible = False
    Worksheets("zadani").Shapes("obr-H3").Visible = False
    Worksheets("zadani").Shapes("obr-H4").Visible = False
    Worksheets("zadani").Shapes("obr-N1").Visible = False
    Worksheets("zadani").Shapes("obr-N2").Visible = False
    Worksheets("zadani").Shapes("obr-N3").Visible = False
    Worksheets("zadani").Shapes("obr-N4").Visible = False
    Worksheets("zadani").Shapes("obr-P1").Visible = False
    Worksheets("zadani").Shapes("obr-P2").Visible = False
    Worksheets("zadani").Shapes("obr-P3").Visible = False
    Worksheets("zadani").Shapes("obr-P4").Visible = False
    
    'z TL - Aktiv / Příčkové / Excellent
    If Worksheets("zadani").Range("C5").Value = "AKTIV 77" Then
        Worksheets("zadani").Shapes("obr-Aktiv").Visible = True
    ElseIf Worksheets("zadani").Range("C5").Value = "AKTIV 77 - příčkové" Then
        Worksheets("zadani").Shapes("obr-Prickove").Visible = True
    ElseIf Worksheets("zadani").Range("C5").Value = "AKTIV 77 - EXCELLENT" Then
        Worksheets("zadani").Shapes("obr-Excellent").Visible = True
    End If
    
    'typy
    If Worksheets("zadani").Range("C8").Value = "jednokřídlové" Then
        Select Case Worksheets("zadani").Range("C9").Value
            Case "1.1L"
                Worksheets("zadani").Shapes("obr-1.1L").Visible = True
            Case "1.1P"
                Worksheets("zadani").Shapes("obr-1.1P").Visible = True
            Case "1.2L"
                Worksheets("zadani").Shapes("obr-1.2L").Visible = True
            Case "1.2P"
                Worksheets("zadani").Shapes("obr-1.2P").Visible = True
            Case "1.3L"
                Worksheets("zadani").Shapes("obr-1.3L").Visible = True
            Case "1.3P"
                Worksheets("zadani").Shapes("obr-1.3P").Visible = True
            Case "1.4L"
                Worksheets("zadani").Shapes("obr-1.4L").Visible = True
            Case "1.4P"
                Worksheets("zadani").Shapes("obr-1.4P").Visible = True
            Case "1.5L"
                Worksheets("zadani").Shapes("obr-1.5L").Visible = True
            Case "1.5P"
                Worksheets("zadani").Shapes("obr-1.5P").Visible = True
            Case "1.6L"
                Worksheets("zadani").Shapes("obr-1.6L").Visible = True
            Case "1.6P"
                Worksheets("zadani").Shapes("obr-1.6P").Visible = True
        End Select
    ElseIf Worksheets("zadani").Range("C8").Value = "dvoukřídlové" Then
        Select Case Worksheets("zadani").Range("C10").Value
            Case "2A.1L"
                Worksheets("zadani").Shapes("obr-2A.1L").Visible = True
            Case "2A.1P"
                Worksheets("zadani").Shapes("obr-2A.1P").Visible = True
            Case "2A.2L"
                Worksheets("zadani").Shapes("obr-2A.2L").Visible = True
            Case "2A.2P"
                Worksheets("zadani").Shapes("obr-2A.2P").Visible = True
            Case "2A.3L"
                Worksheets("zadani").Shapes("obr-2A.3L").Visible = True
            Case "2A.3P"
                Worksheets("zadani").Shapes("obr-2A.3P").Visible = True
            Case "2A.4L"
                Worksheets("zadani").Shapes("obr-2A.4L").Visible = True
            Case "2A.4P"
                Worksheets("zadani").Shapes("obr-2A.4P").Visible = True
            Case "2A.5L"
                Worksheets("zadani").Shapes("obr-2A.5L").Visible = True
            Case "2A.5P"
                Worksheets("zadani").Shapes("obr-2A.5P").Visible = True
            Case "2A.6L"
                Worksheets("zadani").Shapes("obr-2A.6L").Visible = True
            Case "2A.6P"
                Worksheets("zadani").Shapes("obr-2A.6P").Visible = True
        End Select
    End If
    
    If Worksheets("zadani").Range("C5").Value = "AKTIV 77 - příčkové" Then
        Select Case Worksheets("zadani").Range("C11").Value
            Case "B1"
                Worksheets("zadani").Shapes("obr-B1").Visible = True
            Case "B2"
                Worksheets("zadani").Shapes("obr-B2").Visible = True
            Case "B3"
                Worksheets("zadani").Shapes("obr-B3").Visible = True
            Case "C1"
                Worksheets("zadani").Shapes("obr-C1").Visible = True
            Case "C2"
                Worksheets("zadani").Shapes("obr-C2").Visible = True
            Case "C3"
                Worksheets("zadani").Shapes("obr-C3").Visible = True
            Case "C4"
                Worksheets("zadani").Shapes("obr-C4").Visible = True
            Case "D1"
                Worksheets("zadani").Shapes("obr-D1").Visible = True
            Case "D2"
                Worksheets("zadani").Shapes("obr-D2").Visible = True
            Case "D3"
                Worksheets("zadani").Shapes("obr-D3").Visible = True
            Case "H1"
                Worksheets("zadani").Shapes("obr-H1").Visible = True
            Case "H2"
                Worksheets("zadani").Shapes("obr-H2").Visible = True
            Case "H3"
                Worksheets("zadani").Shapes("obr-H3").Visible = True
            Case "H4"
                Worksheets("zadani").Shapes("obr-H4").Visible = True
            Case "N1"
                Worksheets("zadani").Shapes("obr-N1").Visible = True
            Case "N2"
                Worksheets("zadani").Shapes("obr-N2").Visible = True
            Case "N3"
                Worksheets("zadani").Shapes("obr-N3").Visible = True
            Case "N4"
                Worksheets("zadani").Shapes("obr-N4").Visible = True
            Case "P1"
                Worksheets("zadani").Shapes("obr-P1").Visible = True
            Case "P2"
                Worksheets("zadani").Shapes("obr-P2").Visible = True
            Case "P3"
                Worksheets("zadani").Shapes("obr-P3").Visible = True
            Case "P4"
                Worksheets("zadani").Shapes("obr-P4").Visible = True
        End Select
    End If
End Sub

'*************************************************************************************************************************************************
'V Excelu vloženo do modelu "VypocetAktiv"

Sub vypocetDveriAktiv_1()
    'načtení hodnot pro výpočet
    sirka = Worksheets("zadani").Range("sirkaDveri").Value / 1000
    vyska = Worksheets("zadani").Range("vyskaDveri").Value / 1000
    vypln = Worksheets("zadani").Range("vyplnDveri").Value
    vyskaProfiluHorni = Worksheets("vypocet").Range("I4").Value
    If Worksheets("zadani").Range("okop").Value = "S okopem" Then
        vyskaProfiluSpodni = Worksheets("vypocet").Range("J5").Value
        prostupProfiluSpodni = Worksheets("vypocet").Range("M5").Value
    ElseIf Worksheets("zadani").Range("okop").Value = "Průběžný" Then
        vyskaProfiluSpodni = Worksheets("vypocet").Range("J4").Value
        prostupProfiluSpodni = Worksheets("vypocet").Range("M4").Value
    End If
    prostupProfiluHorni = Worksheets("vypocet").Range("L4").Value
    
    Select Case Worksheets("zadani").Range("vyplnDveri").Value
        Case "Prosklená"
            plochaProskleniCreative = 0
            obvodProskleniCreative = 0
            linearniCinitelProstupuProskleniCreative = 0
            prostupVyplne = Worksheets("vypocet").Range("G3").Value
            linearniCinitelProstupu = 0.043
        Case "Plná HPL"
            plochaProskleniCreative = 0
            obvodProskleniCreative = 0
            linearniCinitelProstupuProskleniCreative = 0
            prostupVyplne = Worksheets("vypocet").Range("Q5").Value
            linearniCinitelProstupu = 0
        Case "Plná ALU"
            plochaProskleniCreative = 0
            obvodProskleniCreative = 0
            linearniCinitelProstupuProskleniCreative = 0
            prostupVyplne = Worksheets("vypocet").Range("Q6").Value
            linearniCinitelProstupu = 0
            
            'výplně s okýnky Creative dle soupisu M.Paláta (mailem 1.12.2021)
        Case "HPL - Creative 904"
            plochaProskleniCreative = 0.22 * 1.52
            obvodProskleniCreative = 2 * (0.22 + 1.52)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q5").Value
            linearniCinitelProstupu = 0
        Case "HPL - Creative 906"
            plochaProskleniCreative = 0.076 * 1.401
            obvodProskleniCreative = 2 * (0.076 + 1.401)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q5").Value
            linearniCinitelProstupu = 0
        Case "HPL - Creative 907"
            plochaProskleniCreative = 5 * 0.4 * 0.08
            obvodProskleniCreative = 10 * (0.4 + 0.08)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q5").Value
            linearniCinitelProstupu = 0
        Case "HPL - Creative 908"
            plochaProskleniCreative = 3 * 0.35 * 0.15
            obvodProskleniCreative = 6 * (0.35 + 0.15)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q5").Value
            linearniCinitelProstupu = 0
        Case "HPL - Creative 909"
            plochaProskleniCreative = 0.22 * 1.52
            obvodProskleniCreative = 2 * (0.22 + 1.52)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q5").Value
            linearniCinitelProstupu = 0
        Case "ALU - Creative 904"
            plochaProskleniCreative = 0.22 * 1.52
            obvodProskleniCreative = 2 * (0.22 + 1.52)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q6").Value
            linearniCinitelProstupu = 0
        Case "ALU - Creative 906"
            plochaProskleniCreative = 0.076 * 1.401
            obvodProskleniCreative = 2 * (0.076 + 1.401)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q6").Value
            linearniCinitelProstupu = 0
        Case "ALU - Creative 907"
            plochaProskleniCreative = 5 * 0.4 * 0.08
            obvodProskleniCreative = 10 * (0.4 + 0.08)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q6").Value
            linearniCinitelProstupu = 0
        Case "ALU - Creative 908"
            plochaProskleniCreative = 3 * 0.35 * 0.15
            obvodProskleniCreative = 6 * (0.35 + 0.15)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q6").Value
            linearniCinitelProstupu = 0
        Case "ALU - Creative 909"
            plochaProskleniCreative = 0.22 * 1.52
            obvodProskleniCreative = 2 * (0.22 + 1.52)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q6").Value
            linearniCinitelProstupu = 0

            'výplně s okýnky Creative řady 3xx dle soupisu M.Paláta (mailem 15.12.2021) - také úkol v Heo
        Case "HPL - Creative 301"
            plochaProskleniCreative = 2 * 0.12 * 0.305  
            obvodProskleniCreative = 2 * 2 * (0.305 + 0.12)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q5").Value
            linearniCinitelProstupu = 0
        Case "HPL - Creative 302"
            plochaProskleniCreative = 3 * 0.12 * 0.305
            obvodProskleniCreative = 3 * 2 * (0.305 + 0.12)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q5").Value
            linearniCinitelProstupu = 0
        Case "HPL - Creative 303", "HPL - Creative 306" 
            plochaProskleniCreative = 4 * 0.17 * 0.17
            obvodProskleniCreative = 4 * 2 * (0.17 + 0.17)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q5").Value
            linearniCinitelProstupu = 0
        Case "HPL - Creative 304", "HPL - Creative 305", "HPL - Creative 307"
            plochaProskleniCreative = 4 * 0.22 * 0.12
            obvodProskleniCreative = 4 * 2 * (0.22 + 0.12)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q5").Value
            linearniCinitelProstupu = 0
        Case "HPL - Creative 308"
            plochaProskleniCreative = 0.183088
            obvodProskleniCreative = 3.526768
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q5").Value
            linearniCinitelProstupu = 0
        Case "HPL - Creative 309"
            plochaProskleniCreative = 4 * 0.22 * 0.22
            obvodProskleniCreative = 4 * 2 * (0.22 + 0.22)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q5").Value
            linearniCinitelProstupu = 0
        
        Case "ALU - Creative 301"
            plochaProskleniCreative = 2 * 0.12 * 0.305  
            obvodProskleniCreative = 2 * 2 * (0.305 + 0.12)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q6").Value
            linearniCinitelProstupu = 0
        Case "ALU - Creative 302"
            plochaProskleniCreative = 3 * 0.12 * 0.305
            obvodProskleniCreative = 3 * 2 * (0.305 + 0.12)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q6").Value
            linearniCinitelProstupu = 0
        Case "ALU - Creative 303", "HPL - Creative 306" 
            plochaProskleniCreative = 4 * 0.17 * 0.17
            obvodProskleniCreative = 4 * 2 * (0.17 + 0.17)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q6").Value
            linearniCinitelProstupu = 0
        Case "ALU - Creative 304", "HPL - Creative 305", "HPL - Creative 307"
            plochaProskleniCreative = 4 * 0.22 * 0.12
            obvodProskleniCreative = 4 * 2 * (0.22 + 0.12)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q6").Value
            linearniCinitelProstupu = 0
        Case "ALU - Creative 308"
            plochaProskleniCreative = 0.183088
            obvodProskleniCreative = 3.526768
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q6").Value
            linearniCinitelProstupu = 0
        Case "ALU - Creative 309"
            plochaProskleniCreative = 4 * 0.22 * 0.22
            obvodProskleniCreative = 4 * 2 * (0.22 + 0.22)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q6").Value
            linearniCinitelProstupu = 0

        Case "*****************"
            i = MsgBox ("Vybraná výplň není relevantní.", 0,"Výplň křídla")
            Worksheets("zadani").Range("vyplnDveri").Value = "Prosklená"

    End Select
    
    'výpočty
    plochaProfiluHorni = vyskaProfiluHorni * ((2 * vyska) + (sirka - (2 * vyskaProfiluHorni)))
    plochaProfiluSpodni = vyskaProfiluSpodni * (sirka - (2 * vyskaProfiluHorni))
    plochaVyplne = (sirka * vyska) - plochaProfiluHorni - plochaProfiluSpodni
    obvodZaskleni = 2 * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni + sirka - 2 * vyskaProfiluHorni)
    prostupDveri = ((plochaProfiluHorni * prostupProfiluHorni) + (plochaProfiluSpodni * prostupProfiluSpodni) + ((plochaVyplne - plochaProskleniCreative) * prostupVyplne) + (plochaProskleniCreative * 1.1) + (obvodZaskleni * linearniCinitelProstupu) + (obvodProskleniCreative * linearniCinitelProstupuProskleniCreative)) / (plochaProfiluHorni + plochaProfiluSpodni + plochaVyplne)
    
    Worksheets("zadani").Range("prostupDveri").Value = prostupDveri
    
    'kontrolní výpisy hodnot
    Worksheets("zadani").Range("C101").Value = sirka
    Worksheets("zadani").Range("C102").Value = vyska
    Worksheets("zadani").Range("C103").Value = vypln
    Worksheets("zadani").Range("C104").Value = vyskaProfiluHorni
    Worksheets("zadani").Range("C105").Value = vyskaProfiluSpodni
    Worksheets("zadani").Range("C106").Value = vyskaProfiluPricky
    Worksheets("zadani").Range("C107").Value = vyskaProfiluSrazu
    Worksheets("zadani").Range("C108").Value = delkaProfiluPricky
    Worksheets("zadani").Range("C109").Value = plochaProfiluHorni
    Worksheets("zadani").Range("C110").Value = plochaProfiluSpodni
    Worksheets("zadani").Range("C111").Value = plochaProfiluPricky
    Worksheets("zadani").Range("C112").Value = plochaVyplne
    Worksheets("zadani").Range("C113").Value = plochaVyplneSklo
    Worksheets("zadani").Range("C114").Value = plochaVyplnePlna
    Worksheets("zadani").Range("C115").Value = prostupProfiluHorni
    Worksheets("zadani").Range("C116").Value = prostupProfiluSpodni
    Worksheets("zadani").Range("C117").Value = prostupDveri
    Worksheets("zadani").Range("C118").Value = prostupProfiluSrazu
    Worksheets("zadani").Range("C119").Value = prostupProfiluPricky
    Worksheets("zadani").Range("C120").Value = prostupVyplneSklo
    Worksheets("zadani").Range("C121").Value = prostupVyplnePlna
    Worksheets("zadani").Range("C122").Value = obvodZaskleni
    Worksheets("zadani").Range("C123").Value = linearniCinitelProstupu
    Worksheets("zadani").Range("C124").Value = prostupVyplne
    Worksheets("zadani").Range("C126").Value = plochaProskleniCreative
    Worksheets("zadani").Range("C127").Value = obvodProskleniCreative
    Worksheets("zadani").Range("C128").Value = linearniCinitelProstupuProskleniCreative
    
End Sub

Sub vypocetDveriAktiv_2()
    'načtení hodnot pro výpočet
    sirka = Worksheets("zadani").Range("sirkaDveri").Value / 1000
    vyska = Worksheets("zadani").Range("vyskaDveri").Value / 1000
    vypln = Worksheets("zadani").Range("vyplnDveri").Value
    vyskaProfiluHorni = Worksheets("vypocet").Range("I4").Value
    vyskaProfiluSrazu = Worksheets("vypocet").Range("K4").Value
    If Worksheets("zadani").Range("okop").Value = "S okopem" Then
        vyskaProfiluSpodni = Worksheets("vypocet").Range("J5").Value
        prostupProfiluSpodni = Worksheets("vypocet").Range("M5").Value
    ElseIf Worksheets("zadani").Range("okop").Value = "Průběžný" Then
        vyskaProfiluSpodni = Worksheets("vypocet").Range("J4").Value
        prostupProfiluSpodni = Worksheets("vypocet").Range("M4").Value
    End If
    prostupProfiluHorni = Worksheets("vypocet").Range("L4").Value
    prostupProfiluSrazu = Worksheets("vypocet").Range("N4").Value
   Select Case Worksheets("zadani").Range("vyplnDveri").Value
        Case "Prosklená"
            plochaProskleniCreative = 0
            obvodProskleniCreative = 0
            linearniCinitelProstupuProskleniCreative = 0
            prostupVyplne = Worksheets("vypocet").Range("G3").Value
            linearniCinitelProstupu = 0.043
        Case "Plná HPL"
            plochaProskleniCreative = 0
            obvodProskleniCreative = 0
            linearniCinitelProstupuProskleniCreative = 0
            prostupVyplne = Worksheets("vypocet").Range("Q5").Value
            linearniCinitelProstupu = 0
        Case "Plná ALU"
            plochaProskleniCreative = 0
            obvodProskleniCreative = 0
            linearniCinitelProstupuProskleniCreative = 0
            prostupVyplne = Worksheets("vypocet").Range("Q6").Value
            linearniCinitelProstupu = 0
            
            'výplně s okýnky Creative dle soupisu M.Paláta (mailem 1.12.2021)
        Case "HPL - Creative 904"
            plochaProskleniCreative = 0.22 * 1.52
            obvodProskleniCreative = 2 * (0.22 + 1.52)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q5").Value
            linearniCinitelProstupu = 0
        Case "HPL - Creative 906"
            plochaProskleniCreative = 0.076 * 1.401
            obvodProskleniCreative = 2 * (0.076 + 1.401)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q5").Value
            linearniCinitelProstupu = 0
        Case "HPL - Creative 907"
            plochaProskleniCreative = 5 * 0.4 * 0.08
            obvodProskleniCreative = 10 * (0.4 + 0.08)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q5").Value
            linearniCinitelProstupu = 0
        Case "HPL - Creative 908"
            plochaProskleniCreative = 3 * 0.35 * 0.15
            obvodProskleniCreative = 6 * (0.35 + 0.15)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q5").Value
            linearniCinitelProstupu = 0
        Case "HPL - Creative 909", "ALU - Creative 909"
            plochaProskleniCreative = 0.22 * 1.52
            obvodProskleniCreative = 2 * (0.22 + 1.52)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q5").Value
            linearniCinitelProstupu = 0
        Case "ALU - Creative 904"
            plochaProskleniCreative = 0.22 * 1.52
            obvodProskleniCreative = 2 * (0.22 + 1.52)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q6").Value
            linearniCinitelProstupu = 0
        Case "ALU - Creative 906"
            plochaProskleniCreative = 0.076 * 1.401
            obvodProskleniCreative = 2 * (0.076 + 1.401)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q6").Value
            linearniCinitelProstupu = 0
        Case "ALU - Creative 907"
            plochaProskleniCreative = 5 * 0.4 * 0.08
            obvodProskleniCreative = 10 * (0.4 + 0.08)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q6").Value
            linearniCinitelProstupu = 0
        Case "ALU - Creative 908"
            plochaProskleniCreative = 3 * 0.35 * 0.15
            obvodProskleniCreative = 6 * (0.35 + 0.15)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q6").Value
            linearniCinitelProstupu = 0
        Case "ALU - Creative 909", "ALU - Creative 909"
            plochaProskleniCreative = 0.22 * 1.52
            obvodProskleniCreative = 2 * (0.22 + 1.52)
            linearniCinitelProstupuProskleniCreative = 0.043
            prostupVyplne = Worksheets("vypocet").Range("Q6").Value
            linearniCinitelProstupu = 0
    End Select
    
    'výpočty
    plochaProfiluHorni = vyskaProfiluHorni * ((2 * vyska) + (sirka - (2 * vyskaProfiluHorni) - vyskaProfiluSrazu))
    plochaProfiluSpodni = vyskaProfiluSpodni * (sirka - (2 * vyskaProfiluHorni) - vyskaProfiluSrazu)
    plochaProfiluSrazu = vyskaProfiluSrazu * vyska
    plochaVyplne = (sirka * vyska) - plochaProfiluHorni - plochaProfiluSpodni - plochaProfiluSrazu
    obvodZaskleni = 2 * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni + sirka - 2 * vyskaProfiluHorni - vyskaProfiluSrazu)
    prostupDveri = ((plochaProfiluHorni * prostupProfiluHorni) + (plochaProfiluSpodni * prostupProfiluSpodni) + (plochaProfiluSrazu * prostupProfiluSrazu) + ((plochaVyplne - plochaProskleniCreative) * prostupVyplne) + (plochaProskleniCreative * 1.1) + (obvodZaskleni * linearniCinitelProstupu) + (obvodProskleniCreative * linearniCinitelProstupuProskleniCreative)) / (plochaProfiluHorni + plochaProfiluSpodni + plochaProfiluSrazu + plochaVyplne)
    
    Worksheets("zadani").Range("prostupDveri").Value = prostupDveri

    'kontrolní výpisy hodnot
    Worksheets("zadani").Range("C101").Value = sirka
    Worksheets("zadani").Range("C102").Value = vyska
    Worksheets("zadani").Range("C103").Value = vypln
    Worksheets("zadani").Range("C104").Value = vyskaProfiluHorni
    Worksheets("zadani").Range("C105").Value = vyskaProfiluSpodni
    Worksheets("zadani").Range("C106").Value = vyskaProfiluPricky
    Worksheets("zadani").Range("C107").Value = vyskaProfiluSrazu
    Worksheets("zadani").Range("C108").Value = delkaProfiluPricky
    Worksheets("zadani").Range("C109").Value = plochaProfiluHorni
    Worksheets("zadani").Range("C110").Value = plochaProfiluSpodni
    Worksheets("zadani").Range("C111").Value = plochaProfiluPricky
    Worksheets("zadani").Range("C112").Value = plochaVyplne
    Worksheets("zadani").Range("C113").Value = plochaVyplneSklo
    Worksheets("zadani").Range("C114").Value = plochaVyplnePlna
    Worksheets("zadani").Range("C115").Value = prostupProfiluHorni
    Worksheets("zadani").Range("C116").Value = prostupProfiluSpodni
    Worksheets("zadani").Range("C117").Value = prostupDveri
    Worksheets("zadani").Range("C118").Value = prostupProfiluSrazu
    Worksheets("zadani").Range("C119").Value = prostupProfiluPricky
    Worksheets("zadani").Range("C120").Value = prostupVyplneSklo
    Worksheets("zadani").Range("C121").Value = prostupVyplnePlna
    Worksheets("zadani").Range("C122").Value = obvodZaskleni
    Worksheets("zadani").Range("C123").Value = linearniCinitelProstupu
    Worksheets("zadani").Range("C124").Value = prostupVyplne
    Worksheets("zadani").Range("C125").Value = plochaProfiluSrazu
    Worksheets("zadani").Range("C126").Value = plochaProskleniCreative
    Worksheets("zadani").Range("C127").Value = obvodProskleniCreative
    Worksheets("zadani").Range("C128").Value = linearniCinitelProstupuProskleniCreative

End Sub

'*************************************************************************************************************************************************
'V Excelu vloženo do modelu "VypocetAktivExcellent"

Sub vypocetDveriExcellent()
    'načtení hodnot pro výpočet
    sirka = Worksheets("zadani").Range("sirkaDveri").Value / 1000
    vyska = Worksheets("zadani").Range("vyskaDveri").Value / 1000
    vypln = Worksheets("zadani").Range("vyplnDveri").Value
    
    If Worksheets("zadani").Range("licovani").Value = "jednostranné" Then
        vyskaProfiluHorni = Worksheets("vypocet").Range("I6").Value
        vyskaProfiluSpodni = Worksheets("vypocet").Range("J6").Value
        prostupProfiluHorni = Worksheets("vypocet").Range("L6").Value
        prostupProfiluSpodni = Worksheets("vypocet").Range("M6").Value
        prostupVyplne = Worksheets("vypocet").Range("Q7").Value
    ElseIf Worksheets("zadani").Range("licovani").Value = "oboustranné" Then
        vyskaProfiluHorni = Worksheets("vypocet").Range("I7").Value
        vyskaProfiluSpodni = Worksheets("vypocet").Range("J7").Value
        prostupProfiluHorni = Worksheets("vypocet").Range("L7").Value
        prostupProfiluSpodni = Worksheets("vypocet").Range("M7").Value
        prostupVyplne = Worksheets("vypocet").Range("Q8").Value
    End If
    
    linearniCinitelProstupu = 0
    
    'výpočty
    plochaProfiluHorni = vyskaProfiluHorni * ((2 * vyska) + (sirka - (2 * vyskaProfiluHorni)))
    plochaProfiluSpodni = vyskaProfiluSpodni * (sirka - (2 * vyskaProfiluHorni))
    plochaVyplne = (sirka * vyska) - plochaProfiluHorni - plochaProfiluSpodni
    prostupDveri = ((plochaProfiluHorni * prostupProfiluHorni) + (plochaProfiluSpodni * prostupProfiluSpodni) + (plochaVyplne * prostupVyplne)) / (plochaProfiluHorni + plochaProfiluSpodni + plochaVyplne)
    
    Worksheets("zadani").Range("prostupDveri").Value = prostupDveri
    
    'kontrolní výpisy hodnot
    Worksheets("zadani").Range("C101").Value = sirka
    Worksheets("zadani").Range("C102").Value = vyska
    Worksheets("zadani").Range("C104").Value = vyskaProfiluHorni
    Worksheets("zadani").Range("C105").Value = vyskaProfiluSpodni
    Worksheets("zadani").Range("C106").Value = vyskaProfiluPricky
    Worksheets("zadani").Range("C107").Value = vyskaProfiluSrazu
    Worksheets("zadani").Range("C108").Value = delkaProfiluPricky
    Worksheets("zadani").Range("C109").Value = plochaProfiluHorni
    Worksheets("zadani").Range("C110").Value = plochaProfiluSpodni
    Worksheets("zadani").Range("C111").Value = plochaProfiluPricky
    Worksheets("zadani").Range("C112").Value = plochaVyplne
    Worksheets("zadani").Range("C113").Value = plochaVyplneSklo
    Worksheets("zadani").Range("C114").Value = plochaVyplnePlna
    Worksheets("zadani").Range("C115").Value = prostupProfiluHorni
    Worksheets("zadani").Range("C116").Value = prostupProfiluSpodni
    Worksheets("zadani").Range("C117").Value = prostupDveri
    Worksheets("zadani").Range("C118").Value = prostupProfiluSrazu
    Worksheets("zadani").Range("C119").Value = prostupProfiluPricky
    Worksheets("zadani").Range("C124").Value = prostupVyplne
    
End Sub

'*************************************************************************************************************************************************
'V Excelu vloženo do modelu "VypocetAktivPrickove"

Sub vypocetDveriPrickove()
                    
    sirka = Worksheets("zadani").Range("sirkaDveri").Value / 1000
    vyska = Worksheets("zadani").Range("vyskaDveri").Value / 1000
    vypln = Worksheets("zadani").Range("vyplnDveri").Value
    vyskaProfiluHorni = Worksheets("vypocet").Range("I4").Value
    If Worksheets("zadani").Range("okop").Value = "S okopem" Then
        vyskaProfiluSpodni = Worksheets("vypocet").Range("J5").Value
        prostupProfiluSpodni = Worksheets("vypocet").Range("M5").Value
    ElseIf Worksheets("zadani").Range("okop").Value = "Průběžný" Then
        vyskaProfiluSpodni = Worksheets("vypocet").Range("J4").Value
        prostupProfiluSpodni = Worksheets("vypocet").Range("M4").Value
    End If
    vyskaProfiluPricky = Worksheets("vypocet").Range("I8").Value
    prostupProfiluHorni = Worksheets("vypocet").Range("L4").Value
    prostupProfiluPricky = Worksheets("vypocet").Range("L8").Value
    prostupVyplneSklo = Worksheets("vypocet").Range("G3").Value
    prostupVyplnePlna = Worksheets("vypocet").Range("Q5").Value
    linearniCinitelProstupu = 0.043
    
    'výpočty
    Select Case Worksheets("zadani").Range("model").Value
        Case "B1"
            delkaProfiluPricky = sirka - (2 * vyskaProfiluHorni)
            plochaProfiluPricky = delkaProfiluPricky * vyskaProfiluPricky
            plochaVyplne = (sirka - 2 * vyskaProfiluHorni) * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni) - plochaProfiluPricky
            plochaVyplneSklo = plochaVyplne
            plochaVyplnePlna = 0
            obvodZaskleni = 2 * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni - vyskaProfiluPricky) + 4 * (sirka - 2 * vyskaProfiluHorni)
        Case "B2"
            delkaProfiluPricky = sirka - (2 * vyskaProfiluHorni)
            plochaProfiluPricky = delkaProfiluPricky * vyskaProfiluPricky
            plochaVyplne = (sirka - 2 * vyskaProfiluHorni) * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni) - plochaProfiluPricky
            plochaVyplneSklo = 2 / 3 * plochaVyplne
            plochaVyplnePlna = 1 / 3 * plochaVyplne
            obvodZaskleni = 2 * (2 / 3 * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni - vyskaProfiluPricky)) + 2 * (sirka - 2 * vyskaProfiluHorni)
        Case "B3"
            delkaProfiluPricky = sirka - (2 * vyskaProfiluHorni)
            plochaProfiluPricky = delkaProfiluPricky * vyskaProfiluPricky
            plochaVyplne = (sirka - 2 * vyskaProfiluHorni) * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni) - plochaProfiluPricky
            plochaVyplneSklo = 0
            plochaVyplnePlna = plochaVyplne
            obvodZaskleni = 0
        Case "C1"
            delkaProfiluPricky = 2 * (sirka - (2 * vyskaProfiluHorni))
            plochaProfiluPricky = delkaProfiluPricky * vyskaProfiluPricky
            plochaVyplne = (sirka - 2 * vyskaProfiluHorni) * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni) - plochaProfiluPricky
            plochaVyplneSklo = plochaVyplne
            plochaVyplnePlna = 0
            obvodZaskleni = 2 * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni - 2 * vyskaProfiluPricky) + 6 * (sirka - 2 * vyskaProfiluHorni)
        Case "C2"
            delkaProfiluPricky = 2 * (sirka - (2 * vyskaProfiluHorni))
            plochaProfiluPricky = delkaProfiluPricky * vyskaProfiluPricky
            plochaVyplne = (sirka - 2 * vyskaProfiluHorni) * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni) - plochaProfiluPricky
            plochaVyplneSklo = 2 / 3 * plochaVyplne
            plochaVyplnePlna = 1 / 3 * plochaVyplne
            obvodZaskleni = 2 * (2 / 3 * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni - 2 * vyskaProfiluPricky)) + 4 * (sirka - 2 * vyskaProfiluHorni)
        Case "C3"
            delkaProfiluPricky = 2 * (sirka - (2 * vyskaProfiluHorni))
            plochaProfiluPricky = delkaProfiluPricky * vyskaProfiluPricky
            plochaVyplne = (sirka - 2 * vyskaProfiluHorni) * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni) - plochaProfiluPricky
            plochaVyplneSklo = 1 / 3 * plochaVyplne
            plochaVyplnePlna = 2 / 3 * plochaVyplne
            obvodZaskleni = 2 * (1 / 3 * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni - 2 * vyskaProfiluPricky)) + 2 * (sirka - 2 * vyskaProfiluHorni)
        Case "C4"
            delkaProfiluPricky = 2 * (sirka - (2 * vyskaProfiluHorni))
            plochaProfiluPricky = delkaProfiluPricky * vyskaProfiluPricky
            plochaVyplne = (sirka - 2 * vyskaProfiluHorni) * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni) - plochaProfiluPricky
            plochaVyplneSklo = 0
            plochaVyplnePlna = plochaVyplne
            obvodZaskleni = 0
        Case "D1"
            delkaProfiluPricky = sirka - (2 * vyskaProfiluHorni) + (2 / 3 * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni - vyskaProfiluPricky))
            plochaProfiluPricky = delkaProfiluPricky * vyskaProfiluPricky
            plochaVyplne = (sirka - 2 * vyskaProfiluHorni) * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni) - plochaProfiluPricky
            plochaVyplneSklo = plochaVyplne
            plochaVyplnePlna = 0
            obvodZaskleni = 4 * (2 / 3 * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni - vyskaProfiluPricky)) + 2 * (sirka - 2 * vyskaProfiluHorni) + 4 / 2 * (sirka - vyskaProfiluPricky - 2 * vyskaProfiluHorni)
        Case "D2"
            delkaProfiluPricky = sirka - (2 * vyskaProfiluHorni) + (2 / 3 * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni - vyskaProfiluPricky))
            plochaProfiluPricky = delkaProfiluPricky * vyskaProfiluPricky
            plochaVyplne = (sirka - 2 * vyskaProfiluHorni) * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni) - plochaProfiluPricky
            plochaVyplneSklo = 2 / 3 * plochaVyplne
            plochaVyplnePlna = 1 / 3 * plochaVyplne
            obvodZaskleni = 4 * (2 / 3 * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni - vyskaProfiluPricky)) + 4 / 2 * (sirka - vyskaProfiluPricky - 2 * vyskaProfiluHorni)
        Case "D3"
            delkaProfiluPricky = sirka - (2 * vyskaProfiluHorni) + (2 / 3 * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni - vyskaProfiluPricky))
            plochaProfiluPricky = delkaProfiluPricky * vyskaProfiluPricky
            plochaVyplne = (sirka - 2 * vyskaProfiluHorni) * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni) - plochaProfiluPricky
            plochaVyplneSklo = 0
            plochaVyplnePlna = plochaVyplne
            obvodZaskleni = 0
        Case "H1"
            delkaProfiluPricky = 2 * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni)
            plochaProfiluPricky = delkaProfiluPricky * vyskaProfiluPricky
            plochaVyplne = (sirka - 2 * vyskaProfiluHorni) * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni) - plochaProfiluPricky
            plochaVyplneSklo = plochaVyplne
            plochaVyplnePlna = 0
            obvodZaskleni = 6 * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni) + 6 / 3 * (sirka - 2 * vyskaProfiluPricky - 2 * vyskaProfiluHorni)
        Case "H2"
            delkaProfiluPricky = 2 * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni)
            plochaProfiluPricky = delkaProfiluPricky * vyskaProfiluPricky
            plochaVyplne = (sirka - 2 * vyskaProfiluHorni) * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni) - plochaProfiluPricky
            plochaVyplneSklo = 2 / 3 * plochaVyplne
            plochaVyplnePlna = 1 / 3 * plochaVyplne
            obvodZaskleni = 4 * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni) + 4 / 3 * (sirka - 2 * vyskaProfiluPricky - 2 * vyskaProfiluHorni)
        Case "H3"
            delkaProfiluPricky = 2 * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni)
            plochaProfiluPricky = delkaProfiluPricky * vyskaProfiluPricky
            plochaVyplne = (sirka - 2 * vyskaProfiluHorni) * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni) - plochaProfiluPricky
            plochaVyplneSklo = 1 / 3 * plochaVyplne
            plochaVyplnePlna = 2 / 3 * plochaVyplne
            obvodZaskleni = 2 * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni) + 2 / 3 * (sirka - 2 * vyskaProfiluPricky - 2 * vyskaProfiluHorni)
        Case "H4"
            delkaProfiluPricky = 2 * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni)
            plochaProfiluPricky = delkaProfiluPricky * vyskaProfiluPricky
            plochaVyplne = (sirka - 2 * vyskaProfiluHorni) * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni) - plochaProfiluPricky
            plochaVyplneSklo = 0
            plochaVyplnePlna = plochaVyplne
            obvodZaskleni = 0
        Case "N1"
            delkaProfiluPricky = vyska - vyskaProfiluHorni - vyskaProfiluSpodni
            plochaProfiluPricky = delkaProfiluPricky * vyskaProfiluPricky
            plochaVyplne = (sirka - 2 * vyskaProfiluHorni) * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni) - plochaProfiluPricky
            plochaVyplneSklo = plochaVyplne
            plochaVyplnePlna = 0
            obvodZaskleni = 4 * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni) + 6 / 3 * (sirka - vyskaProfiluPricky - 2 * vyskaProfiluHorni)
        Case "N2"
            delkaProfiluPricky = vyska - vyskaProfiluHorni - vyskaProfiluSpodni
            plochaProfiluPricky = delkaProfiluPricky * vyskaProfiluPricky
            plochaVyplne = (sirka - 2 * vyskaProfiluHorni) * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni) - plochaProfiluPricky
            plochaVyplneSklo = 1 / 3 * plochaVyplne
            plochaVyplnePlna = 2 / 3 * plochaVyplne
            obvodZaskleni = 2 * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni) + 2 / 3 * (sirka - vyskaProfiluPricky - 2 * vyskaProfiluHorni)
        Case "N3"
            delkaProfiluPricky = vyska - vyskaProfiluHorni - vyskaProfiluSpodni
            plochaProfiluPricky = delkaProfiluPricky * vyskaProfiluPricky
            plochaVyplne = (sirka - 2 * vyskaProfiluHorni) * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni) - plochaProfiluPricky
            plochaVyplneSklo = 2 / 3 * plochaVyplne
            plochaVyplnePlna = 1 / 3 * plochaVyplne
            obvodZaskleni = 2 * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni) + 4 / 3 * (sirka - vyskaProfiluPricky - 2 * vyskaProfiluHorni)
        Case "N4"
            delkaProfiluPricky = vyska - vyskaProfiluHorni - vyskaProfiluSpodni
            plochaProfiluPricky = delkaProfiluPricky * vyskaProfiluPricky
            plochaVyplne = (sirka - 2 * vyskaProfiluHorni) * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni) - plochaProfiluPricky
            plochaVyplneSklo = 0
            plochaVyplnePlna = plochaVyplne
            obvodZaskleni = 0
        Case "P1"
            delkaProfiluPricky = 2 * (2 / 3 * (sirka - vyskaProfiluPricky - 2 * vyskaProfiluHorni)) + (vyska - vyskaProfiluHorni - vyskaProfiluSpodni)
            plochaProfiluPricky = delkaProfiluPricky * vyskaProfiluPricky
            plochaVyplne = (sirka - 2 * vyskaProfiluHorni) * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni) - plochaProfiluPricky
            plochaVyplneSklo = plochaVyplne
            plochaVyplnePlna = 0
            obvodZaskleni = 6 * (2 / 3 * (sirka - vyskaProfiluPricky - 2 * vyskaProfiluHorni)) + 6 / 3 * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni - 2 * vyskaProfiluPricky) + 2 * (1 / 3 * (sirka - vyskaProfiluPricky - 2 * vyskaProfiluHorni)) + 2 * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni)
        Case "P2"
            delkaProfiluPricky = 2 * (2 / 3 * (sirka - vyskaProfiluPricky - 2 * vyskaProfiluHorni)) + (vyska - vyskaProfiluHorni - vyskaProfiluSpodni)
            plochaProfiluPricky = delkaProfiluPricky * vyskaProfiluPricky
            plochaVyplne = (sirka - 2 * vyskaProfiluHorni) * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni) - plochaProfiluPricky
            plochaVyplneSklo = 1 / 3 * plochaVyplne
            plochaVyplnePlna = 2 / 3 * plochaVyplne
            obvodZaskleni = 2 * (1 / 3 * (sirka - vyskaProfiluPricky - 2 * vyskaProfiluHorni)) + 2 * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni)
        Case "P3"
            delkaProfiluPricky = 2 * (2 / 3 * (sirka - vyskaProfiluPricky - 2 * vyskaProfiluHorni)) + (vyska - vyskaProfiluHorni - vyskaProfiluSpodni)
            plochaProfiluPricky = delkaProfiluPricky * vyskaProfiluPricky
            plochaVyplne = (sirka - 2 * vyskaProfiluHorni) * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni) - plochaProfiluPricky
            plochaVyplneSklo = 2 / 3 * plochaVyplne
            plochaVyplnePlna = 1 / 3 * plochaVyplne
            obvodZaskleni = 6 * (2 / 3 * (sirka - vyskaProfiluPricky - 2 * vyskaProfiluHorni)) + 6 / 3 * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni - 2 * vyskaProfiluPricky)
        Case "P4"
            delkaProfiluPricky = 2 * (2 / 3 * (sirka - vyskaProfiluPricky - 2 * vyskaProfiluHorni)) + (vyska - vyskaProfiluHorni - vyskaProfiluSpodni)
            plochaProfiluPricky = delkaProfiluPricky * vyskaProfiluPricky
            plochaVyplne = (sirka - 2 * vyskaProfiluHorni) * (vyska - vyskaProfiluHorni - vyskaProfiluSpodni) - plochaProfiluPricky
            plochaVyplneSklo = 0
            plochaVyplnePlna = plochaVyplne
            obvodZaskleni = 0
        Case Else
            i = MsgBox("Model dveří AKTIV 77 - příčkové musí být zvolen." & vbCrLf & "Po zavření této hlášky bude změněn na B1", vbOKOnly + vbExclamation)
            Worksheets("zadani").Range("model").Value = "B1"
    End Select
    
    plochaProfiluHorni = vyskaProfiluHorni * ((2 * vyska) + (sirka - (2 * vyskaProfiluHorni)))
    plochaProfiluSpodni = vyskaProfiluSpodni * (sirka - (2 * vyskaProfiluHorni))
    prostupDveri = ((plochaProfiluHorni * prostupProfiluHorni) + (plochaProfiluSpodni * prostupProfiluSpodni) + (plochaProfiluPricky * prostupProfiluPricky) + (plochaVyplneSklo * prostupVyplneSklo) + (plochaVyplnePlna * prostupVyplnePlna) + (obvodZaskleni * linearniCinitelProstupu)) / (plochaProfiluHorni + plochaProfiluSpodni + plochaProfiluPricky + plochaVyplneSklo + plochaVyplnePlna)
    
    Worksheets("zadani").Range("prostupDveri").Value = prostupDveri
    
    'kontrolní výpisy hodnot
    Worksheets("zadani").Range("C101").Value = sirka
    Worksheets("zadani").Range("C102").Value = vyska
    Worksheets("zadani").Range("C103").Value = vypln
    Worksheets("zadani").Range("C104").Value = vyskaProfiluHorni
    Worksheets("zadani").Range("C105").Value = vyskaProfiluSpodni
    Worksheets("zadani").Range("C106").Value = vyskaProfiluPricky
    Worksheets("zadani").Range("C107").Value = vyskaProfiluSrazu
    Worksheets("zadani").Range("C108").Value = delkaProfiluPricky
    Worksheets("zadani").Range("C109").Value = plochaProfiluHorni
    Worksheets("zadani").Range("C110").Value = plochaProfiluSpodni
    Worksheets("zadani").Range("C111").Value = plochaProfiluPricky
    'Worksheets("zadani").Range("C112").Value = plochaVyplne
    Worksheets("zadani").Range("C113").Value = plochaVyplneSklo
    Worksheets("zadani").Range("C114").Value = plochaVyplnePlna
    Worksheets("zadani").Range("C115").Value = prostupProfiluHorni
    Worksheets("zadani").Range("C116").Value = prostupProfiluSpodni
    Worksheets("zadani").Range("C117").Value = prostupDveri
    Worksheets("zadani").Range("C118").Value = prostupProfiluSrazu
    Worksheets("zadani").Range("C119").Value = prostupProfiluPricky
    Worksheets("zadani").Range("C120").Value = prostupVyplneSklo
    Worksheets("zadani").Range("C121").Value = prostupVyplnePlna
    Worksheets("zadani").Range("C122").Value = obvodZaskleni
    Worksheets("zadani").Range("C123").Value = linearniCinitelProstupu

    
End Sub

'*************************************************************************************************************************************************
'V Excelu vloženo do modelu "VypocetAktivSvetlik"

Sub vypocetSvetlikuB()
                
    vyskaProfiluFix = Worksheets("vypocet").Range("I9").Value
    prostupProfiluFix = Worksheets("vypocet").Range("L9").Value

    'určení rozměrů a ploch
    sirka = Worksheets("zadani").Range("sirkaFixuB").Value / 1000
    vyska = Worksheets("zadani").Range("vyskaFixuB").Value / 1000
    vypln = Worksheets("zadani").Range("vyplnFixuB").Value
    
    'profily a sklo
    plochaProfiluFixu = ((2 * vyska) + (2 * (sirka - (2 * vyskaProfiluFix)))) * vyskaProfiluFix
    plochaVyplneFixu = (sirka * vyska) - plochaProfiluFixu
    obvodZaskleni = 2 * ((sirka + vyska) - (4 * vyskaProfiluFix))
    
    If vypln = "Prosklená" Then
        linearniCinitelProstupu = 0.043
        prostupVyplne = Worksheets("vypocet").Range("G3").Value
    ElseIf vypln = "Plná HPL" Then
        linearniCinitelProstupu = 0
        prostupVyplne = Worksheets("vypocet").Range("Q5").Value
    ElseIf vypln = "Plná ALU" Then
        linearniCinitelProstupu = 0
        prostupVyplne = Worksheets("vypocet").Range("Q6").Value
    End If
    
    prostupFixB = ((plochaProfiluFixu * prostupProfiluFix) + (plochaVyplneFixu * prostupVyplne) + (linearniCinitelProstupu * obvodZaskleni)) / ((plochaProfiluFixu + plochaVyplneFixu))
    
    Worksheets("zadani").Range("prostupFixB").Value = prostupFixB
    
    'kontrolní výpisy hodnot
    Worksheets("zadani").Range("G101").Value = sirka
    Worksheets("zadani").Range("G102").Value = vyska
    Worksheets("zadani").Range("G103").Value = vypln
    Worksheets("zadani").Range("G104").Value = plochaVyplneFixu
    Worksheets("zadani").Range("G105").Value = vyskaProfiluFix
    Worksheets("zadani").Range("G106").Value = plochaProfiluFixu
    Worksheets("zadani").Range("G107").Value = prostupProfiluFix
    Worksheets("zadani").Range("G108").Value = prostupVyplne
    Worksheets("zadani").Range("G109").Value = obvodZaskleni
    Worksheets("zadani").Range("G110").Value = linearniCinitelProstupu
    Worksheets("zadani").Range("G111").Value = prostupFixB
    
End Sub

Sub vypocetSvetlikuC()
    
    'Call nacteniKonstant
                    
    vyskaProfiluFix = Worksheets("vypocet").Range("I9").Value
    prostupProfiluFix = Worksheets("vypocet").Range("L9").Value
    prostupSkla = 0.5
    prostupHpl = Worksheets("vypocet").Range("Q5").Value
    prostupAlu48 = Worksheets("vypocet").Range("Q6").Value

    'určení rozměrů a ploch
    sirka = Worksheets("zadani").Range("sirkaFixuC").Value / 1000
    vyska = Worksheets("zadani").Range("vyskaFixuC").Value / 1000
    vypln = Worksheets("zadani").Range("vyplnFixuC").Value
    
    'profily a sklo
    plochaProfiluFixu = ((2 * vyska) + (2 * (sirka - (2 * vyskaProfiluFix)))) * vyskaProfiluFix
    plochaVyplneFixu = (sirka * vyska) - plochaProfiluFixu
    obvodZaskleni = 2 * ((sirka + vyska) - (4 * vyskaProfiluFix))
    
    If vypln = "Prosklená" Then
        linearniCinitelProstupu = 0.043
        prostupVyplne = Worksheets("vypocet").Range("G3").Value
    ElseIf vypln = "Plná HPL" Then
        linearniCinitelProstupu = 0
        prostupVyplne = Worksheets("vypocet").Range("Q5").Value
    ElseIf vypln = "Plná ALU" Then
        linearniCinitelProstupu = 0
        prostupVyplne = Worksheets("vypocet").Range("Q6").Value
    End If
    
    prostupFixC = ((plochaProfiluFixu * prostupProfiluFix) + (plochaVyplneFixu * prostupVyplne) + (linearniCinitelProstupu * obvodZaskleni)) / ((plochaProfiluFixu + plochaVyplneFixu))
    
    Worksheets("zadani").Range("prostupFixC").Value = prostupFixC
    
    'kontrolní výpisy hodnot
    Worksheets("zadani").Range("J101").Value = sirka
    Worksheets("zadani").Range("J102").Value = vyska
    Worksheets("zadani").Range("J103").Value = vypln
    Worksheets("zadani").Range("J104").Value = plochaVyplneFixu
    Worksheets("zadani").Range("J105").Value = vyskaProfiluFix
    Worksheets("zadani").Range("J106").Value = plochaProfiluFixu
    Worksheets("zadani").Range("J107").Value = prostupProfiluFix
    Worksheets("zadani").Range("J108").Value = prostupVyplne
    Worksheets("zadani").Range("J109").Value = obvodZaskleni
    Worksheets("zadani").Range("J110").Value = linearniCinitelProstupu
    Worksheets("zadani").Range("J111").Value = prostupFixC
    
End Sub
Sub vypocetSvetlikuD()
    
    'Call nacteniKonstant
    vyskaProfiluFix = Worksheets("vypocet").Range("I9").Value
    prostupProfiluFix = Worksheets("vypocet").Range("L9").Value
    prostupSkla = 0.5
    prostupHpl = Worksheets("vypocet").Range("Q5").Value
    prostupAlu48 = Worksheets("vypocet").Range("Q6").Value

    'určení rozměrů a ploch
    sirka = Worksheets("zadani").Range("sirkaFixuD").Value / 1000
    vyska = Worksheets("zadani").Range("vyskaFixuD").Value / 1000
    vypln = Worksheets("zadani").Range("vyplnFixuD").Value
    
    'profily a sklo
    plochaProfiluFixu = ((2 * vyska) + (2 * (sirka - (2 * vyskaProfiluFix)))) * vyskaProfiluFix
    plochaVyplneFixu = (sirka * vyska) - plochaProfiluFixu
    obvodZaskleni = 2 * ((sirka + vyska) - (4 * vyskaProfiluFix))
    
    If vypln = "Prosklená" Then
        linearniCinitelProstupu = 0.043
        prostupVyplne = Worksheets("vypocet").Range("G3").Value
    ElseIf vypln = "Plná HPL" Then
        linearniCinitelProstupu = 0
        prostupVyplne = Worksheets("vypocet").Range("Q5").Value
    ElseIf vypln = "Plná ALU" Then
        linearniCinitelProstupu = 0
        prostupVyplne = Worksheets("vypocet").Range("Q6").Value
    End If
    
    prostupFixD = ((plochaProfiluFixu * prostupProfiluFix) + (plochaVyplneFixu * prostupVyplne) + (linearniCinitelProstupu * obvodZaskleni)) / ((plochaProfiluFixu + plochaVyplneFixu))
    
    Worksheets("zadani").Range("prostupFixD").Value = prostupFixD
    
    'kontrolní výpisy hodnot
    Worksheets("zadani").Range("L101").Value = sirka
    Worksheets("zadani").Range("L102").Value = vyska
    Worksheets("zadani").Range("L103").Value = vypln
    Worksheets("zadani").Range("L104").Value = plochaVyplneFixu
    Worksheets("zadani").Range("L105").Value = vyskaProfiluFix
    Worksheets("zadani").Range("L106").Value = plochaProfiluFixu
    Worksheets("zadani").Range("L107").Value = prostupProfiluFix
    Worksheets("zadani").Range("L108").Value = prostupVyplne
    Worksheets("zadani").Range("L109").Value = obvodZaskleni
    Worksheets("zadani").Range("L110").Value = linearniCinitelProstupu
    Worksheets("zadani").Range("L111").Value = prostupFixD
    
End Sub

'*************************************************************************************************************************************************
'V Excelu vloženo do modelu "VypocetAktivSestavy"

Sub vypocetKompletniSestavy()
    'podle vzorce Uw = (Uf x Af + Ug x Ag + ?g x Ig) / (Af + Ag)
    
    Dim prostupSestavy As Double
    Dim prostupSestavyVrsekZlomku As Double
    Dim prostupSestavySpodekZlomku As Double
    
    Dim prostupVyplneFixuB As Double
    Dim plochaVyplneFixuB As Double
    Dim plochaProfiluFixuB As Double
    Dim prostupProfiluFixuB As Double
        
    Dim prostupVyplneFixuC As Double
    Dim plochaVyplneFixuC As Double
    Dim plochaProfiluFixuC As Double
    Dim prostupProfiluFixuC As Double
        
    Dim prostupVyplneFixuD As Double
    Dim plochaVyplneFixuD As Double
    Dim plochaProfiluFixuD As Double
    Dim prostupProfiluFixuD As Double
        
    'načtení pomocných výpočtů z listu zadani
    'pro dveře
    plochaProfiluHorni = Worksheets("zadani").Range("C109").Value
    plochaProfiluSpodni = Worksheets("zadani").Range("C110").Value
    plochaProfiluPricky = Worksheets("zadani").Range("C111").Value
    plochaVyplne = Worksheets("zadani").Range("C112").Value
    plochaVyplneSklo = Worksheets("zadani").Range("C113").Value
    plochaVyplnePlna = Worksheets("zadani").Range("C114").Value
    prostupProfiluHorni = Worksheets("zadani").Range("C115").Value
    prostupProfiluSpodni = Worksheets("zadani").Range("C116").Value
    prostupProfiluSrazu = Worksheets("zadani").Range("C118").Value
    prostupProfiluPricky = Worksheets("zadani").Range("C119").Value
    prostupVyplneSklo = Worksheets("zadani").Range("C120").Value
    prostupVyplnePlna = Worksheets("zadani").Range("C121").Value
    obvodZaskleni = Worksheets("zadani").Range("C122").Value
    linearniCinitelProstupu = Worksheets("zadani").Range("C123").Value
    prostupVyplne = Worksheets("zadani").Range("C124").Value
    plochaProfiluSrazu = Worksheets("zadani").Range("C125").Value
    plochaProskleniCreative = Worksheets("zadani").Range("C126").Value
    obvodProskleniCreative = Worksheets("zadani").Range("C127").Value
    linearniCinitelProstupuProskleniCreative = Worksheets("zadani").Range("C128").Value

    'pro světlík B
    prostupVyplneFixuB = Worksheets("zadani").Range("G108").Value
    plochaVyplneFixuB = Worksheets("zadani").Range("G104").Value
    plochaProfiluFixuB = Worksheets("zadani").Range("G106").Value
    prostupProfiluFixuB = Worksheets("zadani").Range("G107").Value
    obvodZaskleniFixuB = Worksheets("zadani").Range("G109").Value
    linearniCinitelProstupuFixuB = Worksheets("zadani").Range("G110").Value
    'pro světlík C
    prostupVyplneFixuC = Worksheets("zadani").Range("J108").Value
    plochaVyplneFixuC = Worksheets("zadani").Range("J104").Value
    plochaProfiluFixuC = Worksheets("zadani").Range("J106").Value
    prostupProfiluFixuC = Worksheets("zadani").Range("J107").Value
    obvodZaskleniFixuC = Worksheets("zadani").Range("J109").Value
    linearniCinitelProstupuFixuC = Worksheets("zadani").Range("J110").Value
    'pro světlík D
    prostupVyplneFixuD = Worksheets("zadani").Range("L108").Value
    plochaVyplneFixuD = Worksheets("zadani").Range("L104").Value
    plochaProfiluFixuD = Worksheets("zadani").Range("L106").Value
    prostupProfiluFixuD = Worksheets("zadani").Range("L107").Value
    obvodZaskleniFixuD = Worksheets("zadani").Range("L109").Value
    linearniCinitelProstupuFixuD = Worksheets("zadani").Range("L110").Value

'---------------------------------------------------------------------------------------
    
    'VÝPOČET
    prostupSestavyVrsekZlomku = _
        plochaProfiluHorni * prostupProfiluHorni + _
        plochaProfiluSpodni * prostupProfiluSpodni + _
        plochaProfiluPricky * prostupProfiluPricky + _
        plochaProfiluSrazu * prostupProfiluSrazu + _
        (plochaVyplne - plochaProskleniCreative) * prostupVyplne + _
        plochaProskleniCreative * 1.1 + _
        plochaVyplneSklo * prostupVyplneSklo + _
        plochaVyplnePlna * prostupVyplnePlna + _
        obvodZaskleni * linearniCinitelProstupu + _
        obvodProskleniCreative * linearniCinitelProstupuProskleniCreative + _
        plochaProfiluFixuB * prostupProfiluFixuB + _
        plochaVyplneFixuB * prostupVyplneFixuB + _
        obvodZaskleniFixuB * linearniCinitelProstupuFixuB + _
        plochaProfiluFixuC * prostupProfiluFixuC + _
        plochaVyplneFixuC * prostupVyplneFixuC + _
        obvodZaskleniFixuC * linearniCinitelProstupuFixuC + _
        plochaProfiluFixuD * prostupProfiluFixuD + _
        plochaVyplneFixuD * prostupVyplneFixuD + _
        obvodZaskleniFixuD * linearniCinitelProstupuFixuD
    
    prostupSestavySpodekZlomku = _
        plochaProfiluHorni + _
        plochaProfiluSpodni + _
        plochaProfiluPricky + _
        plochaProfiluSrazu + _
        plochaVyplne + _
        plochaVyplneSklo + _
        plochaVyplnePlna + _
        plochaProfiluFixuB + _
        plochaVyplneFixuB + _
        plochaProfiluFixuC + _
        plochaVyplneFixuC + _
        plochaProfiluFixuD + _
        plochaVyplneFixuD
    
    prostupSestavy = prostupSestavyVrsekZlomku / prostupSestavySpodekZlomku
    
    Worksheets("zadani").Range("G5").Value = prostupSestavy
    
End Sub
