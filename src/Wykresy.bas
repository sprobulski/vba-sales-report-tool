Attribute VB_Name = "Wykresy"
'Buduje wszystkie wykresy w raporcie
Sub ZbudujWykresyNaDashboardzie()

    Dim wsDash As Worksheet, wsCalc As Worksheet
    Dim ptTydzien As PivotTable, ptBrand As PivotTable, ptProdukty As PivotTable
    Dim ptTopReg As PivotTable
    Dim ptTopMarki As PivotTable
    Dim chrt As ChartObject
    
    'Zmienne do trendu
    Dim rngY As Range, rngX As Range
    Dim nachylenie As Double
    Dim kolorTrendu As Long, opisTrendu As String
    Dim tl As Trendline
    Dim ostatniX As Variant
    
    'Sta³e do pozycji wykresów
    Const LEWY_MARGINES As Double = 310
    Const SZEROKOSC_DUZA As Double = 580
    Const SZEROKOSC_MALA As Double = 285
    
    Application.ScreenUpdating = False
    
    On Error Resume Next
    ThisWorkbook.Sheets("DASHBOARD").Unprotect Password:=""
    On Error GoTo 0
    
    Set wsDash = ThisWorkbook.Sheets("DASHBOARD")
    Set wsCalc = ThisWorkbook.Sheets("OBLICZENIA")
    
    'Czyszczenie starych wykresów
    For Each chrt In wsDash.ChartObjects
        chrt.Delete
    Next chrt
    
    On Error Resume Next
    Set ptTydzien = wsCalc.PivotTables("PT_Tydzien")
    Set ptBrand = wsCalc.PivotTables("PT_Brand")
    Set ptProdukt = wsCalc.PivotTables("PT_Produkt")
    Set ptTopReg = wsCalc.PivotTables("PT_TopRegiony")
    Set ptTopProdukty = wsCalc.PivotTables("PT_TopProdukty")
    On Error GoTo 0
    
    If ptTydzien Is Nothing Then MsgBox "Uruchom najpierw Silnik!", vbCritical: Exit Sub

    'TREND SPRZEDA¯Y WED£UG TYGODNI
    nachylenie = 0
    kolorTrendu = RGB(160, 160, 160)
    opisTrendu = " (Stabilny)"
    
    If ptTydzien.DataBodyRange.Rows.Count >= 2 Then
        Set rngX = ptTydzien.RowRange.Offset(1, 0).Resize(ptTydzien.RowRange.Rows.Count - 1, 1)
        Set rngY = ptTydzien.DataBodyRange.Columns(1)
        
        ' Usuniêcie sumy koñcowej
        ostatniX = rngX.Cells(rngX.Rows.Count, 1).Value
        If Not IsNumeric(ostatniX) Then
            Set rngX = rngX.Resize(rngX.Rows.Count - 1, 1)
            Set rngY = rngY.Resize(rngY.Rows.Count - 1, 1)
        End If
        
        'Zmiana koloru w zale¿noœci od trendu
        nachylenie = Application.WorksheetFunction.Slope(rngY, rngX)
        
        If nachylenie > 0.5 Then
            kolorTrendu = RGB(0, 176, 80)
            opisTrendu = " (Wzrost)"
        ElseIf nachylenie < -0.5 Then
            kolorTrendu = RGB(255, 0, 0)
            opisTrendu = " (Spadek)"
        End If
    End If
    
    'Wstawienie wykresu trendu
    Set chrt = wsDash.ChartObjects.Add(Left:=LEWY_MARGINES, Top:=400, Width:=SZEROKOSC_DUZA, Height:=200)
    With chrt.Chart
        FormatujWykres chrt
        
        .ChartType = xlLineMarkers
        .SeriesCollection.NewSeries
        .SeriesCollection(1).XValues = rngX
        .SeriesCollection(1).Values = rngY
        .SeriesCollection(1).Name = "Sprzeda¿"
        
        .HasTitle = True
        .ChartTitle.Text = "Trend Sprzeda¿y" & opisTrendu
        
        ' Zmiana koloru tytu³u w zale¿noœci od nachylenia
        If nachylenie > 0.5 Then
            .ChartTitle.Font.Color = RGB(0, 176, 80)
        ElseIf nachylenie < -0.5 Then
            .ChartTitle.Font.Color = RGB(255, 0, 0)
        Else
            .ChartTitle.Font.Color = RGB(80, 80, 80)
        End If
        
        .Legend.Delete
        
        If rngY.Rows.Count >= 2 Then
            Set tl = .SeriesCollection(1).Trendlines.Add(Type:=xlLinear)
            With tl.Format.Line
                .Visible = msoTrue
                .ForeColor.RGB = kolorTrendu
                .Weight = 2.5
                .DashStyle = msoLineDash
            End With
        End If
    End With
    

    '5 Najlepszych regionów oraz produktów
    
    'RANKING regionów (Lewa strona)
        Set chrt = wsDash.ChartObjects.Add(Left:=LEWY_MARGINES, Top:=215, Width:=SZEROKOSC_MALA, Height:=160)
        With chrt.Chart
            .SetSourceData Source:=ptTopReg.TableRange1
            .ChartType = xlBarClustered
            .HasTitle = True: .ChartTitle.Text = "Top 5 Regionów"
            .Legend.Delete
            .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(31, 78, 120) ' Granat
            .Axes(xlCategory).ReversePlotOrder = True ' Nr 1 na górze
            FormatujWykres chrt
        End With

    'RANKING Produktow (Prawa strona)
        Set chrt = wsDash.ChartObjects.Add(Left:=LEWY_MARGINES + SZEROKOSC_MALA + 15, Top:=215, Width:=SZEROKOSC_MALA, Height:=160)
        With chrt.Chart
            .SetSourceData Source:=ptTopProdukty.TableRange1
            .ChartType = xlBarClustered
            .HasTitle = True
            .ChartTitle.Text = "Top 5 Produktów w kraju"
            .Legend.Delete
            .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(237, 125, 49)
            .Axes(xlCategory).ReversePlotOrder = True
            
            FormatujWykres chrt
        End With




    ' Top Produkty (Ko³owy)
    Set chrt = wsDash.ChartObjects.Add(Left:=LEWY_MARGINES, Top:=610, Width:=SZEROKOSC_MALA, Height:=180)
    With chrt.Chart
        .SetSourceData Source:=ptProdukt.TableRange1
        .ChartType = xlDoughnut
        .HasTitle = True: .ChartTitle.Text = "Top Produkty (Udzia³ %)"
        On Error Resume Next: .ChartGroups(1).DoughnutHoleSize = 50: On Error GoTo 0
        .ApplyDataLabels
        On Error Resume Next
        With .SeriesCollection(1).DataLabels
            .ShowPercentage = True: .ShowValue = False: .ShowCategoryName = False
            .Font.Color = RGB(255, 255, 255): .Font.Bold = True
        End With
        On Error GoTo 0
        .HasLegend = True: .Legend.Position = xlLegendPositionRight: .Legend.Format.Line.Visible = msoFalse
        FormatujWykres chrt
    End With
    
    ' Udzia³ Marek (S³upkowy)
    Set chrt = wsDash.ChartObjects.Add(Left:=LEWY_MARGINES + SZEROKOSC_MALA + 10, Top:=610, Width:=SZEROKOSC_MALA, Height:=180)
    With chrt.Chart
        .SetSourceData Source:=ptBrand.TableRange1
        .ChartType = xlColumnClustered
        .HasTitle = True: .ChartTitle.Text = "Sprzeda¿ wg Marek": .Legend.Delete
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(31, 78, 120)
        FormatujWykres chrt
    End With
    
    'Porównanie œredniego koszyka dla zaznaczonego filtra z ca³ym krajem
    wsCalc.Range("AA1").Value = "Typ": wsCalc.Range("AB1").Value = "Œrednia"
    wsCalc.Range("AA2").Value = "Twój Wybór": wsCalc.Range("AA3").Value = "Kraj"
    wsCalc.Range("AB2").Formula = "=IFERROR(GETPIVOTDATA(""Srednia_Twoja"",$P$3),0)"
    wsCalc.Range("AB3").Formula = "=IFERROR(GETPIVOTDATA(""Srednia_Kraj"",$T$3),0)"
    
    Set chrt = wsDash.ChartObjects.Add(Left:=LEWY_MARGINES, Top:=800, Width:=SZEROKOSC_DUZA, Height:=180)
    With chrt.Chart
        .SetSourceData Source:=wsCalc.Range("AA1:AB3")
        .ChartType = xlBarClustered
        .HasTitle = True: .ChartTitle.Text = "Œredni koszyk (vs Kraj)": .Legend.Delete
        On Error Resume Next
        .SeriesCollection(1).Points(1).Format.Fill.ForeColor.RGB = RGB(237, 125, 49)
        .SeriesCollection(1).Points(2).Format.Fill.ForeColor.RGB = RGB(200, 200, 200)
        On Error GoTo 0
        FormatujWykres chrt
    End With
    


    wsDash.Select
    
    Application.ScreenUpdating = True

End Sub

Sub FormatujWykres(chrt As ChartObject)
    With chrt.Chart
        'Usuwa t³o i ramkê
        .ChartArea.Format.Fill.Visible = msoFalse
        .ChartArea.Format.Line.Visible = msoFalse
        .PlotArea.Format.Fill.Visible = msoFalse

        On Error Resume Next
        .ShowAllFieldButtons = False
        On Error GoTo 0
        
        'Czcionka tytu³u
        If .HasTitle Then
            .ChartTitle.Font.Name = "Arial"
            .ChartTitle.Font.Size = 14
            .ChartTitle.Font.Color = RGB(64, 64, 64)
        End If
        
        'Usuwa linie siatki
        On Error Resume Next
        .Axes(xlValue).MajorGridlines.Format.Line.Visible = msoFalse
        On Error GoTo 0
    End With
End Sub


'Tworzy 3 kafelki KPI: sprzeda¿ ca³kowita, liczba transakcji i œrednia transakcja
Sub RysujKPI()

    Dim wsDash As Worksheet, wsCalc As Worksheet
    Dim shp As Shape

    Set wsDash = ThisWorkbook.Sheets("DASHBOARD")
    Set wsCalc = ThisWorkbook.Sheets("OBLICZENIA")
    
    'Przygotowanie danych (Arkusz OBLICZENIA)

    
    'Ca³kowita Sprzeda¿ (Pobieramy pole "Suma_Tydzien")
    wsCalc.Range("AD1").Formula = "=GETPIVOTDATA(""Suma_Tydzien"",$A$3)"
    wsCalc.Range("AD1").NumberFormat = "#,##0 ""PLN""" ' Formatowanie komórki przeniesie siê na kafelek!
    
    'Liczba Transakcji (Pobieramy pole "Liczba_Transakcji")
    wsCalc.Range("AD2").Formula = "=GETPIVOTDATA(""Liczba_Transakcji"",$A$3)"
    wsCalc.Range("AD2").NumberFormat = "#,##0"
    
    'Œredni Koszyk (Sprzeda¿ / Transakcje)
    wsCalc.Range("AD3").Formula = "=IFERROR(AD1/AD2, 0)" 'zabezpieczenie przed dzieleniem przez 0
    wsCalc.Range("AD3").NumberFormat = "#,##0 ""PLN"""
    
    'Usuniêcie starych KPI
    On Error Resume Next
    For Each shp In wsDash.Shapes
        If shp.Name Like "KPI_*" Then shp.Delete
    Next shp
    On Error GoTo 0
    
    'Rysuje 3 KPI za pomoc¹ RysujKafelek
    'KPI 1: SPRZEDA¯
    RysujKafelek "KPI_Sales", "OBLICZENIA!$AD$1", 312, 120, "SPRZEDA¯ CA£KOWITA"
    
    'KPI 2: TRANSAKCJE
    RysujKafelek "KPI_Trans", "OBLICZENIA!$AD$2", 507, 120, "LICZBA TRANSAKCJI"
    
    'KPI 3: ŒREDNI KOSZYK
    RysujKafelek "KPI_Basket", "OBLICZENIA!$AD$3", 710, 120, "ŒREDNI KOSZYK"
    
    wsDash.Select


End Sub

' Funkcja robi¹ca jeden kafelek
Private Sub RysujKafelek(nazwa As String, adresZrodla As String, l As Double, t As Double, tytul As String)
    Dim ws As Worksheet
    Dim shpBg As Shape, shpVal As Shape, shpTitle As Shape
    
    Set ws = ActiveSheet
    
    'T³o
    Set shpBg = ws.Shapes.AddShape(msoShapeRectangle, l, t, 180, 80)
    With shpBg
        .Name = nazwa & "_BG"
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(31, 78, 120)
    End With
    
    'Tytu³
    Set shpTitle = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, l + 15, t + 10, 150, 20)
    With shpTitle
        .Name = nazwa & "_Title"
        .TextFrame.Characters.Text = tytul
        .TextFrame.Characters.Font.Size = 9
        .TextFrame.Characters.Font.Color = RGB(120, 120, 120) ' Szary kolor tekstu
        .TextFrame.Characters.Font.Name = "Arial"
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
    End With
    
    'Wartoœæ
    Set shpVal = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, l + 15, t + 30, 150, 40)
    With shpVal
        .Name = nazwa & "_Val"
        .DrawingObject.Formula = "=" & adresZrodla
        
        ' Formatowanie wartoœci
        .TextFrame.Characters.Font.Size = 20
        .TextFrame.Characters.Font.Bold = True
        .TextFrame.Characters.Font.Color = RGB(31, 78, 120)
        .TextFrame.Characters.Font.Name = "Arial"
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        .TextFrame.HorizontalAlignment = xlHAlignCenter
    End With
    
End Sub


'Makro odswierzaj¹ce wykrs trendu gdy u¿ytkownik zmieni filtr
Sub OdswiezTylkoTrend()
    Dim wsDash As Worksheet, wsCalc As Worksheet
    Dim ptTydzien As PivotTable
    Dim chrt As ChartObject
    Dim nachylenie As Double
    Dim kolorTrendu As Long, opisTrendu As String
    Dim rngY As Range, rngX As Range
    Dim ostatniX As Variant
    Dim punktyDanych As Long
    
    Application.ScreenUpdating = False
    
    'Zdjêcie has³a
    On Error Resume Next
    ThisWorkbook.Sheets("DASHBOARD").Unprotect Password:=""
    On Error GoTo 0
    
    Set wsDash = ThisWorkbook.Sheets("DASHBOARD")
    Set wsCalc = ThisWorkbook.Sheets("OBLICZENIA")
    
    ' Przypisanie tabeli przestawnej z danymi
    On Error Resume Next
    Set ptTydzien = wsCalc.PivotTables("PT_Tydzien")
    On Error GoTo 0
    If ptTydzien Is Nothing Then Exit Sub
    
    'Pêtla szukaj¹ca na dashboardzie wyrkesu trendu
    Dim c As ChartObject
    For Each c In wsDash.ChartObjects
        If c.Chart.HasTitle Then
            If Left(c.Chart.ChartTitle.Text, 5) = "Trend" Then
                Set chrt = c
                Exit For
            End If
        End If
    Next c
    If chrt Is Nothing Then Exit Sub
    
    
    nachylenie = 0
    kolorTrendu = RGB(160, 160, 160)
    opisTrendu = " (Brak danych)"
    Set rngX = Nothing
    Set rngY = Nothing
    punktyDanych = 0
    
    If ptTydzien.RowRange.Rows.Count > 1 Then
        
        'Pobranie zakresu osi x(tygodni)
        Set rngX = ptTydzien.RowRange.Offset(1, 0).Resize(ptTydzien.RowRange.Rows.Count - 1, 1)
        
        'Pobranie zakresu osi y (wartoœci)
        Set rngY = ptTydzien.DataBodyRange.Columns(1)

        
        'Usuniêcie ostatniego wiersza jeœli to suma koñcowa
        If Not rngX Is Nothing Then
            On Error Resume Next
            ostatniX = rngX.Cells(rngX.Rows.Count, 1).Value
            If Not IsNumeric(ostatniX) Then
                If rngX.Rows.Count > 1 Then
                    Set rngX = rngX.Resize(rngX.Rows.Count - 1, 1)
                    Set rngY = rngY.Resize(rngY.Rows.Count - 1, 1)
                End If
            End If
            On Error GoTo 0
            
            punktyDanych = rngX.Rows.Count
        End If
        
        'Liczenie nachylenia
        If punktyDanych >= 2 Then
            On Error Resume Next
            nachylenie = Application.WorksheetFunction.Slope(rngY, rngX)
            If Err.Number = 0 Then
                opisTrendu = " (Stabilny)"
                'kolor zielony dla wzrostu (>0.5) i czerwony dla spadku (<-0.5)
                If nachylenie > 0.5 Then kolorTrendu = RGB(0, 176, 80): opisTrendu = " (Wzrost)"
                If nachylenie < -0.5 Then kolorTrendu = RGB(255, 0, 0): opisTrendu = " (Spadek)"
            Else
                opisTrendu = " (Zbyt ma³o danych)"
            End If
            On Error GoTo 0
        End If
    End If
    
    ' Aktualizacja wykresu
    With chrt.Chart
        
        'Jeœli wykres jest pusty, dodaje now¹ seriê danych
        If .SeriesCollection.Count = 0 Then
            .SeriesCollection.NewSeries
        End If
        
        'Przypisanie danych do wykresu
        If punktyDanych > 0 Then
            On Error Resume Next
            .SeriesCollection(1).Values = rngY
            .SeriesCollection(1).XValues = rngX
            On Error GoTo 0
        Else
            On Error Resume Next
            .SeriesCollection(1).Values = "={0}"
            .SeriesCollection(1).XValues = "={0}"
            On Error GoTo 0
        End If
        
        .ChartTitle.Text = "Trend Sprzeda¿y" & opisTrendu
        
        ' Kolory tytu³u
        On Error Resume Next
        If nachylenie > 0.5 Then
            .ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 176, 80)
        ElseIf nachylenie < -0.5 Then
            .ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 0, 0)
        Else
            .ChartTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(80, 80, 80)
        End If
        On Error GoTo 0
        
        'Rysowanie linii trendu tylko, gdy s¹ min. 2 punkty
        On Error Resume Next
        If punktyDanych >= 2 Then
            If .SeriesCollection(1).Trendlines.Count = 0 Then .SeriesCollection(1).Trendlines.Add
            With .SeriesCollection(1).Trendlines(1).Format.Line
                .Visible = msoTrue
                .ForeColor.RGB = kolorTrendu
                .DashStyle = msoLineDash
            End With
        Else
            If .SeriesCollection(1).Trendlines.Count > 0 Then .SeriesCollection(1).Trendlines(1).Delete
        End If
        On Error GoTo 0
    End With
    
    Application.ScreenUpdating = True
    
    'Zablokowanie arkusza
    Call ZabezpieczWidok

End Sub
