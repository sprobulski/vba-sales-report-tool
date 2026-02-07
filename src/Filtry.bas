Attribute VB_Name = "Filtry"
Option Explicit

'Buduje tabele przestawne i slicery
Sub ZbudujSilnikDashboardu()

    Dim wsBaza As Worksheet, wsCalc As Worksheet, wsDash As Worksheet
    Dim pc As PivotCache
    Dim pt1 As PivotTable, pt2 As PivotTable, pt3 As PivotTable, pt4 As PivotTable, pt5 As PivotTable
    Dim ptAvgLok As PivotTable, ptAvgKraj As PivotTable
    Dim rngDane As Range
    Dim pf As PivotField
    Dim cache As SlicerCache
    Dim sl As Slicer
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    On Error Resume Next
    ThisWorkbook.Sheets("DASHBOARD").Unprotect Password:=""
    On Error GoTo 0
    


    'Ustawnienia arkuszy
    On Error Resume Next
    Set wsBaza = ThisWorkbook.Sheets("BAZA")
    On Error GoTo 0
    If wsBaza Is Nothing Then
        MsgBox "Brak arkusza BAZA!", vbCritical
        Application.EnableEvents = True
        Exit Sub
    End If
    
    On Error Resume Next
    Set wsCalc = ThisWorkbook.Sheets("OBLICZENIA")
    On Error GoTo 0
    
    If wsCalc Is Nothing Then
        Set wsCalc = ThisWorkbook.Sheets.Add(After:=wsBaza)
        wsCalc.Name = "OBLICZENIA"
    Else
        wsCalc.Rows.Delete ' Czyœci pamiêæ tabeli przestawnych
    End If
    
    On Error Resume Next
    Set wsDash = ThisWorkbook.Sheets("DASHBOARD")
    On Error GoTo 0
    If wsDash Is Nothing Then
        Set wsDash = ThisWorkbook.Sheets.Add(After:=wsCalc)
        wsDash.Name = "DASHBOARD"
    End If
    
    Set rngDane = wsBaza.Range("A1").CurrentRegion
    
    If rngDane.Rows.Count < 2 Then
        MsgBox "Baza danych jest pusta!", vbCritical
        Application.EnableEvents = True
        Exit Sub
    End If
    
    ' Slinik
    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rngDane)
        
    'TABELE PRZESTAWNE
    
    'PT_Tydzien
    Set pt1 = pc.CreatePivotTable(wsCalc.Range("A3"), "PT_Tydzien")
    With pt1
        .PivotFields("Tydzien").Orientation = xlRowField
        Set pf = .AddDataField(.PivotFields("Sprzedaz"), "Suma_Tydzien", xlSum)
        pf.NumberFormat = "#,##0 ""PLN"""
        .AddDataField .PivotFields("Sprzedaz"), "Liczba_Transakcji", xlCount
        .TableStyle2 = ""
    End With
    
    'PT_Brand
    Set pt2 = pc.CreatePivotTable(wsCalc.Range("F3"), "PT_Brand")
    With pt2
        .PivotFields("Brand").Orientation = xlRowField
        Set pf = .AddDataField(.PivotFields("Sprzedaz"), "Suma_Brand", xlSum)
        pf.NumberFormat = "#,##0 ""PLN"""
        .TableStyle2 = ""
    End With
    
    'PT_Produkt
    Set pt3 = pc.CreatePivotTable(wsCalc.Range("K3"), "PT_Produkt")
    With pt3
        .PivotFields("Produkt").Orientation = xlRowField
        Set pf = .AddDataField(.PivotFields("Sprzedaz"), "Suma_Produkt", xlSum)
        pf.NumberFormat = "#,##0 ""PLN"""
        .PivotFields("Produkt").AutoSort xlDescending, "Suma_Produkt"
        .TableStyle2 = ""
    End With
    
    'PT_TopRegiony (Ranking Top 5 regionów)
    Set pt4 = pc.CreatePivotTable(wsCalc.Range("AA20"), "PT_TopRegiony")
    With pt4
        .PivotFields("Wojewodztwo").Orientation = xlRowField

        Set pf = .AddDataField(.PivotFields("Sprzedaz"), "Suma_Woj", xlSum)
        pf.NumberFormat = "#,##0"
        .PivotFields("Wojewodztwo").AutoSort xlDescending, "Suma_Woj" 'sortowanie
        .PivotFields("Wojewodztwo").ClearAllFilters
        .PivotFields("Wojewodztwo").PivotFilters.Add2 Type:=xlTopCount, DataField:=pf, Value1:=5
        
        .TableStyle2 = ""
        .ColumnGrand = False
        .RowGrand = False
    End With
    
    
'PT_TopProdukty (Ranking Top 5 Produktów)
    Set pt5 = pc.CreatePivotTable(wsCalc.Range("AG20"), "PT_TopProdukty")
    
    With pt5
        ' Zmieniamy na PRODUKT
        .PivotFields("Produkt").Orientation = xlRowField
        
        ' dodanie sprzeda¿y
        Set pf = .AddDataField(.PivotFields("Sprzedaz"), "Suma_Prod_Top", xlSum)
        pf.NumberFormat = "#,##0"
        
        .PivotFields("Produkt").AutoSort xlDescending, "Suma_Prod_Top" 'sortowanie
        
        'Filtr na top 5
        .PivotFields("Produkt").ClearAllFilters
        .PivotFields("Produkt").PivotFilters.Add2 Type:=xlTopCount, DataField:=pf, Value1:=5
        
        .TableStyle2 = ""
        .ColumnGrand = False
        .RowGrand = False
    End With
    
    'Porównanie œredniego koszyka zaznaczonych filtrów z krajem
    Set ptAvgLok = pc.CreatePivotTable(wsCalc.Range("P3"), "PT_Srednia_Lok")
    With ptAvgLok
        Set pf = .AddDataField(.PivotFields("Sprzedaz"), "Srednia_Twoja", xlAverage)
        pf.NumberFormat = "#,##0 ""PLN"""
        .TableStyle2 = ""
    End With
    
    Set ptAvgKraj = pc.CreatePivotTable(wsCalc.Range("T3"), "PT_Srednia_Kraj")
    With ptAvgKraj
        Set pf = .AddDataField(.PivotFields("Sprzedaz"), "Srednia_Kraj", xlAverage)
        pf.NumberFormat = "#,##0 ""PLN"""
        .TableStyle2 = ""
    End With

    'SLICERY
    
    'Usuwanie starych
    Dim s As Shape, scTemp As SlicerCache
    For Each s In wsDash.Shapes
        If s.Type = msoSlicer Then s.Delete
    Next s
    For Each scTemp In ThisWorkbook.SlicerCaches
        On Error Resume Next
        scTemp.Delete
        On Error GoTo 0
    Next scTemp
    
    
    ' A. TYDZIEÑ
    Set cache = ThisWorkbook.SlicerCaches.Add(pt1, "Tydzien", "Cache_Tydzien")
    Set sl = cache.Slicers.Add(wsDash, , "Slicer_Tydzien", "Wybierz Tydzieñ", 394, 10, 293, 206)
    sl.NumberOfColumns = 7
    sl.Style = "Moj_styl"
    
    ' B. WOJEWÓDZTWO
    Set cache = ThisWorkbook.SlicerCaches.Add(pt1, "Wojewodztwo", "Cache_Woj")
    Set sl = cache.Slicers.Add(wsCalc, , "Slicer_Woj", "Województwo", 0, 0, 100, 100)
    
    ' C. MARKA
    Set cache = ThisWorkbook.SlicerCaches.Add(pt1, "Brand", "Cache_Brand")
    Set sl = cache.Slicers.Add(wsDash, , "Slicer_Brand", "Marka", 612, 10, 293, 180)
    sl.Style = "Moj_styl"
    
    ' D. PRODUKT
    Set cache = ThisWorkbook.SlicerCaches.Add(pt1, "Produkt", "Cache_Prod")
    Set sl = cache.Slicers.Add(wsDash, , "Slicer_Prod", "Produkt", 804, 10, 293, 180)
    sl.Style = "Moj_styl"
    
    'Po³¹czenie slicerów
    For Each cache In ThisWorkbook.SlicerCaches
        On Error Resume Next
        
        ' Pod³¹czenie tabeli przestawnych
        cache.PivotTables.AddPivotTable pt2
        cache.PivotTables.AddPivotTable pt3
        cache.PivotTables.AddPivotTable ptAvgLok
        
        On Error GoTo 0
    Next cache
    
    
    'Usuwanie pola "puste" ze slicerów
    Dim scClean As SlicerCache
    
    For Each scClean In ThisWorkbook.SlicerCaches
        On Error Resume Next
        
        ' Próbujemy odznaczyæ polsk¹ nazwê
        scClean.SlicerItems("(puste)").Selected = False
        
        
        On Error GoTo 0
    Next scClean
    
    wsCalc.Visible = xlSheetHidden
    wsDash.Visible = xlSheetVisible
    wsDash.Activate
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True


End Sub

'Makro pozwalaj¹ce zaznaczyæ województwo za pomoc¹ mapy
Sub KlikniecieMapy()

    Dim nazwaKsztaltu As String
    Dim sc As SlicerCache
    Dim slItem As SlicerItem
    Dim ws As Worksheet
    Dim czyZnaleziono As Boolean
    
    'Pobranie nazwy wybranego kszta³tu
    On Error Resume Next
    nazwaKsztaltu = Application.Caller
    On Error GoTo 0
    
    If nazwaKsztaltu = "" Then Exit Sub
    
    Set ws = ActiveSheet
    
    ' Znajduje slicer
    On Error Resume Next
    Set sc = ThisWorkbook.SlicerCaches("Cache_Woj")
    On Error GoTo 0
    
    If sc Is Nothing Then MsgBox "B³¹d: Brak Slicera Województw!", vbCritical: Exit Sub
    
    Application.ScreenUpdating = False
    
    sc.ClearManualFilter ' Najpierw czyœcimy, ¿eby szukaæ w pe³nej liœcie
    
    czyZnaleziono = False
    
    ' Sprawdza, czy takie województwo istnieje w danych (w Slicerze)
    On Error Resume Next
    For Each slItem In sc.SlicerItems
        ' porównanie nazwy
        If UCase(slItem.Name) = UCase(nazwaKsztaltu) Then
            
            'gdy znaleziono województwo w danych, zaznacza i koloruje
            sc.VisibleSlicerItemsList = Array(slItem.Name)
            If Err.Number <> 0 Then
                Err.Clear
                slItem.Selected = True
                Dim item As SlicerItem
                For Each item In sc.SlicerItems
                    If item.Name <> slItem.Name Then item.Selected = False
                Next item
            End If
            
            czyZnaleziono = True
            Exit For
        End If
    Next slItem
    On Error GoTo 0
    
    'Zawsze najpierw resetuje wszystko na szaro
    ZresetujKoloryMapy
    
    If czyZnaleziono Then
        'Jeœli dane s¹ koloruje województwo
        ws.Shapes(nazwaKsztaltu).Fill.ForeColor.RGB = RGB(237, 125, 49)
    Else
        'jeœli brak danych (np. Warmiñsko-Mazurskie), kolor pozostaje szary i wyœwietla siê komunikat
        MsgBox "Brak sprzeda¿y w województwie: " & nazwaKsztaltu, vbInformation, "Informacja"
        sc.ClearManualFilter
    End If
    
    Application.ScreenUpdating = True
    Call ZabezpieczWidok

End Sub

'Makro koloruj¹ce wszystkie województwa na szaro (¿adne województwo nie jest zaznaczone)
Sub ZresetujKoloryMapy()

    Dim arrWoj As Variant
    Dim v As Variant
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets("DASHBOARD")
    
    arrWoj = Array("Dolnoslaskie", "Kujawskopomorskie", "Lubelskie", "Lubuskie", _
                   "Lodzkie", "Malopolskie", "Mazowieckie", "Opolskie", _
                   "Podkarpackie", "Podlaskie", "Pomorskie", "Slaskie", _
                   "Swietokrzyskie", "Warminskomazurskie", "Wielkopolskie", "Zachodniopomorskie")
                   
    On Error Resume Next
    For Each v In arrWoj
        ws.Shapes(v).Fill.ForeColor.RGB = RGB(191, 191, 191)
    Next v
    On Error GoTo 0
    Call ZabezpieczWidok
End Sub

'Czysci mapê za pomoc¹ przycisku X
Sub WyczyscMape()

    On Error Resume Next
    ThisWorkbook.SlicerCaches("Cache_Woj").ClearManualFilter
    On Error GoTo 0
    
    ZresetujKoloryMapy
    
    Run "OdswiezTylkoTrend"
    Call ZabezpieczWidok
End Sub
