Attribute VB_Name = "Import_danych"
'Makro zabezpieczaj¹ce arkusz przed ingerencj¹ przez u¿ytkownika
Sub ZabezpieczWidok()
    
    Dim wsDash As Worksheet
    Dim shp As Shape
    Dim sc As SlicerCache
    Dim sl As Slicer
    Dim czyMaMakro As Boolean
    
    Set wsDash = ThisWorkbook.Sheets("DASHBOARD")
    wsDash.Activate
    
    'Zdejmuje star¹ blokadê
    On Error Resume Next
    wsDash.Unprotect Password:=""
    On Error GoTo 0
    
    'Wygl¹d Aplikacji
    With ActiveWindow
        .DisplayGridlines = False
        .DisplayHeadings = False
    End With
    
    'Zabezpieczenia silcerów
    For Each sc In ThisWorkbook.SlicerCaches
        For Each sl In sc.Slicers
            ' Dzia³a tylko na slicerach z tego Dashboardu
            If sl.Shape.Parent.Name = wsDash.Name Then
                sl.Shape.Locked = False       ' Pozwól klikaæ
                sl.DisableMoveResizeUI = True ' Nie pozwala przesuwaæ
            End If
        Next sl
    Next sc
    
    'Reszta zabezpieczeñ
    For Each shp In wsDash.Shapes
        
        'je¿eli to slicer to nic siê nie dzieje
        If shp.Type = msoSlicer Then
            
        Else
            ' Sprawdza czy to przycisk
            czyMaMakro = False
            On Error Resume Next
            If shp.OnAction <> "" Then czyMaMakro = True
            On Error GoTo 0
            
            ' Je¿eli obiekt to przycisk to nie jest blokowany
            If shp.Type = msoFormControl Or czyMaMakro = True Then
                shp.Locked = False
                
            ' Wszystko inne jest blokowane
            Else
                shp.Locked = True
            End If
        End If
        
    Next shp
    
    'Za³o¿enie blokady
    wsDash.Protect Password:="", _
        DrawingObjects:=True, _
        Contents:=True, _
        Scenarios:=True, _
        UserInterfaceOnly:=True, _
        AllowUsingPivotTables:=True
        
    wsDash.Range("A1").Select
    
End Sub


' Makro importuj¹ce dane
Sub ImportujDane()

    Dim wsBaza As Worksheet
    Dim fd As FileDialog
    Dim plik As Variant
    Dim wbZrodlo As Workbook
    Dim wierszZrodlo As Long, nastepnyWiersz As Long, ostatniWiersz As Long, i As Long
    Dim tekst As String
    Dim czesci() As String
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    
    'Przygotowanie arkusza baza
    On Error Resume Next
    Set wsBaza = ThisWorkbook.Sheets("BAZA")
    On Error GoTo 0
    If wsBaza Is Nothing Then
        Set wsBaza = ThisWorkbook.Sheets.Add
        wsBaza.Name = "BAZA"
    End If
    
    wsBaza.Cells.Clear
    
    'Nag³ówki
    With wsBaza
        .Range("A1:F1").Value = Array("Brand", "Produkt", "Tydzien", "Sprzedaz", "Wojewodztwo", "Miasto")
        .Range("A1:F1").Font.Bold = True
        .Range("A1:F1").Interior.Color = RGB(31, 78, 120)
        .Range("A1:F1").Font.Color = vbWhite
    End With
    
    'Import danych
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Wybierz pliki tekstowe"
        .Filters.Clear
        .Filters.Add "Pliki tekstowe", "*.txt"
        .AllowMultiSelect = True
        
        If .Show = -1 Then
            For Each plik In .SelectedItems
                Workbooks.OpenText Filename:=plik, Origin:=65001, DataType:=xlDelimited, Tab:=True, Local:=True
                Set wbZrodlo = ActiveWorkbook
                
                wierszZrodlo = wbZrodlo.Sheets(1).Cells(Rows.Count, "A").End(xlUp).Row
                
                If wierszZrodlo >= 2 Then
                    nastepnyWiersz = wsBaza.Cells(wsBaza.Rows.Count, "A").End(xlUp).Row + 1
                    wbZrodlo.Sheets(1).Range("A2:F" & wierszZrodlo).Copy wsBaza.Cells(nastepnyWiersz, 1)
                End If
                
                wbZrodlo.Close SaveChanges:=False
            Next plik
        Else
            Application.ScreenUpdating = True
            Exit Sub
        End If
    End With
    
    
    'Rozdzielenie kolumny z województwem i miastem na 2 odzielne
    ostatniWiersz = wsBaza.Cells(wsBaza.Rows.Count, "A").End(xlUp).Row
    For i = 2 To ostatniWiersz

        tekst = wsBaza.Cells(i, 5).Value ' Pobieramy tekst z Województwa (Kol E)

        If InStr(tekst, "-") > 0 Then
            czesci = Split(tekst, "-")
            wsBaza.Cells(i, 5).Value = czesci(0)
            If UBound(czesci) >= 1 Then 'Rozdzielenie miast dwucz³onowych
                wsBaza.Cells(i, 6).Value = czesci(1)
            End If
        End If

    Next i

    
    wsBaza.Columns("A:F").AutoFit
    ThisWorkbook.RefreshAll
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

Sub ZbudujRaport()
    Dim wsBaza As Worksheet
    
    'Import Danych
    Call ImportujDane
    
    ' Sprawdzenie, czy w Bazie s¹ jakieœ dane
    Set wsBaza = ThisWorkbook.Sheets("BAZA")
    If wsBaza.Range("A1").CurrentRegion.Rows.Count < 2 Then
        MsgBox "Nie zaimportowano danych. Raport nie zostanie odœwie¿ony.", vbExclamation
        Exit Sub
    End If

    'Budowa Raportu
    Application.Cursor = xlWait
    
    Call ZbudujSilnikDashboardu
    Call ZbudujWykresyNaDashboardzie
    Call RysujKPI
    
    ' 4. Finalne zabezpieczenie
    Call ZabezpieczWidok
    
    Application.Cursor = xlDefault
    MsgBox "Proces zakoñczony pomyœlnie!", vbInformation

End Sub

'Makro generuj¹ce raport jako plik pdf
Sub GenerujRaportPDF()

    Dim wsDash As Worksheet
    Dim sciezka As String, nazwaPliku As String
    Dim pelnaSciezka As String
    
    Set wsDash = ThisWorkbook.Sheets("DASHBOARD")
    wsDash.Activate
    
    Application.ScreenUpdating = False
    
    
    'Ustawienia Strony
    With wsDash.PageSetup
        Application.PrintCommunication = False
        
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        
        ' Marginesy
        .LeftMargin = Application.InchesToPoints(0.2)
        .RightMargin = Application.InchesToPoints(0.2)
        .TopMargin = Application.InchesToPoints(0.2)
        .BottomMargin = Application.InchesToPoints(0.2)
       
        .CenterHorizontally = True
        
        Application.PrintCommunication = True
    End With
    
    'Œcie¿ka pliku
    nazwaPliku = "Raport_Sprzedazy_" & Format(Date, "yyyy-mm-dd") & ".pdf"
    pelnaSciezka = ThisWorkbook.Path & "\" & nazwaPliku
    
    'Zapis do PDF
    On Error GoTo BladZapisu
    wsDash.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=pelnaSciezka, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True  'otwiera plik po wygenerowaniu
    
    Application.ScreenUpdating = True
    MsgBox "Raport PDF wygenerowany pomyœlnie!", vbInformation
    Exit Sub

BladZapisu:
    Application.ScreenUpdating = True
    MsgBox "B³¹d zapisu PDF!" & vbNewLine & _
           "SprawdŸ, czy nie masz otwartego starego pliku PDF o tej samej nazwie.", vbCritical
End Sub


'Resetuje aplikacje do stanu pocz¹tkowego (bez wykresów)
Sub ResetujAplikacje()

    Dim wsDash As Worksheet, wsCalc As Worksheet
    Dim shp As Shape
    Dim pt As PivotTable
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Set wsDash = ThisWorkbook.Sheets("DASHBOARD")
    
    'Zdejmuje ochronê (¿eby móc czyœciæ)
    On Error Resume Next
    wsDash.Unprotect Password:=""
    Set wsCalc = ThisWorkbook.Sheets("OBLICZENIA")
    On Error GoTo 0
    
    'Czyœci arkusz obliczenia
    If Not wsCalc Is Nothing Then
        wsCalc.Cells.Clear
    End If
    
    'Usuwa wykresy i KPI
    For Each shp In wsDash.Shapes
        If shp.Type = msoChart Then
            shp.Delete
        
        ElseIf Left(shp.Name, 4) = "KPI_" Then
            shp.Delete
            
        'nie usuwa slicerów
        End If
    Next shp
    
    ' Zabezpieczenie widoku spowrotem
    Call ZabezpieczWidok
    
    wsDash.Activate
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    
End Sub
