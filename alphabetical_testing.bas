Attribute VB_Name = "Module1"
Public Sub Workbook_Open()
    Dim Stock_sheet As Worksheet
    Dim v_Range As Range
    Dim v_Range_Keyword As Range
    Dim SheetPO_index As Long
    Dim i As Double
    Dim ClipboardDummy As String
    Dim boolean_Divide As Integer
    Dim Title_Column_Output
    Dim List_Stock As Variant
    Dim List_Ticker As Variant
    Dim List_Ticker_S() As String
    Dim b As String
    Dim Delim As String
    Dim strg As String
    Dim dict As Object
    Dim rng As Range, cell As Range, rng_date As Range
    Dim foundFirst As Boolean
    Dim nQuarter As Long
    Dim Quarterly As Double
    Dim nSum As Double
    
        SheetPO_index = 0
        sheetCount = ThisWorkbook.Sheets.Count

   Do
        SheetPO_index = SheetPO_index + 1
        Set Stock_sheet = Worksheets(SheetPO_index)
        Stock_sheet.Activate
        v_Stock_sheet_MaxRow = Stock_sheet.UsedRange.Rows.Count
        
        
        'Find Col Input Quarterly
        With ActiveSheet.Range("A1:Z1")
            Set ra = .Find(What:="Input Quarterly", LookIn:=xlValues, lookat:=xlPart, SearchDirection:=xlNext, MatchCase:=False)
        End With
        col_InputQuarterly = ra.Column + 1
        'Find Col <ticker>
        With ActiveSheet.Range("A1:Z1")
            Set ra = .Find(What:="<ticker>", LookIn:=xlValues, lookat:=xlPart, SearchDirection:=xlNext, MatchCase:=False)
        End With
        col_TickerInput = ra.Column
        'Find Col <date>
        With ActiveSheet.Range("A1:Z1")
            Set ra = .Find(What:="<date>", LookIn:=xlValues, lookat:=xlPart, SearchDirection:=xlNext, MatchCase:=False)
        End With
        col_DateInput = ra.Column
        
        v_InputQuarterly = Stock_sheet.Cells(1, col_InputQuarterly).Value
        If v_InputQuarterly = "1" Then
            month_Start = "202001"
            month_End = "202003"
        ElseIf v_InputQuarterly = "2" Then
            month_Start = "202004"
            month_End = "202006"
        ElseIf v_InputQuarterly = "3" Then
            month_Start = "202007"
            month_End = "202009"
        ElseIf v_InputQuarterly = "4" Then
            month_Start = "202010"
            month_End = "202012"
        End If
        
        j = 1
        col = 10
        nQ = 4
        i = 2
'=======================================================Title col=======================================================
        Title_Column_Output = Array("Ticker", "Quarterly change", " The percentage change", " The total stock volume of the stock", "Ticker", "Value", "Greatest % increase:Greatest % decrease:Greatest total volume")
        For k = 1 To (UBound(Title_Column_Output) + 1)
            If InStr(Title_Column_Output(k - 1), "Greatest % increase") Then
                col = col - 3
                For j = 0 To 2
                    Stock_sheet.Cells(2 + j, col).Value = Split(Title_Column_Output(k - 1), ":")(j)
                Next
            ElseIf InStr(Title_Column_Output(k - 1), "Ticker") And k <> 1 Then
                col = col + 3
                Stock_sheet.Cells(j, col).Value = Title_Column_Output(k - 1)
                col = col + 1
            Else
                Stock_sheet.Cells(j, col).Value = Title_Column_Output(k - 1)
                col = col + 1
            End If
        Next k
'=======================================================Get Arr Ticker =======================================================
        Set dict = CreateObject("Scripting.Dictionary")
        Set rng = Stock_sheet.Range("A1:B" & CStr(v_Stock_sheet_MaxRow))
        ' L?p qua t?ng ? trong ph?m vi
        On Error Resume Next
        For i = 2 To rng.Rows.Count
            If Not dict.exists(Stock_sheet.Cells(i, col_TickerInput).Value) Then
                dict.Add Stock_sheet.Cells(i, col_TickerInput).Value, Nothing
            End If
        On Error GoTo 0
        Next
'=============================================================================================================================
        nGreatest_increase = 0
        nGreatest_decrease = 0
        nGreatest_volum = 0
        col_Ticker = 10
        col_Quarterly = 11
        col_percentage = 12
        col_Sum = 13
        i = 2
        Dim firstKey As Long
        Dim lastKey As Long
'=======================================================Get Input Data =======================================================
        For Each Key In dict.keys
            Stock_sheet.Cells(i, col_Ticker).Value = Key ' Write Ticker
            foundFirst = False
            firstKey = rng.Find(Key, LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlNext).Row
            lastKey = rng.Find(Key, LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
            'Get Arr one Ticker
            For j = firstKey To lastKey
                If Stock_sheet.Cells(j, col_TickerInput).Value = Key And InStr(Stock_sheet.Cells(j, col_DateInput).Value, month_Start) Or Stock_sheet.Cells(j, col_TickerInput).Value = Key And InStr(Stock_sheet.Cells(j, col_DateInput).Value, month_End) Then
                    If Not foundFirst Then
                        firstPos_Ticker = j ' First Ticker quarter
                        foundFirst = True
                    End If
                    lastPos_Ticker = j ' Last Ticker quarter
                End If
            Next
            
            Set rng_Vol = Stock_sheet.Range("G" & CStr(firstPos_Ticker) & ":G" & CStr(lastPos_Ticker)) ' Range Ticker quarter
            nSum = WorksheetFunction.Sum(rng_Vol) ' Sum volum quarter
            
            
            nFirstValueInQuarter = firstPos_Ticker
            nLastValueInQuarter = lastPos_Ticker
            
            PriceOpen = CDbl(Stock_sheet.Cells(nFirstValueInQuarter, 3).Value) 'Price open quarter
            PriceClose = CDbl(Stock_sheet.Cells(nLastValueInQuarter, 6).Value) ' Price close quarter
            
            Quarterly = PriceClose - PriceOpen 'Quarterly change
            Stock_sheet.Cells(i, col_Quarterly).Value = Quarterly
            'Check Color
            If Quarterly < 0 Then
                Stock_sheet.Cells(i, col_Quarterly).Interior.Color = vbRed
            ElseIf Quarterly > 0 Then
                Stock_sheet.Cells(i, col_Quarterly).Interior.Color = vbGreen
            End If
            
            Stock_sheet.Cells(i, col_percentage).Value = Round((Quarterly * 100) / PriceOpen, 2) & "%" ' The percentage change
            ' Greatest % increase
            If nGreatest_increase < Round((Quarterly * 100) / PriceOpen, 2) Then
                nGreatest_increase = Round((Quarterly * 100) / PriceOpen, 2)
                Ticker_increase = Key
            End If
            If nGreatest_decrease > Round((Quarterly * 100) / PriceOpen, 2) Then
                nGreatest_decrease = Round((Quarterly * 100) / PriceOpen, 2)
                Ticker_decrease = Key
            End If
            Stock_sheet.Cells(i, col_Sum).Value = nSum
            If nGreatest_volum < nSum Then
                nGreatest_volum = nSum
                Ticker_volum = Key
            End If
            
            i = i + 1
        Next Key
        
        Stock_sheet.Cells(2, 17).Value = Ticker_increase
        Stock_sheet.Cells(2, 18).Value = nGreatest_increase & "%"
        Stock_sheet.Cells(3, 17).Value = Ticker_decrease
        Stock_sheet.Cells(3, 18).Value = nGreatest_decrease & "%"
        Stock_sheet.Cells(4, 17).Value = Ticker_volum
        Stock_sheet.Cells(4, 18).Value = nGreatest_volum

Loop Until SheetPO_index = sheetCount

End Sub





