Sub WorksheetLoop()

         Dim WS_Count As Integer
         Dim i As Integer

         ' Set WS_Count equal to the number of worksheets in the active
         ' workbook.
         WS_Count = ActiveWorkbook.Worksheets.Count

         ' Begin the loop.
         For i = 1 To WS_Count

            ' Insert your code here.
            ActiveWorkbook.Worksheets(i).Activate
            Main
            MsgBox ActiveWorkbook.Worksheets(i).Name

         Next i

End Sub
Sub Main()
    PrintHeaders
    Dim myArrayOfTickerNames() As String
    GetTickerNames myArrayOfTickerNames
    PrintTickerNames myArrayOfTickerNames
    Dim ticker As String
    Dim i As Integer
    Dim startPoint As Long
    startPoint = 2
    For i = 0 To UBound(myArrayOfTickerNames)
        ticker = myArrayOfTickerNames(i)
        Call Get_QChange_PChange_Volume(ticker, i, startPoint)
        Next i
    Dim length As Integer
    length = UBound(myArrayOfTickerNames)
    Conditional_Formatting (length)
    GetMaxP (length)
    GetMinP (length)
    GetMaxV (length)
    MsgBox startPoint
End Sub
Sub PrintHeaders()
    Cells(1, "K").Value = "Ticker"
    Cells(1, "L").Value = "Quarterly Change"
    Cells(1, "M").Value = "Percent Change"
    Cells(1, "N").Value() = "Volume"
    Cells(1, "Q").Value() = "Ticker"
    Cells(1, "R").Value() = "Value"
    Cells(2, "P").Value() = "Greatest% Increase"
    Cells(3, "P").Value() = "Greatest% Decrease"
    Cells(4, "P").Value() = "Greatest Total Volume"
End Sub

Sub PrintTickerNames(ByRef myArr() As String)
    Dim i As Integer
    For i = 1 To UBound(myArr) + 1
        Cells(i + 1, "K").Value = myArr(i - 1)
        Next i
End Sub

Sub GetTickerNames(ByRef myArray() As String)
    num_of_tickers = 0
    ReDim Preserve myArray(num_of_tickers)
    myArray(0) = Cells(2, "A").Value()
    Dim LR As Long
    LR = Cells(Rows.Count, 1).End(xlUp).Row
    For counter = 2 To LR
        matchResult = Application.Match(Cells(counter, "A").Value, myArray, 0)
        If Not IsError(matchResult) Then
            Cells(counter, "NO").Value = num_of_tickers
        Else
            num_of_tickers = num_of_tickers + 1
            ReDim Preserve myArray(num_of_tickers)
            myArray(num_of_tickers) = Cells(counter, "A").Value()
        End If
        Next counter
End Sub
Sub Get_QChange_PChange_Volume(ticker As String, i As Integer, startPoint As Long)
    Dim myArray2(0) As String
    Dim opener As Double
    Dim high As Double
    Dim volume As Double
    Dim LR As Long
    LR = Cells(Rows.Count, 1).End(xlUp).Row
    myArray2(0) = ticker
    opener = 0
    high = 0
    For counter = startPoint To LR + 1
        matchResult = Application.Match(Cells(counter, "A").Value, myArray2, 0)
        If Not IsError(matchResult) Then
            volume = volume + Cells(counter, "G").Value()
            If opener = 0 Then
                opener = Cells(counter, "C").Value()
            End If
        ElseIf opener = 0 Then
            'Cells(8, "HJ").Value() = "Don't Look At This"
        Else
            If high = 0 Then
                high = Cells(counter - 1, "D").Value()
                startPoint = counter - 10
            End If
            Exit For
        End If
        Next counter
    Cells(i + 2, "L").Value() = high - opener
    Cells(i + 2, "M").Value() = (high - opener) / opener
    Cells(i + 2, "N").Value() = volume
End Sub
Sub Conditional_Formatting(length As Integer)
'
' Conditional_Formatting Macro
'

'
    Dim rangeRowOne As Integer
    
    rangeRowOne = 2
    length = length + 2
    Set myRange = Range("L" & rangeRowOne & ":" & "L" & length)
    myRange.Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub

Sub GetMaxP(length As Integer)
    Dim rangeRowOne As Integer
    Dim maxVal As Double
    Dim maxCell As Range
    
    rangeRowOne = 2
    length = length + 2
    Set myRangeOfValues = Range("M" & rangeRowOne & ":" & "M" & length)
    Set myRangeOfNames = Range("K" & rangeRowOne & ":" & "K" & length)
    
    maxVal = Application.WorksheetFunction.Max(myRangeOfValues)
    Set maxCell = myRangeOfValues.Find(maxVal)
    For i = 1 To myRangeOfValues.Rows.Count
        If myRangeOfValues.Cells(i, 1).Value = maxVal Then
            Set maxCell = myRangeOfValues.Cells(i, 1)
            Exit For
        End If
    Next i
    MaxCellName = myRangeOfNames.Cells(maxCell.Row - 1, 1).Value()
    
    Cells(2, "R").Value() = maxVal
    Cells(2, "Q").Value() = MaxCellName
    
End Sub
Sub GetMinP(length As Integer)
    Dim rangeRowOne As Integer
    Dim minVal As Double
    Dim minCell As Range
    
    rangeRowOne = 2
    length = length + 2
    Set myRangeOfValues = Range("M" & rangeRowOne & ":" & "M" & length)
    Set myRangeOfNames = Range("K" & rangeRowOne & ":" & "K" & length)
    
    minVal = Application.WorksheetFunction.Min(myRangeOfValues)
    Set minCell = myRangeOfValues.Find(minVal)
    For i = 1 To myRangeOfValues.Rows.Count
        If myRangeOfValues.Cells(i, 1).Value = minVal Then
            Set minCell = myRangeOfValues.Cells(i, 1)
            Exit For
        End If
    Next i
    MinCellName = myRangeOfNames.Cells(minCell.Row - 1, 1).Value()
    
    Cells(3, "R").Value() = minVal
    Cells(3, "Q").Value() = MinCellName
    
End Sub
Sub GetMaxV(length As Integer)
    Dim rangeRowOne As Integer
    Dim maxVal As Double
    Dim maxCell As Range
    
    rangeRowOne = 2
    length = length + 2
    Set myRangeOfValues = Range("N" & rangeRowOne & ":" & "N" & length)
    Set myRangeOfNames = Range("K" & rangeRowOne & ":" & "K" & length)
    
    maxVal = Application.WorksheetFunction.Max(myRangeOfValues)
    Set maxCell = myRangeOfValues.Find(maxVal)
    For i = 1 To myRangeOfValues.Rows.Count
        If myRangeOfValues.Cells(i, 1).Value = maxVal Then
            Set maxCell = myRangeOfValues.Cells(i, 1)
            Exit For
        End If
    Next i
    MaxCellName = myRangeOfNames.Cells(maxCell.Row - 1, 1).Value()
    
    Cells(4, "R").Value() = maxVal
    Cells(4, "Q").Value() = MaxCellName
    
End Sub