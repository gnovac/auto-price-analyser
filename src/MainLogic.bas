Attribute VB_Name = "MainLogic"
Option Explicit

' ==========================================================================
' PART 1: HIGH PERFORMANCE DATA CACHE & ENGINE
' ==========================================================================
Private pDictPage1 As Object, pDictKoszyk As Object
Private pDictUzyskane As Object, pDictAfter As Object
Private pDictPodhurt As Object, pDictDead As Object
Private pDictZapas As Object

Sub UpdateAnalysis()
    On Error GoTo ErrorHandler
    
    Dim wsDest As Worksheet
    Dim lastRow As Long, lastUsedRow As Long, i As Long
    Dim arrIndex As Variant, arrOut As Variant, key As String
    
    Set wsDest = ThisWorkbook.Sheets("Analiza")
    Dim screenUpdateState As Boolean, calcState As XlCalculation, eventsState As Boolean
    screenUpdateState = Application.ScreenUpdating
    calcState = Application.Calculation
    eventsState = Application.EnableEvents
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Load Memory Cache Dictionaries
    If pDictPage1 Is Nothing Then
        Set pDictPage1 = LoadDataToDict("Page1", "A", Array("B", "C", "D", "E", "F", "G"))
        Set pDictKoszyk = LoadDataToDict("KOSZYK", "A", Array("C"))
        Set pDictUzyskane = LoadDataToDict("Uzyskane ceny", "A", Array("B", "C"))
        Set pDictAfter = LoadDataToDict("After", "A", Array("B", "C"))
        Set pDictPodhurt = LoadDataToDict("PH", "A", Array("B"))
        Set pDictDead = LoadDataToDict("DEADSTOCK", "A", Array("B", "C", "D"))
        Set pDictZapas = LoadDataToDict("Zapas 3mce<", "A", Array("B", "C"))
    End If
    
    ' Determine dataset row boundaries
    lastRow = wsDest.Cells(wsDest.Rows.Count, "B").End(xlUp).Row
    lastUsedRow = wsDest.UsedRange.Row + wsDest.UsedRange.Rows.Count - 1
    If lastUsedRow < lastRow Then lastUsedRow = lastRow
    
    If lastRow < 2 Then
        ' Clear data range C:S if no indices are provided
        wsDest.Range("C2:S" & lastUsedRow + 1).ClearContents
        wsDest.Range("C2:S" & lastUsedRow + 1).Interior.ColorIndex = xlNone
        wsDest.Range("C2:S" & lastUsedRow + 1).FormatConditions.Delete
        GoTo CleanExit
    End If
    
    If lastRow = 2 Then
        ReDim arrIndex(1 To 1, 1 To 1)
        arrIndex(1, 1) = wsDest.Range("B2").Value2
    Else
        arrIndex = wsDest.Range("B2:B" & lastRow).Value2
    End If
    
    ' Array Matching Engine (Output matrix dimensioned to 17 target columns)
    ReDim arrOut(1 To UBound(arrIndex, 1), 1 To 17)
    Dim tmpArr As Variant
    For i = 1 To UBound(arrIndex, 1)
        key = CStr(arrIndex(i, 1))
        key = Replace(Trim(key), " ", "")
        If Len(key) > 0 Then
            If pDictPage1.Exists(key) Then
                tmpArr = pDictPage1(key)
                arrOut(i, 1) = tmpArr(0): arrOut(i, 2) = tmpArr(1): arrOut(i, 3) = tmpArr(2)
                ' Former Column Q (element 15) is shifted to Column S (element 17)
                arrOut(i, 4) = tmpArr(3): arrOut(i, 5) = tmpArr(4): arrOut(i, 17) = tmpArr(5)
            Else
                arrOut(i, 1) = "#N/D"
            End If
            
            If pDictKoszyk.Exists(key) Then arrOut(i, 6) = pDictKoszyk(key)(0)
            
            If pDictUzyskane.Exists(key) Then
                tmpArr = pDictUzyskane(key)
                arrOut(i, 7) = tmpArr(0): arrOut(i, 8) = tmpArr(1)
            End If
            
            If pDictAfter.Exists(key) Then
                tmpArr = pDictAfter(key)
                arrOut(i, 9) = tmpArr(0): arrOut(i, 10) = tmpArr(1)
            End If
            
            If pDictPodhurt.Exists(key) Then arrOut(i, 11) = pDictPodhurt(key)(0)
            
            ' NEW DATA MAP: Positioned between Sub-Wholesale (Col M) and Deadstock (Col P)
            If pDictZapas.Exists(key) Then
                tmpArr = pDictZapas(key)
                arrOut(i, 12) = tmpArr(0) ' Column N
                arrOut(i, 13) = tmpArr(1) ' Column O
            End If
            
            ' SHIFTED DEADSTOCK MAP: Adjusted index starts at 14 instead of 12
            If pDictDead.Exists(key) Then
                tmpArr = pDictDead(key)
                arrOut(i, 14) = tmpArr(0): arrOut(i, 15) = tmpArr(1): arrOut(i, 16) = tmpArr(2)
            End If
        Else
            Dim c As Long: For c = 1 To 17: arrOut(i, c) = Empty: Next c
        End If
    Next i
    
    ' Bulk memory dump into Excel worksheet range
    wsDest.Range("C2").Resize(UBound(arrOut, 1), 17).Value2 = arrOut
    
    If lastRow < lastUsedRow Then
        ' Clear excess rows down to last used cell
        wsDest.Range("C" & lastRow + 1 & ":S" & lastUsedRow).ClearContents
        wsDest.Range("C" & lastRow + 1 & ":S" & lastUsedRow).FormatConditions.Delete
        wsDest.Range("C" & lastRow + 1 & ":S" & lastUsedRow).Interior.ColorIndex = xlNone
    End If
    
    wsDest.Range("J2:J" & lastRow).NumberFormat = "dd.mm.yyyy"
    wsDest.Range("L2:L" & lastRow).NumberFormat = "dd.mm.yyyy"
    
    ' Apply Formatting Engines (Custom Static Gradients)
    Dim r As Long
    For r = 2 To lastRow
        On Error Resume Next
        Call ApplyStaticColors(wsDest, r)
        If Err.Number <> 0 Then
            Debug.Print "Gradient error at row " & r & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo ErrorHandler
    Next r

CleanExit:
    Application.ScreenUpdating = screenUpdateState
    Application.Calculation = calcState
    Application.EnableEvents = eventsState
    Exit Sub

ErrorHandler:
    Call LogError("UpdateAnalysis", Err.Number, Err.Description, True)
    Resume CleanExit
End Sub

' ==========================================================================
' PART 2: BUSINESS LOGIC (RFQ & PRICING)
' ==========================================================================

Sub RecalculateRow(srcRow As Long)
    On Error GoTo ErrorHandler
    
    Dim wsDOA As Worksheet, wsAnalysis As Worksheet
    Dim dstRow As Long
    Dim indexVal As Variant, sourceChoice As String, discount As Double
    Dim priceStock As Variant, priceBasket As Variant, priceDOA As Variant, priceWholesale As Variant
    Dim priceAfter As Variant, productName As Variant, qtyValue As Variant
    Dim dateDOA As Variant, dateAfter As Variant
    Dim minPrice As Double, lowestSource As String

    Set wsDOA = ThisWorkbook.Sheets("Zapytanie Ofertowe")
    Set wsAnalysis = ThisWorkbook.Sheets("Analiza")

    If srcRow < 2 Then Exit Sub
    dstRow = srcRow + 1

    indexVal = wsAnalysis.Cells(srcRow, "B").Value
    
    If IsEmpty(indexVal) Or indexVal = "" Then
        wsDOA.Range("A" & dstRow & ":G" & dstRow).ClearContents
        On Error Resume Next
        wsDOA.Cells(dstRow, "A").Validation.Delete
        On Error GoTo ErrorHandler
        Exit Sub
    End If

    productName = wsAnalysis.Cells(srcRow, "C").Value
    ' Shifted layout: Quantity tracking column shifted to Column U
    qtyValue = wsAnalysis.Cells(srcRow, "U").Value

    wsDOA.Cells(dstRow, "C").Value = indexVal
    wsDOA.Cells(dstRow, "D").Value = productName
    wsDOA.Cells(dstRow, "E").Value = qtyValue

    If Not IsNumeric(wsDOA.Cells(dstRow, "B").Value) Then
        discount = 0
        wsDOA.Cells(dstRow, "B").Value = discount
    Else
        discount = CDbl(wsDOA.Cells(dstRow, "B").Value)
    End If

    sourceChoice = UCase(Trim(wsDOA.Cells(dstRow, "A").Value & ""))

    ' Baseline extraction matrix (Pricing arrays located prior to shifting columns)
    priceStock = wsAnalysis.Cells(srcRow, "E").Value
    priceBasket = wsAnalysis.Cells(srcRow, "H").Value
    priceDOA = wsAnalysis.Cells(srcRow, "I").Value
    dateDOA = wsAnalysis.Cells(srcRow, "J").Value
    priceAfter = wsAnalysis.Cells(srcRow, "K").Value
    dateAfter = wsAnalysis.Cells(srcRow, "L").Value
    priceWholesale = wsAnalysis.Cells(srcRow, "M").Value

    ' Safe numeric data conversion filters
    If Not IsNumeric(priceStock) Or IsError(priceStock) Then priceStock = 0
     If Not IsNumeric(priceBasket) Or IsError(priceBasket) Then priceBasket = 0
    If Not IsNumeric(priceDOA) Or IsError(priceDOA) Then priceDOA = 0
    If Not IsNumeric(priceAfter) Or IsError(priceAfter) Then priceAfter = 0
    If Not IsNumeric(priceWholesale) Or IsError(priceWholesale) Then priceWholesale = 0

    On Error Resume Next
    UpdatePriceSourceValidation dstRow
    If Err.Number <> 0 Then
        Debug.Print "Validation error at row " & dstRow & ": " & Err.Description
        Err.Clear
    End If
    On Error GoTo ErrorHandler
    
    If sourceChoice = "MANUAL" Then
        If Not IsNumeric(wsDOA.Cells(dstRow, "F").Value) Or wsDOA.Cells(dstRow, "F").Value = 0 Then
            sourceChoice = ""
        Else
            Exit Sub
        End If
    End If

    If sourceChoice = "" Or sourceChoice = "AUTO" Then
        minPrice = GetLowestPrice(priceStock, priceBasket, priceDOA, priceWholesale, priceAfter)
        
        Select Case minPrice
            Case priceDOA: lowestSource = "WYNEGOCJOWANA"
            Case priceAfter: lowestSource = "BULK"
            Case priceStock: lowestSource = "STOCK"
            Case priceBasket: lowestSource = "KAMPANIA"
            Case priceWholesale: lowestSource = "PODHURT"
            Case Else: lowestSource = ""
        End Select
        
        If lowestSource <> "" Then wsDOA.Cells(dstRow, "A").Value = lowestSource

        InsertPrice wsDOA, dstRow, lowestSource, discount, _
                    priceStock, priceBasket, priceDOA, priceWholesale, _
                    priceAfter, dateDOA, dateAfter
    Else
        InsertPrice wsDOA, dstRow, sourceChoice, discount, _
                    priceStock, priceBasket, priceDOA, priceWholesale, _
                    priceAfter, dateDOA, dateAfter
    End If
    
    Exit Sub

CleanExit:
    Exit Sub
ErrorHandler:
    Call LogError("RecalculateRow (row " & srcRow & ")", Err.Number, Err.Description, False)
End Sub

Private Sub InsertPrice(wsDOA As Worksheet, rowIndex As Long, source As String, discount As Double, _
                        priceStock As Variant, priceBasket As Variant, priceDOA As Variant, priceWholesale As Variant, _
                        priceAfter As Variant, dateDOA As Variant, dateAfter As Variant)
    On Error Resume Next
    
    Select Case UCase(source)
        Case "WYNEGOCJOWANA"
            wsDOA.Cells(rowIndex, "F").Value = IIf(discount > 0, Round(priceDOA * (1 - discount), 2), priceDOA)
            If discount = 0 Then wsDOA.Cells(rowIndex, "G").Value = dateDOA Else wsDOA.Cells(rowIndex, "G").ClearContents
        Case "BULK"
            wsDOA.Cells(rowIndex, "F").Value = IIf(discount > 0, Round(priceAfter * (1 - discount), 2), priceAfter)
            If discount = 0 Then wsDOA.Cells(rowIndex, "G").Value = dateAfter Else wsDOA.Cells(rowIndex, "G").ClearContents
        Case "STOCK"
            wsDOA.Cells(rowIndex, "F").Value = Round(priceStock * (1 - discount), 2)
            wsDOA.Cells(rowIndex, "G").ClearContents
        Case "KAMPANIA"
            wsDOA.Cells(rowIndex, "F").Value = Round(priceBasket * (1 - discount), 2)
            wsDOA.Cells(rowIndex, "G").ClearContents
        Case "PODHURT"
            wsDOA.Cells(rowIndex, "F").Value = Round(priceWholesale * (1 - discount), 2)
            wsDOA.Cells(rowIndex, "G").ClearContents
    End Select
    
    On Error GoTo 0
End Sub

Function GetLowestPrice(ParamArray prices()) As Double
    On Error Resume Next
    
    Dim minV As Double, i As Long
    minV = 1E+20
    For i = LBound(prices) To UBound(prices)
        If IsNumeric(prices(i)) And prices(i) > 0 Then
            If prices(i) < minV Then minV = prices(i)
        End If
    Next i
    If minV = 1E+20 Then minV = 0
    GetLowestPrice = minV
    
    On Error GoTo 0
End Function

Sub UpdatePriceSourceValidation(rowIndex As Long)
    On Error GoTo ErrorHandler
    
    Dim wsAnalysis As Worksheet, wsDOA As Worksheet
    Dim srcRow As Long, listItems As String
    Dim priceStock As Variant, priceBasket As Variant, priceDOA As Variant, priceWholesale As Variant, priceAfter As Variant
    Dim currentPrice As Variant, minPrice As Double

    Set wsDOA = ThisWorkbook.Sheets("Zapytanie Ofertowe")
    Set wsAnalysis = ThisWorkbook.Sheets("Analiza")
    If rowIndex < 3 Then Exit Sub
    srcRow = rowIndex - 1

    priceStock = wsAnalysis.Cells(srcRow, "E").Value
    priceBasket = wsAnalysis.Cells(srcRow, "H").Value
    priceDOA = wsAnalysis.Cells(srcRow, "I").Value
    priceAfter = wsAnalysis.Cells(srcRow, "K").Value
    priceWholesale = wsAnalysis.Cells(srcRow, "M").Value

    If Not IsError(priceStock) And IsNumeric(priceStock) And priceStock > 0 Then listItems = listItems & "STOCK,"
    If Not IsError(priceBasket) And IsNumeric(priceBasket) And priceBasket > 0 Then listItems = listItems & "KAMPANIA,"
    If Not IsError(priceWholesale) And IsNumeric(priceWholesale) And priceWholesale > 0 Then listItems = listItems & "PODHURT,"
    If Not IsError(priceDOA) And IsNumeric(priceDOA) And priceDOA > 0 Then listItems = listItems & "WYNEGOCJOWANA,"
    If Not IsError(priceAfter) And IsNumeric(priceAfter) And priceAfter > 0 Then listItems = listItems & "BULK,"

    currentPrice = wsDOA.Cells(rowIndex, "F").Value
    minPrice = GetLowestPrice(priceStock, priceBasket, priceDOA, priceWholesale, priceAfter)

    If Not IsError(currentPrice) And IsNumeric(currentPrice) And Abs(CDbl(currentPrice) - minPrice) > 0.01 Then
        listItems = listItems & "MANUAL,"
    End If
    If Right(listItems, 1) = "," Then listItems = Left(listItems, Len(listItems) - 1)

    On Error Resume Next
    If listItems <> "" And Len(listItems) <= 255 Then
        With wsDOA.Cells(rowIndex, "A").Validation
            .Delete
            .Add Type:=xlValidateList, Formula1:=listItems
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
    Else
        wsDOA.Cells(rowIndex, "A").Validation.Delete
    End If
    
    Exit Sub

ErrorHandler:
    Call LogError("UpdatePriceSourceValidation (row " & rowIndex & ")", Err.Number, Err.Description, False)
End Sub

Sub ResetAllPricesToLowest()
    On Error GoTo ErrorHandler
    
    Dim wsDOA As Worksheet, lastRow As Long, r As Long
    Set wsDOA = ThisWorkbook.Sheets("Zapytanie Ofertowe")
    lastRow = wsDOA.Cells(wsDOA.Rows.Count, "A").End(xlUp).Row
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    For r = 3 To lastRow
        wsDOA.Cells(r, "A").ClearContents
        wsDOA.Cells(r, "B").ClearContents
        On Error Resume Next
        RecalculateRow r - 1
        If Err.Number <> 0 Then
            Debug.Print "Reset error at row " & r & ": " & Err.Description
            Err.Clear
        End If
        On Error GoTo ErrorHandler
    Next r
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Ceny zresetowane do minimalnych z cennika.", vbInformation
    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Call LogError("ResetAllPricesToLowest", Err.Number, Err.Description, True)
End Sub

Sub ResetMemory()
    Set pDictPage1 = Nothing: Set pDictKoszyk = Nothing
    Set pDictUzyskane = Nothing: Set pDictAfter = Nothing
    Set pDictPodhurt = Nothing: Set pDictDead = Nothing
    Set pDictZapas = Nothing ' Memory release for the storage cache dictionary
End Sub

Function LoadDataToDict(sheetName As String, keyCol As String, itemCols As Variant) As Object
    On Error GoTo ErrorHandler
    
    Dim dict As Object, ws As Worksheet, arrData As Variant
    Dim r As Long, i As Long, key As String, items() As Variant, colIndices() As Long
    
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        Call LogError("LoadDataToDict", 0, "Sheet not found: " & sheetName, False)
        Set LoadDataToDict = dict
        Exit Function
    End If
    
    Dim lastR As Long: lastR = ws.Cells(ws.Rows.Count, keyCol).End(xlUp).Row
    If lastR < 2 Then Set LoadDataToDict = dict: Exit Function
    
    arrData = ws.Range("A1:Z" & lastR).Value2
    ReDim colIndices(LBound(itemCols) To UBound(itemCols))
    For i = LBound(itemCols) To UBound(itemCols)
        colIndices(i) = ws.Cells(1, itemCols(i)).Column
    Next i
    Dim keyColIdx As Long: keyColIdx = ws.Cells(1, keyCol).Column
    
    For r = 2 To UBound(arrData, 1)
        key = CStr(arrData(r, keyColIdx))
        key = Replace(Trim(key), " ", "")
        If Len(key) > 0 Then
            If Not dict.Exists(key) Then
                ReDim items(LBound(itemCols) To UBound(itemCols))
                For i = LBound(itemCols) To UBound(itemCols)
                    items(i) = arrData(r, colIndices(i))
                Next i
                dict.Add key, items
            End If
        End If
    Next r
    Set LoadDataToDict = dict
    Exit Function

ErrorHandler:
    Call LogError("LoadDataToDict (" & sheetName & ")", Err.Number, Err.Description, False)
    Set LoadDataToDict = dict
End Function

Sub ApplyStaticColors(wsSheet As Worksheet, targetRow As Long)
    On Error Resume Next
    Dim rngTarget As Range, currentCell As Range
    Dim valMin As Double, valMax As Double, valCurrent As Double
    Dim ratio As Double
    Dim redColor As Integer, greenColor As Integer
    
    ' SAFETY CONDITION: If column B has no index parameter (row cleared/purged)
    If wsSheet.Cells(targetRow, "B").Value = "" Then
        Dim rngPrices As Range
        ' Target column arrays automatically generated via Power Query processes
        Set rngPrices = Union(wsSheet.Range("E" & targetRow & ":F" & targetRow), _
                              wsSheet.Range("H" & targetRow & ":I" & targetRow), _
                              wsSheet.Range("M" & targetRow))
        
        ' Full structural clear for analytical pricing cells (removing backgrounds)
        rngPrices.FormatConditions.Delete
        rngPrices.Interior.ColorIndex = xlNone
        rngPrices.Font.Color = RGB(0, 0, 0)
        rngPrices.Font.Bold = False
        
        ' Clear styling and enforce UI soft-blue background template on manual entry node (Col T)
        With wsSheet.Cells(targetRow, "T")
            .FormatConditions.Delete
            .Interior.Color = RGB(230, 245, 255)
            .Font.Color = RGB(0, 0, 0)
            .Font.Bold = False
        End With
        Exit Sub
    End If
    
    ' --- Standard Execution Engine (Evaluates when active record mapping exists) ---
    Set rngTarget = Union(wsSheet.Range("E" & targetRow & ":F" & targetRow), _
                          wsSheet.Range("H" & targetRow & ":I" & targetRow), _
                          wsSheet.Range("M" & targetRow), _
                          wsSheet.Range("T" & targetRow))
    
    ' Clear old conditional formatting and standard colors
    rngTarget.FormatConditions.Delete
    rngTarget.Interior.ColorIndex = xlNone
    
    valMin = Application.WorksheetFunction.Min(rngTarget)
    valMax = Application.WorksheetFunction.Max(rngTarget)
    
    If Err.Number <> 0 Or IsEmpty(valMin) Then GoTo CleanExit
    
    For Each currentCell In rngTarget
        If Not IsError(currentCell) And Not IsEmpty(currentCell) And IsNumeric(currentCell.Value) Then
            valCurrent = currentCell.Value
            
            ' Calculate 3-color scale elements statically to bypass rendering limits
            If valMin = valMax Then
                currentCell.Interior.Color = RGB(165, 220, 165) ' ERP System Green Theme
                currentCell.Font.Color = RGB(0, 80, 0)
            Else
                ratio = (valCurrent - valMin) / (valMax - valMin)
                If ratio <= 0.5 Then
                    redColor = Int(255 * (ratio / 0.5))
                    greenColor = 130 + Int((255 - 130) * (ratio / 0.5))
                Else
                    redColor = 255
                    greenColor = 255 - Int(255 * ((ratio - 0.5) / 0.5))
                End If
                currentCell.Interior.Color = RGB(redColor, greenColor, 0)
            End If
            
            ' Highlight the absolute minimum row parameter value
            If valCurrent = valMin Then
                currentCell.Font.Bold = True
                currentCell.Font.Color = RGB(255, 255, 0)
            Else
                currentCell.Font.Bold = False
                currentCell.Font.Color = RGB(0, 0, 0)
            End If
        End If
    Next currentCell

CleanExit:
    On Error GoTo 0
End Sub

Sub LoadMemoryOnStart()
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    If pDictPage1 Is Nothing Then
        Set pDictPage1 = LoadDataToDict("Page1", "A", Array("B", "C", "D", "E", "F", "G"))
        Set pDictKoszyk = LoadDataToDict("KOSZYK", "A", Array("C"))
        Set pDictUzyskane = LoadDataToDict("Uzyskane ceny", "A", Array("B", "C"))
        Set pDictAfter = LoadDataToDict("After", "A", Array("B", "C"))
        Set pDictPodhurt = LoadDataToDict("PH", "A", Array("B"))
        Set pDictDead = LoadDataToDict("DEADSTOCK", "A", Array("B", "C", "D"))
        Set pDictZapas = LoadDataToDict("Zapas 3mce<", "A", Array("B", "C"))
    End If
    
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Call LogError("LoadMemoryOnStart", Err.Number, Err.Description, True)
End Sub

