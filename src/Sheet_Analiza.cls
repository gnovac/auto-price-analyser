VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arkusz1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    
    Dim Cell As Range, rngB As Range, Area As Range
    Dim wsDOA As Worksheet
    
    ' Blockade by active filter
    If Me.AutoFilterMode Then
        If Me.AutoFilter.FilterMode Then
            MsgBox "Wy³¹cz filtr przed edycj¹ danych.", vbExclamation
            Exit Sub
        End If
    End If
    
    Application.EnableEvents = False
    
    ' 1. COLUMN B (PRODUCT CODE)
    Set rngB = Intersect(Target, Me.Columns("B"))
    If Not rngB Is Nothing Then
        
        Dim cellCount As Long
        cellCount = 0
        For Each Area In rngB.Areas
            For Each Cell In Area.Cells
                If Not IsEmpty(Cell) And Cell.Value <> "" Then cellCount = cellCount + 1
            Next Cell
        Next Area
        
        If cellCount = 0 Then
            ' Mass delete / Single delete handling
            For Each Area In rngB.Areas
                On Error Resume Next
                ' Clear contents across columns C to T (18 columns)
                Area.Offset(0, 1).Resize(, 18).ClearContents
                ' Reset background fill for columns C to S (17 columns)
                Area.Offset(0, 1).Resize(, 17).Interior.ColorIndex = xlNone
                ' Enforce soft blue background for column T (manual entry input)
                Area.Offset(0, 18).Interior.Color = RGB(230, 245, 255)
                On Error GoTo ErrorHandler
            Next Area
            
            Set wsDOA = ThisWorkbook.Sheets("Zapytanie Ofertowe")
            For Each Area In rngB.Areas
                If Area.Row >= 2 Then
                    On Error Resume Next
                    wsDOA.Range("A" & Area.Row + 1 & ":G" & Area.Row + Area.Rows.Count).ClearContents
                    wsDOA.Range("A" & Area.Row + 1 & ":A" & Area.Row + Area.Rows.Count).Validation.Delete
                    On Error GoTo ErrorHandler
                End If
            Next Area
            
        Else
            ' Single edits / paste operations
            For Each Cell In rngB
                If Cell.Row >= 2 Then
                    If Not IsEmpty(Cell) And Cell.Value <> "" Then
                        If VarType(Cell.Value) = vbString Then
                            On Error Resume Next
                            ' Clean spaces from input string
                            Cell.Value = Replace(WorksheetFunction.Trim(Cell.Value), " ", "")
                            On Error GoTo ErrorHandler
                        End If
                    End If
                End If
            Next Cell
            
            Call UpdateAnalysis
            
            Set wsDOA = ThisWorkbook.Sheets("Zapytanie Ofertowe")
            For Each Cell In rngB
                If Cell.Row >= 2 Then
                    On Error Resume Next
                    wsDOA.Cells(Cell.Row + 1, "A").ClearContents
                    wsDOA.Cells(Cell.Row + 1, "B").ClearContents
                    wsDOA.Cells(Cell.Row + 1, "F").ClearContents
                    wsDOA.Cells(Cell.Row + 1, "G").ClearContents
                    On Error GoTo ErrorHandler
                End If
            Next Cell
            
            For Each Cell In rngB
                If Cell.Row >= 2 Then
                    On Error Resume Next
                    Call RecalculateRow(Cell.Row)
                    If Err.Number <> 0 Then
                        Debug.Print "RecalculateRow error at row " & Cell.Row & ": " & Err.Description
                        Err.Clear
                    End If
                    On Error GoTo ErrorHandler
                End If
            Next Cell
        End If
    End If

    ' 2. GRADIENTS (Trigger formatting recalculation when prices change)
    If Not Intersect(Target, Union(Columns("E:F"), Columns("H:I"), Columns("M"), Columns("T"))) Is Nothing Then
        For Each Cell In Target
            If Cell.Row >= 2 Then
                On Error Resume Next
                ApplyStaticColors Me, Cell.Row
                On Error GoTo ErrorHandler
            End If
        Next Cell
    End If

    ' 3. QUANTITY (Sync order quantities to the RFQ sheet)
    If Not Intersect(Target, Me.Columns("U")) Is Nothing Then
        Set wsDOA = ThisWorkbook.Sheets("Zapytanie Ofertowe")
        For Each Cell In Target
            If Cell.Row >= 2 Then
                On Error Resume Next
                wsDOA.Cells(Cell.Row + 1, "E").Value = Cell.Value
                Call RecalculateRow(Cell.Row)
                On Error GoTo ErrorHandler
            End If
        Next Cell
    End If

CleanExit:
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    Call LogError("Analiza.Worksheet_Change", Err.Number, Err.Description, True)
End Sub
