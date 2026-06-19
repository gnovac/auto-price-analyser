VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arkusz2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo ErrorHandler
    
    Dim Cell As Range
    If Intersect(Target, Me.Range("A:B,F:F")) Is Nothing Then Exit Sub

    Application.EnableEvents = False

    For Each Cell In Target
        If Cell.Row >= 3 Then
            Select Case Cell.Column
                Case 1 ' Source Change Event
                    On Error Resume Next
                    Me.Cells(Cell.Row, "B").ClearContents
                    Call RecalculateRow(Cell.Row - 1)
                    If Err.Number <> 0 Then
                        Debug.Print "Source change error at row " & Cell.Row & ": " & Err.Description
                        Err.Clear
                    End If
                    On Error GoTo ErrorHandler
                    
                Case 2 ' Discount Change Event
                    On Error Resume Next
                    Call RecalculateRow(Cell.Row - 1)
                    If Err.Number <> 0 Then
                        Debug.Print "Discount change error at row " & Cell.Row & ": " & Err.Description
                        Err.Clear
                    End If
                    On Error GoTo ErrorHandler
                    
                Case 6 ' Manual Price Input Handling
                    On Error Resume Next
                    Me.Cells(Cell.Row, "A").Value = "MANUAL"
                    Me.Cells(Cell.Row, "B").ClearContents
                    Me.Cells(Cell.Row, "G").ClearContents
                    Call UpdatePriceSourceValidation(Cell.Row)
                    If Err.Number <> 0 Then
                        Debug.Print "Manual price error at row " & Cell.Row & ": " & Err.Description
                        Err.Clear
                    End If
                    On Error GoTo ErrorHandler
            End Select
        End If
    Next Cell

CleanExit:
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    Application.EnableEvents = True
    Call LogError("ZapytanieOfertowe.Worksheet_Change", Err.Number, Err.Description, False)
End Sub
