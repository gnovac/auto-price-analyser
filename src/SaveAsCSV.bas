Attribute VB_Name = "SaveAsCSV"
Option Explicit

Private Const SAVE_FOLDER As String = "C:\CSV"

Public Sub CSVForCSPS()
    Dim ws As Worksheet, arr As Variant, arrComments As Variant
    Dim lastRow As Long, r As Long
    Dim lines() As String, lineCount As Long
    Dim ts As String, filePath As String
    Dim fso As Object, tsOut As Object
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Set ws = ActiveSheet
    
    ' Find the last active data row intersecting columns A, B, or D
    lastRow = Application.WorksheetFunction.Max( _
        ws.Cells(ws.Rows.Count, "A").End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, "B").End(xlUp).Row, _
        ws.Cells(ws.Rows.Count, "D").End(xlUp).Row)
    If lastRow < 1 Then GoTo EndMacro
    
    ' Retrieve array contents from source blocks (including comments in D)
    arr = ws.Range("A1:B" & lastRow).Value2
    arrComments = ws.Range("D1:D" & lastRow).Value2
    
    ' Prepare output matrix constraints
    ReDim lines(1 To lastRow)
    
    For r = 1 To lastRow
        If Len(arr(r, 1)) > 0 Or Len(arr(r, 2)) > 0 Or Len(arrComments(r, 1)) > 0 Then
            lineCount = lineCount + 1
            If Len(arrComments(r, 1)) > 0 Then
                ' Processing structure containing comments: index;quantity;;comment;
                lines(lineCount) = arr(r, 1) & ";" & arr(r, 2) & ";;" & arrComments(r, 1) & ";"
            Else
                ' Processing standard structure without comments: index;quantity;;;
                lines(lineCount) = arr(r, 1) & ";" & arr(r, 2) & ";;;"
            End If
        End If
    Next r
    
    ' Execute FileSystemObject text write procedures if records exist
    If lineCount > 0 Then
        ts = Format(Now, "yyyy-mm-dd_hh-nn-ss")
        filePath = SAVE_FOLDER & "\CSPS_ORDER_" & ts & ".csv"
        
        Set fso = CreateObject("Scripting.FileSystemObject")
        If Not fso.FolderExists(SAVE_FOLDER) Then fso.CreateFolder SAVE_FOLDER
        
        Set tsOut = fso.CreateTextFile(filePath, True)
        tsOut.Write Join(Application.Index(lines, 1, 0), vbCrLf)
        tsOut.Close
        
        MsgBox "Zapisano: " & filePath, vbInformation
    End If

EndMacro:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

