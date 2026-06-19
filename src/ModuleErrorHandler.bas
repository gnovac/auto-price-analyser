Attribute VB_Name = "ModuleErrorHandler"
Option Explicit

' Errors Handler with log
Sub LogError(procedureName As String, errNumber As Long, errDescription As String, Optional criticalError As Boolean = False)
    Dim errorMsg As String
    
    ' Msg formating
    errorMsg = Now & " | " & procedureName & " | Error " & errNumber & ": " & errDescription
    
    ' Error message for user
    If criticalError Then
        MsgBox "B³¹d krytyczny w: " & procedureName & vbCrLf & vbCrLf & _
               "Error " & errNumber & ": " & errDescription & vbCrLf & vbCrLf & _
               "Skontaktuj siê z administratorem jeœli problem bêdzie siê powtarza³.", vbCritical, "B³¹d Systemu"
    End If
    
    ' Debug
    Debug.Print errorMsg
End Sub


