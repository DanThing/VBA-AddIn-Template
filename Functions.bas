Attribute VB_Name = "kFunct"
Option Explicit

Public Sub Speedup(Optional doit As Boolean = True)
    If doit = True Then
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.DisplayAlerts = False
        Application.EnableEvents = False
        Set thisWB = ThisWorkbook
    Else
        Set thisWB = Nothing
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.StatusBar = False
        Application.DisplayAlerts = True
        Application.EnableEvents = True
    End If
End Sub

Public Sub raiseError(Optional errCode As String, Optional errDesc As String)
If errCode = vbNullString Then
    errCode = Err.Number
    errDesc = Err.Description
End If
Debug.Print errCode; errDesc
If MsgBox("Error:" & nl & errCode & nl & errDesc & nl & "Do you want to continue?", vbYesNo, "Komatsu") = vbNo Then
    Speedup (False)
    End
End If
End Sub

'This Fucntion takes a worksheet as an input and returns the last used row in the sheet
Function Lastrow(Sh As Worksheet)
    On Error Resume Next
    Lastrow = Sh.Cells.Find(What:="*", _
                            after:=Sh.Range("A1"), _
                            lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
    If Err Then raiseError "k404", "Match not found."
    On Error GoTo 0
End Function

'This Fucntion takes a worksheet as an input and returns the last used row in the sheet
Function LastCol(Sh As Worksheet)
    On Error Resume Next
    LastCol = Sh.Cells.Find(What:="*", _
                            after:=Sh.Range("A1"), _
                            lookat:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
    On Error GoTo 0
End Function

' shortcut to add in "
Function qt() As String
qt = Chr(34)
End Function

' shortcut to add in a new line
Function nl() As String
nl = Chr(10)
End Function

' shortcut to add in a number of spaces
Function sp(Optional a As Long = 1) As String
Do While Len(sp) <= a
    sp = sp & Chr(32)
Loop
End Function
