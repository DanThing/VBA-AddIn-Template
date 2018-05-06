Attribute VB_Name = "modFunctions"
Option Explicit


Public Function FileLocation(filename As String) As String
' Shows a dialog box to get the location and file name for the indictaed file.
 ' If cancel selected the this returns vbNullString
    Dim intChoice As Integer    
        If MsgBox( _        
            "Can you show me where the " & filename & " file is located?", _        
            vbOKCancel, _        
            "...") = vbOK Then        
        ' User to select file to open        
            Application.FileDialog(msoFileDialogOpen).AllowMultiSelect = False        
            intChoice = Application.FileDialog(msoFileDialogOpen).Show        
            If intChoice <> 0 Then            
                    FileLocation = Application.FileDialog( _                        
                                    msoFileDialogOpen).SelectedItems(1)            
                    Exit Function        
            End If    
        End If    
        FileLocation = vbNullString
End Function
    
    
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
If MsgBox("Error:" & nl & errCode & nl & errDesc & nl & "Do you want to continue?", vbYesNo) = vbNo Then
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

Function FS() As String  
    FS = Application.PathSeperator
End Function

Sub ClearNameRngs(ws As Worksheet)
' clears the Named Ranged for the indicated worksheet.
Dim xName As Name
    For Each xName In thisWB.Names    
        If InStr(1, xName, ws.Name) Then xName.Delete
    Next xName
End Sub
