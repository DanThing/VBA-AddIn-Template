Attribute VB_Name = "AddIn_Functions"
'===================================
' **Name**|AddIn_Functions
' **Type**|Module
' **Purpose**|Container Module to hold customised Functions
' **Useage**|See each function for details
' **Arthor**|Daniel Boyce
' **Version**| 1.0.20180720
'-----------------------------------
' - AddIn_Functions
'    - OperationCompleted
'    - OperationCancelled
'    - ExportThisWS
'    - DiffBetween
'    - WSExists
'    - IsWorkBookOpen
'    - openWB
'    - getFile
'    - getFolder
'    - Lastrow
'    - LastCol
'    - RndUp
'    - ClearNameRngs
'    - FS
'    - getArrayHeader
'    - DisabledFunction
'-----------------------------------
'===================================

Option Explicit
Option Compare Text
Option Private Module

Sub OperationCompleted()
MsgBox "Operation has been Completed.", vbOKOnly, Company
End Sub

Sub OperationCancelled()
MsgBox "Operation has been cancelled by User.", vbOKOnly, Company
SpeedUp False
End
End Sub

'===================================
' Exports a copy of the active worksheet without formulas
'
' - @method ExportThisWS
'===================================
Sub ExportThisWS()
    If ThisWorkbook.Worksheets("Settings").Range("EnableExportThisWS") = False Then
        DisabledFunction
        Exit Sub
    End If
    With ActiveSheet
        .Copy
        .Cells().Copy
        .Cells().PasteSpecial xlPasteValues
    End With
End Sub

'===================================
' Returns the difference between two numbers
'
' - @method DiffBetween
'   - @param {Variant} firstValue
'   - @param {Variant} secondValue
' - @returns {Long}
'===================================
Function DiffBetween(ByVal firstValue As Variant, ByVal secondValue As Variant) As Long
Dim fst As Long, snd As Long
With Application.WorksheetFunction
    fst = .Max(firstValue, secondValue)
    snd = .Min(firstValue, secondValue)
End With
DiffBetween = fst - snd
End Function

'===================================
' Returns True/False if a  worksheet exists in a given workbook
'
' - @method WSExists
'   - @param {Workbook} TargetWB
'   - @param {String} WSName
' - @returns {Boolean}

Public Function WSExists(ByRef TargetWB As Workbook, ByVal WSName As String) As Boolean
Dim WSO As Worksheet
    On Error GoTo errExit
    Set WSO = TargetWB.Sheets(WSName)
    If Not WSO Is Nothing Then WSExists = True
errExit:
    Set WSO = Nothing
End Function

'===================================
' Returns True/False if a given workbook is currently open
'
' - @method IsWorkBookOpen
'   - @param {String} WorkbookName
' - @returns {Boolean}

Function IsWorkBookOpen(ByVal WorkbookName As String) As Boolean
Dim WBO As Workbook
    On Error GoTo errExit
    Set WBO = Workbooks(WorkbookName)
    If Not WBO Is Nothing Then IsWorkBookOpen = True
    Set WBO = Nothing
    Exit Function
errExit:
    Set WBO = Nothing
End Function

'===================================
' This function checks to see if the workbook is already open.
' If it is then it uses that workbook otherwise it opens the workbook.
'
' - @method openWB
'   - @param {String} fname
' - @return {Workbook}

Function openWB(fname As String) As Workbook
    If IsWorkBookOpen(fname) Then
       Set openWB = Workbooks(fname)
       Exit Function
    Else
        Set openWB = Workbooks.Open(fname, False, True)
        Exit Function
    End If
    Err.Raise CustomError.err3, "openWB", "Cannot find file " & fname & "."
End Function

'===================================
' Shows a dialog box to get the location and file name for the indictaed file.
' If cancel selected then this returns vbNullString
'
' -@method getFile
'   - @param {String} fname
' - @return {String}

Function getFile(fname As String) As String
Dim intChoice As Integer
' User to select file to open
With Application.FileDialog(msoFileDialogOpen)
    .Top = Me.Parent.Application.Top + 100
    .Left = Me.Parent.Application.Left + 100
    .AllowMultiSelect = False
    .Title = fname & " - " & Company
    intChoice = .Show
    If intChoice <> 0 Then
        getFile = .SelectedItems(1)
        Exit Function
    End If
End With
errExit:
getFile = vbNullString
End Function

'===================================
' Shows a dialog box to get the path and name for the indicated folder.
' If cancel selected then this returns vbNullString
'
' - @method getFolder
'   - @param {String} fname
' - @return {String}

Function getFolder(fname As String) As String
Dim intChoice As Integer
With Application.FileDialog(msoFileDialogFolderPicker)
    .AllowMultiSelect = False
    .Title = fname & " - " & Company
    intChoice = .Show
    If intChoice <> 0 Then
        getFolder = .SelectedItems(1)
        Exit Function
    End If
End With
getFolder = vbNullString
End Function

'===================================
' Shows a dialog box to get the path and name for the indicated folder.
' If cancel selected then this returns vbNullString
'
' - @method getFolder
'   - @param {String} fname
' - @return {String}

Sub openFolder(fname As String)
    On Error GoTo errExit
    If fname = vbNullString Then
        Err.Raise CustomError.err3, "openFolder", "Cannot find folder " & fname
    Else
        Call Shell("explorer.exe " & fname, vbNormalFocus)
        Exit Sub
    End If
errExit:
    errHandler Err
End Sub

'===================================
' This Function takes a worksheet as an input
' and returns the last used row in the sheet
'
' - @method Lastrow
'   - @param {Worksheet} sh
' - @return {Long}

Function Lastrow(Sh As Worksheet)
    Lastrow = Sh.Cells.Find(What:="*", _
                            after:=Sh.Range("A1"), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
End Function

'===================================
' This Function takes a worksheet as an input
' and returns the last used column in the sheet
'
' - @method Lastcol
'   - @param {Worksheet} sh
' - @return {Long}

Function LastCol(Sh As Worksheet)
    LastCol = Sh.Cells.Find(What:="*", _
                            after:=Sh.Range("A1"), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
End Function

'===================================
' This Function takes a Number or range as an input
' and rounds it up to the next integer
'
' - @method RndUp
'   - @param {Variant} numbervalue
' - @return {Long}

Function RndUp(numbervalue As Variant) As Long
With Application.WorksheetFunction
    RndUp = .RoundUp(numbervalue, 0)
End With
End Function

'===================================
' clears the Named Ranged for the indicated worksheet.
'
' - @method ClearNameRngs
'   - @param {Worksheet} ws

Sub ClearNameRngs(WS As Worksheet)
Dim xName As Name
For Each xName In thisWB.Names
    If InStr(1, xName, WS.Name) Then xName.Delete
Next xName
End Sub

'===================================
' Returns the default path seperator.
'   '\' for Windows systems
'   ':' for Classic Mac OS
'   '/' for Unix
'
' - @method FS
' - @returns {string}

Function FS() As String
  FS = Application.PathSeparator
End Function

'===================================
' Function to get the position of a
' value within a 2-dimentional Array.
'
' - @method getArrayHeader
'   - @param {String} lookfor
'   - @param {Variant} inArray
' - @returns {Long}

Function getArrayHeader(lookfor As String, inArray As Variant) As Long
On Error GoTo errExit
Dim itm As Long
    For itm = LBound(inArray(2)) To UBound(inArray(2))
        If inArray(1, itm) = lookfor Then
            getArrayHeader = itm
            Exit Function
        End If
    Next itm
errExit:
    getArrayHeader = 0
    Err.Raise CustomError.Err4, "getArrayHeader"
End Function

'===================================
' Function to alert a user that a selected
' function is not yet implemented
'
' - @method DisabledFunction

Sub DisabledFunction()
    MsgBox "This function is currently disabled.", vbOKOnly, Company
    End
End Sub

