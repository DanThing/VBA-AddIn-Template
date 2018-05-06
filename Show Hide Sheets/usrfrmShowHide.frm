VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} usrfrmShowHide 
   Caption         =   "Komatsu - Show/Hide Worksheet Tabs"
   ClientHeight    =   4650
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8025
   OleObjectBlob   =   "usrfrmShowHide.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "usrfrmShowHide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=====================================================
' Workbook Name: Komatsu Report Base Tools & Data.xlsm
'   Module: Main
' Code by Daniel Boyce
' Version: 1.2.20180504
'=====================================================
Option Explicit

Private Sub buildTOC()
Dim ws As Worksheet, rowCount As Long
Dim tip As String
rowCount = 17
With ThisWorkbook.Worksheets("Tools Page")
    .Range("B:C").Clear
    For Each ws In Worksheets
        tip = ""
        If Not ws.Name = "Tools Page" Then
            If ws.Visible = xlSheetHidden Then
                .Cells(rowCount, 2) = ws.Name
                .Cells(rowCount, 3) = "<- Click 'Show/Hide Worksheet Tabs'"
            Else
                .Hyperlinks.Add _
                    Anchor:=.Cells(rowCount, 2), _
                    Address:="", _
                    SubAddress:=ws.Name & "!A1", _
                    ScreenTip:=tip, _
                    TextToDisplay:=ws.Name
                .Cells(rowCount, 3) = tip
            End If
            rowCount = rowCount + 1
        End If
    Next ws
End With
End Sub

Private Sub showSheet(index As Long, sheetname As String)
    Application.ScreenUpdating = False
    With Me
        .notVisible.RemoveItem (index)
        .isVisible.AddItem sheetname
    End With
    ThisWorkbook.Worksheets(sheetname).Visible = xlSheetVisible
    Application.ScreenUpdating = True
End Sub

Private Sub hideSheet(index As Long, sheetname As String)
    Application.ScreenUpdating = False
    With Me
        .isVisible.RemoveItem (index)
        .notVisible.AddItem sheetname
    End With
    Worksheets(sheetname).Visible = xlSheetHidden
    Application.ScreenUpdating = True
End Sub

Private Sub btnHide1_Click()
Dim itm As Long, sht As String
    With Me.isVisible
        For itm = 0 To .ListCount
            If .Selected(itm) = True Then
                .Selected(itm) = False
                sht = .List(itm)
                hideSheet itm, sht
            End If
        Next itm
    End With
End Sub

Private Sub btnHideAll_Click()
Dim ws As Worksheet
    Application.ScreenUpdating = False
    Me.isVisible.Clear
    Me.notVisible.Clear
    For Each ws In Worksheets
        If Not ws.Name = "Tools Page" Then
            ws.Visible = xlSheetHidden
            Me.notVisible.AddItem ws.Name
        End If
    Next ws
    Application.ScreenUpdating = True
End Sub

Private Sub btnShow1_Click()
Dim itm As Long, sht As String
    With Me.notVisible
        For itm = 0 To .ListCount
            If .Selected(itm) = True Then
                .Selected(itm) = False
                sht = .List(itm)
                showSheet itm, sht
            End If
        Next itm
    End With
End Sub

Private Sub btnShowAll_Click()
Dim ws As Worksheet
    Application.ScreenUpdating = False
    Me.isVisible.Clear
    Me.notVisible.Clear
    For Each ws In Worksheets
        If Not ws.Name = "Tools Page" Then
            ws.Visible = xlSheetVisible
            Me.isVisible.AddItem ws.Name
        End If
    Next ws
    Application.ScreenUpdating = True
End Sub

Private Sub isVisible_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim itm As Long, sht As String
    With Me.isVisible
        For itm = 0 To .ListCount
            If .Selected(itm) = True Then
                .Selected(itm) = False
                sht = .List(itm)
                hideSheet itm, sht
                Exit Sub
            End If
        Next itm
    End With
End Sub

Private Sub notVisible_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim itm As Long, sht As String
    With Me.notVisible
        For itm = 0 To .ListCount
            If .Selected(itm) = True Then
                .Selected(itm) = False
                sht = .List(itm)
                showSheet itm, sht
                Exit Sub
            End If
        Next itm
    End With
End Sub

Private Sub UserForm_Initialize()
Dim ws As Worksheet
    For Each ws In Worksheets
        If Not ws.Name = "Tools Page" Then
            Select Case ws.Visible
                Case xlSheetVisible
                    Me.isVisible.AddItem ws.Name
                Case xlSheetHidden
                    Me.notVisible.AddItem ws.Name
                Case xlSheetVeryHidden
                    ' do nothing
                Case Else
                    ' should not be able to get here
                    Err.Raise 448 ' Named argument not found
            End Select
        End If
    Next ws
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Unload Me
End Sub
