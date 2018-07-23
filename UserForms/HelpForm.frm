VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} HelpForm 
   Caption         =   "Add-In Help - Komatsu Austalia"
   ClientHeight    =   5595
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6090
   OleObjectBlob   =   "HelpForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "HelpForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===================================
' **Name**|HelpForm
' **Type**|Userform
' **Author**|Daniel Boyce
' **Version**|1.2.20180627
'-----------------------------------
'- HelpForm
'   - btn_Cancel_Click
'   - btn_New_Click
'   - HelpCombo1_Change
'   - UserForm_Initialize
'   - UserForm_QueryClose
'-----------------------------------
'===================================

Option Explicit

Private Sub btn_cancel_Click()
 Me.Hide
End Sub

Private Sub btn_New_Click()
    Dim frm As New CreateNewHelpForm
    frm.Edit Me.HelpCombo1.listIndex
End Sub

Private Sub HelpCombo1_Change()
    With ThisWorkbook.Worksheets("HelpContents")
        Me.HelpText1.Text = .Cells(HelpCombo1.listIndex + 1, 3).value
    End With
End Sub

Private Sub UserForm_Initialize()
Dim lr As Long
Dim i As Long

    lr = Lastrow(ThisWorkbook.Worksheets("HelpContents"))
    
    With ThisWorkbook.Worksheets("HelpContents")
        For i = 2 To lr
            Me.HelpCombo1.AddItem .Cells(lr, 2).value
        Next i
    End With

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
 Unload Me
End Sub

