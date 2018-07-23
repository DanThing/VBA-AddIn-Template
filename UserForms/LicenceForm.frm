VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LicenceForm 
   Caption         =   "AddIn Licence Information"
   ClientHeight    =   5010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5790
   OleObjectBlob   =   "LicenceForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LicenceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===================================
' **Name**|LicenceForm
' **Type**|Userform
' **Author**|Daniel Boyce
' **Version**|1.2.20180627
'-----------------------------------
'-LicenceForm
'   - UserForm_Initialize
'   - UserForm_QueryClose
'-----------------------------------
'===================================

Option Explicit

Private Sub UserForm_Initialize()
    With ThisWorkbook.Worksheets("Settings")
        Me.Lic_TextBox1.value = _
            .Range("MIT_Header").value & vbNewLine & _
            vbTab & .Range("MIT_copyright").value & vbNewLine & _
            vbNewLine & _
            .Range("MIT_licence1").value & vbNewLine & vbNewLine & _
            vbTab & Chr(149) & Space(1) & .Range("MIT_licence2").value & vbNewLine & _
            vbNewLine & _
            .Range("MIT_Licence3").value & vbNewLine
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Unload Me
End Sub
