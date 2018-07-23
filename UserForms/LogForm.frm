VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} LogForm 
   Caption         =   "Log"
   ClientHeight    =   6765
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   5040
   OleObjectBlob   =   "LogForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "LogForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'===================================
' **Name**|LogForm
' **Type**|Userform
' **Author**|Daniel Boyce
' **Version**|1.2.20180627
'-----------------------------------
'- LogForm
'   - btn_projectdump_Click
'   - cancelBtn_Click
'   - copyBtn_Click
'   - UserForm_Click
'   - UserForm_Initialize
'   - UserForm_QueryClose
'   - updateText
'-----------------------------------
'===================================

Option Explicit
Option Compare Text

Private Sub btn_projectdump_Click()

End Sub

Private Sub cancelBtn_Click()
    Unload Me
End Sub

Private Sub copyBtn_Click()
    With Me.errorText
       .SelStart = 0
       .SelLength = Len(.Text)
       .Copy
    End With
End Sub

Private Sub UserForm_Click()
    With Me.errorText
       .SelStart = 0
       .SelLength = Len(.Text)
    End With
End Sub

Private Sub UserForm_Initialize()
    Debug.Print "New Logger Form Created"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Unload Me
End Sub

Public Sub updateText(errorString As String)
    Me.errorText.value = errorText & errorString & vbNewLine
End Sub

