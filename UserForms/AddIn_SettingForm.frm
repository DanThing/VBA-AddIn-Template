VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddIn_SettingForm 
   Caption         =   "Add-In Settings"
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6045
   OleObjectBlob   =   "AddIn_SettingForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddIn_SettingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'===================================
' **Name**|AddIn_SettingForm
' **Type**|Userform
' **Purpose**|
' **Author**|Daniel Boyce
' **Version**|1.0.20180720
' Default Width of form is to be 314
' Fields not currently in use are
' hidden off to the right of the form
'===================================

'-----------------------------------
' Setting definitions
'
' Core
'    EnableLogging
'    EnableContextMenu
'    EnableSupersession
'    EnableRemoveRMUR
'    EnableAddItemcodeDashes
'    EnableExportThisWS
'    CompanyName
'
' Suplementry
'    URL1           | Kommunity URL
'    URL2           | Supersession folder
'    URL3           | Reman Menu folder
'    URL4           |
'    URL5           |
'    Text1          | Supersession Filename
'    Text2          | Reman Menu Filename
'    Text3          | Reman Menu File Version
'    Text4          | SOH / Inventory Folder
'    Text5          | BOM Folder
'    Text6          | Reports Folder
'    Text7          |
'    Text8          |
'    Text9          |
'    Text10         |
'    Boolean1       |
'    Boolean2       |
'    Boolean3       |
'    Boolean4       |
'    Boolean5       |


Option Explicit

Private Sub btn_cancel_Click()
    Me.Hide
End Sub


Private Sub btn_update_Click()
 
    With ThisWorkbook.Worksheets("Settings")
        .Range("CompanyName") = Me.companyNameUpdate.value
        
        .Range("URL1_") = Me.URL1_.value
        .Range("URL2_") = Me.URL2_.value
        .Range("URL3_") = Me.URL3_.value
        .Range("URL4_") = Me.URL4_.value
        .Range("URL5_") = Me.URL5_.value
        
        .Range("Text1") = Me.Text1_.value
        .Range("Text2") = Me.Text2_.value
        .Range("Text3") = Me.Text3_.value
        .Range("Text4") = Me.Text4_.value
        .Range("Text5") = Me.Text5_.value
        .Range("Text6") = Me.Text6_.value
        .Range("Text7") = Me.Text7_.value
        .Range("Text8") = Me.Text8_.value
        .Range("Text9") = Me.Text9_.value
        .Range("Text10") = Me.Text10_.value
        
        .Range("EnableLogging") = Me.opt_enablelogging.value
        .Range("EnableContextMenu") = Me.opt_enablecontextmenu.value
        .Range("EnableSupersession") = Me.opt_supersession.value
        .Range("EnableRemoveRMUR") = Me.opt_RemoveRMUR.value
        .Range("EnableAddItemcodeDashes") = Me.opt_AddItemcodeDashes.value
        .Range("EnableExportThisWS") = Me.opt_ExportThisWS.value
        
        .Range("Boolean1") = Me.Boolean1_.value
        .Range("Boolean2") = Me.Boolean2_.value
        .Range("Boolean3") = Me.Boolean3_.value
        .Range("Boolean4") = Me.Boolean4_.value
        .Range("Boolean5") = Me.Boolean5_.value
        
    End With
    Me.Hide
    ThisWorkbook.Save
    MsgBox "Settings saved." & vbNewLine & _
            "You may need to reopen this workbook for some settings to take effect.", vbOKOnly
End Sub

Private Sub btn_chkUpdates_Click()
    DisabledFunction
End Sub

'Private Sub btn_partskitsTEST_Click()
'Application.ScreenUpdating = False
'Dim wb As Workbook
'On Error Resume Next
'    Set wb = Workbooks.Open(Me.Text2_.value & FS & Me.Text4_.value & " " & Me.Text3_.value & ".xlsx")
'    If Err Then
'        MsgBox "Connection Failed.", vbOKOnly + vbInformation, Company
'    Else
'        MsgBox "Connection Success.", vbOKOnly + vbInformation, Company
'        wb.Close False
'    End If
'Application.ScreenUpdating = True
'End Sub

Private Sub btn_remanmenuTEST_Click()
Application.ScreenUpdating = False

Dim wb As Workbook
On Error Resume Next
    Set wb = Workbooks.Open(URL3_ & FS & Me.Text2_.value & Space(1) & Me.Text3_.value & ".xlsx")
    If Err Then
        MsgBox "Connection Failed.", vbOKOnly + vbInformation, Company
    Else
        MsgBox "Connection Success.", vbOKOnly + vbInformation, Company
        wb.Close False
    End If
Application.ScreenUpdating = True
End Sub

Private Sub btn_ssTEST_Click()
Application.ScreenUpdating = False
Dim wb As Workbook
On Error Resume Next
    Set wb = Workbooks.Open(Me.URL2_.value & "/" & Me.Text1_.value & ".xlsb")
    If Err Then
        MsgBox "Connection Failed.", vbOKOnly + vbInformation, Company
    Else
        MsgBox "Connection Success.", vbOKOnly + vbInformation, Company
        wb.Close False
    End If
Application.ScreenUpdating = True
End Sub


Private Sub companyLogo__DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    DisabledFunction
End Sub

Private Sub Text1__DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim tempstr As String
    tempstr = InputBox("Enter the Parts Super Session Tool filename", "Supersession Filename", ThisWorkbook.Worksheets("Settings").Range("Text1"))
    If tempstr = vbNullString Or tempstr = " " Then
        MsgBox "Unchanged or Invalid Entry.", vbOKOnly + vbExclamation
        Exit Sub
    End If
    If InStrRev(tempstr, ".") > 2 Then
        tempstr = Left(tempstr, Len(tempstr) - InStrRev(tempstr, "."))
    End If
    If Left(tempstr, 1) = "/" Then
        tempstr = Mid(tempstr, 2, Len(tempstr))
    End If
    Me.Text1_.Text = Trim(tempstr)
End Sub

Private Sub Text2__DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim tempstr As String
    tempstr = InputBox("Enter the Reman Menu filename", "Reman Menu Filename", ThisWorkbook.Worksheets("Settings").Range("Text2"))
    If tempstr = vbNullString Or tempstr = " " Then
        MsgBox "Unchanged or Invalid Entry.", vbOKOnly + vbExclamation
        Exit Sub
    End If
    If InStrRev(tempstr, ".") > 2 Then
        tempstr = Left(tempstr, Len(tempstr) - InStrRev(tempstr, "."))
    End If
    If Left(tempstr, 1) = "/" Then
        tempstr = Mid(tempstr, 2, Len(tempstr))
    End If
    If InStrRev(tempstr, "v") > 0 Then
        Me.Text3_.value = Right(tempstr, InStrRev(tempstr, "v"))
        tempstr = Left(tempstr, Len(tempstr) - InStrRev(tempstr, " v"))
    End If
    Me.Text2_.Text = Trim(tempstr)
End Sub

Private Sub Text3__DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim tempstr As String
    tempstr = InputBox("Enter the new Reman Menu version number [v#.##]", "Reman Menu Version", ThisWorkbook.Worksheets("Settings").Range("Text3"))
    If tempstr = vbNullString Or tempstr = " " Then
        MsgBox "Unchanged or Invalid Entry.", vbOKOnly + vbExclamation
        Exit Sub
    End If
    If Not Left(tempstr, 1) = "v" Then
        tempstr = "v" & tempstr
    End If
    If Right(tempstr, 3) = ".00" Then
        tempstr = Left(tempstr, 2)
    End If
    Me.Text3_.Text = Trim(tempstr)
End Sub

Private Sub Text4__DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim tempstr As String
    tempstr = getFolder("SOH/Inventory")
    If tempstr = vbNullString Or tempstr = " " Then
        MsgBox "Unchanged or Invalid Entry.", vbOKOnly + vbExclamation
        Exit Sub
    End If
    Me.Text4_.Text = Trim(tempstr)
End Sub

Private Sub Text5__DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim tempstr As String
    tempstr = getFolder("BOM Data")
    If tempstr = vbNullString Or tempstr = " " Then
        MsgBox "Unchanged or Invalid Entry.", vbOKOnly + vbExclamation
        Exit Sub
    End If
    Me.Text5_.Text = Trim(tempstr)
End Sub

Private Sub Text6__DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim tempstr As String
    tempstr = getFolder("Reports Folder")
    If tempstr = vbNullString Or tempstr = " " Then
        MsgBox "Unchanged or Invalid Entry.", vbOKOnly + vbExclamation
        Exit Sub
    End If
    Me.Text6_.Text = Trim(tempstr)
End Sub

Private Sub Text7__DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    DisabledFunction
End Sub

Private Sub Text8__DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 DisabledFunction
End Sub

Private Sub Text9__DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 DisabledFunction
End Sub

Private Sub Text10__DblClick(ByVal Cancel As MSForms.ReturnBoolean)
 DisabledFunction
End Sub

Private Sub URL1__DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim tempstr As String
    tempstr = InputBox("Enter the new Kommunity URL", "Kommunity", ThisWorkbook.Worksheets("Settings").Range("URL1_"))
    If tempstr = vbNullString Or tempstr = " " Then
        MsgBox "Unchanged or Invalid Entry.", vbOKOnly + vbExclamation
        Exit Sub
    End If
    If Not Right(tempstr, 1) = "/" Then
        tempstr = tempstr & "/"
    End If
    If Left(tempstr, 7) = "http://" Then
        tempstr = Mid(tempstr, 8, Len(tempstr))
    End If
    Me.URL1_.Text = Trim(tempstr)
End Sub

Private Sub URL2__DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim tempstr As String
    tempstr = InputBox("Enter the new Supersession Data URL", "Kommunity", ThisWorkbook.Worksheets("Settings").Range("URL2_"))
    If tempstr = vbNullString Or tempstr = " " Then
        MsgBox "Unchanged or Invalid Entry.", vbOKOnly + vbExclamation
        Exit Sub
    End If
    If Right(tempstr, 5) = ".xlsb" Then
        tempstr = Left(tempstr, Len(tempstr) - InStrRev(tempstr, "/"))
    End If
    If Not Right(tempstr, 1) = "/" Then
        tempstr = tempstr & "/"
    End If
    If Left(tempstr, 7) = "http://" Then
        tempstr = Mid(tempstr, 8, Len(tempstr))
    End If
    Me.URL2_.Text = Trim(tempstr)
End Sub

Private Sub URL3__DblClick(ByVal Cancel As MSForms.ReturnBoolean)
Dim tempstr As String
    tempstr = InputBox("Enter the new Supersession Data URL", "Kommunity", ThisWorkbook.Worksheets("Settings").Range("URL3_"))
    If tempstr = vbNullString Or tempstr = " " Then
        MsgBox "Unchanged or Invalid Entry.", vbOKOnly + vbExclamation
        Exit Sub
    End If
    If Right(tempstr, 5) = ".xlsx" Then
        tempstr = Left(tempstr, Len(tempstr) - InStrRev(tempstr, "/"))
    End If
    If Not Right(tempstr, 1) = "/" Then
        tempstr = tempstr & "/"
    End If
    If Left(tempstr, 7) = "http://" Then
        tempstr = Mid(tempstr, 8, Len(tempstr))
    End If
    Me.URL3_.Text = Trim(tempstr)
End Sub

Private Sub URL4__DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    DisabledFunction
End Sub

Private Sub URL5__DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    DisabledFunction
End Sub

Private Sub UserForm_Initialize()
    Me.Caption = "AddIn Setting - " & Company
    With ThisWorkbook.Worksheets("Settings")
        Me.opt_enablelogging.value = .Range("EnableLogging").value
        Me.opt_enablecontextmenu.value = .Range("EnableContextMenu").value
        Me.opt_supersession.value = .Range("EnableSupersession").value
        Me.opt_RemoveRMUR.value = .Range("EnableRemoveRMUR").value
        Me.opt_AddItemcodeDashes.value = .Range("EnableAddItemcodeDashes").value
        Me.opt_ExportThisWS.value = .Range("EnableExportThisWS").value
        
        Me.companyNameUpdate.value = .Range("CompanyName").value
        
        Me.URL1_.value = .Range("URL1_").value
        Me.URL2_.value = .Range("URL2_").value
        Me.URL3_.value = .Range("URL3_").value
        Me.URL4_.value = .Range("URL4_").value
        Me.URL5_.value = .Range("URL5_").value
        
        Me.Text1_.value = .Range("Text1").value
        Me.Text2_.value = .Range("Text2").value
        Me.Text3_.value = .Range("Text3").value
        Me.Text4_.value = .Range("Text4").value
        Me.Text5_.value = .Range("Text5").value
        Me.Text6_.value = .Range("Text6").value
        Me.Text7_.value = .Range("Text7").value
        Me.Text8_.value = .Range("Text8").value
        Me.Text9_.value = .Range("Text9").value
        Me.Text10_.value = .Range("Text10").value
        
        Me.Boolean1_.value = .Range("Boolean1").value
        Me.Boolean2_.value = .Range("Boolean2").value
        Me.Boolean3_.value = .Range("Boolean3").value
        Me.Boolean4_.value = .Range("Boolean4").value
        Me.Boolean5_.value = .Range("Boolean5").value
        
    End With
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Unload Me
End Sub
