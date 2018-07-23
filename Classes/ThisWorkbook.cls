VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'===================================
' **Name**|ThisWorkbook
' **Type**|Object
' **Purpose**|Custom settings for when this workbook is opened and closed. This sets up custom Context (Right click) Menus
' **Author**|Daniel Boyce
' **Version**|1.0.20180605
'-----------------------------------
'- ThisWorkbook
'    - Workbook_Open
'    - Workbook_BeforeClose
'    - AddToCellMenu
'    - DeleteFromCellMenu
'-----------------------------------
'===================================

Option Explicit

Private Sub Workbook_Open()
    If ThisWorkbook.Worksheets("Settings").Range("EnableContextMenu") Then
        AddToCellMenu
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
DeleteFromCellMenu
End Sub

Private Sub AddToCellMenu()
Dim ContextMenu As CommandBar
Dim MySubMenu As CommandBarControl
Dim WS As Worksheet
    Set WS = ThisWorkbook.Worksheets("Settings")
' Delete the controls first to avoid duplicates.
    DeleteFromCellMenu
' Set ContextMenu to the Cell context menu.
    Set ContextMenu = Application.CommandBars("Cell")
    ContextMenu.Controls(3).BeginGroup = True
' Add a custom submenu with three buttons.
    Set MySubMenu = ContextMenu.Controls.Add(Type:=msoControlPopup, before:=3)
    With MySubMenu
        .Caption = "AddIn Menu"
        .Tag = "customcontextmenu"

        With .Controls.Add(Type:=msoControlButton)
            If WS.Range("EnableExportThisWS") = True Then
                .OnAction = "'" & ThisWorkbook.Name & "'!" & "ExportThisWS"
                .FaceId = 142
            Else
                .OnAction = "'" & ThisWorkbook.Name & "'!" & "DisabledFunction"
                .FaceId = 1133
            End If
            .Caption = "Export Sheet to new Workbook"
        End With
    
    End With
' Add a separator to the Cell context menu.
    ContextMenu.Controls(3).BeginGroup = True
End Sub
'

Private Sub DeleteFromCellMenu()
Dim ContextMenu As CommandBar
Dim ctrl As CommandBarControl
' Set ContextMenu to the Cell context menu.
    Set ContextMenu = Application.CommandBars("Cell")
' Delete the custom controls with the Tag : My_Cell_Control_Tag.
    For Each ctrl In ContextMenu.Controls
        If ctrl.Tag = "customcontextmenu" Then
            ctrl.Delete
        End If
    Next ctrl
End Sub



