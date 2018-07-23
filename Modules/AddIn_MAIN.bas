Attribute VB_Name = "AddIn_MAIN"
'===================================
' **Name**|AddIn_MAIN
' **Type**|Module
' **Purpose**|
' **Compatibilty**|As of verion 0.0.20180702 this Addin does not support the MAC OS.
' **Author**|Daniel Boyce
' **Version**|1.0.20180702
'-----------------------------------
'- AddIn_MAIN
'    - SpeedUp
'    - openSettings
'-----------------------------------
'===================================

Option Explicit
Option Private Module

Public AddInSettings As AddIn_SettingForm
Public EnableLogging As Boolean
Public Logger As LogForm


'===================================
' Sets the performance paramaters to run Macros
'
' - @method SpeedUp
'   - @param {Boolean} Toggle
'
' @example
'```
' SpeedUp True
'```

Public Sub SpeedUp(ByVal Toggle As Boolean)
    EnableLogging = getSetting("EnableLogging")
    With Application
        .ScreenUpdating = Not Toggle
        .EnableEvents = Not Toggle
        .DisplayAlerts = Not Toggle
        .Calculation = IIf(Toggle, xlCalculationManual, xlCalculationAutomatic)
        .StatusBar = Not Toggle
        .EnableAnimations = Not Toggle
        .DisplayStatusBar = Not Toggle
    End With
    If Toggle And EnableLogging Then
        Set Logger = New LogForm
        LogDebug "----Start of Log----"
    End If
End Sub


