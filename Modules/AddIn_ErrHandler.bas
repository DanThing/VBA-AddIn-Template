Attribute VB_Name = "AddIn_ErrHandler"

'===================================
' **Name**|AddIn_ErrHandler
' **Type**|Module
' **Author**|Daniel Boyce
' **Version**|1.0.20180702
'-----------------------------------
'- AddIn_ErrHandler
'    - errHandler
'    - LogDebug
'    - LogWarning
'    - LogError
'    - ProjectDump
'-----------------------------------
'===================================

Option Explicit
Option Private Module


' Example of use
' ```
'Sub Example()
'    Application.ScreenUpdating = False
'    On Error GoTo ErrorExit
'
'    Dim x As Long
'    x = 1
'    Dim y As Long
'    y = 2
'
'    If x = y Then Err.Raise CustomError.Err1
'
'EOM:
'    Application.ScreenUpdating = True
'    Exit Sub
'
'ErrorExit:
'    ErrHandler Err
'    Resume EOM
'End Sub
' ```



Public Enum CustomError
    Err1 = vbObjectError + 2000
    Err2 = vbObjectError + 3000
    err3 = vbObjectError + 4000
    Err4 = vbObjectError + 5000
    Err5 = vbObjectError + 6000
    Err6 = vbObjectError + 7000
    Err7 = vbObjectError + 8000
    Err8 = vbObjectError + 9000
    Err9 = vbObjectError + 10000
    Err10 = vbObjectError + 11000
End Enum

Private DumpString As String

'===================================
' Customnised error handler.
'
' - @method errHandler
'   - @param {Object} Err
'   - @param {String} [errDetails=""]

Public Sub errHandler(Err As Object, Optional errDetails As String = vbNullString)
Dim errMsg As String
    Select Case Err.Number
        Case CustomError.Err1
            errMsg = "Company details not yet set." & vbNewLine & errDetails
        
        Case CustomError.Err2
            errMsg = "Unable to connect to URL." & vbNewLine & errDetails
            
        Case CustomError.err3
            errMsg = "Unable to open File/Folder." & vbNewLine & errDetails
        
        Case CustomError.Err4
            errMsg = "Invalid array or search." & vbNewLine & errDetails
        
        Case CustomError.Err5
            errMsg = "Selection invalid" & vbNewLine & errDetails
            
        Case CustomError.Err6
            errMsg = "Custom Error Message 6" & vbNewLine & errDetails
        
        Case CustomError.Err7
            errMsg = "Custom Error Message 7" & vbNewLine & errDetails
        
        Case CustomError.Err8
            errMsg = "Custom Error Message 8" & vbNewLine & errDetails
        
        Case CustomError.Err9
            errMsg = "Custom Error Message 9" & vbNewLine & errDetails
            
        Case CustomError.Err10
            errMsg = "Custom Error Message 10" & vbNewLine & errDetails
            
        Case Else
            errMsg = "Unexpected Error: [" & Err.Number & "] " & Err.Description & vbNewLine & errDetails ', vbCritical
    End Select
    If Not Logger Is Nothing Then
        LogError "[" & Err.Number & "] " & errMsg & " :: " & Err.Description
    Else
        MsgBox errMsg & " :: " & Err.Description
        SpeedUp False
        End
    End If
End Sub


'===================================
' Log message (when logging is enabled with `EnableLogging`)
' with optional location where the message is coming from.
'
' @example
' ```
' LogDebug "Executing request..."
' ' -> Log: Executing request...
'
' LogDebug "Executing request...", "Module.Function"
' ' -> Module.Function: Executing request...
' ```
'
' - @method LogDebug
'   - @param {String} Message
'   - @param {String} [From="Log"]

Public Sub LogDebug(ByVal Message As String, Optional From As String = "Log")
    If Not Logger Is Nothing Then
        From = "[" & Format(Now(), "hh:nn:ss") & "] " & From
        If getSetting("EnableLogging") Then
            Debug.Print From & ": " & Message
            DumpString = DumpString & From & ": " & Message & vbNewLine
        End If
        Logger.updateText From & ": " & Message
    End If
End Sub

'===================================
' Log warning (even when logging is disabled with `EnableLogging`)
' with optional location where the message is coming from.
'
' @example
' ```
' LogWarning "Something could go wrong"
' ' -> WARNING: Something could go wrong
'
' LogWarning "Something could go wrong", "Module.Function"
' ' -> WARNING for Module.Function: Something could go wrong
' ```
'
' - @method LogWarning
'   - @param {String} Message
'   - @param {String} [From=""]

Public Sub LogWarning(ByVal Message As String, Optional From As String = vbNullString)
    If From <> vbNullString Then
        From = " for " & From & ": "
    Else
        From = ": "
    End If
    Debug.Print "[" & Format(Now(), "hh:nn:ss") & "] WARNING" & From & Message
    If Not Logger Is Nothing Then
        Logger.updateText "[" & Format(Now(), "hh:nn:ss") & "] WARNING" & From & Message & vbNewLine
    End If
End Sub

'===================================
' Log error (even when logging is disabled with `EnableLogging`)
' with optional location where the message is coming from and error number.
'
' @example
' ```
' LogError "Something went wrong"
' ' -> ERROR: Something went wrong
'
' LogError "Something went wrong", "Module.Function"
' ' -> ERROR in Module.Function: Something went wrong
'
' LogError "Something went wrong", "Module.Function", 100
' ' -> ERROR in Module.Function: [100] Something went wrong
' ```
'
' - @method LogError
'   - @param {String} Message
'   - @param {String} From
'   - @param {Long} [ErrNumber=0]

Public Sub LogError(ByVal Message As String, Optional From As String = vbNullString, Optional ErrNumber As Long = 0)
    Dim web_ErrorValue As String
    If Not From = vbNullString Then
        From = " in " & From & ": "
    Else
        From = ": "
    End If
    If ErrNumber <> 0 Then
        web_ErrorValue = ErrNumber
        If ErrNumber < 0 Then
            web_ErrorValue = web_ErrorValue & " (" & (ErrNumber - vbObjectError) & " / " & VBA.LCase$(VBA.Hex$(ErrNumber)) & ")"
        End If
        web_ErrorValue = "[" & web_ErrorValue & "] "
    End If
    If Logger = Nothing Then
        Logger.updateText "[" & Format(Now(), "hh:nn:ss") & "] ERROR" & From & web_ErrorValue & Message & vbNewLine
    Else
        Debug.Print "[" & Format(Now(), "hh:nn:ss") & "] ERROR" & From & web_ErrorValue & Message
        DumpString = DumpString & "[" & Format(Now(), "hh:nn:ss") & "] ERROR" & From & web_ErrorValue & Message & vbNewLine
    End If
End Sub

'===================================
' Debugging function to get the details of the
' reference libraries, Modules, Proceedures and
' active AddIns in this project.
'
' - @method ProjectDump

Private Function ProjectDump() As String
Dim tempstring As String
Dim lr As Long, lc As Long
Dim i As Long
Dim VBProj As Object  'VBIDE.VBProject
Dim AddinComp As Object 'Office Addin
Dim VBComp As Object 'Modules
Dim procName As String
Dim procKind As Long
    Set VBProj = Application.VBE.ActiveVBProject
    On Error Resume Next
    ProjectDump = "--------Project Dump-----------" & vbNewLine
    ProjectDump = ProjectDump & "Microsoft Excel version " & _
        Application.version & " running on " & _
        Application.OperatingSystem & vbNewLine & "-----------------------------------" & vbNewLine
    ProjectDump = ProjectDump & "List of this Projects Library References" & vbNewLine
    For i = 1 To VBProj.References.Count
        With VBProj.References.Item(i)
            ProjectDump = ProjectDump & "-----------------------------------" & vbNewLine
            ProjectDump = ProjectDump & vbTab & "Description: " & .Description & vbNewLine & _
                        vbTab & "FullPath: " & .FullPath & vbNewLine & _
                        vbTab & "Major.Minor: " & .Major & "." & .Minor & vbNewLine & _
                        vbTab & "Name: " & .Name & vbNewLine & _
                        vbTab & "GUID: " & .GUID & vbNewLine & _
                        vbTab & "Type: " & .Type & vbNewLine
        End With 'VBProj.References.Item(i)
    Next i
    Set AddinComp = Application.AddIns
    ProjectDump = ProjectDump & vbNewLine & "List of currently active AddIns" & vbNewLine
    For i = 1 To AddinComp.Count
        With AddinComp.Item(i)
            If .IsOpen Then
                ProjectDump = ProjectDump & "-----------------------------------" & vbNewLine
                ProjectDump = ProjectDump & vbTab & "FullPath: " & .Path & vbNewLine & _
                            vbTab & "Name: " & .Name & vbNewLine & _
                            vbTab & "CLSID: " & .CLSID & vbNewLine & _
                            vbTab & "progID: " & .progID & vbNewLine
            End If
        End With 'AddinComp.Item(i)
    Next i
    ProjectDump = ProjectDump & "-----------------------------------" & vbNewLine
    ProjectDump = ProjectDump & vbNewLine & "List of Modules and their Proceedures" & vbNewLine
    For Each VBComp In VBProj.VBComponents
        ProjectDump = ProjectDump & "-----------------------------------" & vbNewLine
        With VBComp.CodeModule
            'The Procedures
            i = .CountOfDeclarationLines + 1
            ProjectDump = ProjectDump & .Name & vbNewLine
            Do While i < .CountOfLines
                procName = .ProcOfLine(i, procKind)
                ProjectDump = ProjectDump & vbTab & procName & vbNewLine
                i = .ProcStartLine(procName, procKind) + .ProcCountLines(procName, procKind) + 1
            Loop
        End With
    Next VBComp 'VBProj.VBComponents
    ProjectDump = ProjectDump & "-----------------------------------" & vbNewLine
    lr = Lastrow(ThisWorkbook.Worksheets("Settings"))
    lc = LastCol(ThisWorkbook.Worksheets("Settings"))
    ProjectDump = ProjectDump & vbNewLine & "Current Settings" & vbNewLine
    With ThisWorkbook.Worksheets("Settings")
        For i = 1 To lr
            tempstring = tempstring & _
                        .Cells(i, 1).value & " |" & .Cells(i, 2).value & " |" & .Cells(i, 3).value & " |" & .Cells(i, 4).value & vbNewLine
        Next i
    End With
    ProjectDump = ProjectDump & tempstring
    ProjectDump = ProjectDump & "----------END OF DUMP---------"
    If Logger Is Nothing Then
        Debug.Print ProjectDump
    End If
End Function


