VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SGMExtension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'''''''''''''''''''''''''''''''''''''''''''
' TTDXEdit Plug-in for Saved Game Manager '
' Version 1.20                            '
'                                         '
' Based on TTDEdit Plug-in for SGM        '
' Copyright © Owen Rudge 2001-2012        '
'''''''''''''''''''''''''''''''''''''''''''

Dim PID As Integer
Public Function EnableMenuItem() As Boolean
    On Error GoTo ExtensionError
    
    If FilePath = "" Then
        EnableMenuItem = False
    Else
        If EF.GetSelectedGamesCount(1) = 0 And EF.GetSelectedGamesCount(2) = 0 Then
            EnableMenuItem = False
        Else
            EnableMenuItem = True
        End If
    End If
    
    Exit Function
ExtensionError:
    Select Case EF.ErrorProc(Err)
        Case 1
            Resume
        Case 2
            Resume Next
        Case 3
            MsgBox "Abort!"
    End Select
End Function

Public Function GetHelpCaption() As String
    GetHelpCaption = "&TTDX Editor"
End Function


Public Property Let PluginID(NewVal As Integer)
    PID = NewVal
'    MsgBox "Plugin ID for Test Plugin: " & PID, vbInformation, "Test Plugin v2"
End Property
Public Property Get PluginID() As Integer
    PluginID = PID
End Property
Public Property Get ExpectedVersion() As Double
    ExpectedVersion = 2#
End Property
Public Function GetPluginInformation(InfType As Integer) As String
    If InfType = 1 Then ' Name
        GetPluginInformation = "TTDX Editor"
    ElseIf InfType = 2 Then ' Description
        GetPluginInformation = "Allows you to edit your games"
    End If
End Function

Public Sub AboutProc()
    On Error GoTo ExtensionError
    
    frmAbout.Show 1
    Exit Sub
ExtensionError: ' Call ErrorProc to display the default error dialog. This returns 1 for
    Select Case EF.ErrorProc(Err) ' retry, 2 for ignore and 3 for abort
        Case 1
            Resume
        Case 2
            Resume Next
        Case 3
            MsgBox "Abort!"
    End Select
End Sub

Public Sub ExtensionProc()
    On Error GoTo ExtensionError
    
    Dim DestDir As String, SrcDir As String
    Dim Games As Variant
    Dim i As Integer
    
    If FilePath = "" Then
        MsgBox "The TTDX Editor executable path is not set. Please start TTDX Editor and close it, then restart the Manager.", vbExclamation, "TTDX Editor"
        Exit Sub
    End If
    
    SrcDir = EF.GetSourceDir()
    DestDir = EF.GetDestDir()
    
    Games = EF.GetSelectedGames(1)
    
    For i = 1 To EF.GetSelectedGamesCount(1)
        Shell FilePath & " " & MakePath(SrcDir) & Games(i), 1
    Next i
    
    Games = EF.GetSelectedGames(2)
    
    For i = 1 To EF.GetSelectedGamesCount(2)
        Shell FilePath & " " & MakePath(DestDir) & Games(i), 1
    Next i
    
    Exit Sub
    
ExtensionError:
    Select Case EF.ErrorProc(Err)
        Case 1
            Resume
        Case 2
            Resume Next
        Case 3
            MsgBox "Abort!"
    End Select
End Sub

Public Function InitExtension() As String
'    On Error GoTo ExtensionError
    InitExtension = "&Edit Selected Game(s)..."
    
    Set EF = CreateObject("TTManager.ExtFunctions")
    
    GetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\Owen Rudge\TTDX Editor", "Path", FilePath
    
    If FilePath = "" Then
        GetKeyValue HKEY_CURRENT_USER, "SOFTWARE\Owen Rudge\TTDX Editor", "Path", FilePath
        
        If FilePath = "" Then
            GetKeyValue HKEY_LOCAL_MACHINE, "SOFTWARE\JVP\TTDXEdit", "Path", FilePath
        
            If FilePath = "" Then
                MsgBox "The TTDX Editor plug-in cannot find the TTDXEdit.exe path in the registry. Please start TTDX Editor then quit it, and restart the Manager.", vbExclamation, "TTDX Editor Plug-in"
                Exit Function
            End If
        End If
    End If

    If LCase$(Right$(FilePath, 4)) <> ".exe" Then
        FilePath = FilePath & ".exe"
    End If
    
    Exit Function
ExtensionError: ' You must use your own error handler here
    MsgBox "An error has occurred: " & Err & " (" & Error(Err) & ")", vbCritical, "TTDX Editor Plug-in"
End Function


Public Function TermExtension()
    On Error Resume Next
    Set EF = Nothing
End Function


