Attribute VB_Name = "TTDXEditor"
Option Explicit
Public Const RegBaseKey As String = "Software\Owen Rudge\TTDX Editor"

Public fAutoMode As Boolean
Public fFastMode As Boolean

Private F As New FileSystemObject

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Declare Function SHGetFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwReserved As Long, ppidl As Long) As Long
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListW" (ByVal pIDL As Long, ByVal pszPath As Long) As Long
Declare Function SHParseDisplayName Lib "shell32.dll" (ByVal pszName As Long, ByVal pbc As Long, ByRef ppidl As Long, ByVal sfgaoIn As Long, ByRef psfgaoOut As Long) As Long
Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenW" (ByVal lpString As Long) As Long
Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryW" (ByVal pszPath As Long) As Long
Declare Sub ILFree Lib "shell32.dll" Alias "#155" (ByVal pIDL As Long)

Public Const CSIDL_DESKTOP = &H0
Public Const MAX_PATH = 260

Sub Main()
    Dim Wsa As String, Wsb As String, Wdo As Integer
    
    InitCommonControls
    
    Wsa = Trim(Command$): Wdo = 0
    fAutoMode = False
    If Wsa > " " Then
        If Left(UCase(Wsa), 3) = "/FT" Then
            frmFileTypes.AssociateCmdLine
            End
        ElseIf Wsa = "/SGM" Then
            RegisterSGMPluginStartup
            End
        End If
        
        If InStr(UCase(Wsa), "/SU") Then
            Wdo = 1: Wsa = Replace(Wsa, "/su", "", 1, -1, vbTextCompare)
            fAutoMode = True
        ElseIf InStr(UCase(Wsa), "/S") Then
            Wdo = 2: Wsa = Replace(Wsa, "/s", "", 1, -1, vbTextCompare)
            fAutoMode = True
        End If
        Wsb = Wsa
    End If
    Wsa = Trim(Replace(Wsa, Chr(34), " "))
    frmMDI.Show
    DoEvents
    If F.FileExists(Wsa) Then
        frmMDI.CallFileLoad Wsa
        If Wdo = 1 Then
            frmMDI.CallFileSave 1
            Unload frmMDI
        ElseIf Wdo = 2 Then
            frmMDI.CallFileSave 0
            Unload frmMDI
        End If
    End If
End Sub
