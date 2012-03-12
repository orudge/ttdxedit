Attribute VB_Name = "TTDXEditor"
Option Explicit
Public Const RegBaseKey As String = "Software\Owen Rudge\TTDX Editor"

Public fAutoMode As Boolean
Public fFastMode As Boolean

Private F As New FileSystemObject

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
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
