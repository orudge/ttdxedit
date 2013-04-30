VERSION 5.00
Begin VB.Form frmSGM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Saved Game Manager Plug-in"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6495
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   4575
      Begin VB.Label labC 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4575
      Begin VB.Label LabSt2 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1440
         Width           =   4335
      End
      Begin VB.Label labSt 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.CommandButton cmdDo 
      Caption         =   "&Setup"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   2400
      Width           =   1575
   End
End
Attribute VB_Name = "frmSGM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private SGMPath As String, DllPath As String, DllPath2 As String, MajorVer As Double
Private F As New FileSystemObject

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDo_Click()
    If RunningWin9x() = True Then
        RegisterSGMPlugin MajorVer, DllPath, SGMPath
    Else
        Screen.MousePointer = 11
        DoEvents
        
        StartElevated Me.hwnd, """" & MakePath(App.Path) & App.EXEName & ".exe""", "/SGM", App.Path, 0, "In order to install the Saved Game Manager plug-in, you need to be running as an administrator. If you press Yes, you'll be prompted to enter an Administrator password. If this fails, please try logging out and running TTDX Editor as an administrator." & vbCrLf & vbCrLf & "Do you want to proceed?"
        Screen.MousePointer = 0
    End If

    Update
End Sub

Private Sub Form_Load()
    labC.Caption = "Transport Tycoon Saved Game Manager and the plug-in module are: © Owen Rudge 2000-2012. All rights reserved."
    Update
End Sub

Private Sub Update()
    Dim Wsa As String, wFl As Boolean, SGMVersion As String, Pos As Integer
    labSt.Caption = "": wFl = True
    SGMPath = fReadValue("HKLM", "Software\Owen Rudge\InstalledSoftware\TTSGM", "Path", "S", "")
    DllPath = F.BuildPath(App.Path, "SGMPlugIn\TTDXEdit.dll")
    
    If F.FileExists(DllPath) Then
        labSt.Caption = labSt.Caption + "Plugin file found." + Chr(13) + Chr(13)
    Else
        labSt.Caption = labSt.Caption + "Can't find plugin file (TTDXEdit.dll)." + Chr(13) + Chr(13)
        wFl = False: DllPath = ""
    End If
    
    SGMVersion = fReadValue("HKLM", "Software\Owen Rudge\InstalledSoftware\TTSGM", "Version", "S", "")
    
    On Error Resume Next
    Pos = InStr(3, SGMVersion, ".")
    MajorVer = CDbl(Left(SGMVersion, Pos - 1))
    On Error GoTo 0
    
    If MajorVer <= 0 Then
        labSt.Caption = labSt.Caption + "Saved Game Manager not found."
        wFl = False
    ElseIf MajorVer < 2.3 Then
        If F.FileExists(F.BuildPath(SGMPath, "plugins2.ini")) Then
            labSt.Caption = labSt.Caption + "Saved Game Manager v" + SGMVersion + " found in:" + Chr(13) + SGMPath
            Wsa = fReadValue(F.BuildPath(SGMPath, "plugins2.ini"), "TTDXEdit", "Filename", "S", "")
            If Wsa = "" Then
                LabSt2.Caption = "Plug-in is not installed"
            ElseIf Wsa = DllPath Then
                LabSt2.Caption = "Plug-in is installed."
            Else
                LabSt2.Caption = "Plug-in is not properly installed."
            End If
        Else
            labSt.Caption = labSt.Caption + "Saved Game Manager not found."
            wFl = False
        End If
    Else
        labSt.Caption = labSt.Caption + "Saved Game Manager v" + SGMVersion + " found in:" + Chr(13) + SGMPath
        Wsa = fReadValue("HKLM", "Software\Owen Rudge\Transport Tycoon Saved Game Manager\Plugins\TTDXEdit", "Filename", "S", "")
        If Wsa = "" Then
            LabSt2.Caption = "Plug-in is not installed"
        ElseIf Wsa = DllPath Then
            LabSt2.Caption = "Plug-in is installed."
        Else
            LabSt2.Caption = "Plug-in is not properly installed."
        End If
    End If
    
    cmdDo.Enabled = wFl
    
    If cmdDo.Enabled = True Then
        If IsElevated() = False Then
            SendMessage cmdDo.hwnd, BCM_SETSHIELD, 0, &HFFFFFFFF
        End If
    End If
End Sub

