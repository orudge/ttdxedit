VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About TTDX Editor Plug-in for Saved Game Manager"
   ClientHeight    =   2205
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   5550
   ClipControls    =   0   'False
   Icon            =   "about.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2205
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Copyright © Jens Vang Petersen 2002. All Rights Reserved."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   6
      Top             =   960
      Width           =   4335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © Owen Rudge 2002-2012. All Rights Reserved."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1080
      TabIndex        =   5
      Top             =   1200
      Width           =   4260
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "TTDX Editor:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Copyright © Owen Rudge 2001-2012. All Rights Reserved."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   4260
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "Version 1.20"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "TTDX Editor Plug-in for Saved Game Manager"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3795
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim Version As String
    
    Left = (Screen.Width - Width) / 2
    Top = (Screen.Height - Height) / 2
    
    Command1.Left = (ScaleWidth - Command1.Width) / 2
    
    GetKeyValue HKEY_LOCAL_MACHINE, "Software\Owen Rudge\InstalledSoftware\TTDX Editor", "Version", Version
    
    If Version = "" Then
        GetKeyValue HKEY_LOCAL_MACHINE, "Software\Owen Rudge\TTDX Editor", "Version", Version
        
        If Version = "" Then
            GetKeyValue HKEY_LOCAL_MACHINE, "Software\JVP\TTDXedit", "Version", Version
        
            If Version = "" Then
                Exit Sub
            End If
        End If
    End If
    
    lblVersion.Caption = lblVersion.Caption & " for TTDX Editor " & Version
End Sub


