VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmFileTypes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Associations"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5535
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optTOver 
      Caption         =   "Overwrite"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3900
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.OptionButton optTplug 
      Caption         =   """Plug-In"""
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3900
      TabIndex        =   3
      Top             =   1440
      Value           =   -1  'True
      Width           =   1095
   End
   Begin MSComctlLib.ListView lvTypes 
      Height          =   3255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Filetype"
         Object.Width           =   1677
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Assigned to:"
         Object.Width           =   4498
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
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
      Left            =   3840
      TabIndex        =   1
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton cmdDo 
      Caption         =   "&Associate"
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
      Left            =   3840
      TabIndex        =   0
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Select what to do with filetypes associated with programs other than TTDX Editor:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3840
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmFileTypes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private F As New FileSystemObject

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDo_Click()
    Dim Wa As Integer, wDta As ListItem, Wsa As String
    '
    ' Set The Master Values
    '
    SetAMaster "TTDXEdit.Save", 1, False
    SetAMaster "TTDXEdit.Scenario", 2, False
    SetAMaster "TTDXEdit.Unpack", 3, False
    '
    ' Loop through the filetypes
    '
    For Each wDta In lvTypes.ListItems
        '
        ' Only do something to selected types
        '
        If wDta.Checked Then
            Wa = Val(wDta.Tag)
            Wsa = fReadValue("HKCR", "." + wDta.Text, "", "S", "TTDXEdit")
            If (Wsa Like "TTDXEdit*") Or optTOver.Value Then
                '
                ' Handle my own, unassigned and overwrite
                '
                SetThisFileType "." + wDta.Text, Wa
            Else
                '
                ' Do plug-in
                '
                SetAMaster Wsa, Wa, True
            End If
        End If
    Next wDta
    Wa = MsgBox("The file associations you have selected have been set.", vbInformation)
    RefreshTypes
End Sub


Private Sub SetAMaster(wName As String, wMode As Integer, wPlug As Boolean)
    Dim Wa As Long, Wsa As String
    
    Wsa = F.BuildPath(App.Path, App.EXEName + ".exe")
    If Not wPlug Then
        If wMode = 1 Then Wa = fWriteValue("HKCR", wName, "", "S", "TTDX Savegame")
        If wMode = 2 Then Wa = fWriteValue("HKCR", wName, "", "S", "TTDX Scenario")
        If wMode = 3 Then Wa = fWriteValue("HKCR", wName, "", "S", "TTDX Uncompressed Data")
        Wa = fWriteValue("HKCR", wName + "\Shell", "", "S", "TTDXEdit")
    End If
    Wa = fWriteValue("HKCR", wName + "\Shell\TTDXEdit", "", "S", "Edit With TTDXEdit")
    Wa = fWriteValue("HKCR", wName + "\Shell\TTDXEdit\Command", "", "S", Chr(34) + Wsa + Chr(34) + " " + Chr(34) + "%1" + Chr(34))
    If wMode = 3 Then
        Wa = fWriteValue("HKCR", wName + "\Shell\TTDXPack", "", "S", "Pack to Gameformat")
    Else
        Wa = fWriteValue("HKCR", wName + "\Shell\TTDXUnpack", "", "S", "Make Unpacked File")
        Wa = fWriteValue("HKCR", wName + "\Shell\TTDXUnpack\Command", "", "S", Chr(34) + Wsa + Chr(34) + " " + Chr(34) + "%1" + Chr(34) + " /SU")
        Wa = fWriteValue("HKCR", wName + "\Shell\TTDXPack", "", "S", "Repack to Gameformat")
    End If
    Wa = fWriteValue("HKCR", wName + "\Shell\TTDXPack\Command", "", "S", Chr(34) + Wsa + Chr(34) + " " + Chr(34) + "%1" + Chr(34) + " /S")

End Sub

Private Sub SetThisFileType(wExt As String, wMode As Integer)
    Dim Wa As Long
    '
    ' Associate fileextension with predefined master
    '
    If wMode = 1 Then: Wa = fWriteValue("HKCR", wExt, "", "S", "TTDXEdit.Save")
    If wMode = 2 Then: Wa = fWriteValue("HKCR", wExt, "", "S", "TTDXEdit.Scenario")
    If wMode = 3 Then: Wa = fWriteValue("HKCR", wExt, "", "S", "TTDXEdit.Unpack")
End Sub

Private Sub RefreshTypes()
    '
    ' Show a list of filetypes and their curent assign
    '
    Dim Wva As Variant, Wa As Long, Wsa As String, Wsb As String, wDta As ListItem
    ' The number before each type is used to select the types of events availeble.
    Const Types As String = "1sv1|1sv2|1ss1|2sv0|2ss0|3sv1dta|3sv2dta|3ss1dta|3sv0dta|3ss0dta|3sv1hdr|3sv2hdr|3ss1hdr|3sv0hdr|3ss0hdr"
    Wva = Split(Types, "|")
    lvTypes.ListItems.Clear
    For Wa = 0 To UBound(Wva)
        Set wDta = lvTypes.ListItems.Add
        ' Get curent file reference (if any)
        Wsa = Mid(Wva(Wa), 2): Wsb = fReadValue("HKCR", "." + Wsa, "", "S", "")
        wDta.Text = Wsa
        wDta.Checked = True: wDta.Tag = Left(Wva(Wa), 1)
        If Wsb Like "TTDXEdit*" Then
            Wsb = "This editor"
            wDta.Bold = True
        ElseIf Wsb > " " Then
            ' Get the descriptionname curently assigned
            Wsb = fReadValue("HKCR", Wsb, "", "S", "")
        End If
        wDta.SubItems(1) = Wsb
    Next Wa
End Sub

Private Sub Form_Load()
    RefreshTypes
End Sub

