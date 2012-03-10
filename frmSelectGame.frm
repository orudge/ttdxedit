VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmSelectGame 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select SaveGame"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8655
   ControlBox      =   0   'False
   Icon            =   "frmSelectGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkHideTTD 
      Caption         =   "Hide TTD Info"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   7
      Top             =   3780
      Width           =   2655
   End
   Begin VB.ComboBox cmbFtypes 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   4140
      Width           =   3735
   End
   Begin VB.TextBox txtSelected 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3360
      TabIndex        =   5
      Top             =   3780
      Width           =   3735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7920
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   8
      ImageHeight     =   8
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectGame.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectGame.frx":062E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSelectGame.frx":0A82
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   7200
      TabIndex        =   4
      Top             =   4140
      Width           =   1335
   End
   Begin VB.DriveListBox Drive1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3255
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3240
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   3255
   End
   Begin MSComctlLib.ListView lvFiles 
      Height          =   3615
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   6376
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ColHdrIcons     =   "ImageList1"
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "File"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Date"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Name Of The Game"
         Object.Width           =   5716
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
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
      Height          =   315
      Left            =   7200
      TabIndex        =   0
      Top             =   3780
      Width           =   1335
   End
End
Attribute VB_Name = "frmSelectGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public CurPath As String
Public Selected As String
Public FileSet As Integer
Public FileMode As Integer ' Use 0 For Load, 1 For save as, 2 For "save as" with owerwrite test

Private F As New FileSystemObject
Private wWork(48) As Byte
Private fInit As Boolean
Private LastPath As String

Private Sub UpdateList()
    Dim wFo As Folder, wFf As File, wDta As ListItem, Wsa As String
    If CurPath = LastPath Then Exit Sub
    lvFiles.Enabled = False
    lvFiles.ListItems.Clear
    Selected = ""
    
    If F.FolderExists(CurPath) Then
        Set wFo = F.GetFolder(CurPath)
        Wsa = ""
        Select Case FileSet
            Case 0: Wsa = ".sv1.sv2."
            Case 1: Wsa = ".sv0.ss0.ss1."
            Case 2: Wsa = ".sv1hdr.sv2hdr."
            Case 3: Wsa = ".sv0hdr.ss0hdr.ss1hdr."
            Case 4: Wsa = ".sv1.sv2.sv0.ss0.ss1.sv1hdr.sv2hdr.sv0hdr.ss0hdr.ss1hdr."
        End Select
        For Each wFf In wFo.Files
            If InStr(Wsa, "." + F.GetExtensionName(wFf.Path) + ".") Or Wsa = "" Then
                If wFf.Name Like "tr?##.*" Then
                    Set wDta = lvFiles.ListItems.Add
                    wDta.Text = wFf.Name
                    wDta.SubItems(1) = Format(wFf.DateLastModified, "YYYY-MM-DD HH:MM:SS")
                    If chkHideTTD.Value <> 1 Then wDta.SubItems(2) = GetName(wFf.Path)
                Else
                    Set wDta = lvFiles.ListItems.Add
                    wDta.Text = wFf.Name
                    wDta.SubItems(1) = Format(wFf.DateLastModified, "YYYY-MM-DD HH:MM:SS")
                End If
            End If
        Next wFf
        LastPath = CurPath
        lvFiles.Enabled = True
    End If
    UpdCols
End Sub
Private Function GetName(wFile As String) As String
    Dim Wa As Integer, Wb As Long
    
    GetName = ""
    If F.FileExists(wFile) Then
        If F.GetFile(wFile).Size < 49 Then Exit Function
        Open wFile For Binary As 1
        Get 1, , wWork()
        Close 1
        Wb = TTDXCalcHdCheck(wWork)
        If Wb = wWork(47) + wWork(48) * 256& Then
            For Wa = 0 To 46
                If wWork(Wa) < 32 Then Exit For
                GetName = GetName + Chr(wWork(Wa))
            Next Wa
        End If
    End If
End Function

Private Sub chkHideTTD_Click()
    LastPath = "": UpdateList
End Sub

Private Sub cmbFtypes_Click()
    LastPath = "": FileSet = cmbFtypes.ListIndex
    UpdateList
End Sub

Private Sub cmdCancel_Click()
    Selected = ""
    Me.Hide
End Sub

Private Sub cmdOK_Click()
    If Selected > " " Then
        Selected = F.BuildPath(CurPath, lvFiles.SelectedItem.Text)
        If F.FileExists(Selected) Then
            Me.Hide
        End If
    End If
End Sub

Private Sub Dir1_Change()
    CurPath = Dir1.Path: Dir1.Enabled = True: UpdateList
End Sub

Private Sub Dir1_Click()
    CurPath = Dir1.List(Dir1.ListIndex): UpdateList
End Sub

Private Sub Drive1_Change()
    Dim Wsa As String
    If fInit Then Exit Sub
    
    Wsa = Left(Drive1.Drive, 1)
    Dir1.Enabled = False
    If Wsa > "" Then
        If F.GetDrive(Wsa).IsReady Then Dir1.Path = F.GetDrive(Wsa).RootFolder: Dir1.Enabled = True
    End If
End Sub

Private Sub Form_Activate()
    Dim Wa As Integer, Wsa As String
    fInit = True
    
    LastPath = ""
    Wa = InStr(CurPath, ":")
    Drive1.ListIndex = -1
    If Wa > 0 Then
        Wsa = Left(CurPath, Wa)
        For Wa = 0 To Drive1.ListCount - 1
            If Left(Drive1.List(Wa), Len(Wsa)) = Wsa Then Drive1.ListIndex = Wa: Exit For
        Next Wa
        If Drive1.ListIndex > -1 Then
            Wsa = CurPath
            While Not F.FolderExists(Wsa)
                Wsa = F.GetParentFolderName(Wsa)
            Wend
            Dir1.Path = Wsa: CurPath = Wsa
        End If
    End If
    Selected = ""
    If FileMode = 0 Then txtSelected.Enabled = False Else txtSelected.Enabled = True
    cmbFtypes.ListIndex = FileSet
    UpdateList
    fInit = False
End Sub

Private Sub Form_Load()
    Dim Wa As Long
    chkHideTTD.Value = fReadValue("HKCU", RegBaseKey + "\Selector", "HideTTD", "D", 0)
    'lvFiles.ColumnHeaders(2).Width = 0
    cmbFtypes.Clear
    cmbFtypes.AddItem "Savegames"
    cmbFtypes.AddItem "Scenarios"
    cmbFtypes.AddItem "Uncompressed Savegames"
    cmbFtypes.AddItem "Uncompressed Scenarios"
    cmbFtypes.AddItem "All usable files"
    cmbFtypes.AddItem "All files"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Wa As Long
    Wa = fWriteValue("HKCU", RegBaseKey + "\Selector", "HideTTD", "D", chkHideTTD.Value)
End Sub

Private Sub lvFiles_ColumnClick(ByVal Wa As MSComctlLib.ColumnHeader)
    If lvFiles.Sorted = False Then
        lvFiles.Sorted = True
        lvFiles.SortKey = Wa.Index - 1
    Else
        If lvFiles.SortKey = Wa.Index - 1 Then
            lvFiles.SortOrder = lvFiles.SortOrder * -1 + 1
        Else
            lvFiles.SortKey = Wa.Index - 1
            lvFiles.SortOrder = 1
        End If
    End If
    UpdCols
End Sub
Private Sub UpdCols()
    Dim Wa As Integer
    
    For Wa = 1 To lvFiles.ColumnHeaders.Count
        lvFiles.ColumnHeaders.Item(Wa).Icon = 1
    Next Wa
    If lvFiles.Sorted Then
        If lvFiles.SortOrder = 0 Then Wa = 3 Else Wa = 2
        lvFiles.ColumnHeaders.Item(lvFiles.SortKey + 1).Icon = Wa
    End If
End Sub

Private Sub lvFiles_DblClick()
    If Selected > " " Then
        cmdOK_Click
    End If
End Sub

Private Sub lvFiles_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Selected = Item.Text
    txtSelected.Text = Selected
End Sub
