VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm frmMDI 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   8160
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11400
   Icon            =   "frmMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   OLEDropMode     =   1  'Manual
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar stbar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   7905
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   847
            MinWidth        =   441
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   609
            MinWidth        =   353
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picTools 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   11400
      TabIndex        =   0
      Top             =   0
      Width           =   11400
      Begin VB.CommandButton Command1 
         Caption         =   "Check 44BBD"
         Height          =   375
         Left            =   5160
         TabIndex        =   8
         Top             =   120
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton cmdFinances 
         Height          =   540
         Left            =   720
         Picture         =   "frmMDI.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         Width           =   540
      End
      Begin VB.CommandButton cmdVeh 
         Height          =   540
         Left            =   3120
         Picture         =   "frmMDI.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   540
      End
      Begin VB.CommandButton cmdStations 
         Height          =   540
         Left            =   2520
         Picture         =   "frmMDI.frx":1086
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   540
      End
      Begin VB.CommandButton cmdPlayer 
         Height          =   540
         Left            =   120
         Picture         =   "frmMDI.frx":1B90
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   0
         Width           =   540
      End
      Begin VB.CommandButton cmdIndu 
         Height          =   540
         Left            =   1920
         Picture         =   "frmMDI.frx":269A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   540
      End
      Begin VB.CommandButton cmdCity 
         Height          =   540
         Left            =   1320
         Picture         =   "frmMDI.frx":31A4
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   0
         Width           =   555
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Save Game As"
      Filter          =   $"frmMDI.frx":3CAE
   End
   Begin VB.Menu mnFile 
      Caption         =   "&File"
      Begin VB.Menu mnFLoad 
         Caption         =   "&Load Game"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnFsave 
         Caption         =   "&Save Game"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnFsaveAs 
         Caption         =   "Save Game As"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu mnSep1 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnFsaveU 
         Caption         =   "Save &Uncompressed"
      End
      Begin VB.Menu mnSep1a 
         Caption         =   "-"
      End
      Begin VB.Menu mnQuit 
         Caption         =   "&Quit"
      End
      Begin VB.Menu mnCleanQuit 
         Caption         =   "&Cleanout and Quit"
      End
   End
   Begin VB.Menu mnView 
      Caption         =   "&Options"
      Begin VB.Menu mnVtool 
         Caption         =   "&Toolbar"
         Checked         =   -1  'True
         Shortcut        =   ^T
      End
      Begin VB.Menu mnVTech 
         Caption         =   "Tech-&Info"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnVmap 
         Caption         =   "&Map"
         Begin VB.Menu mnVMnone 
            Caption         =   "&None"
            Shortcut        =   ^{F1}
         End
         Begin VB.Menu mnVMsmall 
            Caption         =   "&Small"
            Shortcut        =   ^{F2}
         End
         Begin VB.Menu mnVMlarge 
            Caption         =   "&Large"
            Shortcut        =   ^{F3}
         End
         Begin VB.Menu mnVMextr 
            Caption         =   "&Extreme"
            Shortcut        =   ^{F4}
         End
         Begin VB.Menu mnSep9 
            Caption         =   "-"
         End
      End
      Begin VB.Menu mnuOCurrency 
         Caption         =   "&Currency"
         Begin VB.Menu mnuOCCur 
            Caption         =   "£ - &Pounds"
            Index           =   0
         End
         Begin VB.Menu mnuOCCur 
            Caption         =   "$ - &Dollars"
            Index           =   1
         End
         Begin VB.Menu mnuOCCur 
            Caption         =   "¥ - &Yen"
            Index           =   2
         End
         Begin VB.Menu mnuOCCur 
            Caption         =   "Fr - &Francs"
            Index           =   3
         End
         Begin VB.Menu mnuOCCur 
            Caption         =   "DM - D&eutschmark"
            Index           =   4
         End
         Begin VB.Menu mnuOCCur 
            Caption         =   "Pt - Pe&setas"
            Index           =   5
         End
         Begin VB.Menu mnuOCCur 
            Caption         =   "€ - E&uro"
            Index           =   6
         End
         Begin VB.Menu mnuOCCur 
            Caption         =   "Ft - &Hungarian Forint"
            Index           =   7
         End
         Begin VB.Menu mnuOCCur 
            Caption         =   "zl - Polish &Zloty"
            Index           =   8
         End
         Begin VB.Menu mnuOCCur 
            Caption         =   "ATS - Aus&trian Shilling"
            Index           =   9
         End
         Begin VB.Menu mnuOCCur 
            Caption         =   "BEF - Bel&gian Franc"
            Index           =   10
         End
         Begin VB.Menu mnuOCCur 
            Caption         =   "DKK - Da&nish Krone"
            Index           =   11
         End
         Begin VB.Menu mnuOCCur 
            Caption         =   "FIM - Finnish Mar&kka"
            Index           =   12
         End
         Begin VB.Menu mnuOCCur 
            Caption         =   "GRD - &Greek Drachma"
            Index           =   13
         End
         Begin VB.Menu mnuOCCur 
            Caption         =   "CHF - S&wiss Franc"
            Index           =   14
         End
         Begin VB.Menu mnuOCCur 
            Caption         =   "NLG - Dutch Guilde&r"
            Index           =   15
         End
         Begin VB.Menu mnuOCCur 
            Caption         =   "ITL - Itali&an Lira"
            Index           =   16
         End
         Begin VB.Menu mnuOCCur 
            Caption         =   "SEK - Swed&ish Krona"
            Index           =   17
         End
         Begin VB.Menu mnuOCCur 
            Caption         =   "RUB - &Russian Rubel"
            Index           =   18
         End
      End
      Begin VB.Menu mnSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnOpFileAss 
         Caption         =   "File &Associations"
      End
      Begin VB.Menu mnOsgm 
         Caption         =   "&SGM Plugin"
      End
   End
   Begin VB.Menu mnPlayer 
      Caption         =   "&Players"
      Begin VB.Menu mnPedit 
         Caption         =   "&Edit"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnuPFinances 
         Caption         =   "&Finances"
      End
      Begin VB.Menu mnSep6 
         Caption         =   "-"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnTerrain 
      Caption         =   "&Terrain"
      Begin VB.Menu mnTEremwood 
         Caption         =   "Remove &Trees"
      End
      Begin VB.Menu mnTownAIr 
         Caption         =   "Own AI Roads"
      End
      Begin VB.Menu mnTownCbridge 
         Caption         =   "Own City Bridges"
      End
   End
   Begin VB.Menu mnCity 
      Caption         =   "&Cities"
      Begin VB.Menu mnCedit 
         Caption         =   "&Edit"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnSep10 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnCmaxrat 
         Caption         =   "Maximize Ratings"
         Visible         =   0   'False
         Begin VB.Menu mnCmaxrate 
            Caption         =   "0"
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnIndu 
      Caption         =   "&Industries"
      Begin VB.Menu mnIedit 
         Caption         =   "&Edit"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnSep11 
         Caption         =   "-"
         Index           =   0
      End
      Begin VB.Menu mnImaxPro 
         Caption         =   "M&aximize Production"
         Begin VB.Menu mnMaxProAll 
            Caption         =   "All"
         End
         Begin VB.Menu mnSep4 
            Caption         =   "-"
         End
         Begin VB.Menu mnMaxPro 
            Caption         =   "0"
            Index           =   0
         End
      End
      Begin VB.Menu mnIminPro 
         Caption         =   "M&inimize Production"
         Begin VB.Menu mnMinProAll 
            Caption         =   "All"
         End
         Begin VB.Menu mnSep5 
            Caption         =   "-"
         End
         Begin VB.Menu mnMinPro 
            Caption         =   "0"
            Index           =   0
         End
      End
   End
   Begin VB.Menu mnStations 
      Caption         =   "&Stations"
      Begin VB.Menu mnSedit 
         Caption         =   "&Edit"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnSep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnSownUn 
         Caption         =   "Own Unowned"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnVehicles 
      Caption         =   "&Vehicles"
      Begin VB.Menu mnVedit 
         Caption         =   "&Edit"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "&Windows"
      WindowList      =   -1  'True
   End
   Begin VB.Menu mnHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnHAbout 
         Caption         =   "&About"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private F As New FileSystemObject, CleanUp As Boolean
Sub CheckCurrencyItem(Item As Integer)
    On Error GoTo Error
    
    Dim MID As Long, TLMID As Long, SMID As Long
    
    MID = GetMenu(hwnd)
    TLMID = GetSubMenu(MID, 1)
    SMID = GetSubMenu(TLMID, 3)
    
    CheckMenuRadioItem SMID, 0, 18, Item, MF_BYPOSITION
    Exit Sub
Error:
    Select Case ErrorProc(Err, "Function: CheckCurrencyItem(" & Item & ")")
        Case 3:
            End
        Case 2:
            Resume Next
        Case 1:
            Resume
    End Select
End Sub


Private Sub cmdCity_Click()
    frmCity.Show
    frmCity.SetFocus
End Sub

Private Sub cmdFinances_Click()
    frmFinances.Show
    frmFinances.SetFocus
End Sub

Private Sub cmdIndu_Click()
    frmIndu.Show
    frmIndu.SetFocus
End Sub

Private Sub cmdPlayer_Click()
    frmPlayer.Show
    frmPlayer.SetFocus
End Sub

Private Sub cmdStations_Click()
    frmStation.Show
    frmStation.SetFocus
End Sub

Private Sub cmdVeh_Click()
    frmVehicle.Show
    frmVehicle.SetFocus
End Sub

Private Sub Command1_Click()
    Dim Tmpstr As String
    Dim i As Integer
    
    Tmpstr = wData(&H44BBD) & vbCrLf & vbCrLf

    For i = 0 To 7
        If wData(&H44BBD) And (2 ^ i) Then
            Tmpstr = Tmpstr & "i (" & (2 ^ i) & ") = Yes" & vbCrLf
        Else
            Tmpstr = Tmpstr & "i (" & (2 ^ i) & ") = No" & vbCrLf
        End If
    Next i
    
    MsgBox Tmpstr
End Sub

Private Sub MDIForm_Load()
    'On Error GoTo Error
    Dim Wa As Long, Wsa As String
    
    Me.Caption = "TTDX Editor" ' V" + Format(App.Major) + "." + Format(App.Minor, "00") + "." + Format(App.Revision, "000") + "  (C)Jens Vang Petersen 2002"
    Wa = fWriteValue("HKLM", RegBaseKey, "Version", "S", Format(App.Major) + "." + Format(App.Minor, "00") + "." + Format(App.Revision, "0000"))
    
    For Wa = 1 To 11: Load mnMaxPro(Wa): Load mnMinPro(Wa): Next Wa
    
    Width = fReadValue("HKCU", RegBaseKey & "\Options", "LastWidth", "D", Width)
    Height = fReadValue("HKCU", RegBaseKey & "\Options", "LastHeight", "D", Height)
    Left = CInt(fReadValue("HKCU", RegBaseKey & "\Options", "LastLeft", "S", CStr(Left)))
    Top = CInt(fReadValue("HKCU", RegBaseKey & "\Options", "LastTop", "S", CStr(Top)))
    WindowState = fReadValue("HKCU", RegBaseKey & "\Options", "LastWindowState", "D", WindowState)
    
    Me.Show: DoEvents: Me.Show
    SetMenus
    BasTests
    CleanUp = False
    
    ' Load currency defaults
    Wa = fReadValue("HKCU", RegBaseKey & "\Options", "Currency", "D", 0)
    mnuOCCur_Click (Wa)
End Sub

Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Wa As Long, Wsa As String
    
    If Data.GetFormat(15) Then
        Me.OLEDropMode = 0
        Wa = Data.Files.Count
        If Wa > 1 Then
            Wa = MsgBox("This program only supports the dropping of one file.", 48)
        ElseIf Wa = 1 Then
            Wsa = Data.Files(1)
            If F.FileExists(Wsa) Then
                If InStr(".sv0.sv1.sv2.ss0.ss1.", "." + Left(F.GetExtensionName(Wsa), 3) + ".") Then
                    CallFileLoad Wsa
                End If
            End If
        End If
        Me.OLEDropMode = 1
    End If
End Sub

Private Sub MDIForm_Resize()
    frmTechInfo.SetMe
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    Dim Wa As Long
    If FileChanged Then
        Wa = MsgBox("Exit without saving changes?", 33)
        If Wa = 2 Then Cancel = 1: Exit Sub
    End If
    Unload frmTechInfo
    Wa = fWriteValue("HKCU", RegBaseKey + "\Map", "Size", "D", frmMap.MapSize)
    Wa = fWriteValue("HKLM", RegBaseKey, "LastPath", "S", frmSelectGame.CurPath)
    Wa = fWriteValue("HKCU", RegBaseKey + "\View", "ToolBar", "B", mnVtool.Checked)
    Wa = fWriteValue("HKCU", RegBaseKey + "\View", "TechBar", "B", mnVTech.Checked)
    Wa = fWriteValue("HKCU", RegBaseKey, "FastMode", "B", fFastMode)
    
    Wa = fWriteValue("HKCU", RegBaseKey + "\Options", "LastWidth", "D", Width)
    Wa = fWriteValue("HKCU", RegBaseKey + "\Options", "LastHeight", "D", Height)
    Wa = fWriteValue("HKCU", RegBaseKey + "\Options", "LastLeft", "S", frmMDI.Left)
    Wa = fWriteValue("HKCU", RegBaseKey + "\Options", "LastTop", "S", frmMDI.Top)
    Wa = fWriteValue("HKCU", RegBaseKey + "\Options", "LastWindowState", "D", WindowState)

    If CleanUp Then Wa = fRecDeleteKey("HKCU", "Software\Owen Rudge", "TTDX Editor")
    If CleanUp Then Wa = fRecDeleteKey("HKLM", "Software\Owen Rudge", "TTDX Editor")
End Sub

Private Sub mnCedit_Click()
    cmdCity_Click
End Sub

Private Sub mnCleanQuit_Click()
    Dim Wa As Long
    Wa = MsgBox("All TTDX Editor settings are now being removed." + Chr(10) + "NOTE: File associations and SGM Plugin are NOT affected", 33)
    If Wa = 1 Then
        CleanUp = True
        Unload Me
    End If
End Sub

Private Sub mnFLoad_Click()
    frmSelectGame.FileSet = 4
    frmSelectGame.FileMode = 0
    frmSelectGame.Caption = "Load Game"
    frmSelectGame.Show vbModal, Me
    
    If frmSelectGame.Selected > " " Then
        CallFileLoad frmSelectGame.Selected
    End If
End Sub

Public Sub CallFileLoad(wFile As String)
    Dim Wa As Integer, Wua As TTDXgeneral, Wv As TTDXplayer
    DoEvents
    frmWSplash.labText.Caption = "Loading File."
    Me.Enabled = False: DoEvents
    Unload frmCity
    Unload frmIndu
    Unload frmPlayer
    Unload frmStation
    Unload frmFinances
    frmWSplash.Show 0, Me
    frmWSplash.Refresh: DoEvents
    Wa = TTDXLoadFile(wFile)
    Me.Enabled = True
    If (Wa = 0) And (Not fAutoMode) Then
        frmMap.UpdateInfo
        Wua = TTDXGeneralInfo
        stbar.Panels(1).Text = "File: " + F.GetFileName(CurFile)
        stbar.Panels(2).Text = "Climate: " + Wua.ClimName
        stbar.Panels(3).Text = "City Names: " + Wua.CityNames
        stbar.Panels(4).Text = "Vehicle Array: " + Format(Wua.VehSize)
    ElseIf Wa < 100 Then
        stbar.Panels(1).Text = "Load Error: " + TTDXLoadError(Wa)
    End If
    SetMenus
    Unload frmWSplash
End Sub


Private Sub mnFsave_Click()
    CallFileSave 0
End Sub

Private Sub mnFsaveAs_Click()
    On Error Resume Next
    
    CommonDialog1.Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNOverwritePrompt
    CommonDialog1.FileName = CurFile
    CommonDialog1.ShowSave
    
    If Err = 0 Then
        CallFileSave 0, CommonDialog1.FileName
    End If
End Sub

Private Sub mnFsaveU_Click()
    CallFileSave 1
End Sub

Public Sub CallFileSave(wMode As Integer, Optional ByVal wFile As String)
    Dim Wa As Integer
    If CurFile > " " Then
        If wFile = "" Then wFile = CurFile
        Select Case wMode
            Case 0: frmWSplash.labText.Caption = "Saving File.."
            Case 1: frmWSplash.labText.Caption = "Saving Uncompressed.."
        End Select
        Me.Enabled = False
        frmWSplash.Show 0, Me
        frmWSplash.Refresh
        DoEvents
        If Not fAutoMode Then
            If frmIndu.Visible Then frmIndu.PrepSave
            If frmIndu.Visible Then frmCity.PrepSave
            If frmPlayer.Visible Then frmPlayer.PrepSave
            If frmStation.Visible Then frmStation.PrepSave
            If frmVehicle.Visible Then frmVehicle.PrepSave
        End If
        Me.Enabled = False
        frmWSplash.Show 0, Me
        frmWSplash.Refresh
        DoEvents
        Select Case wMode
            Case 0: Wa = TTDXSaveFile(CurFile)
            Case 1: Wa = TTDXSaveUncom(CurFile)
        End Select
        DoEvents
        Me.Enabled = True
        Unload frmWSplash
        If Wa <> 0 Then Wa = MsgBox("Save Failed", 48)
    End If

End Sub

Private Sub mnHAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnIedit_Click()
    cmdIndu_Click
End Sub

Private Sub mnMaxPro_Click(Index As Integer)
    MacroIndustry 1, Index
    frmIndu.UpdateInfo
End Sub
Private Sub mnMaxProAll_Click()
    MacroIndustry 1, -1
    frmIndu.UpdateInfo
End Sub

Private Sub mnMinPro_Click(Index As Integer)
    MacroIndustry 2, Index
    frmIndu.UpdateInfo
End Sub

Private Sub mnMinProAll_Click()
    MacroIndustry 2, -1
    frmIndu.UpdateInfo
End Sub

Private Sub mnOpFileAss_Click()
    frmFileTypes.Show vbModal, Me
End Sub

Private Sub mnOsgm_Click()
    frmSGM.Show vbModal, Me
End Sub

Private Sub mnPedit_Click()
    cmdPlayer_Click
End Sub

Private Sub BasTests()
    Dim Wsa As String, Wa As Long
    Wsa = F.BuildPath(App.Path, App.EXEName)
    If fReadValue("HKLM", RegBaseKey, "Path", "S", "") <> Wsa Then
        Wa = MsgBox("This program has been moved since last run" + Chr(10) + "If you have assigned filetypes you should update them now.", 48)
        Wa = fWriteValue("HKLM", RegBaseKey, "Path", "S", Wsa)
    End If
    'Wa = fReadValue("HKLM", RegBaseKey, "Version", "S", App.Revision)
    frmMap.MapSize = fReadValue("HKCU", RegBaseKey + "\Map", "Size", "D", 0)
    frmMap.Move 0, 0
    frmSelectGame.CurPath = fReadValue("HKLM", RegBaseKey, "LastPath", "S", App.Path)
    mnVtool.Checked = Not fReadValue("HKCU", RegBaseKey + "\View", "ToolBar", "B", True): mnVtool_Click
    mnVTech.Checked = Not fReadValue("HKCU", RegBaseKey + "\View", "TechBar", "B", False): mnVTech_Click
    fFastMode = True
End Sub


Private Sub mnQuit_Click()
    Unload Me
End Sub

Private Sub mnSedit_Click()
    cmdStations_Click
End Sub


Private Sub mnTEremwood_Click()
    Dim Wa As Integer
    TTDXtermacRemWood
    Wa = MsgBox("Trees Removed")
    frmMap.UpdateInfo
End Sub

Private Sub mnTownAIr_Click()
    Dim Wa As Integer
    TTDXtermacOwnAIRoad
End Sub

Private Sub mnTownCbridge_Click()
    TTDXtermacOwnCityBridge
End Sub

Private Sub mnuOCCur_Click(Index As Integer)
    On Error GoTo Error
    Dim LangCharset As Byte
    LangCharset = 0
    
    Select Case Index
        Case 0
            CurrencyMultiplier = 1
            
            If LangCharset <> 0 Then
                CurrencyLabel = "GBP "
            Else
                CurrencyLabel = "£"
            End If
            
            CurrencySeparator = ","
            CurrencySymbolBefore = True
        Case 1
            CurrencyMultiplier = 2
            CurrencyLabel = "$"
            CurrencySeparator = ","
            CurrencySymbolBefore = True
        Case 2
            CurrencyMultiplier = 20
            
            If LangCharset <> 0 Then
                CurrencyLabel = "JPY "
            Else
                CurrencyLabel = "¥"
            End If
      
            CurrencySeparator = ","
            CurrencySymbolBefore = True
        Case 3
            CurrencyMultiplier = 10
            CurrencyLabel = "Fr "
            CurrencySeparator = "."
            CurrencySymbolBefore = True
        Case 4
            CurrencyMultiplier = 4
            CurrencyLabel = "DM "
            CurrencySeparator = "."
            CurrencySymbolBefore = True
        Case 5
            CurrencyMultiplier = 20
            CurrencyLabel = "Pt "
            CurrencySeparator = "."
            CurrencySymbolBefore = True
        Case 6
            CurrencyMultiplier = 2
            CurrencyLabel = "€"
            CurrencySeparator = ","
            CurrencySymbolBefore = True
        Case 7
            CurrencyMultiplier = 375.62
            CurrencyLabel = " Ft"
            CurrencySeparator = "."
            CurrencySymbolBefore = False
        Case 8
            CurrencyMultiplier = 6.079
            CurrencyLabel = " zl"
            CurrencySeparator = " "
            CurrencySymbolBefore = False
        Case 9
            CurrencyMultiplier = 19.41
            CurrencyLabel = "ATS "
            CurrencySeparator = ","
            CurrencySymbolBefore = True
        Case 10
            CurrencyMultiplier = 56.89
            CurrencyLabel = "BEF "
            CurrencySeparator = ","
            CurrencySymbolBefore = True
        Case 11
            CurrencyMultiplier = 10.48
            CurrencyLabel = "DKK "
            CurrencySeparator = ","
            CurrencySymbolBefore = True
        Case 12
            CurrencyMultiplier = 8.38
            CurrencyLabel = "FIM "
            CurrencySeparator = ","
            CurrencySymbolBefore = True
        Case 13
            CurrencyMultiplier = 480.47
            CurrencyLabel = "GRD "
            CurrencySeparator = ","
            CurrencySymbolBefore = True
        Case 14
            CurrencyMultiplier = 2.16
            CurrencyLabel = "CHF "
            CurrencySeparator = ","
            CurrencySymbolBefore = True
        Case 15
            CurrencyMultiplier = 3.11
            CurrencyLabel = "NLG "
            CurrencySeparator = ","
            CurrencySymbolBefore = True
        Case 16
            CurrencyMultiplier = 2730.58
            CurrencyLabel = "ITL "
            CurrencySeparator = ","
            CurrencySymbolBefore = True
        Case 17
            CurrencyMultiplier = 13.09
            CurrencyLabel = "SEK "
            CurrencySeparator = ","
            CurrencySymbolBefore = True
        Case 18
            CurrencyMultiplier = 4.859
            CurrencyLabel = " rur"
            CurrencySeparator = " "
            CurrencySymbolBefore = False
    End Select
    
    CheckCurrencyItem (Index)
    'SaveRegEntry "Options", "Currency", CStr(Index)
     fWriteValue "HKCU", RegBaseKey & "\Options", "Currency", "D", Index
    
    Exit Sub
Error:
    Select Case ErrorProc(Err, "Function: frmMDI.mnuOCCur_Click(" & Index & ")")
        Case 3:
            End
        Case 2:
            Resume Next
        Case 1:
            Resume
    End Select
End Sub

Private Sub mnuPFinances_Click()
    cmdFinances_Click
End Sub

Private Sub mnVedit_Click()
    frmVehicle.Show
    frmVehicle.SetFocus
End Sub

Private Sub mnVMextr_Click()
    frmMap.MapSize = 4: frmMap.UpdateInfo
End Sub
Private Sub mnVMlarge_Click()
    frmMap.MapSize = 2: frmMap.UpdateInfo
End Sub
Private Sub mnVMnone_Click()
    frmMap.MapSize = 0: frmMap.UpdateInfo
End Sub
Private Sub mnVMsmall_Click()
    frmMap.MapSize = 1: frmMap.UpdateInfo
End Sub


Private Sub mnVTech_Click()
    If mnVTech.Checked = True Then
        mnVTech.Checked = False
        frmTechInfo.Hide
    Else
        mnVTech.Checked = True
        frmTechInfo.Show 0, Me
        frmTechInfo.SetMe
    End If
End Sub

Private Sub mnVtool_Click()
    If mnVtool.Checked = True Then
        mnVtool.Checked = False
        picTools.Visible = False
    Else
        mnVtool.Checked = True
        picTools.Visible = True
    End If
End Sub

Private Sub SetMenus()
    Dim Wa As Integer
    If CurFile > " " Then
        '
        ' Update Menus
        '
        mnImaxPro.Enabled = True
        mnIminPro.Enabled = True
        For Wa = 0 To 11
            If CargoTypes(Wa) > ">" Then
                mnMaxPro(Wa).Visible = True: mnMaxPro(Wa).Caption = CargoTypes(Wa)
                mnMinPro(Wa).Visible = True: mnMinPro(Wa).Caption = CargoTypes(Wa)
            Else
                mnMaxPro(Wa).Visible = False
                mnMinPro(Wa).Visible = False
            End If
        Next Wa
        mnFsave.Enabled = True
        mnFsaveU.Enabled = True
        mnFsaveAs.Enabled = True
    Else
        mnIminPro.Enabled = False
        mnImaxPro.Enabled = False
        mnFsave.Enabled = False
        mnFsaveU.Enabled = False
        mnFsaveAs.Enabled = False
    End If
End Sub
