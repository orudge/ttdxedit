VERSION 5.00
Object = "{9FC6639B-4237-4FB5-93B8-24049D39DF74}#1.5#0"; "ExLvwU.ocx"
Object = "{1F9B9092-BEE4-4CAF-9C7B-9384AF087C63}#1.4#0"; "ShBrowserCtlsU.ocx"
Object = "{1F8F0FE7-2CFB-4466-A2BC-ABB441ADEDD5}#2.3#0"; "ExTvwU.ocx"
Object = "{A8F9B8E7-E699-4FCE-A647-72C877F8E632}#1.8#0"; "EditCtlsU.ocx"
Begin VB.Form frmSelectGame 
   Caption         =   "Open Game"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10545
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSelectGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   10545
   StartUpPosition =   1  'CenterOwner
   Begin ExLVwLibUCtl.ExplorerListView lvFiles 
      Height          =   3735
      Left            =   3195
      TabIndex        =   1
      Top             =   0
      Width           =   7215
      _cx             =   12726
      _cy             =   6588
      AbsoluteBkImagePosition=   0   'False
      AllowHeaderDragDrop=   -1  'True
      AllowLabelEditing=   -1  'True
      AlwaysShowSelection=   -1  'True
      Appearance      =   1
      AutoArrangeItems=   0
      AutoSizeColumns =   0   'False
      BackColor       =   -2147483643
      BackgroundDrawMode=   0
      BkImagePositionX=   0
      BkImagePositionY=   0
      BkImageStyle    =   2
      BlendSelectionLasso=   -1  'True
      BorderSelect    =   0   'False
      BorderStyle     =   0
      CallBackMask    =   0
      CheckItemOnSelect=   0   'False
      ClickableColumnHeaders=   -1  'True
      ColumnHeaderVisibility=   1
      DisabledEvents  =   3144701
      DontRedraw      =   0   'False
      DragScrollTimeBase=   -1
      DrawImagesAsynchronously=   0   'False
      EditBackColor   =   -2147483643
      EditForeColor   =   -2147483640
      EditHoverTime   =   -1
      EditIMEMode     =   -1
      EmptyMarkupTextAlignment=   1
      Enabled         =   -1  'True
      FilterChangedTimeout=   -1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      FullRowSelect   =   2
      GridLines       =   0   'False
      GroupFooterForeColor=   -2147483640
      GroupHeaderForeColor=   -2147483640
      GroupMarginBottom=   0
      GroupMarginLeft =   0
      GroupMarginRight=   0
      GroupMarginTop  =   12
      GroupSortOrder  =   0
      HeaderFullDragging=   -1  'True
      HeaderHotTracking=   0   'False
      HeaderHoverTime =   -1
      HeaderOLEDragImageStyle=   0
      HideLabels      =   0   'False
      HotForeColor    =   -1
      HotMousePointer =   0
      HotTracking     =   0   'False
      HotTrackingHoverTime=   -1
      HoverTime       =   -1
      IMEMode         =   -1
      IncludeHeaderInTabOrder=   0   'False
      InsertMarkColor =   0
      ItemActivationMode=   2
      ItemAlignment   =   0
      ItemBoundingBoxDefinition=   70
      ItemHeight      =   17
      JustifyIconColumns=   0   'False
      LabelWrap       =   -1  'True
      MinItemRowsVisibleInGroups=   0
      MousePointer    =   0
      MultiSelect     =   -1  'True
      OLEDragImageStyle=   0
      OutlineColor    =   -2147483633
      OwnerDrawn      =   0   'False
      ProcessContextMenuKeys=   -1  'True
      Regional        =   0   'False
      RegisterForOLEDragDrop=   0   'False
      ResizableColumns=   -1  'True
      RightToLeft     =   0
      ScrollBars      =   1
      SelectedColumnBackColor=   -1
      ShowFilterBar   =   0   'False
      ShowGroups      =   0   'False
      ShowHeaderChevron=   0   'False
      ShowHeaderStateImages=   0   'False
      ShowStateImages =   0   'False
      ShowSubItemImages=   0   'False
      SimpleSelect    =   0   'False
      SingleRow       =   0   'False
      SnapToGrid      =   0   'False
      SortOrder       =   0
      SupportOLEDragImages=   -1  'True
      TextBackColor   =   -1
      TileViewItemLines=   1
      TileViewLabelMarginBottom=   0
      TileViewLabelMarginLeft=   0
      TileViewLabelMarginRight=   0
      TileViewLabelMarginTop=   0
      TileViewSubItemForeColor=   -1
      TileViewTileHeight=   -1
      TileViewTileWidth=   -1
      ToolTips        =   3
      UnderlinedItems =   0
      UseMinColumnWidths=   0   'False
      UseSystemFont   =   0   'False
      UseWorkAreas    =   0   'False
      View            =   3
      VirtualMode     =   0   'False
      EmptyMarkupText =   "frmSelectGame.frx":030A
      FooterIntroText =   "frmSelectGame.frx":032A
   End
   Begin EditCtlsLibUCtl.TextBox txtSelected 
      Height          =   315
      Left            =   3195
      TabIndex        =   3
      Top             =   3840
      Width           =   4095
      _cx             =   7223
      _cy             =   556
      AcceptNumbersOnly=   0   'False
      AcceptTabKey    =   0   'False
      AllowDragDrop   =   -1  'True
      AlwaysShowSelection=   0   'False
      Appearance      =   1
      AutoScrolling   =   2
      BackColor       =   -2147483643
      BorderStyle     =   0
      CancelIMECompositionOnSetFocus=   0   'False
      CharacterConversion=   0
      CompleteIMECompositionOnKillFocus=   0   'False
      DisabledBackColor=   -2147483643
      DisabledEvents  =   7179
      DisabledForeColor=   -2147483640
      DisplayCueBannerOnFocus=   0   'False
      DontRedraw      =   0   'False
      DoOEMConversion =   0   'False
      DragScrollTimeBase=   -1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      FormattingRectangleHeight=   0
      FormattingRectangleLeft=   0
      FormattingRectangleTop=   0
      FormattingRectangleWidth=   0
      HAlignment      =   0
      HoverTime       =   -1
      IMEMode         =   -1
      InsertMarkColor =   0
      InsertSoftLineBreaks=   0   'False
      LeftMargin      =   -1
      MaxTextLength   =   -1
      Modified        =   0   'False
      MousePointer    =   0
      MultiLine       =   0   'False
      OLEDragImageStyle=   0
      PasswordChar    =   0
      ProcessContextMenuKeys=   -1  'True
      ReadOnly        =   -1  'True
      RegisterForOLEDragDrop=   0   'False
      RightMargin     =   -1
      RightToLeft     =   0
      ScrollBars      =   0
      SelectedTextMousePointer=   0
      SupportOLEDragImages=   -1  'True
      TabWidth        =   -1
      UseCustomFormattingRectangle=   0   'False
      UsePasswordChar =   0   'False
      UseSystemFont   =   0   'False
      CueBanner       =   "frmSelectGame.frx":034A
      Text            =   "frmSelectGame.frx":036A
   End
   Begin ExTVwLibUCtl.ExplorerTreeView tvDirs 
      Height          =   3735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      _cx             =   5530
      _cy             =   6588
      AllowDragDrop   =   -1  'True
      AllowLabelEditing=   -1  'True
      AlwaysShowSelection=   -1  'True
      Appearance      =   1
      AutoHScroll     =   -1  'True
      AutoHScrollPixelsPerSecond=   150
      AutoHScrollRedrawInterval=   15
      BackColor       =   -2147483643
      BlendSelectedItemsIcons=   0   'False
      BorderStyle     =   0
      BuiltInStateImages=   0
      CaretChangedDelayTime=   500
      DisabledEvents  =   1023
      DontRedraw      =   0   'False
      DragExpandTime  =   -1
      DragScrollTimeBase=   -1
      DrawImagesAsynchronously=   0   'False
      EditBackColor   =   -2147483643
      EditForeColor   =   -2147483640
      EditHoverTime   =   -1
      EditIMEMode     =   -1
      Enabled         =   -1  'True
      FadeExpandos    =   0   'False
      FavoritesStyle  =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      FullRowSelect   =   0   'False
      GroupBoxColor   =   -2147483632
      HotTracking     =   0
      HoverTime       =   -1
      IMEMode         =   -1
      Indent          =   16
      IndentStateImages=   -1  'True
      InsertMarkColor =   0
      ItemBoundingBoxDefinition=   70
      ItemHeight      =   17
      ItemXBorder     =   3
      ItemYBorder     =   0
      LineColor       =   -2147483632
      LineStyle       =   0
      MaxScrollTime   =   100
      MousePointer    =   0
      OLEDragImageStyle=   0
      ProcessContextMenuKeys=   -1  'True
      RegisterForOLEDragDrop=   0   'False
      RichToolTips    =   0   'False
      RightToLeft     =   0
      ScrollBars      =   2
      ShowStateImages =   0   'False
      ShowToolTips    =   -1  'True
      SingleExpand    =   0
      SortOrder       =   0
      SupportOLEDragImages=   -1  'True
      TreeViewStyle   =   3
      UseSystemFont   =   0   'False
   End
   Begin VB.CheckBox chkHideTTD 
      Caption         =   "&Hide TTD Info"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   3840
      Width           =   1455
   End
   Begin VB.ComboBox cmbFtypes 
      Height          =   315
      Left            =   7440
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   3840
      Width           =   3015
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   9120
      TabIndex        =   7
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   7680
      TabIndex        =   6
      Top             =   4320
      Width           =   1335
   End
   Begin VB.Image imgSplitter 
      Height          =   3735
      Left            =   3120
      MousePointer    =   9  'Size W E
      Top             =   0
      Width           =   105
   End
   Begin VB.Label lblFilename 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Filename:"
      Height          =   195
      Left            =   2460
      TabIndex        =   4
      Top             =   3840
      Width           =   675
   End
   Begin ShBrowserCtlsLibUCtl.ShellTreeView tvShell 
      Left            =   1440
      Top             =   4320
      Version         =   256
      AutoEditNewItems=   -1  'True
      ColorCompressedItems=   -1  'True
      ColorEncryptedItems=   -1  'True
      DefaultManagedItemProperties=   511
      BeginProperty DefaultNamespaceEnumSettings {CC889E2B-5A0D-42F0-AA08-D5FD5863410C} 
         EnumerationFlags=   161
         ExcludedFileItemFileAttributes=   0
         ExcludedFileItemShellAttributes=   0
         ExcludedFolderItemFileAttributes=   0
         ExcludedFolderItemShellAttributes=   0
         IncludedFileItemFileAttributes=   0
         IncludedFileItemShellAttributes=   536870912
         IncludedFolderItemFileAttributes=   0
         IncludedFolderItemShellAttributes=   0
      EndProperty
      DisabledEvents  =   111
      DisplayElevationShieldOverlays=   -1  'True
      HandleOLEDragDrop=   7
      HiddenItemsStyle=   2
      InfoTipFlags    =   536870912
      ItemEnumerationTimeout=   3000
      ItemTypeSortOrder=   0
      LimitLabelEditInput=   -1  'True
      LoadOverlaysOnDemand=   -1  'True
      PreselectBasenameOnLabelEdit=   -1  'True
      ProcessShellNotifications=   -1  'True
      UseGenericIcons =   1
      UseSystemImageList=   1
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

Private bResizing As Boolean

Private Sub SizeControls(Optional ByVal SplitterPos As Long = 0)
    If SplitterPos <> 0 Then
        imgSplitter.Left = SplitterPos - 15
        tvDirs.Width = SplitterPos
        lvFiles.Left = imgSplitter.Left + imgSplitter.Width - 15
    End If
    
    lvFiles.Width = Me.ScaleWidth - lvFiles.Left - 60
    
    cmbFtypes.Left = Me.ScaleWidth - cmbFtypes.Width - 60
    txtSelected.Width = Me.ScaleWidth - txtSelected.Left - cmbFtypes.Width - 100
    
    cmdCancel.Left = Me.ScaleWidth - cmdCancel.Width - 60
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 100
    
    cmdCancel.Top = Me.ScaleHeight - cmdCancel.Height - 60
    cmdOK.Top = cmdCancel.Top
    
    txtSelected.Top = cmdOK.Top - txtSelected.Height - 100
    cmbFtypes.Top = txtSelected.Top
    chkHideTTD.Top = txtSelected.Top
    
    lblFilename.Top = txtSelected.Top + ((txtSelected.Height - lblFilename.Height) / 2)
    
    tvDirs.Height = txtSelected.Top - 100
    lvFiles.Height = tvDirs.Height
    imgSplitter.Height = tvDirs.Height
End Sub


Private Sub UpdateList(Optional ByVal Force = False)
    Dim wFo As Folder, wFf As File, wDta As ListViewItem, Wsa As String
    
    If (CurPath = LastPath) And (Not Force) Then
        Exit Sub
    End If
    
    lvFiles.Enabled = False
    lvFiles.ListItems.RemoveAll
    Selected = ""
    
    If CurPath = "" Then
        Exit Sub
    End If
    
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
                Set wDta = lvFiles.ListItems.Add(wFf.Name)
                wDta.SubItems(1) = FormatDateTime(wFf.DateLastModified)
                
                If chkHideTTD.Value <> 1 Then
                    wDta.SubItems(2) = GetName(wFf.Path)
                End If
            End If
        Next wFf
        LastPath = CurPath
        lvFiles.Enabled = True
    End If
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
        Selected = F.BuildPath(CurPath, Selected)
        If F.FileExists(Selected) Then
            LastPath = CurPath
            Me.Hide
        End If
    End If
End Sub





Private Sub Form_Activate()
    fInit = True
    
    Dim pIDLCurPath As Long
    Dim stvi As ShellTreeViewItem
    
    If CurPath <> "" Then
        If SHParseDisplayName(StrPtr(CurPath), 0, pIDLCurPath, 0, 0) = 0 Then
            Screen.MousePointer = MousePointerConstants.vbHourglass
            
            tvShell.EnsureItemIsLoaded pIDLCurPath
            Set stvi = tvShell.TreeItems(pIDLCurPath, ShTvwItemIdentifierTypeConstants.stiitEqualPIDL)
            
            If Not (stvi Is Nothing) Then
                Set tvDirs.CaretItem = stvi.TreeViewItemObject
            End If
            
            ILFree pIDLCurPath
        
            Screen.MousePointer = MousePointerConstants.vbDefault
        End If
    End If
    
    Selected = ""
    
    txtSelected.Enabled = IIf(FileMode = 0, False, True)
    
    cmbFtypes.ListIndex = FileSet
    UpdateList True
    fInit = False
End Sub

Private Sub Form_Load()
    Dim Wa As Long
    
    chkHideTTD.Value = fReadValue("HKCU", RegBaseKey + "\Selector", "HideTTD", "D", 0)
    
    lvFiles.Columns.Add "File", , 120
    lvFiles.Columns.Add "Date", , 120
    lvFiles.Columns.Add "Game Name", , 230
    
    cmbFtypes.Clear
    cmbFtypes.AddItem "Savegames"
    cmbFtypes.AddItem "Scenarios"
    cmbFtypes.AddItem "Uncompressed Savegames"
    cmbFtypes.AddItem "Uncompressed Scenarios"
    cmbFtypes.AddItem "All usable files"
    cmbFtypes.AddItem "All files"
    
    Dim itm As ShellTreeViewItem
    Dim pIDLDesktop As Long, pIDLCurPath As Long
    Dim OldPath As String
    
    tvShell.Attach tvDirs.hWnd
    tvShell.hWndShellUIParentWindow = Me.hWnd
    
    OldPath = CurPath
    
    SHGetFolderLocation Me.hWnd, CSIDL_DESKTOP, 0, 0, pIDLDesktop
    
    Set itm = tvShell.TreeItems.Add(pIDLDesktop, , InsertAfterConstants.iaFirst, , , HasExpandoConstants.heYes)
    
    If Not (itm Is Nothing) Then
        Set tvDirs.CaretItem = itm.TreeViewItemObject
        tvDirs.CaretItem.Expand
    End If
    
    If OldPath <> "" Then
        If SHParseDisplayName(StrPtr(OldPath), 0, pIDLCurPath, 0, 0) = 0 Then
            Screen.MousePointer = MousePointerConstants.vbHourglass
            
            tvShell.EnsureItemIsLoaded pIDLCurPath
            Set itm = tvShell.TreeItems(pIDLCurPath, ShTvwItemIdentifierTypeConstants.stiitEqualPIDL)
            
            If Not (itm Is Nothing) Then
                Set tvDirs.CaretItem = itm.TreeViewItemObject
            End If
            
            ILFree pIDLCurPath
        
            Screen.MousePointer = MousePointerConstants.vbDefault
        End If
    End If
End Sub

Private Sub Form_Resize()
    If Me.Width < 10635 Then
        Me.Width = 10635
        Exit Sub
    End If
    
    If Me.Height < 5175 Then
        Me.Height = 5175
        Exit Sub
    End If
    
    SizeControls
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Wa As Long
    Wa = fWriteValue("HKCU", RegBaseKey + "\Selector", "HideTTD", "D", chkHideTTD.Value)
    
    tvShell.Detach
End Sub


Private Sub imgSplitter_MouseDown(button As Integer, shift As Integer, x As Single, y As Single)
    bResizing = True
End Sub


Private Sub imgSplitter_MouseMove(button As Integer, shift As Integer, x As Single, y As Single)
    Dim Pos As Long
    
    If bResizing Then
        Pos = x + imgSplitter.Left
        
        If Pos < 2055 Then
            Pos = 2055
        ElseIf Pos > Me.ScaleWidth - 2055 Then
            Pos = Me.ScaleWidth - 2055
        End If
        
        SizeControls Pos
    End If
End Sub


Private Sub imgSplitter_MouseUp(button As Integer, shift As Integer, x As Single, y As Single)
    SizeControls
    bResizing = False
End Sub


Private Sub lvFiles_ColumnClick(ByVal column As ExLVwLibUCtl.IListViewColumn, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HeaderHitTestConstants)
    If lvFiles.SortOrder = soAscending Then
        lvFiles.SortOrder = soDescending
    Else
        lvFiles.SortOrder = soAscending
    End If
    
    lvFiles.SortItems sobText, , , , , column
End Sub

Private Sub lvFiles_DblClick(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
    If Selected > " " Then
        cmdOK_Click
    End If
End Sub

Private Sub lvFiles_ItemSelectionChanged(ByVal listItem As ExLVwLibUCtl.IListViewItem)
    Selected = listItem.Text
    txtSelected.Text = listItem.Text
End Sub

Private Sub tvDirs_CaretChanged(ByVal previousCaretItem As ExTVwLibUCtl.ITreeViewItem, ByVal newCaretItem As ExTVwLibUCtl.ITreeViewItem, ByVal caretChangeReason As ExTVwLibUCtl.CaretChangeCausedByConstants)
    Dim itm As ShellTreeViewItem
    Dim slvns As ShellListViewNamespace
    Dim Path As String
    
    Set itm = newCaretItem.ShellTreeViewItemObject
    
    If Not (itm Is Nothing) Then
        Path = String$(MAX_PATH, Chr$(0))
        SHGetPathFromIDList itm.FullyQualifiedPIDL, StrPtr(Path)
        Path = Left$(Path, lstrlen(StrPtr(Path)))
        
        If PathIsDirectory(StrPtr(Path)) = 0 Then
            CurPath = ""
        Else
            CurPath = Path
        End If
        
        UpdateList
    End If
End Sub


