VERSION 5.00
Object = "{9FC6639B-4237-4FB5-93B8-24049D39DF74}#1.5#0"; "ExLvwU.ocx"
Begin VB.Form frmFileTypes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Associations"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7650
   Icon            =   "frmFileTypes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   7650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ExLVwLibUCtl.ExplorerListView lvTypes 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      _cx             =   10186
      _cy             =   7011
      AbsoluteBkImagePosition=   0   'False
      AllowHeaderDragDrop=   0   'False
      AllowLabelEditing=   0   'False
      AlwaysShowSelection=   -1  'True
      Appearance      =   1
      AutoArrangeItems=   0
      AutoSizeColumns =   -1  'True
      BackColor       =   -2147483643
      BackgroundDrawMode=   0
      BkImagePositionX=   0
      BkImagePositionY=   0
      BkImageStyle    =   2
      BlendSelectionLasso=   -1  'True
      BorderSelect    =   0   'False
      BorderStyle     =   0
      CallBackMask    =   0
      CheckItemOnSelect=   -1  'True
      ClickableColumnHeaders=   -1  'True
      ColumnHeaderVisibility=   1
      DisabledEvents  =   3145725
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   -2147483640
      FullRowSelect   =   0
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
      MultiSelect     =   0   'False
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
      ShowStateImages =   -1  'True
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
      UseSystemFont   =   -1  'True
      UseWorkAreas    =   0   'False
      View            =   3
      VirtualMode     =   0   'False
      EmptyMarkupText =   "frmFileTypes.frx":000C
      FooterIntroText =   "frmFileTypes.frx":002C
   End
   Begin VB.OptionButton optTOver 
      Caption         =   "&Overwrite"
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
      Left            =   6060
      TabIndex        =   2
      Top             =   1200
      Width           =   1095
   End
   Begin VB.OptionButton optTplug 
      Caption         =   """&Plug-In"""
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
      Left            =   6060
      TabIndex        =   3
      Top             =   1440
      Value           =   -1  'True
      Width           =   1095
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
      Left            =   6000
      TabIndex        =   5
      Top             =   3720
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
      Left            =   6000
      TabIndex        =   4
      Top             =   3240
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
      Left            =   6000
      TabIndex        =   1
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

Public Sub AssociateCmdLine()
    On Error GoTo Error
    
    Dim CmdLine As String
    Dim CmdSplit() As String
    Dim i As Integer
    
    CmdLine = Trim$(Right$(Command$, Len(Command$) - 4))
    CmdSplit = Split(CmdLine, " ")
    
    SetAMaster "TTDXEdit.Save", 0, False
    SetAMaster "TTDXEdit.Scenario", 1, False
    SetAMaster "TTDXEdit.Unpack", 2, False

    For i = LBound(CmdSplit) To UBound(CmdSplit)
        Dim CmdOptions() As String
        
        CmdOptions = Split(CmdSplit(i), ":")
        
        If CmdOptions(0) = "o" Then
            SetThisFileType CmdOptions(1), CInt(CmdOptions(2))
        ElseIf CmdOptions(0) = "m" Then
            SetAMaster CmdOptions(1), CInt(CmdOptions(2)), True
        End If
    Next i
    
    Exit Sub
Error:
    End
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDo_Click()
    Dim Wa As Integer, wDta As ListViewItem, Wsa As String
    Dim FiletypeList As String
    
    If RunningWin9x() Then
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
            If wDta.StateImageIndex = 2 Then
                Wa = wDta.ItemData
                Wsa = fReadValue("HKCR", "." + wDta.Text, "", "S", "TTDXEdit")
                If (Wsa Like "TTDXEdit*") Or optTOver.Value Then
                    '
                    ' Handle my own, unassigned and overwrite
                    '
                    SetThisFileType "." + wDta.Text, Wa
                    FiletypeList = FiletypeList & "o:." & wDta.Text & ":" & Wa & " "
                Else
                    '
                    ' Do plug-in
                    '
                    SetAMaster Wsa, Wa, True
                    FiletypeList = FiletypeList & "m:" & Wsa & ":" & Wa & " "
                End If
            End If
        Next wDta
    Else
        '
        ' Loop through the filetypes
        '
        For Each wDta In lvTypes.ListItems
            '
            ' Only do something to selected types
            '
            If wDta.StateImageIndex = 2 Then
                Wa = wDta.ItemData
                Wsa = fReadValue("HKCR", "." + wDta.Text, "", "S", "TTDXEdit")
                If (Wsa Like "TTDXEdit*") Or optTOver.Value Then
                    '
                    ' Handle my own, unassigned and overwrite
                    '
                    FiletypeList = FiletypeList & "o:." & wDta.Text & ":" & Wa & " "
                Else
                    '
                    ' Do plug-in
                    '
                    FiletypeList = FiletypeList & "m:" & Wsa & ":" & Wa & " "
                End If
            End If
        Next wDta
        
        Screen.MousePointer = 11
        DoEvents
        
        StartElevated Me.hwnd, """" & MakePath(App.Path) & App.EXEName & ".exe""", "/FT " & FiletypeList, App.Path, 0, "In order to modify file associations, you need to be running as an administrator. If you press Yes, you'll be prompted to enter an Administrator password. If this fails, please try logging out and running TTDX Editor as an administrator." & vbCrLf & vbCrLf & "Do you want to proceed?"
        Screen.MousePointer = 0
     End If

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
    Dim Wva As Variant, Wa As Long, Wsa As String, Wsb As String, wDta As ListViewItem
    ' The number before each type is used to select the types of events availeble.
    Const Types As String = "1sv1|1sv2|1ss1|2sv0|2ss0|3sv1dta|3sv2dta|3ss1dta|3sv0dta|3ss0dta|3sv1hdr|3sv2hdr|3ss1hdr|3sv0hdr|3ss0hdr"
    Wva = Split(Types, "|")
    lvTypes.ListItems.RemoveAll
    For Wa = 0 To UBound(Wva)
        Wsa = MID(Wva(Wa), 2)
       
        ' Get current file reference (if any)
        Wsb = fReadValue("HKCR", "." + Wsa, "", "S", "")
        
        Set wDta = lvTypes.ListItems.Add(Wsa)
        wDta.Text = Wsa
        wDta.StateImageIndex = 2
        wDta.ItemData = CLng(Left(Wva(Wa), 1))
        If Wsb Like "TTDXEdit*" Then
            Wsb = "This editor"
            'wDta.Bold = True
        ElseIf Wsb > " " Then
            ' Get the description name currently assigned
            Wsb = fReadValue("HKCR", Wsb, "", "S", "")
        End If
        wDta.SubItems(1) = Wsb
    Next Wa
End Sub



Private Sub Form_Load()
    lvTypes.Columns.Add "File Type"
    lvTypes.Columns.Add "Assigned To"
    
    RefreshTypes

    If IsElevated() = False Then
        SendMessage cmdDo.hwnd, BCM_SETSHIELD, 0, &HFFFFFFFF
    End If
End Sub

Private Sub lvTypes_DblClick(ByVal listItem As ExLVwLibUCtl.IListViewItem, ByVal listSubItem As ExLVwLibUCtl.IListViewSubItem, ByVal button As Integer, ByVal shift As Integer, ByVal x As Single, ByVal y As Single, ByVal hitTestDetails As ExLVwLibUCtl.HitTestConstants)
    If listItem Is Nothing Then
        Exit Sub
    End If
    
    If listItem.StateImageIndex = 1 Then
        listItem.StateImageIndex = 2
    Else
        listItem.StateImageIndex = 1
    End If
End Sub


