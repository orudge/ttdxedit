VERSION 5.00
Object = "{956B5A46-C53F-45A7-AF0E-EC2E1CC9B567}#1.5#0"; "TrackBarCtlU.ocx"
Object = "{1F8F0FE7-2CFB-4466-A2BC-ABB441ADEDD5}#2.3#0"; "ExTvwU.ocx"
Begin VB.Form frmVehicle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vehicles"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVehicle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmDta 
      Caption         =   "General"
      Height          =   615
      Index           =   2
      Left            =   3750
      TabIndex        =   19
      Top             =   960
      Width           =   3135
      Begin VB.TextBox txtVal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "&Value (£):"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   270
         Width           =   1455
      End
   End
   Begin VB.Frame frmDta 
      Caption         =   "Engine Related"
      Height          =   1815
      Index           =   1
      Left            =   3750
      TabIndex        =   11
      Top             =   1560
      Width           =   3135
      Begin TrackBarCtlLibUCtl.TrackBar sliRel 
         Height          =   255
         Left            =   30
         TabIndex        =   13
         Top             =   450
         Width           =   3075
         _cx             =   5424
         _cy             =   450
         Appearance      =   0
         AutoTickFrequency=   26
         AutoTickMarks   =   -1  'True
         BackColor       =   -2147483633
         BackgroundDrawMode=   0
         BorderStyle     =   0
         CurrentPosition =   0
         DetectDoubleClicks=   -1  'True
         DisabledEvents  =   779
         DontRedraw      =   0   'False
         DownIsLeft      =   -1  'True
         Enabled         =   -1  'True
         HoverTime       =   -1
         LargeStepWidth  =   25
         Maximum         =   255
         Minimum         =   0
         MousePointer    =   0
         Orientation     =   0
         ProcessContextMenuKeys=   -1  'True
         RangeSelectionEnd=   0
         RangeSelectionStart=   0
         RegisterForOLEDragDrop=   0   'False
         Reversed        =   0   'False
         RightToLeftLayout=   0   'False
         SelectionType   =   0
         ShowSlider      =   -1  'True
         SliderLength    =   -1
         SmallStepWidth  =   1
         SupportOLEDragImages=   -1  'True
         TickMarksPosition=   1
         ToolTipPosition =   2
      End
      Begin VB.TextBox txtAge 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   16
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtAgeMax 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1680
         TabIndex        =   18
         Top             =   1440
         Width           =   1335
      End
      Begin TrackBarCtlLibUCtl.TrackBar sliRdr 
         Height          =   255
         Left            =   30
         TabIndex        =   14
         Top             =   720
         Width           =   3075
         _cx             =   5424
         _cy             =   450
         Appearance      =   0
         AutoTickFrequency=   80
         AutoTickMarks   =   -1  'True
         BackColor       =   -2147483633
         BackgroundDrawMode=   0
         BorderStyle     =   0
         CurrentPosition =   0
         DetectDoubleClicks=   -1  'True
         DisabledEvents  =   779
         DontRedraw      =   0   'False
         DownIsLeft      =   -1  'True
         Enabled         =   -1  'True
         HoverTime       =   -1
         LargeStepWidth  =   200
         Maximum         =   4000
         Minimum         =   0
         MousePointer    =   0
         Orientation     =   0
         ProcessContextMenuKeys=   -1  'True
         RangeSelectionEnd=   0
         RangeSelectionStart=   0
         RegisterForOLEDragDrop=   0   'False
         Reversed        =   0   'False
         RightToLeftLayout=   0   'False
         SelectionType   =   0
         ShowSlider      =   -1  'True
         SliderLength    =   -1
         SmallStepWidth  =   40
         SupportOLEDragImages=   -1  'True
         TickMarksPosition=   1
         ToolTipPosition =   2
      End
      Begin VB.Label Label2 
         Caption         =   "A&ge (Days):"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1110
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "&Max Age (Days):"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1470
         Width           =   1575
      End
      Begin VB.Label labRel 
         Caption         =   "&Reliability:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Owner"
      Height          =   615
      Left            =   5280
      TabIndex        =   4
      Top             =   0
      Width           =   1575
      Begin VB.ComboBox cmbOwner 
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame frmDta 
      Caption         =   "&Cargo"
      Height          =   1815
      Index           =   0
      Left            =   3750
      TabIndex        =   6
      Top             =   3360
      Width           =   3135
      Begin VB.ComboBox cmbCargo 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   2895
      End
      Begin TrackBarCtlLibUCtl.TrackBar sliMaxLoad 
         Height          =   255
         Left            =   30
         TabIndex        =   9
         Top             =   840
         Width           =   3060
         _cx             =   5397
         _cy             =   450
         Appearance      =   0
         AutoTickFrequency=   12
         AutoTickMarks   =   -1  'True
         BackColor       =   -2147483633
         BackgroundDrawMode=   0
         BorderStyle     =   0
         CurrentPosition =   0
         DetectDoubleClicks=   -1  'True
         DisabledEvents  =   779
         DontRedraw      =   0   'False
         DownIsLeft      =   -1  'True
         Enabled         =   -1  'True
         HoverTime       =   -1
         LargeStepWidth  =   20
         Maximum         =   600
         Minimum         =   0
         MousePointer    =   0
         Orientation     =   0
         ProcessContextMenuKeys=   -1  'True
         RangeSelectionEnd=   0
         RangeSelectionStart=   0
         RegisterForOLEDragDrop=   0   'False
         Reversed        =   0   'False
         RightToLeftLayout=   0   'False
         SelectionType   =   0
         ShowSlider      =   -1  'True
         SliderLength    =   -1
         SmallStepWidth  =   1
         SupportOLEDragImages=   -1  'True
         TickMarksPosition=   1
         ToolTipPosition =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Liquid cargos are multiplied by 100 in TTD, 1000 in TTDPatch."
         Height          =   390
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   2895
         WordWrap        =   -1  'True
      End
      Begin VB.Label labCMax 
         Caption         =   "Ma&x:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   630
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "&Active Vehicles"
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin ExTVwLibUCtl.ExplorerTreeView tvVeh 
         Height          =   4455
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   3375
         _cx             =   5953
         _cy             =   7858
         AllowDragDrop   =   -1  'True
         AllowLabelEditing=   0   'False
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
         DisabledEvents  =   263167
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
            Name            =   "Microsoft Sans Serif"
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
         ItemBoundingBoxDefinition=   94
         ItemHeight      =   17
         ItemXBorder     =   3
         ItemYBorder     =   0
         LineColor       =   -2147483632
         LineStyle       =   1
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
      Begin VB.ComboBox cmbShowC 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbShowO 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmVehicle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CurItm As TTDXVehicle, CurItmNo As Long, fInit As Boolean

Implements ISubclassedWindow

Private Sub Subclass()
    If Not SubclassWindow(Me.hwnd, Me, EnumSubclassID.escidCity) Then
        Debug.Print "Subclassing failed!"
    End If
    
    ' tell the controls to negotiate the correct format with the form
    SendMessageAsLong tvVeh.hwnd, WM_NOTIFYFORMAT, Me.hwnd, NF_REQUERY
End Sub

Private Function ISubclassedWindow_HandleMessage(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal eSubclassID As EnumSubclassID, bCallDefProc As Boolean) As Long
    Dim lRet As Long
    
    On Error GoTo StdHandler_End
    
    If eSubclassID = EnumSubclassID.escidVehicle Then
        lRet = HandleMessage_Form(hwnd, uMsg, wParam, lParam, bCallDefProc)
    End If
    
StdHandler_End:
    ISubclassedWindow_HandleMessage = lRet
End Function

Private Function HandleMessage_Form(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, bCallDefProc As Boolean) As Long
    Dim lRet As Long
    
    On Error GoTo StdHandler_End
    
    If uMsg = WM_NOTIFYFORMAT Then
        ' give the control a chance to request Unicode notifications
        lRet = SendMessageAsLong(wParam, OCM__BASE + uMsg, wParam, lParam)
        
        bCallDefProc = False
    End If
    
StdHandler_End:
    HandleMessage_Form = lRet
End Function
Public Sub UpdateInfo()
    Dim Wa As Integer, Wv As TTDXplayer
    
    fInit = True
    cmbShowO.Clear
    cmbShowC.Clear
    If CurFile > " " Then
        cmbShowO.AddItem "<All>": cmbShowO.ItemData(0) = -1
        For Wa = 0 To 7
            Wv = TTDXPlayerInfo(Wa)
            If Wv.Id > 0 Then
                cmbShowO.AddItem "Player " + Format(Wa + 1)
                cmbShowO.ItemData(cmbShowO.NewIndex) = Wa
                cmbOwner.AddItem "Player " + Format(Wa + 1)
                cmbOwner.ItemData(cmbOwner.NewIndex) = Wa
            End If
        Next Wa
        cmbShowO.ListIndex = 0
        
        cmbShowC.AddItem "<All>": cmbShowC.ItemData(0) = -1
        cmbShowC.AddItem "Rail": cmbShowC.ItemData(1) = &H10
        cmbShowC.AddItem "Road": cmbShowC.ItemData(2) = &H11
        cmbShowC.AddItem "Ship": cmbShowC.ItemData(3) = &H12
        cmbShowC.AddItem "Aircraft": cmbShowC.ItemData(4) = &H13
        cmbShowC.ListIndex = 0
        
        For Wa = 0 To UBound(CargoTypes)
            If CargoTypes(Wa) > ">" Then
                cmbCargo.AddItem CargoTypes(Wa)
                cmbCargo.ItemData(cmbCargo.NewIndex) = Wa
            End If
        Next Wa
        
        ShowList
        Frame2.Enabled = True
    Else
        Frame2.Enabled = False
    End If
    For Wa = 0 To frmDta.UBound: frmDta(Wa).Enabled = False: Next Wa
    CurItmNo = -1
    fInit = False
End Sub

Private Sub cmbCargo_Click()
    CurItm.CargoT = cmbCargo.ItemData(cmbCargo.ListIndex)
    MarkGame 16
End Sub

Private Sub cmbShowC_Click()
    If Not fInit Then ShowList
End Sub

Private Sub cmbShowO_Click()
    If Not fInit Then ShowList
End Sub

Private Sub Form_Load()
    Subclass
    UpdateInfo
End Sub

Public Sub PrepSave()
    If CurItmNo > -1 Then TTDXPutVeh CurItm
End Sub

Private Sub ShowList()
    Dim Wa As Long, Wb As Long, Wc As Integer, Wva As TreeViewItem, Wva2 As TreeViewItem
    Dim OK As Boolean
    
    Screen.MousePointer = 11
    
    tvVeh.TreeItems.RemoveAll
    tvVeh.SortItems sobNone
    
    If CurFile > " " Then
        Wb = (CLng(TTDXGeneralInfo.VehSize) * 850) - 1
        For Wa = 0 To Wb
            CurItm = TTDXGetVeh(Wa): Wc = cmbShowC.ItemData(cmbShowC.ListIndex)
            If (CurItm.Class < &H14) And ((CurItm.Class = Wc) Or (Wc = -1)) Then
                Wc = cmbShowO.ItemData(cmbShowO.ListIndex)
                If (CurItm.Class > 0) And ((CurItm.Owner = Wc) Or (Wc = -1)) Then  ' removed And (CurItm.SubClass = 0)
                    If CurItm.Class = &H13 Then
                        If CurItm.Subclass <= 2 Then
                            OK = True
                        Else
                            OK = False
                        End If
                    Else
                        If CurItm.Subclass <> 0 Then
                            OK = False
                        Else
                            OK = True
                        End If
                    End If
                    
                    If OK = True Then
                        Set Wva = tvVeh.TreeItems.Add(CurItm.Name + " (" + Format(Wa + 1) + ")", , , heNo, , , , Wa)
                        Set Wva2 = Wva
                        While CurItm.Next < &HFFFF&
                            CurItm = TTDXGetVeh(CurItm.Next)
                            'If CurItm.SubClass > 0 Then
                                Set Wva2 = Wva.SubItems.Add(CurItm.Name + " (" + Format(CurItm.Number + 1) + ")", Wva, , heNo, , , , CurItm.Number)
                                Wva.HasExpando = heYes
                            'Else
                                'Wa = 1.23456789054646E+21
                            'End If
                        Wend
                    End If
                    
                    OK = False
                End If
            End If
        Next Wa
    End If
    
    tvVeh.SortOrder = soAscending
    tvVeh.SortItems sobText
    
    Screen.MousePointer = 0
End Sub

Private Sub UpdateFields()
    Dim Wa As Integer
    'txtMoney.Text = CurItm.Money
    'txtDebt.Text = CurItm.Debt
    For Wa = 0 To cmbCargo.ListCount - 1
        If cmbCargo.ItemData(Wa) = CurItm.CargoT Then cmbCargo.ListIndex = Wa: Exit For
    Next Wa
    For Wa = 0 To cmbOwner.ListCount - 1
        If cmbOwner.ItemData(Wa) = CurItm.Owner Then cmbOwner.ListIndex = Wa: Exit For
    Next Wa
    frmTechInfo.ShowInfo 5, CurItm.Name, CurItm.Offset
    sliMaxLoad.CurrentPosition = CurItm.CargoMax
    sliRdr.CurrentPosition = CurItm.RelDropRate
    sliRel.CurrentPosition = CurItm.Rel
    If CurItm.Subclass <> 0 Then frmDta(1).Enabled = False
    txtAge.Text = CStr(CurItm.Age)
    txtAgeMax.Text = CStr(CurItm.AgeMax)
    txtVal.Text = CStr(CurItm.Value)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PrepSave
    UnSubclassWindow Me.hwnd, EnumSubclassID.escidVehicle
End Sub

Private Sub sliMaxLoad_PositionChanged(ByVal changeType As TrackBarCtlLibUCtl.PositionChangeTypeConstants, ByVal newPosition As Long)
    labCMax.Caption = "Max Units (" + Format(newPosition) + ")"
    CurItm.CargoMax = newPosition
    MarkGame 16
End Sub

Private Sub sliRdr_PositionChanged(ByVal changeType As TrackBarCtlLibUCtl.PositionChangeTypeConstants, ByVal newPosition As Long)
    CurItm.RelDropRate = newPosition
    labRel.Caption = "&Reliability: " + Format(Fix(sliRel.CurrentPosition / 2.55)) + "%" + " Drop Rate " + Format(CurItm.RelDropRate)
    MarkGame 16
End Sub

Private Sub sliRel_PositionChanged(ByVal changeType As TrackBarCtlLibUCtl.PositionChangeTypeConstants, ByVal newPosition As Long)
    CurItm.Rel = CByte(newPosition)
    labRel.Caption = "&Reliability: " + Format(Fix(newPosition / 2.55)) + "%" + " Drop Rate " + Format(CurItm.RelDropRate)
    MarkGame 16
End Sub

Private Sub tvVeh_CaretChanged(ByVal previousCaretItem As ExTVwLibUCtl.ITreeViewItem, ByVal newCaretItem As ExTVwLibUCtl.ITreeViewItem, ByVal caretChangeReason As ExTVwLibUCtl.CaretChangeCausedByConstants)
    Dim Wa As Integer, Wb As Long
    If CurItmNo > -1 Then TTDXPutVeh CurItm
    Wb = newCaretItem.ItemData
    CurItm = TTDXGetVeh(Wb)
    CurItmNo = Wb
    For Wa = 0 To frmDta.UBound: frmDta(Wa).Enabled = True: Next Wa
    UpdateFields
End Sub

Private Sub txtAge_Change()
    If jBetween(-1, Val(txtAge.Text), 65536) Then
        CurItm.Age = Val(txtAge.Text)
    Else
        txtAge.Text = CStr(CurItm.Age)
    End If
    txtAge.ToolTipText = CStr(Fix(CurItm.Age / 364.25)) + " Years"
    MarkGame 16
End Sub
Private Sub txtAge_KeyPress(KeyAscii As Integer)
    KeyAscii = CheckNumInput(txtAge, KeyAscii)
End Sub

Private Sub txtAgeMax_Change()
    If jBetween(-1, Val(txtAgeMax.Text), 65536) Then
        CurItm.AgeMax = Val(txtAgeMax.Text)
    Else
        txtAgeMax.Text = CStr(CurItm.AgeMax)
    End If
    txtAgeMax.ToolTipText = CStr(Fix(CurItm.AgeMax / 364.25)) + " Years"
    MarkGame 16
End Sub
Private Sub txtAgeMax_KeyPress(KeyAscii As Integer)
    KeyAscii = CheckNumInput(txtAgeMax, KeyAscii)
End Sub

Private Sub txtVal_Change()
    If jBetween(-1, Val(txtVal.Text), 2000000000) Then
        CurItm.Value = Val(txtVal.Text)
    Else
        txtVal.Text = CStr(CurItm.Value)
    End If
    
    MarkGame 16
End Sub
Private Sub txtVal_KeyPress(KeyAscii As Integer)
    KeyAscii = CheckNumInput(txtVal, KeyAscii)
End Sub
