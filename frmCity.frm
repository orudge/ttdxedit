VERSION 5.00
Object = "{956B5A46-C53F-45A7-AF0E-EC2E1CC9B567}#1.5#0"; "TrackBarCtlU.ocx"
Begin VB.Form frmCity 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cities"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7575
   Icon            =   "frmCity.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmDta 
      Caption         =   "Ratings"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Index           =   1
      Left            =   3720
      TabIndex        =   6
      Top             =   1320
      Width           =   3735
      Begin TrackBarCtlLibUCtl.TrackBar sliRat 
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   17
         Top             =   240
         Width           =   2775
         _cx             =   4895
         _cy             =   450
         Appearance      =   0
         AutoTickFrequency=   100
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
         LargeStepWidth  =   50
         Maximum         =   999
         Minimum         =   -999
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
      Begin VB.CheckBox chkRatE 
         Caption         =   "Pl. 8"
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
         Index           =   7
         Left            =   120
         TabIndex        =   14
         Top             =   2760
         Width           =   615
      End
      Begin VB.CheckBox chkRatE 
         Caption         =   "Pl. 7"
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
         Index           =   6
         Left            =   120
         TabIndex        =   13
         Top             =   2400
         Width           =   615
      End
      Begin VB.CheckBox chkRatE 
         Caption         =   "Pl. 6"
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
         Index           =   5
         Left            =   120
         TabIndex        =   12
         Top             =   2040
         Width           =   615
      End
      Begin VB.CheckBox chkRatE 
         Caption         =   "Pl. 5"
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
         Index           =   4
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   615
      End
      Begin VB.CheckBox chkRatE 
         Caption         =   "Pl. 4"
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
         Index           =   3
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   615
      End
      Begin VB.CheckBox chkRatE 
         Caption         =   "Pl. 3"
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
         Index           =   2
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   735
      End
      Begin VB.CheckBox chkRatE 
         Caption         =   "Pl. 2"
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
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   615
      End
      Begin VB.CheckBox chkRatE 
         Caption         =   "Pl. 1"
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
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin TrackBarCtlLibUCtl.TrackBar sliRat 
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   19
         Top             =   600
         Width           =   2775
         _cx             =   4895
         _cy             =   450
         Appearance      =   0
         AutoTickFrequency=   100
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
         LargeStepWidth  =   50
         Maximum         =   999
         Minimum         =   -999
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
      Begin TrackBarCtlLibUCtl.TrackBar sliRat 
         Height          =   255
         Index           =   2
         Left            =   840
         TabIndex        =   20
         Top             =   960
         Width           =   2775
         _cx             =   4895
         _cy             =   450
         Appearance      =   0
         AutoTickFrequency=   100
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
         LargeStepWidth  =   50
         Maximum         =   999
         Minimum         =   -999
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
      Begin TrackBarCtlLibUCtl.TrackBar sliRat 
         Height          =   255
         Index           =   3
         Left            =   840
         TabIndex        =   21
         Top             =   1320
         Width           =   2775
         _cx             =   4895
         _cy             =   450
         Appearance      =   0
         AutoTickFrequency=   100
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
         LargeStepWidth  =   50
         Maximum         =   999
         Minimum         =   -999
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
      Begin TrackBarCtlLibUCtl.TrackBar sliRat 
         Height          =   255
         Index           =   4
         Left            =   840
         TabIndex        =   22
         Top             =   1680
         Width           =   2775
         _cx             =   4895
         _cy             =   450
         Appearance      =   0
         AutoTickFrequency=   100
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
         LargeStepWidth  =   50
         Maximum         =   999
         Minimum         =   -999
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
      Begin TrackBarCtlLibUCtl.TrackBar sliRat 
         Height          =   255
         Index           =   5
         Left            =   840
         TabIndex        =   23
         Top             =   2040
         Width           =   2775
         _cx             =   4895
         _cy             =   450
         Appearance      =   0
         AutoTickFrequency=   100
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
         LargeStepWidth  =   50
         Maximum         =   999
         Minimum         =   -999
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
      Begin TrackBarCtlLibUCtl.TrackBar sliRat 
         Height          =   255
         Index           =   6
         Left            =   840
         TabIndex        =   24
         Top             =   2400
         Width           =   2775
         _cx             =   4895
         _cy             =   450
         Appearance      =   0
         AutoTickFrequency=   100
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
         LargeStepWidth  =   50
         Maximum         =   999
         Minimum         =   -999
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
      Begin TrackBarCtlLibUCtl.TrackBar sliRat 
         Height          =   255
         Index           =   7
         Left            =   840
         TabIndex        =   18
         Top             =   2760
         Width           =   2775
         _cx             =   4895
         _cy             =   450
         Appearance      =   0
         AutoTickFrequency=   100
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
         LargeStepWidth  =   50
         Maximum         =   999
         Minimum         =   -999
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
   End
   Begin VB.Frame frmDta 
      Caption         =   "Population"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   5520
      TabIndex        =   15
      Top             =   120
      Width           =   1935
      Begin VB.TextBox txtPop 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame frmDta 
      Caption         =   "Location"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   3720
      TabIndex        =   1
      Top             =   120
      Width           =   1815
      Begin VB.TextBox txtX 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   3
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtY 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "X:"
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
         Left            =   120
         TabIndex        =   5
         Top             =   285
         Width           =   255
      End
      Begin VB.Label Label2 
         Caption         =   "Y:"
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
         Left            =   960
         TabIndex        =   4
         Top             =   285
         Width           =   255
      End
   End
   Begin VB.ListBox lstCities 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4350
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmCity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CurItm As TTDXCitInfo, CurItmNo

Public Sub UpdateInfo()
    Dim Wa As Integer, Wv As TTDXplayer
    
    lstCities.Clear
    If CurFile > " " Then
        For Wa = 0 To 69
            CurItm = CityInfo(Wa)
            If CInt(CurItm.X) + CInt(CurItm.Y) > 0& Then
                lstCities.AddItem CurItm.Name + " (" + Format(Wa, "00") + ") "
                lstCities.ItemData(lstCities.NewIndex) = Wa
            End If
        Next Wa
        For Wa = 0 To 7
            Wv = TTDXPlayerInfo(Wa)
            If Wv.Id > 0 Then
                chkRatE(Wa).Enabled = True
            Else
                chkRatE(Wa).Enabled = False
            End If
        Next Wa
    End If
    For Wa = 0 To frmDta.UBound: frmDta(Wa).Enabled = False: Next Wa
    CurItmNo = lstCities.ListIndex
End Sub
Public Sub PrepSave()
    If CurItmNo > -1 Then TTDXCityPut CurItm
End Sub
Private Sub chkRatE_Click(Index As Integer)
    If chkRatE(Index).Value = 1 Then
        sliRat(Index).Enabled = True
    Else
        sliRat(Index).Enabled = False
    End If
    sliRat(Index).CurrentPosition = CurItm.CRate(Index)
    
    MarkGame 2
End Sub

Private Sub Form_Load()
    UpdateInfo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PrepSave
End Sub

Private Sub lstCities_Click()
    Dim Wa As Integer, Wb As Integer
    If CurItmNo > -1 Then TTDXCityPut CurItm
    Wb = lstCities.ItemData(lstCities.ListIndex)
    CurItm = CityInfo(Wb)
    CurItmNo = Wb
    UpdateFields
    For Wa = 0 To frmDta.UBound: frmDta(Wa).Enabled = True: Next Wa
End Sub
Private Sub UpdateFields()
    Dim Wa As Integer
    txtX.Text = CStr(CurItm.X)
    txtY.Text = CStr(CurItm.Y)
    txtPop.Text = CStr(CurItm.Population)
    For Wa = 0 To 7
        If CurItm.CRateE(Wa) Then chkRatE(Wa).Value = 1 Else: chkRatE(Wa).Value = 0
        chkRatE_Click Wa
    Next Wa
    
    frmTechInfo.ShowInfo 2, CurItm.Name, CurItm.Offset
    frmMap.SetHighlight CInt(CurItm.X), CInt(CurItm.Y)
End Sub


Private Sub sliRat_PositionChanged(Index As Integer, ByVal changeType As TrackBarCtlLibUCtl.PositionChangeTypeConstants, ByVal newPosition As Long)
    CurItm.CRate(Index) = newPosition
    MarkGame 2
End Sub

