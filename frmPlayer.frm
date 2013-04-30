VERSION 5.00
Begin VB.Form frmPlayer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Players"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4695
   Icon            =   "frmPlayer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmDta 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   2760
      TabIndex        =   5
      Top             =   1560
      Width           =   1815
      Begin VB.CommandButton cmdRemHQ 
         Caption         =   "Remove HQ"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame frmDta 
      Caption         =   "Debt (£)"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   2760
      TabIndex        =   3
      Top             =   840
      Width           =   1815
      Begin VB.TextBox txtDebt 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame frmDta 
      Caption         =   "Money (£)"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   2760
      TabIndex        =   1
      Top             =   120
      Width           =   1815
      Begin VB.TextBox txtMoney 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.ListBox lstPlay 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frmPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private CurItm As TTDXplayer, CurItmNo As Integer

Public Sub UpdateInfo()
    Dim Wa As Integer, Wv As TTDXplayer
    
    lstPlay.Clear
    If CurFile > " " Then
        For Wa = 0 To 7
            CurItm = TTDXPlayerInfo(Wa)
            If CurItm.Id > 0 Then
                lstPlay.AddItem "Player " + Format(Wa + 1)
                lstPlay.ItemData(lstPlay.NewIndex) = Wa
            End If
        Next Wa
    End If
    For Wa = 0 To frmDta.UBound: frmDta(Wa).Enabled = False: Next Wa
    CurItmNo = -1
End Sub

Private Sub cmdRemHQ_Click()
    Dim Wx As Integer, Wy As Integer, Wva As TTDXlandscape
    For Wx = CurItm.HQx To CurItm.HQx + 1
        For Wy = CurItm.HQy To CurItm.HQy + 1
            Wva = TTDXgetLandscape(Wx, Wy)
            Wva.Object = 0: Wva.Owner = &H10: Wva.L5 = 0
            TTDXputLandscape Wva
        Next Wy
    Next Wx
    CurItm.HQx = 255: CurItm.HQy = 255
    
    MarkGame 1
End Sub

Private Sub Form_Load()
    UpdateInfo
End Sub

Public Sub PrepSave()
    If CurItmNo > -1 Then TTDXputPlayer CurItm
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PrepSave
End Sub

Private Sub lstplay_Click()
    Dim Wa As Integer, Wb As Integer
    If CurItmNo > -1 Then TTDXputPlayer CurItm
    Wb = lstPlay.ItemData(lstPlay.ListIndex)
    CurItm = TTDXPlayerInfo(Wb)
    CurItmNo = Wb
    UpdateFields
    For Wa = 0 To frmDta.UBound: frmDta(Wa).Enabled = True: Next Wa
End Sub
Private Sub UpdateFields()
    Dim Wa As Integer
    txtMoney.Text = CurItm.Money
    txtDebt.Text = CurItm.Debt
    'frmMDI.ShowDebug "Player " + CStr(CurItm.Number + 1) + Chr(10) + "Offset: " + Hex(CurItm.Offset), CurItm.DebugData, CurItm.DebugData2
    frmTechInfo.ShowInfo 1, "Player " + CStr(CurItm.Number + 1), CurItm.Offset
End Sub

Private Sub txtDebt_Change()
    If jBetween(-1, Val(txtDebt.Text), 2000000000) Then
        CurItm.Debt = Val(txtDebt.Text)
    Else
        txtDebt = CStr(CurItm.Debt)
    End If

    MarkGame 1
End Sub

Private Sub txtDebt_KeyPress(KeyAscii As Integer)
    KeyAscii = CheckNumInput(txtDebt, KeyAscii)
End Sub

Private Sub txtMoney_Change()
    If jBetween(-1000000, Val(txtMoney.Text), 2000000000) Then
        CurItm.Money = Val(txtMoney.Text)
    Else
        txtMoney = CStr(CurItm.Money)
    End If
    
    MarkGame 1
End Sub

Private Sub txtMoney_KeyPress(KeyAscii As Integer)
    KeyAscii = CheckNumInput(txtMoney, KeyAscii)
End Sub

