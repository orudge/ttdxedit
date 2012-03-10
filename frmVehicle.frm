VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmVehicle 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vehicles"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6975
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
      Left            =   3750
      TabIndex        =   19
      Top             =   960
      Width           =   3135
      Begin VB.TextBox txtVal 
         Alignment       =   1  'Right Justify
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
         Left            =   1680
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Value (£):"
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
         TabIndex        =   21
         Top             =   270
         Width           =   1455
      End
   End
   Begin VB.Frame frmDta 
      Caption         =   "Engine Related"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   1
      Left            =   3750
      TabIndex        =   11
      Top             =   1560
      Width           =   3135
      Begin VB.TextBox txtAge 
         Alignment       =   1  'Right Justify
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
         Left            =   1680
         TabIndex        =   16
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtAgeMax 
         Alignment       =   1  'Right Justify
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
         Left            =   1680
         TabIndex        =   15
         Top             =   1440
         Width           =   1335
      End
      Begin MSComctlLib.Slider sliRDr 
         Height          =   255
         Left            =   30
         TabIndex        =   14
         Top             =   720
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   200
         SmallChange     =   40
         Max             =   4000
         TickFrequency   =   80
      End
      Begin MSComctlLib.Slider sliRel 
         Height          =   255
         Left            =   30
         TabIndex        =   12
         Top             =   450
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   25
         Max             =   255
         TickFrequency   =   26
      End
      Begin VB.Label Label2 
         Caption         =   "Age (Days):"
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
         TabIndex        =   18
         Top             =   1110
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Max Age (Days):"
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
         TabIndex        =   17
         Top             =   1470
         Width           =   1575
      End
      Begin VB.Label labRel 
         Caption         =   "Reliability"
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
         TabIndex        =   13
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Owner"
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
      Left            =   5280
      TabIndex        =   6
      Top             =   0
      Width           =   1575
      Begin VB.ComboBox cmbOwner 
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
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame frmDta 
      Caption         =   "Cargo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   0
      Left            =   3750
      TabIndex        =   4
      Top             =   3360
      Width           =   3135
      Begin MSComctlLib.Slider sliMaxLoad 
         Height          =   255
         Left            =   30
         TabIndex        =   8
         Top             =   840
         Width           =   3060
         _ExtentX        =   5398
         _ExtentY        =   450
         _Version        =   393216
         Max             =   600
         TickFrequency   =   12
      End
      Begin VB.ComboBox cmbCargo 
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
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Note: Some types of cargo have a bit odd values, like oil where 1 unit = 100 l."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         TabIndex        =   10
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label labCMax 
         Caption         =   "Max"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   630
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Active Vehicles"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.ComboBox cmbShowC 
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
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.ComboBox cmbShowO 
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
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
      Begin MSComctlLib.TreeView tvVeh 
         Height          =   4455
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   7858
         _Version        =   393217
         Indentation     =   441
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   6
         FullRowSelect   =   -1  'True
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
    wData(&H44BBD) = wData(&H44BBD) Or 16
End Sub

Private Sub cmbShowC_Click()
    If Not fInit Then ShowList
End Sub

Private Sub cmbShowO_Click()
    If Not fInit Then ShowList
End Sub

Private Sub Form_Load()
    UpdateInfo
End Sub

Public Sub PrepSave()
    If CurItmNo > -1 Then TTDXPutVeh CurItm
End Sub

Private Sub ShowList()
    Dim Wa As Long, Wb As Long, Wc As Integer, Wva As Node, Wva2 As Node
    Dim OK As Boolean
    
    Screen.MousePointer = 11
    tvVeh.Nodes.Clear
    If CurFile > " " Then
        Wb = (CLng(TTDXGeneralInfo.VehSize) * 850) - 1
        For Wa = 0 To Wb
            CurItm = TTDXGetVeh(Wa): Wc = cmbShowC.ItemData(cmbShowC.ListIndex)
            If (CurItm.Class < &H14) And ((CurItm.Class = Wc) Or (Wc = -1)) Then
                Wc = cmbShowO.ItemData(cmbShowO.ListIndex)
                If (CurItm.Class > 0) And ((CurItm.Owner = Wc) Or (Wc = -1)) Then  ' removed And (CurItm.SubClass = 0)
                    If CurItm.Class = &H13 Then
                        If CurItm.SubClass <= 2 Then
                            OK = True
                        Else
                            OK = False
                        End If
                    Else
                        If CurItm.SubClass <> 0 Then
                            OK = False
                        Else
                            OK = True
                        End If
                    End If
                    
                    If OK = True Then
                        Set Wva = tvVeh.Nodes.Add(, , "x" + Format(Wa, "00000"), CurItm.Name + " (" + Format(Wa + 1) + ")")
                        Wva.Sorted = False
                        Set Wva2 = Wva
                        While CurItm.Next < &HFFFF&
                            CurItm = TTDXGetVeh(CurItm.Next)
                            'If CurItm.SubClass > 0 Then
                                Set Wva2 = tvVeh.Nodes.Add(Wva.Index, 4, "x" + Format(CurItm.Number, "00000"), CurItm.Name + " (" + Format(CurItm.Number + 1) + ")")
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
    sliMaxLoad.Value = CurItm.CargoMax: sliMaxLoad_Change
    sliRDr.Value = CurItm.RelDropRate: sliRDr_Change
    sliRel.Value = CurItm.Rel: sliRel_Change
    If CurItm.SubClass <> 0 Then frmDta(1).Enabled = False
    txtAge.Text = CStr(CurItm.Age)
    txtAgeMax.Text = CStr(CurItm.AgeMax)
    txtVal.Text = CStr(CurItm.Value)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PrepSave
End Sub

Private Sub sliMaxLoad_Change()
    labCMax.Caption = "Max Units (" + Format(sliMaxLoad.Value) + ")"
    CurItm.CargoMax = sliMaxLoad.Value
    wData(&H44BBD) = wData(&H44BBD) Or 16
End Sub

Private Sub sliRDr_Change()
    CurItm.RelDropRate = sliRDr.Value
    labRel.Caption = "Reliability: " + Format(Fix(sliRel.Value / 2.55)) + "%" + " Droprate " + Format(CurItm.RelDropRate)
    wData(&H44BBD) = wData(&H44BBD) Or 16
End Sub

Private Sub sliRel_Change()
    CurItm.Rel = CByte(sliRel.Value)
    labRel.Caption = "Reliability: " + Format(Fix(sliRel.Value / 2.55)) + "%" + " Droprate " + Format(CurItm.RelDropRate)
    wData(&H44BBD) = wData(&H44BBD) Or 16
End Sub

Private Sub tvVeh_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim Wa As Integer, Wb As Long
    If CurItmNo > -1 Then TTDXPutVeh CurItm
    Wb = Val(MID(Node.Key, 2))
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
    wData(&H44BBD) = wData(&H44BBD) Or 16
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
    wData(&H44BBD) = wData(&H44BBD) Or 16
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
    
    wData(&H44BBD) = wData(&H44BBD) Or 16
End Sub
Private Sub txtVal_KeyPress(KeyAscii As Integer)
    KeyAscii = CheckNumInput(txtVal, KeyAscii)
End Sub
