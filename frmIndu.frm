VERSION 5.00
Begin VB.Form frmIndu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Industries"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7425
   Icon            =   "frmIndu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrmDta 
      Caption         =   "HomeTown"
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
      Top             =   840
      Width           =   1815
      Begin VB.ComboBox cmbTown 
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
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame FrmDta 
      Caption         =   "Accepts"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   3
      Left            =   3600
      TabIndex        =   18
      Top             =   840
      Width           =   1815
      Begin VB.ComboBox cmbDel 
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
         Index           =   2
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox cmbDel 
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
         Index           =   1
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox cmbDel 
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
         Index           =   0
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame FrmDta 
      Caption         =   "Production Type And Rate"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Index           =   4
      Left            =   3600
      TabIndex        =   10
      Top             =   2400
      Width           =   3615
      Begin VB.TextBox txtProR 
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
         Height          =   315
         Index           =   1
         Left            =   1920
         TabIndex        =   14
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtProR 
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
         Height          =   315
         Index           =   0
         Left            =   1920
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.ComboBox cmbProd 
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
         Index           =   1
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   600
         Width           =   1695
      End
      Begin VB.ComboBox cmbProd 
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
         Index           =   0
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblRate 
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
         Left            =   2520
         TabIndex        =   27
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblRate 
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
         Left            =   2520
         TabIndex        =   26
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   $"frmIndu.frx":0442
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
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   3375
      End
   End
   Begin VB.Frame FrmDta 
      Caption         =   "Size"
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
      Index           =   1
      Left            =   5520
      TabIndex        =   3
      Top             =   120
      Width           =   1815
      Begin VB.TextBox txtH 
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
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtW 
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
         TabIndex        =   4
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label4 
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
         TabIndex        =   9
         Top             =   285
         Width           =   255
      End
      Begin VB.Label Label3 
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
         TabIndex        =   8
         Top             =   285
         Width           =   255
      End
   End
   Begin VB.Frame FrmDta 
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
      Left            =   3600
      TabIndex        =   0
      Top             =   120
      Width           =   1815
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
         TabIndex        =   1
         Top             =   240
         Width           =   495
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
         TabIndex        =   7
         Top             =   285
         Width           =   255
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
         TabIndex        =   6
         Top             =   285
         Width           =   255
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Active Industries"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   120
      TabIndex        =   22
      Top             =   120
      Width           =   3375
      Begin VB.ComboBox cmbShowT 
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
         TabIndex        =   25
         Top             =   240
         Width           =   1575
      End
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
         Left            =   1680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   240
         Width           =   1575
      End
      Begin VB.ListBox lstIndu 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3570
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmIndu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CurItm As TTDXIndInfo, CurItmNo As Integer


Public Sub UpdateInfo()
    Dim Wa As Integer, Wb As Integer, Wsa As String, Wva As TTDXCitInfo
    
    For Wa = 0 To 1: cmbProd(Wa).Clear: Next Wa
    For Wa = 0 To 2: cmbDel(Wa).Clear: Next Wa
    If CurFile > " " Then
        For Wa = 0 To frmDta.UBound: frmDta(Wa).Enabled = True: Next Wa
        cmbProd(0).AddItem "<Nothing>": cmbProd(0).ItemData(0) = 255
        cmbProd(1).AddItem "<Nothing>": cmbProd(1).ItemData(0) = 255
        For Wa = 0 To 2
            cmbDel(Wa).AddItem "<Nothing>": cmbDel(Wa).ItemData(0) = 255
        Next Wa
        cmbTown.Clear: cmbShowC.Clear
        cmbShowC.AddItem " <Show All>": cmbShowC.ItemData(0) = 255
        For Wa = 0 To 69
            If Cities(Wa) > "" Then
                cmbTown.AddItem Cities(Wa): cmbTown.ItemData(cmbTown.NewIndex) = Wa
                cmbShowC.AddItem Cities(Wa): cmbShowC.ItemData(cmbShowC.NewIndex) = Wa
            End If
        Next Wa
        For Wa = 0 To UBound(CargoTypes)
            If CargoTypes(Wa) > ">" Then
                For Wb = 0 To 1
                    cmbProd(Wb).AddItem CargoTypes(Wa): cmbProd(Wb).ItemData(cmbProd(Wb).NewIndex) = Wa
                Next Wb
                For Wb = 0 To 2
                    cmbDel(Wb).AddItem CargoTypes(Wa): cmbDel(Wb).ItemData(cmbDel(Wb).NewIndex) = Wa
                Next Wb
            End If
        Next Wa
        cmbShowT.Clear: cmbShowT.AddItem "<Show All>": cmbShowT.ItemData(0) = 255
        For Wa = 0 To UBound(IndustryTypes)
            If IndustryTypes(Wa) > ">" Then cmbShowT.AddItem IndustryTypes(Wa): cmbShowT.ItemData(cmbShowT.NewIndex) = Wa
        Next Wa
        Frame2.Enabled = True
    Else
        Frame2.Enabled = False
    End If
    ShowList
End Sub
Public Sub PrepSave()
    If CurItmNo > -1 Then TTDXIndustryPut CurItm
End Sub
Private Sub ShowList()
    '
    ' Display the wanted Industries in the list
    '
    Dim Wa As Integer, wXa As TTDXIndInfo
    
    lstIndu.Clear
    If CurFile > " " Then
        If cmbShowC.ListIndex = -1 Then cmbShowC.ListIndex = 0: Exit Sub
        If cmbShowT.ListIndex = -1 Then cmbShowT.ListIndex = 0: Exit Sub
        For Wa = 0 To 89
            wXa = TTDXIndustryInfo(Wa)
            If wXa.H + wXa.W > 0 Then
                If wXa.HomeTown = cmbShowC.ItemData(cmbShowC.ListIndex) Or cmbShowC.ListIndex = 0 Then
                    If wXa.Type = cmbShowT.ItemData(cmbShowT.ListIndex) Or cmbShowT.ListIndex = 0 Then
                        lstIndu.AddItem wXa.Name: lstIndu.ItemData(lstIndu.NewIndex) = Wa
                    End If
                End If
            End If
        Next Wa
    End If
    For Wa = 0 To frmDta.UBound: frmDta(Wa).Enabled = False: Next Wa
    CurItmNo = lstIndu.ListIndex
End Sub

Private Sub cmbDel_Click(Wb As Integer)
    If cmbDel(Wb).ListIndex > -1 Then
        CurItm.del(Wb) = cmbDel(Wb).ItemData(cmbDel(Wb).ListIndex)
    Else
        cmbDel(Wb).ListIndex = 0
    End If
    MarkGame 4
End Sub

Private Sub cmbProd_Click(Wb As Integer)
    If cmbProd(Wb).ListIndex > -1 Then
        CurItm.Prod(Wb) = cmbProd(Wb).ItemData(cmbProd(Wb).ListIndex)
        If cmbProd(Wb).ItemData(cmbProd(Wb).ListIndex) = 255 Then
            txtProR(Wb).Enabled = False
        Else
            txtProR(Wb).Enabled = True
        End If
    
        MarkGame 4
    Else
        cmbProd(Wb).ListIndex = 0
    End If
End Sub

Private Sub cmbShowC_Click()
    ShowList
End Sub

Private Sub cmbShowT_Click()
    ShowList
End Sub

Private Sub cmbTown_Click()
    If cmbTown.ListIndex > -1 Then
        CurItm.HomeTown = cmbTown.ItemData(cmbTown.ListIndex)
        MarkGame 4
    End If
End Sub

Private Sub Form_Load()
    UpdateInfo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PrepSave
End Sub

Private Sub lstIndu_Click()
    Dim Wa As Integer, Wb As Integer
    If CurItmNo > -1 Then TTDXIndustryPut CurItm
    Wb = lstIndu.ItemData(lstIndu.ListIndex)
    CurItm = TTDXIndustryInfo(Wb)
    CurItmNo = Wb
    UpdateFields
    For Wa = 0 To frmDta.UBound: frmDta(Wa).Enabled = True: Next Wa
End Sub

Private Sub UpdateFields()
    Dim Wa As Integer, Wb As Integer
    
    txtX.Text = Str(CurItm.X)
    txtY.Text = Str(CurItm.Y)
    txtW.Text = Str(CurItm.W)
    txtH.Text = Str(CurItm.H)
    For Wa = 0 To cmbProd(0).ListCount - 1
        If CurItm.Prod(0) = cmbProd(0).ItemData(Wa) Then cmbProd(0).ListIndex = Wa
        If CurItm.Prod(1) = cmbProd(1).ItemData(Wa) Then cmbProd(1).ListIndex = Wa
        For Wb = 0 To 2
            If CurItm.del(Wb) = cmbDel(Wb).ItemData(Wa) Then cmbDel(Wb).ListIndex = Wa
        Next Wb
    Next Wa
    For Wa = 0 To cmbTown.ListCount - 1
        If CurItm.HomeTown = cmbTown.ItemData(Wa) Then cmbTown.ListIndex = Wa
    Next Wa
    txtProR(0).Text = CStr(CurItm.ProR(0))
    txtProR(1).Text = CStr(CurItm.ProR(1))
    
    'frmMDI.ShowDebug "Industry: " + CurItm.Name + Chr(10) + "Offset: " + Hex(CurItm.Offset), CurItm.DebugData, CurItm.DebugData2
    frmTechInfo.ShowInfo 3, CurItm.Name, CurItm.Offset
    frmMap.SetHighlight CInt(CurItm.X), CInt(CurItm.Y)
End Sub


Private Sub txtProR_Change(Index As Integer)
    If jBetween(-1, Val(txtProR(Index).Text), 241) Then
        CurItm.ProR(Index) = Val(txtProR(Index))
        lblRate(Index).Caption = CStr(CurItm.ProR(Index) * 8) & " / " & CStr(CurItm.ProR(Index) * 9)
    Else
        txtProR(Index) = CStr(CurItm.ProR(Index))
        lblRate(Index).Caption = CStr(CurItm.ProR(Index) * 8) & " / " & CStr(CurItm.ProR(Index) * 9)
    End If
End Sub
Private Sub txtProR_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = CheckNumInput(txtProR(Index), KeyAscii)
End Sub

