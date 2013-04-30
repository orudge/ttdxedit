VERSION 5.00
Begin VB.Form frmFinances 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Finances"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9855
   Icon            =   "finances.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   9855
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstPlayers 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      ItemData        =   "finances.frx":0442
      Left            =   120
      List            =   "finances.frx":0444
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblThisYearTrainIncome 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7800
      TabIndex        =   32
      Top             =   3240
      Width           =   45
   End
   Begin VB.Label lblLastYearTrainIncome 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5520
      TabIndex        =   31
      Top             =   3240
      Width           =   45
   End
   Begin VB.Label lblTwoYearsAgoTrainIncome 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   30
      Top             =   3240
      Width           =   45
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Train Income:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1560
      TabIndex        =   29
      Top             =   3240
      Width           =   1680
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblThisYearPropertyMaintenance 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7800
      TabIndex        =   28
      Top             =   2880
      Width           =   45
   End
   Begin VB.Label lblLastYearPropertyMaintenance 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5520
      TabIndex        =   27
      Top             =   2880
      Width           =   45
   End
   Begin VB.Label lblTwoYearsAgoPropertyMaintenance 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   26
      Top             =   2880
      Width           =   45
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Property Maintenance:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1560
      TabIndex        =   25
      Top             =   2760
      Width           =   1680
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblThisYearShipRunningCosts 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7800
      TabIndex        =   24
      Top             =   2520
      Width           =   45
   End
   Begin VB.Label lblLastYearShipRunningCosts 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5520
      TabIndex        =   23
      Top             =   2520
      Width           =   45
   End
   Begin VB.Label lblTwoYearsAgoShipRunningCosts 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   22
      Top             =   2520
      Width           =   45
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Ship Running Costs:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1560
      TabIndex        =   21
      Top             =   2520
      Width           =   1800
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Aircraft Running Costs:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1560
      TabIndex        =   36
      Top             =   2040
      Width           =   1800
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTwoYearsAgoAircraftRunningCosts 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   35
      Top             =   2160
      Width           =   45
   End
   Begin VB.Label lblLastYearAircraftRunningCosts 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5520
      TabIndex        =   34
      Top             =   2160
      Width           =   45
   End
   Begin VB.Label lblThisYearAircraftRunningCosts 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7800
      TabIndex        =   33
      Top             =   2160
      Width           =   45
   End
   Begin VB.Label lblThisYearRoadVehRunningCosts 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7800
      TabIndex        =   20
      Top             =   1680
      Width           =   45
   End
   Begin VB.Label lblLastYearRoadVehRunningCosts 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5520
      TabIndex        =   19
      Top             =   1680
      Width           =   45
   End
   Begin VB.Label lblTwoYearsAgoRoadVehRunningCosts 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   18
      Top             =   1680
      Width           =   45
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Road Vehicle Running Costs:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1560
      TabIndex        =   17
      Top             =   1560
      Width           =   1800
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Train Running Costs:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1560
      TabIndex        =   16
      Top             =   1200
      Width           =   1800
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTwoYearsAgoTrainRunningCosts 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   15
      Top             =   1200
      Width           =   45
   End
   Begin VB.Label lblLastYearTrainRunningCosts 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5520
      TabIndex        =   14
      Top             =   1200
      Width           =   45
   End
   Begin VB.Label lblThisYearTrainRunningCosts 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7800
      TabIndex        =   13
      Top             =   1200
      Width           =   45
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "New Vehicles:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1560
      TabIndex        =   12
      Top             =   840
      Width           =   1125
   End
   Begin VB.Label lblTwoYearsAgoNewVehicles 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   11
      Top             =   840
      Width           =   45
   End
   Begin VB.Label lblLastYearNewVehicles 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5520
      TabIndex        =   10
      Top             =   840
      Width           =   45
   End
   Begin VB.Label lblThisYearNewVehicles 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7800
      TabIndex        =   9
      Top             =   840
      Width           =   45
   End
   Begin VB.Label lblThisYearConstruction 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7800
      TabIndex        =   8
      Top             =   480
      Width           =   45
   End
   Begin VB.Label lblLastYearConstruction 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5520
      TabIndex        =   7
      Top             =   480
      Width           =   45
   End
   Begin VB.Label lblTwoYearsAgoConstruction 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   6
      Top             =   480
      Width           =   45
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "This Year"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   7800
      TabIndex        =   5
      Top             =   120
      Width           =   780
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Last Year"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5520
      TabIndex        =   4
      Top             =   120
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Two Years Ago"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Construction:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   1125
   End
   Begin VB.Label lblPlayer 
      AutoSize        =   -1  'True
      Caption         =   "&Player:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmFinances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CurItm As TTDXplayer, CurItmNo As Integer

Private Sub UpdateFields()
    Dim Wa As Integer
    'txtMoney.Text = CurItm.Money
    'txtDebt.Text = CurItm.Debt
    lblTwoYearsAgoConstruction.Caption = FormatMoney(CurItm.FinancesTwoYearsAgo.Construction)
    lblLastYearConstruction.Caption = FormatMoney(CurItm.FinancesLastYear.Construction)
    lblThisYearConstruction.Caption = FormatMoney(CurItm.FinancesThisYear.Construction)
    
    lblTwoYearsAgoNewVehicles.Caption = FormatMoney(CurItm.FinancesTwoYearsAgo.NewVehicles)
    lblLastYearNewVehicles.Caption = FormatMoney(CurItm.FinancesLastYear.NewVehicles)
    lblThisYearNewVehicles.Caption = FormatMoney(CurItm.FinancesThisYear.NewVehicles)
    
    lblTwoYearsAgoTrainRunningCosts.Caption = FormatMoney(CurItm.FinancesTwoYearsAgo.TrainRunningCosts)
    lblLastYearTrainRunningCosts.Caption = FormatMoney(CurItm.FinancesLastYear.TrainRunningCosts)
    lblThisYearTrainRunningCosts.Caption = FormatMoney(CurItm.FinancesThisYear.TrainRunningCosts)
    
    lblTwoYearsAgoRoadVehRunningCosts.Caption = FormatMoney(CurItm.FinancesTwoYearsAgo.RoadVehRunningCosts)
    lblLastYearRoadVehRunningCosts.Caption = FormatMoney(CurItm.FinancesLastYear.RoadVehRunningCosts)
    lblThisYearRoadVehRunningCosts.Caption = FormatMoney(CurItm.FinancesThisYear.RoadVehRunningCosts)
    
    lblTwoYearsAgoAircraftRunningCosts.Caption = FormatMoney(CurItm.FinancesTwoYearsAgo.AircraftRunningCosts)
    lblLastYearAircraftRunningCosts.Caption = FormatMoney(CurItm.FinancesLastYear.AircraftRunningCosts)
    lblThisYearAircraftRunningCosts.Caption = FormatMoney(CurItm.FinancesThisYear.AircraftRunningCosts)
    
    lblTwoYearsAgoShipRunningCosts.Caption = FormatMoney(CurItm.FinancesTwoYearsAgo.ShipRunningCosts)
    lblLastYearShipRunningCosts.Caption = FormatMoney(CurItm.FinancesLastYear.ShipRunningCosts)
    lblThisYearShipRunningCosts.Caption = FormatMoney(CurItm.FinancesThisYear.ShipRunningCosts)
    
    lblTwoYearsAgoPropertyMaintenance.Caption = FormatMoney(CurItm.FinancesTwoYearsAgo.PropertyMaintenance)
    lblLastYearPropertyMaintenance.Caption = FormatMoney(CurItm.FinancesLastYear.PropertyMaintenance)
    lblThisYearPropertyMaintenance.Caption = FormatMoney(CurItm.FinancesThisYear.PropertyMaintenance)
    
    lblTwoYearsAgoTrainIncome.Caption = FormatMoney(CurItm.FinancesTwoYearsAgo.TrainIncome)
    lblLastYearTrainIncome.Caption = FormatMoney(CurItm.FinancesLastYear.TrainIncome)
    lblThisYearTrainIncome.Caption = FormatMoney(CurItm.FinancesThisYear.TrainIncome)
    
    'frmMDI.ShowDebug "Player " + CStr(CurItm.Number + 1) + Chr(10) + "Offset: " + Hex(CurItm.Offset), CurItm.DebugData, CurItm.DebugData2
    frmTechInfo.ShowInfo 1, "Player " + CStr(CurItm.Number + 1), CurItm.Offset
End Sub
Private Sub Form_Load()
    Dim Wa As Integer, Wv As TTDXplayer
    
    lstPlayers.Clear
    If CurFile > " " Then
        For Wa = 0 To 7
            CurItm = TTDXPlayerInfo(Wa)
            If CurItm.Id > 0 Then
                lstPlayers.AddItem "Player " + Format(Wa + 1)
                lstPlayers.ItemData(lstPlayers.NewIndex) = Wa
            End If
        Next Wa
    End If
    CurItmNo = -1
End Sub


Private Sub lstPlayers_Click()
    Dim Wa As Integer, Wb As Integer
'    If CurItmNo > -1 Then TTDXputPlayer CurItm
    Wb = lstPlayers.ItemData(lstPlayers.ListIndex)
    CurItm = TTDXPlayerInfo(Wb)
    CurItmNo = Wb
    UpdateFields
    'For Wa = 0 To frmDta.UBound: frmDta(Wa).Enabled = True: Next Wa
End Sub


