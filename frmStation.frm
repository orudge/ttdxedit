VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmStation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stations"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7215
   Icon            =   "frmStation.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmDta 
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
      Index           =   3
      Left            =   3480
      TabIndex        =   16
      Top             =   5040
      Width           =   3615
      Begin VB.CommandButton cmdRemStop 
         Caption         =   "Add Rem. Part"
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
         Left            =   1440
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdAddHouse 
         Caption         =   "Add Buildings"
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
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Stations"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   0
      TabIndex        =   13
      Top             =   120
      Width           =   3375
      Begin VB.ListBox lstStation 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4740
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   15
         Top             =   600
         Width           =   3135
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
         TabIndex        =   14
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame frmDta 
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
      Index           =   2
      Left            =   5280
      TabIndex        =   10
      Top             =   120
      Width           =   1815
      Begin VB.TextBox txtOwner 
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
         TabIndex        =   11
         Top             =   240
         Width           =   1575
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
      Height          =   4095
      Index           =   1
      Left            =   3480
      TabIndex        =   5
      Top             =   840
      Width           =   3615
      Begin VB.TextBox txtCam 
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
         Index           =   0
         Left            =   2640
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin MSComctlLib.Slider sliCrate 
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   12
         Top             =   360
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   26
         Max             =   255
         TickFrequency   =   26
      End
      Begin VB.CheckBox chkChasrate 
         Caption         =   "Check1"
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
         TabIndex        =   6
         Top             =   375
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Amount:"
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
         Left            =   2640
         TabIndex        =   9
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Rating:"
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
         Left            =   1560
         TabIndex        =   8
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.Frame frmDta 
      Caption         =   "Location (Sign)"
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
      Left            =   3480
      TabIndex        =   0
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
         TabIndex        =   2
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
         TabIndex        =   1
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   285
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CurItmNo As Integer, CurItm As TTDXStation

Public Sub UpdateInfo()
    Dim Wa As Integer, Wb As Integer, Wsa As String, Wv As TTDXplayer
    
    CurItmNo = -1
    If CurFile > " " Then
        Wb = chkChasrate(0).Top
        cmbShowO.Clear
        cmbShowO.AddItem "<All>"
        cmbShowO.ItemData(cmbShowO.NewIndex) = 255
        cmbShowO.AddItem "<No One>"
        cmbShowO.ItemData(cmbShowO.NewIndex) = &H10
        
        For Wa = 0 To 7
            Wv = TTDXPlayerInfo(Wa)
            If Wv.Id > 0 Then
                cmbShowO.AddItem "Player " + Format(Wa + 1)
                cmbShowO.ItemData(cmbShowO.NewIndex) = Wa
            End If
        Next Wa
        For Wa = 0 To 11
            If CargoTypes(Wa) > "<z" Then
                sliCrate(Wa).Top = Wb: sliCrate(Wa).Visible = True
                txtCam(Wa).Top = Wb - 15: txtCam(Wa).Visible = True
                chkChasrate(Wa).Top = Wb: Wb = Wb + 300
                chkChasrate(Wa).Visible = True
                chkChasrate(Wa).Caption = CargoTypes(Wa)
                chkChasrate(Wa).Value = 0:  sliCrate(Wa).Value = 0: txtCam(Wa).Text = ""
            Else
                sliCrate(Wa).Visible = False
                txtCam(Wa).Visible = False
                chkChasrate(Wa).Visible = False
            End If
        Next Wa
        Frame2.Enabled = True
        cmbShowO.ListIndex = 0
    Else
        Frame2.Enabled = False
        For Wa = 0 To 11: chkChasrate(Wa).Visible = False: sliCrate(Wa).Visible = False: txtCam(Wa).Visible = False: Next Wa
    End If
    ShowList
End Sub
Public Sub PrepSave()
    If CurItmNo > -1 Then TTDXStationPut CurItm
End Sub
Private Sub ShowList()
    Dim Wa As Integer, wXa As TTDXStation
    
    lstStation.Clear
    If CurFile > " " Then
        For Wa = 0 To 249
            wXa = TTDXStationInfo(Wa)
            If wXa.BaseX > 0 Then
                If (cmbShowO.ItemData(cmbShowO.ListIndex) = wXa.Owner) Or (cmbShowO.ListIndex = 0) Then
                    lstStation.AddItem wXa.Name
                    lstStation.ItemData(lstStation.NewIndex) = Wa
                End If
            End If
        Next Wa
    End If
    For Wa = 0 To frmDta.UBound: frmDta(Wa).Enabled = False: Next Wa
    CurItmNo = lstStation.ListIndex
End Sub

Private Sub chkChasrate_Click(Index As Integer)
    If chkChasrate(Index).Value = 1 Then
        sliCrate(Index).Enabled = True: sliCrate(Index).Value = CurItm.CRate(Index)
        txtCam(Index).Enabled = True: txtCam(Index).Text = Format(CurItm.Cargo(Index))
        CurItm.CEnrout(Index) = CurItm.CEnroutOrg(Index)
    Else
        sliCrate(Index).Enabled = False: sliCrate(Index).Value = 0
        txtCam(Index).Enabled = False: txtCam(Index).Text = ""
        CurItm.CEnrout(Index) = 255
    End If
    
    MarkGame 8
End Sub

Private Sub cmbShowO_Click()
    ShowList
End Sub

Private Sub cmdAddHouse_Click()
    Dim Wx As Integer, Wy As Integer, Wva As TTDXlandscape, Wc As Long
    If CurItmNo > -1 Then
        With CurItm
            Wc = 0
            If .RailDir Then
                Wc = Wc + DoAddBuild(.RailX, .RailY, .RailTrackLen, .RailTracks)
            Else
                Wc = Wc + DoAddBuild(.RailX, .RailY, .RailTracks, .RailTrackLen)
            End If
            Select Case .AirportType
                Case 0: Wc = Wc + DoAddBuild(.AirX, .AirY, 4, 3)
                Case 1: Wc = Wc + DoAddBuild(.AirX, .AirY, 6, 6)
                Case 2: Wc = Wc + DoAddBuild(.AirX, .AirY, 1, 1)
            End Select
            Wc = Wc + DoAddBuild(.BusX, .BusY, 1, 1)
            Wc = Wc + DoAddBuild(.TruckX, .TruckY, 1, 1)
            Wc = Wc + DoAddBuild(.DockX, .DockY, 2, 2)
        End With
        Wc = MsgBox(Format(Wc) + " buildings added.")
        
        MarkGame 8
    End If
End Sub
Private Function DoAddBuild(vX As Byte, vY As Byte, vW As Byte, vH As Byte) As Long
    Dim Wx As Integer, Wy As Integer, Wc As Long, Wva As TTDXlandscape
    DoAddBuild = 0
    If vX < 1 Then Exit Function
    
    For Wx = vX - 1 To vX + vW
        For Wy = vY - 1 To vY + vH
            Wva = TTDXgetLandscape(Wx, Wy)
            If Wva.Object = 0 Then
                Wva.Object = 3: Wva.Owner = 0
                Select Case CInt(Rnd * 6)
                    Case 0: Wva.L2 = &HA
                    Case 1: Wva.L2 = &H4
                    Case 2: Wva.L2 = &H11
                    Case 3: Wva.L2 = &H1E
                    Case 4: Wva.L2 = &H24
                    Case Else: Wva.L2 = &H1C
                End Select
                Wva.L5 = 1: Wva.L3 = &H40
                Wc = Wc + 1
                TTDXputLandscape Wva
            End If
        Next Wy
    Next Wx
    For Wx = vX - 2 To vX + vW + 1
        For Wy = vY - 2 To vY + vH + 1
            Wva = TTDXgetLandscape(Wx, Wy)
            If Wva.Object = 0 Then
                Wva.Object = 3: Wva.Owner = 0
                Select Case CInt(Rnd * 5)
                    Case 0: Wva.L2 = &H9
                    Case 1: Wva.L2 = &H2
                    Case 2: Wva.L2 = &H3
                    Case 3: Wva.L2 = &H12
                    Case Else: Wva.L2 = &H1B
                End Select
                Wva.L5 = 1: Wva.L3 = &H40
                Wc = Wc + 1
                TTDXputLandscape Wva
            End If
        Next Wy
    Next Wx
    For Wx = vX - 4 To vX + vW + 3
        For Wy = vY - 4 To vY + vH + 3
            Wva = TTDXgetLandscape(Wx, Wy)
            If Wva.Object = 0 Then
                Wva.Object = 3: Wva.Owner = 0
                Select Case CInt(Rnd * 5)
                    Case 0: Wva.L2 = &H2
                    Case 1: Wva.L2 = &H1A
                    Case Else: Wva.L2 = &H6
                End Select
                Wva.L3 = &H40
                Wva.L5 = 1
                Wc = Wc + 1
                TTDXputLandscape Wva
            End If
        Next Wy
    Next Wx
    DoAddBuild = Wc
End Function

Private Sub cmdRemStop_Click()
    Dim Wx As Integer, Wy As Integer, Wva As TTDXlandscape, Wc As Long, Wfa As Boolean
    If CurItmNo > -1 Then
        With CurItm
            If .BusX > 0 And .TruckX > 0 Then
                Wc = MsgBox("Bus and truck stop already present.")
            Else
                Wfa = False
                If .RailDir Then
                    Wfa = AddRem(.RailX, .RailY, .RailTrackLen, .RailTracks)
                Else
                    Wfa = AddRem(.RailX, .RailY, .RailTracks, .RailTrackLen)
                End If
                Select Case .AirportType
                    Case 0: Wfa = Wfa Or AddRem(.AirX, .AirY, 4, 3)
                    Case 1: Wfa = Wfa Or AddRem(.AirX, .AirY, 6, 6)
                    Case 2: Wfa = Wfa Or AddRem(.AirX, .AirY, 1, 1)
                End Select
                Wfa = Wfa Or AddRem(.BusX, .BusY, 1, 1)
                Wfa = Wfa Or AddRem(.TruckX, .TruckY, 1, 1)
                Wfa = Wfa Or AddRem(.DockX, .DockY, 2, 2)
                If Wfa Then Wc = MsgBox("Station part(s) added.") Else: Wc = MsgBox("No new locations found.")
                
                MarkGame 8
            End If
        End With
    End If
End Sub
Private Function AddRem(vX As Byte, vY As Byte, vW As Byte, vH As Byte) As Boolean
    Dim Wx As Integer, Wy As Integer, Wc As Long, Wva As TTDXlandscape
    AddRem = False
    If vX < 1 Then Exit Function
    
    For Wx = vX - 4 To vX + vW + 3
        For Wy = vY - 4 To vY + vH + 3
            Wva = TTDXgetLandscape(Wx, Wy)
            If Wva.Object = 1 And Wva.Owner = CurItm.Owner Then
                If Wva.L5 = &H3F Then
                    If CurItm.BusX = 0 Then
                        Wva.Object = 5: Wva.L2 = CurItm.Number: Wva.L5 = &H47: Wva.L3 = 0
                        CurItm.BusX = Wva.X: CurItm.BusY = Wva.Y
                        CurItm.BusStatus = 3
                        CurItm.Parts = CurItm.Parts Or 4
                    ElseIf CurItm.TruckX = 0 Then
                        Wva.Object = 5: Wva.L2 = CurItm.Number: Wva.L5 = &H43: Wva.L3 = 0
                        CurItm.Parts = CurItm.Parts Or 2
                        CurItm.TruckX = Wva.X: CurItm.TruckY = Wva.Y
                    End If
                    AddRem = True
                    TTDXputLandscape Wva
                End If
            End If
        Next Wy
    Next Wx
End Function

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'If Shift = 0 Then
    '    If KeyCode = vbKeyF1 Then ViewFile Me, App.Path + "/docs/3e.html"
    'End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PrepSave
End Sub

Private Sub lstStation_Click()
    Dim Wa As Integer, Wb As Integer
    If CurItmNo > -1 Then TTDXStationPut CurItm
    Wb = lstStation.ItemData(lstStation.ListIndex)
    CurItm = TTDXStationInfo(Wb)
    CurItmNo = Wb
    UpdateFields
    For Wa = 0 To frmDta.UBound: frmDta(Wa).Enabled = True: Next Wa
End Sub
Private Sub UpdateFields()
    Dim Wa As Integer, Wb As Integer
    
    txtX.Text = CurItm.BaseX
    txtY.Text = CurItm.BaseY
    For Wa = 0 To 11
        If CurItm.CEnrout(Wa) = 255 Then chkChasrate(Wa).Value = 0 Else: chkChasrate(Wa).Value = 1
        chkChasrate_Click Wa
    Next Wa
    If CurItm.Owner = &H10 Then
        txtOwner.Text = "<No One>"
    Else
        txtOwner.Text = "Player " + Format(CurItm.Owner + 1)
    End If
    
    frmTechInfo.ShowInfo 4, CurItm.Name, CurItm.Offset
    frmMap.SetHighlight CInt(CurItm.BaseX), CInt(CurItm.BaseY)
End Sub

Private Sub Form_Load()
    Dim Wa As Integer, Wb As Integer
    For Wa = 1 To 11
        Load chkChasrate(Wa)
        Load sliCrate(Wa)
        Load txtCam(Wa)
    Next Wa
    UpdateInfo
End Sub


Private Sub sliCrate_Change(Index As Integer)
    CurItm.CRate(Index) = sliCrate(Index).Value
    MarkGame 8
End Sub

Private Sub txtCam_Change(Index As Integer)
    If txtCam(Index).Enabled Then
        If jBetween(-1, Val(txtCam(Index).Text), 4096) Then
            CurItm.Cargo(Index) = Val(txtCam(Index))
        Else
            txtCam(Index) = CStr(CurItm.Cargo(Index))
        End If
    End If

    MarkGame 8
End Sub
Private Sub txtCam_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = CheckNumInput("", KeyAscii)
End Sub


