VERSION 5.00
Begin VB.Form frmMap 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Map"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7845
   FillColor       =   &H0000FF00&
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000C000&
   Icon            =   "frmMap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   7845
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tim1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   4200
      Top             =   7800
   End
   Begin VB.PictureBox picMap 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   0
      MousePointer    =   2  'Cross
      ScaleHeight     =   413
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   413
      TabIndex        =   0
      Top             =   0
      Width           =   6255
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private hX1 As Integer, hX2 As Integer, hY1 As Integer, hY2 As Integer, hFl As Boolean, hT As Integer
Private ColSet(15) As Long

'
' SetPixel is faster than pset
'
Private Declare Function SetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Public MapSize As Integer

Public Sub UpdateInfo()
    Dim Wx As Integer, Wy As Integer, Wva As TTDXlandscape
    Dim Wax As Integer, Way As Integer, Wd As Long, We As Long
    Dim Wbx As Integer, Wby As Integer
    '
    tim1.Enabled = False
    frmMDI.mnVMlarge.Checked = False: frmMDI.mnVMnone.Checked = False
    frmMDI.mnVMsmall.Checked = False: frmMDI.mnVMextr.Checked = False
    If MapSize = 0 Then
        frmMDI.mnVMnone.Checked = True
        Me.Hide
    ElseIf CurFile > " " Then
        picMap.Width = 60 + 15 * 255 * MapSize
        picMap.Height = picMap.Width
        Me.Show: Me.Refresh: DoEvents
        If MapSize = 1 Then frmMDI.mnVMsmall.Checked = True
        If MapSize = 2 Then frmMDI.mnVMlarge.Checked = True
        If MapSize = 4 Then frmMDI.mnVMextr.Checked = True
        picMap.BackColor = &H0
        picMap.Cls: picMap.Refresh: DoEvents
        Me.Refresh
        If CurFile > "" Then
            For Wx = 0 To 254
                For Wy = 0 To 254
                    We = ColSet(TTDXGetLandMap(Wx, Wy))
                    Wbx = (254 - Wx) * MapSize: Wby = Wy * MapSize
                    For Wax = 0 To (MapSize - 1)
                        For Way = 0 To (MapSize - 1)
                            Wd = SetPixel(picMap.hDC, Wbx + Wax, Wby + Way, We)
                        Next Way
                    Next Wax
                Next Wy
                If Not fFastMode Then: DoEvents: picMap.Refresh
            Next Wx
            picMap.Refresh
        End If
    Else
        Me.Hide
    End If
End Sub
Private Sub SetPoint(wXa As Integer, wYa As Integer, wT As Byte)
    Dim Wax As Integer, Way As Integer, Wx As Integer, Wy As Integer, Wc As Long, Wd As Long
    Select Case wT
        Case 0: Wc = &H7000&    ' Nothing
        Case 1: Wc = &H0&       ' Rails
        Case 2: Wc = &H80&      ' Road (crossing)
        Case 3: Wc = &H606060   ' House
        Case 4: Wc = &HA000&    ' Tree(s)
        Case 5: Wc = &HAAAAAA   ' Station/Depot
        Case 6: Wc = &H900000   ' Water
        Case 8: Wc = &HAAAA&    ' Industry
        Case 9: Wc = &H0&       ' Tunnel/Bridge
        Case 10: Wc = &H2000&   ' Solid objects (Antenna, Lighttower, HQ)
        
        Case Else: Wc = &HFFFFFF
    End Select
    
    For Wax = 0 To (MapSize - 1)
        For Way = 0 To (MapSize - 1)
            Wx = (254 - wXa) * MapSize + Wax: Wy = wYa * MapSize + Way
            Wd = SetPixel(picMap.hDC, Wx, Wy, Wc)
        Next Way
    Next Wax
End Sub

Private Sub SetPointOLD(wXa As Integer, wYa As Integer, wT As Byte)
    Dim Wax As Integer, Way As Integer, Wx As Integer, Wy As Integer
    For Wax = 0 To (MapSize - 1)
        For Way = 0 To (MapSize - 1)
            Wx = (254 - wXa) * MapSize + Wax: Wy = wYa * MapSize + Way
            Select Case wT
                Case 0: picMap.PSet (Wx, Wy), &H7000&
                Case 1: picMap.PSet (Wx, Wy), &H0&
                Case 2: picMap.PSet (Wx, Wy), &H80&
                Case 3: picMap.PSet (Wx, Wy), &H606060
                Case 4: picMap.PSet (Wx, Wy), &HA000&
                Case 5: picMap.PSet (Wx, Wy), &HAAAAAA
                Case 6: picMap.PSet (Wx, Wy), &H900000
                Case 8: picMap.PSet (Wx, Wy), &HAAAA&
                Case 9: picMap.PSet (Wx, Wy), &H0&
                Case 10: picMap.PSet (Wx, Wy), &H2000&
                
                Case Else: picMap.PSet (Wx, Wy), &HFFFFFF
            End Select
        Next Way
    Next Wax
End Sub

Public Sub SetHighlight(wx1 As Integer, wy1 As Integer)
    StopHighlight
    If Me.Visible Then
        hX1 = wx1
        hY1 = wy1
        tim1.Enabled = True
    End If
End Sub
Public Sub StopHighlight()
    If tim1.Enabled Then
        tim1.Enabled = False: hFl = False: tim1_Timer
    End If
End Sub
Private Sub tim1_Timer()
    Dim Wa As Integer, Wb As Integer, Wva As TTDXlandscape
    For Wa = hX1 - 4 To hX1 + 4
        If hFl Then SetPointOLD Wa, hY1, 20 Else Wva = TTDXgetLandscape(Wa, hY1): SetPoint Wa, hY1, Wva.Object
    Next Wa
    For Wa = hY1 - 4 To hY1 + 4
        If hFl Then SetPointOLD hX1, Wa, 20 Else Wva = TTDXgetLandscape(hX1, Wa): SetPoint hX1, Wa, Wva.Object
    Next Wa
    hFl = Not hFl
End Sub


Private Sub Form_Load()
    ColSet(0) = &H7000&
    ColSet(1) = &H0&
    ColSet(2) = &H80&
    ColSet(3) = &H606060
    ColSet(4) = &HA000&
    ColSet(5) = &HAAAAAA
    ColSet(6) = &H900000
    ColSet(7) = &HFFFFFF
    ColSet(8) = &HAAAA&
    ColSet(9) = &H0&
    ColSet(10) = &H2000&
    ColSet(11) = &HFFFFFF
    ColSet(12) = &HFFFFFF
    ColSet(13) = &HFFFFFF
    ColSet(14) = &HFFFFFF
    ColSet(15) = &HFFFFFF
    UpdateInfo
End Sub

Private Sub picMap_MouseMove(Button As Integer, shift As Integer, X As Single, Y As Single)
    Dim Wva As TTDXlandscape
    
    Wva = TTDXgetLandscape(254 - Fix(X / MapSize), Fix(Y / MapSize))
    picMap.ToolTipText = "X:" + Format(Wva.X) + " Y:" + Format(Wva.Y) + "  L1:" + Hex(Wva.L1) + " L2:" + Hex(Wva.L2) + " L3:" + Hex(Wva.L3) + " L4:" + Hex(Wva.L4) + " L5:" + Hex(Wva.L5)
End Sub

Private Sub picMap_Resize()
    Me.ScaleMode = 1
    Me.Width = picMap.Width + (Me.Width - Me.ScaleWidth)
    Me.Height = picMap.Height + (Me.Height - Me.ScaleHeight)
End Sub

