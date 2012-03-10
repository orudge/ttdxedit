VERSION 5.00
Begin VB.Form frmTechInfo 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "TechInfo"
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   3165
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   12  'No Drop
   ScaleHeight     =   7005
   ScaleWidth      =   3165
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picTechCon 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   6495
      Left            =   -15
      ScaleHeight     =   6465
      ScaleWidth      =   3420
      TabIndex        =   0
      Top             =   480
      Width           =   3450
      Begin VB.VScrollBar vsDebug 
         Height          =   6855
         Left            =   2910
         TabIndex        =   2
         Top             =   0
         Width           =   255
      End
      Begin VB.PictureBox picTech 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6735
         Left            =   0
         ScaleHeight     =   6735
         ScaleWidth      =   2910
         TabIndex        =   1
         Top             =   0
         Width           =   2910
      End
   End
   Begin VB.Label labName 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   3135
   End
End
Attribute VB_Name = "frmTechInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    SetMe
End Sub

Public Sub SetMe()
    If Me.Visible Then
        Me.Refresh
    End If
End Sub

Private Sub Form_Load()
    Me.Left = (frmMDI.Left + frmMDI.Width) - Me.Width
    Me.Top = frmMDI.Top
End Sub

Private Sub Form_Resize()
    If Me.ScaleHeight - picTechCon.Top > 32 Then
        If Me.Width <> 3255 Then Me.Width = 3255: Exit Sub
        picTechCon.Height = Me.ScaleHeight - picTechCon.Top + 15
        vsDebug.Height = (picTechCon.ScaleHeight)
        vsDebug.Max = (picTech.Height - vsDebug.Height) / 15
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMDI.mnVTech.Checked = False
End Sub

Private Sub vsDebug_Change()
    picTech.Top = -vsDebug.Value * 15
End Sub
Public Sub ShowInfo(wType As Integer, wName As String, wOffset As Long)
    Dim Wa As Long, Wb As Long, Wsa As String, Wsb As String, Wsc As String
    picTech.Cls
    If Me.Visible Then
        labName.Caption = wName + Chr(10) + "Offset: " + Hex(wOffset)
        Select Case wType
            Case 1: Me.Caption = "TechInfo (Player)": Wb = 945
            Case 2: Me.Caption = "TechInfo (City)": Wb = &H5D
            Case 3: Me.Caption = "TechInfo (Industry)": Wb = &H35
            Case 4: Me.Caption = "TechInfo (Station)": Wb = &H8D
            Case 5: Me.Caption = "TechInfo (Vehicle)": Wb = &H80
        End Select
        '
        ' This sets the bytes that are known and should be shown in different color
        '
        Select Case wType
            '              0                               0                               0                               0                               0                               0                               0                               0                               1                               1                               1                               1                               1                               1                               1                               1                               2                               2                               2                               2                               2                               2                               2                               2                               3                               3                               3                               3                               3                               3                               3
            '              0               1               2               3               4               5               6               7               8               9               A               B               C               D               E               F               0               1               2               3               4               5               6               7               8               9               A               B               C               D               E               F               0               1               2               3               4               5               6               7               8               9               A               B               C               D               E               F               0               1               2               3               4               5               6               7               8               9               A               B               C
            '              0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0       8       0
            Case 1: Wsa = "                xxxxxxxx                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            xx"
            Case 2: Wsa = "xxxxxxxxxx                    xxxxxxxxxxxxxxxxx                                                "
            Case 3: Wsa = "xxxx  xxxx    xxxxx xxxxxxxx          x              "
            Case 4: Wsa = "xxxx  xxxxxxxxxxx   xx      xx xx   xx xx   xx xx   xx xx   xx xx   xx xx   xx xx   xx xx   xx xx   xx xx   xx xx   xx xx      xxxxx   xx"
            Case 5: Wsa = "xxxxxx                               x                   xxxxx  xxxxxx        xxxx        xxxxxxxx"
        End Select
        Wsb = "000: "
        Wsc = "     "

        For Wa = 0 To Wb
            If Wa > 0 And Wa Mod 8 = 0 Then
                Wsb = Wsb + Chr(10) + Right("0000" + Hex(Wa), 3) + ": "
                Wsc = Wsc + Chr(10) + "     "
            ElseIf Wa > 0 And Wa Mod 2 = 0 Then
                Wsb = Wsb + " ": Wsc = Wsc + " "
            End If
            If Mid(Wsa, Wa + 1, 1) = "x" Then
                Wsb = Wsb + "  "
                Wsc = Wsc + Right("00" + Hex(TTDXGetByte(wOffset + Wa)), 2)
            Else
                Wsb = Wsb + Right("00" + Hex(TTDXGetByte(wOffset + Wa)), 2)
                Wsc = Wsc + "  "
            End If
        Next Wa
        '
        ' Adapt the required size and put the text on
        '
        picTech.ForeColor = vbButtonText
        picTech.Height = picTech.TextHeight(Wsb)
        vsDebug.Max = (picTech.Height - vsDebug.Height) / 15
        picTech.Print Wsb
        picTech.CurrentX = 0: picTech.CurrentY = 0: picTech.ForeColor = vbInactiveCaptionText
        picTech.Print Wsc
    End If
End Sub
