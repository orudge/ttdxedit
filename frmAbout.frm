VERSION 5.00
Begin VB.Form frmAboutOld 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer timerTW 
      Enabled         =   0   'False
      Interval        =   80
      Left            =   7080
      Top             =   2160
   End
   Begin VB.Timer tim2 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   7080
      Top             =   1800
   End
   Begin VB.Timer timerMain 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   7080
      Top             =   1440
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1305
      Left            =   0
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   895.475
      ScaleMode       =   0  'User
      ScaleWidth      =   1369.55
      TabIndex        =   1
      Top             =   0
      Width           =   1980
   End
   Begin VB.PictureBox PicWork 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1695
      Left            =   0
      ScaleHeight     =   113
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   505
      TabIndex        =   0
      Top             =   1320
      Width           =   7575
   End
   Begin VB.Label labdummy 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   4080
      TabIndex        =   8
      Top             =   3030
      Width           =   3495
   End
   Begin VB.Label labMes 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Show Messages"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   225
      Left            =   2040
      TabIndex        =   7
      Top             =   3030
      Width           =   2025
   End
   Begin VB.Label labLegal 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Show Legal stuff"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C0C0&
      Height          =   225
      Left            =   0
      TabIndex        =   6
      Top             =   3030
      Width           =   2025
   End
   Begin VB.Label labMail 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "E-Mail Author"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   4800
      TabIndex        =   5
      Top             =   1080
      Width           =   2790
   End
   Begin VB.Label labWeb 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Go To Webpage"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   1995
      TabIndex        =   4
      Top             =   1080
      Width           =   2790
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   360
      Left            =   1995
      TabIndex        =   3
      Top             =   705
      Width           =   5595
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "TTDX Editor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   1995
      TabIndex        =   2
      Top             =   0
      Width           =   5580
   End
End
Attribute VB_Name = "frmAboutOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const Jh As Integer = 80
Private Char(48) As String, CharPosX(48) As Integer, CharPosY(48) As Single, CharJump(48) As Integer
Private Bfl(48) As Boolean, Chardel(48) As Integer
Private Clen As Integer, Cst As Integer, Wtxt As String, TxtPos As Integer

Private Wa As Integer
Private Sub SetString()
    Dim Wsa As String, Wx As Integer, wFade As Integer, Wmod As Integer
    
    timerMain.Enabled = False: tim2.Enabled = False: timerTW.Enabled = False
    PicWork.Cls
    Wsa = "This Is A T-T-T-T-T-T-T-T-T-T-T-T-T-Test": wFade = 0: Wmod = 0
    Select Case Cst
        Case 1: Wsa = "Written by: Jens Vang Petersen": wFade = 3
        Case 2: Wsa = "Maintained by: Owen Rudge": wFade = 3
        Case 3: Wsa = "I would like to thank the following people:"
        Case 4: Wsa = "Josef (TTDPatch) Drexler.": wFade = 1
        Case 5: Wsa = "Marcin Grzegorczyk.": wFade = 1
        Case 6: Wsa = "Owen Rudge & Nick Hundley.": wFade = 1
        Case 7: Wsa = "Ehh, you're not still reading this, are you ??": wFade = 2
        Case 8: Wsa = "This text inspired by Turbo Outrun - C64.": wFade = 2
        Case 9: Wsa = "The Legal Blah-Blah-Blah coming up:"
        Case 10: Wsa = "By using this program the user accepts full£liability for any damage that could possibly£be caused by its use or misuse.¤¤£The author cannot be held responsible in any way£for damage to either hardware or software.¤¤¤¤": Wmod = 1
        Case 11: Wsa = "[In plain words:]£If you break anything or anybody by using this£editor, it's your own fault..¤¤¤¤": Wmod = 1
        Case 12: Wsa = "This program is released as Mailware.¤£If you find it useful please send me an E-mail£with your opinions and new ideas.£The only way to improve programs is getting£input from people using them.¤¤¤¤": Wmod = 1
        Case 13: Wsa = "This software may never be sold for profit.£It may however be included in PD/Shareware£collections on cd's or websites, this also£includes coverdiscs..¤¤¤¤": Wmod = 1
        Case 14: Wsa = "Transport Tycoon Deluxe is Copyrighted by£Microprose, FishUK and/or Chris Sawyer.£This Editor IS NOT a crack, it won't help you to run£an illegal copy of the game, but it won't prevent/lyou from doing so either.¤¤¤¤": Wmod = 1
        Case 15: Wsa = "I can't belive you're still reading.": wFade = 2
        'Case 15: Wsa = "[Special Message To Ben:]¤¤£I'm to sexy for your proxy,¤£to sexy for your proxy.¤£It's going to kill me.¤.¤.¤.¤.¤¤": Wmod = 1
        Case 16: Wsa = "One time in Denver, I ate at this resturant. They£had a big sign on the wall: Watch your hat and£coat. While I was watching, someone stole my£steak.¤¤£ - Boone, Where The Hell's That Gold¤¤¤¤": Wmod = 1: Cst = 0
    End Select
    Cst = Cst + 1
    If Wmod = 0 Then
        PicWork.FontSize = 18
        Erase Char
        Wx = (PicWork.ScaleWidth - PicWork.TextWidth(Wsa)) / 2
        Clen = Len(Wsa) - 1
        For Wa = 0 To Clen
            Char(Wa) = Mid(Wsa, Wa + 1, 1)
            CharPosX(Wa) = Wx: Wx = Wx + PicWork.TextWidth(Char(Wa))
            CharPosY(Wa) = 90
            Select Case wFade
                Case 0: Chardel(Wa) = Wa
                Case 1: Chardel(Wa) = Rnd * 60
                Case 2: Chardel(Wa) = (Clen - Wa)
                Case 3: Chardel(Wa) = (Wa Mod 2) * 24
            End Select
            CharJump(Wa) = Jh + 30
        Next Wa
        timerMain.Enabled = True
    ElseIf Wmod = 1 Then
        PicWork.FontSize = 16
        PicWork.Cls
        Wtxt = Wsa: TxtPos = 0
        timerTW.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    Label2.Caption = "Version " + Format(App.Major, "0") + "." + Format(App.Minor, "00") + "." + Format(App.Revision, "000")
    Cst = 1
    SetString
End Sub

Private Sub labLegal_Click()
    Cst = 8: SetString
End Sub
Private Sub labMes_Click()
    Cst = 15: SetString
End Sub

Private Sub labMail_Click()
    Dim rc As Long
    rc = ShellExecute(Screen.ActiveForm.hwnd, "Open", "mailto:owen@owenrudge.net", vbNullString, App.Path, vbNormalFocus)
End Sub


Private Sub labWeb_Click()
    ViewFile Me, "http://www.owenrudge.net/TT/"
End Sub

Private Sub tim2_Timer()
    PicWork.Cls
    For Wa = 0 To Clen
        If Chardel(Wa) = 0 Then
            CharJump(Wa) = CharJump(Wa) + 3
        Else
            Chardel(Wa) = Chardel(Wa) - 1
        End If
        PicWork.CurrentX = CharPosX(Wa)
        PicWork.CurrentY = CharJump(Wa) + Jh
        PicWork.Print Char(Wa)
    Next Wa
    If CharJump(Clen) > 20 Then SetString
End Sub

Private Sub timerMain_Timer()
    Dim Wb As Integer, Sj As Boolean
    
    PicWork.Cls
    Sj = False
    For Wa = 0 To Clen
        If Chardel(Wa) = 0 Then
            PicWork.CurrentX = CharPosX(Wa)
            Wb = Abs(Sin(CharPosY(Wa) * 0.017453293)) * CharJump(Wa)
            PicWork.CurrentY = Jh - Wb
            If Wb < 1 And Bfl(Wa) Then
                If CharJump(Wa) > 4 Then CharJump(Wa) = CharJump(Wa) / 2.5 Else CharJump(Wa) = 0
                Bfl(Wa) = False
            Else
                Bfl(Wa) = True
            End If
            CharPosY(Wa) = CharPosY(Wa) + 7.5
            PicWork.Print Char(Wa)
            If CharJump(Wa) > 0 Then Sj = True
        Else
            Chardel(Wa) = Chardel(Wa) - 1: Sj = True
        End If
    Next Wa
    If Not Sj Then: SetOut
End Sub

Private Sub SetOut()
    timerMain.Enabled = False
    For Wa = 0 To Clen
        Chardel(Wa) = Wa
        CharJump(Wa) = 0
    Next Wa
    tim2.Enabled = True
End Sub

Private Sub timerTW_Timer()
    Dim Wsa As String, Wsb As String
    timerTW.Interval = 60
    TxtPos = TxtPos + 1
    If TxtPos > 1 Then PicWork.Line -Step(-12, -18), &H0&, BF: PicWork.CurrentY = PicWork.CurrentY - 4
    Wsa = Mid(Wtxt, TxtPos, 1): Wsb = Mid(Wtxt, TxtPos, 2)
    If Wsb = "/l" Then Wsa = Chr(13): TxtPos = TxtPos + 1: PicWork.CurrentY = PicWork.CurrentY - 3
    If Wsb = "/p" Then Wsa = "": TxtPos = TxtPos + 1
    If Wsa = "¤" Then timerTW.Interval = 750: Wsa = ""
    If Wsa = "£" Then Wsa = Chr(13):  PicWork.CurrentY = PicWork.CurrentY - 3
    If Wsa = "[" Then PicWork.ForeColor = &HAAFFAA: Wsa = ""
    If Wsa = "]" Then PicWork.ForeColor = &HFF00&: Wsa = ""
    PicWork.Print Wsa;
    PicWork.Line Step(0, 4)-Step(12, 18), , BF
    If TxtPos > Len(Wtxt) Then SetString
End Sub
