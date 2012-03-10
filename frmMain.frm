VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   9390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12855
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   12855
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstIndu 
      Height          =   1425
      Left            =   5160
      TabIndex        =   1
      Top             =   240
      Width           =   4455
   End
   Begin VB.ListBox lstCities 
      Height          =   5325
      Left            =   8280
      TabIndex        =   0
      Top             =   3960
      Width           =   4575
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   2535
      Left            =   5040
      TabIndex        =   3
      Top             =   3000
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   855
      Left            =   5160
      TabIndex        =   2
      Top             =   1920
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





Private Sub Form_Load()
    Dim Wa As Integer
    Debug.Print TTDXLoadFile("C:\Program Files\Microsoft Visual Studio\VB98\0J-TTDXEdit\Sv1tool\trp03.sv1")
    'Debug.Print TTDXdata.SaveFile("C:\Program Files\Microsoft Visual Studio\VB98\0J-TTDXEdit\Sv1tool\trp04.sv1")
    
    For Wa = 0 To 69
        lstCities.AddItem Format(Wa, "00") + " " + CityName(Wa)
        lstCities.ItemData(lstCities.NewIndex) = Wa
       'Debug.Print
    Next Wa
End Sub

Private Sub lstCities_Click()
    Dim Wa As Integer, Wb As Integer
    Label2.Caption = ""
    Wb = lstCities.ItemData(lstCities.ListIndex)
    For Wa = 0 To &H5E
        Label2.Caption = Label2.Caption + " " + Hex(CityData(Wb, Wa))
    Next Wa
End Sub

Private Sub lstIndu_Click()
    Dim Wa As Integer, Wb As Integer
    Label1.Caption = ""
    Wb = lstIndu.ItemData(lstIndu.ListIndex)
    For Wa = 0 To &H35
        Label1.Caption = Label1.Caption + " " + Hex(TTDXIndustryData(Wb, Wa))
    Next Wa
    Label1.Caption = Label1.Caption + Chr(10) + CityName((TTDXIndustryInfo(Wb).HomeTown))
End Sub
