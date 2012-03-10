Attribute VB_Name = "JVPToolProcs"
'-----------------------------------------------------------------------------------
' JVP ToolProcedures V2.00 by Jens Vang Petersen (VB6 Edition)
'-----------------------------------------------------------------------------------
' Discs and files:
'  Boolan = jExistFile(Filename)       Check if a file exist
'  String = BuildFileName(Path,File)   Build a full filename (adds \ to path if needed)
'  String = GetDiscVolume(RootPath)    Get the Volumename of a disc (HD, FLOPPY, CD, etc.)
'  Long   = GetDiscSerial(RootPath)    Get the Serialnumber of a disc
'  ExecuteFile(File)                   Execute another program.
'  ViewFile(CallingForm,File)          Open a file in it's default program.
'                                      (Normaly you'll use "Me" for CallingForm)
'
' Sound (Remember to do a WAVStop before you leave the program):
'  WAVPlay(File)                       Play a Wav sound
'  WAVLoop(File)                       Play a Wav sound and repeat it forever
'  WAVStop(File)                       Stop the sound playing (if any)
'
' Strings
'  String = Removechar(Source,Remove)  Removes part of a string like a$=a$-"STR" in Amos
'  String = Replacechar(Src,Rem,New)   Replaces 'Rem' with 'new' in Src string.
'
'-----------------------------------------------------------------------------------
' IMPORTAINT: Some of these functions and procedures make calls to other in the
' module. Especialy the string functions are called often by other functions.
'-----------------------------------------------------------------------------------
' The JVP ToolProcedures is something I began making in my old Amiga-Amos days, And
' I've now moved them onto my new world in VB5.0. The Idea is basicaly to have a
' shared module of procedures (I usualy include the SAME module in all my projects,
' so if I find a bug it's automaticaly adjusted in all other projects.) doing stuff
' I often need to do, and so I don't have to invent the same twice, or go looking
' through older projects to try and find the needed piece of code. Especialy when
' one's used to working with AMOS and the many extensions from there VB seems a bit
' limited in functions, so many of these 'tools' is in fact a remake of some old
' but usefull functions from other coding systems.
'-----------------------------------------------------------------------------------
' Please note that some of these functions are based on methods I've found on the
' web, if I've been able to find the original author he'll be give credit in the
' proper procedure..
'-----------------------------------------------------------------------------------




Public Declare Function ShellExecute Lib "Shell32.dll" Alias "ShellExecuteA" _
(ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long

Private Declare Function FindExecutableA Lib "Shell32.dll" _
(ByVal lpFile As String, ByVal lpDirectory As String, ByVal lpResult As String) As Long

Private Declare Function SetVolumeLabelA Lib "kernel32" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long


'Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As LARGE_INTEGER, lpTotalNumberOfBytes As LARGE_INTEGER, lpTotalNumberOfFreeBytes As LARGE_INTEGER) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
    Public Const SND_SYNC = &H0
    Public Const SND_ASYNC = &H1
    Public Const SND_NODEFAULT = &H2
    Public Const SND_MEMORY = &H4
    Public Const SND_LOOP = &H8
    Public Const SND_NOSTOP = &H10
                            
Public Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long

Public Const GWL_WNDPROC = (-4)


Private MsgID_QueryCancelAutoPlay As Long, AROldProc As Long

Private Type SHITEMID
    cb   As Long
    abID As Byte
End Type

Private Type ITEMIDLIST
    mkid As SHITEMID
End Type

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Retrieves the ID of a special folder.
Private Declare Function SHGetSpecialFolderLocation Lib "Shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
' Converts an item identifier list to a file system path.
Private Declare Function SHGetPathFromIDList Lib "Shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long


Public Function fGetSpecialFolder(vFormHwnd As Long, CSIDL As Long) As String
    Dim sPath As String
    Dim IDL   As ITEMIDLIST
    '
    ' Retrieve info about system folders such as the
    ' "Recent Documents" folder.  Info is stored in
    ' the IDL structure.
    '
    fGetSpecialFolder = ""
    If SHGetSpecialFolderLocation(vFormHwnd, CSIDL, IDL) = 0 Then
        '
        ' Get the path from the ID list, and return the folder.
        '
        sPath = Space$(260)
        If SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath) Then
            fGetSpecialFolder = Left$(sPath, InStr(sPath, vbNullChar) - 1) & "\"
        End If
    End If
End Function

'
' Functions regarding numbers
'
Public Function RandomNumber(wFirst, wLast) As Long
    RandomNumber = Int((wLast - wFirst + 1) * Rnd + wFirst)
End Function

Public Function jRndPick(ParamArray wValues() As Variant) As Variant
    Debug.Print UBound(wValues)
    jRndPick = wValues(Int((UBound(wValues) + 1) * Rnd))
End Function


Public Function SecondsToTime(wSec As Long) As String
    Wa = Int(wSec / 3600)
    Wc = wSec Mod 3600
    Wb = Int(Wc / 60)
    Wc = Wc Mod 60
    If Wa > 0 Then
        SecondsToTime = Format(Wa, "00") + ":" + Format(Wb, "00") + ":" + Format(Wc, "00")
    Else
        SecondsToTime = Format(Wb, "00") + ":" + Format(Wc, "00")
    End If
End Function

Public Sub ExecuteFile(ByVal FilePath As String)
    '
    ' Execute a file
    ' Based on code from planet-source-code.com
    '
    On Error GoTo 0
    ret = Shell("rundll32.exe url.dll,FileProtocolHandler " & (FilePath), vbNormalFocus)
End Sub

Public Sub ViewFile(ByVal wCaller As Form, ByVal wFile As String)
    rc = ShellExecute(wCaller.hwnd, "Open", wFile, vbNullString, App.Path, vbNormalFocus)
End Sub

Public Function FindExecutable(ByVal s As String) As String
    Dim i As Integer, s2 As String

    s2 = String(1024, 31)
    i = FindExecutableA(s, vbNullString, s2)
    s2 = Trim(s2)
    If i > 32 Then
        'FindExecutable = Left$(s2, InStr(s2, Chr$(0)) - 1)
        FindExecutable = s2
    Else
        FindExecutable = ""
    End If
    Debug.Print s2
End Function



Public Sub InitStopAutoRun(ByVal wCaller As Form)
    '
    '   We ask for a message from the system just before autorun is called
    '   Then we store the normal eventhandler and setup our own..
    '
    MsgID_QueryCancelAutoPlay = RegisterWindowMessage(ByVal "QueryCancelAutoPlay")
    AROldProc = GetWindowLong(wCaller.hwnd, GWL_WNDPROC)
    SetWindowLong wCaller.hwnd, GWL_WNDPROC, AddressOf StopAutoRun
End Sub
Private Function StopAutoRun(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    '
    '   First we check to see if we've got the right event, if not we call the
    '   normal eventhandler..
    '
    If wMsg = MsgID_QueryCancelAutoPlay Then
        sText = Format(wMsg) + " - " + Format(wparm) + " - " + Format(lparm)
        Form1.Text1 = Form1.Text1 & sText & vbCrLf
        StopAutoRun = 1
    Else
        StopAutoRun = CallWindowProc(AROldProc, hwnd, wMsg, wParam, lParam)
    End If
End Function
Public Sub CloseStopAutoRun(ByVal wCaller As Form)
    SetWindowLong wCaller.hwnd, GWL_WNDPROC, AROldProc
End Sub



'-----------------------------------------------------------------------------------
'
' Stuff for strings
'
'-----------------------------------------------------------------------------------
Public Function jRandomName(ByVal minLen As Integer, ByVal maxLen As Integer) As String
    Dim Wa As Integer, Wb As Integer
    Randomize
    Wa = Int((maxLen - minLen + 1) * Rnd) + minLen:
    jRandomName = ""
    For Wb = 1 To Wa
        jRandomName = jRandomName + Chr(Int(Rnd * 26) + 97)
    Next Wb
End Function

Public Function jRandomString(ByVal minWords As Integer, ByVal maxWords As Integer) As String
    Dim Wa As Integer

    Wa = Int((maxWords - minWords + 1) * Rnd) + minWords:
    For Wb = 1 To Wa - 1
        jRandomString = jRandomString + jRandomName(1, 14) + " "
    Next Wb
    jRandomString = jRandomString + jRandomName(2, 18)
End Function

Public Function jAscUL(ByVal wStr As String) As Double
    wStr = wStr + ChrB(0) + ChrB(0) + ChrB(0) + ChrB(0)
    jAscUL = jAscUW(wStr) + jAscUW(Mid(wStr, 3)) * 65536
End Function

Public Function jAscbL(ByVal wStr As String) As Long
    Dim W1 As Long, W2 As Long, W3 As Long, W4 As Long
    '
    ' Take a 4 byte string into a Long number (Intel format with Lo-byte first)
    '
    wStr = wStr + ChrB(0) + ChrB(0) + ChrB(0) + ChrB(0)
        
    W1 = AscB(wStr): W2 = AscB(MidB(wStr, 2)): W3 = AscB(MidB(wStr, 3)): W4 = AscB(MidB(wStr, 4))
    
    'Debug.Print W1, W2, W3, W4
    If W4 > 127 Then
        jAscbL = W1 - 256 + (W2 - 255) * 256 + (W3 - 255) * 65536 + (W4 - 255) * 16777216
    Else
        jAscbL = W1 + W2 * 256 + W3 * 65536 + W4 * 16777216
    End If
End Function

Public Function jChrbL(ByVal wVal As Long) As String
    Dim W1 As Integer, W2 As Integer, W3 As Integer, W4 As Integer
    
    If wVal < 0 Then
        W4 = Int(wVal / 16777216) + 255: wVal = wVal Mod 16777216
        W3 = Int(wVal / 65536) + 255: wVal = wVal Mod 65536
        W2 = Int(wVal / 256) + 255: wVal = wVal Mod 256
        W1 = wVal + 256
    Else
        W4 = Int(wVal / 16777216): wVal = wVal Mod 16777216
        W3 = Int(wVal / 65536): wVal = wVal Mod 65536
        W2 = Int(wVal / 256): wVal = wVal Mod 256
        W1 = wVal
    End If
    'Debug.Print W1, W2, W3, W4
    jChrbL = ChrB(W1) + ChrB(W2) + ChrB(W3) + ChrB(W4)
End Function

Public Function jAscUW(ByVal wStr As String) As Double
    wStr = wStr + ChrB(0) + ChrB(0)
    jAscUW = AscB(wStr) + AscB(MidB(wStr, 2)) * 256
End Function

Public Function jAscW(ByVal wStr As String) As Double
    Dim W1 As Long, W2 As Long
    '
    ' Take a 4 byte string into a Long number (Intel format with Lo-byte first)
    '
    wStr = wStr + ChrB(0) + ChrB(0)
        
    W1 = AscB(wStr): W2 = AscB(MidB(wStr, 2))
    
    'Debug.Print W1, W2
    If W2 > 127 Then
        jAscW = W1 - 256 + (W2 - 255) * 256
    Else
        jAscW = W1 + W2 * 256
    End If
End Function

Public Function jChrbW(ByVal wVal As Long) As String
    Dim W1 As Integer, W2 As Integer
    
    Debug.Print wVal
    
    If wVal < 65536 And wVal > -32769 Then
        If wVal < 0 Then
            W2 = Int(wVal / 256) + 255: wVal = wVal Mod 256
            W1 = wVal + 256
        Else
            W2 = Int(wVal / 256): wVal = wVal Mod 256
            W1 = wVal
        End If
        'Debug.Print W1, W2
        jChrbW = ChrB(W1) + ChrB(W2)
    End If
End Function

Public Function jMakeString(ByVal wStr As String)
    Dim Wa As Integer
    For Wa = 1 To LenB(wStr)
        If AscB(MidB(wStr, Wa)) = 0 Then Exit For
        jMakeString = jMakeString + Chr(AscB(MidB(wStr, Wa)))
    Next Wa
End Function

Public Function jRange(ByVal wMin As Double, ByVal wVal As Double, ByVal wMax As Double) As Double
    '
    ' Range function..
    '
    Dim Wa As Double
    If wMin > wMax Then Wa = wMin: wMin = wMax: wMax = Wa
    If wVal > wMax Then wVal = wMax
    If wVal < wMin Then wVal = wMin
    jRange = wVal
End Function

Public Function jMin(ByRef wVal1 As Variant, ByRef wVal2 As Variant) As Variant
    jMin = wVal1
    If wVal2 < wVal1 Then jMin = wVal2
End Function

Public Function jMax(ByRef wVal1 As Variant, ByRef wVal2 As Variant) As Variant
    jMax = wVal1
    If wVal2 > wVal1 Then jMax = wVal2
End Function

Public Function jBetween(ByVal wMin As Variant, ByVal wVal As Variant, ByVal wMax As Variant) As Boolean
    jBetween = False
    If wMin < wVal And wVal < wMax Then jBetween = True
End Function
           


Public Sub WAVStop()
    Call WAVPlay(vbNullString)
End Sub


Public Sub WAVLoop(ByVal File)
    Dim SoundName As String
    SoundName$ = File
    wFlags% = &HB
    X = sndPlaySound(SoundName$, wFlags%)
End Sub


Public Sub WAVPlay(ByVal File)
    Dim SoundName As String, wFlags As Long
    SoundName = File
    wFlags = &H3
    X = sndPlaySound(SoundName, wFlags)
End Sub

Public Function jFrmIsLoaded(wFormName As String) As Boolean
  'This function returns true if a form is loaded
  '
  'Parameters:
  '   Formname: Name of the form
  '
  
  Dim wi As Integer
  wFormName = UCase$(wFormName)
  jFrmIsLoaded = False
  For wi = 0 To (Forms.Count - 1)
    If UCase$(Forms(wi).Name) = wFormName Then
      jFrmIsLoaded = True
      Exit For
    End If
  Next wi
End Function

Public Function jSecsToTime(wSecs As Long) As Date
    '
End Function
