Attribute VB_Name = "TTDXeditProcs"
Option Explicit
Option Compare Text

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetUNIXTime Lib "TTDXHelp.dll" () As Long
Declare Function CheckMenuRadioItem Lib "user32" (ByVal hMenu As Long, ByVal un1 As Long, ByVal un2 As Long, ByVal un3 As Long, ByVal un4 As Long) As Boolean
Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal NPos As Long) As Long
Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal NPos As Long) As Long
Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long

Declare Function ShellExecuteEx Lib "Shell32.dll" Alias "ShellExecuteExA" (lpSEI As SHELLEXECUTEINFO) As Long
Declare Function ShellExecuteElevated Lib "elevate.dll" Alias "ShellExecuteElevatedA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function ShellExecuteExElevated Lib "elevate.dll" Alias "ShellExecuteExElevatedA" (lpSEI As SHELLEXECUTEINFO) As Long

Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Declare Function IsUserAnAdmin Lib "TTDXHelp.dll" () As Long

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function GetTempFileNameAPI Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Global Const MF_BYPOSITION = &H400&

Global Const BCM_FIRST = &H1600&
Global Const BCM_SETSHIELD = (BCM_FIRST + &HC&)

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Type SHELLEXECUTEINFO
    cbSize        As Long
    fMask         As Long
    hwnd          As Long
    lpVerb        As String
    lpFile        As String
    lpParameters  As String
    lpDirectory   As String
    nShow         As Long
    hInstApp      As Long
    lpIDList      As Long     'Optional
    lpClass       As String   'Optional
    hkeyClass     As Long     'Optional
    dwHotKey      As Long     'Optional
    hIcon         As Long     'Optional
    hProcess      As Long     'Optional
End Type

Global Const SEE_MASK_NOCLOSEPROCESS = &H40&
Global Const INFINITE = -1&

Private Const Poff As Long = &H52A62

Public Type TTDXgeneral
    GameName As String
    ClimType As Byte
    ClimName As String
    CityNameSet As Byte
    CityNames As String
    VehSize As Byte
End Type
Public Type TTDXlandscape
    X As Integer
    Y As Integer
    Owner As Byte
    Object As Byte
    Height As Byte
    L1 As Byte
    L2 As Byte
    L3 As Long
    L4 As Byte
    L5 As Byte
End Type
Public Type FinanceInfo
    Construction As Long
    NewVehicles As Long
    TrainRunningCosts As Long
    RoadVehRunningCosts As Long
    AircraftRunningCosts As Long
    ShipRunningCosts As Long
    PropertyMaintenance As Long
    TrainIncome As Long
    RoadVehIncome As Long
    AircraftIncome As Long
    ShipIncome As Long
    LoanInterest As Long
    Other As Long
End Type
Public Type TTDXplayer
    Number As Integer
    Id As Integer
    Money As Long
    Debt As Long
    HQx As Integer
    HQy As Integer
    Offset As Long
    FinancesThisYear As FinanceInfo
    FinancesLastYear As FinanceInfo
    FinancesTwoYearsAgo As FinanceInfo
End Type
Public Type TTDXIndInfo
    Number As Integer
    X As Byte
    Y As Byte
    W As Byte
    H As Byte
    Type As Byte
    HomeTown As Byte
    Prod(1) As Byte
    ProR(1) As Byte
    del(2) As Byte
    Name As String
    Offset As Long
End Type
Public Type TTDXCitInfo
    Number As Integer
    X As Byte
    Y As Byte
    Population As Long
    CRate(7) As Long
    CRateE(7) As Boolean
    Name As String
    Offset As Long
End Type
Public Type TTDXStation
    Number As Integer
    BaseX As Byte
    BaseY As Byte
    Parts As Byte
    BusX As Byte
    BusY As Byte
    BusStatus As Byte
    TruckX As Byte
    TruckY As Byte
    TruckStatus As Byte
    RailX As Byte
    RailY As Byte
    RailDir As Boolean
    RailTracks As Byte
    RailTrackLen As Byte
    AirX As Byte
    AirY As Byte
    AirportType As Byte
    DockX As Byte
    DockY As Byte
    Cargo(11) As Integer
    CRate(11) As Byte
    CAcc(11) As Byte
    CEnrout(11) As Byte
    CEnroutOrg(11) As Byte
    Cgood2(11) As Byte
    Cgood5(11) As Byte
    Cgood6(11) As Byte
    Cgood7(11) As Byte
    HomeTown As Byte
    Owner As Byte
    Name As String
    Offset As Long
End Type
Public Type TTDXVehicle
    Number As Long
    Class As Byte
    SubClass As Byte
    Owner As Byte
    CargoT As Byte
    CargoMax As Long
    CargoCur As Long
    SpeedMax As Long
    Value As Long
    Age As Long
    AgeMax As Long
    Rel As Byte
    RelDropRate As Long
    Type As Byte
    Next As Long
    Name As String
    Offset As Long
End Type

Public CurFile As String, FileChanged As Boolean
Public CargoTypes(25) As String, IndustryTypes(50) As String
Public Cities(70) As String, Stations(250) As String

Private F As New FileSystemObject
'
' The importaint internal values
'
Public wData() As Byte, wHeadData(46) As Byte
Private wTTDPext As Integer, wClimate As Integer
Private wError As Integer, wExtraChunks

Global CurrencyMultiplier As Double
Global CurrencyLabel As String
Global CurrencySeparator As String
Global CurrencySymbolBefore As Boolean

Function GetTempFileName()
    On Error GoTo Error
    
    Dim TmpPath As String * 250
    Dim TmpFN As String * 260
    
    GetTempPath 250, TmpPath
    
    If GetTempFileNameAPI(TmpPath, "TDX", 0, TmpFN) = 0 Then
        MsgBox "Unable to create temporary filename.", vbCritical, "Error"
        Exit Function
    Else
        GetTempFileName = APITrim(TmpFN)
    End If
    
    Exit Function
Error:
    Select Case ErrorProc(Err, "Function: TTDXeditProcs.GetTempFileName")
        Case 3:
            End
        Case 2:
            Resume Next
        Case 1:
            Resume
    End Select
End Function

Function APITrim(Buf As String) As String
    On Error GoTo Error
    
    Dim Tmp As String
    Dim NullPos As Integer
       
    Buf = Trim$(Buf)
    NullPos = InStr(1, Buf, Chr$(0), 0)
    
    If NullPos <> 0 Then
        Tmp = Left$(Buf, NullPos - 1)
    Else
        Tmp = Buf
    End If
    
    If Right$(Tmp, 1) = Chr$(0) Then
        Tmp = Left$(Tmp, Len(Tmp) - 1)
    End If
    
    APITrim = Tmp
    Exit Function
Error:
    Select Case ErrorProc(Err, "Function: TTDXeditProcs.APITrim(""" & Buf & """)")
        Case 3:
            End
        Case 2:
            Resume Next
        Case 1:
            Resume
    End Select
End Function



Function IsElevated() As Boolean
    If RunningWin9x() = True Then
        IsElevated = True
    Else
        If IsUserAnAdmin() = 1 Then
            IsElevated = True
        Else
            IsElevated = False
        End If
    End If
End Function


Public Sub RegisterSGMPlugin(ByVal MajorVer As Double, ByVal DllPath As String, ByVal SGMPath As String)
    Dim Wa As Long, Wb As Double
    Wb = Shell("Regsvr32 /s " + Chr(34) + DllPath + Chr(34))
    
    If MajorVer < 2.3 Then
        Wa = fWriteValue(F.BuildPath(SGMPath, "plugins2.ini"), "TTDXEdit", "Class", "S", "SGMTTDXEdit")
        Wa = fWriteValue(F.BuildPath(SGMPath, "plugins2.ini"), "TTDXEdit", "Enabled", "S", 1)
        Wa = fWriteValue(F.BuildPath(SGMPath, "plugins2.ini"), "TTDXEdit", "Filename", "S", DllPath)
        Wa = fWriteValue(F.BuildPath(SGMPath, "plugins2.ini"), "Plugins", "TTDXEdit", "S", "TTDXEdit")
    Else
        Wa = fWriteValue("HKLM", "Software\Owen Rudge\Transport Tycoon Saved Game Manager\Plugins\TTDXEdit", "Class", "S", "SGMTTDXEdit")
        Wa = fWriteValue("HKLM", "Software\Owen Rudge\Transport Tycoon Saved Game Manager\Plugins\TTDXEdit", "Filename", "S", DllPath)
        Wa = fWriteValue("HKLM", "Software\Owen Rudge\Transport Tycoon Saved Game Manager\Plugins\TTDXEdit", "Type", "S", "COM")
    End If
End Sub


Public Sub RegisterSGMPluginStartup()
    Dim Wsa As String, wFl As Boolean, SGMVersion As String, Pos As Integer
    Dim MajorVer As Double, DllPath As String, SGMPath As String
    
    SGMPath = fReadValue("HKLM", "Software\Owen Rudge\InstalledSoftware\TTSGM", "Path", "S", "")
    DllPath = F.BuildPath(App.Path, "SGMPlugIn\TTDXEdit.dll")
    
    If F.FileExists(DllPath) = False Then
        Exit Sub
    End If
    
    SGMVersion = fReadValue("HKLM", "Software\Owen Rudge\InstalledSoftware\TTSGM", "Version", "S", "")
    
    On Error Resume Next
    Pos = InStr(3, SGMVersion, ".")
    MajorVer = CDbl(Left(SGMVersion, Pos - 1))
    On Error GoTo 0
    
    If MajorVer <= 0 Then
        Exit Sub
    End If
    
    RegisterSGMPlugin MajorVer, DllPath, SGMPath
End Sub
Public Function StartElevated(ByVal hwnd As Long, ByVal AppName As String, ByVal Params As String, ByVal WorkingDir As String, ByVal Show As Integer, ByVal Message As String) As Boolean
    On Error GoTo Error
    
    Dim sei As SHELLEXECUTEINFO
    Dim osinfo As OSVERSIONINFO
    Dim retvalue As Integer
    Dim RetVal As Long
    
    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)
    
    sei.cbSize = Len(sei)
    sei.fMask = SEE_MASK_NOCLOSEPROCESS
    sei.hwnd = hwnd
    sei.lpVerb = "open"
    sei.lpFile = AppName
    sei.lpParameters = Params
    sei.lpDirectory = WorkingDir
    sei.nShow = Show
    
    If osinfo.dwPlatformId = 2 Then
        RetVal = IsUserAnAdmin()
        
        If RetVal = 0 Then
            If MsgBox(Message, vbExclamation Or vbYesNo, "TTDX Editor") = vbNo Then
                StartElevated = False
                Exit Function
            End If
            
            If osinfo.dwMajorVersion >= 6 Then
                RetVal = ShellExecuteExElevated(sei)
            Else
                sei.lpVerb = "runas"
                RetVal = ShellExecuteEx(sei)
            End If
        
            GoTo WaitForTermination
        End If
    End If
    
    RetVal = ShellExecuteEx(sei)

WaitForTermination:
    If RetVal = 1 Then
        WaitForSingleObject sei.hProcess, INFINITE
        StartElevated = True
    End If
    
    StartElevated = False
    Exit Function
    
Error:
    Select Case ErrorProc(Err, "Function: TTDXeditProcs.StartElevated(" & hwnd & ", """ & AppName & """, """ & Params & """, """ & WorkingDir & """, " & Show & ")")
        Case 3:
            End
        Case 2:
            Resume Next
        Case 1:
            Resume
    End Select
End Function

Public Function RunningWin9x() As Boolean
    On Error GoTo Error
    
    Dim osinfo As OSVERSIONINFO
    Dim retvalue As Integer
    
    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)
    
    If osinfo.dwPlatformId = 1 Then
        RunningWin9x = True
    Else
        RunningWin9x = False
    End If
    
    Exit Function
Error:
    Select Case ErrorProc(Err, "Function: TTDXeditProcs.RunningWin9x()")
        Case 3:
            End
        Case 2:
            Resume Next
        Case 1:
            Resume
    End Select
End Function
Function MakePath(Path As String) As String
    On Error Resume Next
    
    If Right$(Path, 1) = "\" Then
        MakePath = Path
    Else
        MakePath = Path & "\"
    End If
End Function
Function FlipSign(ByVal Value As Long) As Long
    FlipSign = -Value
    
    'If Value > 0 Then
        'FlipSign = 0 - Value
    'Else
        'FlipSign = 0 - ((2 * Value))
    'End If
End Function

Function FormatMoney(ByVal Money As Long) As String
    On Error GoTo Error
    
    Dim MoneyStr As String
    Dim NewMoneyStr As String
    Dim SetOf3 As Byte
    Dim i As Integer
    Dim Negative As Byte
    
    MoneyStr = CStr(CCur(Money * CurrencyMultiplier))
    
    If Left$(MoneyStr, 1) = "-" Then
        MoneyStr = Right$(MoneyStr, Len(MoneyStr) - 1)
        Negative = True
    Else
        Negative = False
    End If
    
    NewMoneyStr = ""
    SetOf3 = 0
    
    For i = Len(MoneyStr) To 1 Step -1
        SetOf3 = SetOf3 + 1
        
        NewMoneyStr = MID$(MoneyStr, i, 1) & NewMoneyStr
        
        If SetOf3 = 3 And i <> 1 Then
            NewMoneyStr = CurrencySeparator & NewMoneyStr
            SetOf3 = 0
        End If
    Next i
    
    If CurrencySymbolBefore = True Then
        FormatMoney = CurrencyLabel & NewMoneyStr
'        FormatMoney = CurrencyLabel & Format(Money * CurrencyMultiplier, "###" & CurrencySeparator & "###" & CurrencySeparator & "###" & CurrencySeparator & "###" & CurrencySeparator & "###")
    Else
        FormatMoney = NewMoneyStr & CurrencyLabel
        'FormatMoney = Format(Money * CurrencyMultiplier, "###,###,###,###,###") & CurrencyLabel
    End If
    
    If Negative = True Then
        FormatMoney = "-" & FormatMoney
    End If
    
    Exit Function
Error:
    Select Case ErrorProc(Err, "Function: TTDXeditProcs.FormatMoney(" & Money & ")")
        Case 3:
            End
        Case 2:
            Resume Next
        Case 1:
            Resume
    End Select
End Function


Function ErrorProc(Code As Long, Optional Extra)
    Dim ErrDesc As String
    ErrDesc = Err.Description
    
    On Error GoTo Error
    
    Dim mb As Integer
    Dim ExtraText As String
    
    If IsMissing(Extra) = False Then
        ExtraText = vbCrLf & vbCrLf & Extra
    Else
        ExtraText = ""
    End If
    
    mb = MsgBox("The following error occurred:" & vbCrLf & vbCrLf & ErrDesc & vbCrLf & "Error Number " & CStr(Code) & ExtraText, vbCritical Or vbAbortRetryIgnore, "TTDX Editor")
    
    If mb = vbAbort Then
        ErrorProc = 3
    ElseIf mb = vbRetry Then
        ErrorProc = 1
    ElseIf mb = vbIgnore Then
        ErrorProc = 2
    End If
    
'    ErrorProc = mb
    Exit Function
Error:
    Select Case ErrorProc(Err, "WARNING! This error occurred during the error procedure. Something is very wrong!")
        Case 3:
            End
        Case 2:
            Resume Next
        Case 1:
            Resume
    End Select
End Function

Public Function TTDXLoadFile(ByVal vPath As String) As Integer
    Dim Wa As Long, Wb As Long, Wc As Integer, Wd As Integer, wWork() As Byte
    Dim Wsa As String, NoMore As Boolean, Tmpstr As String, TmpPos As Long
    Dim j As Integer
    
    TTDXLoadFile = 0: Wa = 0: Wb = 0
    
    If Not F.FileExists(vPath) Then TTDXLoadFile = 1: Exit Function
     If (vPath Like "*.s??hdr") Or (vPath Like "*.s??dta") Then
        '
        ' Loading an uncompressed file
        '
        Wsa = Left(vPath, Len(vPath) - 3)
        If F.FileExists(Wsa + "hdr") And F.FileExists(Wsa + "dta") Then
            Erase wData: CurFile = "": wTTDPext = 0
            Wb = F.GetFile(Wsa + "dta").Size
        End If
        If Wb < 49 Then TTDXLoadFile = 1: Exit Function
        If (Wb Mod 108800) <> 74873 Then TTDXLoadFile = 2: Exit Function
        '
        ' Reserve Mem and Load Data
        '
        ReDim wData(Wb - 1)
        Open Wsa + "hdr" For Binary As 1
        Get 1, , wHeadData()
        Close 1
        Open Wsa + "dta" For Binary As 1
        Get 1, , wData()
        Close 1
        vPath = Wsa
    ElseIf InStr(".sv0.sv1.sv2.ss0.ss1.", "." + F.GetExtensionName(vPath) + ".") > 0 Then
        '
        ' Load a normal savegame
        '
        Erase wData: CurFile = "": wTTDPext = 0
        Wa = F.GetFile(vPath).Size
        '
        ' Load The Data
        '
        ReDim wWork(Wa - 1)
        Open vPath For Binary As 1
        Get 1, , wWork()
        Close 1
        '
        ' Copy Headerinformation
        '
        For Wa = 0 To UBound(wHeadData)
            wHeadData(Wa) = wWork(Wa)
        Next Wa
        '
        ' Decompression
        '
        ReDim wData(618873 + 256)
        Wa = 49: Wb = 0
        While Wa < UBound(wWork) - 4
            '
            ' Increase dataarea if needed
            '
            If Wb > UBound(wData) - 256 Then ReDim Preserve wData(UBound(wData) + 108800)
            Wc = wWork(Wa)
            If Wc > 127 Then
                Wc = Abs(Wc - 256): Wa = Wa + 1
                For Wd = 0 To Wc
                    wData(Wb) = wWork(Wa): Wb = Wb + 1
                Next Wd
                Wa = Wa + 1
            Else
                Wa = Wa + 1
                For Wd = 0 To Wc
                    wData(Wb) = wWork(Wa): Wb = Wb + 1: Wa = Wa + 1
                Next Wd
            End If
            If Not fFastMode Then DoEvents
        Wend
        '
        ' Check result and reduce memory array size to minimum
        '
        If Wa <> UBound(wWork) - 3 Then TTDXLoadFile = 4: Exit Function
        'OCR'If (Wb Mod 108800) <> 74873 Then TTDXLoadFile = 2: Exit Function
        ReDim Preserve wData(Wb - 1)
    Else
        TTDXLoadFile = 5: Exit Function
    End If
    '
    ' Check filesize, check for extended vehicles
    '
    If Wb < 618873 Then
        '
        ' Not enough data
        '
        TTDXLoadFile = 2
    ElseIf Wb = 618873 Then
        '
        ' No extended vehicles
        '
        CurFile = vPath: wTTDPext = 1
        wClimate = wData(&H77131)
    Else
        If GetLong(&H44CB4) = &H70445454 Then
            wExtraChunks = wData(&H44CB8)
        Else
            wExtraChunks = wData(&H24CBC)
        End If
        
        Dim ExtraChunkPos As Long, i As Integer, ExtraChunkLength As Long
        Dim ExtraChunkSingleLength As Long
        Dim MoreVehicles As Long
        Dim Chunk6Broken As Boolean
        
        MoreVehicles = wData(&H24CBA)
        
        If MoreVehicles = 0 Then
            MoreVehicles = 1
        ElseIf MoreVehicles = 1 Then
            MoreVehicles = 2
        End If
        
        'ExtraChunkPos = ((wData(&H24CBA) - 1) * &H1A900) + &H97179
        ExtraChunkPos = ((MoreVehicles - 1) * &H1A900) + &H97179

        'If wData(&H24CBA) = 0 Then
            'ExtraChunkPos = &H97179
        'End If

        For i = 1 To wExtraChunks
            If wData(ExtraChunkPos) = 1 And wData(ExtraChunkPos + 1) = 128 Then
                If GetLong(ExtraChunkPos + 6) And 1073741824 Then '31 Then ' and 2147483648
                    Chunk6Broken = False
                Else
                    Chunk6Broken = True
                End If
            End If
                
            ExtraChunkSingleLength = GetLong(ExtraChunkPos + 2)
            ExtraChunkLength = ExtraChunkLength + ExtraChunkSingleLength + 6
            ExtraChunkPos = ExtraChunkPos + ExtraChunkSingleLength + 6
            
            If Chunk6Broken = True Then
                ExtraChunkPos = ExtraChunkPos + 4
                ExtraChunkLength = ExtraChunkLength + 4
                Chunk6Broken = False
            End If
        Next i
        
'        If Wb = (618873 + ExtraChunkLength) + ((wData(&H24CBA) - 1) * 108800) Then
        If Wb = (618873 + ExtraChunkLength) + ((MoreVehicles - 1) * 108800) Then
            'wTTDPext = wData(&H24CBA)
            wTTDPext = MoreVehicles
            CurFile = vPath
            wClimate = wData(&H77131 + (wTTDPext - 1) * 108800)
        Else
            If Wb = (618873 + ExtraChunkLength) Then
                wTTDPext = 1
                CurFile = vPath
                wClimate = wData(&H77131)
            Else
                TTDXLoadFile = 3
            End If
        End If
    End If
    
#If MarkGame = 1 Then
    'If wData(&H24CCB) >= 1 Then
        'MsgBox "This game has been modified by TTDX Editor " & wData(&H24CCB) & _
           '"." & wData(&H24CCC) & "." & wData(&H24CCD) & "!", vbExclamation, "TTDX Editor"
    'End If
    
    If GetLong(&H44CB4) = &H70445454 Then
        wData(&H44BBA) = App.Major
        wData(&H44BBB) = App.Minor
        wData(&H44BBC) = App.Revision
    Else
        wData(&H24CCB) = App.Major
        wData(&H24CCC) = App.Minor
        wData(&H24CCD) = App.Revision
    End If
#End If
    
#If MarkGameOldCode = 1 Then
    wExtraChunks = wData(&H24CBC)
    NoMore = False
    
    ExtraChunkPos = ((wData(&H24CBA) - 1) * &H1A900) + &H97179
    
    If wData(&H24CBA) > 0 Then
        For i = 1 To wExtraChunks
            If wData(ExtraChunkPos) = 99 Then
                Tmpstr = ""
                TmpPos = 16
                
                For j = 1 To wData(ExtraChunkPos + 6)
                    Tmpstr = Tmpstr & vbCrLf
                    Tmpstr = Tmpstr & _
                             "v" & wData(ExtraChunkPos + TmpPos) & "." _
                                 & wData(ExtraChunkPos + TmpPos + 2) & "." _
                                 & wData(ExtraChunkPos + TmpPos + 4)
                                    
                    Tmpstr = Tmpstr & vbTab & GetLong(ExtraChunkPos + TmpPos + 6)
                    
                    TmpPos = TmpPos + 16
                Next j
                
                MsgBox "This game has been edited using TTDX Editor as follows:" & Tmpstr, vbInformation, "TTDX Editor"
                NoMore = True
            End If
            
            ExtraChunkSingleLength = GetLong(ExtraChunkPos + 2)
            ExtraChunkLength = ExtraChunkLength + ExtraChunkSingleLength + 6
            ExtraChunkPos = ExtraChunkPos + ExtraChunkSingleLength + 6
        Next i
        
        If NoMore = False Then
            wData(&H24CBC) = wData(&H24CBC) + 1
            
            ReDim Preserve wData(UBound(wData) + 32)
            
            wData(ExtraChunkPos) = 99
            PutLong wData, ExtraChunkPos + 2, 26
            wData(ExtraChunkPos + 6) = 1
            ' padding space (8 bytes)
            
            wData(ExtraChunkPos + 16) = App.Major
            wData(ExtraChunkPos + 18) = App.Minor
            wData(ExtraChunkPos + 20) = App.Revision
            PutLong wData, ExtraChunkPos + 22, GetUNIXTime()
            ' padding space -> 32 (6 bytes)
        Else
            ReDim Preserve wData(UBound(wData) + 16)
            
            'ExtraChunkPos = (ExtraChunkPos - ExtraChunkSingleLength) - 6
            PutLong wData, ExtraChunkPos + 2, GetLong(ExtraChunkPos + 2) + 16
            
            wData(ExtraChunkPos + 6) = wData(ExtraChunkPos + 6) + 1
            wData((ExtraChunkPos + 6) - 1 * 16) = App.Major
            wData(((ExtraChunkPos + 6) - 1 * 16) + 2) = App.Minor
            wData(((ExtraChunkPos + 6) - 1 * 16) + 4) = App.Revision
            PutLong wData, ((ExtraChunkPos + 6) - 1 * 16) + 4, GetUNIXTime()
        End If
    End If
#End If
    
    SetVars
End Function
Public Function TTDXLoadError(vNo As Integer) As String
    If vNo = 0 Then TTDXLoadError = "No Error."
    If vNo = 1 Then TTDXLoadError = "File Not Found."
    If vNo = 2 Then TTDXLoadError = "Unexpected Filesize."
    If vNo = 3 Then TTDXLoadError = "Unexpected Extended Vehicle Array Size."
    If vNo = 4 Then TTDXLoadError = "Decompression Error, Unexpected EOF."
    If vNo = 5 Then TTDXLoadError = "Unknown fileformat."
End Function

Public Function TTDXSaveFile(ByVal vPath As String) As Integer
    Dim Wa As Long, Wb As Long, Wc As Long, Wd As Byte, We As Byte, wWork() As Byte
    Dim lUncompressedDataSize As Long
    Dim bTempByte As Byte
    Dim wHeadCheck(1) As Byte
    TTDXSaveFile = 0
    On Error GoTo TTDXsaveFileErr
    On Error GoTo 0
    If vPath > " " Then
        If F.FileExists(vPath) Then F.DeleteFile vPath
        '
        ' Reserve workarea, calculate header checksum and copy the header
        '
        ReDim wWork(UBound(wData))
        Wc = TTDXCalcHdCheck(wHeadData)
        wWork(47) = Wc Mod 256
        wWork(48) = Fix(Wc / 256)
        For Wa = 0 To 46: wWork(Wa) = wHeadData(Wa): Next Wa
        Wc = 49: Wb = 0
        '
        ' Compress into RLE data
        '
             
' Old code
'        While Wb <= UBound(wData)
'            If wData(Wb) = wData(Wb + 1) Then
'                We = 124
'                For Wd = 2 To 124
'                    If Wb + Wd > UBound(wData) Then We = Wd - 1: Exit For
'                    If wData(Wb) <> wData(Wb + Wd) Then We = Wd - 1: Exit For
'                Next Wd
'                wWork(Wc) = 256 - We: wWork(Wc + 1) = wData(Wb): Wc = Wc + 2: Wb = Wb + We + 1
'            Else
'                We = 124
'                For Wd = 0 To 124
'                    If Wb + Wd > UBound(wData) Then We = Wd - 1: Exit For
'                    If wData(Wb + Wd) = wData(Wb + Wd + 1) Then We = Wd - 1: Exit For
'                Next Wd
'                wWork(Wc) = We: Wc = Wc + 1
'                For Wd = 0 To We: wWork(Wc) = wData(Wb): Wc = Wc + 1: Wb = Wb + 1: Next Wd
'            End If
'            If Not fFastMode Then DoEvents ' Give some room to other programs running
'        Wend
'


' New Code based on old code
' Wc current byte in outputstream
' Wb current byte in inputstream
        lUncompressedDataSize = UBound(wData)
        While Wb <= lUncompressedDataSize
            If (wData(Wb) = wData(Wb + 1)) Then
                bTempByte = wData(Wb)
                If (Wb + 124 >= lUncompressedDataSize) Then
                    We = lUncompressedDataSize - Wb
                    For Wd = 2 To We
                        If bTempByte <> wData(Wb + Wd) Then We = Wd - 1: Exit For
                    Next Wd
                Else
                    We = 124
                    For Wd = 2 To 124
                        If bTempByte <> wData(Wb + Wd) Then We = Wd - 1: Exit For
                    Next Wd
                End If
'                We = 124
'                For Wd = 2 To 124
'                    If Wb + Wd > lUncompressedDataSize Then We = Wd - 1: Exit For
'                    If wData(Wb) <> wData(Wb + Wd) Then We = Wd - 1: Exit For
'                Next Wd
                wWork(Wc) = 256 - We        ' write down how many bytes are same
                wWork(Wc + 1) = wData(Wb)  ' the actual byte data
                Wc = Wc + 2
                Wb = Wb + We + 1
            Else
                ' Different Bytes
                
                If (Wb + 124 >= lUncompressedDataSize) Then
                    We = lUncompressedDataSize - Wb
                    For Wd = 0 To We - 1
                        If wData(Wb + Wd) = wData(Wb + Wd + 1) Then We = Wd - 1: Exit For
                    Next Wd
                Else
                    We = 124
                    For Wd = 0 To 124
                        If wData(Wb + Wd) = wData(Wb + Wd + 1) Then We = Wd - 1: Exit For
                    Next Wd
                End If
                               
'                We = 124
'                For Wd = 0 To 124
'                    If Wb + Wd > lUncompressedDataSize Then We = Wd - 1: Exit For
' // new crash fix:
'                    If Wb + Wd = lUncompressedDataSize Then We = Wd: Exit For
' //
'                    If wData(Wb + Wd) = wData(Wb + Wd + 1) Then We = Wd - 1: Exit For
'                Next Wd

                wWork(Wc) = We              ' write down how many different bytes we have
                Wc = Wc + 1

                For Wd = 0 To We            ' write the actuall different bytes
                    wWork(Wc) = wData(Wb)
                    Wc = Wc + 1:
                    Wb = Wb + 1:
                Next Wd
            End If
            If Not fFastMode Then DoEvents ' Give some room to other programs running
        Wend
        '
        ' Set proper size and calculate checksum
        '
        ReDim Preserve wWork(Wc - 1)
        Wc = 0
        For Wa = 0 To UBound(wWork)
            Wc = (Wc And &HFFFFFF00) Or ((Wc And &HFF) + wWork(Wa) And &HFF)
            If Wc And &H10000000 Then
                Wc = (((Wc And &HFFFFFFF) * 8) Or ((Wc And &HE0000000) / 2 ^ 29) And &H7) Or &H80000000
            Else
                Wc = ((Wc And &HFFFFFFF) * 8) Or ((Wc And &HE0000000) / 2 ^ 29) And &H7
            End If
        Next Wa
        '
        ' Puttin´ on the disc
        '
        Open vPath For Binary As 1
        Put 1, , wWork()
        Put 1, , CLng(Wc + 201100)
        Close 1
        FileChanged = False
    Else
        TTDXSaveFile = 1
    End If
    Exit Function
    
TTDXsaveFileErr:
    ' would be good to look what error was here...
    TTDXSaveFile = 1
End Function
Public Function TTDXCalcHdCheck(vHd() As Byte) As Long
    Dim Wc As Long, Wa As Integer
    Wc = 0
    For Wa = 0 To 46
        Wc = (Wc + vHd(Wa)) * 2
        If Wc And &H10000 Then Wc = (Wc Or 1) And &HFFFF&
    Next Wa
    TTDXCalcHdCheck = (Wc Xor &HAAAA&)
End Function
Public Function TTDXSaveUncom(ByVal vPath As String) As Integer
    Dim Wsa As String, Wsb As String, Wc As Long

    TTDXSaveUncom = 0
    If vPath > " " Then
        Wsa = F.GetExtensionName(vPath)
        If InStr(".sv0.sv1.sv2.ss0.ss1.", "." + Wsa + ".") Then
        Else
            vPath = vPath + ".sv1"
        End If
        If F.FileExists(vPath + "hdr") Then F.DeleteFile (vPath + "hdr")
        Wc = TTDXCalcHdCheck(wHeadData)
        Open vPath + "hdr" For Binary As 1
        Put 1, , wHeadData()
        Put 1, , CByte(Wc Mod 256)
        Put 1, , CByte(Fix(Wc / 256))
        Close 1
        If F.FileExists(vPath + "dta") Then F.DeleteFile (vPath + "dta")
        Open vPath + "dta" For Binary As 1
        Put 1, , wData()
        Close 1
    End If
    Exit Function
TTDXsaveFileErr:
    TTDXSaveUncom = 1
End Function


'****************************************************************************************************
'**** Landscape Procedures                                                                       ****
'****************************************************************************************************

Public Function TTDXgetLandscape(Wx As Integer, Wy As Integer) As TTDXlandscape
    Dim Wa As Integer, Woff As Long
    If Wx < 0 Or Wy < 0 Then TTDXgetLandscape.Object = 7: Exit Function
    If Wx > 255 Or Wy > 255 Then TTDXgetLandscape.Object = 7: Exit Function
    Woff = Wx + 256& * Wy
    If CurFile > " " Then
        With TTDXgetLandscape
            .X = Wx
            .Y = Wy
            .Owner = wData(&H4CBA + Woff)
            .Object = (wData(&H77179 + (wTTDPext - 1) * 108800 + Woff) And &HF0) / 16
            .Height = wData(&H77179 + (wTTDPext - 1) * 108800 + Woff) And &HF
            .L1 = wData(&H4CBA + Woff)
            .L2 = wData(&H14CBA + Woff)
            .L3 = wData(&H24CBA + Woff * 2) + wData(&H24CBB + Woff * 2) * 256&
            .L4 = wData(&H77179 + (wTTDPext - 1) * 108800 + Woff)
            .L5 = wData(&H87179 + (wTTDPext - 1) * 108800 + Woff)
        End With
    End If
End Function
Public Sub TTDXputLandscape(wDta As TTDXlandscape)
    Dim Wa As Byte, Woff As Long
    With wDta
        Woff = .X + 256& * .Y
        If .Owner <> wData(&H4CBA + Woff) Then FileChanged = True: wData(&H4CBA + Woff) = .Owner
        Wa = (.Object And &HF) * 16 + (.Height And &HF)
        If Wa <> wData(&H77179 + (wTTDPext - 1) * 108800 + Woff) Then FileChanged = True: wData(&H77179 + (wTTDPext - 1) * 108800 + Woff) = Wa
        If .L2 <> wData(&H14CBA + Woff) Then FileChanged = True: wData(&H14CBA + Woff) = .L2
        If .L3 <> wData(&H24CBA + Woff * 2) + wData(&H24CBB + Woff * 2) * 256& Then
            wData(&H24CBA + Woff * 2) = CByte(.L3 Mod 256)
            wData(&H24CBB + Woff * 2) = CByte(Fix(.L3 / 256) And &HFF&)
            FileChanged = True
        End If
        If .L5 <> wData(&H87179 + (wTTDPext - 1) * 108800 + Woff) Then FileChanged = True: wData(&H87179 + (wTTDPext - 1) * 108800 + Woff) = .L5
    End With
End Sub
Public Function TTDXGetLandMap(Wx As Integer, Wy As Integer) As Integer
    TTDXGetLandMap = (wData(&H77179 + (wTTDPext - 1) * 108800 + Wx + Wy * 256&) And &HF0&) / 16
End Function

'****************************************************************************************************
'**** City procedures                                                                            ****
'****************************************************************************************************

Public Function CityData(vCityNo As Integer, vDataoff As Integer) As Byte
    Dim Wa As Long
    CityData = wData(&H264 + &H5E * vCityNo + vDataoff)
End Function
Public Function CityInfo(vCityNo As Integer) As TTDXCitInfo
    Dim Wa As Integer
    '
    ' Return information about a city
    '
    With CityInfo
        .Number = vCityNo
        .Offset = &H264 + &H5E * vCityNo
        .X = CityData(vCityNo, 0)
        .Y = CityData(vCityNo, 1)
        .Population = CityData(vCityNo, 2) + CityData(vCityNo, 3) * 256&
        .Name = Cities(vCityNo)
        For Wa = 0 To 7
            .CRate(Wa) = CityData(vCityNo, Wa * 2 + &H1E) + CityData(vCityNo, Wa * 2 + &H1F) * 256&
            If .CRate(Wa) > 32768 Then .CRate(Wa) = .CRate(Wa) - 65536
            If (CityData(vCityNo, &H2E) And 2 ^ Wa) > 0 Then .CRateE(Wa) = True Else .CRateE(Wa) = False
        Next Wa
        '
    End With
End Function
Public Sub TTDXCityPut(wDta As TTDXCitInfo)
    Dim wWork(&H5D) As Byte, Wa As Integer, Wb As Long
    With wDta
        For Wa = 0 To &H5D: wWork(Wa) = CityData(.Number, Wa): Next Wa
        wWork(&H2E) = 0
        For Wa = 0 To 7
            Wb = .CRate(Wa)
            If Wb < 0 Then Wb = Wb + 65536
            wWork(&H1E + Wa * 2) = CByte(Wb Mod 256)
            wWork(&H1F + Wa * 2) = CByte(Fix(Wb / 256))
            If .CRateE(Wa) Then wWork(&H2E) = wWork(&H2E) Or 2 ^ Wa
        Next Wa
        For Wa = 0 To &H5D
            If wWork(Wa) <> wData(&H264 + &H5E * .Number + Wa) Then
                FileChanged = True: wData(&H264 + &H5E * .Number + Wa) = wWork(Wa)
                'Debug.Print Hex(Wa)
            End If
        Next Wa
    End With
End Sub
Public Function CityName(vCityNo As Integer) As String
    Dim Wa As Long, Wb As Long
    Wb = CityData(vCityNo, 4)
    If CityData(vCityNo, 5) = &H7D Then
        CityName = GetString(Wb)
    ElseIf CityData(vCityNo, 5) = &H20 Then
        If Wb = &HC1 Then
            CityName = TTDXMakeC1CityName(CityData(vCityNo, 6), CityData(vCityNo, 7), CityData(vCityNo, 8), CityData(vCityNo, 9))
        ElseIf Wb = &HC2 Then
            CityName = TTDXMakeC2CityName(CityData(vCityNo, 6), CityData(vCityNo, 7), CityData(vCityNo, 8), CityData(vCityNo, 9))
        End If
    End If
    'If CityName = "" Then: CityName = "<Unknown Name " + Right("00" + Hex(vCityNo), 2) + ">"
End Function

'****************************************************************************************************
'**** Industry procedures                                                                        ****
'****************************************************************************************************

Public Function TTDXIndustryData(vIndNo As Integer, vDataoff As Integer) As Byte
    Dim Wa As Long
    TTDXIndustryData = wData(333670 + &H36 * vIndNo + vDataoff)
End Function

Public Function TTDXIndustryInfo(vIndNo As Integer) As TTDXIndInfo
    Dim Wa As Integer
    '                                          1               2               3
    With TTDXIndustryInfo
        .Number = vIndNo
        .X = TTDXIndustryData(vIndNo, 0)
        .Y = TTDXIndustryData(vIndNo, 1)
        .W = TTDXIndustryData(vIndNo, 6)
        .H = TTDXIndustryData(vIndNo, 7)
        .HomeTown = CByte((((TTDXIndustryData(vIndNo, 2) + TTDXIndustryData(vIndNo, 3) * 256) - &H264) / &H5E) And &HFF)
        .Prod(0) = TTDXIndustryData(vIndNo, 8)
        .Prod(1) = TTDXIndustryData(vIndNo, 9)
        .ProR(0) = TTDXIndustryData(vIndNo, &HE)
        .ProR(1) = TTDXIndustryData(vIndNo, &HF)
        .del(0) = TTDXIndustryData(vIndNo, &H10)
        .del(1) = TTDXIndustryData(vIndNo, &H11)
        .del(2) = TTDXIndustryData(vIndNo, &H12)
        .Type = TTDXIndustryData(vIndNo, &H26)
        .Name = IndustryTypes(.Type)
        
        .Offset = 333670 + &H36 * vIndNo
    End With
End Function
Public Sub TTDXIndustryPut(wDta As TTDXIndInfo)
    Dim wWork(&H35) As Byte, Wa As Integer
    With wDta
        For Wa = 0 To &H35: wWork(Wa) = TTDXIndustryData(.Number, Wa): Next Wa
        wWork(8) = .Prod(0): wWork(9) = .Prod(1)
        wWork(&HE) = .ProR(0): wWork(&HF) = .ProR(1)
        wWork(&H10) = .del(0): wWork(&H11) = .del(1): wWork(&H12) = .del(2)
        wWork(2) = CByte(((.HomeTown * &H5E) + &H264) Mod 256)
        wWork(3) = CByte(Fix(((.HomeTown * &H5E) + &H264) / 256))
        For Wa = 0 To &H35
            If wWork(Wa) <> wData(333670 + &H36 * .Number + Wa) Then
                FileChanged = True: wData(333670 + &H36 * .Number + Wa) = wWork(Wa)
            End If
        Next Wa
    End With
End Sub

'****************************************************************************************************
'**** Station procedures                                                                         ****
'****************************************************************************************************
Public Function TTDXStationInfo(wNo As Integer) As TTDXStation
    Dim sOff As Long, Wa As Integer, Wx As Byte, Wy As Byte
    
    sOff = &H48CBA + wNo * &H8E&

    With TTDXStationInfo
        .BaseX = wData(sOff)
        .BaseY = wData(sOff + 1)
        '.AreaX1 = 255: .AreaX2 = 0: .AreaY1 = 255: .AreaY2 = 0
        .Number = wNo
        '
        .Parts = wData(sOff + &H80)
        .TruckStatus = wData(sOff + &H82)
        .BusStatus = wData(sOff + &H83)
        .BusX = wData(sOff + 6): .BusY = wData(sOff + 7)
        .TruckX = wData(sOff + 8): .TruckY = wData(sOff + 9)
        .RailX = wData(sOff + 10): .RailY = wData(sOff + 11)
        .AirX = wData(sOff + 12): .AirY = wData(sOff + 13)
        .DockX = wData(sOff + 14): .DockY = wData(sOff + 15)
        .RailTracks = CByte(wData(sOff + 16) And 7)
        .RailTrackLen = CByte((wData(sOff + 16) And 248) / 8)
        If TTDXgetLandscape(CInt(Wx), CInt(Wy)).L5 And 1& Then: .RailDir = False: Else: .RailDir = True
        .AirportType = wData(sOff + 81)
                    
        For Wa = 0 To 11
            .Cargo(Wa) = CInt((wData(sOff + &H1C + Wa * 8) + wData(sOff + &H1D + Wa * 8) * 256&) And &HFFF)
            .CAcc(Wa) = (wData(sOff + &H1D + Wa * 8) And &HF0) / 16
            .Cgood2(Wa) = wData(sOff + &H1C + Wa * 8 + 2)
            .CRate(Wa) = wData(sOff + &H1C + Wa * 8 + 3)
            .CEnrout(Wa) = wData(sOff + &H1C + Wa * 8 + 4): .CEnroutOrg(Wa) = .CEnrout(Wa)
            .Cgood5(Wa) = wData(sOff + &H1C + Wa * 8 + 5)
            .Cgood6(Wa) = wData(sOff + &H1C + Wa * 8 + 6)
            .Cgood7(Wa) = wData(sOff + &H1C + Wa * 8 + 7)
        Next Wa
        .Owner = wData(sOff + &H7F)
        .HomeTown = CByte((((wData(sOff + 2) + wData(sOff + 3) * 256) - &H264) / &H5E) And &HFF)
        .Name = Stations(.Number)
        .Offset = sOff
    End With
End Function
Public Sub TTDXStationPut(wDta As TTDXStation)
    Dim wWork(&H8D) As Byte, Wa As Integer, sOff As Long
    With wDta
        sOff = &H48CBA + wDta.Number * &H8E&
        '
        ' Copy Original Data
        '
        For Wa = 0 To &H8D: wWork(Wa) = wData(sOff + Wa): Next Wa
        '
        wWork(&H80) = .Parts
        wWork(&H82) = .TruckStatus
        wWork(&H83) = .BusStatus
        wWork(6) = .BusX: wWork(7) = .BusY
        wWork(8) = .TruckX: wWork(9) = .TruckY
        '
        For Wa = 0 To 11
            If .CEnrout(Wa) = 255 Then
                .Cargo(Wa) = 0: .Cgood2(Wa) = 0: .CRate(Wa) = &HAF
                .Cgood5(Wa) = 0: .Cgood6(Wa) = 0: .Cgood7(Wa) = 255
            End If
            wWork(&H1C + Wa * 8) = CByte(.Cargo(Wa) Mod 256)
            wWork(&H1D + Wa * 8) = CByte(Fix(.Cargo(Wa) / 256) + .CAcc(Wa) * 16)
            wWork(&H1E + Wa * 8) = .Cgood2(Wa)
            wWork(&H1F + Wa * 8) = .CRate(Wa)
            wWork(&H20 + Wa * 8) = .CEnrout(Wa)
            wWork(&H1C + Wa * 8 + 5) = .Cgood5(Wa)
            wWork(&H1C + Wa * 8 + 6) = .Cgood6(Wa)
            wWork(&H1C + Wa * 8 + 7) = .Cgood7(Wa)
        Next Wa
        '
        ' Check to see if Data has changed
        '
        For Wa = 0 To &H8D
            If wWork(Wa) <> wData(sOff + Wa) Then
                FileChanged = True: wData(sOff + Wa) = wWork(Wa)
            End If
        Next Wa
    End With
End Sub

'****************************************************************************************************
'**** Vehicle procedures                                                                         ****
'****************************************************************************************************
Public Function TTDXGetVeh(wNo As Long) As TTDXVehicle
    Dim sOff As Long
    sOff = &H547F2 + wNo * 128
    
    With TTDXGetVeh
        .Number = wNo
        .Class = wData(sOff)
        'If .Class = 19 Then
            'MsgBox "aircraft"
        'End If
        .SubClass = wData(sOff + 1)
        .SpeedMax = wData(sOff + &H18) + wData(sOff + &H19) * 256&
        .Owner = wData(sOff + &H25)
        .CargoT = wData(sOff + &H39)
        .CargoMax = wData(sOff + &H3A) + wData(sOff + &H3B) * 256&
        .CargoCur = wData(sOff + &H3C) + wData(sOff + &H3D) * 256&
        .Age = wData(sOff + &H40) + wData(sOff + &H41) * 256&
        .AgeMax = wData(sOff + &H42) + wData(sOff + &H43) * 256&
        .Rel = wData(sOff + &H4F)
        .RelDropRate = wData(sOff + &H50) + wData(sOff + &H51) * 256&
        .Next = wData(sOff + &H5A) + wData(sOff + &H5B) * 256&
        .Value = GetLong(sOff + &H5C)
        .Offset = sOff
        .Name = "Unknown"
        Select Case .Class
            Case &H10
                Select Case .SubClass
                    Case 0
                        .Name = "Train " + Format(wData(sOff + &H45), "000")
                End Select
            Case &H11
                .Name = "Road Vehicle " + Format(wData(sOff + &H45), "000")
            Case &H12
                .Name = "Ship " + Format(wData(sOff + &H45), "000")
            Case &H13
                Select Case .SubClass
                    Case 0 ' copter
                        .Name = "Aircraft " + Format(wData(sOff + &H45), "000")
                    Case 2 ' plane
                        .Name = "Aircraft " + Format(wData(sOff + &H45), "000")
                    Case 4
                        .Name = "Aircraft Second " + Format(wData(sOff + &H45), "000")
                    Case 6
                        .Name = "Chopper Rotor " + Format(wData(sOff + &H45), "000")
                End Select
        End Select
'        Select Case .SubClass
'            Case 0
'                Select Case wData(sOff + &H60) + wData(sOff + &H61) * 256&
'                    Case &H8864&: .Name = "Train " + Format(wData(sOff + &H45), "000")
'                    Case &H902B&: .Name = "Road Vehicle " + Format(wData(sOff + &H45), "000")
'                    Case &H9830&: .Name = "Ship " + Format(wData(sOff + &H45), "000")
'                    Case &HA02F&: .Name = "Aircraft " + Format(wData(sOff + &H45), "000")
'                End Select
'            Case 2: .Name = "Car"
'            Case 4: .Name = "Aircraft Second"
'            Case 6: .Name = "Chopper Rotor"
'        End Select
        
    End With
End Function
Public Sub TTDXPutVeh(wDta As TTDXVehicle)
    Dim wWork(&H7F) As Byte, Wa As Integer, sOff As Long
    With wDta
        sOff = &H547F2 + .Number * &H80
        '
        ' Copy Original Data
        '
        For Wa = 0 To &H7F: wWork(Wa) = wData(sOff + Wa): Next Wa
        '
        ' Store values
        '
        wWork(&H39) = .CargoT
        wWork(&H3A) = CByte(.CargoMax Mod 256)
        wWork(&H3B) = CByte(Fix(.CargoMax / 256))
        wWork(&H40) = CByte(.Age Mod 256)
        wWork(&H41) = CByte(Fix(.Age / 256))
        wWork(&H42) = CByte(.AgeMax Mod 256)
        wWork(&H43) = CByte(Fix(.AgeMax / 256))
        wWork(&H4F) = CByte(.Rel)
        wWork(&H50) = CByte(.RelDropRate Mod 256)
        wWork(&H51) = CByte(Fix(.RelDropRate / 256))
        PutLong wWork(), &H5C, .Value
        '
        ' Check to see if Data has changed
        '
        For Wa = 0 To &H7F
            If wWork(Wa) <> wData(sOff + Wa) Then
                FileChanged = True: wData(sOff + Wa) = wWork(Wa)
            End If
        Next Wa
    End With
End Sub

'****************************************************************************************************
'**** General & Player                                                                           ****
'****************************************************************************************************

Public Function TTDXGetByte(wPos) As Byte
    TTDXGetByte = wData(wPos)
End Function
Public Function TTDXGeneralInfo() As TTDXgeneral
    Dim Wa As Integer
    With TTDXGeneralInfo
        .VehSize = wTTDPext
        .ClimType = wClimate
        Select Case wClimate
            Case 0: .ClimName = "Temperate"
            Case 1: .ClimName = "Sub-Arctic"
            Case 2: .ClimName = "Sub-Tropic"
            Case 3: .ClimName = "Toyland"
            Case Else: .ClimName = "<Undefined " + Format(wClimate) + ">"
        End Select
        .CityNameSet = wData(&H5C80D + &H1A900 * .VehSize)
        Select Case .CityNameSet
            Case 0: .CityNames = "English"
            Case 1: .CityNames = "French"
            Case 2: .CityNames = "German"
            Case 3: .CityNames = "American"
            Case 4: .CityNames = "Latin-American"
            Case 5: .CityNames = "Silly"
        End Select
        .GameName = ""
        For Wa = 0 To 46
            If wHeadData(Wa) < 32 Then Exit For
            .GameName = .GameName + Chr(wHeadData(Wa))
        Next Wa
    End With
End Function

Public Function TTDXPlayerInfo(vPlayer As Integer) As TTDXplayer
    Dim Wa As Integer
    '                                          1               2               3               4               5
    '                          0       8       0       8       0       8       0       8       0       8       0       8     E
    Const BytKnow As String = "                xxxxxxxx  "
    
    With TTDXPlayerInfo
        .Number = vPlayer
        .Offset = Poff + vPlayer * 946
        .Id = wData(Poff + vPlayer * 946) + CInt(wData(Poff + vPlayer * 946 + 1))
        .Money = GetLong(Poff + vPlayer * 946 + &H10)
        .Debt = GetLong(Poff + vPlayer * 946 + &H14)
        .HQx = wData(Poff + vPlayer * 946 + &H3A4)
        .HQy = wData(Poff + vPlayer * 946 + &H3A5)
        
        .FinancesThisYear.Construction = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 0))
        .FinancesThisYear.NewVehicles = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 4))
        .FinancesThisYear.TrainRunningCosts = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 8))
        .FinancesThisYear.RoadVehRunningCosts = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 12))
        .FinancesThisYear.AircraftRunningCosts = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 16))
        .FinancesThisYear.ShipRunningCosts = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 20))
        .FinancesThisYear.PropertyMaintenance = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 24))
        .FinancesThisYear.TrainIncome = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 28))
        .FinancesThisYear.RoadVehIncome = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 32))
        .FinancesThisYear.AircraftIncome = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 36))
        .FinancesThisYear.ShipIncome = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 40))
        .FinancesThisYear.LoanInterest = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 44))
        .FinancesThisYear.Other = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 48))
    
        .FinancesLastYear.Construction = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 52 + 0))
        .FinancesLastYear.NewVehicles = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 52 + 4))
        .FinancesLastYear.TrainRunningCosts = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 52 + 8))
        .FinancesLastYear.RoadVehRunningCosts = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 52 + 12))
        .FinancesLastYear.AircraftRunningCosts = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 52 + 16))
        .FinancesLastYear.ShipRunningCosts = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 52 + 20))
        .FinancesLastYear.PropertyMaintenance = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 52 + 24))
        .FinancesLastYear.TrainIncome = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 52 + 28))
        .FinancesLastYear.RoadVehIncome = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 52 + 32))
        .FinancesLastYear.AircraftIncome = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 52 + 36))
        .FinancesLastYear.ShipIncome = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 52 + 40))
        .FinancesLastYear.LoanInterest = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 52 + 44))
        .FinancesLastYear.Other = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 52 + 48))
    
        .FinancesTwoYearsAgo.Construction = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 104 + 0))
        .FinancesTwoYearsAgo.NewVehicles = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 104 + 4))
        .FinancesTwoYearsAgo.TrainRunningCosts = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 104 + 8))
        .FinancesTwoYearsAgo.RoadVehRunningCosts = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 104 + 12))
        .FinancesTwoYearsAgo.AircraftRunningCosts = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 104 + 16))
        .FinancesTwoYearsAgo.ShipRunningCosts = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 104 + 20))
        .FinancesTwoYearsAgo.PropertyMaintenance = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 104 + 24))
        .FinancesTwoYearsAgo.TrainIncome = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 104 + 28))
        .FinancesTwoYearsAgo.RoadVehIncome = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 104 + 32))
        .FinancesTwoYearsAgo.AircraftIncome = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 104 + 36))
        .FinancesTwoYearsAgo.ShipIncome = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 104 + 40))
        .FinancesTwoYearsAgo.LoanInterest = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 104 + 44))
        .FinancesTwoYearsAgo.Other = FlipSign(GetLong(Poff + vPlayer * 946 + &H26 + 104 + 48))
    End With
End Function
Public Function TTDXputPlayer(wDta As TTDXplayer)
    Dim wWork(945) As Byte, Wa As Integer
    With wDta
        For Wa = 0 To 945: wWork(Wa) = wData(Poff + .Number * 946 + Wa): Next Wa
        PutLong wWork(), &H10, .Money
        PutLong wWork(), &H14, .Debt
        wWork(&H3A4) = .HQx
        wWork(&H3A5) = .HQy
        For Wa = 0 To 945
            If wWork(Wa) <> wData(Poff + .Number * 946 + Wa) Then
                FileChanged = True: wData(Poff + .Number * 946 + Wa) = wWork(Wa)
            End If
        Next Wa
    End With
End Function


'****************************************************************************************************
'**** Internal                                                                                   ****
'****************************************************************************************************

Private Sub SetVars()
    Dim Wva As Variant, Wsb As String, Wa As Integer, Wb As Integer, Woff As Long
    '
    ' Cargo types
    '
    Wsb = " | "
    Select Case wClimate
        Case 0: Wsb = "Passengers|Coal|Mail|Oil|Livestock|Goods|Grain|Wood|Iron Ore|Steel|Valuables"
        Case 1: Wsb = "Passengers|Coal|Mail|Oil|Livestock|Goods|Wheat|Wood||Paper|Gold|Food"
        Case 2: Wsb = "Passengers|Rubber|Mail|Oil|Fruit|Goods|Maize|Wood|Copper|Water|Diamonds|Food"
        Case 3: Wsb = "Passengers|Sugar|Mail|Toys|Batteries|Sweets|Toffee|Cola|Candyfloss|Bubbles|Plastics|Fizzy Drinks"
    End Select
    Wva = Split(Wsb, "|")
    For Wa = 0 To UBound(CargoTypes): CargoTypes(Wa) = "<unknown " + Format(Wa) + ">": Next Wa
    For Wa = 0 To UBound(Wva)
        If Trim(CStr(Wva(Wa))) > " " Then CargoTypes(Wa) = CStr(Wva(Wa))
    Next Wa
    '
    ' Industry Types
    '
    Wsb = " | "
    Select Case wClimate
        Case 0: Wsb = "Coal Mine|Power Station|Saw Mill|Forest|Oil Refinery|Oil Rig|Factory||Steel Mill|Farm| |Oil Wells|Bank||||||Iron Ore Mine"
        Case 1: Wsb = "Coal Mine|Power Station||Forest|Oil Refinery|||Printing Works||Farm||Oil Wells||Food Processing Plant|Paper Mill|Gold Mine|Bank"
        Case 2: Wsb = "||||Oil Refinery||||||Copper Ore Mine|Oil Wells||Food Processing Plant|||Bank|Diamond Mine| |Food Plantation|Rubber Plantation|Water Supply|Water Tower|Factory|Farm|Lumber Mill"
        Case 3: Wsb = "||||||||||||||||||||||||||Candyfloss Factory|Sweet Factory|Battery Farm|Cola Wells|Toy Shop|Toy Factory|Plastic Fountains|Fizzy Drinks Factory|Bubble Generator|Toffee Quarry|Sugar Mine"
    End Select
    Wva = Split(Wsb, "|")
    For Wa = 0 To UBound(IndustryTypes): IndustryTypes(Wa) = "<unknown " + Format(Wa) + ">": Next Wa
    For Wa = 0 To UBound(Wva)
        If Trim(CStr(Wva(Wa))) > " " Then IndustryTypes(Wa) = CStr(Wva(Wa))
    Next Wa
    '
    ' Get Citynames
    '
    For Wa = 0 To 69
        Cities(Wa) = CityName(Wa)
    Next Wa
    '
    ' Get StationNames
    '
    For Wa = 0 To 249
        Woff = &H48CBA + Wa * &H8E&
        If wData(Woff + &H15) = &H30 Then
            Wb = (((wData(Woff + 2) + wData(Woff + 3) * 256) - &H264) / &H5E) And &HFF
            Stations(Wa) = Cities(Wb) + TTDXStationExtension(wData(Woff + &H14))
        ElseIf wData(Woff + &H15) = &H7F Then
            Stations(Wa) = GetString(CInt(wData(Woff + &H14)))
        End If
    Next Wa
End Sub
Private Function GetString(wNo As Long) As String
    Dim Wa As Long, Wb As Integer
    GetString = ""
    Wa = (463090 + (wTTDPext - 1) * 108800) + wNo * 32
    For Wb = 0 To 31
        If wData(Wa + Wb) = 0 Then Exit For
        GetString = GetString + Chr(wData(Wa + Wb))
    Next Wb
End Function
Private Function GetLong(wAdr) As Long
    Dim Wda As Double
    Wda = wData(wAdr) + wData(wAdr + 1) * 256# + wData(wAdr + 2) * 256 ^ 2# + wData(wAdr + 3) * 256 ^ 3#
    If Wda > 2147483647# Then
        GetLong = CLng(Wda - 4294967296#)
    Else
        GetLong = CLng(Wda)
    End If
End Function
Private Sub PutLong(ByRef vDta() As Byte, vOff As Long, vVal As Long)
    Dim Wda As Double
    If vVal > -1 Then Wda = vVal Else Wda = vVal + 4294967296#
    vDta(vOff + 3) = Fix(Wda / 2 ^ 24): Wda = Wda - Fix(Wda / 2 ^ 24) * 2 ^ 24
    vDta(vOff + 2) = Fix(Wda / 2 ^ 16): Wda = Wda Mod 2 ^ 16
    vDta(vOff + 1) = Fix(Wda / 2 ^ 8)
    vDta(vOff) = Wda Mod 256
End Sub
