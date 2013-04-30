Attribute VB_Name = "UniInteraction"
'*************************************************************************************************
'* UniInteraction - Unicode & improved equivalents of VBA.Interaction
'* ------------------------------------------------------------------
'* By Vesa Piittinen aka Merri, http://vesa.piittinen.name/ <vesa@piittinen.name>
'*************************************************************************************************
Option Explicit

' Command$
Private Declare Function CommandLineToArgvW Lib "shell32" (ByVal lpCmdLine As Long, pNumArgs As Integer) As Long
Private Declare Function GetCommandLineW Lib "kernel32" () As Long
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal Addr As Long, RetAddr As Long)
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Function SysAllocStringLen Lib "oleaut32" (ByVal BSTR As Long, ByVal WLen As Long) As Long

' MsgBox
Private Const MB_USERICON = &H80&

Private Type MsgBoxParams
    cbSize As Long
    hWndOwner As Long
    hInstance As Long
    lpszText As Long
    lpszCaption As Long
    dwStyle As Long
    lpszIcon As Long
    dwContextHelpId As Long
    lpfnMsgBoxCallback As Long
    dwLanguageId As Long
End Type

Private Declare Function MessageBoxIndirectW Lib "user32" (lpMsgBoxParams As MsgBoxParams) As Long

' ProcedureReplace
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef lpvDest As Any, ByRef lpvSrc As Any, ByVal cbLen As Long)
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long

' 100% compatible Unicode version of Command$
Public Function Command() As String
    Dim intArguments As Integer, lngPtr As Long, lngStrEnd As Long, lngStrStart As Long
    ' get command line
    lngPtr = CommandLineToArgvW(GetCommandLineW, intArguments)
    ' we need more than one argument: the first one is the exe name
    If (intArguments > 0) And (lngPtr <> 0) Then
        ' get the starting position of the first string
        GetMem4 lngPtr + 4, lngStrStart
        ' get the starting position of the last string
        GetMem4 lngPtr + (intArguments - 1) * 4, lngStrEnd
        ' allocate memory; also for compability we replace null characters with a space
        Command = Replace(StringFromPtr(SysAllocStringLen(lngStrStart, (lngStrEnd - lngStrStart) \ 2 + lstrlenW(lngStrEnd))), vbNullChar, " ")
    End If
    ' free memory if necessary
    If lngPtr Then LocalFree lngPtr
End Function
Public Function IIf(ByVal Expression As Boolean, TruePart As Variant, FalsePart As Variant) As Variant
    ' as silly as it sounds, this IIf is actually faster than the native implementation
    If Expression Then IIf = TruePart Else IIf = FalsePart
End Function
' note: I didn't bother to go ahead and start hacking with HelpFile and Context
Public Function MsgBox(ByVal Prompt As String, Optional Buttons As VbMsgBoxStyle = vbOKOnly, Optional ByVal Title As String, Optional ResourceIcon As String, Optional hWndOwner As Long) As VbMsgBoxResult
    Dim udtMsgBox As MsgBoxParams
    ' if no owner is specified, try to use the active form
    If hWndOwner = 0 Then If Not Screen.ActiveForm Is Nothing Then hWndOwner = Screen.ActiveForm.hWnd
    With udtMsgBox
        .cbSize = Len(udtMsgBox)
        ' important to set owner to get behavior similar to the native MsgBox
        .hWndOwner = hWndOwner
        .hInstance = App.hInstance
        ' set the message
        .lpszText = StrPtr(Prompt)
        ' if no title is given, use the application title like the native MsgBox
        If LenB(Title) = 0 Then Title = App.Title
        .lpszCaption = StrPtr(Title)
        ' thought this would be a nice feature addition
        If LenB(ResourceIcon) = 0& Then
            .dwStyle = Buttons
        Else
            .dwStyle = (Buttons Or MB_USERICON) And Not (&H70&)
            .lpszIcon = StrPtr(ResourceIcon)
        End If
    End With
    ' show the message box
    MsgBox = MessageBoxIndirectW(udtMsgBox)
End Function
Private Sub ProcedureReplace(ByVal AddressOfDest As Long, ByVal AddressOfSrc As Long)
    Dim lngJMPASM(1) As Long, lngBytesWritten As Long, lngProcessHandle As Long
    ' get a handle for current process
    lngProcessHandle = OpenProcess(&H1F0FFF, 0&, GetCurrentProcessId)
    ' if failed, we can't do anything
    If lngProcessHandle = 0 Then Exit Sub
    ' check if we are in the IDE
    If App.LogMode = 0 Then
        ' get the real locations of the procedures
        CopyMemory AddressOfDest, ByVal AddressOfDest + &H16&, 4&
        CopyMemory AddressOfSrc, ByVal AddressOfSrc + &H16&, 4&
    End If
    ' set ASM JMP
    lngJMPASM(0) = &HE9000000
    ' set JMP parameter (how many bytes to jump)
    lngJMPASM(1) = AddressOfSrc - AddressOfDest - 5
    ' replace original procedure with the JMP
    WriteProcessMemory lngProcessHandle, ByVal AddressOfDest, ByVal VarPtr(lngJMPASM(0)) + 3, 5, lngBytesWritten
    ' close handle for current process
    CloseHandle lngProcessHandle
End Sub
Private Function StringFromPtr(ByVal AllocatedPtr As Long) As String
    ProcedureReplace AddressOf UniInteraction.StringFromPtr, AddressOf SwapParamOut
    StringFromPtr = StringFromPtr(AllocatedPtr)
End Function
Private Function SwapParamOut(ByVal Value1 As Long) As Long
    SwapParamOut = Value1
End Function
