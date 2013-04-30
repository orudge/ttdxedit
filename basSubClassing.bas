Attribute VB_Name = "basSubclassing"
Option Explicit

Public Enum EnumSubclassID
    escidMain = 1
    escidCity = 2
    escidFileTypes = 3
    escidSelectGame = 4
    escidStation = 5
    escidVehicle = 6
End Enum

Private Declare Function RemoveWindowSubclass Lib "comctl32.dll" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long) As Long
Private Declare Function SetWindowSubclass Lib "comctl32.dll" (ByVal hWnd As Long, ByVal pfnSubclass As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByVal pDest As Long, ByVal sz As Long)

Public Declare Function DefSubclassProc Lib "comctl32.dll" (ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Declare Function SendMessageAsLong Lib "user32.dll" Alias "SendMessageW" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const NF_REQUERY = 4
Public Const WM_NOTIFYFORMAT = &H55

Public Const WM_USER = &H400
Public Const OCM__BASE = WM_USER + &H1C00

Public Function SubclassProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal uIdSubclass As Long, ByVal dwRefData As Long) As Long
    Dim bCallDefProc As Boolean
    Dim oClient As ISubclassedWindow
    Dim lRet As Long
    
    On Error GoTo StdHandler_Error
    
    bCallDefProc = True
    
    If dwRefData Then
        Set oClient = GetObjectFromPointer(dwRefData)
        
        If Not (oClient Is Nothing) Then
            lRet = oClient.HandleMessage(hWnd, uMsg, wParam, lParam, uIdSubclass, bCallDefProc)
        End If
    End If
    
StdHandler_Ende:
    On Error Resume Next
    
    If bCallDefProc Then
      lRet = DefSubclassProc(hWnd, uMsg, wParam, lParam)
    End If
    
    SubclassProc = lRet
    Exit Function
    
StdHandler_Error:
    Debug.Print "Error in SubclassProc: ", Err.Number, Err.Description
    Resume StdHandler_Ende
End Function

Public Function SubclassWindow(ByVal hWnd As Long, oClient As ISubclassedWindow, ByVal eSubclassID As EnumSubclassID) As Boolean
    Dim bRet As Boolean
    
    On Error GoTo StdHandler_Error
    
#If Subclassing = 1 Then
    If App.LogMode <> 0 Then  ' not running in IDE
        If SetWindowSubclass(hWnd, AddressOf basSubclassing.SubclassProc, eSubclassID, ObjPtr(oClient)) Then
            bRet = True
        End If
    End If
#End If

StdHandler_Ende:
    SubclassWindow = bRet
    Exit Function
    
StdHandler_Error:
    Debug.Print "Error in SubclassWindow: ", Err.Number, Err.Description
    bRet = False
    Resume StdHandler_Ende
End Function

Public Function UnSubclassWindow(ByVal hWnd As Long, ByVal eSubclassID As EnumSubclassID) As Boolean
    Dim bRet As Boolean
    
    On Error GoTo StdHandler_Error
    
#If Subclassing = 1 Then
    If App.LogMode <> 0 Then  ' not running in IDE
        If RemoveWindowSubclass(hWnd, AddressOf basSubclassing.SubclassProc, eSubclassID) Then
            bRet = True
        End If
    End If
#End If

StdHandler_Ende:
    UnSubclassWindow = bRet
    Exit Function
    
StdHandler_Error:
    Debug.Print "Error in UnSubclassWindow: ", Err.Number, Err.Description
    bRet = False
    Resume StdHandler_Ende
End Function

' returns the object <lPtr> points to
Private Function GetObjectFromPointer(ByVal lPtr As Long) As Object
    Dim oRet As Object
    
    ' get the object <lPtr> points to
    CopyMemory VarPtr(oRet), VarPtr(lPtr), LenB(lPtr)
    Set GetObjectFromPointer = oRet
    ' free memory
    ZeroMemory VarPtr(oRet), 4
End Function
