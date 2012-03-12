VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CGUID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "GUID Generator Class"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Transport Tycoon Saved Game Manager                       '
' Version 2.2                                               '
'                                                           '
' Copyright © Owen Rudge 2000-2002. All Rights Reserved.    '
' Web site: www.transporttycoon.uk.tt  tt@orudge.freeuk.com '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Declare Function CoCreateGuid Lib "ole32.dll" (GUID As LongGUID) As Long
Attribute CoCreateGuid.VB_Description = "Creates a new GUID"

Private Type LongGUID
    Value(15) As Byte
End Type

Private GUID As LongGUID
Private GS As String
Private GR As String

Private Result As Long
Public Sub CreateNew()
Attribute CreateNew.VB_Description = "Creates a new GUID"
    On Error Resume Next  ' OK, we'll end up with a deformed GUID this way...
    
    Dim i As Integer
    Result = CoCreateGuid(GUID)
    
    GS = ""
    Dim Temp As String
    Temp = ""
    For i = 0 To 3
       Temp = Right("00" + Hex(GUID.Value(i)), 2) + Temp
    Next i
    
    GS = Temp
    Temp = ""
    For i = 4 To 5
        Temp = Right("00" + Hex(GUID.Value(i)), 2) + Temp
    Next i
    GS = GS + Temp
    
    Temp = ""
    For i = 6 To 7
        Temp = Right("00" + Hex(GUID.Value(i)), 2) + Temp
    Next i
    GS = GS + Temp

    Temp = ""
    For i = 8 To 15
        Temp = Right("00" + Hex(GUID.Value(i)), 2) + Temp
    Next i
    GS = GS + Temp
    
    GR = "{"
    GR = GR + MID(GS, 1, 8) + "-"
    GR = GR + MID(GS, 9, 4) + "-"
    GR = GR + MID(GS, 13, 4) + "-"
    GR = GR + MID(GS, 17, 4) + "-"
    GR = GR + MID(GS, 21) + "}"
End Sub

Public Property Get RegVal() As String
Attribute RegVal.VB_Description = "Returns the new GUID"
    On Error Resume Next
    RegVal = GR
End Property


