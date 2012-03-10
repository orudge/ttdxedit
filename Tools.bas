Attribute VB_Name = "Tools"
Option Explicit
Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647
Private Const OFFSET_2 = 65536
Private Const MAXINT_2 = 32767


Public Function CheckNumInput(wText As String, wNew As Integer) As Integer
    'Debug.Print wNew
    CheckNumInput = wNew

    If wNew = 8 Then
    ElseIf wNew = Asc("-") Then
    ElseIf wNew < 48 Then
        CheckNumInput = 0
    ElseIf wNew > 57 Then
        CheckNumInput = 0
    End If
End Function

Public Function jROL_l(wVal As Long, wS) As Long
    Dim Wa As Integer, Wb As Integer
    jROL_l = wVal
    If wS < 1 Then Exit Function
    For Wa = 1 To wS
        Wb = CInt((wVal And &H80000000) / 2 ^ 31)
        wVal = UnsignedToLong(LongToUnsigned(wVal And &H7FFFFFFF) * 2)
        If Wb <> 0 Then wVal = wVal Or 1
    Next Wa
    jROL_l = wVal
End Function

Public Function j2Bin(wVal As Long) As String
    Dim Wa As Integer
    If wVal < 0 Then j2Bin = "1" Else j2Bin = "0"
    For Wa = 30 To 0 Step -1
        If (wVal And 2 ^ Wa) <> 0 Then j2Bin = j2Bin + "1" Else j2Bin = j2Bin + "0"
    Next Wa
End Function


Function UnsignedToLong(Value As Double) As Long
  If Value < 0 Or Value >= OFFSET_4 Then Error 6 ' Overflow
  If Value <= MAXINT_4 Then
    UnsignedToLong = Value
  Else
    UnsignedToLong = Value - OFFSET_4
  End If
End Function

Function LongToUnsigned(Value As Long) As Double
  If Value < 0 Then
    LongToUnsigned = Value + OFFSET_4
  Else
    LongToUnsigned = Value
  End If
End Function

Function UnsignedToInteger(Value As Long) As Integer
  If Value < 0 Or Value >= OFFSET_2 Then Error 6 ' Overflow
  If Value <= MAXINT_2 Then
    UnsignedToInteger = Value
  Else
    UnsignedToInteger = Value - OFFSET_2
  End If
End Function

Function IntegerToUnsigned(Value As Integer) As Long
  If Value < 0 Then
    IntegerToUnsigned = Value + OFFSET_2
  Else
    IntegerToUnsigned = Value
  End If
End Function
