Attribute VB_Name = "TTDXMacros"
Option Explicit

Public Sub MacroIndustry(wNo As Integer, wDta As Integer)
    '
    ' Industry Macros
    ' wNo is the number of the macro to run
    ' wDta is data given to the macro
    '
    ' wNo=1: Max production, wDta = Cargo code or -1
    ' wNo=2: Min production, wDta = Cargo code or -1
    '
    Dim wXa As TTDXIndInfo, Wa As Integer, Wb As Integer, Wc As Integer
    
    If wNo = 1 Then Wc = 240
    If wNo = 2 Then Wc = 1
    
    For Wa = 0 To 89
        wXa = TTDXIndustryInfo(Wa)
        If wXa.H + wXa.W > 0 Then
            Select Case wNo
            Case 1, 2
                For Wb = 0 To 1
                    If wXa.Prod(Wb) = wDta Then
                        wXa.ProR(Wb) = Wc
                    ElseIf wXa.Prod(Wb) < 20 And wDta = -1 Then
                        If wXa.ProR(Wb) > 0 Then: wXa.ProR(Wb) = Wc
                    End If
                Next Wb
            End Select
            TTDXIndustryPut wXa
        End If
    Next Wa
End Sub

Public Sub TTDXtermacRemWood()
    Dim Wva As TTDXlandscape, Wa As Integer, Wb As Integer
    
    If CurFile > " " Then
        For Wa = 0 To 254
            For Wb = 0 To 254
                Wva = TTDXgetLandscape(Wa, Wb)
                If Wva.Object = 4 Then Wva.Object = 0: Wva.L5 = 0: TTDXputLandscape Wva
            Next Wb
            DoEvents
        Next Wa
    
        wData(&H44BBD) = wData(&H44BBD) Or 32
    End If
End Sub
Public Sub TTDXtermacOwnAIRoad()
    Dim Wva As TTDXlandscape, Wa As Integer, Wb As Integer, Wc As Long
    
    If CurFile > " " Then
        Wc = 0
        For Wa = 0 To 254
            For Wb = 0 To 254
                Wva = TTDXgetLandscape(Wa, Wb)
                If Wva.Object = 2 Then
                    If (Wva.L5 And 24) = 16 Then
                        If jBetween(0, (Wva.L3 And 255&), &H10) Then
                            Wva.L3 = Wva.L3 And &HFF00&: TTDXputLandscape Wva: Wc = Wc + 1
                        End If
                    ElseIf jBetween(0, Wva.Owner, &H10) And ((Wva.L5 And 32) = 0) Then
                        Wva.Owner = 0: TTDXputLandscape Wva: Wc = Wc + 1
                    End If
                ElseIf Wva.Object = 9 Then
                    If (Wva.L5 And 192) = 192 Then
                        '
                        ' Middle piece of bridge, check for road under
                        '
                        If (Wva.L5 And 40) = 40 Then
                            Debug.Print "."
                            If jBetween(0, Wva.Owner, &H10) Then
                                Wva.Owner = 0: TTDXputLandscape Wva: Wc = Wc + 1
                            End If
                        End If
                    ElseIf (Wva.L5 And 130) = 130 Then
                        '
                        ' The ending of a road bridge
                        '
                        If jBetween(0, Wva.Owner, &H10) Then
                            Wva.Owner = 0: TTDXputLandscape Wva: Wc = Wc + 1
                        End If
                    ElseIf (Wva.L5 And 4) Then
                        '
                        ' Road Tunnel
                        '
                        If jBetween(0, Wva.Owner, &H10) Then
                            Wva.Owner = 0: TTDXputLandscape Wva: Wc = Wc + 1
                        End If
                    End If
                End If
            Next Wb
            DoEvents
        Next Wa
        Wa = MsgBox(Format(Wc) + " AI road tiles now owned by Player 1.")
        wData(&H44BBD) = wData(&H44BBD) Or 32
    End If
End Sub

Public Sub TTDXtermacOwnCityBridge()
    Dim Wva As TTDXlandscape, Wa As Integer, Wb As Integer, Wc As Long
    
    If CurFile > " " Then
        Wc = 0
        For Wa = 0 To 254
            For Wb = 0 To 254
                Wva = TTDXgetLandscape(Wa, Wb)
                If Wva.Object = 9 Then
                    If (Wva.L5 And 194) = 130 Then
                        '
                        ' The endings of a road bridge
                        '
                        If &H10 < Wva.Owner Then
                            Wva.Owner = 0: TTDXputLandscape Wva: Wc = Wc + 1
                        End If
                    End If
                End If
            Next Wb
            DoEvents
        Next Wa
        
        wData(&H44BBD) = wData(&H44BBD) Or 32
        Wa = MsgBox(Format(Wc / 2) + " city bridges now owned by Player 1.")
    End If
End Sub


