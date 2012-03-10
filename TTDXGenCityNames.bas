Attribute VB_Name = "TTDXGenCityNames"
Option Explicit
'
' ******* ******* ******  *    * ******* ******  * *******
'    *       *     *    *  *  *  *        *    * *    *
'    *       *     *    *   **   *****    *    * *    *
'    *       *     *    *  *  *  *        *    * *    *
'    *       *    ******  *    * ******* ******  *    *
'*****************************************************************************************************
'**** Functions for generating city and station names                                             ****
'*****************************************************************************************************
'**** Written By: Jens Vang Petersen (C)2002                                                      ****
'**** Thanks to:  Josef Drexler for the basic information.                                        ****
'*****************************************************************************************************

Public Function TTDXMakeC1CityName(wB1 As Byte, wB2 As Byte, wB3 As Byte, wB4 As Byte) As String
    Dim Wa As Long, Wsa As String, Wsb As String, Wsc As String, Wva As Variant
    '
    ' C1 covers American/English and German names..
    ' (Call this if CityOffset+5=&h20 and CityOffset+4=&hC1)
    '  wB1 is found in CityOffset+6
    '  wB2 is found in CityOffset+7
    '  wB3 is found in CityOffset+8
    '  wB4 is found in CityOffset+9
    '
    ' The buildingblocks used for English/American Names:
    '
    Const EngP1 As String = "Great_.Little_.New_.Fort_"
    Const EngP2 As String = "Wr.B.C.Ch.Br.D.Dr.F.Fr.Fl.G.Gr.H.L.M.N.P.Pr.Pl.R.S.S.Sl.T.Tr.W"
    Const EngP3 As String = "ar.a.e.in.on.u.un.en"
    Const EngP4 As String = "n.ning.ding.d..t.fing"
    Const EngP5 As String = "ville.ham.field.ston.town.bridge.bury.wood.ford.hall..way.stone.borough.ley.head.bourne.pool.worth.hill.well.hattan.burg"
    Const EngP6 As String = "-on-sea._Bay._Market._Cross._Bridge._Falls._City._Ridge._Springs"
    '
    ' Buildingblocks for german names:
    '
    Const GerP1 As String = "Bruns.Lim.Han.Brun.Ober.Frank.Rott.Frei.Ravens.Schwein.Osna.Düssel.Wester.Flens.Mühl.Heidel.Inns.Cloppen.Pfung.Michel.Wert.Wildes.Freuden"
    Const GerP2 As String = "gen.stadt.heim.burg.haven.ford.feld.mund.münster.hausen.furt.brück.over.wald"
    Wsa = ""
    If (wB4 And &H80) = 0 Then
        '
        ' English/American Name
        '
        Wa = Fix((((wB1 + wB2 * 256& + wB3 * 2 ^ 16) And &HFFFF0) / 8 * 13) / 2 ^ 16)
        Wva = Split(EngP2, ".")
        If (Wa > -1) And (Wa <= UBound(Wva)) Then Wsa = Wsa + Wva(Wa)
        '
        Wa = Fix((((wB1 + wB2 * 256& + wB3 * 2 ^ 16) And &H7FFF80) / 16) / 2 ^ 16)
        Wva = Split(EngP3, ".")
        If (Wa > -1) And (Wa <= UBound(Wva)) Then Wsa = Wsa + Wva(Wa)
        '
        Wa = Fix(((((wB2 + wB3 * 256& + wB4 * 2 ^ 16) And &H3FFFC) / 4) * 7) / 2 ^ 16)
        Wva = Split(EngP4, ".")
        If (Wa > -1) And (Wa <= UBound(Wva)) Then Wsa = Wsa + Wva(Wa)
        '
        Wa = Fix(((((wB2 + wB3 * 256& + wB4 * 2 ^ 16) And &H1FFFE0) / 32) * 23) / 2 ^ 16)
        Wva = Split(EngP5, ".")
        If (Wa > -1) And (Wa <= UBound(Wva)) Then Wsa = Wsa + Wva(Wa)
        '
        ' Check for some illegal names:
        '
        If Left(Wsa, 2) = "Ce" Then Wsa = "Ke" + Mid(Wsa, 3)
        If Left(Wsa, 2) = "Ci" Then Wsa = "Ki" + Mid(Wsa, 3)
        Wsb = Left(Wsa, 4): Wsc = Mid(Wsa, 5)
        Select Case Wsb
            Case "Cunt": Wsa = "East" + Wsc
            Case "Slag": Wsa = "Pits" + Wsc
            Case "Slut": Wsa = "Edin" + Wsc
            Case "Fart": Wsa = "Boot" + Wsc
            Case "Drar": Wsa = "Quar" + Wsc
            Case "Dreh": Wsa = "Bash" + Wsc
            Case "Frar": Wsa = "Shor" + Wsc
            Case "Grar": Wsa = "Aber" + Wsc
            Case "Brar": Wsa = "Over" + Wsc
            Case "Wrar": Wsa = "Inve" + Wsc
        End Select
        '
        ' Suffixname (If any)
        '
        Wa = Fix(((((wB2 + wB3 * 256& + wB4 * 2 ^ 16) And &H7FFF80) / 128) * 69) / 2 ^ 16) - 60
        Wva = Split(EngP6, "."): If (Wa > -1) And (Wa <= UBound(Wva)) Then Wsa = Wsa + Wva(Wa)
        '
        ' Prefixname (If any)
        '
        Wa = Fix(((wB1 + wB2 * 256&) * 27) / 2 ^ 15) - 50
        Wva = Split(EngP1, "."): If (Wa > -1) And (Wa <= UBound(Wva)) Then Wsa = Wva(Wa) + Wsa
    Else
        '
        ' German Name
        '
        Wa = Fix((((wB1 + wB2 * 256&) And &HFFFF) * 23) / 2 ^ 16)
        Wva = Split(GerP1, "."): If (Wa > -1) And (Wa <= UBound(Wva)) Then Wsa = Wsa + Wva(Wa)
        '
        Wa = Fix((((wB1 + wB2 * 256& + wB3 * 2 ^ 16) And &H1FFFE0) / 32 * 14) / 2 ^ 16)
        Wva = Split(GerP2, "."): If (Wa > -1) And (Wa <= UBound(Wva)) Then Wsa = Wsa + Wva(Wa)
    End If
    '
    ' The last touch for the perfect name :-)
    '
    TTDXMakeC1CityName = Replace(Wsa, "_", " ")
End Function

Public Function TTDXMakeC2CityName(wB1 As Byte, wB2 As Byte, wB3 As Byte, wB4 As Byte) As String
    Dim Wa As Long, Wsa As String, Wsb As String, Wsc As String, Wva As Variant
    '
    ' C2 covers Spanish/French and Silly names..
    ' (Call this if CityOffset+5=&h20 and CityOffset+4=&hC2)
    ' wB1 is found in CityOffset+6
    ' wB2 is found in CityOffset+7
    ' wB3 is found in CityOffset+8
    ' wB4 is found in CityOffset+9
    '
    ' The Spanish/French Names:
    '
    Const SpFr As String = "Caracas,Maracay,Maracaibo,Velencia,El Dorado,Morrocoy,Cata,Cataito," + _
                           "Ciudad Bolivar,Barquisimeto,Merida,Puerto Ordaz,Santa Elena,San Juan," + _
                           "San Luis,San Rafael,Santiago,Barcelona,Barinas,San Cristobal," + _
                           "San Fransisco,San Martin,Guayana,San Carlos,El Limon,Coro,Corocoro," + _
                           "Puerto Ayacucho,Elorza,Arismendi,Trujillo,Carupano,Anaco,Lima,Cuzco," + _
                           "Iquitos,Callao,Huacho,Camana,Puerto Chala,Santa Cruz,Quito,Cuenca,Huacho," + _
                           "Tulcan,Esmereldas,Ibarra,San Lorenzo,Macas,Morana,Machala,Zamora," + _
                           "Latacunga,Tena,Cochabamba,Ascencion,Magdalena,Santa Ana,Manoa,Sucre," + _
                           "Oruro,Uyuni,Potosi,Tupiza,La Quiaca,Yacuiba,San Borja,Fuerte Olimpio," + _
                           "Fortin Esteros,Campo Grande,Bogota,El Banco,Zaragosa,Neiva,Mariano," + _
                           "Cali,La Palma,Andoas,Barranca,Montevideo,Valdivia,Arica,Temuco,Tocopilla," + _
                           "Mendoza,Santa Rosa,Agincourt,Lille,Dinan,Aubusson,Rodez,Bergerac," + _
                           "Bordeaux,Bayonne,Montpellier,Montelimar,Valence,Digne,Nice,Cannes," + _
                           "St. Tropez,Marseilles,Narbonne,Sète,Aurillac,Gueret,Le Creusot,Nevers," + _
                           "Auxerre,Versailles,Meaux,Châlons,Compiègne,Metz,Chaumont,Langres,Bourg," + _
                           "Lyons,Vienne,Grenoble,Toulon,Rennes,Le Mans,Angers,Nantes,Châteauroux," + _
                           "Orléans,Lisieux,Cherbourg,Morlaix,Cognac,Agen,Tulle,Blois,Troyes," + _
                           "Charolles,Grenoble,Chambéry,Tours,St. Brieuc,St. Malo,La Rochelle," + _
                           "St. Flour,Le Puy,Vichy,St. Valery,Beaujolais,Narbonne,Albi,St. Valery," + _
                           "Biarritz,Béziers,Nîmes,Chamonix,Angoulême,Alencon"
    '
    ' Silly buildingblocks
    '
    Const SillyP1 As String = "Binky.Blubber.Bumble.Crinkle.Crusty.Dangle.Dribble.Flippety.Goggle." + _
                              "Muffin.Nosey.Pinker.Quack.Rumble.Sleepy.Sliggles.Snooze.Teddy.Tinkle." + _
                              "Twister.Pinker.Hippo.Itchy.Jelly.Jingle.Jolly.Kipper.Lazy.Frogs.Mouse." + _
                              "Quack.Cheeky.Lumpy.Grumpy.Mangle.Fiddle.Slugs.Noodles.Poodles.Shiver." + _
                              "Rumble.Pixie.Puddle.Riddle.Rattle.Rickety.Waffle.Sagging.Sausage.Egg." + _
                              "Sleepy.Scatter.Scramble.Silly.Simple.Trickle.Slippery.Slimey.Slumber." + _
                              "Soggy.Sliggles.Splutter.Sulky.Swindle.Swivel.Tasty.Tangle.Toggle." + _
                              "Trotting.Tumble.Snooze.Water.Windy.Amble.Bubble.Cheery.Cheese.Cockle." + _
                              "Cracker.Crumple.Teddy.Evil.Fairy.Falling.Fishy.Fizzle.Frosty.Griddle"
    Const SillyP2 As String = "ton.bury.bottom.ville.well.weed.worth.wig.wick.wood.pool.head.burg.gate.bridge"
    Wsa = ""
    If (wB4 And &H80) = 0 Then
        '
        ' Spanish/French names
        '
        Wa = wB1 + wB2 * 256
        Wva = Split(SpFr, ","): If (Wa > -1) And (Wa <= UBound(Wva)) Then Wsa = Wva(Wa)
    Else
        '
        ' Silly Names
        '
        Wa = wB1 + wB2 * 256&
        Wva = Split(SillyP1, "."): If (Wa > -1) And (Wa <= UBound(Wva)) Then Wsa = Wva(Wa)
        '
        Wa = (wB3 + wB4 * 256&) And &H7FFF
        Wva = Split(SillyP2, "."): If (Wa > -1) And (Wa <= UBound(Wva)) Then Wsa = Wsa + Wva(Wa)
        '
    End If
    TTDXMakeC2CityName = Replace(Wsa, "_", " ")
End Function

Public Function TTDXStationExtension(wCode As Byte) As String
    Dim Wsa As String
    '
    ' Get the extensionname of a station,
    '  wCode is found in "StationOffset+&h14"
    '
    Select Case wCode
        Case &H10: Wsa = " North"
        Case &H11: Wsa = " South"
        Case &H12: Wsa = " East"
        Case &H13: Wsa = " West"
        Case &H14: Wsa = " Central"
        Case &H15: Wsa = " Transfer"
        Case &H16: Wsa = " Halt"
        Case &H17: Wsa = " Valley"
        Case &H18: Wsa = " Heights"
        Case &H19: Wsa = " Woods"
        Case &H1A: Wsa = " Lakeside"
        Case &H1B: Wsa = " Exchange"
        Case &H1C: Wsa = " Airport"
        Case &H1D: Wsa = " Oilfield"
        Case &H1E: Wsa = " Mines"
        Case &H1F: Wsa = " Docks"
        Case &H20 To &H28: Wsa = " Buoy " + Format(wCode - &H1F)
        Case &H29: Wsa = " Annexe"
        Case &H2A: Wsa = " Sidings"
        Case &H2B: Wsa = " Branch"
        Case &H2C: Wsa = " Upper"
        Case &H2D: Wsa = " Lower"
        Case &H2E: Wsa = " Heliport"
        Case &H2F: Wsa = " Forest"
    End Select
    TTDXStationExtension = Wsa
End Function
