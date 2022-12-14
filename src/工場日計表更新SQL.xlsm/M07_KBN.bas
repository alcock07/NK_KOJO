Attribute VB_Name = "M07_KBN"
Option Explicit

Public KBN_NAME As String

Function KBN_CHG(strTCD As String, strGCD As String) As String
    '============================================================================================================
    'ćŞĎX
    '============================================================================================================

    Select Case strGCD
        Case "0000000710004"
            If strTCD = "0000000710001" Then
                KBN_CHG = "G01" '¸ŢŮ°Ěßč
            ElseIf strTCD = "0000000710002" Then
                KBN_CHG = "G02" '¸ŢŮ°Ěßčĺă
            ElseIf strTCD = "0000000710003" Then
                KBN_CHG = "G03" '¸ŢŮ°ĚßčC
            ElseIf strTCD = "0000000710004" Then
                KBN_CHG = "G04" '¸ŢŮ°Ěßč{
            ElseIf strTCD = "0000000710005" Then
                KBN_CHG = "G05" '¸ŢŮ°ĚßčěÖ
            ElseIf strTCD = "0000000710007" Then
                KBN_CHG = "G07" '¸ŢŮ°ĚßčŞ
            ElseIf strTCD = "0000000710009" Then
                KBN_CHG = "G09" '¸ŢŮ°ĚßčźĂŽ
             ElseIf strTCD = "0000000710011" Then
                KBN_CHG = "G11" '¸ŢŮ°ĚßčkÖ
            Else
                KBN_CHG = "G99" '¸ŢŮ°ĚßčťĚź
            End If
        Case "0000000710505"
            KBN_CHG = "B01" 'PCACX^[sŽY
            KBN_NAME = "PCACX^[sŽY"
        Case "0000000212012"
            KBN_CHG = "B02" 'ęÝ
            KBN_NAME = "ęÝ"
        Case "0000000717000"
            KBN_CHG = "B03" '¸ŢfBnEX
            KBN_NAME = "¸Ţ×ÝĂŢ¨Ęł˝"
        Case "0000000711001"
            KBN_CHG = "B04" '§Ż¤Ď
            KBN_NAME = "§Ż¤Ď"
        Case "0000000110775"
            KBN_CHG = "B05" 'hZî
            KBN_NAME = "hZî"
        Case "0000000711014"
            KBN_CHG = "B06" 'AQ
            KBN_NAME = "AQ"
        Case "0000000210215"
            KBN_CHG = "B07" 'AC_Ýv
            KBN_NAME = "AC_Ýv"
        Case "0000000710014"
            KBN_CHG = "B08" 'Cu[nEX
            KBN_NAME = "Cu[nEX"
        Case Else
            KBN_CHG = "B99"
            KBN_NAME = "ťĚź"
    End Select

End Function

Function KBN_CHGT(strTCD As String, strGCD As String, strKBN As String, strKBP As String) As String
    '============================================================================================================
    'ćŞĎX
    '============================================================================================================
    'Gş°ÄŢĹžÓćđťč
    Select Case strGCD
        Case "0000000115024"
            KBN_CHGT = "A01" 'ŔĎÎ°Ń
            KBN_NAME = "^}z["
        Case "0000000212012"
            KBN_CHGT = "A02" 'ęÝ
            KBN_NAME = "ęÝ"
        Case "0000000811701"
            KBN_CHGT = "A03" 'AVANTIA
            KBN_NAME = "AVANTIA"
        Case "0000000210215"
            KBN_CHGT = "A04" 'AC_Ýv
            KBN_NAME = "AC_Ýv"
        Case "0000000210700"
            KBN_CHGT = "A05" 'ŃcYĆ
            KBN_NAME = "ŃcYĆ"
        Case "0000000812301"
            KBN_CHGT = "A06" 'ą´×Î°Ń
            KBN_NAME = "ą´×Î°Ń"
        Case "0000000811101"
            KBN_CHGT = "A07" 'ąŻÄĘłźŢÝ¸Ţ(¸ŢŻÄŢŘËŢÝ¸Ţ)
            KBN_NAME = "ąŻÄĘłźŢÝ¸Ţ"
        Case "0000000812105"
            KBN_CHGT = "A09" 'Ąz
            KBN_NAME = "ťĚź"
        Case "0000000812204"
            KBN_CHGT = "A09" 'Ŕ˛şłĘł˝
            KBN_NAME = "ťĚź"
        Case "0000000811401"
            KBN_CHGT = "A09" '¸×ź˝Î°Ń
            KBN_NAME = "ťĚź"
        Case "0000000212305"
            KBN_CHGT = "A08" 'ą˛ĂŢ¨Î°Ń
            KBN_NAME = "ą˛ĂŢ¨Î°Ń"
        Case "0000000115320"
            KBN_CHGT = "A09" 'ĘP
            KBN_NAME = "ťĚź"
        Case "0000000111450"
            KBN_CHGT = "A09" 'ť°×
            KBN_NAME = "ťĚź"
        Case "0000000216450"
            KBN_CHGT = "B04" 'Î¸´Â
            KBN_NAME = "zNGc"
        Case "0000000810001"
            KBN_CHGT = "B05" 'lźÜ°¸˝
            KBN_NAME = "lźÜ°¸˝"
        Case "0000000810014"
            KBN_CHGT = "B09" 'ăť|
            KBN_NAME = "ťĚź"
        Case "0000000819011"
            KBN_CHGT = "A08" 'šŕŽ(ĺăč )
            KBN_NAME = "šŕŽ(ĺăč )"
        Case "0000000812803"
            KBN_CHGT = "A08" 'šŕŽ(ĺăč )
            KBN_NAME = "šŕŽ(ĺăč )"
        Case "0000000819001" 'šŕŽ
            If strTCD = "0000000812301" Or strTCD = "0000000812303" Or strTCD = "0000000812308" Or strTCD = "0000000812350" Then  'ą´×Î°Ń(oR)
                KBN_CHGT = "A06"
                KBN_NAME = "ą´×Î°Ń"
            ElseIf strTCD = "0000000819004" Then
                KBN_CHGT = "B04"
                KBN_NAME = "{(Î¸´Âź)"
            ElseIf strTCD = "0000000819005" Then
                KBN_CHGT = "C01"
                KBN_NAME = "šŕŽiźĂŽj"
            Else
                '¤ićŞ(HINKB)Ĺťč
                If Trim(strKBN) = "07" Or Trim(strKBN) = "16" Then      'ĚŢŘŻźŢ
                    KBN_CHGT = "B01"
                    KBN_NAME = "šŕŽĚŢŘŻźŢ"
                ElseIf Trim(strKBN) = "08" Then 'ĐĆŰ°ÄŢ
                    KBN_CHGT = "B02"
                    KBN_NAME = "šŕŽĐĆŰ°ÄŢ"
                ElseIf Trim(strKBN) = "09" Then 'TL
                    KBN_CHGT = "B03"
                    KBN_NAME = "šŕŽTL"
                Else
                    KBN_CHGT = "C09"
                    KBN_NAME = "ťĚź"
                End If
            End If
        Case "0000000819010"           'ÖAL
            KBN_CHGT = "C02"
            KBN_NAME = "ÖAL"
        Case "0000000830035"           'ÁHnľ
            KBN_CHGT = "W99"
            KBN_NAME = "ÁHnľ"
        Case "0000000830099"           'ĚŢŘŻźŢŢżřľ
            KBN_CHGT = "Z99"
            KBN_NAME = "ĚŢŘŻźŢŢżřľ"
        Case Else
            If strTCD = "0000000810999" Then 'Gă
                KBN_CHGT = "C09"
                KBN_NAME = "ťĚź"
            ElseIf Trim(strKBP) = "01" Then
                KBN_CHGT = "A09" 'č ťĚź
                KBN_NAME = "ťĚź"
            Else
                Select Case Trim(strKBN)
                    Case "11", "14", "16", "17"
                        KBN_CHGT = "A09" 'č ťĚź
                        KBN_NAME = "ťĚź"
                    Case "07", "08", "09", "10"
                        KBN_CHGT = "B09" 'ubWťĚź
                        KBN_NAME = "ťĚź"
                    Case Else
                        KBN_CHGT = "C09" 'ťĚź
                        KBN_NAME = "ťĚź"
                End Select
            End If
        End Select
End Function
