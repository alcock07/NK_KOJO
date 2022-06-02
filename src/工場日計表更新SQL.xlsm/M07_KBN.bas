Attribute VB_Name = "M07_KBN"
Option Explicit

Public KBN_NAME As String

Function KBN_CHG(strTCD As String, strGCD As String) As String
    '============================================================================================================
    '区分変更処理
    '============================================================================================================

    Select Case strGCD
        Case "0000000710004"
            If strTCD = "0000000710001" Then
                KBN_CHG = "G01" 'ｸﾞﾙｰﾌﾟ売り東京
            ElseIf strTCD = "0000000710002" Then
                KBN_CHG = "G02" 'ｸﾞﾙｰﾌﾟ売り大阪
            ElseIf strTCD = "0000000710003" Then
                KBN_CHG = "G03" 'ｸﾞﾙｰﾌﾟ売り東海
            ElseIf strTCD = "0000000710004" Then
                KBN_CHG = "G04" 'ｸﾞﾙｰﾌﾟ売り本部
            ElseIf strTCD = "0000000710005" Then
                KBN_CHG = "G05" 'ｸﾞﾙｰﾌﾟ売り南関東
            ElseIf strTCD = "0000000710007" Then
                KBN_CHG = "G07" 'ｸﾞﾙｰﾌﾟ売り福岡
            ElseIf strTCD = "0000000710009" Then
                KBN_CHG = "G09" 'ｸﾞﾙｰﾌﾟ売り名古屋
             ElseIf strTCD = "0000000710011" Then
                KBN_CHG = "G11" 'ｸﾞﾙｰﾌﾟ売り北関東
            Else
                KBN_CHG = "G99" 'ｸﾞﾙｰﾌﾟ売りその他
            End If
        Case "0000000710505"
            KBN_CHG = "B01" 'ケイアイスター不動産
            KBN_NAME = "ケイアイスター不動産"
        Case "0000000212012"
            KBN_CHG = "B02" '一建設
            KBN_NAME = "一建設"
        Case "0000000717000"
            KBN_CHG = "B03" 'ｸﾞランディハウス
            KBN_NAME = "ｸﾞﾗﾝﾃﾞｨﾊｳｽ"
        Case "0000000711001"
            KBN_CHG = "B04" '県民共済
            KBN_NAME = "県民共済"
        Case "0000000110775"
            KBN_CHG = "B05" '東栄住宅
            KBN_NAME = "東栄住宅"
        Case "0000000711014"
            KBN_CHG = "B06" 'アゲル
            KBN_NAME = "アゲル"
        Case "0000000210215"
            KBN_CHG = "B07" 'アイダ設計
            KBN_NAME = "アイダ設計"
        Case "0000000710014"
            KBN_CHG = "B08" 'ライブリーハウス
            KBN_NAME = "ライブリーハウス"
        Case Else
            KBN_CHG = "B99"
            KBN_NAME = "その他"
    End Select

End Function

Function KBN_CHGT(strTCD As String, strGCD As String, strKBN As String, strKBP As String) As String
    '============================================================================================================
    '区分変更処理
    '============================================================================================================
    'Gｺｰﾄﾞで得意先を判定
'    If strGCD = "0000000110215" Then
'        Stop
'    End If
    Select Case strGCD
        Case "0000000115024"
            KBN_CHGT = "A01" 'ﾀﾏﾎｰﾑ
            KBN_NAME = "タマホーム"
        Case "0000000212012"
            KBN_CHGT = "A02" '一建設
            KBN_NAME = "一建設"
        Case "0000000811701"
            KBN_CHGT = "A03" 'AVANTIA
            KBN_NAME = "AVANTIA"
        Case "0000000210215"
            KBN_CHGT = "A04" 'アイダ設計
            KBN_NAME = "アイダ設計"
        Case "0000000210700"
            KBN_CHGT = "A05" '飯田産業
            KBN_NAME = "飯田産業"
        Case "0000000812301"
            KBN_CHGT = "A06" 'ｱｴﾗﾎｰﾑ
            KBN_NAME = "ｱｴﾗﾎｰﾑ"
        Case "0000000811101"
            KBN_CHGT = "A07" 'ｱｯﾄﾊｳｼﾞﾝｸﾞ(ｸﾞｯﾄﾞﾘﾋﾞﾝｸﾞ)
            KBN_NAME = "ｱｯﾄﾊｳｼﾞﾝｸﾞ"
        Case "0000000812105"
            KBN_CHGT = "A09" '今高建築
            KBN_NAME = "その他"
        Case "0000000812204"
            KBN_CHGT = "A09" 'ﾀｲｺｳﾊｳｽ
            KBN_NAME = "その他"
        Case "0000000811401"
            KBN_CHGT = "A09" 'ｸﾗｼｽﾎｰﾑ
            KBN_NAME = "その他"
        Case "0000000212305"
            KBN_CHGT = "A08" 'ｱｲﾃﾞｨﾎｰﾑ
            KBN_NAME = "ｱｲﾃﾞｨﾎｰﾑ"
        Case "0000000115320"
            KBN_CHGT = "A09" '玉善
            KBN_NAME = "その他"
        Case "0000000111450"
            KBN_CHGT = "A09" 'ｻｰﾗ
            KBN_NAME = "その他"
        Case "0000000216450"
            KBN_CHGT = "B04" 'ﾎｸｴﾂ
            KBN_NAME = "ホクエツ"
        Case "0000000810001"
            KBN_CHGT = "B05" '浜名ﾜｰｸｽ
            KBN_NAME = "浜名ﾜｰｸｽ"
        Case "0000000810014"
            KBN_CHGT = "B09" '糸代製鋼
            KBN_NAME = "その他"
        Case "0000000819011"
            KBN_CHGT = "A08" '鳥居金属(大阪手摺)
            KBN_NAME = "鳥居金属(大阪手摺)"
        Case "0000000812803"
            KBN_CHGT = "A08" '鳥居金属(大阪手摺)
            KBN_NAME = "鳥居金属(大阪手摺)"
        Case "0000000819001" '鳥居金属
            If strTCD = "0000000812301" Or strTCD = "0000000812303" Or strTCD = "0000000812308" Or strTCD = "0000000812350" Then  'ｱｴﾗﾎｰﾑ(東京経由)
                KBN_CHGT = "A06"
                KBN_NAME = "ｱｴﾗﾎｰﾑ"
            ElseIf strTCD = "0000000819004" Then
                KBN_CHGT = "B04"
                KBN_NAME = "本部(ﾎｸｴﾂ他)"
            Else
                '商品区分(HINKB)で判定
                If Trim(strKBN) = "07" Or Trim(strKBN) = "16" Then      'ﾌﾞﾘｯｼﾞ
                    KBN_CHGT = "B01"
                    KBN_NAME = "鳥居金属ﾌﾞﾘｯｼﾞ"
                ElseIf Trim(strKBN) = "08" Then 'ﾐﾆﾛｰﾄﾞ
                    KBN_CHGT = "B02"
                    KBN_NAME = "鳥居金属ﾐﾆﾛｰﾄﾞ"
                ElseIf Trim(strKBN) = "09" Then 'TL
                    KBN_CHGT = "B03"
                    KBN_NAME = "鳥居金属TL"
                Else
                    KBN_CHGT = "C09"
                    KBN_NAME = "その他"
                End If
            End If
        Case "0000000819010"           '関東AL
            KBN_CHGT = "C02"
            KBN_NAME = "関東AL"
        Case "0000000830035"           '加工渡し
            KBN_CHGT = "W99"
            KBN_NAME = "加工渡し"
        Case "0000000830099"           'ﾌﾞﾘｯｼﾞ材料引落し
            KBN_CHGT = "Z99"
            KBN_NAME = "ﾌﾞﾘｯｼﾞ材料引落し"
        Case Else
            If strTCD = "0000000810999" Then '雑売上
                KBN_CHGT = "C09"
                KBN_NAME = "その他"
            ElseIf Trim(strKBP) = "01" Then
                KBN_CHGT = "A09" '手摺その他
                KBN_NAME = "その他"
            Else
                Select Case Trim(strKBN)
                    Case "11", "14", "16", "17"
                        KBN_CHGT = "A09" '手摺その他
                        KBN_NAME = "その他"
                    Case "07", "08", "09", "10"
                        KBN_CHGT = "B09" 'ブリッジその他
                        KBN_NAME = "その他"
                    Case Else
                        KBN_CHGT = "C09" 'その他
                        KBN_NAME = "その他"
                End Select
            End If
        End Select
End Function
