Attribute VB_Name = "M08_TOK"
Option Explicit

'=== 作業テーブルに売上、受注、仕入、在庫を入れる ======

Sub GET_TOK_K(strDate As String)

    Dim cnA    As New ADODB.Connection
    Dim rsA    As New ADODB.Recordset
    Dim rsX    As New ADODB.Recordset
    Dim Cmd    As New ADODB.Command
    Dim strSQL As String
    
    start_time = Timer
    
'☆=============== 売上データ処理 ======================================☆

    'W_KA_URIから当月実績にデータ追加
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnA.Open
    strSQL = ""
    strSQL = strSQL & "INSERT INTO"
    strSQL = strSQL & "        W_KA_NK    (SMADT"
    strSQL = strSQL & "                   ,GCODE"
    strSQL = strSQL & "                   ,NKBN"
    strSQL = strSQL & "                   ,NKNM"
    strSQL = strSQL & "                   ,URIKNR"
    strSQL = strSQL & "                   ,GENKNR)"
    strSQL = strSQL & "           SELECT '" & strDate & "'"
    strSQL = strSQL & "                   ,GCODE"
    strSQL = strSQL & "                   ,NKBN"
    strSQL = strSQL & "                   ,NKNM"
    strSQL = strSQL & "                   ,sum(URIKIN) - sum(ZKMUZEKN)"
    strSQL = strSQL & "                   ,sum(GENKIN)"
    strSQL = strSQL & "           FROM W_KA_URI"
    strSQL = strSQL & "           WHERE TOKCD < '0000000730000'"
    strSQL = strSQL & "           GROUP BY GCODE"
    strSQL = strSQL & "                   ,NKBN"
    strSQL = strSQL & "                   ,NKNM"
    Set Cmd.ActiveConnection = cnA
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
    '渡しﾃﾞｰﾀ追加
    strSQL = ""
    strSQL = strSQL & "INSERT INTO"
    strSQL = strSQL & "        W_KA_NK    (SMADT"
    strSQL = strSQL & "                   ,GCODE"
    strSQL = strSQL & "                   ,NKBN"
    strSQL = strSQL & "                   ,NKNM"
    strSQL = strSQL & "                   ,URIKNR)"
    strSQL = strSQL & "           SELECT '" & strDate & "'"
    strSQL = strSQL & "                   ,GCODE"
    strSQL = strSQL & "                   ,'W99'"
    strSQL = strSQL & "                   ,'渡し'"
    strSQL = strSQL & "                   ,sum(URIKIN)"
    strSQL = strSQL & "           FROM W_KA_URI"
    strSQL = strSQL & "        　 WHERE SMADT = '" & strDate & "'"
    strSQL = strSQL & "           And (TOKCD = '0000000730030'"
    strSQL = strSQL & "           Or TOKCD = '0000000730035')"
    strSQL = strSQL & "           GROUP BY GCODE"
    Set Cmd.ActiveConnection = cnA
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
        
    '/// 売上の最終日付を取得して当月実績テーブルの売上日付を更新する ///
    '最終日付取得
    strSQL = ""
    strSQL = strSQL & "SELECT UDNDT"
    strSQL = strSQL & "       FROM W_KA_URI"
    strSQL = strSQL & "       ORDER BY UDNDT DESC"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
    ' /// 当日売りﾃﾞｰﾀ取得 ///
    Dim strDay   As String '最終売上日付
    strDay = ""
    If rsA.EOF = False Then '当日日付取得
        rsA.MoveFirst
        strDay = rsA(0)
    End If
    If rsA.State = adStateOpen Then rsA.Close
    
    If strDay <> "" Then
        strSQL = ""
        strSQL = strSQL & "UPDATE W_KA_NK"
        strSQL = strSQL & "       SET W_KA_NK.URIKN = URI.UK"
        strSQL = strSQL & "          ,W_KA_NK.GENKN = URI.GK"
        strSQL = strSQL & "          ,W_KA_NK.UDNDT = URI.UDNDT"
        strSQL = strSQL & "       FROM (SELECT GCODE"
        strSQL = strSQL & "                    ,sum(URIKIN) - sum(ZKMUZEKN) as UK"
        strSQL = strSQL & "                    ,sum(GENKIN) as GK"
        strSQL = strSQL & "                    ,UDNDT"
        strSQL = strSQL & "             FROM W_KA_URI"
        strSQL = strSQL & "             WHERE UDNDT = " & strDay
        strSQL = strSQL & "             And TOKCD < '0000000730000'"
        strSQL = strSQL & "             GROUP BY GCODE"
        strSQL = strSQL & "                     ,UDNDT"
        strSQL = strSQL & "             ) as URI"
        strSQL = strSQL & "       WHERE W_KA_NK.GCODE = URI.GCODE"
        Cmd.CommandText = strSQL
        Set rsA = Cmd.Execute
    End If

'☆=============== 計画データ処理 ======================================☆
    
    Dim lngPU As Long
    Dim lngPA As Long
    
    strSQL = ""
    strSQL = strSQL & "SELECT GCODE"
    strSQL = strSQL & "       ,NKBN"
    strSQL = strSQL & "       ,NKNM"
    strSQL = strSQL & "       ,sum(PUKN)"
    strSQL = strSQL & "       ,sum(PAKN)"
    strSQL = strSQL & "   FROM W_PLN"
    strSQL = strSQL & "   GROUP BY GCODE"
    strSQL = strSQL & "   　　　   ,NKBN "
    strSQL = strSQL & "   　　　   ,NKNM"
    strSQL = strSQL & "   ORDER BY GCODE"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
    If rsA.EOF = False Then
        rsA.MoveFirst
        Do Until rsA.EOF
            strSQL = ""
            strSQL = strSQL & "SELECT * "
            strSQL = strSQL & "      FROM W_KA_NK"
            strSQL = strSQL & "           WHERE GCODE = '" & rsA.Fields("GCODE") & "'"
            rsX.Open strSQL, cnA, adOpenStatic, adLockOptimistic
            If rsX.EOF Then
                rsX.AddNew
                rsX.Fields("SMADT") = strDate '締日付
                rsX.Fields("GCODE") = rsA(0)  'Gｺｰﾄﾞ
                rsX.Fields("NKBN") = rsA(1)   '日計区分
                rsX.Fields("NKNM") = rsA(2)   '名称
                rsX.Fields("PUKN") = rsA(3)   '計画売り
                rsX.Fields("PAKN") = rsA(4)   '計画粗利
            Else
                rsX.Fields("PUKN") = rsX.Fields("PUKN") + rsA(3)
                rsX.Fields("PAKN") = rsX.Fields("PAKN") + rsA(4)
            End If
            rsX.Update
            rsX.Close
            rsA.MoveNext
        Loop
       
    End If
    If rsA.State = adStateOpen Then rsA.Close
    
'☆=============== 受注データ処理 ======================================☆

    '受注残ﾃﾞｰﾀ取得
    strSQL = ""
    strSQL = strSQL & "SELECT GCODE"
    strSQL = strSQL & "       ,NKBN"
    strSQL = strSQL & "       ,NKNM"
    strSQL = strSQL & "       ,sum(ZANKN) as ZAN"
    strSQL = strSQL & "   FROM W_KA_JUZ"
    strSQL = strSQL & "   GROUP BY GCODE"
    strSQL = strSQL & "   　　　   ,NKBN "
    strSQL = strSQL & "   　　　   ,NKNM"

    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
    If rsA.EOF = False Then
        rsA.MoveFirst
        Do Until rsA.EOF
            strSQL = ""
            strSQL = strSQL & "SELECT * "
            strSQL = strSQL & "      FROM W_KA_NK"
            strSQL = strSQL & "           WHERE GCODE = '" & rsA.Fields("GCODE") & "'"
            rsX.Open strSQL, cnA, adOpenStatic, adLockOptimistic
            If rsX.EOF Then
                rsX.AddNew
                rsX.Fields("SMADT") = strDate '締日付
                rsX.Fields("GCODE") = rsA(0)  'Gｺｰﾄﾞ
                rsX.Fields("NKBN") = rsA(1)   '日計区分
                rsX.Fields("NKNM") = rsA(2)   '名称
                rsX.Fields("JUZ") = rsA(3)    '受注残
            Else
                rsX.Fields("JUZ") = rsX.Fields("JUZ") + rsA(3)
            End If
            rsX.Update
            rsX.Close
            rsA.MoveNext
        Loop
       
    End If
    If rsA.State = adStateOpen Then rsA.Close
    
'☆=============== 仕入データ処理 ======================================☆

    Dim lngSIR   As Long   '仕入金額
    
    strSQL = ""
    strSQL = strSQL & "SELECT GKBN"
    strSQL = strSQL & "       ,sum(SIRKIN)"
    strSQL = strSQL & "       FROM W_KA_SRE"
    strSQL = strSQL & "       WHERE GKBN = 'S'"
    strSQL = strSQL & "       GROUP BY GKBN"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    If rsA.EOF Then
        lngSIR = 0
    Else
        rsA.MoveFirst
        lngSIR = rsA(1)
    End If
    rsA.Close
    If lngSIR <> 0 Then
        strSQL = ""
        strSQL = strSQL & "SELECT *"
        strSQL = strSQL & "       FROM W_KA_NK"
        strSQL = strSQL & "       WHERE NKBN = 'S99'"
        rsA.Open strSQL, cnA, adOpenStatic, adLockOptimistic
        If rsA.EOF Then
            rsA.AddNew
            rsA.Fields("SMADT") = strDate
            rsA.Fields("GCODE") = "0000000729999"
        Else
            rsA.MoveFirst
        End If
        rsA.Fields("NKBN") = "S99"
        rsA.Fields("NKNM") = "仕入"
        rsA.Fields("SIRKN") = lngSIR
        rsA.Update
    End If
    If rsA.State = adStateOpen Then rsA.Close
    
    '受けﾃﾞｰﾀ追加
    strSQL = ""
    strSQL = strSQL & "INSERT INTO"
    strSQL = strSQL & "           W_KA_NK (SMADT"
    strSQL = strSQL & "                   ,GCODE"
    strSQL = strSQL & "                   ,NKBN"
    strSQL = strSQL & "                   ,NKNM"
    strSQL = strSQL & "                   ,SIRKN)"
    strSQL = strSQL & "           SELECT '" & strDate & "'"
    strSQL = strSQL & "                   ,'0000000749999'"
    strSQL = strSQL & "                   ,'U99'"
    strSQL = strSQL & "                   ,'受け'"
    strSQL = strSQL & "                   ,sum(SIRKIN)"
    strSQL = strSQL & "           FROM W_KA_SRE"
    strSQL = strSQL & "        　 WHERE SMADT = '" & strDate & "'"
    strSQL = strSQL & "           And GKBN = 'U'"
    strSQL = strSQL & "           GROUP BY GKBN"
    Set Cmd.ActiveConnection = cnA
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
Exit_DB:

    If Not rsX Is Nothing Then
        If rsX.State = adStateOpen Then rsX.Close
        Set rsX = Nothing
    End If
    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    If Not cnA Is Nothing Then
        If cnA.State = adStateOpen Then cnA.Close
        Set cnA = Nothing
    End If
    
    end_time = Timer
    Debug.Print "W_KA_NK   " & (end_time - start_time)
    
End Sub

Sub GET_TOK_T(strDate As String)

    Dim cnA    As New ADODB.Connection
    Dim rsA    As New ADODB.Recordset
    Dim rsX    As New ADODB.Recordset
    Dim Cmd    As New ADODB.Command
    Dim strSQL As String
   
    Dim strDay   As String '最終売上日付
    Dim lngSIR   As Long   '仕入金額

    start_time = Timer
    
'☆=============== 売上データ処理 ======================================☆

    'W_TA_URIからW_KA_NKにデータ追加
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnA.Open
    strSQL = ""
    strSQL = strSQL & "INSERT INTO"
    strSQL = strSQL & "           W_TA_NK (SMADT"
    strSQL = strSQL & "                   ,NKBN"
    strSQL = strSQL & "                   ,NKNM"
    strSQL = strSQL & "                   ,URIKNR"
    strSQL = strSQL & "                   ,GENKNR)"
    strSQL = strSQL & "           SELECT '" & strDate & "'"
    strSQL = strSQL & "                   ,NKBN"
    strSQL = strSQL & "                   ,NKNM"
    strSQL = strSQL & "                   ,sum(URIKIN) - sum(ZKMUZEKN)"
    strSQL = strSQL & "                   ,sum(GENKIN)"
    strSQL = strSQL & "           FROM W_TA_URI"
    strSQL = strSQL & "           WHERE TOKCD < '0000000820000'"
    strSQL = strSQL & "           GROUP BY NKBN"
    strSQL = strSQL & "                   ,NKNM"
    Set Cmd.ActiveConnection = cnA
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
    '渡しﾃﾞｰﾀ追加
    strSQL = ""
    strSQL = strSQL & "INSERT INTO"
    strSQL = strSQL & "           W_TA_NK (SMADT"
    strSQL = strSQL & "                   ,NKBN"
    strSQL = strSQL & "                   ,NKNM"
    strSQL = strSQL & "                   ,URIKNR)"
    strSQL = strSQL & "           SELECT '" & strDate & "'"
    strSQL = strSQL & "                   ,NKBN"
    strSQL = strSQL & "                   ,NKNM"
    strSQL = strSQL & "                   ,sum(URIKIN)"
    strSQL = strSQL & "           FROM W_TA_URI"
    strSQL = strSQL & "        　 WHERE SMADT = '" & strDate & "'"
    strSQL = strSQL & "           And TOKCD = '0000000830035'"
    strSQL = strSQL & "           GROUP BY NKBN"
    strSQL = strSQL & "                   ,NKNM"
    Set Cmd.ActiveConnection = cnA
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
    'ﾌﾞﾘｯｼﾞ材料引落しﾃﾞｰﾀ追加
    strSQL = ""
    strSQL = strSQL & "INSERT INTO"
    strSQL = strSQL & "          W_TA_NK  (SMADT"
    strSQL = strSQL & "                   ,NKBN"
    strSQL = strSQL & "                   ,NKNM"
    strSQL = strSQL & "                   ,URIKNR)"
    strSQL = strSQL & "           SELECT '" & strDate & "'"
    strSQL = strSQL & "                   ,NKBN"
    strSQL = strSQL & "                   ,NKNM"
    strSQL = strSQL & "                   ,sum(URIKIN)"
    strSQL = strSQL & "           FROM W_TA_URI"
    strSQL = strSQL & "        　 WHERE SMADT = '" & strDate & "'"
    strSQL = strSQL & "           And TOKCD = '0000000830099'"
    strSQL = strSQL & "           GROUP BY NKBN"
    strSQL = strSQL & "                   ,NKNM"
    Set Cmd.ActiveConnection = cnA
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
    '/// 売上の最終日付を取得して当月実績テーブルの売上日付を更新する ///
    '最終日付取得
    strSQL = ""
    strSQL = strSQL & "SELECT UDNDT"
    strSQL = strSQL & "       FROM W_TA_URI"
    strSQL = strSQL & "       ORDER BY UDNDT DESC"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
    ' /// 当日売りﾃﾞｰﾀ取得 ///
    strDay = ""
    If rsA.EOF = False Then '当日日付取得
        rsA.MoveFirst
        strDay = rsA(0)
    End If
    If rsA.State = adStateOpen Then rsA.Close
    
    If strDay <> "" Then
        strSQL = ""
        strSQL = strSQL & "UPDATE W_TA_NK"
        strSQL = strSQL & "       SET W_TA_NK.URIKN = URI.UK"
        strSQL = strSQL & "          ,W_TA_NK.GENKN = URI.GK"
        strSQL = strSQL & "          ,W_TA_NK.UDNDT = URI.UDNDT"
        strSQL = strSQL & "       FROM (SELECT NKBN"
        strSQL = strSQL & "                   ,sum(URIKIN) - sum(ZKMUZEKN) as UK"
        strSQL = strSQL & "                   ,sum(GENKIN) as GK"
        strSQL = strSQL & "                   ,UDNDT"
        strSQL = strSQL & "             FROM W_TA_URI"
        strSQL = strSQL & "             WHERE UDNDT = " & strDay
        strSQL = strSQL & "             And TOKCD < '0000000820000'"
        strSQL = strSQL & "             GROUP BY NKBN"
        strSQL = strSQL & "                     ,UDNDT"
        strSQL = strSQL & "             ) as URI"
        strSQL = strSQL & "       WHERE W_TA_NK.NKBN = URI.NKBN"
        Cmd.CommandText = strSQL
        Set rsA = Cmd.Execute
    End If

'☆=============== 計画データ処理 ======================================☆
    
    Dim lngPU As Long
    Dim lngPA As Long
    
    strSQL = ""
    strSQL = strSQL & "SELECT NKBN"
    strSQL = strSQL & "      ,NKNM"
    strSQL = strSQL & "      ,sum(PUKN)"
    strSQL = strSQL & "      ,sum(PAKN)"
    strSQL = strSQL & "   FROM W_PLN"
    strSQL = strSQL & "   GROUP BY NKBN"
    strSQL = strSQL & "   　　　   ,NKNM"
    strSQL = strSQL & "   ORDER BY NKBN"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
    If rsA.EOF = False Then
        rsA.MoveFirst
        Do Until rsA.EOF
            strSQL = ""
            strSQL = strSQL & "SELECT * "
            strSQL = strSQL & "      FROM W_TA_NK"
            strSQL = strSQL & "           WHERE NKBN = '" & rsA.Fields("NKBN") & "'"
            rsX.Open strSQL, cnA, adOpenStatic, adLockOptimistic
            If rsX.EOF Then
                rsX.AddNew
                rsX.Fields("SMADT") = strDate '締日付
                rsX.Fields("NKBN") = rsA(0)   '日計区分
                rsX.Fields("NKNM") = rsA(1)   '名称
                rsX.Fields("PUKN") = rsA(2)   '計画売り
                rsX.Fields("PAKN") = rsA(3)   '計画粗利
            Else
                rsX.Fields("PUKN") = rsX.Fields("PUKN") + rsA(2)
                rsX.Fields("PAKN") = rsX.Fields("PAKN") + rsA(3)
            End If
            rsX.Update
            rsX.Close
            rsA.MoveNext
        Loop
       
    End If
    If rsA.State = adStateOpen Then rsA.Close
    
'☆=============== 受注データ処理 ======================================☆

    '受注残ﾃﾞｰﾀ取得
    strSQL = ""
    strSQL = strSQL & "SELECT NKBN"
    strSQL = strSQL & "      ,NKNM"
    strSQL = strSQL & "      ,sum(ZANKN) as ZAN"
    strSQL = strSQL & "   FROM W_TA_JUZ"
    strSQL = strSQL & "   GROUP BY NKBN"
    strSQL = strSQL & "   　　　  ,NKNM"

    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
    If rsA.EOF = False Then
        rsA.MoveFirst
        Do Until rsA.EOF
            strSQL = ""
            strSQL = strSQL & "SELECT * "
            strSQL = strSQL & "      FROM W_TA_NK"
            strSQL = strSQL & "           WHERE NKBN = '" & rsA.Fields("NKBN") & "'"
            rsX.Open strSQL, cnA, adOpenStatic, adLockOptimistic
            If rsX.EOF Then
                rsX.AddNew
                rsX.Fields("SMADT") = strDate '締日付
                rsX.Fields("NKBN") = rsA(0)   '日計区分
                rsX.Fields("NKNM") = rsA(1)   '名称
                rsX.Fields("JUZ") = rsA(2)    '受注残
            Else
                rsX.Fields("JUZ") = rsX.Fields("JUZ") + rsA(2)
            End If
            rsX.Update
            rsX.Close
            rsA.MoveNext
        Loop
       
    End If
    If rsA.State = adStateOpen Then rsA.Close
    
'☆=============== 仕入データ処理 ======================================☆
    
    strSQL = ""
    strSQL = strSQL & "SELECT GKBN"
    strSQL = strSQL & "       ,sum(SIRKIN)"
    strSQL = strSQL & "     FROM W_TA_SRE"
    strSQL = strSQL & "          WHERE GKBN = 'S'"
    strSQL = strSQL & "          GROUP BY GKBN"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    If rsA.EOF Then
        lngSIR = 0
    Else
        rsA.MoveFirst
        lngSIR = rsA(1)
    End If
    rsA.Close
    If lngSIR <> 0 Then
        strSQL = ""
        strSQL = strSQL & "SELECT *"
        strSQL = strSQL & "       FROM W_TA_NK"
        strSQL = strSQL & "       WHERE NKBN = 'S99'"
        rsA.Open strSQL, cnA, adOpenStatic, adLockOptimistic
        If rsA.EOF Then
            rsA.AddNew
            rsA.Fields("SMADT") = strDate
        Else
            rsA.MoveFirst
        End If
        rsA.Fields("NKBN") = "S99"
        rsA.Fields("NKNM") = "仕入"
        rsA.Fields("SIRKN") = lngSIR
        rsA.Update
    End If
    If rsA.State = adStateOpen Then rsA.Close
    
    '受けﾃﾞｰﾀ追加
    strSQL = ""
    strSQL = strSQL & "INSERT INTO"
    strSQL = strSQL & "           W_TA_NK (SMADT"
    strSQL = strSQL & "                   ,NKBN"
    strSQL = strSQL & "                   ,NKNM"
    strSQL = strSQL & "                   ,SIRKN)"
    strSQL = strSQL & "           SELECT '" & strDate & "'"
    strSQL = strSQL & "                   ,'U99'"
    strSQL = strSQL & "                   ,'受け'"
    strSQL = strSQL & "                   ,sum(SIRKIN)"
    strSQL = strSQL & "           FROM W_TA_SRE"
    strSQL = strSQL & "        　 WHERE SMADT = '" & strDate & "'"
    strSQL = strSQL & "           And GKBN = 'U'"
    strSQL = strSQL & "           GROUP BY GKBN"
    Set Cmd.ActiveConnection = cnA
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
    '加工受けﾃﾞｰﾀ追加
    strSQL = ""
    strSQL = strSQL & "INSERT INTO"
    strSQL = strSQL & "           W_TA_NK (SMADT"
    strSQL = strSQL & "                   ,NKBN"
    strSQL = strSQL & "                   ,NKNM"
    strSQL = strSQL & "                   ,SIRKN)"
    strSQL = strSQL & "           SELECT '" & strDate & "'"
    strSQL = strSQL & "                   ,'R99'"
    strSQL = strSQL & "                   ,'加工受け'"
    strSQL = strSQL & "                   ,sum(SIRKIN)"
    strSQL = strSQL & "           FROM W_KA_SRE"
    strSQL = strSQL & "        　 WHERE SMADT = '" & strDate & "'"
    strSQL = strSQL & "           And GKBN = 'R'"
    strSQL = strSQL & "           GROUP BY GKBN"
    Set Cmd.ActiveConnection = cnA
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
    
Exit_DB:

    If Not rsX Is Nothing Then
        If rsX.State = adStateOpen Then rsX.Close
        Set rsX = Nothing
    End If
    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    If Not cnA Is Nothing Then
        If cnA.State = adStateOpen Then cnA.Close
        Set cnA = Nothing
    End If
    
    end_time = Timer
    Debug.Print "W_TA_NK  " & (end_time - start_time)
    
End Sub
