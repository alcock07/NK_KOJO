Attribute VB_Name = "M09_TAN"
Option Explicit

Sub GET_TAN_K(strDate As String)
    
    Dim cnA    As New ADODB.Connection
    Dim rsA    As New ADODB.Recordset
    Dim rsW    As New ADODB.Recordset
    Dim Cmd    As New ADODB.Command
    Dim strSQL As String
    Dim strDay As String '最終売上日付
    Dim strTAN As String '担当者コード
    Dim strTNM As String '担当者名

    '☆=============== 売上データ処理 ======================================☆

    'T_売上から当月実績Tにデータ追加
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnA.Open
    strSQL = ""
    strSQL = strSQL & "INSERT INTO"
    strSQL = strSQL & "         W_KA_NKT  (SMADT"
    strSQL = strSQL & "                   ,TANCD"
    strSQL = strSQL & "                   ,URIKNR"
    strSQL = strSQL & "                   ,GENKNR)"
    strSQL = strSQL & "           SELECT '" & strDate & "'"
    strSQL = strSQL & "                   ,TANCD"
    strSQL = strSQL & "                   ,sum(URIKIN) - sum(ZKMUZEKN)"
    strSQL = strSQL & "                   ,sum(GENKIN)"
    strSQL = strSQL & "           FROM W_KA_URI"
    strSQL = strSQL & "           WHERE TOKCD < '0000000730000'"
    strSQL = strSQL & "           GROUP BY TANCD"
    Set Cmd.ActiveConnection = cnA
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
   '売上の最終日付を取得して当月実績Tテーブルの売上日付を更新する
    
    '最終日付取得
    strSQL = ""
    strSQL = strSQL & "SELECT UDNDT"
    strSQL = strSQL & "       FROM W_KA_URI"
    strSQL = strSQL & "       ORDER BY UDNDT DESC"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    If rsA.EOF Then 'データなければ受注処理へ
        rsA.Close
        GoTo JUC_ZAN
    Else
        rsA.MoveFirst
        strDay = rsA(0)
    End If
    If rsA.State = adStateOpen Then rsA.Close
    
     ' /// 当日売りﾃﾞｰﾀ取得 ///
    strSQL = ""
    strSQL = strSQL & "UPDATE W_KA_NKT"
    strSQL = strSQL & "       SET W_KA_NKT.URIKN = URI.UK"
    strSQL = strSQL & "          ,W_KA_NKT.GENKN = URI.GK"
    strSQL = strSQL & "          ,W_KA_NKT.UDNDT = URI.UDNDT"
    strSQL = strSQL & "       FROM (SELECT TANCD"
    strSQL = strSQL & "                   ,sum(URIKIN) - sum(ZKMUZEKN) as UK"
    strSQL = strSQL & "                   ,sum(GENKIN) as GK"
    strSQL = strSQL & "                   ,UDNDT"
    strSQL = strSQL & "             FROM W_KA_URI"
    strSQL = strSQL & "             WHERE UDNDT = " & strDay
    strSQL = strSQL & "             And TOKCD < '0000000730000'"
    strSQL = strSQL & "             GROUP BY TANCD"
    strSQL = strSQL & "                     ,UDNDT"
    strSQL = strSQL & "             ) as URI"
    strSQL = strSQL & "       WHERE W_KA_NKT.TANCD = URI.TANCD"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute

'☆=============== 計画データ処理 ======================================☆
PLN_DATA:
    
    strSQL = ""
    strSQL = strSQL & "SELECT TANCD"
    strSQL = strSQL & "       ,sum(PUKN) as URI"
    strSQL = strSQL & "       ,sum(PAKN) as ARA"
    strSQL = strSQL & "       FROM W_PLN"
    strSQL = strSQL & "            GROUP BY TANCD"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    If rsA.EOF = False Then
        rsA.MoveFirst
        Do Until rsA.EOF
            'その他担当者で計画だけある場合の処理
            strTAN = Trim(rsA(0))
            If strTAN = "00000708" Then
                strTNM = "その他"
            End If
            strSQL = ""
            strSQL = strSQL & "SELECT * "
            strSQL = strSQL & "       FROM W_KA_NKT"
            strSQL = strSQL & "       WHERE TANCD = '" & strTAN & "'"
            rsW.Open strSQL, cnA, adOpenStatic, adLockOptimistic
            If rsW.EOF Then
                rsW.AddNew
                rsW.Fields("SMADT") = strDate
                rsW.Fields("TANCD") = rsA(0)
                rsW.Fields("PUKN") = rsA(1)
                rsW.Fields("PAKN") = rsA(2)
                rsW.Fields("TANNM") = strTNM
            Else
                rsW.Fields("PUKN") = rsA(1)
                rsW.Fields("PAKN") = rsA(2)
            End If
            rsW.Update
            rsW.Close
            rsA.MoveNext
        Loop
    End If
    If rsA.State = adStateOpen Then rsA.Close
    
    
'☆=============== 受注データ処理 ======================================☆
JUC_ZAN:

    strSQL = ""
    strSQL = strSQL & "UPDATE W_KA_NKT"
    strSQL = strSQL & "       SET W_KA_NKT.JUZ = JUC.ZAN"
    strSQL = strSQL & "       FROM (SELECT TANCD"
    strSQL = strSQL & "                   ,sum(ZANKN) as ZAN"
    strSQL = strSQL & "             FROM W_KA_JUZ"
    strSQL = strSQL & "             GROUP BY TANCD"
    strSQL = strSQL & "             ) as JUC"
    strSQL = strSQL & "       WHERE W_KA_NKT.TANCD = JUC.TANCD"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
    
    '担当者名更新
    strSQL = ""
    strSQL = strSQL & "UPDATE W_KA_NKT"
    strSQL = strSQL & "       SET TANNM = MST.TANNM"
    strSQL = strSQL & "       FROM OPENQUERY ([ORA],"
    strSQL = strSQL & "                       'SELECT TANCD,"
    strSQL = strSQL & "                               TANNM"
    strSQL = strSQL & "                        FROM   TANMTA"
    strSQL = strSQL & "                        WHERE  DATKB = ''1''"
    strSQL = strSQL & "                        And    TANCD <> ''00000708''"
    strSQL = strSQL & "                               ') as MST"
    strSQL = strSQL & "       INNER JOIN W_KA_NKT"
    strSQL = strSQL & "                  ON MST.TANCD = W_KA_NKT.TANCD"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
Exit_DB:

    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    If Not cnA Is Nothing Then
        If cnA.State = adStateOpen Then cnA.Close
        Set cnA = Nothing
    End If
    
    end_time = Timer
    Debug.Print "W_KA_NKT  " & (end_time - start_time)

End Sub

Sub GET_TAN_T(strDate As String)
    
    Dim cnA    As New ADODB.Connection
    Dim rsA    As New ADODB.Recordset
    Dim rsW    As New ADODB.Recordset
    Dim Cmd    As New ADODB.Command
    Dim strSQL As String
    Dim strDay As String '最終売上日付
    Dim strTAN As String '担当者コード
    Dim strTNM As String '担当者名

    '☆=============== 売上データ処理 ======================================☆

    'T_売上から当月実績Tにデータ追加
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnA.Open
    strSQL = ""
    strSQL = strSQL & "INSERT INTO"
    strSQL = strSQL & "           W_TA_NKT (SMADT"
    strSQL = strSQL & "                    ,TANCD"
    strSQL = strSQL & "                    ,URIKNR"
    strSQL = strSQL & "                    ,GENKNR)"
    strSQL = strSQL & "           SELECT '" & strDate & "'"
    strSQL = strSQL & "                    ,TANCD"
    strSQL = strSQL & "                    ,sum(URIKIN) - sum(ZKMUZEKN)"
    strSQL = strSQL & "                    ,sum(GENKIN)"
    strSQL = strSQL & "           FROM W_TA_URI"
    strSQL = strSQL & "           WHERE TOKCD < '0000000820000'"
    strSQL = strSQL & "           GROUP BY TANCD"
    Set Cmd.ActiveConnection = cnA
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
   '売上の最終日付を取得して当月実績Tテーブルの売上日付を更新する
   
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
    
     ' /// 当日売りﾃﾞｰﾀ取得 ///
    strSQL = ""
    strSQL = strSQL & "UPDATE W_TA_NKT"
    strSQL = strSQL & "       SET W_TA_NKT.URIKN = URI.UK"
    strSQL = strSQL & "          ,W_TA_NKT.GENKN = URI.GK"
    strSQL = strSQL & "          ,W_TA_NKT.UDNDT = URI.UDNDT"
    strSQL = strSQL & "       FROM (SELECT TANCD"
    strSQL = strSQL & "                   ,sum(URIKIN) - sum(ZKMUZEKN) as UK"
    strSQL = strSQL & "                   ,sum(GENKIN) as GK"
    strSQL = strSQL & "                   ,UDNDT"
    strSQL = strSQL & "             FROM W_TA_URI"
    strSQL = strSQL & "             WHERE UDNDT = " & strDay
    strSQL = strSQL & "             And TOKCD < '0000000820000'"
    strSQL = strSQL & "             GROUP BY TANCD"
    strSQL = strSQL & "                     ,UDNDT"
    strSQL = strSQL & "             ) as URI"
    strSQL = strSQL & "       WHERE W_TA_NKT.TANCD = URI.TANCD"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
'☆=============== 計画データ処理 ======================================☆
    
    strSQL = ""
    strSQL = strSQL & "SELECT TANCD"
    strSQL = strSQL & "      ,sum(PUKN) as URI"
    strSQL = strSQL & "      ,sum(PAKN) as ARA"
    strSQL = strSQL & "       FROM W_PLN"
    strSQL = strSQL & "            GROUP BY TANCD"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    If rsA.EOF = False Then
        rsA.MoveFirst
        Do Until rsA.EOF
            'その他担当者で計画だけある場合の処理
            strTAN = Trim(rsA(0)) & ""
            If strTAN <> "" Then
                strSQL = ""
                strSQL = strSQL & "SELECT * "
                strSQL = strSQL & "       FROM W_TA_NKT"
                strSQL = strSQL & "       WHERE TANCD = '" & strTAN & "'"
                rsW.Open strSQL, cnA, adOpenStatic, adLockOptimistic
                If rsW.EOF = False Then
                    rsW.Fields("PUKN") = rsA(1)
                    rsW.Fields("PAKN") = rsA(2)
                    rsW.Update
                End If
                rsW.Close
            End If
            rsA.MoveNext
        Loop
    End If
    If rsA.State = adStateOpen Then rsA.Close
    
    '☆=============== 受注データ処理 ======================================☆

    strSQL = ""
    strSQL = strSQL & "UPDATE W_TA_NKT"
    strSQL = strSQL & "       SET W_TA_NKT.JUZ = JUC.ZAN"
    strSQL = strSQL & "       FROM (SELECT TANCD"
    strSQL = strSQL & "                   ,sum(ZANKN) as ZAN"
    strSQL = strSQL & "             FROM W_TA_JUZ"
    strSQL = strSQL & "             GROUP BY TANCD"
    strSQL = strSQL & "             ) as JUC"
    strSQL = strSQL & "       WHERE W_TA_NKT.TANCD = JUC.TANCD"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
    '担当者名更新
    strSQL = ""
    strSQL = strSQL & "UPDATE W_TA_NKT"
    strSQL = strSQL & "       SET TANNM = MST.TANNM"
    strSQL = strSQL & "       FROM OPENQUERY ([ORA],"
    strSQL = strSQL & "                       'SELECT TANCD,"
    strSQL = strSQL & "                               TANNM"
    strSQL = strSQL & "                        FROM   TANMTA"
    strSQL = strSQL & "                        WHERE  DATKB = ''1''"
    strSQL = strSQL & "                               ') as MST"
    strSQL = strSQL & "       INNER JOIN W_TA_NKT"
    strSQL = strSQL & "                  ON MST.TANCD = W_TA_NKT.TANCD"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
Exit_DB:

    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    If Not cnA Is Nothing Then
        If cnA.State = adStateOpen Then cnA.Close
        Set cnA = Nothing
    End If
    
    end_time = Timer
    Debug.Print "W_TA_NKT " & (end_time - start_time)

End Sub
