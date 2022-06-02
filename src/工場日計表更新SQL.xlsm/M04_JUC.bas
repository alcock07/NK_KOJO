Attribute VB_Name = "M04_JUC"
Option Explicit

Sub GET_JUC_K(strDate As String)
    '============================================================================================================
    '受注トラン（納期、商品ｺｰﾄﾞ、受注残、得意先ｺｰﾄﾞ）
    '部門ｺｰﾄﾞ=070701、消費税除く
    '============================================================================================================
    '変数の宣言
    Dim cnA    As New ADODB.Connection
    Dim rsA    As New ADODB.Recordset
    Dim rsW    As New ADODB.Recordset
    Dim Cmd    As New ADODB.Command
    Dim strSQL As String
    Dim strNOK As String
    
    start_time = Timer
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnA.Open

    strNOK = Left(strDate, 6)
    strSQL = ""
    strSQL = strSQL & "INSERT INTO W_KA_JUZ(tancd"
    strSQL = strSQL & "                 ,tokcd"
    strSQL = strSQL & "                 ,toknm"
    strSQL = strSQL & "                 ,nokdt"
    strSQL = strSQL & "                 ,zankn"
    strSQL = strSQL & "                 ,gnkkn)"
    strSQL = strSQL & "            SELECT "
    strSQL = strSQL & "                  tancd"
    strSQL = strSQL & "                 ,tokcd"
    strSQL = strSQL & "                 ,toknm"
    strSQL = strSQL & "                 ,nokdt"
    strSQL = strSQL & "                 ,sum(zankn) as 受注残"
    strSQL = strSQL & "                 ,sum(gnkkn) as 原価"
    strSQL = strSQL & "            FROM JUZTBZ_Hybrid"
    strSQL = strSQL & "            WHERE bmncd = '070701'"
    strSQL = strSQL & "            And   Left(nokdt,6)  = " & strNOK
    strSQL = strSQL & "            GROUP BY tancd,tokcd,toknm,nokdt"

    Set Cmd.ActiveConnection = cnA
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
    'ｸﾞﾙｰﾌﾟｺｰﾄﾞ更新
    strSQL = ""
    strSQL = strSQL & "UPDATE W_KA_JUZ"
    strSQL = strSQL & "       SET GCODE = TOK.GRPCD,"
    strSQL = strSQL & "           TANCD = TOK.TANCD"
    strSQL = strSQL & "       FROM OPENQUERY ([ORA],"
    strSQL = strSQL & "                       'SELECT GRPCD,"
    strSQL = strSQL & "                               TANCD,"
    strSQL = strSQL & "                               TOKCD"
    strSQL = strSQL & "                        FROM TOKMTA') as TOK"
    strSQL = strSQL & "       INNER JOIN W_KA_JUZ"
    strSQL = strSQL & "       ON TOK.TOKCD = W_KA_JUZ.TOKCD"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
    strSQL = "SELECT * FROM W_KA_JUZ"
    rsW.Open strSQL, cnA, adOpenStatic, adLockOptimistic
    If rsW.EOF = False Then
        rsW.MoveFirst
    End If
    Do Until rsW.EOF
        If Trim(rsW.Fields("GCODE")) = "" Then
            rsW.Fields("GCODE") = rsW.Fields("TOKCD")
        End If
        rsW.Fields("NKBN") = KBN_CHG(rsW.Fields("TOKCD"), rsW.Fields("GCODE"))
        rsW.Fields("NKNM") = KBN_NAME
        rsW.Update
        rsW.MoveNext
    Loop
    
    end_time = Timer
    Debug.Print "W_KA_JUZ  " & (end_time - start_time)

Exit_DB:

    If Not rsW Is Nothing Then
        If rsW.State = adStateOpen Then rsW.Close
        Set rsW = Nothing
    End If
    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    If Not cnA Is Nothing Then
        If cnA.State = adStateOpen Then cnA.Close
        Set cnA = Nothing
    End If

End Sub

Sub GET_JUC_T(strDate As String)
    '============================================================================================================
    '受注トラン（納期、商品ｺｰﾄﾞ、受注残、得意先ｺｰﾄﾞ）
    '部門ｺｰﾄﾞ=070701、消費税除く
    '============================================================================================================
    '変数の宣言
    Dim cnA    As New ADODB.Connection
    Dim rsA    As New ADODB.Recordset
    Dim rsW    As New ADODB.Recordset
    Dim Cmd    As New ADODB.Command
    Dim strSQL As String
    
    start_time = Timer
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnA.Open

    strSQL = ""
    strSQL = strSQL & "INSERT INTO W_TA_JUZ("
    strSQL = strSQL & "                  tancd"
    strSQL = strSQL & "                 ,tokcd"
    strSQL = strSQL & "                 ,nokdt"
    strSQL = strSQL & "                 ,hincd"
    strSQL = strSQL & "                 ,zankn"
    strSQL = strSQL & "                 ,gnkkn)"
    strSQL = strSQL & "            SELECT "
    strSQL = strSQL & "                  tancd"
    strSQL = strSQL & "                 ,tokcd"
    strSQL = strSQL & "                 ,nokdt"
    strSQL = strSQL & "                 ,hincd"
    strSQL = strSQL & "                 ,sum(zankn) as 受注残"
    strSQL = strSQL & "                 ,sum(gnkkn) as 原価"
    strSQL = strSQL & "            FROM JUZTBZ_Hybrid"
    strSQL = strSQL & "            WHERE (bmncd = '080808' or bmncd = '080880')"
    strSQL = strSQL & "            And  nokdt  <= " & strDate
    strSQL = strSQL & "            GROUP BY tancd,tokcd,nokdt,hincd"

    Set Cmd.ActiveConnection = cnA
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
    'ｸﾞﾙｰﾌﾟｺｰﾄﾞ更新
    strSQL = ""
    strSQL = strSQL & "UPDATE W_TA_JUZ"
    strSQL = strSQL & "       SET GCODE = TOK.GRPCD,"
    strSQL = strSQL & "           TANCD = TOK.TANCD"
    strSQL = strSQL & "       FROM OPENQUERY ([ORA],"
    strSQL = strSQL & "                       'SELECT GRPCD,"
    strSQL = strSQL & "                               TANCD,"
    strSQL = strSQL & "                               TOKCD"
    strSQL = strSQL & "                        FROM TOKMTA') as TOK"
    strSQL = strSQL & "       INNER JOIN W_TA_JUZ"
    strSQL = strSQL & "       ON TOK.TOKCD = W_TA_JUZ.TOKCD"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
    '商品区分更新
    strSQL = ""
    strSQL = strSQL & "UPDATE W_TA_JUZ"
    strSQL = strSQL & "       SET HINBID = HIN.HINCLBID,"
    strSQL = strSQL & "           HINCID = HIN.HINCLCID"
    strSQL = strSQL & "       FROM OPENQUERY ([ORA],"
    strSQL = strSQL & "                       'SELECT HINCLBID,"
    strSQL = strSQL & "                               HINCLCID,"
    strSQL = strSQL & "                               HINCD"
    strSQL = strSQL & "                        FROM HINMTA') as HIN"
    strSQL = strSQL & "       INNER JOIN W_TA_JUZ"
    strSQL = strSQL & "       ON HIN.HINCD = W_TA_JUZ.HINCD"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
    strSQL = "SELECT * FROM W_TA_JUZ"
    rsW.Open strSQL, cnA, adOpenStatic, adLockOptimistic
    If rsW.EOF = False Then
        rsW.MoveFirst
    End If
    
    Dim strHINBID As String
    Do Until rsW.EOF
        If Trim(rsW.Fields("GCODE")) = "" Then
            rsW.Fields("GCODE") = rsW.Fields("TOKCD")
        End If
        KBN_NAME = ""
        
        If IsNull(rsW.Fields("HINBID")) Then
            strHINBID = ""
        Else
            strHINBID = rsW.Fields("HINBID")
        End If
            
        rsW.Fields("NKBN") = KBN_CHGT(rsW.Fields("TOKCD"), rsW.Fields("GCODE"), strHINBID, "")
        rsW.Fields("NKNM") = KBN_NAME
        rsW.Update
        rsW.MoveNext
    Loop
    
    end_time = Timer
    Debug.Print "W_TA_JUZ " & (end_time - start_time)

Exit_DB:

    If Not rsW Is Nothing Then
        If rsW.State = adStateOpen Then rsW.Close
        Set rsW = Nothing
    End If
    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    If Not cnA Is Nothing Then
        If cnA.State = adStateOpen Then cnA.Close
        Set cnA = Nothing
    End If

End Sub
