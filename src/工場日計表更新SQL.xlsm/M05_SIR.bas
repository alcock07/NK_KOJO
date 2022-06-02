Attribute VB_Name = "M05_SIR"
Option Explicit

Sub GET_SIR_K(strDate As String)
'============================================================================================================
'仕入トラン（商品ｺｰﾄﾞ、仕入金額、仕入先ｺｰﾄﾞ）
'    部門ｺｰﾄﾞ=010398（関東納入）
'             070791（外部仕入）
'             070792（外注先）
'             070785（加工）
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
    strSQL = "SELECT * FROM W_KA_SRE"
    rsW.Open strSQL, cnA, adOpenStatic, adLockOptimistic

    strSQL = ""
    strSQL = strSQL & "SELECT * FROM OPENQUERY ([ORA],"
    strSQL = strSQL & "        'SELECT HINCD"
    strSQL = strSQL & "               ,Sum(SREKN)"
    strSQL = strSQL & "               ,SIRCD"
    strSQL = strSQL & "               ,SIRBMNCD"
    strSQL = strSQL & "         FROM SDNTRA"
    strSQL = strSQL & "              WHERE  DATKB = ''1''" '削除区分
    strSQL = strSQL & "              And    LINNO < ''990''"  '消費税
    strSQL = strSQL & "              And    SMADT = ''" & strDate & "''"  '月指定
    strSQL = strSQL & "              And    SIRBMNCD IN(''070785'',''010398'',''070791'',''070792'')"  '関東仕入指定
    strSQL = strSQL & "         GROUP BY HINCD"
    strSQL = strSQL & "                 ,SIRCD"
    strSQL = strSQL & "                 ,SIRBMNCD"
    strSQL = strSQL & "         HAVING Sum(SREKN) <> 0"
    strSQL = strSQL & "         ')"
    Set Cmd.ActiveConnection = cnA
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    If rsA.EOF = False Then
        rsA.MoveFirst
    End If

    Do Until rsA.EOF
        rsW.AddNew
        rsW.Fields("SMADT") = strDate
        rsW.Fields("HINCD") = Trim(rsA.Fields(0))
        rsW.Fields("SIRKIN") = rsA.Fields(1)
        rsW.Fields("SIRCD") = rsA.Fields(2)
        '外注仕入の部門ｺｰﾄﾞ指定
        If rsA(3) = "070792" Then
            rsW.Fields("GKBN") = "G"
        ElseIf rsA(3) = "070785" Then
            rsW.Fields("GKBN") = "U"
        Else
            rsW.Fields("GKBN") = "S"
        End If
        rsW.Update
        rsA.MoveNext
    Loop

    end_time = Timer
    Debug.Print "W_KA_SRE  " & (end_time - start_time)
    
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

Sub GET_SIR_T(strDate As String)
    '============================================================================================================
    '仕入トラン
    '010397 --- 納入
    '080885 --- 加工
    '080891 --- 仕入先
    '080892 --- ﾌﾞﾘｯｼﾞ外注先
    '080893 --- 笠木外注先
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
    strSQL = "SELECT * FROM W_TA_SRE"
    rsW.Open strSQL, cnA, adOpenStatic, adLockOptimistic

    strSQL = ""
    strSQL = strSQL & "SELECT * FROM OPENQUERY ([ORA],"
    strSQL = strSQL & "             'SELECT SIRCD"
    strSQL = strSQL & "                     ,Sum(SREKN)"
    strSQL = strSQL & "                     ,SIRBMNCD"
    strSQL = strSQL & "              FROM   SDNTRA"
    strSQL = strSQL & "                     WHERE LINNO < ''990''"  '消費税
    strSQL = strSQL & "                     And   DATKB = ''1''"
    strSQL = strSQL & "                     And   SMADT = ''" & strDate & "''"  '月指定
    strSQL = strSQL & "                     And   SIRBMNCD IN(''080885'',''010397'',''080891'',''080892'',''080893'')"
    strSQL = strSQL & "              GROUP BY SIRCD"
    strSQL = strSQL & "                       ,SIRBMNCD"
    strSQL = strSQL & "              HAVING Sum(SREKN) <> 0"
    strSQL = strSQL & "             ')"
    Set Cmd.ActiveConnection = cnA
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    If rsA.EOF = False Then
        rsA.MoveFirst
    End If

    Do Until rsA.EOF
        rsW.AddNew
        rsW.Fields("SMADT") = strDate
        rsW.Fields("SIRCD") = rsA.Fields(0)
        rsW.Fields("SIRKIN") = rsA.Fields(1)
        '外注仕入の部門ｺｰﾄﾞ指定
        If rsA(2) = "080892" Or rsA(2) = "080893" Then
            rsW.Fields("GKBN") = "G"
        ElseIf rsA(0) = "0000000840001" Then
            rsW.Fields("GKBN") = "U"
        ElseIf rsA(0) = "0000000840011" Then
            rsW.Fields("GKBN") = "B"
        Else
            rsW.Fields("GKBN") = "S"
        End If
        rsW.Update
        rsA.MoveNext
    Loop

    end_time = Timer
    Debug.Print "W_TA_SRE " & (end_time - start_time)
    
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

Sub GET_SIR_T2(strDate As String)
'============================================================================================================
'仕入トラン
'080885 --- 加工
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
    strSQL = "SELECT * FROM W_KA_SRE"
    rsW.Open strSQL, cnA, adOpenStatic, adLockOptimistic

    strSQL = ""
    strSQL = strSQL & "SELECT * FROM OPENQUERY ([ORA],"
    strSQL = strSQL & "        'SELECT HINCD"
    strSQL = strSQL & "               ,Sum(SREKN)"
    strSQL = strSQL & "               ,SIRCD"
    strSQL = strSQL & "               ,SIRBMNCD"
    strSQL = strSQL & "         FROM SDNTRA"
    strSQL = strSQL & "              WHERE  DATKB = ''1''" '削除区分
    strSQL = strSQL & "              And    LINNO < ''990''"  '消費税
    strSQL = strSQL & "              And    SMADT = ''" & strDate & "''"  '月指定
    strSQL = strSQL & "              And    SIRBMNCD = ''080885''"        '東海加工
    strSQL = strSQL & "         GROUP BY HINCD"
    strSQL = strSQL & "                 ,SIRCD"
    strSQL = strSQL & "                 ,SIRBMNCD"
    strSQL = strSQL & "         HAVING Sum(SREKN) <> 0"
    strSQL = strSQL & "         ')"
    Set Cmd.ActiveConnection = cnA
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
    If rsA.EOF = False Then
        rsA.MoveFirst
    End If

    Do Until rsA.EOF
        rsW.AddNew
        rsW.Fields("SMADT") = strDate
        rsW.Fields("HINCD") = Trim(rsA.Fields(0))
        rsW.Fields("SIRKIN") = rsA.Fields(1)
        rsW.Fields("SIRCD") = rsA.Fields(2)
        rsW.Update
        rsA.MoveNext
    Loop
    
    '商品区分更新
    strSQL = ""
    strSQL = strSQL & "UPDATE W_KA_SRE"
    strSQL = strSQL & "       SET HINKB = HIN.HINCLCID"
    strSQL = strSQL & "       FROM OPENQUERY ([ORA],"
    strSQL = strSQL & "                       'SELECT HINCLCID,"
    strSQL = strSQL & "                               HINCD"
    strSQL = strSQL & "                        FROM HINMTA') as HIN"
    strSQL = strSQL & "       INNER JOIN W_KA_SRE"
    strSQL = strSQL & "       ON HIN.HINCD = W_KA_SRE.HINCD"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
    If rsW.EOF = False Then
        rsW.MoveFirst
    End If

    Do Until rsW.EOF
        If Left(rsW.Fields("HINKB"), 1) = "R" Or Left(rsW.Fields("HINKB"), 1) = "H" Then
            rsW.Fields("GKBN") = "R"
        Else
            rsW.Fields("GKBN") = "X"
        End If
        rsW.Update
        rsW.MoveNext
    Loop
   
    
    end_time = Timer
    Debug.Print "W_KA_SRE " & (end_time - start_time)
    
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
