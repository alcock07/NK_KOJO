Attribute VB_Name = "M05_SIR"
Option Explicit

Sub GET_SIR_K(strDate As String)
'============================================================================================================
'�d���g�����i���i���ށA�d�����z�A�d���溰�ށj
'    ���庰��=010398�i�֓��[���j
'             070791�i�O���d���j
'             070792�i�O����j
'             070785�i���H�j
'============================================================================================================
    
    '�ϐ��̐錾
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
    strSQL = strSQL & "              WHERE  DATKB = ''1''" '�폜�敪
    strSQL = strSQL & "              And    LINNO < ''990''"  '�����
    strSQL = strSQL & "              And    SMADT = ''" & strDate & "''"  '���w��
    strSQL = strSQL & "              And    SIRBMNCD IN(''070785'',''010398'',''070791'',''070792'')"  '�֓��d���w��
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
        '�O���d���̕��庰�ގw��
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
    '�d���g����
    '010397 --- �[��
    '080885 --- ���H
    '080891 --- �d����
    '080892 --- ��د�ފO����
    '080893 --- �}�؊O����
    '============================================================================================================
    '�ϐ��̐錾
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
    strSQL = strSQL & "                     WHERE LINNO < ''990''"  '�����
    strSQL = strSQL & "                     And   DATKB = ''1''"
    strSQL = strSQL & "                     And   SMADT = ''" & strDate & "''"  '���w��
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
        '�O���d���̕��庰�ގw��
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
'�d���g����
'080885 --- ���H
'============================================================================================================
    
    '�ϐ��̐錾
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
    strSQL = strSQL & "              WHERE  DATKB = ''1''" '�폜�敪
    strSQL = strSQL & "              And    LINNO < ''990''"  '�����
    strSQL = strSQL & "              And    SMADT = ''" & strDate & "''"  '���w��
    strSQL = strSQL & "              And    SIRBMNCD = ''080885''"        '���C���H
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
    
    '���i�敪�X�V
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
