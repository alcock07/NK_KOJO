Attribute VB_Name = "M03_URI"
Option Explicit

Sub GET_URI_K(strDate As String)
'============================================================================================================
'γgiγϊtA€iΊ°ΔήAγΰzA΄Ώΰzj
'εΊ°Δή=070701idόj
'         070709(ΈήΩ°Μίθ)
'         070781(n΅)
'`[ζͺ=2(γ)or 3(Ό)AΑοΕ­
'============================================================================================================
    Dim cnA    As New ADODB.Connection
    Dim rsP    As New ADODB.Recordset
    Dim rsA    As New ADODB.Recordset
    Dim rsW    As New ADODB.Recordset
    Dim Cmd    As New ADODB.Command
    Dim strSQL As String
    
    start_time = Timer
    
    'μΖΓ°ΜήΩ(SQL Server)
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnA.Open
    strSQL = "SELECT * FROM W_KA_URI"
    rsW.Open strSQL, cnA, adOpenStatic, adLockOptimistic
    
    'SQLΜSMADTΙϊΎ―Δ
    strSQL = ""
    strSQL = strSQL & "SELECT * FROM OPENQUERY ([ORA],"
    strSQL = strSQL & "                       'SELECT UDNDT"
    strSQL = strSQL & "                              ,Sum(URIKN) "
    strSQL = strSQL & "                              ,Sum(GNKKN) "
    strSQL = strSQL & "                              ,TOKCD "
    strSQL = strSQL & "                              ,Sum(ZKMUZEKN) "
    strSQL = strSQL & "                        FROM   UDNTRA"
    strSQL = strSQL & "                               WHERE  DATKB = ''1''"
    strSQL = strSQL & "                               And    TOKBMNCD IN(''070701'',''070709'',''070781'')"
    strSQL = strSQL & "                               And    DENKB IN(''2'',''3'')"
    strSQL = strSQL & "                               And    LINNO < ''990''"
    strSQL = strSQL & "                               And    SMADT = ''" & strDate & "''"
    strSQL = strSQL & "                        GROUP BY UDNDT"
    strSQL = strSQL & "                                ,TOKCD"
    strSQL = strSQL & "                        ')"
    Set Cmd.ActiveConnection = cnA
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    If rsA.EOF = False Then
        rsA.MoveFirst
    End If
    
    Do Until rsA.EOF
        'ΰz ιͺΎ―
        If rsA.Fields(1) <> 0 Or rsA.Fields(2) <> 0 Or rsA.Fields(4) <> 0 Then
            rsW.AddNew
            rsW.Fields("SMADT") = strDate
            rsW.Fields("UDNDT") = rsA.Fields(0)
            rsW.Fields("URIKIN") = rsA.Fields(1)
            rsW.Fields("GENKIN") = rsA.Fields(2)
            rsW.Fields("TOKCD") = rsA.Fields(3)
            rsW.Fields("ZKMUZEKN") = rsA.Fields(4)
            rsW.Update
        End If
        rsA.MoveNext
    Loop
    rsA.Close
    
    'ΈήΩ°ΜίΊ°ΔήXV
    strSQL = ""
    strSQL = strSQL & "UPDATE W_KA_URI"
    strSQL = strSQL & "       SET GCODE = TOK.GRPCD,"
    strSQL = strSQL & "           TANCD = TOK.TANCD"
    strSQL = strSQL & "       FROM OPENQUERY ([ORA],"
    strSQL = strSQL & "                       'SELECT GRPCD,"
    strSQL = strSQL & "                               TANCD,"
    strSQL = strSQL & "                               TOKCD"
    strSQL = strSQL & "                        FROM TOKMTA') as TOK"
    strSQL = strSQL & "       INNER JOIN W_KA_URI"
    strSQL = strSQL & "       ON TOK.TOKCD = W_KA_URI.TOKCD"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
    If rsW.EOF = False Then
        rsW.MoveFirst
    End If
    Do Until rsW.EOF
        If Trim(rsW.Fields("GCODE")) = "" Then
            rsW.Fields("GCODE") = rsW.Fields("TOKCD")
        End If
        KBN_NAME = ""
        rsW.Fields("NKBN") = KBN_CHG(rsW.Fields("TOKCD"), rsW.Fields("GCODE"))
        rsW.Fields("NKNM") = KBN_NAME
        rsW.Update
        rsW.MoveNext
    Loop

    end_time = Timer
    Debug.Print "W_KA_URI  " & (end_time - start_time)
    
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

Sub GET_URI_T(strDate As String)
'====================================================================================================================
'γgiγϊtA€iΊ°ΔήAγΰzA΄Ώΰzj
'εΊ°Δή=080808(Oθj
'         080880iΈήΩ°Μίθj
'         080881(n΅)j
'         080885(ΑH)
'`[ζͺ=2(γ)or 3(Ό)AΑοΕ­
'====================================================================================================================
    
    Dim cnA    As New ADODB.Connection
    Dim rsP    As New ADODB.Recordset
    Dim rsA    As New ADODB.Recordset
    Dim rsW    As New ADODB.Recordset
    Dim Cmd    As New ADODB.Command
    Dim strSQL As String
    Dim strCD  As String 'S?R[h
    Dim lngD   As Long   '
    
    start_time = Timer
    
    'μΖΓ°ΜήΩ(SQL Server)
    
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnA.Open
    strSQL = "SELECT * FROM W_TA_URI"
    rsW.Open strSQL, cnA, adOpenStatic, adLockOptimistic
    
    'SQLΜSMADTΙϊΎ―Δ
    strSQL = ""
    strSQL = strSQL & "SELECT * FROM OPENQUERY ([ORA],"
    strSQL = strSQL & "             'SELECT UDNDT"
    strSQL = strSQL & "                    ,Sum(URIKN)"
    strSQL = strSQL & "                    ,Sum(GNKKN)"
    strSQL = strSQL & "                    ,TOKCD"
    strSQL = strSQL & "                    ,HINCD"
    strSQL = strSQL & "                    ,Sum(ZKMUZEKN) "
    strSQL = strSQL & "              FROM UDNTRA"
    strSQL = strSQL & "                   WHERE  DATKB = ''1''"
    strSQL = strSQL & "                     And  TOKBMNCD IN(''080808'',''080880'',''080881'')"
    strSQL = strSQL & "                     And  DENKB IN(''2'',''3'')"
    strSQL = strSQL & "                     And  LINNO < ''990''"
    strSQL = strSQL & "                     And  SMADT = ''" & strDate & "''"
    strSQL = strSQL & "                   GROUP BY UDNDT"
    strSQL = strSQL & "                           ,TOKCD"
    strSQL = strSQL & "                           ,HINCD"
    strSQL = strSQL & "              ')"
    Set Cmd.ActiveConnection = cnA
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    If rsA.EOF = False Then
        rsA.MoveFirst
    End If

    Do Until rsA.EOF
        'ΰz ιͺΎ―
        If rsA.Fields(1) <> 0 Or rsA.Fields(2) <> 0 Or rsA.Fields(5) <> 0 Then
            rsW.AddNew
            rsW.Fields("SMADT") = strDate
            rsW.Fields("UDNDT") = rsA.Fields(0)
            rsW.Fields("URIKIN") = rsA.Fields(1)
            rsW.Fields("GENKIN") = rsA.Fields(2)
            rsW.Fields("TOKCD") = rsA.Fields(3)
            rsW.Fields("HINCD") = rsA.Fields(4)
            rsW.Fields("ZKMUZEKN") = rsA.Fields(5)
            rsW.Update
        End If
        rsA.MoveNext
    Loop
    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    
    'ΈήΩ°ΜίΊ°ΔήXV
    strSQL = ""
    strSQL = strSQL & "UPDATE W_TA_URI"
    strSQL = strSQL & "       SET GCODE = TOK.GRPCD,"
    strSQL = strSQL & "           TANCD = TOK.TANCD"
    strSQL = strSQL & "       FROM OPENQUERY ([ORA],"
    strSQL = strSQL & "                       'SELECT GRPCD,"
    strSQL = strSQL & "                               TANCD,"
    strSQL = strSQL & "                               TOKCD"
    strSQL = strSQL & "                        FROM TOKMTA') as TOK"
    strSQL = strSQL & "       INNER JOIN W_TA_URI"
    strSQL = strSQL & "       ON TOK.TOKCD = W_TA_URI.TOKCD"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
        
    '€iζͺXV
    strSQL = ""
    strSQL = strSQL & "UPDATE W_TA_URI"
    strSQL = strSQL & "       SET HINBID = HIN.HINCLBID,"
    strSQL = strSQL & "           HINCID = HIN.HINCLCID"
    strSQL = strSQL & "       FROM OPENQUERY ([ORA],"
    strSQL = strSQL & "                       'SELECT HINCLBID,"
    strSQL = strSQL & "                               HINCLCID,"
    strSQL = strSQL & "                               HINCD"
    strSQL = strSQL & "                        FROM HINMTA') as HIN"
    strSQL = strSQL & "       INNER JOIN W_TA_URI"
    strSQL = strSQL & "       ON HIN.HINCD = W_TA_URI.HINCD"
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
    If rsW.EOF = False Then
        rsW.MoveFirst
    End If
    Do Until rsW.EOF
        If Trim(rsW.Fields("GCODE")) = "" Then
            rsW.Fields("GCODE") = rsW.Fields("TOKCD")
        End If
        KBN_NAME = ""
        rsW.Fields("NKBN") = KBN_CHGT(rsW.Fields("TOKCD"), rsW.Fields("GCODE"), rsW.Fields("HINBID"), "")
        rsW.Fields("NKNM") = KBN_NAME
        rsW.Update
        rsW.MoveNext
    Loop

    end_time = Timer
    Debug.Print "W_TA_URI " & (end_time - start_time)
    
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
