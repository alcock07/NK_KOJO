Attribute VB_Name = "M10_UP_NK"
Option Explicit

Sub UP_DATA_K(strDate As String)

    Dim cnA    As New ADODB.Connection
    Dim rsA    As New ADODB.Recordset
    Dim Cmd    As New ADODB.Command
    Dim strSQL As String

    start_time = Timer
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnA.Open
    
    '日計ﾃﾞｰﾀ当月分削除
    strSQL = ""
    strSQL = strSQL & "DELETE"
    strSQL = strSQL & "       FROM NK_KJR"
    strSQL = strSQL & "       WHERE SMADT = '" & strDate & "'"
    strSQL = strSQL & "       AND   FKBN = 'KANTO'"
    Set Cmd.ActiveConnection = cnA
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
    'ﾃﾞｰﾀ更新
    strSQL = ""
    strSQL = strSQL & "INSERT INTO"
    strSQL = strSQL & "       NK_KJR(SMADT"
    strSQL = strSQL & "             ,FKBN"
    strSQL = strSQL & "             ,GCODE"
    strSQL = strSQL & "             ,NKBN"
    strSQL = strSQL & "             ,NKNM"
    strSQL = strSQL & "             ,UDNDT"
    strSQL = strSQL & "             ,URIKN"
    strSQL = strSQL & "             ,URIKNR"
    strSQL = strSQL & "             ,GENKN"
    strSQL = strSQL & "             ,GENKNR"
    strSQL = strSQL & "             ,JUZ"
    strSQL = strSQL & "             ,SIRKN"
    strSQL = strSQL & "             ,PUKN"
    strSQL = strSQL & "             ,PAKN"
    strSQL = strSQL & "              )"
    strSQL = strSQL & "        SELECT SMADT"
    strSQL = strSQL & "             ,'KANTO'"
    strSQL = strSQL & "             ,GCODE"
    strSQL = strSQL & "             ,NKBN"
    strSQL = strSQL & "             ,NKNM"
    strSQL = strSQL & "             ,UDNDT"
    strSQL = strSQL & "             ,URIKN"
    strSQL = strSQL & "             ,URIKNR"
    strSQL = strSQL & "             ,GENKN"
    strSQL = strSQL & "             ,GENKNR"
    strSQL = strSQL & "             ,JUZ"
    strSQL = strSQL & "             ,SIRKN"
    strSQL = strSQL & "             ,PUKN"
    strSQL = strSQL & "             ,PAKN"
    strSQL = strSQL & "       FROM W_KA_NK"
    Set Cmd.ActiveConnection = cnA
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
    Debug.Print "NK_KJ     " & (end_time - start_time)
    
End Sub

Sub UP_DATA_T(strDate As String)

    Dim cnA    As New ADODB.Connection
    Dim rsA    As New ADODB.Recordset
    Dim Cmd    As New ADODB.Command
    Dim strSQL As String

    start_time = Timer
    cnA.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnA.Open
    
    '日計ﾃﾞｰﾀ当月分削除
    strSQL = ""
    strSQL = strSQL & "DELETE"
    strSQL = strSQL & "       FROM NK_TAR"
    strSQL = strSQL & "       WHERE SMADT = '" & strDate & "'"
    Set Cmd.ActiveConnection = cnA
    Cmd.CommandText = strSQL
    Set rsA = Cmd.Execute
    
    'ﾃﾞｰﾀ更新
    strSQL = ""
    strSQL = strSQL & "INSERT INTO"
    strSQL = strSQL & "       NK_TAR(SMADT"
    strSQL = strSQL & "             ,NKBN"
    strSQL = strSQL & "             ,NKNM"
    strSQL = strSQL & "             ,UDNDT"
    strSQL = strSQL & "             ,URIKN"
    strSQL = strSQL & "             ,URIKNR"
    strSQL = strSQL & "             ,GENKN"
    strSQL = strSQL & "             ,GENKNR"
    strSQL = strSQL & "             ,JUZ"
    strSQL = strSQL & "             ,SIRKN"
    strSQL = strSQL & "             ,PUKN"
    strSQL = strSQL & "             ,PAKN"
    strSQL = strSQL & "              )"
    strSQL = strSQL & "       SELECT " & strDate
    strSQL = strSQL & "             ,NKBN"
    strSQL = strSQL & "             ,NKNM"
    strSQL = strSQL & "             ,UDNDT"
    strSQL = strSQL & "             ,URIKN"
    strSQL = strSQL & "             ,URIKNR"
    strSQL = strSQL & "             ,GENKN"
    strSQL = strSQL & "             ,GENKNR"
    strSQL = strSQL & "             ,JUZ"
    strSQL = strSQL & "             ,SIRKN"
    strSQL = strSQL & "             ,PUKN"
    strSQL = strSQL & "             ,PAKN"
    strSQL = strSQL & "       FROM W_TA_NK"
    Set Cmd.ActiveConnection = cnA
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
    Debug.Print "NK_TAR   " & (end_time - start_time)
    
End Sub
