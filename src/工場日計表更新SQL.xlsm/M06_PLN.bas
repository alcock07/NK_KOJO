Attribute VB_Name = "M06_PLN"
Option Explicit

Sub GET_PLN_K(ByVal strDate As String)
    
    Dim cnW    As ADODB.Connection
    Dim rsW    As ADODB.Recordset
    Dim rsP    As ADODB.Recordset
    Dim Cmd    As New ADODB.Command
    Dim strSQL As String
    Dim strCD  As String '担当者コード
    Dim lngD   As Long   '月
    
    start_time = Timer
    
    '日計作業テーブル
    Set cnW = New ADODB.Connection
    Set rsW = New ADODB.Recordset
    cnW.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnW.Open
    strSQL = "SELECT * FROM W_PLN"
    rsW.Open strSQL, cnW, adOpenStatic, adLockOptimistic
    
    '計画データ
    Set rsP = New ADODB.Recordset
    strSQL = ""
    strSQL = strSQL & "SELECT 得意先コード,"
    strSQL = strSQL & "       商品区分A,"
    strSQL = strSQL & "       Sum(売上01),"
    strSQL = strSQL & "       Sum(売上02),"
    strSQL = strSQL & "       Sum(売上03),"
    strSQL = strSQL & "       Sum(売上04),"
    strSQL = strSQL & "       Sum(売上05),"
    strSQL = strSQL & "       Sum(売上06),"
    strSQL = strSQL & "       Sum(売上07),"
    strSQL = strSQL & "       Sum(売上08),"
    strSQL = strSQL & "       Sum(売上09),"
    strSQL = strSQL & "       Sum(売上10),"
    strSQL = strSQL & "       Sum(売上11),"
    strSQL = strSQL & "       Sum(売上12),"
    strSQL = strSQL & "       Sum(粗利01),"
    strSQL = strSQL & "       Sum(粗利02),"
    strSQL = strSQL & "       Sum(粗利03),"
    strSQL = strSQL & "       Sum(粗利04),"
    strSQL = strSQL & "       Sum(粗利05),"
    strSQL = strSQL & "       Sum(粗利06),"
    strSQL = strSQL & "       Sum(粗利07),"
    strSQL = strSQL & "       Sum(粗利08),"
    strSQL = strSQL & "       Sum(粗利09),"
    strSQL = strSQL & "       Sum(粗利10),"
    strSQL = strSQL & "       Sum(粗利11),"
    strSQL = strSQL & "       Sum(粗利12)"
    strSQL = strSQL & "       FROM 年度計画"
    strSQL = strSQL & "       WHERE 支店  = '関東'"
    strSQL = strSQL & "       GROUP BY 得意先コード,"
    strSQL = strSQL & "                商品区分A"
    Set Cmd.ActiveConnection = cnW
    Cmd.CommandText = strSQL
    Set rsP = Cmd.Execute
    If rsP.EOF = False Then rsP.MoveFirst
    
    lngD = CLng(Mid(strDate, 5, 2))
    Do Until rsP.EOF
        If rsP.Fields(lngD + 1) = 0 And rsP.Fields(lngD + 13) = 0 Then
        Else
            rsW.AddNew
            rsW.Fields("TOKCD") = rsP.Fields(0)
            rsW.Fields("KBN") = rsP.Fields(1)
            rsW.Fields("PUKN") = rsP.Fields(lngD + 1) * 10000
            rsW.Fields("PAKN") = rsP.Fields(lngD + 13) * 10000
            rsW.Update
        End If
        rsP.MoveNext
    Loop
    rsW.Close
    
    'ｸﾞﾙｰﾌﾟｺｰﾄﾞ更新
    strSQL = ""
    strSQL = strSQL & "Update W_PLN"
    strSQL = strSQL & "       SET GCODE = TOK.GRPCD,"
    strSQL = strSQL & "           TANCD = TOK.TANCD"
    strSQL = strSQL & "       FROM OPENQUERY ([ORA],"
    strSQL = strSQL & "                       'SELECT GRPCD,"
    strSQL = strSQL & "                               TANCD,"
    strSQL = strSQL & "                               TOKCD"
    strSQL = strSQL & "                        FROM TOKMTA ') as TOK"
    strSQL = strSQL & "       INNER JOIN W_PLN"
    strSQL = strSQL & "       ON TOK.TOKCD = W_PLN.TOKCD"
    rsW.Open strSQL, cnW, adOpenStatic, adLockOptimistic
    
    strSQL = "SELECT * FROM W_PLN"
    rsW.Open strSQL, cnW, adOpenStatic, adLockOptimistic
    If rsW.EOF = False Then
        rsW.MoveFirst
    End If
    Do Until rsW.EOF
        If IsNull(rsW.Fields("GCODE")) Then
            rsW.Fields("GCODE") = ""
            rsW.Fields("TANCD") = ""
        End If
        If Trim(rsW.Fields("GCODE")) = "" Then
            rsW.Fields("GCODE") = rsW.Fields("TOKCD")
        End If
        KBN_NAME = ""
        rsW.Fields("NKBN") = KBN_CHG(rsW.Fields("TOKCD"), rsW.Fields("GCODE"))
        rsW.Fields("NKNM") = KBN_NAME
        If Trim(rsW.Fields("TANCD")) = "" Then
            rsW.Fields("TANCD") = "00000708"
        End If
        rsW.Update
        rsW.MoveNext
    Loop
    
    end_time = Timer
    Debug.Print "W_PLN     " & (end_time - start_time)

Exit_DB:

    If Not rsP Is Nothing Then
        If rsP.State = adStateOpen Then rsP.Close
        Set rsP = Nothing
    End If
    If Not rsW Is Nothing Then
        If rsW.State = adStateOpen Then rsW.Close
        Set rsW = Nothing
    End If
    If Not cnW Is Nothing Then
        If cnW.State = adStateOpen Then cnW.Close
        Set cnW = Nothing
    End If
    
End Sub

Sub GET_PLN_T(ByVal strDate As String)
    
    Dim cnW    As ADODB.Connection
    Dim rsW    As ADODB.Recordset
    Dim rsP    As ADODB.Recordset
    Dim Cmd    As New ADODB.Command
    Dim strSQL As String
    Dim strCD  As String '担当者コード
    Dim lngD   As Long   '月
    Dim lngR   As Long
    
    start_time = Timer
    
    '日計作業テーブル
    Set cnW = New ADODB.Connection
    Set rsW = New ADODB.Recordset
    cnW.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnW.Open
    strSQL = "SELECT * FROM W_PLN"
    rsW.Open strSQL, cnW, adOpenStatic, adLockOptimistic
    
    '計画データ
    Set rsP = New ADODB.Recordset
    strSQL = ""
    strSQL = strSQL & "SELECT 得意先コード,"
    strSQL = strSQL & "       商品区分A,"
    strSQL = strSQL & "       Sum(売上01),"
    strSQL = strSQL & "       Sum(売上02),"
    strSQL = strSQL & "       Sum(売上03),"
    strSQL = strSQL & "       Sum(売上04),"
    strSQL = strSQL & "       Sum(売上05),"
    strSQL = strSQL & "       Sum(売上06),"
    strSQL = strSQL & "       Sum(売上07),"
    strSQL = strSQL & "       Sum(売上08),"
    strSQL = strSQL & "       Sum(売上09),"
    strSQL = strSQL & "       Sum(売上10),"
    strSQL = strSQL & "       Sum(売上11),"
    strSQL = strSQL & "       Sum(売上12),"
    strSQL = strSQL & "       Sum(粗利01),"
    strSQL = strSQL & "       Sum(粗利02),"
    strSQL = strSQL & "       Sum(粗利03),"
    strSQL = strSQL & "       Sum(粗利04),"
    strSQL = strSQL & "       Sum(粗利05),"
    strSQL = strSQL & "       Sum(粗利06),"
    strSQL = strSQL & "       Sum(粗利07),"
    strSQL = strSQL & "       Sum(粗利08),"
    strSQL = strSQL & "       Sum(粗利09),"
    strSQL = strSQL & "       Sum(粗利10),"
    strSQL = strSQL & "       Sum(粗利11),"
    strSQL = strSQL & "       Sum(粗利12),"
    strSQL = strSQL & "       担当者コード"
    strSQL = strSQL & "       FROM 修正計画"
    strSQL = strSQL & "       WHERE 支店  = '東海'"
    strSQL = strSQL & "       GROUP BY 得意先コード,"
    strSQL = strSQL & "                商品区分A,"
    strSQL = strSQL & "                担当者コード"
    Set Cmd.ActiveConnection = cnW
    Cmd.CommandText = strSQL
    Set rsP = Cmd.Execute
    If rsP.EOF = False Then rsP.MoveFirst
    
    lngD = CLng(Mid(strDate, 5, 2))
    Do Until rsP.EOF
        If rsP.Fields(lngD + 1) = 0 And rsP.Fields(lngD + 13) = 0 Then
        Else
            rsW.AddNew
            rsW.Fields("TOKCD") = rsP.Fields(0)
            rsW.Fields("KBN") = rsP.Fields(1)
            rsW.Fields("PUKN") = rsP.Fields(lngD + 1) * 10000
            rsW.Fields("PAKN") = rsP.Fields(lngD + 13) * 10000
            rsW.Fields("TANCD") = rsP.Fields(26)
            rsW.Update
        End If
        rsP.MoveNext
    Loop
    rsW.Close
    
    'ｸﾞﾙｰﾌﾟｺｰﾄﾞ更新
    strSQL = ""
    strSQL = strSQL & "Update W_PLN"
    strSQL = strSQL & "       SET GCODE = TOK.GRPCD"
    strSQL = strSQL & "       FROM OPENQUERY ([ORA],"
    strSQL = strSQL & "                       'SELECT GRPCD,"
    strSQL = strSQL & "                               TOKCD"
    strSQL = strSQL & "                        FROM TOKMTA ') as TOK"
    strSQL = strSQL & "       INNER JOIN W_PLN"
    strSQL = strSQL & "       ON TOK.TOKCD = W_PLN.TOKCD"
    rsW.Open strSQL, cnW, adOpenStatic, adLockOptimistic

    strSQL = "SELECT * FROM W_PLN"
    rsW.Open strSQL, cnW, adOpenStatic, adLockOptimistic
    If rsW.EOF = False Then
        rsW.MoveFirst
    End If
    Do Until rsW.EOF
        If IsNull(rsW.Fields("GCODE")) Then
            rsW.Fields("GCODE") = ""
        End If
        If Trim(rsW.Fields("GCODE")) = "" Then
            rsW.Fields("GCODE") = rsW.Fields("TOKCD")
        End If
        rsW.Fields("NKBN") = KBN_CHGT(rsW.Fields("TOKCD"), rsW.Fields("GCODE"), "", rsW.Fields("KBN"))
        rsW.Fields("NKNM") = KBN_NAME
        rsW.Update
        rsW.MoveNext
    Loop
    
    'ｼｰﾄ上のﾌﾞﾘｯｼﾞの計画をDBに入れる
    For lngR = 2 To 5
        rsW.AddNew
        rsW.Fields("TOKCD") = "0000000819001"               '得意先コード
        rsW.Fields("GCODE") = "0000000819001"               'Gコード
        rsW.Fields("TANCD") = ""
        rsW.Fields("KBN") = Sheets("Plan").Cells(lngR, 2)   '商品区分A
        rsW.Fields("NKBN") = Sheets("Plan").Cells(lngR, 2)  '日計区分
        rsW.Fields("NKNM") = Sheets("Plan").Cells(lngR, 1)  '日計区分名
        rsW.Fields("PUKN") = Sheets("Plan").Cells(lngR, lngD + 2) * 10000     '売上
        rsW.Fields("PAKN") = Sheets("Plan").Cells(lngR + 5, lngD + 2) * 10000 '粗利
        rsW.Update
    Next lngR
    rsW.Close
    
    end_time = Timer
    Debug.Print "W_PLN    " & (end_time - start_time)

Exit_DB:

    If Not rsP Is Nothing Then
        If rsP.State = adStateOpen Then rsP.Close
        Set rsP = Nothing
    End If
    If Not rsW Is Nothing Then
        If rsW.State = adStateOpen Then rsW.Close
        Set rsW = Nothing
    End If
    If Not cnW Is Nothing Then
        If cnW.State = adStateOpen Then cnW.Close
        Set cnW = Nothing
    End If
    
End Sub
