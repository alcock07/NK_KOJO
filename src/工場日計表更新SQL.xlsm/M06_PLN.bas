Attribute VB_Name = "M06_PLN"
Option Explicit

Sub GET_PLN_K(ByVal strDate As String)
    
    Dim cnW    As ADODB.Connection
    Dim rsW    As ADODB.Recordset
    Dim rsP    As ADODB.Recordset
    Dim Cmd    As New ADODB.Command
    Dim strSQL As String
    Dim strCD  As String '�S���҃R�[�h
    Dim lngD   As Long   '��
    
    start_time = Timer
    
    '���v��ƃe�[�u��
    Set cnW = New ADODB.Connection
    Set rsW = New ADODB.Recordset
    cnW.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnW.Open
    strSQL = "SELECT * FROM W_PLN"
    rsW.Open strSQL, cnW, adOpenStatic, adLockOptimistic
    
    '�v��f�[�^
    Set rsP = New ADODB.Recordset
    strSQL = ""
    strSQL = strSQL & "SELECT ���Ӑ�R�[�h,"
    strSQL = strSQL & "       ���i�敪A,"
    strSQL = strSQL & "       Sum(����01),"
    strSQL = strSQL & "       Sum(����02),"
    strSQL = strSQL & "       Sum(����03),"
    strSQL = strSQL & "       Sum(����04),"
    strSQL = strSQL & "       Sum(����05),"
    strSQL = strSQL & "       Sum(����06),"
    strSQL = strSQL & "       Sum(����07),"
    strSQL = strSQL & "       Sum(����08),"
    strSQL = strSQL & "       Sum(����09),"
    strSQL = strSQL & "       Sum(����10),"
    strSQL = strSQL & "       Sum(����11),"
    strSQL = strSQL & "       Sum(����12),"
    strSQL = strSQL & "       Sum(�e��01),"
    strSQL = strSQL & "       Sum(�e��02),"
    strSQL = strSQL & "       Sum(�e��03),"
    strSQL = strSQL & "       Sum(�e��04),"
    strSQL = strSQL & "       Sum(�e��05),"
    strSQL = strSQL & "       Sum(�e��06),"
    strSQL = strSQL & "       Sum(�e��07),"
    strSQL = strSQL & "       Sum(�e��08),"
    strSQL = strSQL & "       Sum(�e��09),"
    strSQL = strSQL & "       Sum(�e��10),"
    strSQL = strSQL & "       Sum(�e��11),"
    strSQL = strSQL & "       Sum(�e��12)"
    strSQL = strSQL & "       FROM �N�x�v��"
    strSQL = strSQL & "       WHERE �x�X  = '�֓�'"
    strSQL = strSQL & "       GROUP BY ���Ӑ�R�[�h,"
    strSQL = strSQL & "                ���i�敪A"
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
    
    '��ٰ�ߺ��ލX�V
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
    Dim strCD  As String '�S���҃R�[�h
    Dim lngD   As Long   '��
    Dim lngR   As Long
    
    start_time = Timer
    
    '���v��ƃe�[�u��
    Set cnW = New ADODB.Connection
    Set rsW = New ADODB.Recordset
    cnW.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnW.Open
    strSQL = "SELECT * FROM W_PLN"
    rsW.Open strSQL, cnW, adOpenStatic, adLockOptimistic
    
    '�v��f�[�^
    Set rsP = New ADODB.Recordset
    strSQL = ""
    strSQL = strSQL & "SELECT ���Ӑ�R�[�h,"
    strSQL = strSQL & "       ���i�敪A,"
    strSQL = strSQL & "       Sum(����01),"
    strSQL = strSQL & "       Sum(����02),"
    strSQL = strSQL & "       Sum(����03),"
    strSQL = strSQL & "       Sum(����04),"
    strSQL = strSQL & "       Sum(����05),"
    strSQL = strSQL & "       Sum(����06),"
    strSQL = strSQL & "       Sum(����07),"
    strSQL = strSQL & "       Sum(����08),"
    strSQL = strSQL & "       Sum(����09),"
    strSQL = strSQL & "       Sum(����10),"
    strSQL = strSQL & "       Sum(����11),"
    strSQL = strSQL & "       Sum(����12),"
    strSQL = strSQL & "       Sum(�e��01),"
    strSQL = strSQL & "       Sum(�e��02),"
    strSQL = strSQL & "       Sum(�e��03),"
    strSQL = strSQL & "       Sum(�e��04),"
    strSQL = strSQL & "       Sum(�e��05),"
    strSQL = strSQL & "       Sum(�e��06),"
    strSQL = strSQL & "       Sum(�e��07),"
    strSQL = strSQL & "       Sum(�e��08),"
    strSQL = strSQL & "       Sum(�e��09),"
    strSQL = strSQL & "       Sum(�e��10),"
    strSQL = strSQL & "       Sum(�e��11),"
    strSQL = strSQL & "       Sum(�e��12),"
    strSQL = strSQL & "       �S���҃R�[�h"
    strSQL = strSQL & "       FROM �C���v��"
    strSQL = strSQL & "       WHERE �x�X  = '���C'"
    strSQL = strSQL & "       GROUP BY ���Ӑ�R�[�h,"
    strSQL = strSQL & "                ���i�敪A,"
    strSQL = strSQL & "                �S���҃R�[�h"
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
    
    '��ٰ�ߺ��ލX�V
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
    
    '��ď����د�ނ̌v���DB�ɓ����
    For lngR = 2 To 5
        rsW.AddNew
        rsW.Fields("TOKCD") = "0000000819001"               '���Ӑ�R�[�h
        rsW.Fields("GCODE") = "0000000819001"               'G�R�[�h
        rsW.Fields("TANCD") = ""
        rsW.Fields("KBN") = Sheets("Plan").Cells(lngR, 2)   '���i�敪A
        rsW.Fields("NKBN") = Sheets("Plan").Cells(lngR, 2)  '���v�敪
        rsW.Fields("NKNM") = Sheets("Plan").Cells(lngR, 1)  '���v�敪��
        rsW.Fields("PUKN") = Sheets("Plan").Cells(lngR, lngD + 2) * 10000     '����
        rsW.Fields("PAKN") = Sheets("Plan").Cells(lngR + 5, lngD + 2) * 10000 '�e��
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
