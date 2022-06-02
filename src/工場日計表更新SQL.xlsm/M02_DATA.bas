Attribute VB_Name = "M02_DATA"
Option Explicit

Public Const MYPROVIDERE = "Provider=SQLOLEDB;"
Public Const MYSERVER = "Data Source=192.168.128.9\SQLEXPRESS;"
Public Const USER = "User ID=sa;"
Public Const PSWD = "Password=ALCadmin!;"
Public Const strNT = "Initial Catalog=process_os;"

Sub Proc_TZ()

    Dim lngYYY  As Long    '�N
    Dim lngMMM  As Long    '��
    Dim strMon  As String  '�������t
    Dim DateA   As Date    '���t��Ɨp
    
    '�O�����擾
    lngYYY = CInt(Format(Now(), "yyyy"))
    lngMMM = CInt(Format(Now(), "mm"))
    DateA = CDate(lngYYY & "/" & lngMMM & "/01")
    DateA = DateA - 1
    strMon = Format(DateA, "yyyymmdd")
    
    
    Sheets("Wait").Range("D15") = "�O���f�[�^�X�V���E�E�E"
    
    Sheets("Wait").Range("D12") = "���֓��A���R�b�N��"
    DoEvents
'    strMon = "20210630"
    Call Proc_DataK(strMon)
    
    Sheets("Wait").Range("D12") = "�����C�A���R�b�N��"
    DoEvents
    Call Proc_DataT(strMon)
    
    '�����擾
    lngYYY = CLng(Format(Now(), "yyyy"))
    lngMMM = CLng(Format(Now(), "mm"))
    lngMMM = lngMMM + 1
    If lngMMM = 13 Then
        lngMMM = 1
        lngYYY = lngYYY + 1
    End If
    DateA = CDate(lngYYY & "/" & lngMMM & "/01")
    DateA = DateA - 1
    strMon = Format(DateA, "yyyymmdd")

    Sheets("Wait").Range("D15") = "�����f�[�^�X�V���E�E�E"
    Sheets("Wait").Range("D12") = "���֓��A���R�b�N��"
    DoEvents
'    strMon = "20210731"
    Call Proc_DataK(strMon)
    
    Sheets("Wait").Range("D12") = "�����C�A���R�b�N��"
    DoEvents
    Call Proc_DataT(strMon)
    
    Sheets("Wait").Range("D15") = "�I���I�I"
    DoEvents
    
End Sub

Sub Proc_DataK(strMon As String)

    '��ƃe�[�u���쐬
    Call CR_TBL_URI
    Call CR_TBL_JUZ
    Call CR_TBL_SRE
    Call CR_TBL_PLN
    Call CR_TBL_NK
    Call CR_TBL_NKT

    '�f�[�^�擾
    Call GET_URI_K(strMon)  '������݂���֓��̔��ゾ����W_KA_URI�֒��o
    Call GET_JUC_K(strMon)  '����݂���֓��̎󒍂�����W_KA_JUZ�֒��o
    Call GET_SIR_K(strMon)  '�d����݂���֓��̎d��������W_KA_SRE�֒��o
    Call GET_PLN_K(strMon)  '�c�ƌv�悩��֓��̌v�悾����W_KA_PLN�֒��o
    Call GET_TOK_K(strMon)  'W_KA_NK�e�[�u���ɔ���A�󒍁A�d���A�v�������
    Call GET_TAN_K(strMon)  'W_KA_NKT�e�[�u���ɔ���A�󒍁A�d���A�v�������
    Call UP_DATA_K(strMon)  'W_KA_NK�e�[�u������NK_KJR��
    Call UP_TAN_K(strMon)   'W_KA_NKT�e�[�u������NK_KJT��

    '��ƃe�[�u���폜
    Call DR_TBL_URI
    Call DR_TBL_JUZ
    Call DR_TBL_SRE
    Call DR_TBL_PLN
    Call DR_TBL_NK
    Call DR_TBL_NKT


End Sub

Sub Proc_DataT(strMon As String)

    '��ƃe�[�u���쐬
    Call CR_TBL_TAU
    Call CR_TBL_TAJ
    Call CR_TBL_TAS
    Call CR_TBL_SRE
    Call CR_TBL_PLN
    Call CR_TBL_NKTA
    Call CR_TBL_NKTT

    '�f�[�^�擾
    Call GET_URI_T(strMon)  '������݂��瓌�C�̔��ゾ����W_TA_URI�֒��o
    Call GET_JUC_T(strMon)  '����݂��瓌�C�̎󒍂�����W_TA_JUZ�֒��o
    Call GET_SIR_T(strMon)  '�d����݂��瓌�C�̎d��������W_TA_SRE�֒��o
    Call GET_SIR_T2(strMon) '�d����݂��瓌�C�̉��H�󂯂�����W_KA_SRE�֒��o
    Call GET_PLN_T(strMon)  '�c�ƌv�悩�瓌�C�̌v�悾����W_TA_PLN�֒��o
    Call GET_TOK_T(strMon)  'W_TA_NK�e�[�u���ɔ���A�󒍁A�d���A�v�������
    Call GET_TAN_T(strMon)  'W_TA_NKT�e�[�u���ɔ���A�󒍁A�d���A�v�������
    Call UP_DATA_T(strMon)  'W_TA_NK�e�[�u������NK_KJR��
    Call UP_TAN_T(strMon)   'W_TA_NKT�e�[�u������NK_KJT��

    '��ƃe�[�u���폜
    Call DR_TBL_TAU
    Call DR_TBL_TAJ
    Call DR_TBL_TAS
    Call DR_TBL_SRE
    Call DR_TBL_PLN
    Call DR_TBL_NKTA
    Call DR_TBL_NKTT


End Sub

Sub CR_TBL_URI()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strSQL As String

    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'W_KA_URI�e�[�u���폜
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_KA_URI]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_KA_URI]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'W_KA_URI�e�[�u���쐬
    strSQL = ""
    strSQL = strSQL & "CREATE TABLE [dbo].[W_KA_URI]( "
    strSQL = strSQL & "    [SMADT]     [nchar](8)  NOT NULL, "
    strSQL = strSQL & "    [UDNDT]     [nchar](8)  NOT NULL,"
    strSQL = strSQL & "    [TOKCD]     [nchar](13) NOT NULL, "
    strSQL = strSQL & "    [GCODE]     [nchar](13) NULL, "
    strSQL = strSQL & "    [TANCD]     [nchar](8)  NULL, "
    strSQL = strSQL & "    [NKBN]      [nchar](3)  NULL, "
    strSQL = strSQL & "    [NKNM]      [nchar](20) NULL, "
    strSQL = strSQL & "    [URIKIN]    [real]      NULL, "
    strSQL = strSQL & "    [GENKIN]    [real]      NULL, "
    strSQL = strSQL & "    [ZKMUZEKN]  [real]      NULL, "
    strSQL = strSQL & "CONSTRAINT [PK_KURI] PRIMARY KEY CLUSTERED "
    strSQL = strSQL & "( "
    strSQL = strSQL & "[UDNDT] ASC, "
    strSQL = strSQL & "[TOKCD] ASC "
    strSQL = strSQL & ") WITH "
    strSQL = strSQL & "(PAD_INDEX = OFF, "
    strSQL = strSQL & " STATISTICS_NORECOMPUTE = OFF, "
    strSQL = strSQL & " IGNORE_DUP_KEY = OFF, "
    strSQL = strSQL & " ALLOW_ROW_LOCKS = ON, "
    strSQL = strSQL & " ALLOW_PAGE_LOCKS = ON, "
    strSQL = strSQL & " OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF "
    strSQL = strSQL & ") ON [PRIMARY]"
    strSQL = strSQL & ") ON [PRIMARY]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

Sub DR_TBL_URI()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strSQL As String

    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'W_KA_URI�e�[�u���폜
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_KA_URI]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_KA_URI]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

Sub CR_TBL_JUZ()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strSQL As String

    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'W_KA_JUZ�e�[�u���폜
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_KA_JUZ]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_KA_JUZ]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'W_KA_JUZ�e�[�u���쐬
    strSQL = ""
    strSQL = strSQL & "CREATE TABLE [dbo].[W_KA_JUZ]( "
    strSQL = strSQL & "    [TANCD]     [nchar](8)  NULL, "
    strSQL = strSQL & "    [GCODE]     [nchar](13) NULL, "
    strSQL = strSQL & "    [TOKCD]     [nchar](13) NOT NULL, "
    strSQL = strSQL & "    [TOKNM]     [nchar](40) NULL, "
    strSQL = strSQL & "    [NOKDT]     [nchar](20) NOT NULL, "
    strSQL = strSQL & "    [ZANKN]     [real]      NULL, "
    strSQL = strSQL & "    [GNKKN]     [real]      NULL, "
    strSQL = strSQL & "    [NKBN]      [nchar](3)  NULL, "
    strSQL = strSQL & "    [NKNM]      [nchar](20) NULL, "
    strSQL = strSQL & "CONSTRAINT [PK_KJUZ] PRIMARY KEY CLUSTERED "
    strSQL = strSQL & "( "
    strSQL = strSQL & "[TOKCD] ASC, "
    strSQL = strSQL & "[NOKDT] ASC "
    strSQL = strSQL & ") WITH "
    strSQL = strSQL & "(PAD_INDEX = OFF, "
    strSQL = strSQL & " STATISTICS_NORECOMPUTE = OFF, "
    strSQL = strSQL & " IGNORE_DUP_KEY = OFF, "
    strSQL = strSQL & " ALLOW_ROW_LOCKS = ON, "
    strSQL = strSQL & " ALLOW_PAGE_LOCKS = ON, "
    strSQL = strSQL & " OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF "
    strSQL = strSQL & ") ON [PRIMARY]"
    strSQL = strSQL & ") ON [PRIMARY]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

Sub DR_TBL_JUZ()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strSQL As String

    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'W_KA_JUZ�e�[�u���폜
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_KA_JUZ]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_KA_JUZ]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

Sub CR_TBL_SRE()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strSQL As String

    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'W_KA_SRE�e�[�u���폜
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_KA_SRE]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_KA_SRE]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'W_KA_SRE�e�[�u���쐬
    strSQL = ""
    strSQL = strSQL & "CREATE TABLE [dbo].[W_KA_SRE]( "
    strSQL = strSQL & "    [SMADT]     [nchar](8)  NULL, "
    strSQL = strSQL & "    [HINCD]     [nchar](10) NOT NULL, "
    strSQL = strSQL & "    [HINKB]     [nchar](10) NULL, "
    strSQL = strSQL & "    [SIRKIN]    [real]      NULL, "
    strSQL = strSQL & "    [SIRCD]     [nchar](13) NOT NULL, "
    strSQL = strSQL & "    [GKBN]      [nchar](1)  NULL, "
    strSQL = strSQL & "CONSTRAINT [PK_KSRE] PRIMARY KEY CLUSTERED "
    strSQL = strSQL & "( "
    strSQL = strSQL & "[HINCD] ASC, "
    strSQL = strSQL & "[SIRCD] ASC "
    strSQL = strSQL & ") WITH "
    strSQL = strSQL & "(PAD_INDEX = OFF, "
    strSQL = strSQL & " STATISTICS_NORECOMPUTE = OFF, "
    strSQL = strSQL & " IGNORE_DUP_KEY = OFF, "
    strSQL = strSQL & " ALLOW_ROW_LOCKS = ON, "
    strSQL = strSQL & " ALLOW_PAGE_LOCKS = ON, "
    strSQL = strSQL & " OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF "
    strSQL = strSQL & ") ON [PRIMARY]"
    strSQL = strSQL & ") ON [PRIMARY]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

Sub DR_TBL_SRE()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strSQL As String

    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'W_KA_SRE�e�[�u���폜
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_KA_SRE]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_KA_SRE]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

Sub CR_TBL_PLN()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strSQL As String

    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'W_PLN�e�[�u���폜
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_PLN]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_PLN]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'W_PLN�e�[�u���쐬
    strSQL = ""
    strSQL = strSQL & "CREATE TABLE [dbo].[W_PLN]( "
    strSQL = strSQL & "    [TOKCD]     [nchar](13) NOT NULL, "
    strSQL = strSQL & "    [KBN]       [nchar](10) NOT NULL, "
    strSQL = strSQL & "    [GCODE]     [nchar](13) NULL, "
    strSQL = strSQL & "    [NKBN]      [nchar](3)  NULL, "
    strSQL = strSQL & "    [NKNM]      [nchar](20) NULL, "
    strSQL = strSQL & "    [PUKN]      [real]      NULL, "
    strSQL = strSQL & "    [PAKN]      [real]      NULL, "
    strSQL = strSQL & "    [TANCD]     [nchar](8)  NULL, "
    strSQL = strSQL & "CONSTRAINT [PK_PLN] PRIMARY KEY CLUSTERED "
    strSQL = strSQL & "( "
    strSQL = strSQL & "[TOKCD] ASC, "
    strSQL = strSQL & "[KBN] ASC "
    strSQL = strSQL & ") WITH "
    strSQL = strSQL & "(PAD_INDEX = OFF, "
    strSQL = strSQL & " STATISTICS_NORECOMPUTE = OFF, "
    strSQL = strSQL & " IGNORE_DUP_KEY = OFF, "
    strSQL = strSQL & " ALLOW_ROW_LOCKS = ON, "
    strSQL = strSQL & " ALLOW_PAGE_LOCKS = ON, "
    strSQL = strSQL & " OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF "
    strSQL = strSQL & ") ON [PRIMARY]"
    strSQL = strSQL & ") ON [PRIMARY]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

Sub DR_TBL_PLN()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strSQL As String

    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'W_PLN�e�[�u���폜
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_PLN]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_PLN]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

Sub CR_TBL_NK()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strSQL As String

    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'W_KA_NK�e�[�u���폜
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_KA_NK]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_KA_NK]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'W_KA_NK�e�[�u���쐬
    strSQL = ""
    strSQL = strSQL & "CREATE TABLE [dbo].[W_KA_NK]( "
    strSQL = strSQL & "    [SMADT]     [nchar](8)  NOT NULL, "
    strSQL = strSQL & "    [GCODE]     [nchar](13) NOT NULL, "
    strSQL = strSQL & "    [NKBN]      [nchar](10) NOT NULL, "
    strSQL = strSQL & "    [NKNM]      [nchar](20) NULL, "
    strSQL = strSQL & "    [UDNDT]     [nchar](8)  NULL DEFAULT '', "
    strSQL = strSQL & "    [URIKN]     [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [URIKNR]    [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [GENKN]     [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [GENKNR]    [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [JUZ]       [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [SIRKN]     [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [PUKN]      [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [PAKN]      [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "CONSTRAINT [PK_KANK] PRIMARY KEY CLUSTERED "
    strSQL = strSQL & "( "
    strSQL = strSQL & "[GCODE] ASC, "
    strSQL = strSQL & "[NKBN] ASC "
    strSQL = strSQL & ") WITH "
    strSQL = strSQL & "(PAD_INDEX = OFF, "
    strSQL = strSQL & " STATISTICS_NORECOMPUTE = OFF, "
    strSQL = strSQL & " IGNORE_DUP_KEY = OFF, "
    strSQL = strSQL & " ALLOW_ROW_LOCKS = ON, "
    strSQL = strSQL & " ALLOW_PAGE_LOCKS = ON, "
    strSQL = strSQL & " OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF "
    strSQL = strSQL & ") ON [PRIMARY]"
    strSQL = strSQL & ") ON [PRIMARY]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

Sub DR_TBL_NK()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strSQL As String

    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'W_KA_NK�e�[�u���폜
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_KA_NK]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_KA_NK]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

Sub CR_TBL_NKT()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strSQL As String

    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'W_KA_NKT�e�[�u���폜
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_KA_NKT]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_KA_NKT]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'W_KA_NKT�e�[�u���쐬
    strSQL = ""
    strSQL = strSQL & "CREATE TABLE [dbo].[W_KA_NKT]( "
    strSQL = strSQL & "    [SMADT]     [nchar](8)  NOT NULL, "
    strSQL = strSQL & "    [TANCD]     [nchar](8)  NOT NULL, "
    strSQL = strSQL & "    [TANNM]     [nchar](20) NULL, "
    strSQL = strSQL & "    [UDNDT]     [nchar](8)  NULL DEFAULT '', "
    strSQL = strSQL & "    [URIKN]     [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [URIKNR]    [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [GENKN]     [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [GENKNR]    [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [JUZ]       [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [PUKN]      [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [PAKN]      [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "CONSTRAINT [PK_KANKT] PRIMARY KEY CLUSTERED "
    strSQL = strSQL & "( "
    strSQL = strSQL & "[TANCD] ASC "
    strSQL = strSQL & ") WITH "
    strSQL = strSQL & "(PAD_INDEX = OFF, "
    strSQL = strSQL & " STATISTICS_NORECOMPUTE = OFF, "
    strSQL = strSQL & " IGNORE_DUP_KEY = OFF, "
    strSQL = strSQL & " ALLOW_ROW_LOCKS = ON, "
    strSQL = strSQL & " ALLOW_PAGE_LOCKS = ON, "
    strSQL = strSQL & " OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF "
    strSQL = strSQL & ") ON [PRIMARY]"
    strSQL = strSQL & ") ON [PRIMARY]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

Sub DR_TBL_NKT()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strSQL As String

    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'W_KA_NKT�e�[�u���폜
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_KA_NKT]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_KA_NKT]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

Sub CR_TBL_KJ()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strSQL As String

    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'NK_KJR�e�[�u���폜
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[NK_KJR]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[NK_KJR]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'NK_KJR�e�[�u���쐬
    strSQL = ""
    strSQL = strSQL & "CREATE TABLE [dbo].[NK_KJR]( "
    strSQL = strSQL & "    [SMADT]     [nchar](8)  NOT NULL, "
    strSQL = strSQL & "    [FKBN]      [nchar](5)  NOT NULL, "
    strSQL = strSQL & "    [GCODE]     [nchar](13) NOT NULL, "
    strSQL = strSQL & "    [NKBN]      [nchar](10) NOT NULL, "
    strSQL = strSQL & "    [NKNM]      [nchar](20) NULL, "
    strSQL = strSQL & "    [UDNDT]     [nchar](8)  NULL DEFAULT '', "
    strSQL = strSQL & "    [URIKN]     [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [URIKNR]    [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [GENKN]     [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [GENKNR]    [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [JUZ]       [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [SIRKN]     [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [PUKN]      [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [PAKN]      [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "CONSTRAINT [PK_KKR] PRIMARY KEY CLUSTERED "
    strSQL = strSQL & "( "
    strSQL = strSQL & "[SMADT] ASC, "
    strSQL = strSQL & "[FKBN] ASC, "
    strSQL = strSQL & "[GCODE] ASC, "
    strSQL = strSQL & "[NKBN] ASC "
    strSQL = strSQL & ") WITH "
    strSQL = strSQL & "(PAD_INDEX = OFF, "
    strSQL = strSQL & " STATISTICS_NORECOMPUTE = OFF, "
    strSQL = strSQL & " IGNORE_DUP_KEY = OFF, "
    strSQL = strSQL & " ALLOW_ROW_LOCKS = ON, "
    strSQL = strSQL & " ALLOW_PAGE_LOCKS = ON, "
    strSQL = strSQL & " OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF "
    strSQL = strSQL & ") ON [PRIMARY]"
    strSQL = strSQL & ") ON [PRIMARY]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

Sub CR_TBL_KJT()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strSQL As String

    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'NK_KJT�e�[�u���폜
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[NK_KJT]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[NK_KJT]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'NK_KJT�e�[�u���쐬
    strSQL = ""
    strSQL = strSQL & "CREATE TABLE [dbo].[NK_KJT]( "
    strSQL = strSQL & "    [SMADT]     [nchar](8)  NOT NULL, "
    strSQL = strSQL & "    [FKBN]      [nchar](5)  NOT NULL, "
    strSQL = strSQL & "    [TANCD]     [nchar](8)  NOT NULL, "
    strSQL = strSQL & "    [TANNM]     [nchar](20) NULL, "
    strSQL = strSQL & "    [UDNDT]     [nchar](8)  NULL DEFAULT '', "
    strSQL = strSQL & "    [URIKN]     [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [URIKNR]    [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [GENKN]     [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [GENKNR]    [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [JUZ]       [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [SIRKN]     [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [PUKN]      [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [PAKN]      [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "CONSTRAINT [PK_KKJT] PRIMARY KEY CLUSTERED "
    strSQL = strSQL & "( "
    strSQL = strSQL & "[SMADT] ASC, "
    strSQL = strSQL & "[FKBN] ASC, "
    strSQL = strSQL & "[TANCD] ASC "
    strSQL = strSQL & ") WITH "
    strSQL = strSQL & "(PAD_INDEX = OFF, "
    strSQL = strSQL & " STATISTICS_NORECOMPUTE = OFF, "
    strSQL = strSQL & " IGNORE_DUP_KEY = OFF, "
    strSQL = strSQL & " ALLOW_ROW_LOCKS = ON, "
    strSQL = strSQL & " ALLOW_PAGE_LOCKS = ON, "
    strSQL = strSQL & " OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF "
    strSQL = strSQL & ") ON [PRIMARY]"
    strSQL = strSQL & ") ON [PRIMARY]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

Sub CR_TBL_TAU()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strSQL As String

    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'W_TA_URI�e�[�u���폜
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_TA_URI]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_TA_URI]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'W_TA_URI�e�[�u���쐬
    strSQL = ""
    strSQL = strSQL & "CREATE TABLE [dbo].[W_TA_URI]( "
    strSQL = strSQL & "    [SMADT]     [nchar](8)  NOT NULL, "
    strSQL = strSQL & "    [UDNDT]     [nchar](8)  NOT NULL,"
    strSQL = strSQL & "    [TOKCD]     [nchar](13) NOT NULL, "
    strSQL = strSQL & "    [GCODE]     [nchar](13) NULL, "
    strSQL = strSQL & "    [HINCD]     [nchar](20) NOT NULL, "
    strSQL = strSQL & "    [HINBID]    [nchar](10) NULL, "
    strSQL = strSQL & "    [HINCID]    [nchar](10) NULL, "
    strSQL = strSQL & "    [TANCD]     [nchar](8)  NULL, "
    strSQL = strSQL & "    [NKBN]      [nchar](3)  NULL, "
    strSQL = strSQL & "    [NKNM]      [nchar](20) NULL, "
    strSQL = strSQL & "    [URIKIN]    [real]      NULL, "
    strSQL = strSQL & "    [GENKIN]    [real]      NULL, "
    strSQL = strSQL & "    [ZKMUZEKN]  [real]      NULL, "
    strSQL = strSQL & "CONSTRAINT [PK_TURI] PRIMARY KEY CLUSTERED "
    strSQL = strSQL & "( "
    strSQL = strSQL & "[UDNDT] ASC, "
    strSQL = strSQL & "[TOKCD] ASC, "
    strSQL = strSQL & "[HINCD] ASC "
    strSQL = strSQL & ") WITH "
    strSQL = strSQL & "(PAD_INDEX = OFF, "
    strSQL = strSQL & " STATISTICS_NORECOMPUTE = OFF, "
    strSQL = strSQL & " IGNORE_DUP_KEY = OFF, "
    strSQL = strSQL & " ALLOW_ROW_LOCKS = ON, "
    strSQL = strSQL & " ALLOW_PAGE_LOCKS = ON, "
    strSQL = strSQL & " OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF "
    strSQL = strSQL & ") ON [PRIMARY]"
    strSQL = strSQL & ") ON [PRIMARY]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

Sub DR_TBL_TAU()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strSQL As String

    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'W_TA_URI�e�[�u���폜
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_TA_URI]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_TA_URI]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

Sub CR_TBL_TAJ()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strSQL As String

    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'W_TA_JUZ�e�[�u���폜
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_TA_JUZ]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_TA_JUZ]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'W_TA_JUZ�e�[�u���쐬
    strSQL = ""
    strSQL = strSQL & "CREATE TABLE [dbo].[W_TA_JUZ]( "
    strSQL = strSQL & "    [TANCD]     [nchar](8)  NULL, "
    strSQL = strSQL & "    [GCODE]     [nchar](13) NULL, "
    strSQL = strSQL & "    [HINCD]     [nchar](20) NOT NULL, "
    strSQL = strSQL & "    [HINBID]    [nchar](10) NULL, "
    strSQL = strSQL & "    [HINCID]    [nchar](10) NULL, "
    strSQL = strSQL & "    [TOKCD]     [nchar](13) NOT NULL, "
    strSQL = strSQL & "    [NOKDT]     [nchar](20) NOT NULL, "
    strSQL = strSQL & "    [ZANKN]     [real]      NULL, "
    strSQL = strSQL & "    [GNKKN]     [real]      NULL, "
    strSQL = strSQL & "    [NKBN]      [nchar](3)  NULL, "
    strSQL = strSQL & "    [NKNM]      [nchar](20) NULL, "
    strSQL = strSQL & "CONSTRAINT [PK_TJUZ] PRIMARY KEY CLUSTERED "
    strSQL = strSQL & "( "
    strSQL = strSQL & "[TOKCD] ASC, "
    strSQL = strSQL & "[HINCD] ASC, "
    strSQL = strSQL & "[NOKDT] ASC "
    strSQL = strSQL & ") WITH "
    strSQL = strSQL & "(PAD_INDEX = OFF, "
    strSQL = strSQL & " STATISTICS_NORECOMPUTE = OFF, "
    strSQL = strSQL & " IGNORE_DUP_KEY = OFF, "
    strSQL = strSQL & " ALLOW_ROW_LOCKS = ON, "
    strSQL = strSQL & " ALLOW_PAGE_LOCKS = ON, "
    strSQL = strSQL & " OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF "
    strSQL = strSQL & ") ON [PRIMARY]"
    strSQL = strSQL & ") ON [PRIMARY]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

Sub DR_TBL_TAJ()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strSQL As String

    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'W_TA_JUZ�e�[�u���폜
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_TA_JUZ]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_TA_JUZ]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

Sub CR_TBL_TAS()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strSQL As String

    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'W_TA_SRE�e�[�u���폜
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_TA_SRE]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_TA_SRE]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'W_TA_SRE�e�[�u���쐬
    strSQL = ""
    strSQL = strSQL & "CREATE TABLE [dbo].[W_TA_SRE]( "
    strSQL = strSQL & "    [SMADT]     [nchar](8)  NULL, "
    strSQL = strSQL & "    [SIRKIN]    [real]      NULL, "
    strSQL = strSQL & "    [SIRCD]     [nchar](13) NOT NULL, "
    strSQL = strSQL & "    [GKBN]      [nchar](1)  NULL, "
    strSQL = strSQL & "CONSTRAINT [PK_TSRE] PRIMARY KEY CLUSTERED "
    strSQL = strSQL & "( "
    strSQL = strSQL & "[SIRCD] ASC "
    strSQL = strSQL & ") WITH "
    strSQL = strSQL & "(PAD_INDEX = OFF, "
    strSQL = strSQL & " STATISTICS_NORECOMPUTE = OFF, "
    strSQL = strSQL & " IGNORE_DUP_KEY = OFF, "
    strSQL = strSQL & " ALLOW_ROW_LOCKS = ON, "
    strSQL = strSQL & " ALLOW_PAGE_LOCKS = ON, "
    strSQL = strSQL & " OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF "
    strSQL = strSQL & ") ON [PRIMARY]"
    strSQL = strSQL & ") ON [PRIMARY]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

Sub DR_TBL_TAS()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strSQL As String

    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'W_TA_SRE�e�[�u���폜
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_TA_SRE]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_TA_SRE]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

Sub CR_TBL_NKTA()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strSQL As String

    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'W_TA_NK�e�[�u���폜
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_TA_NK]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_TA_NK]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'W_KA_NK�e�[�u���쐬
    strSQL = ""
    strSQL = strSQL & "CREATE TABLE [dbo].[W_TA_NK]( "
    strSQL = strSQL & "    [SMADT]     [nchar](8)  NOT NULL, "
    strSQL = strSQL & "    [NKBN]      [nchar](10) NOT NULL, "
    strSQL = strSQL & "    [NKNM]      [nchar](20) NULL, "
    strSQL = strSQL & "    [UDNDT]     [nchar](8)  NULL DEFAULT '', "
    strSQL = strSQL & "    [URIKN]     [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [URIKNR]    [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [GENKN]     [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [GENKNR]    [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [JUZ]       [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [SIRKN]     [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [PUKN]      [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [PAKN]      [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "CONSTRAINT [PK_TANK] PRIMARY KEY CLUSTERED "
    strSQL = strSQL & "( "
    strSQL = strSQL & "[NKBN] ASC "
    strSQL = strSQL & ") WITH "
    strSQL = strSQL & "(PAD_INDEX = OFF, "
    strSQL = strSQL & " STATISTICS_NORECOMPUTE = OFF, "
    strSQL = strSQL & " IGNORE_DUP_KEY = OFF, "
    strSQL = strSQL & " ALLOW_ROW_LOCKS = ON, "
    strSQL = strSQL & " ALLOW_PAGE_LOCKS = ON, "
    strSQL = strSQL & " OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF "
    strSQL = strSQL & ") ON [PRIMARY]"
    strSQL = strSQL & ") ON [PRIMARY]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

Sub DR_TBL_NKTA()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strSQL As String

    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'W_TA_NK�e�[�u���폜
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_TA_NK]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_TA_NK]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

Sub CR_TBL_NKTT()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strSQL As String

    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'W_TA_NKT�e�[�u���폜
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_TA_NKT]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_TA_NKT]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'W_TA_NKT�e�[�u���쐬
    strSQL = ""
    strSQL = strSQL & "CREATE TABLE [dbo].[W_TA_NKT]( "
    strSQL = strSQL & "    [SMADT]     [nchar](8)  NOT NULL, "
    strSQL = strSQL & "    [TANCD]     [nchar](8)  NOT NULL, "
    strSQL = strSQL & "    [TANNM]     [nchar](20) NULL, "
    strSQL = strSQL & "    [UDNDT]     [nchar](8)  NULL DEFAULT '', "
    strSQL = strSQL & "    [URIKN]     [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [URIKNR]    [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [GENKN]     [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [GENKNR]    [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [JUZ]       [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [PUKN]      [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [PAKN]      [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "CONSTRAINT [PK_TANKT] PRIMARY KEY CLUSTERED "
    strSQL = strSQL & "( "
    strSQL = strSQL & "[TANCD] ASC "
    strSQL = strSQL & ") WITH "
    strSQL = strSQL & "(PAD_INDEX = OFF, "
    strSQL = strSQL & " STATISTICS_NORECOMPUTE = OFF, "
    strSQL = strSQL & " IGNORE_DUP_KEY = OFF, "
    strSQL = strSQL & " ALLOW_ROW_LOCKS = ON, "
    strSQL = strSQL & " ALLOW_PAGE_LOCKS = ON, "
    strSQL = strSQL & " OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF "
    strSQL = strSQL & ") ON [PRIMARY]"
    strSQL = strSQL & ") ON [PRIMARY]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

Sub DR_TBL_NKTT()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strSQL As String

    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'W_KA_NKT�e�[�u���폜
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_TA_NKT]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_TA_NKT]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

Sub CR_TBL_TAR()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strSQL As String

    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'NK_TAR�e�[�u���폜
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[NK_TAR]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[NK_TAR]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'NK_TAR�e�[�u���쐬
    strSQL = ""
    strSQL = strSQL & "CREATE TABLE [dbo].[NK_TAR]( "
    strSQL = strSQL & "    [SMADT]     [nchar](8)  NOT NULL, "
    strSQL = strSQL & "    [NKBN]      [nchar](10) NOT NULL, "
    strSQL = strSQL & "    [NKNM]      [nchar](20) NULL, "
    strSQL = strSQL & "    [UDNDT]     [nchar](8)  NULL DEFAULT '', "
    strSQL = strSQL & "    [URIKN]     [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [URIKNR]    [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [GENKN]     [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [GENKNR]    [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [JUZ]       [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [SIRKN]     [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [PUKN]      [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [PAKN]      [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "CONSTRAINT [PK_TAR] PRIMARY KEY CLUSTERED "
    strSQL = strSQL & "( "
    strSQL = strSQL & "[SMADT] ASC, "
    strSQL = strSQL & "[NKBN] ASC "
    strSQL = strSQL & ") WITH "
    strSQL = strSQL & "(PAD_INDEX = OFF, "
    strSQL = strSQL & " STATISTICS_NORECOMPUTE = OFF, "
    strSQL = strSQL & " IGNORE_DUP_KEY = OFF, "
    strSQL = strSQL & " ALLOW_ROW_LOCKS = ON, "
    strSQL = strSQL & " ALLOW_PAGE_LOCKS = ON, "
    strSQL = strSQL & " OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF "
    strSQL = strSQL & ") ON [PRIMARY]"
    strSQL = strSQL & ") ON [PRIMARY]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub

Sub CR_TBL_TAT()

Dim cnG    As New ADODB.Connection
Dim rsG    As New ADODB.Recordset
Dim strSQL As String

    cnG.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnG.Open
    
    'NK_TAT�e�[�u���폜
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[NK_TAT]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[NK_TAT]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'NK_TAT�e�[�u���쐬
    strSQL = ""
    strSQL = strSQL & "CREATE TABLE [dbo].[NK_TAT]( "
    strSQL = strSQL & "    [SMADT]     [nchar](8)  NOT NULL, "
    strSQL = strSQL & "    [TANCD]     [nchar](8)  NOT NULL, "
    strSQL = strSQL & "    [TANNM]     [nchar](20) NULL, "
    strSQL = strSQL & "    [UDNDT]     [nchar](8)  NULL DEFAULT '', "
    strSQL = strSQL & "    [URIKN]     [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [URIKNR]    [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [GENKN]     [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [GENKNR]    [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [JUZ]       [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [SIRKN]     [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [PUKN]      [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "    [PAKN]      [real]      NULL DEFAULT 0, "
    strSQL = strSQL & "CONSTRAINT [PK_TAT] PRIMARY KEY CLUSTERED "
    strSQL = strSQL & "( "
    strSQL = strSQL & "[SMADT] ASC, "
    strSQL = strSQL & "[TANCD] ASC "
    strSQL = strSQL & ") WITH "
    strSQL = strSQL & "(PAD_INDEX = OFF, "
    strSQL = strSQL & " STATISTICS_NORECOMPUTE = OFF, "
    strSQL = strSQL & " IGNORE_DUP_KEY = OFF, "
    strSQL = strSQL & " ALLOW_ROW_LOCKS = ON, "
    strSQL = strSQL & " ALLOW_PAGE_LOCKS = ON, "
    strSQL = strSQL & " OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF "
    strSQL = strSQL & ") ON [PRIMARY]"
    strSQL = strSQL & ") ON [PRIMARY]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic

    If Not rsG Is Nothing Then
        If rsG.State = adStateOpen Then rsG.Close
        Set rsG = Nothing
    End If
    If Not cnG Is Nothing Then
        If cnG.State = adStateOpen Then cnG.Close
        Set cnG = Nothing
    End If
    
End Sub
