Attribute VB_Name = "M02_DATA"
Option Explicit

Public Const MYPROVIDERE = "Provider=SQLOLEDB;"
Public Const MYSERVER = "Data Source=192.168.128.9\SQLEXPRESS;"
Public Const USER = "User ID=sa;"
Public Const PSWD = "Password=ALCadmin!;"
Public Const strNT = "Initial Catalog=process_os;"

Sub Proc_TZ()

    Dim lngYYY  As Long    '年
    Dim lngMMM  As Long    '月
    Dim strMon  As String  '月末日付
    Dim DateA   As Date    '日付作業用
    
    '前月末取得
    lngYYY = CInt(Format(Now(), "yyyy"))
    lngMMM = CInt(Format(Now(), "mm"))
    DateA = CDate(lngYYY & "/" & lngMMM & "/01")
    DateA = DateA - 1
    strMon = Format(DateA, "yyyymmdd")
    
    
    Sheets("Wait").Range("D15") = "前月データ更新中・・・"
    
    Sheets("Wait").Range("D12") = "☆関東アルコック☆"
    DoEvents
'    strMon = "20210630"
    Call Proc_DataK(strMon)
    
    Sheets("Wait").Range("D12") = "☆東海アルコック☆"
    DoEvents
    Call Proc_DataT(strMon)
    
    '月末取得
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

    Sheets("Wait").Range("D15") = "当月データ更新中・・・"
    Sheets("Wait").Range("D12") = "☆関東アルコック☆"
    DoEvents
'    strMon = "20210731"
    Call Proc_DataK(strMon)
    
    Sheets("Wait").Range("D12") = "☆東海アルコック☆"
    DoEvents
    Call Proc_DataT(strMon)
    
    Sheets("Wait").Range("D15") = "終了！！"
    DoEvents
    
End Sub

Sub Proc_DataK(strMon As String)

    '作業テーブル作成
    Call CR_TBL_URI
    Call CR_TBL_JUZ
    Call CR_TBL_SRE
    Call CR_TBL_PLN
    Call CR_TBL_NK
    Call CR_TBL_NKT

    'データ取得
    Call GET_URI_K(strMon)  '売上ﾄﾗﾝから関東の売上だけをW_KA_URIへ抽出
    Call GET_JUC_K(strMon)  '受注ﾄﾗﾝから関東の受注だけをW_KA_JUZへ抽出
    Call GET_SIR_K(strMon)  '仕入ﾄﾗﾝから関東の仕入だけをW_KA_SREへ抽出
    Call GET_PLN_K(strMon)  '営業計画から関東の計画だけをW_KA_PLNへ抽出
    Call GET_TOK_K(strMon)  'W_KA_NKテーブルに売上、受注、仕入、計画を入れる
    Call GET_TAN_K(strMon)  'W_KA_NKTテーブルに売上、受注、仕入、計画を入れる
    Call UP_DATA_K(strMon)  'W_KA_NKテーブルからNK_KJRへ
    Call UP_TAN_K(strMon)   'W_KA_NKTテーブルからNK_KJTへ

    '作業テーブル削除
    Call DR_TBL_URI
    Call DR_TBL_JUZ
    Call DR_TBL_SRE
    Call DR_TBL_PLN
    Call DR_TBL_NK
    Call DR_TBL_NKT


End Sub

Sub Proc_DataT(strMon As String)

    '作業テーブル作成
    Call CR_TBL_TAU
    Call CR_TBL_TAJ
    Call CR_TBL_TAS
    Call CR_TBL_SRE
    Call CR_TBL_PLN
    Call CR_TBL_NKTA
    Call CR_TBL_NKTT

    'データ取得
    Call GET_URI_T(strMon)  '売上ﾄﾗﾝから東海の売上だけをW_TA_URIへ抽出
    Call GET_JUC_T(strMon)  '受注ﾄﾗﾝから東海の受注だけをW_TA_JUZへ抽出
    Call GET_SIR_T(strMon)  '仕入ﾄﾗﾝから東海の仕入だけをW_TA_SREへ抽出
    Call GET_SIR_T2(strMon) '仕入ﾄﾗﾝから東海の加工受けだけをW_KA_SREへ抽出
    Call GET_PLN_T(strMon)  '営業計画から東海の計画だけをW_TA_PLNへ抽出
    Call GET_TOK_T(strMon)  'W_TA_NKテーブルに売上、受注、仕入、計画を入れる
    Call GET_TAN_T(strMon)  'W_TA_NKTテーブルに売上、受注、仕入、計画を入れる
    Call UP_DATA_T(strMon)  'W_TA_NKテーブルからNK_KJRへ
    Call UP_TAN_T(strMon)   'W_TA_NKTテーブルからNK_KJTへ

    '作業テーブル削除
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
    
    'W_KA_URIテーブル削除
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_KA_URI]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_KA_URI]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'W_KA_URIテーブル作成
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
    
    'W_KA_URIテーブル削除
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
    
    'W_KA_JUZテーブル削除
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_KA_JUZ]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_KA_JUZ]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'W_KA_JUZテーブル作成
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
    
    'W_KA_JUZテーブル削除
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
    
    'W_KA_SREテーブル削除
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_KA_SRE]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_KA_SRE]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'W_KA_SREテーブル作成
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
    
    'W_KA_SREテーブル削除
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
    
    'W_PLNテーブル削除
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_PLN]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_PLN]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'W_PLNテーブル作成
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
    
    'W_PLNテーブル削除
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
    
    'W_KA_NKテーブル削除
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_KA_NK]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_KA_NK]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'W_KA_NKテーブル作成
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
    
    'W_KA_NKテーブル削除
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
    
    'W_KA_NKTテーブル削除
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_KA_NKT]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_KA_NKT]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'W_KA_NKTテーブル作成
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
    
    'W_KA_NKTテーブル削除
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
    
    'NK_KJRテーブル削除
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[NK_KJR]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[NK_KJR]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'NK_KJRテーブル作成
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
    
    'NK_KJTテーブル削除
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[NK_KJT]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[NK_KJT]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'NK_KJTテーブル作成
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
    
    'W_TA_URIテーブル削除
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_TA_URI]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_TA_URI]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'W_TA_URIテーブル作成
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
    
    'W_TA_URIテーブル削除
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
    
    'W_TA_JUZテーブル削除
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_TA_JUZ]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_TA_JUZ]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'W_TA_JUZテーブル作成
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
    
    'W_TA_JUZテーブル削除
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
    
    'W_TA_SREテーブル削除
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_TA_SRE]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_TA_SRE]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'W_TA_SREテーブル作成
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
    
    'W_TA_SREテーブル削除
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
    
    'W_TA_NKテーブル削除
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_TA_NK]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_TA_NK]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'W_KA_NKテーブル作成
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
    
    'W_TA_NKテーブル削除
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
    
    'W_TA_NKTテーブル削除
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[W_TA_NKT]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[W_TA_NKT]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'W_TA_NKTテーブル作成
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
    
    'W_KA_NKTテーブル削除
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
    
    'NK_TARテーブル削除
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[NK_TAR]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[NK_TAR]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'NK_TARテーブル作成
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
    
    'NK_TATテーブル削除
    strSQL = ""
    strSQL = strSQL & "if exists (select * from sysobjects where id = "
    strSQL = strSQL & "object_id(N'[dbo].[NK_TAT]') and "
    strSQL = strSQL & "OBJECTPROPERTY(id, N'IsUserTable') = 1) "
    strSQL = strSQL & "DROP TABLE [dbo].[NK_TAT]"
    rsG.Open strSQL, cnG, adOpenStatic, adLockOptimistic
    
    'NK_TATテーブル作成
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
