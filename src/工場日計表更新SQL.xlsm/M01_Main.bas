Attribute VB_Name = "M01_Main"
Option Explicit

Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Const MAX_COMPUTERNAME_LENGTH = 15
Public start_time As Double
Public end_time   As Double

'===== �T�� =====
'�֓��A���R�b�N�A�y�ѓ��C�A���R�b�N�̔�����v�f�[�^���쐬����A�����X�V����
'����A�O���[�v�R�[�h�A���Ӑ�R�[�h�ŋ敪�𔻕ʁA�W�v����B
'�f�[�^�͔���A�󒍎c�A�v��A�d��
'�ݐσe�[�u���iNK_KJR�ENK_KJT�j�ɃZ�b�g
'�g�p�e�[�u���FOracle�iUDNTRA�ESDNTRA�ETOKMTA�EHINMTA�j
'              SQLServer�i�N�x�v��E�C���v��EJUZTBZ_Hybrid�j

Public Function CP_NAME() As String

    Const COMPUTERNAMBUFFER_LENGTH = MAX_COMPUTERNAME_LENGTH + 1
    Dim strComputerNameBuffer As String * COMPUTERNAMBUFFER_LENGTH
    Dim lngComputerNameLength As Long
    Dim lngWin32apiResultCode As Long
    
    ' �R���s���[�^�[���̒�����ݒ�
    lngComputerNameLength = Len(strComputerNameBuffer)
    ' �R���s���[�^�[�����擾
    lngWin32apiResultCode = GetComputerName(strComputerNameBuffer, _
                                            lngComputerNameLength)
    ' �R���s���[�^�[����\��
    CP_NAME = Left(strComputerNameBuffer, InStr(strComputerNameBuffer, vbNullChar) - 1)

End Function

Sub AP_END()
   
    Dim myBook As Workbook
    Dim strFN  As String
    Dim boolB  As Boolean
    
    'Excell���ɂ��̃u�b�N�ȊO�̃u�b�N���L���Excell���I�����Ȃ�
    ThisWorkbook.Save

    strFN = ThisWorkbook.Name '���̃u�b�N�̖��O
    boolB = False
    For Each myBook In Workbooks
        If myBook.Name <> strFN Then boolB = True
    Next
    If boolB Then
        ThisWorkbook.Close False  '�t�@�C�������
    Else
        Application.Quit  'Excell���I��
        ThisWorkbook.Saved = True
        ThisWorkbook.Close False
    End If
    
End Sub
