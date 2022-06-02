Attribute VB_Name = "M07_KBN"
Option Explicit

Public KBN_NAME As String

Function KBN_CHG(strTCD As String, strGCD As String) As String
    '============================================================================================================
    '�敪�ύX����
    '============================================================================================================

    Select Case strGCD
        Case "0000000710004"
            If strTCD = "0000000710001" Then
                KBN_CHG = "G01" '��ٰ�ߔ��蓌��
            ElseIf strTCD = "0000000710002" Then
                KBN_CHG = "G02" '��ٰ�ߔ�����
            ElseIf strTCD = "0000000710003" Then
                KBN_CHG = "G03" '��ٰ�ߔ��蓌�C
            ElseIf strTCD = "0000000710004" Then
                KBN_CHG = "G04" '��ٰ�ߔ���{��
            ElseIf strTCD = "0000000710005" Then
                KBN_CHG = "G05" '��ٰ�ߔ����֓�
            ElseIf strTCD = "0000000710007" Then
                KBN_CHG = "G07" '��ٰ�ߔ��蕟��
            ElseIf strTCD = "0000000710009" Then
                KBN_CHG = "G09" '��ٰ�ߔ��薼�É�
             ElseIf strTCD = "0000000710011" Then
                KBN_CHG = "G11" '��ٰ�ߔ���k�֓�
            Else
                KBN_CHG = "G99" '��ٰ�ߔ��肻�̑�
            End If
        Case "0000000710505"
            KBN_CHG = "B01" '�P�C�A�C�X�^�[�s���Y
            KBN_NAME = "�P�C�A�C�X�^�[�s���Y"
        Case "0000000212012"
            KBN_CHG = "B02" '�ꌚ��
            KBN_NAME = "�ꌚ��"
        Case "0000000717000"
            KBN_CHG = "B03" '�ރ����f�B�n�E�X
            KBN_NAME = "�����ިʳ�"
        Case "0000000711001"
            KBN_CHG = "B04" '��������
            KBN_NAME = "��������"
        Case "0000000110775"
            KBN_CHG = "B05" '���h�Z��
            KBN_NAME = "���h�Z��"
        Case "0000000711014"
            KBN_CHG = "B06" '�A�Q��
            KBN_NAME = "�A�Q��"
        Case "0000000210215"
            KBN_CHG = "B07" '�A�C�_�݌v
            KBN_NAME = "�A�C�_�݌v"
        Case "0000000710014"
            KBN_CHG = "B08" '���C�u���[�n�E�X
            KBN_NAME = "���C�u���[�n�E�X"
        Case Else
            KBN_CHG = "B99"
            KBN_NAME = "���̑�"
    End Select

End Function

Function KBN_CHGT(strTCD As String, strGCD As String, strKBN As String, strKBP As String) As String
    '============================================================================================================
    '�敪�ύX����
    '============================================================================================================
    'G���ނœ��Ӑ�𔻒�
'    If strGCD = "0000000110215" Then
'        Stop
'    End If
    Select Case strGCD
        Case "0000000115024"
            KBN_CHGT = "A01" '��ΰ�
            KBN_NAME = "�^�}�z�[��"
        Case "0000000212012"
            KBN_CHGT = "A02" '�ꌚ��
            KBN_NAME = "�ꌚ��"
        Case "0000000811701"
            KBN_CHGT = "A03" 'AVANTIA
            KBN_NAME = "AVANTIA"
        Case "0000000210215"
            KBN_CHGT = "A04" '�A�C�_�݌v
            KBN_NAME = "�A�C�_�݌v"
        Case "0000000210700"
            KBN_CHGT = "A05" '�ѓc�Y��
            KBN_NAME = "�ѓc�Y��"
        Case "0000000812301"
            KBN_CHGT = "A06" '���ΰ�
            KBN_NAME = "���ΰ�"
        Case "0000000811101"
            KBN_CHGT = "A07" '���ʳ��ݸ�(�ޯ�����ݸ�)
            KBN_NAME = "���ʳ��ݸ�"
        Case "0000000812105"
            KBN_CHGT = "A09" '�������z
            KBN_NAME = "���̑�"
        Case "0000000812204"
            KBN_CHGT = "A09" '����ʳ�
            KBN_NAME = "���̑�"
        Case "0000000811401"
            KBN_CHGT = "A09" '�׼�ΰ�
            KBN_NAME = "���̑�"
        Case "0000000212305"
            KBN_CHGT = "A08" '���ިΰ�
            KBN_NAME = "���ިΰ�"
        Case "0000000115320"
            KBN_CHGT = "A09" '�ʑP
            KBN_NAME = "���̑�"
        Case "0000000111450"
            KBN_CHGT = "A09" '���
            KBN_NAME = "���̑�"
        Case "0000000216450"
            KBN_CHGT = "B04" 'θ��
            KBN_NAME = "�z�N�G�c"
        Case "0000000810001"
            KBN_CHGT = "B05" '�l��ܰ��
            KBN_NAME = "�l��ܰ��"
        Case "0000000810014"
            KBN_CHGT = "B09" '���㐻�|
            KBN_NAME = "���̑�"
        Case "0000000819011"
            KBN_CHGT = "A08" '��������(���萠)
            KBN_NAME = "��������(���萠)"
        Case "0000000812803"
            KBN_CHGT = "A08" '��������(���萠)
            KBN_NAME = "��������(���萠)"
        Case "0000000819001" '��������
            If strTCD = "0000000812301" Or strTCD = "0000000812303" Or strTCD = "0000000812308" Or strTCD = "0000000812350" Then  '���ΰ�(�����o�R)
                KBN_CHGT = "A06"
                KBN_NAME = "���ΰ�"
            ElseIf strTCD = "0000000819004" Then
                KBN_CHGT = "B04"
                KBN_NAME = "�{��(θ��)"
            Else
                '���i�敪(HINKB)�Ŕ���
                If Trim(strKBN) = "07" Or Trim(strKBN) = "16" Then      '��د��
                    KBN_CHGT = "B01"
                    KBN_NAME = "����������د��"
                ElseIf Trim(strKBN) = "08" Then '��۰��
                    KBN_CHGT = "B02"
                    KBN_NAME = "����������۰��"
                ElseIf Trim(strKBN) = "09" Then 'TL
                    KBN_CHGT = "B03"
                    KBN_NAME = "��������TL"
                Else
                    KBN_CHGT = "C09"
                    KBN_NAME = "���̑�"
                End If
            End If
        Case "0000000819010"           '�֓�AL
            KBN_CHGT = "C02"
            KBN_NAME = "�֓�AL"
        Case "0000000830035"           '���H�n��
            KBN_CHGT = "W99"
            KBN_NAME = "���H�n��"
        Case "0000000830099"           '��د�ލޗ�������
            KBN_CHGT = "Z99"
            KBN_NAME = "��د�ލޗ�������"
        Case Else
            If strTCD = "0000000810999" Then '�G����
                KBN_CHGT = "C09"
                KBN_NAME = "���̑�"
            ElseIf Trim(strKBP) = "01" Then
                KBN_CHGT = "A09" '�萠���̑�
                KBN_NAME = "���̑�"
            Else
                Select Case Trim(strKBN)
                    Case "11", "14", "16", "17"
                        KBN_CHGT = "A09" '�萠���̑�
                        KBN_NAME = "���̑�"
                    Case "07", "08", "09", "10"
                        KBN_CHGT = "B09" '�u���b�W���̑�
                        KBN_NAME = "���̑�"
                    Case Else
                        KBN_CHGT = "C09" '���̑�
                        KBN_NAME = "���̑�"
                End Select
            End If
        End Select
End Function
