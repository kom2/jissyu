Attribute VB_Name = "Module1"
'��ǎ��K�X�P�W���[����\
'                                           H. Komatsu
'------------------------------------------------------------
'                                           create 2014.01.09
'
'2019.02.18) �S���҂̕\�����e�L�X�g�{�b�N�X�Ƃ��Ԋ|�����s��
'2019.02.12) �u�����Z���I�������K�̘A�����{�v���͋@�\�ǉ�
'2019.02.11) �c�蓖�Ԃ̃��[�e�[�V�����ύX
'2019.02.04) ���K���ԕύX�ɑΉ�
'2015.05.31) ��\�Ώہi�T�j��I���\�Ƃ���
'

Const CST_SHT_KIHON         As String = "��{�ݒ�"
Const CST_SHT_KOMOKU        As String = "���K����"
Const CST_SHT_TOPPAGE       As String = "�\��"
Const CST_SHT_SCHEDULE1     As String = "�X�P�W���[��1"
Const CST_SHT_SCHEDULE2     As String = "�X�P�W���[��2"

'��{�ݒ�V�[�g�̃f�[�^�̈�
Const CST_RANGE_K_NENDO     As String = "C5"            '�N�x
Const CST_RANGE_K_KI        As String = "C6"            '��
Const CST_RANGE_K_KAISI     As String = "C7"            '���ԁi�J�n���j
Const CST_RANGE_K_MANRYO    As String = "C8"            '���ԁi�������j
Const CST_RANGE_K_GAKUSEI   As String = "B12:F14"       '�w���i���K���j
Const CST_RANGE_K_YAKUZAISI As String = "B20:F28"       '�E���i��܎t�j
Const CST_RANGE_K_JIMU      As String = "B30:F38"       '�E���i�����j
Const CST_RANGE_K_YOTEI     As String = "B43:F55"       '���ԓ��̗\��
Const CST_RANGE_K_MOKUHYO   As String = "J6:J55"        '�T���̖ڕW�i�ۑ�j

'���K���ڃV�[�g�̃f�[�^�̈�
Const CST_RANGE_J_KAISI_MON As String = "D2"            '��P�T�̌��j���ƂȂ�����Z�b�g����Z��
Const CST_RANGE_J_CAL_TOP   As String = "K1"            '�J�����_�[�̈�̍���i�g�b�v�j
Const CST_RANGE_J_CAL_DATE  As String = "K3:CP3"        '�J�����_�[���t��
Const CST_RANGE_J_CHECK     As String = "A5"            '�`�F�b�N��
Const CST_RANGE_J_ITEM_TOP  As String = "B5"            '���K���ڂ̍���i�g�b�v�j
Const CST_RANGE_J_KOMOKU    As String = "D5"            '���K���ځi���o���j
Const CST_RANGE_J_PRINT     As String = "E5"            '�󎚓��e��
Const CST_RANGE_J_JIKAN     As String = "F5"            '���ԑ�
Const CST_RANGE_J_FUKUSU    As String = "G5"            '�������K��
Const CST_RANGE_J_TANTO     As String = "H5"            '�S��

'�\���V�[�g�̃f�[�^�̈�
Const CST_RANGE_TP_TITLE    As String = "B3"            '�^�C�g�����i�����m�m�N�x�m���@��ǁc�X�P�W���[���j
Const CST_RANGE_TP_KIKAN    As String = "B5"            '����
Const CST_RANGE_TP_GAKUSEI  As String = "C9"            '���K����
Const CST_RANGE_TP_SIDO     As String = "F10"           '�w����܎t��
Const CST_RANGE_TP_SYOKUIN  As String = "C14"           '�݂ȂݐE����
Const CST_RANGE_TP_JIMU     As String = "F15"           '�����E����

'�X�P�W���[���P�i�P�`�W�T�j�̃f�[�^�̈�
Const CST_RANGE_S1_CREATE   As String = "P1"            '�쐬��
Const CST_RANGE_S1_TITLE    As String = "B2"            '�^�C�g����
Const CST_RANGE_S1_MOKUHYO  As String = "B3:B5"         '�ڕW�i�ۑ�j��
Const CST_RANGE_S1_TABLE    As String = "B7:Q39"        '�\
Const CST_RANGE_S1_AM       As String = "B8:Q23"        '�ߑO�̗̈�
Const CST_RANGE_S1_PM       As String = "B24:Q39"       '�ߌ�̗̈�

'�X�P�W���[���Q�i9�`11�T�j�̃f�[�^�̈�
Const CST_RANGE_S2_CREATE   As String = "P1"            '�쐬��
Const CST_RANGE_S2_TITLE    As String = "B2"            '�^�C�g����
Const CST_RANGE_S2_MOKUHYO  As String = "B3:B5"         '�ڕW�i�ۑ�j��
Const CST_RANGE_S2_TABLE1   As String = "B7:Q16"        '�\(9�T��)
Const CST_RANGE_S2_AM       As String = "B8:G12"        '�ߑO�̗̈�
Const CST_RANGE_S2_PM       As String = "B13:G16"       '�ߌ�̗̈�
Const CST_RANGE_S2_TABLE2   As String = "B19:Q28"       '�\(10�T��)
Const CST_RANGE_S2_TABLE3   As String = "B31:Q40"       '�\(11�T��)

'���K���ڃV�[�g�ɂ�����f�[�^�s�̎��ʋL��
Const CST_MARK_NOKORI_TOBAN As String = "��1"           '�c�蓖��
Const CST_MARK_NOKORI_SUB   As String = "��2"           '�c�蓖�ԁi�T�u�j
Const CST_MARK_KYUJITU      As String = "��"            '�x��
Const CST_MARK_ITEM         As String = "��"            '���K����

'���j���[�R�}���h
Const CST_MENU_1            As String = "���K��A�����{" '�E�N���b�N�����ۂɕ\�����郁�j���[

Const CST_LINE_WEIGHT       As Single = 0.75            '���i�R�l�N�g�}�`�j�̐��̑���(0.25pt/0.5pt/0.75pt/1.0pt/1.25pt/1.5pt/1.75pt/2.0pt��)
Const CST_TANTO_SHAPE       As Boolean = True           '�S���҂̕\����}�`�`��(�Ԋ|��)�Ƃ��邩�ǂ���

'�ҏW�f�[�^
Type tpData
    rowIdx                  As Long                     '�sindex
    title                   As String                   '���ڃ^�C�g��
    prnData(20)             As String                   '�󎚓��e
    timetbl                 As Long                     '���Ԋ�(�ߑO:1�`5,�ߌ�6�`10)
    strTime                 As String                   '���Ԋ��i������j
    fukusu                  As String                   '�������K��
    tanto                   As String                   '�S��
    settei                  As String                   '�J�����_�����͒l
    seqno                   As Long                     '�\�[�g�p
    sortkey                 As String                   '�\�[�g�p
    prnStartRow             As Long                     '����J�n�s
    prnRowsCount            As Long                     '����s���i���א��j
    yajirusi                As Long                     '���̒����i�e�ɃZ�b�g�j
    yajirusi2               As Long                     '���̒����i�q�ɃZ�b�g�j
    dmyline                 As Long                     '�ʒu���킹�ׂ̈̋�s��
End Type

    
'���b�Z�[�W�\��
Public Sub dispMessage(strMessage)
    With UserForm1
        .setMessage (strMessage)
        .Caption = ThisWorkbook.Name
        .Show vbModeless
        .Repaint
    End With
End Sub

'���b�Z�[�W����
Public Sub hideMessage()
    Unload UserForm1
End Sub

'��{�ݒ�V�[�g�̕ύX
Public Sub wSheetChage_Kihon(ByVal target As Range)
    Dim cell1       As Range
    
    If target.Columns.Count > 100 Or target.Rows.Count > 100 Then
        '�s�I���Ȃǂ͏������Ȃ�
        Exit Sub
    End If
    
    For Each cell1 In target.Cells
        Call cellChange_Kihon(cell1)
    Next
End Sub

'���K���ڃV�[�g�̕ύX
Public Sub wSheetChage_Jissyu(ByVal target As Range)
    Dim cell1       As Range
    
    If target.Columns.Count > 100 Or target.Rows.Count > 100 Then
        '�s�I���Ȃǂ͏������Ȃ�
        Exit Sub
    End If
    
    For Each cell1 In target.Cells
        Call cellChange_Jissyu(cell1)
    Next
End Sub

'���K���ڃV�[�g�̃Z���N�V�����ύX
Public Sub wSheetSelectionChage_Jissyu(ByVal target As Range)
    Dim obj1        As Object
    
    '�E�N���b�N���j���[�̓Ǝ����j���[���폜
    On Error Resume Next
    Application.CommandBars("Cell").Controls(CST_MENU_1).Delete
    
    If target.Cells.Count = 1 Then
        Call selectionChange_Jissyu(target)
    Else
        If target.Cells.Count > 1 And target.Areas.Count = 1 And target.Rows.Count = 1 Then
            '�P��s�ɂ����鉡���������Z���̑I�����A�����{����
            Set obj1 = Application.CommandBars("Cell").Controls.Add()
            With obj1
                .Caption = CST_MENU_1
                .OnAction = "renzoku_jissi"
                .BeginGroup = False
            End With
        End If
    End If
End Sub

'���K�̘A�����{�i�E�N���b�N���j���[�j
Public Sub renzoku_jissi()
    Dim rng1        As Range
    Dim i           As Long
    Dim str1        As String
    Dim str2        As String
    Dim c           As Long
    Dim r1          As Long
    Dim r2          As Long
    Dim r3          As Long
    Dim d1          As Date
    
    If ActiveSheet.Name <> CST_SHT_KOMOKU Then
        MsgBox ("���K���ڃV�[�g��I�����Ă�������")
        Exit Sub
    End If
    
    Set rng1 = Selection
    
    If rng1.Cells.Count < 2 Or rng1.Areas.Count <> 1 Or rng1.Rows.Count <> 1 Then
        Set rng1 = Nothing
        Exit Sub
    End If
    
    If rng1.Cells(1).Column < Range(CST_RANGE_J_CAL_TOP).Column Or _
       rng1.Cells(1).Row < Range(CST_RANGE_J_ITEM_TOP).Row Then
        MsgBox ("�I��͈͂��s���ł��B")
        Exit Sub
    End If
    
    c = Range(CST_RANGE_J_JIKAN).Column                 '���ԑт̌�
    r1 = rng1.Cells(1).Row                              '���͑Ώۂ̍s�i�I���s�j
    r2 = Range(CST_RANGE_J_CAL_DATE).Row                '�J�����_�̍s
    r3 = fnGetKujitsuRow                                '�x���̍s
    
    If r3 = 0 Then
        MsgBox ("�x���̃f�[�^�̈悪������܂���")
        Exit Sub
    End If
    
    '���ԑ�
    If ActiveSheet.Cells(r1, c).Value <> "" Then
        str1 = Left(CStr(ActiveSheet.Cells(r1, c).Value), 1)
    Else
        str1 = "A"
    End If
    c = rng1.Cells(1).Column
    If c > Range(CST_RANGE_J_CAL_DATE).Column Then
        For i = c To Range(CST_RANGE_J_CAL_DATE).Column Step -1
            If Left(UCase(ActiveSheet.Cells(r1, i).Value), 1) = "A" Or _
               Left(UCase(ActiveSheet.Cells(r1, i).Value), 1) = "P" Then
                str1 = ActiveSheet.Cells(r1, i).Value
                Exit For
            End If
        Next i
    End If
    
    For i = 1 To rng1.Cells.Count
        c = rng1.Cells(i).Column
        If c >= Range(CST_RANGE_J_CAL_DATE).Column And _
            c <= (Range(CST_RANGE_J_CAL_DATE).Column + Range(CST_RANGE_J_CAL_DATE).Columns.Count - 1) Then
            '���K�\��̓��͗̈�
            d1 = ActiveSheet.Cells(r2, c).Value
            If Weekday(d1) = 1 Or Weekday(d1) = 7 Or ActiveSheet.Cells(r3, c) <> "" Then
                '���j�E�y�j�E�x��
                str2 = ""
            Else
                If d1 = fnKaisi Or Weekday(d1) = 2 Then
                    '���K�J�n���܂��͌��j��
                    str2 = str1
                Else
                    If ActiveSheet.Cells(r1, c - 1).Value = "" Then
                        str2 = str1
                    Else
                        str2 = "��"
                    End If
                End If
            End If
            ActiveSheet.Cells(r1, c).Value = str2
        End If
    Next i
    
    
    
    Set rng1 = Nothing
    
End Sub


'�x���s�̎擾
Public Function fnGetKujitsuRow()
    Dim r                   As Long
    Dim c                   As Long
    Dim rMax                As Long
    
    fnGetKujitsuRow = 0
    With ThisWorkbook.Sheets(CST_SHT_KOMOKU)
        rMax = .UsedRange.Rows.Count
        c = Range(CST_RANGE_J_ITEM_TOP).Column
        For r = Range(CST_RANGE_J_ITEM_TOP).Row To rMax
            If .Cells(r, c).Value = CST_MARK_KYUJITU Then
                fnGetKujitsuRow = r
                Exit For
            End If
        Next r
    End With
    
End Function



'��{�ݒ�V�[�g�̃Z���ύX
Private Sub cellChange_Kihon(ByVal target As Range)
    Dim r1                  As Long
    Dim c1                  As Long
    Dim i                   As Long

    '���̓Z���̃A�h���X���`�F�b�N���āA���͂��ꂽ���ڂɉ����ď������s��
    
    '�C�x���g�𖳌��i�A���̗}�~�j
    Application.EnableEvents = False
    
    
    '���ԊJ�n���̓��́��J�����_�[�ύX
    If target.Address(False, False) = Range(CST_RANGE_K_KAISI).Address(False, False) Then
        Call changeKaisi(target)
        GoTo term_change_kihon
    End If
    
    '���Ԗ������̓��́��J�����_�[�ύX
    If target.Address(False, False) = Range(CST_RANGE_K_MANRYO).Address(False, False) Then
        Call changeManryo(target)
        GoTo term_change_kihon
    End If
    
    
    '�T���̖ڕW�i�ۑ�j�^�C�g��
    r1 = Range(CST_RANGE_K_MOKUHYO).Row
    c1 = Range(CST_RANGE_K_MOKUHYO).Column
    For i = 1 To 10
        If target.Row = r1 And target.Column = c1 Then
            Call changeMokuhyo(i)
        End If
        r1 = r1 + 5
    Next i
    
    '�\��
    If target.Row >= Range(CST_RANGE_K_YOTEI).Row And _
        target.Row <= Range(CST_RANGE_K_YOTEI).Row + Range(CST_RANGE_K_YOTEI).Rows.Count And _
        target.Column >= Range(CST_RANGE_K_YOTEI).Column And _
        target.Column <= Range(CST_RANGE_K_YOTEI).Column + Range(CST_RANGE_K_YOTEI).Columns.Count Then
        Call setYotei2Comment
    End If
        
        
    

term_change_kihon:

    '�C�x���g�L���֖߂�
    Application.EnableEvents = True

End Sub


'���K���ڃV�[�g�̃Z���ύX
Private Sub cellChange_Jissyu(ByVal target As Range)
    Dim r1                  As Long
    Dim c1                  As Long
    Dim i                   As Long

    '�C�x���g�𖳌��i�A���̗}�~�j
    Application.EnableEvents = False
    
    '�ΏۊO���H�i���o�������j
    If target.Row < Range(CST_RANGE_J_ITEM_TOP).Row Then
        GoTo term_change_Jissyu
    End If
    
    '���K���ځA�^�C�g���ύX
    If target.Column >= Range(CST_RANGE_J_ITEM_TOP).Column And _
        target.Column <= Range(CST_RANGE_J_ITEM_TOP).Column + 3 Then
        Call changeJissyuKomoku(target.Row)
        GoTo term_change_Jissyu
    End If
    
    '�J�����_��
    If target.Column >= Range(CST_RANGE_J_CAL_DATE).Column And _
        target.Column <= Range(CST_RANGE_J_CAL_DATE).Column + Range(CST_RANGE_J_CAL_DATE).Columns.Count - 1 Then
        Call changeCalendar(target.Row, target.Column)
        GoTo term_change_Jissyu
    End If

term_change_Jissyu:

    '�C�x���g�L���֖߂�
    Application.EnableEvents = True

End Sub


'���K���ڃV�[�g�̃Z���N�V�����ύX�i�T�u�j
Private Sub selectionChange_Jissyu(ByVal target As Range)
    Dim colMark             As Long                     '�}�[�N��index
    Dim colTitle            As Long                     '���K���ځi�^�C�g���j��index
    Dim colPrn              As Long                     '�󎚓��e
    Dim colTimetbl          As Long                     '���Ԋ��敪
    Dim colTanto            As Long                     '�S��
    Dim colFukusu           As Long                     '�������K��
    Dim rowDate             As Long                     '�J�����_�̓��t�s
    Dim date1               As Long
    Dim kyujitu             As String
    Dim nokori              As String
    Dim nokoriSub           As String
    Dim cnt                 As Long
    Dim data1(100)          As tpData
    Dim tmpdata             As tpData
    Dim strMark             As String
    Dim str1                As String
    Dim str2                As String
    Dim str3                As String
    Dim str4                As String
    Dim v1                  As Variant
    Dim timtbl              As Long
    Dim flg1                As Boolean
    Dim i                   As Long
    Dim j                   As Long
    Dim columnIdx           As Long
    
    '�C�x���g�𖳌��i�A���̗}�~�j
    Application.EnableEvents = False

    '�ΏۊO���H�i���t�����j
    If target.Row < Range(CST_RANGE_J_CAL_DATE).Row Or _
        target.Row > Range(CST_RANGE_J_CAL_DATE).Row + Range(CST_RANGE_J_CAL_DATE).Rows.Count - 1 Or _
        target.Column < Range(CST_RANGE_J_CAL_DATE).Column Or _
        target.Column > Range(CST_RANGE_J_CAL_DATE).Column + Range(CST_RANGE_J_CAL_DATE).Columns.Count - 1 Then
        GoTo term_selectionChange_Jissyu
    End If
    
    columnIdx = target.Column
    
    colMark = Range(CST_RANGE_J_ITEM_TOP).Column
    colTitle = Range(CST_RANGE_J_KOMOKU).Column
    colTimetbl = Range(CST_RANGE_J_JIKAN).Column
    colFukusu = Range(CST_RANGE_J_FUKUSU).Column
    colTanto = Range(CST_RANGE_J_TANTO).Column
    colPrn = Range(CST_RANGE_J_PRINT).Column
    rowDate = Range(CST_RANGE_J_CAL_DATE).Row
    
    With ThisWorkbook.Sheets(CST_SHT_KOMOKU)
        
        date1 = .Cells(rowDate, columnIdx).Value
        cnt = 0
        Erase data1
        kyujitu = ""
        nokori = ""
        nokoriSub = ""
        For i = Range(CST_RANGE_J_ITEM_TOP).Row To .UsedRange.Rows.Count
            str1 = CStr(.Cells(i, columnIdx).Value)
            If str1 <> "" Then
                strMark = CStr(.Cells(i, colMark).Value)
                Select Case strMark
                    Case CST_MARK_KYUJITU
                        kyujitu = "���x��"
                    Case CST_MARK_NOKORI_TOBAN
                        nokori = str1
                    Case CST_MARK_NOKORI_SUB
                        nokoriSub = str1
                    Case CST_MARK_ITEM
                        If str1 = "��" Then
                            str2 = lookLeft(i, columnIdx)
                        Else
                            str2 = str1
                        End If
                        If CStr(.Cells(i, colTimetbl).Value) <> "" Then
                            str2 = CStr(.Cells(i, colTimetbl).Value)
                        End If
                        v1 = Split(str2, ",")
                        For k = 0 To UBound(v1)
                            str3 = v1(k)
                            timtbl = cnvTimetbl(str3)
                            If timtbl > 0 Then
                                cnt = cnt + 1
                                data1(cnt).rowIdx = i
                                data1(cnt).title = CStr(.Cells(i, colTitle).Value)
                                data1(cnt).settei = str1
                                data1(cnt).timetbl = timtbl
                                If IsNumeric(.Cells(i, colFukusu).Value) Then
                                    data1(cnt).fukusu = fnGakuseiSan(.Cells(i, colFukusu).Value)
                                Else
                                    data1(cnt).fukusu = CStr(.Cells(i, colFukusu).Value)
                                End If
                                data1(cnt).tanto = CStr(.Cells(i, colTanto).Value)
                                'data1(cnt).seqno = timtbl * 100 + cnt
                                str4 = UCase(str3)
                                If str4 = "A" Then
                                    str4 = "A2"
                                End If
                                If str4 = "P" Then
                                    str4 = "P3"
                                End If
                                data1(cnt).strTime = str4
                                data1(cnt).sortkey = str4 & CStr(1000 + cnt)
                                data1(cnt).prnRowsCount = 1
                                data1(cnt).prnData(1) = CStr(.Cells(i, colPrn).Value)
                                For l = 1 To 19
                                    If CStr(.Cells(i + l, colMark).Value) = CST_MARK_ITEM Or _
                                        CStr(.Cells(i + l, colPrn).Value) = "" Then
                                        Exit For
                                    End If
                                    data1(cnt).prnRowsCount = l + 1
                                    data1(cnt).prnData(l + 1) = CStr(.Cells(i + l, colPrn).Value)
                                Next l
                                
                            End If
                        Next k
                    Case Else
                End Select
                
            End If
        Next i
        
        If cnt > 1 Then
            'SORT
            For i = 1 To cnt
                flg1 = True
                For j = 2 To cnt
                    If data1(j - 1).sortkey > data1(j).sortkey Then
                        flg1 = False
                        tmpdata = data1(j - 1)
                        data1(j - 1) = data1(j)
                        data1(j) = tmpdata
                    End If
                Next j
                If flg1 Then
                    Exit For
                End If
            Next i
        End If
        str1 = ""
        If kyujitu <> "" Then
            str1 = kyujitu & vbLf
        End If
        str2 = ""
        If cnt > 0 Then
            For i = 1 To cnt
                str1 = str1 & "���@" & data1(i).title & vbLf
                For j = 1 To data1(i).prnRowsCount
                    str3 = data1(i).prnData(j)
                    If j = 1 And data1(i).tanto <> "" Then
                        str3 = str3 & "  �y" & data1(i).tanto & "�z"
                    End If
                    If j = 1 Then
                        str3 = "�E" & str3
                    Else
                        str3 = "�@" & str3
                    End If
                    If fnGakuseiSu() > 1 Then
                        If data1(i).fukusu <> "" Then
                            str3 = Replace(str3, "[@]", data1(i).fukusu)
                        End If
                    Else
                        str3 = Replace(str3, "[@]", "")
                    End If
                    str2 = str2 & str3 & vbLf
                Next j
            Next i
        End If
'        If str1 = "" Then
'            .Cells(rowDate, columnIdx).ClearComments
'        Else
'            str1 = Format(date1, "m/d(aaa)") & " " & nokori & "/" & nokoriSub & vbLf & str1
'            .Cells(rowDate, columnIdx).ClearComments
'            .Cells(rowDate, columnIdx).AddComment str1
'            .Cells(rowDate, columnIdx).Comment.Shape.TextFrame.AutoSize = True
'        End If
    End With
    
    If str1 <> "" Then
        Call UserForm2.setValues(StrConv(Format(date1, "m��d��(aaa)"), vbWide), nokori & "/" & nokoriSub, str1, str2)
        UserForm2.Show vbModal
    End If




term_selectionChange_Jissyu:

    '�C�x���g�L���֖߂�
    Application.EnableEvents = True
    
End Sub

'�T���̖ڕW�i�ۑ�j�^�C�g�����ύX���ꂽ
Private Sub changeMokuhyo(weekNo As Long)
    Dim str1                As String
    Dim r1                  As Long
    Dim c1                  As Long
    Dim r2                  As Long
    Dim c2                  As Long
    
    If weekNo < 1 Or weekNo > 10 Then
        Exit Sub
    End If
    
    r1 = Range(CST_RANGE_K_MOKUHYO).Row + (weekNo - 1) * 5
    c1 = Range(CST_RANGE_K_MOKUHYO).Column
    str1 = CStr(ThisWorkbook.Sheets(CST_SHT_KIHON).Cells(r1, c1).Value)
    
    r2 = Range(CST_RANGE_J_CAL_TOP).Row
    c2 = Range(CST_RANGE_J_CAL_TOP).Column + (weekNo - 1) * 7
    
    With ThisWorkbook.Sheets(CST_SHT_KOMOKU).Cells(r2, c2)
        .ClearComments
        .AddComment (str1)
'        .Comment.Shape.TextFrame.AutoSize = True
    End With
    
    
End Sub


'���K���Ԃ̊J�n�����ύX���ꂽ���J�����_�[�ݒ�
Private Sub changeKaisi(target As Range)
    Dim day1                As Long
    Dim day2                As Long
    Dim day3                As Long
    Dim strNendo            As String
    Dim strKi               As String
    
    If IsDate(target.Value) = False Then
        MsgBox ("�����������t����͂��ĉ�����")
        Exit Sub
    End If
    
    'Call dispMessage("�������ł�...")
    
    day1 = target.Value
    
    '�N�x�A�����Z�o
    '------------------------------------------------------------
    ' (2018�N�x�܂Ł���1���F5���`7���A��2���F9���`11���A��3���F1���`3��)
    ' (2019�N�x���灨��1���F2���`5���A��2���F5���`8���A��3���F8���`11���A��4���F11���`2��)
    '
    If day1 < CDate("2019/01/31") Then
        '<<<2018�N�x�ȑO>>>
        If Month(day1) < 4 Then
            day2 = CDate(CStr(Year(day1) - 1) & "/4/1")
            strKi = "��R��"
        Else
            day2 = CDate(CStr(Year(day1) & "/4/1"))
            If Month(day1) > 7 Then
                strKi = "��Q��"
            Else
                strKi = "��P��"
            End If
        End If
    Else
        '<<<2019�N�x�ȍ~>>>
        Select Case True
            Case Month(day1) = 1
                day2 = CDate(CStr(Year(day1) - 1) & "/4/1")
                strKi = "��S��"
            Case Month(day1) < 5
                day2 = CDate(CStr(Year(day1)) & "/4/1")
                strKi = "��P��"
            Case Month(day1) < 8
                day2 = CDate(CStr(Year(day1)) & "/4/1")
                strKi = "��Q��"
            Case Month(day1) < 11
                day2 = CDate(CStr(Year(day1)) & "/4/1")
                strKi = "��R��"
            Case Else
                day2 = CDate(CStr(Year(day1)) & "/4/1")
                strKi = "��S��"
        End Select
    End If
    strNendo = Format(day2, "ggge�N�x")
    ThisWorkbook.Sheets(CST_SHT_KIHON).Range(CST_RANGE_K_NENDO).Value = strNendo
    ThisWorkbook.Sheets(CST_SHT_KIHON).Range(CST_RANGE_K_KI).Value = strKi
    
    
    
    '�J�����_�̊J�n���A���Ԗ��������Z�o����i�Փ��ȂǂŌ��j���ȊO����̊J�n�����蓾�邽�߁j
    day2 = day1 - ((Weekday(day1) + 5) Mod 7)           'WeekDay�֐��Ō��j���́u2�v
    day3 = day2 + (7 * 11) - 3
    MsgBox ("���Ԗ��������u" & Format(day3, "yyyy/m/d") & "�v�Ƃ��܂��B" & vbLf & "���t���قȂ�ꍇ�͏C�����ĉ������B")
    ThisWorkbook.Sheets(CST_SHT_KIHON).Range(CST_RANGE_K_MANRYO).Value = day3
    
    '�J�����_���t�̃Z�b�g
    Call setDate2Calendar
    
    '�J�����_�[��̗\����`�F�b�N
    Call setYotei2Comment
    
    'Call hideMessage
    
End Sub


'���K���Ԃ̖��������ύX���ꂽ���J�����_�[�ݒ�
Private Sub changeManryo(target As Range)
    
    If IsDate(target.Value) = False Then
        MsgBox ("�����������t����͂��ĉ�����")
        Exit Sub
    End If
    
    '�J�����_���t�̃Z�b�g
    Call setDate2Calendar
    
    '�J�����_�[��̗\����`�F�b�N
    Call setYotei2Comment
    
    
End Sub


'�J�����_���t�̃Z�b�g�A�x���`�F�b�N
Private Sub setDate2Calendar()
    Dim date1               As Long         '���ԊJ�n��
    Dim date2               As Long         '���Ԗ�����
    Dim date3               As Long         '�P�T�ڂ̌��j�i�J�����_�̐擪�j
    Dim i                   As Long
    Dim c                   As Long
    Dim r                   As Long
    Dim rKyujitu            As Long
    Dim flg1                As Boolean      '���K�Ώۓ��i���ԊO�A�y���E�x����False�j
    
    date1 = fnKaisi()
    date2 = fnManryo()
    date3 = fn1stMonday()
    
    If date1 > date2 Then
        MsgBox ("���K���Ԃ̓��͂��s���ł��B")
        Exit Sub
    End If
    
    c = Range(CST_RANGE_J_CAL_DATE).Column
    r = Range(CST_RANGE_J_CAL_DATE).Row
    rKyujitu = markedRow(CST_MARK_KYUJITU)
    
    With ThisWorkbook.Sheets(CST_SHT_KOMOKU)
        For i = date3 To date3 + Range(CST_RANGE_J_CAL_DATE).Columns.Count - 1
            flg1 = True
            .Cells(r, c).Value = i
            .Cells(r, c).ClearComments
            .Cells(r + 1, c).Value = i
            If i < date1 Or i > date2 Then
                flg1 = False
            End If
            If Weekday(i) = 1 Or Weekday(i) = 7 Then
                flg1 = False
            End If
            If CStr(.Cells(rKyujitu, c).Value) <> "" Then
                flg1 = False
            End If
                
            With .Cells(r, c).Interior
                If flg1 Then
                    '���K�Ώۓ�
                    .Pattern = xlNone
                    .PatternColorIndex = xlAutomatic
                    '.PatternTintAndShade = 0
                Else
                    '���ԊO�܂��͓y���E�x��
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = -0.14996795556505
                    .PatternTintAndShade = 0
                End If
            End With
            c = c + 1
        Next i
    End With
    
End Sub


'���K���ڂ̕ύX
Private Sub changeJissyuKomoku(rowIdx As Long)
    Call checkJissyuKomoku(rowIdx)
End Sub

'�J�����_���̕ύX
Private Sub changeCalendar(rowIdx As Long, columnIdx As Long)
    Dim colMark             As Long
    Dim rowDate             As Long
    Dim date1               As Long
    Dim flg1                As Boolean
    
    Call checkJissyuKomoku(rowIdx)
    
    'Call setCommentOnCalendar(columnIdx)
    
    colMark = Range(CST_RANGE_J_ITEM_TOP).Column
    rowDate = Range(CST_RANGE_J_CAL_DATE).Row
    
    With ThisWorkbook.Sheets(CST_SHT_KOMOKU)
        date1 = CLng(.Cells(rowDate, columnIdx).Value)
        If CStr(.Cells(rowIdx, colMark).Value) = CST_MARK_KYUJITU Then
            '�x���̍s
            flg1 = False
            If CStr(.Cells(rowIdx, columnIdx).Value) <> "" Then
                '�x��
                flg1 = True
            End If
            If Weekday(date1) = 1 Or Weekday(date1) = 7 Then
                flg1 = True
            End If
            With .Cells(rowDate, columnIdx).Interior
                If flg1 Then
                    '�x���܂��͓y��
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = -0.14996795556505
                    .PatternTintAndShade = 0
                Else
                    '�x���ȊO
                    .Pattern = xlNone
                    .PatternColorIndex = xlAutomatic
                    '.PatternTintAndShade = 0
                End If
            End With
        End If
    End With

End Sub


'�w��s�̎��K���ڂɂ��ăX�P�W���[���ݒ�̗L�����`�F�b�N����
Private Sub checkJissyuKomoku(rowIdx As Long)
    Dim colCheck            As Long                     '�`�F�b�N��
    Dim colMark             As Long                     '�}�[�N�i���ڎ��ʁj��
    Dim colTitle            As Long                     '�^�C�g����
    Dim colPrn              As Long                     '�󎚓��e��
    Dim colCalendar1        As Long                     '�J�����_���̊J�n��
    Dim colCalendar2        As Long                     '�J�����_���̏I�[��
    Dim strDate             As String
    Dim i                   As Long
    Dim j                   As Long
    Dim c                   As Long
    Dim str1                As String
    
    colCheck = Range(CST_RANGE_J_CHECK).Column
    colMark = Range(CST_RANGE_J_ITEM_TOP).Column
    colTitle = Range(CST_RANGE_J_KOMOKU).Column
    colPrn = Range(CST_RANGE_J_PRINT).Column
    colCalendar1 = Range(CST_RANGE_J_CAL_DATE).Column
    colCalendar2 = colCalendar1 + Range(CST_RANGE_J_CAL_DATE).Columns.Count - 1
    
    With ThisWorkbook.Sheets(CST_SHT_KOMOKU)
        If CStr(.Cells(rowIdx, colMark).Value) = CST_MARK_ITEM Then
            c = colCalendar1
            For i = 1 To 12
                str1 = ""
                For j = 1 To 7
                    If CStr(.Cells(rowIdx, c).Value) <> "" Then
                        str1 = str1 & Mid("���ΐ��؋��y��", j, 1)
                    End If
                    c = c + 1
                Next j
                If str1 <> "" Then
                    strDate = strDate & "��" & CStr(i) & "�T(" & str1 & ") "
                End If
            Next i
            If Len(strDate) > 50 Then
                strDate = Left(strDate, 50) & "..."
            End If
            If strDate = "" Then
                .Cells(rowIdx, colCheck).Value = "��"
                .Cells(rowIdx, colTitle).ClearComments
            Else
                .Cells(rowIdx, colCheck).Value = ""
                .Cells(rowIdx, colTitle).ClearComments
                .Cells(rowIdx, colTitle).AddComment (strDate)
                .Cells(rowIdx, colTitle).Comment.Shape.TextFrame.AutoSize = True
            End If
            If .Cells(rowIdx, colTitle).Value = "" Then
                .Cells(rowIdx, colTitle).Interior.ColorIndex = 38
            Else
                .Cells(rowIdx, colTitle).Interior.ColorIndex = 0
            End If

            If .Cells(rowIdx, colPrn).Value = "" Then
                .Cells(rowIdx, colPrn).Interior.ColorIndex = 38
            Else
                .Cells(rowIdx, colPrn).Interior.ColorIndex = 0
            End If
        End If
        If CStr(.Cells(rowIdx, colMark).Value) = "" Then
                .Cells(rowIdx, colCheck).Value = ""
                .Cells(rowIdx, colTitle).ClearComments
                .Cells(rowIdx, colTitle).Interior.ColorIndex = 0
                .Cells(rowIdx, colPrn).Interior.ColorIndex = 0
        End If
    End With

End Sub



'�C�x���g�\��Ȃǂ��J�����_�[�̓��t���̃R�����g�փZ�b�g
Public Sub setYotei2Comment()
    Dim evtNo               As Long
    Dim evtDate(20)         As Long
    Dim evtStr(20)          As String
    Dim r                   As Long
    Dim c                   As Long
    Dim i                   As Long
    Dim str1                As String
    Dim str2                As String
    Dim date1               As Long
    
    r = Range(CST_RANGE_K_YOTEI).Row
    c = Range(CST_RANGE_K_YOTEI).Column
    
    evtNo = 0
    For i = 1 To Range(CST_RANGE_K_YOTEI).Columns.Count
        str1 = CStr(ThisWorkbook.Sheets(CST_SHT_KIHON).Cells(r + i - 1, c))     '�\��i�\��j
        str2 = CStr(ThisWorkbook.Sheets(CST_SHT_KIHON).Cells(r + i - 1, c + 2)) '���t
        If IsDate(str2) Then
            date1 = CDate(str2)
        Else
            date1 = 0
        End If
        If str1 <> "" And date1 > 0 Then
            evtNo = evtNo + 1
            evtDate(evtNo) = date1
            evtStr(evtNo) = str1
        End If
    Next i
    
    r = Range(CST_RANGE_J_CAL_DATE).Row
    c = Range(CST_RANGE_J_CAL_DATE).Column
    date1 = fn1stMonday()
    For i = 1 To (7 * 12)
        With ThisWorkbook.Sheets(CST_SHT_KOMOKU).Cells(r, c)
            .ClearComments
            For j = 1 To evtNo
                If date1 = evtDate(j) Then
                    .AddComment (evtStr(j))
                    .Comment.Shape.TextFrame.AutoSize = True
                End If
            Next j
            c = c + 1
            date1 = date1 + 1
        End With
    Next i
        
End Sub


'�c�蓖�Ԃ̃N���A
Public Sub clearNokori()
    Dim r                   As Long
    Dim c                   As Long
    
    If MsgBox("�c�蓖�ԁE�T�u�̓��͂�S�ď������܂��B" & vbLf & "��낵���ł����H", vbOKCancel) = vbCancel Then
        Exit Sub
    End If
    
    Application.EnableEvents = False
    
    c = Range(CST_RANGE_J_CAL_DATE).Column
    c2 = c + (7 * 12) - 1
    
    '�c��N���A
    r = markedRow(CST_MARK_NOKORI_TOBAN)
    If r > 0 Then
        ThisWorkbook.Sheets(CST_SHT_KOMOKU).Range(Cells(r, c), Cells(r, c2)).Select
        Selection.ClearContents
        ThisWorkbook.Sheets(CST_SHT_KOMOKU).Cells(r, c).Select
    End If
    
    '�T�u�N���A
    r = markedRow(CST_MARK_NOKORI_SUB)
    If r > 0 Then
        ThisWorkbook.Sheets(CST_SHT_KOMOKU).Range(Cells(r, c), Cells(r, c2)).Select
        Selection.ClearContents
        ThisWorkbook.Sheets(CST_SHT_KOMOKU).Cells(r, c).Select
    End If
    
    Application.EnableEvents = True
    
End Sub

'�c�蓖�Ԃ̃Z�b�g
'2019.02.11 �c�蓖�Ԃ��T�l�̏ꍇ�͋x���ɍS��炸�T���ɃV�t�g����
' ��j��:1 ��:2 ��:3 ��:4 ��:5   ��:2 ��:3 ��:4 ��:5 ��:1   ��:3 ��:4 ��:5 ��:1 ��:2
Public Sub setNokori()
    Dim r1                  As Long
    Dim r2                  As Long
    Dim c1                  As Long
    Dim c2                  As Long
    Dim i                   As Long
    Dim j                   As Long
    Dim k                   As Long
    Dim str1                As String
    Dim str2                As String
    Dim cnt                 As Long
    Dim arNokori(10)        As String
    Dim date1               As Long
    Dim date2               As Long
    Dim x                   As Long
    Dim kyujitu             As Boolean
    Dim weekcnt             As Long
    
    Application.EnableEvents = False
    
    '�c�蓖�Ԃ̐ݒ���擾
    r1 = Range(CST_RANGE_K_YAKUZAISI).Row
    c1 = Range(CST_RANGE_K_YAKUZAISI).Column
    cnt = 0
    For i = 1 To 9
        str1 = CStr(ThisWorkbook.Sheets(CST_SHT_KIHON).Cells(r1 + i - 1, c1 + 3).Value)     '�c�蓖�ԏ�
        str2 = CStr(ThisWorkbook.Sheets(CST_SHT_KIHON).Cells(r1 + i - 1, c1 + 4).Value)     '�S������
        If IsNumeric(str1) And str2 <> "" Then
            j = CLng(str1)
            If j > 0 And j < 10 Then
                arNokori(j) = str2
                If j > cnt Then
                    cnt = j
                End If
            End If
        End If
    Next i
    
    If cnt = 0 Then
        '���Ԃ̐ݒ肪����
        Exit Sub
    End If
    
    '�J�����_�T��
    date1 = fnKaisi()           '���ԁi�J�n���j
    date2 = fnManryo()          '���ԁi�������j
    If date2 < date1 Then
        Exit Sub
    End If
    
    If cnt <> 5 Then
        '<<<<<�c�Ɠ����ɏ��Ԃ̃��[�e�[�V����(�]�O)>>>>>>
        For i = 1 To 2
            If i = 1 Then
                r1 = markedRow(CST_MARK_NOKORI_TOBAN)
            Else
                r1 = markedRow(CST_MARK_NOKORI_SUB)
            End If
            r2 = markedRow(CST_MARK_KYUJITU)    '�x���̍sindex
            x = 0                               '�O���i�O�c�Ɠ��j�̎c�蓖��index
            c1 = date2ColumnIdx(date1)          '�ݒ�J�n���index
            For j = date1 To date2
                str1 = CStr(ThisWorkbook.Sheets(CST_SHT_KOMOKU).Cells(r1, c1).Value)
                str2 = CStr(ThisWorkbook.Sheets(CST_SHT_KOMOKU).Cells(r2, c1).Value)
                If Weekday(j) <> 1 And Weekday(j) <> 7 And str2 = "" Then
                    '�y�E���E�x���ȊO
                    If str1 <> "" Then
                        '���͍ς݁�����index�̎擾
                        For k = 1 To cnt
                            If str1 = arNokori(k) Then
                                x = k
                                Exit For
                            End If
                        Next k
                    Else
                        If x > 0 Then
                            '���̓��Ԃ��Z�b�g
                            x = (x Mod cnt) + 1
                            ThisWorkbook.Sheets(CST_SHT_KOMOKU).Cells(r1, c1).Value = arNokori(x)
                        End If
                    End If
                End If
                c1 = c1 + 1
            Next j
        Next i
    Else
        '<<<<<�T�l�̐��̃��[�e�[�V����(�T���ɃV�t�g���x���͍l�����Ȃ�)>>>>>>
        For i = 1 To 2
            If i = 1 Then
                r1 = markedRow(CST_MARK_NOKORI_TOBAN)
            Else
                r1 = markedRow(CST_MARK_NOKORI_SUB)
            End If
            r2 = markedRow(CST_MARK_KYUJITU)    '�x���̍sindex
            x = 0                               '�O���i�O�c�Ɠ��j�̎c�蓖��index
            c1 = date2ColumnIdx(date1)          '�ݒ�J�n���index
            For j = date1 To date2
                str1 = CStr(ThisWorkbook.Sheets(CST_SHT_KOMOKU).Cells(r1, c1).Value)    '�c�蓖��
                str2 = CStr(ThisWorkbook.Sheets(CST_SHT_KOMOKU).Cells(r2, c1).Value)    '�x��
                If Weekday(j) = 1 Then
                    weekcnt = weekcnt + 1
                End If
                If Weekday(j) <> 1 And Weekday(j) <> 7 And str2 = "" Then
                    '�y�E���E�x���ȊO
                    If str1 <> "" Then
                        '���͍ς݁�����index�̎擾
                        For k = 1 To cnt
                            If str1 = arNokori(k) Then
                                x = k
                                weekcnt = x + 9 - Weekday(j)
                                Exit For
                            End If
                        Next k
                    Else
                        If x > 0 Then
                            '���̓��Ԃ��Z�b�g
                            x = ((weekcnt + Weekday(j)) Mod cnt) + 1
                            ThisWorkbook.Sheets(CST_SHT_KOMOKU).Cells(r1, c1).Value = arNokori(x)
                        End If
                    End If
                End If
                c1 = c1 + 1
            Next j
        Next i
    End If
    Application.EnableEvents = True

End Sub


'================
'�X�P�W���[����\
'================
Public Sub makeSchedule()
    Dim newWorkBook         As Workbook
    Dim i                   As Long
    Dim weeks               As Long
    Dim str1                As String
    Dim sht                 As Long
    Dim targetWeek          As Variant
    
    weeks = fnWeeks()
    If weeks < 1 Then
        MsgBox ("���K���Ԃ��s���ł�")
        Exit Sub
    End If


'    '��\�����J�n�̊m�F
'    If MsgBox("�X�P�W���[���\���쐬���܂��B" & vbLf & "��낵���ł����H", vbOKCancel) = vbCancel Then
'        Exit Sub
'    End If
    
    '��\�Ώۂ̑I��
    UserForm3.Show
    If UserForm3.Tag = "" Then
        Unload UserForm3
        Exit Sub
    End If
    targetWeek = UserForm3.selctionData
    Unload UserForm3
    
    Call dispMessage("�������ł�...")
    
    '�V�K�u�b�N�𐶐����A�\���y�уX�P�W���[���̊e�V�[�g��{�u�b�N���R�s�[���A���e��ݒ肵�܂�
    sht = 0
    If targetWeek(0) = True Then
        ThisWorkbook.Sheets(CST_SHT_TOPPAGE).Copy
        Set newWorkBook = ActiveWorkbook
        ActiveSheet.Name = "�\��"
        '�\���V�[�g�̓��e�Z�b�g
        Call setTopPage(ActiveSheet)
        sht = 1
    End If
    For i = 1 To weeks
        If i <= 8 Then
            If targetWeek(i) = True Then
                '�X�P�W���[���P�i��P�T�`��W�T�j�̓��e�Z�b�g
                If sht = 0 Then
                    ThisWorkbook.Sheets(CST_SHT_SCHEDULE1).Copy
                    Set newWorkBook = ActiveWorkbook
                Else
                    ThisWorkbook.Sheets(CST_SHT_SCHEDULE1).Copy after:=newWorkBook.Sheets(sht)
                End If
                ActiveSheet.Name = "��" & CStr(i) & "�T��"
                sht = sht + 1
                Call setSchedule1(ActiveSheet, i)
            End If
        End If
        If i >= 9 And (i Mod 3) = 0 Then
            If targetWeek(i) = True Then
                '�X�P�W���[���Q�i��9�T�`�j�̓��e�Z�b�g
                If sht = 0 Then
                    ThisWorkbook.Sheets(CST_SHT_SCHEDULE2).Copy
                    Set newWorkBook = ActiveWorkbook
                Else
                    ThisWorkbook.Sheets(CST_SHT_SCHEDULE2).Copy after:=newWorkBook.Sheets(sht)
                End If
                If i + 3 < weeks Then
                    ActiveSheet.Name = "��" & CStr(i) & "�T�ڂ����" & CStr(i + 2) & "�T��"
                Else
                    ActiveSheet.Name = "��" & CStr(i) & "�T�ڈȍ~"
                End If
                sht = sht + 1
                Call setSchedule2(ActiveSheet, i)
            End If
        End If
    Next i
    
    newWorkBook.Sheets(1).Select
    
    Set newWorkBook = Nothing

    Call hideMessage

End Sub


'�\���V�[�g�̓��e�Z�b�g
Private Sub setTopPage(targetSheet As Worksheet)
    Dim strNendo            As String
    Dim strKi               As String
    Dim dateKaisi           As Long
    Dim dateManryo          As Long
    Dim nissu               As Long
    Dim str1                As String
    Dim str2                As String
    Dim r1                  As Long
    Dim r2                  As Long
    Dim c1                  As Long
    Dim c2                  As Long
    Dim c3                  As Long
    Dim i                   As Long
    Dim cnt                 As Long
    Dim cnt2                As Long
    
    With ThisWorkbook.Sheets(CST_SHT_KIHON)
        '���K�N�x�E���E����
        strNendo = CStr(.Range(CST_RANGE_K_NENDO).Value)
        strKi = CStr(.Range(CST_RANGE_K_KI).Value)
        dateKaisi = .Range(CST_RANGE_K_KAISI).Value
        dateManryo = .Range(CST_RANGE_K_MANRYO).Value
        nissu = nannichime(dateManryo)
        targetSheet.Range(CST_RANGE_TP_TITLE).Value = strNendo & strKi & " ��ǎ��K���@�X�P�W���[��"
        targetSheet.Range(CST_RANGE_TP_KIKAN).Value = "��" & Format(dateKaisi, "ggge�Nm��d���iaaaa�j") & _
                                                    " �` " & Format(dateManryo, "ggge�Nm��d���iaaaa�j") & _
                                                    "�F�v" & CStr(nissu) & "���ԁ�"
        '���K��
        cnt = 0
        r1 = Range(CST_RANGE_K_GAKUSEI).Row
        c1 = Range(CST_RANGE_K_GAKUSEI).Column
        r2 = Range(CST_RANGE_TP_GAKUSEI).Row
        c2 = Range(CST_RANGE_TP_GAKUSEI).Column
        c3 = Range(CST_RANGE_TP_SIDO).Column
        For i = 1 To 3
            str1 = CStr(.Cells(r1 + i - 1, c1).Value)
            If str1 <> "" Then
                cnt = cnt + 1
                str1 = CStr(.Cells(r1 + i - 1, c1 + 2).Value) & "   " & str1 & _
                        StrConv(" (" & CStr(.Cells(r1 + i - 1, c1 + 1).Value) & ")", vbNarrow)
                str2 = "�S����܎t�F" & CStr(.Cells(r1 + i - 1, c1 + 4).Value)
                targetSheet.Cells(r2 + i, c2).Value = str1
                targetSheet.Cells(r2 + i, c3).Value = str2
            End If
        Next i
        targetSheet.Cells(r2, c2) = "�����K���i" & StrConv(CStr(cnt), vbWide) & "���j"
        
        '�݂ȂݐE��
        cnt = 0
        r1 = Range(CST_RANGE_K_YAKUZAISI).Row
        c1 = Range(CST_RANGE_K_YAKUZAISI).Column
        r2 = Range(CST_RANGE_TP_SYOKUIN).Row
        c2 = Range(CST_RANGE_TP_SYOKUIN).Column
        '��܎t
        For i = 1 To 9
            str1 = CStr(.Cells(r1 + i - 1, c1).Value)
            If str1 <> "" Then
                cnt = cnt + 1
                str1 = str1 & StrConv(" (" & CStr(.Cells(r1 + i - 1, c1 + 1).Value) & ")", vbNarrow)
                str2 = CStr(.Cells(r1 + i - 1, c1 + 2).Value)
                If str2 <> "" Then
                    str1 = str1 & " �F " & str2
                End If
                targetSheet.Cells(r2 + i + 1, c2).Value = str1
            End If
        Next i
        targetSheet.Cells(r2 + 1, c2).Value = "�E��܎t�i" & StrConv(CStr(cnt), vbWide) & "���j"
        '����
        cnt2 = 0
        r1 = Range(CST_RANGE_K_JIMU).Row
        c1 = Range(CST_RANGE_K_JIMU).Column
        c2 = Range(CST_RANGE_TP_JIMU).Column
        For i = 1 To 9
            str1 = CStr(.Cells(r1 + i - 1, c1).Value)
            If str1 <> "" Then
                cnt2 = cnt2 + 1
                str1 = str1 & StrConv(" (" & CStr(.Cells(r1 + i - 1, c1 + 1).Value) & ")", vbNarrow)
                str2 = CStr(.Cells(r1 + i - 1, c1 + 2).Value)
                If str2 <> "" Then
                    str1 = str1 & " �F " & str2
                End If
                targetSheet.Cells(r2 + i + 1, c2).Value = str1
            End If
        Next i
        targetSheet.Cells(r2 + 1, c2).Value = "�E�����i" & StrConv(CStr(cnt2), vbWide) & "���j"
        targetSheet.Range(CST_RANGE_TP_SYOKUIN).Value = "���H�c�݂Ȃ݉�c��ǐE���i" & StrConv(CStr(cnt + cnt2), vbWide) & "���j"
        
    
    End With

End Sub


'�w��̓��t�����K�J�n�����牽���ڂɂ����邩��Ԃ�
Private Function nannichime(jdate As Long) As Long
    Dim i                   As Long
    Dim j                   As Long
    Dim c                   As Long
    Dim dateKaisi           As Long
    Dim dateManryo          As Long
    Dim date1st             As Long
    Dim cnt                 As Long
    Dim cx                  As Long
    
    nannichime = 0
    
    dateKaisi = fnKaisi()
    dateManryo = fnManryo()
    date1st = fn1stMonday()
    
    If jdate < dateKaisi Or jdate > dateManryo Then
        '���K���ԊO
        Exit Function
    End If
    
    cnt = 0
    
    '�͂��߂ɋx���f�[�^�̍s��T��
    i = markedRow(CST_MARK_KYUJITU)
    If i > 0 Then
        With ThisWorkbook.Sheets(CST_SHT_KOMOKU)
            cx = date2ColumnIdx(dateKaisi)
            For j = dateKaisi To jdate
                '�x���A�y�j���A���j���ȊO�̓����J�E���g
                If CStr(.Cells(i, cx).Value) = "" And _
                    Weekday(j) <> 7 And _
                    Weekday(j) <> 1 Then
                    cnt = cnt + 1
                End If
                cx = cx + 1
            Next j
        End With
    End If
    
    nannichime = cnt
    
End Function


'���K���ڃV�[�g�ɂ�����}�[�N�i�f�[�^���ʁj���T�����A�Y������sindex��Ԃ�
'������Ȃ��ꍇ�́u0�v��Ԃ�
Private Function markedRow(strMark As String) As Long
    Dim c                   As Long
    Dim i                   As Long
    
    markedRow = 0
    
    c = Range(CST_RANGE_J_ITEM_TOP).Column
    
    With ThisWorkbook.Sheets(CST_SHT_KOMOKU)
        r = 0
        For i = Range(CST_RANGE_J_ITEM_TOP).Row To .UsedRange.Rows.Count
            If .Cells(i, c).Value = strMark Then
                markedRow = i
                Exit For
            End If
        Next i
    End With

End Function


'�J�����_�ɂ����āA�w����̗�(index)��Ԃ�
Private Function date2ColumnIdx(jdate As Long) As Long
    Dim dateKaisi           As Long
    
    date2ColumnIdx = 0
    
    dateKaisi = fn1stMonday()
    If jdate < dateKaisi Or jdate > fnManryo() Then
        '�J�����_�͈͊O
        Exit Function
    End If
    
    date2ColumnIdx = (jdate - dateKaisi) + Range(CST_RANGE_J_CAL_TOP).Column
    
End Function


'���K���̐l��
Private Function fnGakuseiSu() As Long
    Dim r                   As Long
    Dim c                   As Long
    Dim i                   As Long
    Dim cnt                 As Long
    
    fnGakuseiSu = 0
    
    r = Range(CST_RANGE_K_GAKUSEI).Row
    c = Range(CST_RANGE_K_GAKUSEI).Column
    
    cnt = 0
    For i = r To r + 2
        If CStr(ThisWorkbook.Sheets(CST_SHT_KIHON).Cells(i, c).Value) <> "" Then
            cnt = cnt + 1
        End If
    Next i
    
    fnGakuseiSu = cnt
    
End Function

'N�Ԗڂ̊w������i�����{�u����v�j
Private Function fnGakuseiSan(no As Long) As String
    If no = 1 Or no = 2 Or no = 3 Then
        fnGakuseiSan = " (" & ThisWorkbook.Sheets(CST_SHT_KIHON).Cells(Range(CST_RANGE_K_GAKUSEI).Row + no - 1, Range(CST_RANGE_K_GAKUSEI).Column).Value & " ����)"
    Else
        fnGakuseiSan = ""
    End If
End Function


'���K���Ԃ̊J�n��
Private Function fnKaisi() As Long
    fnKaisi = CLng(ThisWorkbook.Sheets(CST_SHT_KIHON).Range(CST_RANGE_K_KAISI).Value)
End Function


'���K���Ԃ̖�����
Private Function fnManryo() As Long
    fnManryo = CLng(ThisWorkbook.Sheets(CST_SHT_KIHON).Range(CST_RANGE_K_MANRYO).Value)
End Function

'���K���Ԃ͉��T�Ԗڂ܂ł��邩��Ԃ�
Public Function fnWeeks() As Long
    Dim d1                  As Long
    Dim d2                  As Long
    
    d1 = fn1stMonday()
    d2 = fnManryo()
    
    If d1 > d2 Then
        fnWeeks = 0
    Else
        fnWeeks = Int((d2 - d1) / 7) + 1
    End If
End Function


'�J�����_�[�̐擪���i�J�n�����܂ޏT�̌��j���j
Public Function fn1stMonday() As Long
    Dim d                   As Long
    
    d = fnKaisi()
    fn1stMonday = d - ((Weekday(d) + 5) Mod 7)  'WeekDay�֐��Ō��j���́u2�v
End Function



'�X�P�W���[���V�[�g�P�i��P�T�`��W�T�j�̓��e�Z�b�g
Private Sub setSchedule1(targetSheet As Worksheet, weekNo As Long)
    Dim str1                As String
    Dim r1                  As Long
    Dim r2                  As Long
    Dim c1                  As Long
    Dim c2                  As Long
    Dim i                   As Long
    
    '�^�C�g���E�ۑ�i�ڕW�j
    r1 = Range(CST_RANGE_K_MOKUHYO).Row + (weekNo - 1) * 5
    c1 = Range(CST_RANGE_K_MOKUHYO).Column
    With ThisWorkbook.Sheets(CST_SHT_KIHON)
        '�^�C�g��
        targetSheet.Range(CST_RANGE_S1_TITLE).Value = .Cells(r1, c1).Value
        '�ۑ�i�ڕW�j
        r2 = Range(CST_RANGE_S1_MOKUHYO).Row
        c2 = Range(CST_RANGE_S1_MOKUHYO).Column
        For i = 1 To 4
            targetSheet.Cells(r2 + i - 1, c2).Value = .Cells(r1 + i, c1).Value
        Next i
    End With
    
    '���t�E�c�蓖��
    targetSheet.Range(CST_RANGE_S1_CREATE).Value = Format(Now(), "(ge.m.d�쐬)")
    Call setDateNokori(targetSheet, CST_RANGE_S1_TABLE, weekNo)
    
    '�X�P�W���[�����e
    Call setScheduleData(targetSheet, CST_RANGE_S1_TABLE, CST_RANGE_S1_AM, weekNo)
    
    
End Sub

'�X�P�W���[���V�[�g�Q�i��9�T�`�j�̓��e�Z�b�g�i�R�T�ԁ^�V�[�g�j
Private Sub setSchedule2(targetSheet As Worksheet, weekNo As Long)
    Dim str1                As String
    Dim r1                  As Long
    Dim r2                  As Long
    Dim c1                  As Long
    Dim c2                  As Long
    Dim i                   As Long
    Dim weeks               As Long
    Dim rw                  As Long
    Dim v1                  As Variant
    
    
    '�^�C�g���E�ۑ�i�ڕW�j
    If weekNo <= 9 Then
        rw = 9
    Else
        rw = 10
    End If
    r1 = Range(CST_RANGE_K_MOKUHYO).Row + (rw - 1) * 5
    c1 = Range(CST_RANGE_K_MOKUHYO).Column
    With ThisWorkbook.Sheets(CST_SHT_KIHON)
        '�^�C�g��
        targetSheet.Range(CST_RANGE_S2_TITLE).Value = .Cells(r1, c1).Value
        '�ۑ�i�ڕW�j
        r2 = Range(CST_RANGE_S2_MOKUHYO).Row
        c2 = Range(CST_RANGE_S2_MOKUHYO).Column
        For i = 1 To 4
            targetSheet.Cells(r2 + i - 1, c2).Value = .Cells(r1 + i, c1).Value
        Next i
    End With
    
    weeks = fnWeeks()
    
    '���t�E�c�蓖��,�X�P�W���[�����e
    targetSheet.Range(CST_RANGE_S2_CREATE).Value = Format(Now(), "(ge.m.d�쐬)")
    If weekNo <= weeks Then
        Call setDateNokori(targetSheet, CST_RANGE_S2_TABLE1, weekNo)
        Call setScheduleData(targetSheet, CST_RANGE_S2_TABLE1, CST_RANGE_S2_AM, weekNo)
    End If
    If weekNo + 1 <= weeks Then
        Call setDateNokori(targetSheet, CST_RANGE_S2_TABLE2, weekNo + 1)
        Call setScheduleData(targetSheet, CST_RANGE_S2_TABLE2, CST_RANGE_S2_AM, weekNo + 1)
    End If
    If weekNo + 2 <= weeks Then
        Call setDateNokori(targetSheet, CST_RANGE_S2_TABLE3, weekNo + 2)
        Call setScheduleData(targetSheet, CST_RANGE_S2_TABLE3, CST_RANGE_S2_AM, weekNo + 2)
    End If
    
End Sub

'�X�P�W���[���V�[�g�ւ̓��t�E�c�蓖�ԃZ�b�g
Private Sub setDateNokori(targetSheet As Worksheet, targetRange As String, weekNo As Long)
    Dim str1                As String
    Dim r1                  As Long
    Dim r2                  As Long
    Dim c1                  As Long
    Dim c2                  As Long
    Dim i                   As Long
    Dim date1               As Long
    Dim date2               As Long
    Dim nisu                As Long
    Dim rKyujitu            As Long
    Dim rNokori             As Long
    Dim rNokoriSub          As Long
    Dim strNokori           As String
    Dim strNokoriSub        As String
    
    
    '���t�E�c�蓖��
    With ThisWorkbook.Sheets(CST_SHT_KOMOKU)
        r1 = Range(targetRange).Row
        c1 = Range(targetRange).Column
        targetSheet.Cells(r1, c1).Value = " " & StrConv(CStr(weekNo), vbWide) & "�T��"
        c1 = c1 + 2
        rKyujitu = markedRow(CST_MARK_KYUJITU)          '�x���̍sindex
        rNokori = markedRow(CST_MARK_NOKORI_TOBAN)      '�c�蓖�Ԃ̍sindex
        rNokoriSub = markedRow(CST_MARK_NOKORI_SUB)     '�c�蓖�ԁi�T�u�j�̍sindex
        date1 = fn1stMonday() + (weekNo - 1) * 7
        date2 = date1 + 4
        If date2 > fnManryo() Then
            date2 = fnManryo()
        End If
        c2 = date2ColumnIdx(date1)
        For i = date1 To date2
            str1 = StrConv(Format(i, "m/d(aaa)"), vbWide)
            If CStr(.Cells(rKyujitu, c2).Value) <> "" Then
                '�x��
                str1 = str1 & "�F�x��"
            Else
                nisu = nannichime(i)                    '�����ڂ�
                str1 = str1 & "�F" & StrConv(CStr(nisu), vbWide) & "����"
                strNokori = CStr(.Cells(rNokori, c2).Value)
                strNokoriSub = CStr(.Cells(rNokoriSub, c2).Value)
                If strNokori <> "" Then
                    '�c�蓖��
                    str1 = str1 & "�@" & strNokori
                    If strNokoriSub <> "" Then
                        str1 = str1 & "/" & strNokoriSub & ""
                    End If
                End If
                    
            End If
            targetSheet.Cells(r1, c1).Value = str1
            c1 = c1 + 3
            c2 = c2 + 1
        Next i
    End With

End Sub


'�X�P�W���[�����e�̃Z�b�g
Private Sub setScheduleData(targetSheet As Worksheet, targetRange As String, amRange As String, weekNo As Long)
    Dim date1               As Long
    Dim date2               As Long
    Dim colidx              As Long
    Dim i                   As Long
    Dim j                   As Long
    Dim k                   As Long
    Dim l                   As Long
    Dim d                   As Long
    Dim r1                  As Long
    Dim r2                  As Long
    Dim r3                  As Long
    Dim c1                  As Long
    Dim str1                As String
    Dim str2                As String
    Dim str3                As String
    Dim colMark             As Long             '���ڂ����ʂ���}�[�N�̗�index
    Dim colKomoku           As Long             '���K���ځi���o���j�̗�index
    Dim colPrn              As Long             '�󎚓��e�̗�index
    Dim colJikan            As Long             '���ԑт̗�index
    Dim colFukusu           As Long             '����
    Dim colTanto            As Long             '�S��
    Dim v1                  As Variant
    Dim timetbl             As Long
    Dim rc                  As Long
    
    
    
    Dim dcount              As Long
    Dim data1(5, 100)       As tpData
    Dim dcnt1(5)            As Long
    Dim tmpdata             As tpData
    Dim flg1                As Boolean
    Dim timeNum(10)         As Long                 '���ԑі��̈󎚍s��
    Dim timeRow(10)         As Long                 '���ԑі��̈󎚊J�n�sindex
    Dim tmpNum(10)          As Long
    Dim timeStr(10)         As String               '���ԕ\��
    
    colMark = Range(CST_RANGE_J_ITEM_TOP).Column
    colKomoku = Range(CST_RANGE_J_KOMOKU).Column
    colPrn = Range(CST_RANGE_J_PRINT).Column
    colJikan = Range(CST_RANGE_J_JIKAN).Column
    colFukusu = Range(CST_RANGE_J_FUKUSU).Column
    colTanto = Range(CST_RANGE_J_TANTO).Column
    
    With ThisWorkbook.Sheets(CST_SHT_KOMOKU)
    
        date1 = fn1stMonday() + (weekNo - 1) * 7
        date2 = date1 + 4
        
        If date2 > fnManryo() Then
            date2 = fnManryo()
        End If
        
        '�ҏW�f�[�^�̈�N���A
        Erase data1
        Erase timeNum
        
        '�����Ƀf�[�^��ҏW�̈�փZ�b�g
        For i = 1 To 5
            d = date1 + i - 1
            
            dcount = 0
            
            If d < fnKaisi() Or d > fnManryo() Then
                '�ΏۊO
            Else
                colidx = date2ColumnIdx(d)
                r1 = Range(CST_RANGE_J_ITEM_TOP).Row
                '�J�����_����T��
                For j = r1 To .UsedRange.Rows.Count
                    str1 = CStr(.Cells(j, colidx).Value)
                    If str1 <> "" Then
                        str2 = CStr(.Cells(j, colMark).Value)
                        If str2 = CST_MARK_ITEM Then
                            '���K����
                            str3 = str1
                            If str3 = "��" Then
                                str3 = lookLeft(j, colidx)  '�������擾
                            End If
                            If CStr(.Cells(j, colJikan).Value) <> "" Then
                                str3 = CStr(.Cells(j, colJikan).Value)
                            End If
                            v1 = Split(str3, ",")       '�ua,p�v�ȂǌߑO���ߌ�̃p�^�[���Ή�
                            For k = 0 To UBound(v1)
                                str3 = v1(k)
                                
                                timetbl = cnvTimetbl(str3)
                                If timetbl > 0 Then
                                    '�L���Ȏ��ԑ�
                                    dcount = dcount + 1
                                    data1(i, dcount).rowIdx = j
                                    data1(i, dcount).title = CStr(.Cells(j, colKomoku).Value)
                                    data1(i, dcount).timetbl = timetbl
                                    'data1(i, dcount).prnData
                                    If IsNumeric(.Cells(j, colFukusu).Value) Then
                                        data1(i, dcount).fukusu = fnGakuseiSan(.Cells(j, colFukusu).Value)
                                    Else
                                        data1(i, dcount).fukusu = CStr(.Cells(j, colFukusu).Value)
                                    End If
                                    data1(i, dcount).tanto = CStr(.Cells(j, colTanto).Value)
                                    data1(i, dcount).settei = str1
                                    'data1(i, dcount).seqno = timetbl * 100 + dcount
                                    str4 = UCase(str3)
                                    If str4 = "A" Then
                                        str4 = "A2"
                                    End If
                                    If str4 = "P" Then
                                        str4 = "P3"
                                    End If
                                    data1(i, dcount).strTime = str4
                                    data1(i, dcount).sortkey = str4 & CStr(1000 + dcount)
                                    data1(i, dcount).prnRowsCount = 1
                                    data1(i, dcount).prnData(1) = CStr(.Cells(j, colPrn).Value)
                                    data1(i, dcount).yajirusi = 0
                                    data1(i, dcount).yajirusi2 = 0
                                    data1(i, dcount).dmyline = 0
                                    For l = 1 To 19
                                        If CStr(.Cells(j + l, colMark).Value) = CST_MARK_ITEM Or _
                                            CStr(.Cells(j + l, colPrn).Value) = "" Then
                                            Exit For
                                        End If
                                        data1(i, dcount).prnRowsCount = l + 1
                                        data1(i, dcount).prnData(l + 1) = CStr(.Cells(j + l, colPrn).Value)
                                    Next l
                                End If
                            Next k
                        End If
                    End If
                Next j
                dcnt1(i) = dcount
                
                '�ҏW�f�[�^�����Ԗ��̋敪�Ń\�[�g
                For j = 1 To dcount
                    flg1 = True
                    For k = 2 To dcount
                        If data1(i, k - 1).sortkey > data1(i, k).sortkey Then
                            flg1 = False
                            tmpdata = data1(i, k - 1)
                            data1(i, k - 1) = data1(i, k)
                            data1(i, k) = tmpdata
                        End If
                    Next k
                    If flg1 Then
                        Exit For
                    End If
                Next j
                
                '�e���ԑі��̍ő�s�����擾
                Erase tmpNum
                For j = 1 To dcount
                    k = data1(i, j).timetbl
                    l = k
                    tmpNum(k) = tmpNum(k) + data1(i, j).prnRowsCount
                    If tmpNum(k) > timeNum(k) Then
                        timeNum(k) = tmpNum(k)
                    End If
                Next j
                
            End If
            
        Next i
    End With
    
    '���ԑі��̈󎚊J�n�s��ݒ�
    For i = 1 To 5
        k = 1
        timetbl = data1(i, 1).timetbl
        For j = 1 To dcnt1(i)
            If data1(i, j).timetbl <> timetbl Then
                timetbl = data1(i, j).timetbl
                k = 1
            End If
            data1(i, j).prnStartRow = k + data1(i, j).dmyline
            k = k + data1(i, j).prnRowsCount
        Next j
    Next i
            
    '�Ηj���ȍ~�́u���v�ɂ��Ĉ󎚊J�n�s�𒲐�
    For i = 2 To 5
        For j = 1 To dcnt1(i)
            If data1(i, j).settei = "��" Then
                Call checkYajirusi(dcnt1, data1, i, j)
            End If
        Next j
    Next i
     
    
    
    

    '���ԑі��̊J�n�s�������Z�b�g
    For i = 1 To 5
        Erase tmpNum
        For j = 1 To dcount
            k = data1(i, j).timetbl
            tmpNum(k) = tmpNum(k) + data1(i, j).prnRowsCount
            If tmpNum(k) > timeNum(k) Then
                timeNum(k) = tmpNum(k)
            End If
        Next j
    Next i
    timeRow(1) = 1
    For i = 2 To 10
        timeRow(i) = timeRow(i - 1) + timeNum(i - 1)
    Next i
    '����(a5)�̈ʒu�ɗ]�T����������A   a4�ȍ~�̈ʒu�������Ē�������
    k = Range(amRange).Rows.Count + 1 - (timeRow(5) + timeNum(5))
    If k > 0 Then
        For i = 4 To 10
            timeRow(i) = timeRow(i) + k
        Next i
    End If
    '16��(p4),17��(p5)�̈ʒu�𒲐�
    k = (Range(targetRange).Rows.Count) - (timeRow(10) + timeNum(10))
    If timeNum(10) = 0 Then
        k = k - 1
    End If
    If k > 0 Then
        timeRow(9) = timeRow(9) + k
        timeRow(10) = timeRow(10) + k
    End If

    '���ԑт̈�
    Erase timeStr
    timeStr(1) = "A1"
    timeStr(10) = "P9"
    For i = 1 To 5
        For j = 1 To dcnt1(i)
            k = cnvTimetbl(data1(i, j).strTime)
            If timeStr(k) = "" Then
                timeStr(k) = data1(i, j).strTime
            Else
                If data1(i, j).strTime < timeStr(k) Then
                    timeStr(k) = data1(i, j).strTime
                End If
            End If
        Next j
    Next i
    j = 1
    For i = 1 To 10
        If j < timeRow(i) Then
            j = timeRow(i)
        End If
        Select Case timeStr(i)
            Case "A1"
                str1 = "(�ߑO) 8:30�`"
            Case "A2"
                str1 = " 9:00���`"
            Case "A3"
                str1 = " 9:30���`"
            Case "A4"
                str1 = "10:00���`"
            Case "A5"
                str1 = "10:30���`"
            Case "A6"
                str1 = "11:00���`"
            Case "A7"
                str1 = "11:30���`"
            Case "A8"
                str1 = "12:00�`"
            Case "A9"
                str1 = "12:30�`"
            Case "P1"
                str1 = "(�ߌ�)13:00�`"
            Case "P2"
                str1 = "13:30���`"
            Case "P3"
                str1 = "14:00���`"
            Case "P4"
                str1 = "14:30���`"
            Case "P5"
                str1 = "15:00���`"
            Case "P6"
                str1 = "15:30���`"
            Case "P7"
                str1 = "16:00���`"
            Case "P8"
                str1 = "16:30���`"
            Case "P9"
                str1 = "17:00 �I��"
            Case Else
                str1 = ""
        End Select
        targetSheet.Cells(j + Range(targetRange).Row, 2).Value = str1
    Next i
    
'    str1 = "#" & CStr(weekNo)
'    If weekNo < 9 Then
'        str1 = "#" & CStr(weekNo)
'    Else
'        str1 = "#9"
'    End If
'    i = markedRow(str1)
'    If i > 0 Then
'        v = Split(CStr(ThisWorkbook.Sheets(CST_SHT_KOMOKU).Cells(i, colPrn).Value), ",")
'        l = 1
'        For j = 1 To UBound(v) + 1
'            If l < timeRow(j) Then
'                l = timeRow(j)
'            End If
'            targetSheet.Cells(l + Range(targetRange).Row, 2).Value = v(j - 1)
'        Next j
'    End If
           
    '��
    For i = 1 To 5
        l = 1
        For j = 1 To dcnt1(i)
            If l < timeRow(data1(i, j).timetbl) Then
                l = timeRow(data1(i, j).timetbl)
            End If
            l = l + data1(i, j).dmyline
            rc = data1(i, j).prnRowsCount - data1(i, j).dmyline
            r1 = Range(targetRange).Row + l
            r2 = r1 + rc - 1
            c1 = Range(targetRange).Column + (i * 3) - 2
            For k = 1 To rc
                If data1(i, j).settei = "��" Then
                    If k = 1 Then
                        'targetSheet.Cells(l + Range(targetRange).Row, (i * 3) + 1).Value = "��"
                    End If
                Else
                    str1 = data1(i, j).prnData(k)
                    If fnGakuseiSu() > 1 Then
                        '�w���Q���ȏ�̏ꍇ
                        If data1(i, j).fukusu <> "" Then
                            str1 = Replace(str1, "[@]", data1(i, j).fukusu)
                        End If
                    Else
                        '�w������l
                        str1 = Replace(str1, "[@]", "")
                    End If
                    If k = 1 And data1(i, j).tanto <> "" Then
                        str1 = str1 & "  (" & data1(i, j).tanto & ")"
                    End If
                    '�󎚓��e�Z�b�g
                    targetSheet.Cells(l + Range(targetRange).Row, (i * 3) + 1).Value = str1
                    If data1(i, j).settei = LCase(data1(i, j).settei) Then
                        '�J�����_���ɏ������œ��͂���Ă���ꍇ�̓C�^���b�N�̂Ƃ���
                        targetSheet.Cells(l + Range(targetRange).Row, (i * 3) + 1).Font.Italic = True
                    End If
                    '�S���҂�Ԋ|���\��
                    If k = 1 And data1(i, j).tanto <> "" Then
                        Call tanto_amikake(targetSheet, targetSheet.Cells(l + Range(targetRange).Row, (i * 3) + 1))
                    End If
                End If
                l = l + 1
            Next k
            If rc > 1 And data1(i, j).settei <> "��" Then
                '���J�b�R
                With targetSheet.Range(Cells(r1, c1), Cells(r2, c1))
                    targetSheet.Shapes.AddShape(Type:=msoShapeLeftBracket, Left:=.Left + 3, Top:=.Top + 3, Width:=.Width - 6, Height:=.Height - 6).Select
                    
                End With
            End If
            If data1(i, j).yajirusi > 0 Then
                '�E�J�b�R����
                With targetSheet.Range(Cells(r1, c1 + 2), Cells(r2, c1 + 2))
                    targetSheet.Shapes.AddShape(Type:=msoShapeRightBracket, Left:=.Left + 3, Top:=.Top + 3, Width:=.Width - 6, Height:=.Height - 6).Select
                End With
                With targetSheet.Range(Cells(r1, c1 + 3), Cells(r2, c1 + 3 + data1(i, j).yajirusi * 3 - 2))
                    'targetSheet.Shapes.AddShape(Type:=msoShapeRightArrow, Left:=.Left + 3, Top:=.Top + Int(.Height / 2) - 1, Width:=.Width - 6, Height:=0).Select
                    targetSheet.Shapes.AddConnector(Type:=msoConnectorStraight, beginx:=.Left + 3, beginy:=.Top + Int(.Height / 2), endx:=.Left + .Width - 3, endy:=.Top + Int(.Height / 2)).Select
                    Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadTriangle
                    Selection.ShapeRange.Line.Visible = msoTrue
                    Selection.ShapeRange.Line.Weight = CST_LINE_WEIGHT
                End With
                
            End If
        Next j
    Next i
    
    targetSheet.Range("A1").Select
    
End Sub

'�J�����_�����͒l���u���v�̏ꍇ�ɁA���Y�Z���̍����̗���l���擾���܂�
Private Function lookLeft(rowIdx As Long, colidx As Long) As String
    Dim str1                As String
    
    lookLeft = ""
    
    If colidx > Range(CST_RANGE_J_CAL_TOP).Column Then
        str1 = CStr(ThisWorkbook.Sheets(CST_SHT_KOMOKU).Cells(rowIdx, colidx - 1).Value)
        If str1 = "��" Then
            '�ċA�I�ɌĂяo��
            lookLeft = lookLeft(rowIdx, colidx - 1)
        Else
            lookLeft = str1
        End If
    End If
    
End Function

'���ԑы敪�i1����10�j�����ڒ�`�l�܂��̓J�����_�����͒l���ϊ����ĕԂ��܂�
Public Function cnvTimetbl(strSettei As String) As Long
    Dim str1                As String
    Dim tbl                 As Long
    
    cnvTimetbl = 0
    
    
    Select Case UCase(strSettei)
        Case "A1"
            tbl = 1
        Case "A2", "A3", "A"
            tbl = 2
        Case "A4", "A5"
            tbl = 3
        Case "A6", "A7"
            tbl = 4
        Case "A8", "A9"
            tbl = 5
        Case "P1", "P2"
            tbl = 6
        Case "P3", "P4", "P"
            tbl = 7
        Case "P5", "P6"
            tbl = 8
        Case "P7", "P8"
            tbl = 9
        Case "P9"
            tbl = 10
        Case Else
            Debug.Print "NG (" & strSettei & ")"
    End Select
    
    cnvTimetbl = tbl
            
End Function


'���̈ʒu����
Private Sub checkYajirusi(ByRef dcnt1() As Long, ByRef data1() As tpData, idx1 As Long, idx2 As Long)
    Dim i                   As Long
    Dim j                   As Long
    Dim k                   As Long
    
    Dim rowIdx              As Long
    Dim timetbl             As Long
    Dim pos                 As Long
    Dim minPos              As Long
    Dim maxPos              As Long
    Dim rIdx(5)             As Long
    Dim p1                  As Long
    Dim cnt                 As Long
    
    rowIdx = data1(idx1, idx2).rowIdx
    timetbl = data1(idx1, idx2).timetbl
    minPos = data1(idx1, idx2).prnStartRow
    maxPos = data1(idx1, idx2).prnStartRow
    
    '�͂��߂ɓ��Y���ڂɂ��Ĉ󎚊J�n�s�̂�����m�F
    For i = 1 To 5
        rIdx(i) = 0
        For j = 1 To dcnt1(i)
            If data1(i, j).rowIdx = rowIdx And data1(i, j).timetbl = timetbl Then
                '����f�[�^
                rIdx(i) = j
                pos = data1(i, j).prnStartRow
                If pos < minPos Then
                    minPos = pos
                End If
                If pos > maxPos Then
                    maxPos = pos
                End If
            End If
        Next j
    Next i
    
    '�e�q�֌W�̃Z�b�g
    For i = 1 To 5
        If rIdx(i) > 0 Then
            If data1(i, rIdx(i)).settei <> "��" Then
                '�e
                p1 = i
                cnt = 0
                data1(i, rIdx(i)).yajirusi = 0
            Else
                '�q
                cnt = cnt + 1
                data1(p1, rIdx(p1)).yajirusi = cnt
                data1(i, rIdx(i)).yajirusi2 = cnt
            End If
        End If
    Next i
    
    If minPos = maxPos Then
        '����͖���
        Exit Sub
    End If
    
    '�󎚊J�n�s��傫���l�ɍ��킹��
    For i = 1 To 5
        For j = 1 To dcnt1(i)
            If data1(i, j).rowIdx = rowIdx And data1(i, j).timetbl = timetbl Then
                '����f�[�^
                pos = data1(i, j).prnStartRow
                If pos < maxPos Then
                    data1(i, j).dmyline = maxPos - pos
                    data1(i, j).prnRowsCount = data1(i, j).prnRowsCount + (maxPos - pos)
                End If
            End If
        Next j
    Next i
    
    '���ԑі��̈󎚊J�n�s���Đݒ�
    For i = 1 To 5
        k = 1
        timetbl = data1(i, 1).timetbl
        For j = 1 To dcnt1(i)
            If data1(i, j).timetbl <> timetbl Then
                timetbl = data1(i, j).timetbl
                k = 1
            End If
            data1(i, j).prnStartRow = k + data1(i, j).dmyline
            k = k + data1(i, j).prnRowsCount
        Next j
    Next i
    
    
    
End Sub


'�S���ҕ\����Ԋ|���Ƃ���
'�Z�������̃e�L�X�g�̊Y�����𔒐F�i�e�[�}�J���[�F�P�j�Ƃ��A�Z�������ӏ��Ƀe�L�X�g�}�`�ŕ`�悷��
Private Sub tanto_amikake(sht1 As Worksheet, rng1 As Range)
    Dim str1            As String
    Dim str2            As String
    Dim i               As Long
    Dim p1              As Long
    Dim p2              As Long
    Dim w1              As Double
    Dim h1              As Double
    
    If CST_TANTO_SHAPE = False Then
        Exit Sub
    End If
    
    str1 = rng1.Value
    If str1 = "" Then
        Exit Sub
    End If
    
    '(�S����)��T��
    p1 = 0
    p2 = 0
    For i = 1 To Len(str1)
        If Mid(str1, i, 1) = "(" Then
            p1 = i
        ElseIf Mid(str1, i, 1) = ")" Then
            p2 = i
        End If
    Next i
    If p1 <> 0 And p2 <> 0 And p1 < p2 Then
        '�S���Ҕ���
        str2 = Mid(str1, p1, p2 - p1 + 1)
        
        '�S���҂̕\���F�𔒐F�֕ύX�i�����Ȃ�����j
        rng1.Characters(Start:=p1, Length:=Len(str2)).Font.ThemeColor = 1
        
        '�S���҂̃e�L�X�g�{�b�N�X��`��
        sht1.Shapes.AddTextbox(msoTextOrientationHorizontal, rng1.Left + rng1.Width - 200, rng1.Top, 200, rng1.Height).Select
        With Selection
            w1 = .Width
            h1 = .Height
            .ShapeRange.TextFrame2.TextRange.Characters.Text = str2
        End With
        With Selection.ShapeRange.TextFrame2
            .VerticalAnchor = msoAnchorMiddle
            .HorizontalAnchor = msoAnchorCenter
            .WordWrap = msoFalse
            .AutoSize = msoAutoSizeShapeToFitText
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
        End With
        With Selection.ShapeRange.TextFrame2.TextRange.Font
            .NameComplexScript = "���C���I"
            .NameFarEast = "���C���I"
            .Name = "���C���I"
            .Size = 8
        End With
        
        With Selection.ShapeRange.Fill
            .Visible = msoTrue
            .ForeColor.ObjectThemeColor = msoThemeColorAccent4
            .ForeColor.TintAndShade = 0
            .ForeColor.Brightness = 0.6000000238
            .Transparency = 0
            .Solid
        End With
        If Selection.Width < w1 Then
            Selection.ShapeRange.IncrementLeft w1 - Selection.Width
        End If
            
    End If
End Sub






