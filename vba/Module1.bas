Attribute VB_Name = "Module1"
'薬局実習スケジュール作表
'                                           H. Komatsu
'------------------------------------------------------------
'                                           create 2014.01.09
'
'2019.02.18) 担当者の表示をテキストボックスとし網掛けを行う
'2019.02.12) 「複数セル選択→実習の連続実施」入力機能追加
'2019.02.11) 残り当番のローテーション変更
'2019.02.04) 実習期間変更に対応
'2015.05.31) 作表対象（週）を選択可能とする
'

Const CST_SHT_KIHON         As String = "基本設定"
Const CST_SHT_KOMOKU        As String = "実習項目"
Const CST_SHT_TOPPAGE       As String = "表紙"
Const CST_SHT_SCHEDULE1     As String = "スケジュール1"
Const CST_SHT_SCHEDULE2     As String = "スケジュール2"

'基本設定シートのデータ領域
Const CST_RANGE_K_NENDO     As String = "C5"            '年度
Const CST_RANGE_K_KI        As String = "C6"            '期
Const CST_RANGE_K_KAISI     As String = "C7"            '期間（開始日）
Const CST_RANGE_K_MANRYO    As String = "C8"            '期間（満了日）
Const CST_RANGE_K_GAKUSEI   As String = "B12:F14"       '学生（実習生）
Const CST_RANGE_K_YAKUZAISI As String = "B20:F28"       '職員（薬剤師）
Const CST_RANGE_K_JIMU      As String = "B30:F38"       '職員（事務）
Const CST_RANGE_K_YOTEI     As String = "B43:F55"       '期間内の予定
Const CST_RANGE_K_MOKUHYO   As String = "J6:J55"        '週毎の目標（課題）

'実習項目シートのデータ領域
Const CST_RANGE_J_KAISI_MON As String = "D2"            '第１週の月曜日となる日をセットするセル
Const CST_RANGE_J_CAL_TOP   As String = "K1"            'カレンダー領域の左上（トップ）
Const CST_RANGE_J_CAL_DATE  As String = "K3:CP3"        'カレンダー日付欄
Const CST_RANGE_J_CHECK     As String = "A5"            'チェック欄
Const CST_RANGE_J_ITEM_TOP  As String = "B5"            '実習項目の左上（トップ）
Const CST_RANGE_J_KOMOKU    As String = "D5"            '実習項目（見出し）
Const CST_RANGE_J_PRINT     As String = "E5"            '印字内容欄
Const CST_RANGE_J_JIKAN     As String = "F5"            '時間帯
Const CST_RANGE_J_FUKUSU    As String = "G5"            '複数実習生
Const CST_RANGE_J_TANTO     As String = "H5"            '担当

'表紙シートのデータ領域
Const CST_RANGE_TP_TITLE    As String = "B3"            'タイトル欄（平成ＮＮ年度Ｎ期　薬局…スケジュール）
Const CST_RANGE_TP_KIKAN    As String = "B5"            '期間
Const CST_RANGE_TP_GAKUSEI  As String = "C9"            '実習生欄
Const CST_RANGE_TP_SIDO     As String = "F10"           '指導薬剤師欄
Const CST_RANGE_TP_SYOKUIN  As String = "C14"           'みなみ職員欄
Const CST_RANGE_TP_JIMU     As String = "F15"           '事務職員欄

'スケジュール１（１～８週）のデータ領域
Const CST_RANGE_S1_CREATE   As String = "P1"            '作成日
Const CST_RANGE_S1_TITLE    As String = "B2"            'タイトル欄
Const CST_RANGE_S1_MOKUHYO  As String = "B3:B5"         '目標（課題）欄
Const CST_RANGE_S1_TABLE    As String = "B7:Q39"        '表
Const CST_RANGE_S1_AM       As String = "B8:Q23"        '午前の領域
Const CST_RANGE_S1_PM       As String = "B24:Q39"       '午後の領域

'スケジュール２（9～11週）のデータ領域
Const CST_RANGE_S2_CREATE   As String = "P1"            '作成日
Const CST_RANGE_S2_TITLE    As String = "B2"            'タイトル欄
Const CST_RANGE_S2_MOKUHYO  As String = "B3:B5"         '目標（課題）欄
Const CST_RANGE_S2_TABLE1   As String = "B7:Q16"        '表(9週目)
Const CST_RANGE_S2_AM       As String = "B8:G12"        '午前の領域
Const CST_RANGE_S2_PM       As String = "B13:G16"       '午後の領域
Const CST_RANGE_S2_TABLE2   As String = "B19:Q28"       '表(10週目)
Const CST_RANGE_S2_TABLE3   As String = "B31:Q40"       '表(11週目)

'実習項目シートにおけるデータ行の識別記号
Const CST_MARK_NOKORI_TOBAN As String = "※1"           '残り当番
Const CST_MARK_NOKORI_SUB   As String = "※2"           '残り当番（サブ）
Const CST_MARK_KYUJITU      As String = "☆"            '休日
Const CST_MARK_ITEM         As String = "□"            '実習項目

'メニューコマンド
Const CST_MENU_1            As String = "実習を連続実施" '右クリックした際に表示するメニュー

Const CST_LINE_WEIGHT       As Single = 0.75            '→（コネクト図形）の線の太さ(0.25pt/0.5pt/0.75pt/1.0pt/1.25pt/1.5pt/1.75pt/2.0pt等)
Const CST_TANTO_SHAPE       As Boolean = True           '担当者の表示を図形描画(網掛け)とするかどうか

'編集データ
Type tpData
    rowIdx                  As Long                     '行index
    title                   As String                   '項目タイトル
    prnData(20)             As String                   '印字内容
    timetbl                 As Long                     '時間割(午前:1～5,午後6～10)
    strTime                 As String                   '時間割（文字列）
    fukusu                  As String                   '複数実習生
    tanto                   As String                   '担当
    settei                  As String                   'カレンダ欄入力値
    seqno                   As Long                     'ソート用
    sortkey                 As String                   'ソート用
    prnStartRow             As Long                     '印刷開始行
    prnRowsCount            As Long                     '印刷行数（明細数）
    yajirusi                As Long                     '→の長さ（親にセット）
    yajirusi2               As Long                     '→の長さ（子にセット）
    dmyline                 As Long                     '位置合わせの為の空行数
End Type

    
'メッセージ表示
Public Sub dispMessage(strMessage)
    With UserForm1
        .setMessage (strMessage)
        .Caption = ThisWorkbook.Name
        .Show vbModeless
        .Repaint
    End With
End Sub

'メッセージ消去
Public Sub hideMessage()
    Unload UserForm1
End Sub

'基本設定シートの変更
Public Sub wSheetChage_Kihon(ByVal target As Range)
    Dim cell1       As Range
    
    If target.Columns.Count > 100 Or target.Rows.Count > 100 Then
        '行選択などは処理しない
        Exit Sub
    End If
    
    For Each cell1 In target.Cells
        Call cellChange_Kihon(cell1)
    Next
End Sub

'実習項目シートの変更
Public Sub wSheetChage_Jissyu(ByVal target As Range)
    Dim cell1       As Range
    
    If target.Columns.Count > 100 Or target.Rows.Count > 100 Then
        '行選択などは処理しない
        Exit Sub
    End If
    
    For Each cell1 In target.Cells
        Call cellChange_Jissyu(cell1)
    Next
End Sub

'実習項目シートのセレクション変更
Public Sub wSheetSelectionChage_Jissyu(ByVal target As Range)
    Dim obj1        As Object
    
    '右クリックメニューの独自メニューを削除
    On Error Resume Next
    Application.CommandBars("Cell").Controls(CST_MENU_1).Delete
    
    If target.Cells.Count = 1 Then
        Call selectionChange_Jissyu(target)
    Else
        If target.Cells.Count > 1 And target.Areas.Count = 1 And target.Rows.Count = 1 Then
            '単一行における横方向複数セルの選択→連続実施入力
            Set obj1 = Application.CommandBars("Cell").Controls.Add()
            With obj1
                .Caption = CST_MENU_1
                .OnAction = "renzoku_jissi"
                .BeginGroup = False
            End With
        End If
    End If
End Sub

'実習の連続実施（右クリックメニュー）
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
        MsgBox ("実習項目シートを選択してください")
        Exit Sub
    End If
    
    Set rng1 = Selection
    
    If rng1.Cells.Count < 2 Or rng1.Areas.Count <> 1 Or rng1.Rows.Count <> 1 Then
        Set rng1 = Nothing
        Exit Sub
    End If
    
    If rng1.Cells(1).Column < Range(CST_RANGE_J_CAL_TOP).Column Or _
       rng1.Cells(1).Row < Range(CST_RANGE_J_ITEM_TOP).Row Then
        MsgBox ("選択範囲が不正です。")
        Exit Sub
    End If
    
    c = Range(CST_RANGE_J_JIKAN).Column                 '時間帯の桁
    r1 = rng1.Cells(1).Row                              '入力対象の行（選択行）
    r2 = Range(CST_RANGE_J_CAL_DATE).Row                'カレンダの行
    r3 = fnGetKujitsuRow                                '休日の行
    
    If r3 = 0 Then
        MsgBox ("休日のデータ領域が見つかりません")
        Exit Sub
    End If
    
    '時間帯
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
            '実習予定の入力領域
            d1 = ActiveSheet.Cells(r2, c).Value
            If Weekday(d1) = 1 Or Weekday(d1) = 7 Or ActiveSheet.Cells(r3, c) <> "" Then
                '日曜・土曜・休日
                str2 = ""
            Else
                If d1 = fnKaisi Or Weekday(d1) = 2 Then
                    '実習開始日または月曜日
                    str2 = str1
                Else
                    If ActiveSheet.Cells(r1, c - 1).Value = "" Then
                        str2 = str1
                    Else
                        str2 = "→"
                    End If
                End If
            End If
            ActiveSheet.Cells(r1, c).Value = str2
        End If
    Next i
    
    
    
    Set rng1 = Nothing
    
End Sub


'休日行の取得
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



'基本設定シートのセル変更
Private Sub cellChange_Kihon(ByVal target As Range)
    Dim r1                  As Long
    Dim c1                  As Long
    Dim i                   As Long

    '入力セルのアドレスをチェックして、入力された項目に応じて処理を行う
    
    'イベントを無効（連鎖の抑止）
    Application.EnableEvents = False
    
    
    '期間開始日の入力→カレンダー変更
    If target.Address(False, False) = Range(CST_RANGE_K_KAISI).Address(False, False) Then
        Call changeKaisi(target)
        GoTo term_change_kihon
    End If
    
    '期間満了日の入力→カレンダー変更
    If target.Address(False, False) = Range(CST_RANGE_K_MANRYO).Address(False, False) Then
        Call changeManryo(target)
        GoTo term_change_kihon
    End If
    
    
    '週毎の目標（課題）タイトル
    r1 = Range(CST_RANGE_K_MOKUHYO).Row
    c1 = Range(CST_RANGE_K_MOKUHYO).Column
    For i = 1 To 10
        If target.Row = r1 And target.Column = c1 Then
            Call changeMokuhyo(i)
        End If
        r1 = r1 + 5
    Next i
    
    '予定
    If target.Row >= Range(CST_RANGE_K_YOTEI).Row And _
        target.Row <= Range(CST_RANGE_K_YOTEI).Row + Range(CST_RANGE_K_YOTEI).Rows.Count And _
        target.Column >= Range(CST_RANGE_K_YOTEI).Column And _
        target.Column <= Range(CST_RANGE_K_YOTEI).Column + Range(CST_RANGE_K_YOTEI).Columns.Count Then
        Call setYotei2Comment
    End If
        
        
    

term_change_kihon:

    'イベント有効へ戻す
    Application.EnableEvents = True

End Sub


'実習項目シートのセル変更
Private Sub cellChange_Jissyu(ByVal target As Range)
    Dim r1                  As Long
    Dim c1                  As Long
    Dim i                   As Long

    'イベントを無効（連鎖の抑止）
    Application.EnableEvents = False
    
    '対象外か？（見出し部分）
    If target.Row < Range(CST_RANGE_J_ITEM_TOP).Row Then
        GoTo term_change_Jissyu
    End If
    
    '実習項目、タイトル変更
    If target.Column >= Range(CST_RANGE_J_ITEM_TOP).Column And _
        target.Column <= Range(CST_RANGE_J_ITEM_TOP).Column + 3 Then
        Call changeJissyuKomoku(target.Row)
        GoTo term_change_Jissyu
    End If
    
    'カレンダ部
    If target.Column >= Range(CST_RANGE_J_CAL_DATE).Column And _
        target.Column <= Range(CST_RANGE_J_CAL_DATE).Column + Range(CST_RANGE_J_CAL_DATE).Columns.Count - 1 Then
        Call changeCalendar(target.Row, target.Column)
        GoTo term_change_Jissyu
    End If

term_change_Jissyu:

    'イベント有効へ戻す
    Application.EnableEvents = True

End Sub


'実習項目シートのセレクション変更（サブ）
Private Sub selectionChange_Jissyu(ByVal target As Range)
    Dim colMark             As Long                     'マーク列index
    Dim colTitle            As Long                     '実習項目（タイトル）列index
    Dim colPrn              As Long                     '印字内容
    Dim colTimetbl          As Long                     '時間割区分
    Dim colTanto            As Long                     '担当
    Dim colFukusu           As Long                     '複数実習生
    Dim rowDate             As Long                     'カレンダの日付行
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
    
    'イベントを無効（連鎖の抑止）
    Application.EnableEvents = False

    '対象外か？（日付欄か）
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
                        kyujitu = "☆休日"
                    Case CST_MARK_NOKORI_TOBAN
                        nokori = str1
                    Case CST_MARK_NOKORI_SUB
                        nokoriSub = str1
                    Case CST_MARK_ITEM
                        If str1 = "→" Then
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
                str1 = str1 & "□　" & data1(i).title & vbLf
                For j = 1 To data1(i).prnRowsCount
                    str3 = data1(i).prnData(j)
                    If j = 1 And data1(i).tanto <> "" Then
                        str3 = str3 & "  【" & data1(i).tanto & "】"
                    End If
                    If j = 1 Then
                        str3 = "・" & str3
                    Else
                        str3 = "　" & str3
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
        Call UserForm2.setValues(StrConv(Format(date1, "m月d日(aaa)"), vbWide), nokori & "/" & nokoriSub, str1, str2)
        UserForm2.Show vbModal
    End If




term_selectionChange_Jissyu:

    'イベント有効へ戻す
    Application.EnableEvents = True
    
End Sub

'週毎の目標（課題）タイトルが変更された
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


'実習期間の開始日が変更された→カレンダー設定
Private Sub changeKaisi(target As Range)
    Dim day1                As Long
    Dim day2                As Long
    Dim day3                As Long
    Dim strNendo            As String
    Dim strKi               As String
    
    If IsDate(target.Value) = False Then
        MsgBox ("ただしい日付を入力して下さい")
        Exit Sub
    End If
    
    'Call dispMessage("処理中です...")
    
    day1 = target.Value
    
    '年度、期を算出
    '------------------------------------------------------------
    ' (2018年度まで→第1期：5月～7月、第2期：9月～11月、第3期：1月～3月)
    ' (2019年度から→第1期：2月～5月、第2期：5月～8月、第3期：8月～11月、第4期：11月～2月)
    '
    If day1 < CDate("2019/01/31") Then
        '<<<2018年度以前>>>
        If Month(day1) < 4 Then
            day2 = CDate(CStr(Year(day1) - 1) & "/4/1")
            strKi = "第３期"
        Else
            day2 = CDate(CStr(Year(day1) & "/4/1"))
            If Month(day1) > 7 Then
                strKi = "第２期"
            Else
                strKi = "第１期"
            End If
        End If
    Else
        '<<<2019年度以降>>>
        Select Case True
            Case Month(day1) = 1
                day2 = CDate(CStr(Year(day1) - 1) & "/4/1")
                strKi = "第４期"
            Case Month(day1) < 5
                day2 = CDate(CStr(Year(day1)) & "/4/1")
                strKi = "第１期"
            Case Month(day1) < 8
                day2 = CDate(CStr(Year(day1)) & "/4/1")
                strKi = "第２期"
            Case Month(day1) < 11
                day2 = CDate(CStr(Year(day1)) & "/4/1")
                strKi = "第３期"
            Case Else
                day2 = CDate(CStr(Year(day1)) & "/4/1")
                strKi = "第４期"
        End Select
    End If
    strNendo = Format(day2, "ggge年度")
    ThisWorkbook.Sheets(CST_SHT_KIHON).Range(CST_RANGE_K_NENDO).Value = strNendo
    ThisWorkbook.Sheets(CST_SHT_KIHON).Range(CST_RANGE_K_KI).Value = strKi
    
    
    
    'カレンダの開始日、期間満了日を算出する（祭日などで月曜日以外からの開始もあり得るため）
    day2 = day1 - ((Weekday(day1) + 5) Mod 7)           'WeekDay関数で月曜日は「2」
    day3 = day2 + (7 * 11) - 3
    MsgBox ("期間満了日を「" & Format(day3, "yyyy/m/d") & "」とします。" & vbLf & "日付が異なる場合は修正して下さい。")
    ThisWorkbook.Sheets(CST_SHT_KIHON).Range(CST_RANGE_K_MANRYO).Value = day3
    
    'カレンダ日付のセット
    Call setDate2Calendar
    
    'カレンダー上の予定をチェック
    Call setYotei2Comment
    
    'Call hideMessage
    
End Sub


'実習期間の満了日が変更された→カレンダー設定
Private Sub changeManryo(target As Range)
    
    If IsDate(target.Value) = False Then
        MsgBox ("ただしい日付を入力して下さい")
        Exit Sub
    End If
    
    'カレンダ日付のセット
    Call setDate2Calendar
    
    'カレンダー上の予定をチェック
    Call setYotei2Comment
    
    
End Sub


'カレンダ日付のセット、休日チェック
Private Sub setDate2Calendar()
    Dim date1               As Long         '期間開始日
    Dim date2               As Long         '期間満了日
    Dim date3               As Long         '１週目の月曜（カレンダの先頭）
    Dim i                   As Long
    Dim c                   As Long
    Dim r                   As Long
    Dim rKyujitu            As Long
    Dim flg1                As Boolean      '実習対象日（期間外、土日・休日→False）
    
    date1 = fnKaisi()
    date2 = fnManryo()
    date3 = fn1stMonday()
    
    If date1 > date2 Then
        MsgBox ("実習期間の入力が不正です。")
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
                    '実習対象日
                    .Pattern = xlNone
                    .PatternColorIndex = xlAutomatic
                    '.PatternTintAndShade = 0
                Else
                    '期間外または土日・休日
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


'実習項目の変更
Private Sub changeJissyuKomoku(rowIdx As Long)
    Call checkJissyuKomoku(rowIdx)
End Sub

'カレンダ部の変更
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
            '休日の行
            flg1 = False
            If CStr(.Cells(rowIdx, columnIdx).Value) <> "" Then
                '休日
                flg1 = True
            End If
            If Weekday(date1) = 1 Or Weekday(date1) = 7 Then
                flg1 = True
            End If
            With .Cells(rowDate, columnIdx).Interior
                If flg1 Then
                    '休日または土日
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .ThemeColor = xlThemeColorDark1
                    .TintAndShade = -0.14996795556505
                    .PatternTintAndShade = 0
                Else
                    '休日以外
                    .Pattern = xlNone
                    .PatternColorIndex = xlAutomatic
                    '.PatternTintAndShade = 0
                End If
            End With
        End If
    End With

End Sub


'指定行の実習項目についてスケジュール設定の有無をチェックする
Private Sub checkJissyuKomoku(rowIdx As Long)
    Dim colCheck            As Long                     'チェック欄
    Dim colMark             As Long                     'マーク（項目識別）欄
    Dim colTitle            As Long                     'タイトル欄
    Dim colPrn              As Long                     '印字内容欄
    Dim colCalendar1        As Long                     'カレンダ欄の開始列
    Dim colCalendar2        As Long                     'カレンダ欄の終端列
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
                        str1 = str1 & Mid("月火水木金土日", j, 1)
                    End If
                    c = c + 1
                Next j
                If str1 <> "" Then
                    strDate = strDate & "第" & CStr(i) & "週(" & str1 & ") "
                End If
            Next i
            If Len(strDate) > 50 Then
                strDate = Left(strDate, 50) & "..."
            End If
            If strDate = "" Then
                .Cells(rowIdx, colCheck).Value = "未"
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



'イベント予定などをカレンダーの日付欄のコメントへセット
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
        str1 = CStr(ThisWorkbook.Sheets(CST_SHT_KIHON).Cells(r + i - 1, c))     '予定（表題）
        str2 = CStr(ThisWorkbook.Sheets(CST_SHT_KIHON).Cells(r + i - 1, c + 2)) '日付
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


'残り当番のクリア
Public Sub clearNokori()
    Dim r                   As Long
    Dim c                   As Long
    
    If MsgBox("残り当番・サブの入力を全て消去します。" & vbLf & "よろしいですか？", vbOKCancel) = vbCancel Then
        Exit Sub
    End If
    
    Application.EnableEvents = False
    
    c = Range(CST_RANGE_J_CAL_DATE).Column
    c2 = c + (7 * 12) - 1
    
    '残りクリア
    r = markedRow(CST_MARK_NOKORI_TOBAN)
    If r > 0 Then
        ThisWorkbook.Sheets(CST_SHT_KOMOKU).Range(Cells(r, c), Cells(r, c2)).Select
        Selection.ClearContents
        ThisWorkbook.Sheets(CST_SHT_KOMOKU).Cells(r, c).Select
    End If
    
    'サブクリア
    r = markedRow(CST_MARK_NOKORI_SUB)
    If r > 0 Then
        ThisWorkbook.Sheets(CST_SHT_KOMOKU).Range(Cells(r, c), Cells(r, c2)).Select
        Selection.ClearContents
        ThisWorkbook.Sheets(CST_SHT_KOMOKU).Cells(r, c).Select
    End If
    
    Application.EnableEvents = True
    
End Sub

'残り当番のセット
'2019.02.11 残り当番が５人の場合は休日に拘わらず週毎にシフトする
' 例）月:1 火:2 水:3 木:4 金:5   月:2 火:3 水:4 木:5 金:1   月:3 火:4 水:5 木:1 金:2
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
    
    '残り当番の設定を取得
    r1 = Range(CST_RANGE_K_YAKUZAISI).Row
    c1 = Range(CST_RANGE_K_YAKUZAISI).Column
    cnt = 0
    For i = 1 To 9
        str1 = CStr(ThisWorkbook.Sheets(CST_SHT_KIHON).Cells(r1 + i - 1, c1 + 3).Value)     '残り当番順
        str2 = CStr(ThisWorkbook.Sheets(CST_SHT_KIHON).Cells(r1 + i - 1, c1 + 4).Value)     '担当略名
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
        '当番の設定が無い
        Exit Sub
    End If
    
    'カレンダ探索
    date1 = fnKaisi()           '期間（開始日）
    date2 = fnManryo()          '期間（満了日）
    If date2 < date1 Then
        Exit Sub
    End If
    
    If cnt <> 5 Then
        '<<<<<営業日毎に順番のローテーション(従前)>>>>>>
        For i = 1 To 2
            If i = 1 Then
                r1 = markedRow(CST_MARK_NOKORI_TOBAN)
            Else
                r1 = markedRow(CST_MARK_NOKORI_SUB)
            End If
            r2 = markedRow(CST_MARK_KYUJITU)    '休日の行index
            x = 0                               '前日（前営業日）の残り当番index
            c1 = date2ColumnIdx(date1)          '設定開始列のindex
            For j = date1 To date2
                str1 = CStr(ThisWorkbook.Sheets(CST_SHT_KOMOKU).Cells(r1, c1).Value)
                str2 = CStr(ThisWorkbook.Sheets(CST_SHT_KOMOKU).Cells(r2, c1).Value)
                If Weekday(j) <> 1 And Weekday(j) <> 7 And str2 = "" Then
                    '土・日・休日以外
                    If str1 <> "" Then
                        '入力済み→当番indexの取得
                        For k = 1 To cnt
                            If str1 = arNokori(k) Then
                                x = k
                                Exit For
                            End If
                        Next k
                    Else
                        If x > 0 Then
                            '次の当番をセット
                            x = (x Mod cnt) + 1
                            ThisWorkbook.Sheets(CST_SHT_KOMOKU).Cells(r1, c1).Value = arNokori(x)
                        End If
                    End If
                End If
                c1 = c1 + 1
            Next j
        Next i
    Else
        '<<<<<５人体制のローテーション(週毎にシフトし休日は考慮しない)>>>>>>
        For i = 1 To 2
            If i = 1 Then
                r1 = markedRow(CST_MARK_NOKORI_TOBAN)
            Else
                r1 = markedRow(CST_MARK_NOKORI_SUB)
            End If
            r2 = markedRow(CST_MARK_KYUJITU)    '休日の行index
            x = 0                               '前日（前営業日）の残り当番index
            c1 = date2ColumnIdx(date1)          '設定開始列のindex
            For j = date1 To date2
                str1 = CStr(ThisWorkbook.Sheets(CST_SHT_KOMOKU).Cells(r1, c1).Value)    '残り当番
                str2 = CStr(ThisWorkbook.Sheets(CST_SHT_KOMOKU).Cells(r2, c1).Value)    '休日
                If Weekday(j) = 1 Then
                    weekcnt = weekcnt + 1
                End If
                If Weekday(j) <> 1 And Weekday(j) <> 7 And str2 = "" Then
                    '土・日・休日以外
                    If str1 <> "" Then
                        '入力済み→当番indexの取得
                        For k = 1 To cnt
                            If str1 = arNokori(k) Then
                                x = k
                                weekcnt = x + 9 - Weekday(j)
                                Exit For
                            End If
                        Next k
                    Else
                        If x > 0 Then
                            '次の当番をセット
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
'スケジュール作表
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
        MsgBox ("実習期間が不正です")
        Exit Sub
    End If


'    '作表処理開始の確認
'    If MsgBox("スケジュール表を作成します。" & vbLf & "よろしいですか？", vbOKCancel) = vbCancel Then
'        Exit Sub
'    End If
    
    '作表対象の選択
    UserForm3.Show
    If UserForm3.Tag = "" Then
        Unload UserForm3
        Exit Sub
    End If
    targetWeek = UserForm3.selctionData
    Unload UserForm3
    
    Call dispMessage("処理中です...")
    
    '新規ブックを生成し、表紙及びスケジュールの各シートを本ブックよりコピーし、内容を設定します
    sht = 0
    If targetWeek(0) = True Then
        ThisWorkbook.Sheets(CST_SHT_TOPPAGE).Copy
        Set newWorkBook = ActiveWorkbook
        ActiveSheet.Name = "表紙"
        '表紙シートの内容セット
        Call setTopPage(ActiveSheet)
        sht = 1
    End If
    For i = 1 To weeks
        If i <= 8 Then
            If targetWeek(i) = True Then
                'スケジュール１（第１週～第８週）の内容セット
                If sht = 0 Then
                    ThisWorkbook.Sheets(CST_SHT_SCHEDULE1).Copy
                    Set newWorkBook = ActiveWorkbook
                Else
                    ThisWorkbook.Sheets(CST_SHT_SCHEDULE1).Copy after:=newWorkBook.Sheets(sht)
                End If
                ActiveSheet.Name = "第" & CStr(i) & "週目"
                sht = sht + 1
                Call setSchedule1(ActiveSheet, i)
            End If
        End If
        If i >= 9 And (i Mod 3) = 0 Then
            If targetWeek(i) = True Then
                'スケジュール２（第9週～）の内容セット
                If sht = 0 Then
                    ThisWorkbook.Sheets(CST_SHT_SCHEDULE2).Copy
                    Set newWorkBook = ActiveWorkbook
                Else
                    ThisWorkbook.Sheets(CST_SHT_SCHEDULE2).Copy after:=newWorkBook.Sheets(sht)
                End If
                If i + 3 < weeks Then
                    ActiveSheet.Name = "第" & CStr(i) & "週目から第" & CStr(i + 2) & "週目"
                Else
                    ActiveSheet.Name = "第" & CStr(i) & "週目以降"
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


'表紙シートの内容セット
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
        '実習年度・期・期間
        strNendo = CStr(.Range(CST_RANGE_K_NENDO).Value)
        strKi = CStr(.Range(CST_RANGE_K_KI).Value)
        dateKaisi = .Range(CST_RANGE_K_KAISI).Value
        dateManryo = .Range(CST_RANGE_K_MANRYO).Value
        nissu = nannichime(dateManryo)
        targetSheet.Range(CST_RANGE_TP_TITLE).Value = strNendo & strKi & " 薬局実習生　スケジュール"
        targetSheet.Range(CST_RANGE_TP_KIKAN).Value = "＜" & Format(dateKaisi, "ggge年m月d日（aaaa）") & _
                                                    " ～ " & Format(dateManryo, "ggge年m月d日（aaaa）") & _
                                                    "：計" & CStr(nissu) & "日間＞"
        '実習生
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
                str2 = "担当薬剤師：" & CStr(.Cells(r1 + i - 1, c1 + 4).Value)
                targetSheet.Cells(r2 + i, c2).Value = str1
                targetSheet.Cells(r2 + i, c3).Value = str2
            End If
        Next i
        targetSheet.Cells(r2, c2) = "◆実習生（" & StrConv(CStr(cnt), vbWide) & "名）"
        
        'みなみ職員
        cnt = 0
        r1 = Range(CST_RANGE_K_YAKUZAISI).Row
        c1 = Range(CST_RANGE_K_YAKUZAISI).Column
        r2 = Range(CST_RANGE_TP_SYOKUIN).Row
        c2 = Range(CST_RANGE_TP_SYOKUIN).Column
        '薬剤師
        For i = 1 To 9
            str1 = CStr(.Cells(r1 + i - 1, c1).Value)
            If str1 <> "" Then
                cnt = cnt + 1
                str1 = str1 & StrConv(" (" & CStr(.Cells(r1 + i - 1, c1 + 1).Value) & ")", vbNarrow)
                str2 = CStr(.Cells(r1 + i - 1, c1 + 2).Value)
                If str2 <> "" Then
                    str1 = str1 & " ： " & str2
                End If
                targetSheet.Cells(r2 + i + 1, c2).Value = str1
            End If
        Next i
        targetSheet.Cells(r2 + 1, c2).Value = "・薬剤師（" & StrConv(CStr(cnt), vbWide) & "名）"
        '事務
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
                    str1 = str1 & " ： " & str2
                End If
                targetSheet.Cells(r2 + i + 1, c2).Value = str1
            End If
        Next i
        targetSheet.Cells(r2 + 1, c2).Value = "・事務（" & StrConv(CStr(cnt2), vbWide) & "名）"
        targetSheet.Range(CST_RANGE_TP_SYOKUIN).Value = "◆秋田みなみ会営薬局職員（" & StrConv(CStr(cnt + cnt2), vbWide) & "名）"
        
    
    End With

End Sub


'指定の日付が実習開始日から何日目にあたるかを返す
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
        '実習期間外
        Exit Function
    End If
    
    cnt = 0
    
    'はじめに休日データの行を探索
    i = markedRow(CST_MARK_KYUJITU)
    If i > 0 Then
        With ThisWorkbook.Sheets(CST_SHT_KOMOKU)
            cx = date2ColumnIdx(dateKaisi)
            For j = dateKaisi To jdate
                '休日、土曜日、日曜日以外の日数カウント
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


'実習項目シートにおけるマーク（データ識別）列を探索し、該当する行indexを返す
'見つからない場合は「0」を返す
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


'カレンダにおいて、指定日の列(index)を返す
Private Function date2ColumnIdx(jdate As Long) As Long
    Dim dateKaisi           As Long
    
    date2ColumnIdx = 0
    
    dateKaisi = fn1stMonday()
    If jdate < dateKaisi Or jdate > fnManryo() Then
        'カレンダ範囲外
        Exit Function
    End If
    
    date2ColumnIdx = (jdate - dateKaisi) + Range(CST_RANGE_J_CAL_TOP).Column
    
End Function


'実習生の人数
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

'N番目の学生さん（氏名＋「さん」）
Private Function fnGakuseiSan(no As Long) As String
    If no = 1 Or no = 2 Or no = 3 Then
        fnGakuseiSan = " (" & ThisWorkbook.Sheets(CST_SHT_KIHON).Cells(Range(CST_RANGE_K_GAKUSEI).Row + no - 1, Range(CST_RANGE_K_GAKUSEI).Column).Value & " さん)"
    Else
        fnGakuseiSan = ""
    End If
End Function


'実習期間の開始日
Private Function fnKaisi() As Long
    fnKaisi = CLng(ThisWorkbook.Sheets(CST_SHT_KIHON).Range(CST_RANGE_K_KAISI).Value)
End Function


'実習期間の満了日
Private Function fnManryo() As Long
    fnManryo = CLng(ThisWorkbook.Sheets(CST_SHT_KIHON).Range(CST_RANGE_K_MANRYO).Value)
End Function

'実習期間は何週間目まであるかを返す
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


'カレンダーの先頭日（開始日を含む週の月曜日）
Public Function fn1stMonday() As Long
    Dim d                   As Long
    
    d = fnKaisi()
    fn1stMonday = d - ((Weekday(d) + 5) Mod 7)  'WeekDay関数で月曜日は「2」
End Function



'スケジュールシート１（第１週～第８週）の内容セット
Private Sub setSchedule1(targetSheet As Worksheet, weekNo As Long)
    Dim str1                As String
    Dim r1                  As Long
    Dim r2                  As Long
    Dim c1                  As Long
    Dim c2                  As Long
    Dim i                   As Long
    
    'タイトル・課題（目標）
    r1 = Range(CST_RANGE_K_MOKUHYO).Row + (weekNo - 1) * 5
    c1 = Range(CST_RANGE_K_MOKUHYO).Column
    With ThisWorkbook.Sheets(CST_SHT_KIHON)
        'タイトル
        targetSheet.Range(CST_RANGE_S1_TITLE).Value = .Cells(r1, c1).Value
        '課題（目標）
        r2 = Range(CST_RANGE_S1_MOKUHYO).Row
        c2 = Range(CST_RANGE_S1_MOKUHYO).Column
        For i = 1 To 4
            targetSheet.Cells(r2 + i - 1, c2).Value = .Cells(r1 + i, c1).Value
        Next i
    End With
    
    '日付・残り当番
    targetSheet.Range(CST_RANGE_S1_CREATE).Value = Format(Now(), "(ge.m.d作成)")
    Call setDateNokori(targetSheet, CST_RANGE_S1_TABLE, weekNo)
    
    'スケジュール内容
    Call setScheduleData(targetSheet, CST_RANGE_S1_TABLE, CST_RANGE_S1_AM, weekNo)
    
    
End Sub

'スケジュールシート２（第9週～）の内容セット（３週間／シート）
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
    
    
    'タイトル・課題（目標）
    If weekNo <= 9 Then
        rw = 9
    Else
        rw = 10
    End If
    r1 = Range(CST_RANGE_K_MOKUHYO).Row + (rw - 1) * 5
    c1 = Range(CST_RANGE_K_MOKUHYO).Column
    With ThisWorkbook.Sheets(CST_SHT_KIHON)
        'タイトル
        targetSheet.Range(CST_RANGE_S2_TITLE).Value = .Cells(r1, c1).Value
        '課題（目標）
        r2 = Range(CST_RANGE_S2_MOKUHYO).Row
        c2 = Range(CST_RANGE_S2_MOKUHYO).Column
        For i = 1 To 4
            targetSheet.Cells(r2 + i - 1, c2).Value = .Cells(r1 + i, c1).Value
        Next i
    End With
    
    weeks = fnWeeks()
    
    '日付・残り当番,スケジュール内容
    targetSheet.Range(CST_RANGE_S2_CREATE).Value = Format(Now(), "(ge.m.d作成)")
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

'スケジュールシートへの日付・残り当番セット
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
    
    
    '日付・残り当番
    With ThisWorkbook.Sheets(CST_SHT_KOMOKU)
        r1 = Range(targetRange).Row
        c1 = Range(targetRange).Column
        targetSheet.Cells(r1, c1).Value = " " & StrConv(CStr(weekNo), vbWide) & "週目"
        c1 = c1 + 2
        rKyujitu = markedRow(CST_MARK_KYUJITU)          '休日の行index
        rNokori = markedRow(CST_MARK_NOKORI_TOBAN)      '残り当番の行index
        rNokoriSub = markedRow(CST_MARK_NOKORI_SUB)     '残り当番（サブ）の行index
        date1 = fn1stMonday() + (weekNo - 1) * 7
        date2 = date1 + 4
        If date2 > fnManryo() Then
            date2 = fnManryo()
        End If
        c2 = date2ColumnIdx(date1)
        For i = date1 To date2
            str1 = StrConv(Format(i, "m/d(aaa)"), vbWide)
            If CStr(.Cells(rKyujitu, c2).Value) <> "" Then
                '休日
                str1 = str1 & "：休日"
            Else
                nisu = nannichime(i)                    '何日目か
                str1 = str1 & "：" & StrConv(CStr(nisu), vbWide) & "日目"
                strNokori = CStr(.Cells(rNokori, c2).Value)
                strNokoriSub = CStr(.Cells(rNokoriSub, c2).Value)
                If strNokori <> "" Then
                    '残り当番
                    str1 = str1 & "　" & strNokori
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


'スケジュール内容のセット
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
    Dim colMark             As Long             '項目を識別するマークの列index
    Dim colKomoku           As Long             '実習項目（見出し）の列index
    Dim colPrn              As Long             '印字内容の列index
    Dim colJikan            As Long             '時間帯の列index
    Dim colFukusu           As Long             '複数
    Dim colTanto            As Long             '担当
    Dim v1                  As Variant
    Dim timetbl             As Long
    Dim rc                  As Long
    
    
    
    Dim dcount              As Long
    Dim data1(5, 100)       As tpData
    Dim dcnt1(5)            As Long
    Dim tmpdata             As tpData
    Dim flg1                As Boolean
    Dim timeNum(10)         As Long                 '時間帯毎の印字行数
    Dim timeRow(10)         As Long                 '時間帯毎の印字開始行index
    Dim tmpNum(10)          As Long
    Dim timeStr(10)         As String               '時間表示
    
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
        
        '編集データ領域クリア
        Erase data1
        Erase timeNum
        
        '日毎にデータを編集領域へセット
        For i = 1 To 5
            d = date1 + i - 1
            
            dcount = 0
            
            If d < fnKaisi() Or d > fnManryo() Then
                '対象外
            Else
                colidx = date2ColumnIdx(d)
                r1 = Range(CST_RANGE_J_ITEM_TOP).Row
                'カレンダ部を探索
                For j = r1 To .UsedRange.Rows.Count
                    str1 = CStr(.Cells(j, colidx).Value)
                    If str1 <> "" Then
                        str2 = CStr(.Cells(j, colMark).Value)
                        If str2 = CST_MARK_ITEM Then
                            '実習項目
                            str3 = str1
                            If str3 = "→" Then
                                str3 = lookLeft(j, colidx)  '左側より取得
                            End If
                            If CStr(.Cells(j, colJikan).Value) <> "" Then
                                str3 = CStr(.Cells(j, colJikan).Value)
                            End If
                            v1 = Split(str3, ",")       '「a,p」など午前＆午後のパターン対応
                            For k = 0 To UBound(v1)
                                str3 = v1(k)
                                
                                timetbl = cnvTimetbl(str3)
                                If timetbl > 0 Then
                                    '有効な時間帯
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
                
                '編集データを時間毎の区分でソート
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
                
                '各時間帯毎の最大行数を取得
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
    
    '時間帯毎の印字開始行を設定
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
            
    '火曜日以降の「→」について印字開始行を調整
    For i = 2 To 5
        For j = 1 To dcnt1(i)
            If data1(i, j).settei = "→" Then
                Call checkYajirusi(dcnt1, data1, i, j)
            End If
        Next j
    Next i
     
    
    
    

    '時間帯毎の開始行を初期セット
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
    'お昼(a5)の位置に余裕があったら、   a4以降の位置を下げて調整する
    k = Range(amRange).Rows.Count + 1 - (timeRow(5) + timeNum(5))
    If k > 0 Then
        For i = 4 To 10
            timeRow(i) = timeRow(i) + k
        Next i
    End If
    '16時(p4),17時(p5)の位置を調整
    k = (Range(targetRange).Rows.Count) - (timeRow(10) + timeNum(10))
    If timeNum(10) = 0 Then
        k = k - 1
    End If
    If k > 0 Then
        timeRow(9) = timeRow(9) + k
        timeRow(10) = timeRow(10) + k
    End If

    '時間帯の印字
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
                str1 = "(午前) 8:30～"
            Case "A2"
                str1 = " 9:00頃～"
            Case "A3"
                str1 = " 9:30頃～"
            Case "A4"
                str1 = "10:00頃～"
            Case "A5"
                str1 = "10:30頃～"
            Case "A6"
                str1 = "11:00頃～"
            Case "A7"
                str1 = "11:30頃～"
            Case "A8"
                str1 = "12:00～"
            Case "A9"
                str1 = "12:30～"
            Case "P1"
                str1 = "(午後)13:00～"
            Case "P2"
                str1 = "13:30頃～"
            Case "P3"
                str1 = "14:00頃～"
            Case "P4"
                str1 = "14:30頃～"
            Case "P5"
                str1 = "15:00頃～"
            Case "P6"
                str1 = "15:30頃～"
            Case "P7"
                str1 = "16:00頃～"
            Case "P8"
                str1 = "16:30頃～"
            Case "P9"
                str1 = "17:00 終了"
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
           
    '印字
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
                If data1(i, j).settei = "→" Then
                    If k = 1 Then
                        'targetSheet.Cells(l + Range(targetRange).Row, (i * 3) + 1).Value = "→"
                    End If
                Else
                    str1 = data1(i, j).prnData(k)
                    If fnGakuseiSu() > 1 Then
                        '学生２名以上の場合
                        If data1(i, j).fukusu <> "" Then
                            str1 = Replace(str1, "[@]", data1(i, j).fukusu)
                        End If
                    Else
                        '学生が一人
                        str1 = Replace(str1, "[@]", "")
                    End If
                    If k = 1 And data1(i, j).tanto <> "" Then
                        str1 = str1 & "  (" & data1(i, j).tanto & ")"
                    End If
                    '印字内容セット
                    targetSheet.Cells(l + Range(targetRange).Row, (i * 3) + 1).Value = str1
                    If data1(i, j).settei = LCase(data1(i, j).settei) Then
                        'カレンダ部に小文字で入力されている場合はイタリック体とする
                        targetSheet.Cells(l + Range(targetRange).Row, (i * 3) + 1).Font.Italic = True
                    End If
                    '担当者を網掛け表示
                    If k = 1 And data1(i, j).tanto <> "" Then
                        Call tanto_amikake(targetSheet, targetSheet.Cells(l + Range(targetRange).Row, (i * 3) + 1))
                    End If
                End If
                l = l + 1
            Next k
            If rc > 1 And data1(i, j).settei <> "→" Then
                '左カッコ
                With targetSheet.Range(Cells(r1, c1), Cells(r2, c1))
                    targetSheet.Shapes.AddShape(Type:=msoShapeLeftBracket, Left:=.Left + 3, Top:=.Top + 3, Width:=.Width - 6, Height:=.Height - 6).Select
                    
                End With
            End If
            If data1(i, j).yajirusi > 0 Then
                '右カッコ＆→
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

'カレンダ部入力値が「→」の場合に、当該セルの左側の列より値を取得します
Private Function lookLeft(rowIdx As Long, colidx As Long) As String
    Dim str1                As String
    
    lookLeft = ""
    
    If colidx > Range(CST_RANGE_J_CAL_TOP).Column Then
        str1 = CStr(ThisWorkbook.Sheets(CST_SHT_KOMOKU).Cells(rowIdx, colidx - 1).Value)
        If str1 = "→" Then
            '再帰的に呼び出し
            lookLeft = lookLeft(rowIdx, colidx - 1)
        Else
            lookLeft = str1
        End If
    End If
    
End Function

'時間帯区分（1から10）を項目定義値またはカレンダ部入力値より変換して返します
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


'→の位置調整
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
    
    'はじめに当該項目について印字開始行のずれを確認
    For i = 1 To 5
        rIdx(i) = 0
        For j = 1 To dcnt1(i)
            If data1(i, j).rowIdx = rowIdx And data1(i, j).timetbl = timetbl Then
                '同一データ
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
    
    '親子関係のセット
    For i = 1 To 5
        If rIdx(i) > 0 Then
            If data1(i, rIdx(i)).settei <> "→" Then
                '親
                p1 = i
                cnt = 0
                data1(i, rIdx(i)).yajirusi = 0
            Else
                '子
                cnt = cnt + 1
                data1(p1, rIdx(p1)).yajirusi = cnt
                data1(i, rIdx(i)).yajirusi2 = cnt
            End If
        End If
    Next i
    
    If minPos = maxPos Then
        'ずれは無い
        Exit Sub
    End If
    
    '印字開始行を大きい値に合わせる
    For i = 1 To 5
        For j = 1 To dcnt1(i)
            If data1(i, j).rowIdx = rowIdx And data1(i, j).timetbl = timetbl Then
                '同一データ
                pos = data1(i, j).prnStartRow
                If pos < maxPos Then
                    data1(i, j).dmyline = maxPos - pos
                    data1(i, j).prnRowsCount = data1(i, j).prnRowsCount + (maxPos - pos)
                End If
            End If
        Next j
    Next i
    
    '時間帯毎の印字開始行を再設定
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


'担当者表示を網掛けとする
'セル無いのテキストの該当部を白色（テーマカラー：１）とし、セル末尾箇所にテキスト図形で描画する
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
    
    '(担当者)を探索
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
        '担当者発見
        str2 = Mid(str1, p1, p2 - p1 + 1)
        
        '担当者の表示色を白色へ変更（見えなくする）
        rng1.Characters(Start:=p1, Length:=Len(str2)).Font.ThemeColor = 1
        
        '担当者のテキストボックスを描画
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
            .NameComplexScript = "メイリオ"
            .NameFarEast = "メイリオ"
            .Name = "メイリオ"
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






