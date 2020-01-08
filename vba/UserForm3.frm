VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "UserForm3"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7695
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'----------------------------------------------------------------
'薬局実習スケジュール作表
'   作表対象の選択
'                                           H. Komatsu
'                                           2015.5.31
'----------------------------------------------------------------

Private flgInit             As Boolean          '初期化済みフラグ
Private flgDisEvents        As Boolean          'イベント制御用


'フォーム初期化時
Private Sub UserForm_Initialize()
    
    flgInit = False
    
End Sub


'フォームアクティベート時
Private Sub UserForm_Activate()

    If flgInit = False Then
        Call init
        flgInit = True
    End If
    
    Call checkEntry

End Sub


'初期化
Private Sub init()
    Dim vCtrl               As Control
    Dim weeks               As Long
    Dim n                   As Long
    Dim str1                As String
    Dim date1               As Long
    Dim date2               As Long
    
    flgDisEvents = True                         'イベント抑止
    
    '何週目まであるか
    weeks = fnWeeks
    
    '標題セット
    Me.Caption = "実習スケジュール作表　作表対象の選択"
    Label1.Caption = "作表対象を選択して下さい。"
    For Each vCtrl In Me.Controls
        If Left(vCtrl.Name, 8) = "CheckBox" Then
            n = CLng(Mid(vCtrl.Name, 9)) - 1
            Select Case True
                Case (n = 0)
                    str1 = "表紙"
                Case ((n <= 8) And (weeks >= n))
                    date1 = fn1stMonday + (n - 1) * 7
                    date2 = date1 + 4
                    str1 = "第@week週　（@start～@end）"
                    str1 = Replace(str1, "@week", StrConv(CStr(n), vbWide))
                    str1 = Replace(str1, "@start", StrConv(Format(date1, "m/d"), vbWide))
                    str1 = Replace(str1, "@end", StrConv(Format(date2, "m/d"), vbWide))
                Case (n >= 9 And (weeks >= ((n - 9) * 3 + 9)))
                    date1 = fn1stMonday + (((n - 9) * 3 + 9) - 1) * 7
                    date2 = date1 + 18
                    str1 = "第@week週～（@start～@end）"
                    str1 = Replace(str1, "@week", StrConv(CStr(((n - 9) * 3 + 9)), vbWide))
                    str1 = Replace(str1, "@start", StrConv(Format(date1, "m/d"), vbWide))
                    str1 = Replace(str1, "@end", StrConv(Format(date2, "m/d"), vbWide))
                Case Else
                    str1 = ""
            End Select
            If str1 = "" Then
                vCtrl.Caption = ""
                vCtrl.Value = False
                vCtrl.Enabled = False
                vCtrl.Visible = False
            Else
                vCtrl.Caption = str1
                vCtrl.Value = False
                vCtrl.Enabled = True
                vCtrl.Visible = True
            End If
        End If
    Next
    CommandButton1.Caption = "全選択/全解除"
    CommandButton2.Caption = "作表の開始"
    CommandButton3.Caption = "キャンセル"
    
    flgDisEvents = False                        'イベント抑止解除

End Sub


'入力チェック
Private Sub checkEntry()
    Dim vCtrl               As Control
    Dim flgchk              As Boolean
    
    flgchk = False
    For Each vCtrl In Me.Controls
        If Left(vCtrl.Name, 8) = "CheckBox" Then
            If vCtrl.Enabled = True And vCtrl.Value = True Then
                flgchk = True
                Exit For
            End If
        End If
    Next
    If flgchk = True Then
        CommandButton2.Enabled = True
    Else
        CommandButton2.Enabled = False
    End If

End Sub


'作表開始ボタン
Private Sub CommandButton2_Click()
    
    Me.Tag = "OK"
    Me.Hide

End Sub


'キャンセルボタン
Private Sub CommandButton3_Click()
    
    Me.Tag = ""
    Me.Hide

End Sub


'全選択／全解除ボタン
Private Sub CommandButton1_Click()
    Dim vCtrl               As Control
    Dim cntOff              As Long
    Dim flg1                As Boolean
    
    If flgDisEvents Then
        Exit Sub
    End If
    
    flgDisEvents = True
    
    flg1 = False
    
    For Each vCtrl In Me.Controls
        If Left(vCtrl.Name, 8) = "CheckBox" Then
            If vCtrl.Enabled = True Then
                If vCtrl.Value = False Then
                    'チェックされていないものがあった
                    flg1 = True
                    Exit For
                End If
            End If
        End If
    Next
    
    For Each vCtrl In Me.Controls
        If Left(vCtrl.Name, 8) = "CheckBox" Then
            If vCtrl.Enabled = True Then
                vCtrl.Value = flg1
            End If
        End If
    Next
    
    Call checkEntry
    
    flgDisEvents = False
    
End Sub


'チェックボックス
Private Sub CheckBox1_Click()
    Call checkEntry
End Sub

Private Sub CheckBox2_Click()
    Call checkEntry
End Sub

Private Sub CheckBox3_Click()
    Call checkEntry
End Sub

Private Sub CheckBox4_Click()
    Call checkEntry
End Sub

Private Sub CheckBox5_Click()
    Call checkEntry
End Sub

Private Sub CheckBox6_Click()
    Call checkEntry
End Sub

Private Sub CheckBox7_Click()
    Call checkEntry
End Sub

Private Sub CheckBox8_Click()
    Call checkEntry
End Sub

Private Sub CheckBox9_Click()
    Call checkEntry
End Sub

Private Sub CheckBox10_Click()
    Call checkEntry
End Sub

Private Sub CheckBox11_Click()
    Call checkEntry
End Sub

Private Sub CheckBox12_Click()
    Call checkEntry
End Sub


'選択された作表対象のデータ
Public Function selctionData() As Variant
    Dim vData(17)           As Boolean
    Dim i                   As Long
    Dim j                   As Long
    Dim vCtrl               As Control
    
    For i = 0 To 17
        vData(i) = False
    Next i
    
    For Each vCtrl In Me.Controls
        If Left(vCtrl.Name, 8) = "CheckBox" Then
            i = CLng(Mid(vCtrl.Name, 9)) - 1
            If i <= 8 Then
                vData(i) = vCtrl.Value
            Else
                j = (i - 9) * 3 + 9
                vData(j) = vCtrl.Value
                vData(j + 1) = vCtrl.Value
                vData(j + 2) = vCtrl.Value
            End If
        End If
    Next
            
    selctionData = vData
    
    
End Function

