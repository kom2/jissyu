VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'薬局実習スケジュール作表「実習項目」シート
'                                           H. Komatsu
'                                           2014.1.9

'実習項目シートの変更
Private Sub Worksheet_Change(ByVal target As Range)
    Call wSheetChage_Jissyu(target)
End Sub

'実習項目シートのセレクション変更
Private Sub Worksheet_SelectionChange(ByVal target As Range)
    Call wSheetSelectionChage_Jissyu(target)
End Sub



'印字項目表示のチェックボックス
Private Sub CheckBox1_Click()

    If CheckBox1.Value = True Then
        '印字項目表示
        If ActiveSheet.AutoFilterMode Then
            ActiveSheet.AutoFilterMode = False
        End If
        Columns(5).EntireColumn.Hidden = False
    Else
        '印字項目を隠す
        Columns(5).EntireColumn.Hidden = True
        If ActiveSheet.AutoFilterMode Then
            ActiveSheet.AutoFilterMode = False
        End If
        Columns(2).Select
        Selection.AutoFilter Field:=1, Criteria1:="<>"
        
    End If

    Cells(3, 1).Select
End Sub

'作表ボタン
Private Sub CommandButton1_Click()
    Dim strAdr              As String
    
    CommandButton1.Enabled = False
    strAdr = ActiveCell.Address
    Call makeSchedule
    CommandButton1.Enabled = True
    ThisWorkbook.Activate
    Me.Range(strAdr).Select
End Sub

'残り当番のクリア
Private Sub CommandButton2_Click()
    CommandButton2.Enabled = False
    Call clearNokori
    CommandButton2.Enabled = True
End Sub

'残り当番のセット
Private Sub CommandButton3_Click()
    CommandButton3.Enabled = False
    Call setNokori
    CommandButton3.Enabled = True
End Sub

'実習項目シートのセレクション変更
