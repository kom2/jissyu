VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'薬局実習スケジュール作表「基本設定」シート
'                                           H. Komatsu
'                                           2014.1.9

'基本設定シートの変更
Private Sub Worksheet_Change(ByVal target As Range)
    Call wSheetChage_Kihon(target)
End Sub

