VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'��ǎ��K�X�P�W���[����\�u��{�ݒ�v�V�[�g
'                                           H. Komatsu
'                                           2014.1.9

'��{�ݒ�V�[�g�̕ύX
Private Sub Worksheet_Change(ByVal target As Range)
    Call wSheetChage_Kihon(target)
End Sub
