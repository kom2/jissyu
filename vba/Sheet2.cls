VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'��ǎ��K�X�P�W���[����\�u���K���ځv�V�[�g
'                                           H. Komatsu
'                                           2014.1.9

'���K���ڃV�[�g�̕ύX
Private Sub Worksheet_Change(ByVal target As Range)
    Call wSheetChage_Jissyu(target)
End Sub

'���K���ڃV�[�g�̃Z���N�V�����ύX
Private Sub Worksheet_SelectionChange(ByVal target As Range)
    Call wSheetSelectionChage_Jissyu(target)
End Sub



'�󎚍��ڕ\���̃`�F�b�N�{�b�N�X
Private Sub CheckBox1_Click()

    If CheckBox1.Value = True Then
        '�󎚍��ڕ\��
        If ActiveSheet.AutoFilterMode Then
            ActiveSheet.AutoFilterMode = False
        End If
        Columns(5).EntireColumn.Hidden = False
    Else
        '�󎚍��ڂ��B��
        Columns(5).EntireColumn.Hidden = True
        If ActiveSheet.AutoFilterMode Then
            ActiveSheet.AutoFilterMode = False
        End If
        Columns(2).Select
        Selection.AutoFilter Field:=1, Criteria1:="<>"
        
    End If

    Cells(3, 1).Select
End Sub

'��\�{�^��
Private Sub CommandButton1_Click()
    Dim strAdr              As String
    
    CommandButton1.Enabled = False
    strAdr = ActiveCell.Address
    Call makeSchedule
    CommandButton1.Enabled = True
    ThisWorkbook.Activate
    Me.Range(strAdr).Select
End Sub

'�c�蓖�Ԃ̃N���A
Private Sub CommandButton2_Click()
    CommandButton2.Enabled = False
    Call clearNokori
    CommandButton2.Enabled = True
End Sub

'�c�蓖�Ԃ̃Z�b�g
Private Sub CommandButton3_Click()
    CommandButton3.Enabled = False
    Call setNokori
    CommandButton3.Enabled = True
End Sub

'���K���ڃV�[�g�̃Z���N�V�����ύX