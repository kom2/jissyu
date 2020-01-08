VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
   Caption         =   "UserForm3"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7695
   OleObjectBlob   =   "UserForm3.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'----------------------------------------------------------------
'��ǎ��K�X�P�W���[����\
'   ��\�Ώۂ̑I��
'                                           H. Komatsu
'                                           2015.5.31
'----------------------------------------------------------------

Private flgInit             As Boolean          '�������ς݃t���O
Private flgDisEvents        As Boolean          '�C�x���g����p


'�t�H�[����������
Private Sub UserForm_Initialize()
    
    flgInit = False
    
End Sub


'�t�H�[���A�N�e�B�x�[�g��
Private Sub UserForm_Activate()

    If flgInit = False Then
        Call init
        flgInit = True
    End If
    
    Call checkEntry

End Sub


'������
Private Sub init()
    Dim vCtrl               As Control
    Dim weeks               As Long
    Dim n                   As Long
    Dim str1                As String
    Dim date1               As Long
    Dim date2               As Long
    
    flgDisEvents = True                         '�C�x���g�}�~
    
    '���T�ڂ܂ł��邩
    weeks = fnWeeks
    
    '�W��Z�b�g
    Me.Caption = "���K�X�P�W���[����\�@��\�Ώۂ̑I��"
    Label1.Caption = "��\�Ώۂ�I�����ĉ������B"
    For Each vCtrl In Me.Controls
        If Left(vCtrl.Name, 8) = "CheckBox" Then
            n = CLng(Mid(vCtrl.Name, 9)) - 1
            Select Case True
                Case (n = 0)
                    str1 = "�\��"
                Case ((n <= 8) And (weeks >= n))
                    date1 = fn1stMonday + (n - 1) * 7
                    date2 = date1 + 4
                    str1 = "��@week�T�@�i@start�`@end�j"
                    str1 = Replace(str1, "@week", StrConv(CStr(n), vbWide))
                    str1 = Replace(str1, "@start", StrConv(Format(date1, "m/d"), vbWide))
                    str1 = Replace(str1, "@end", StrConv(Format(date2, "m/d"), vbWide))
                Case (n >= 9 And (weeks >= ((n - 9) * 3 + 9)))
                    date1 = fn1stMonday + (((n - 9) * 3 + 9) - 1) * 7
                    date2 = date1 + 18
                    str1 = "��@week�T�`�i@start�`@end�j"
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
    CommandButton1.Caption = "�S�I��/�S����"
    CommandButton2.Caption = "��\�̊J�n"
    CommandButton3.Caption = "�L�����Z��"
    
    flgDisEvents = False                        '�C�x���g�}�~����

End Sub


'���̓`�F�b�N
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


'��\�J�n�{�^��
Private Sub CommandButton2_Click()
    
    Me.Tag = "OK"
    Me.Hide

End Sub


'�L�����Z���{�^��
Private Sub CommandButton3_Click()
    
    Me.Tag = ""
    Me.Hide

End Sub


'�S�I���^�S�����{�^��
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
                    '�`�F�b�N����Ă��Ȃ����̂�������
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


'�`�F�b�N�{�b�N�X
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


'�I�����ꂽ��\�Ώۂ̃f�[�^
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

