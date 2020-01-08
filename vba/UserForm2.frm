VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6945
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strItemValue        As String
Private strPrntValue        As String

'表示データ設定
Public Sub setValues(strDate As String, strNokori As String, strItem As String, strPrnt As String)
    Dim v1                  As Variant
    Dim i                   As Long
    
    Me.Caption = Application.Name
    CheckBox1.Caption = "印字内容表示"
    CommandButton1.Caption = "OK"
    
    Label1.Caption = strDate
    Label2.Caption = strNokori
    strItemValue = strItem
    strPrntValue = strPrnt

    Call viewdata
    
End Sub




Private Sub CheckBox1_Click()
    Call viewdata
End Sub

Private Sub CommandButton1_Click()
    Me.Hide
End Sub

Private Sub viewdata()
    Dim v1                  As Variant
    
    If CheckBox1.Value Then
        v1 = Split(strPrntValue, vbLf)
    Else
        v1 = Split(strItemValue, vbLf)
    End If

    ListBox1.Clear
    For i = 0 To UBound(v1)
        ListBox1.AddItem (v1(i))
    Next i
End Sub
