VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CopyWorkSheetDialogForm 
   Caption         =   "ワークシートを複数コピー"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5205
   OleObjectBlob   =   "CopyWorkSheetDialogForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "CopyWorkSheetDialogForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' DialogResult プロパティのバッキングフィールドです。
Private valueOfDialogResult As Boolean

'------------------------------------------------------------------------
' NewNames プロパティ
'------------------------------------------------------------------------
'   目的：      ユーザーによって指定された新しいワークシートの名前の配列を取得または設定します。
'   注釈：
'   変更履歴：  s-koga 2010/07/05 - オリジナル
'------------------------------------------------------------------------
Public Property Get NewNames() As Variant
    ' Return
    NewNames = Split(newNamesBox.Text, vbCrLf)
End Property
Public Property Let NewNames(ByVal vNewValue As Variant)
    newNamesBox.Text = Join(vNewValue, vbCrLf)
    
    newNamesBox.SelStart = 0
    newNamesBox.SelLength = Len(newNamesBox.Text)
    
End Property

'------------------------------------------------------------------------
' DialogResult プロパティ
'------------------------------------------------------------------------
'   目的：      ユーザーがOKボタンを押したかどうかを取得または設定します。
'   注釈：
'   変更履歴：  s-koga 2010/07/05 - オリジナル
'------------------------------------------------------------------------
Public Property Get DialogResult() As Boolean
    ' Return
    DialogResult = valueOfDialogResult
End Property
Public Property Let DialogResult(ByVal vNewValue As Boolean)
    valueOfDialogResult = vNewValue
End Property

' キャンセルボタンが押されたときの処理を実行します。
Private Sub cancelButton_Click()
    Me.DialogResult = False
    Call Me.Hide
End Sub

' OKボタンが押されたときの処理を実行します。
Private Sub okButton_Click()
    If newNamesBox.Text = "" Then
        newNamesBox.SetFocus
        Exit Sub
    End If
    
    Me.DialogResult = True
    Call Me.Hide
End Sub

'------------------------------------------------------------------------
'   目的：      ダイアログを表示し、OKが押されたかどうかを返します
'   戻り値：    ユーザーがOKボタンを押したかどうかを Boolean として返します。
'   注釈：
'   使用方法：If dialog.ShowDialog() Then
'------------------------------------------------------------------------
Public Function ShowDialog() As Boolean
    Me.Show (vbModal)
    ShowDialog = Me.DialogResult
End Function

Private Sub UserForm_Click()

End Sub
