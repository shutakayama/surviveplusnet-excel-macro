VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReplaceForm 
   Caption         =   "置換"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5370
   OleObjectBlob   =   "ReplaceForm.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "ReplaceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' DialogResult プロパティのバッキングフィールドです。
Private valueOfDialogResult As Boolean

'------------------------------------------------------------------------
' FindText プロパティ
'------------------------------------------------------------------------
'   目的：      検索する文字列を取得または設定します。
'   注釈：
'   変更履歴：  s-koga 2010/06/23 - オリジナル
'------------------------------------------------------------------------
Public Property Get FindText() As String
    ' Return
    FindText = find.Text
End Property
Public Property Let FindText(ByVal vNewValue As String)
    find.Text = vNewValue
End Property

'------------------------------------------------------------------------
' ReplaceText プロパティ
'------------------------------------------------------------------------
'   目的：      置換後の文字列を取得または設定します。
'   注釈：
'   変更履歴：  s-koga 2010/06/23 - オリジナル
'------------------------------------------------------------------------
Public Property Get ReplaceText() As String
    ' Return
    ReplaceText = Replace.Text
End Property
Public Property Let ReplaceText(ByVal vNewValue As String)
    Replace.Text = vNewValue
End Property

'------------------------------------------------------------------------
' DialogResult プロパティ
'------------------------------------------------------------------------
'   目的：      キャンセルの時は False、それ以外は True を取得または設定します。
'   注釈：
'   変更履歴：  s-koga 2010/06/23 - オリジナル
'------------------------------------------------------------------------
Public Property Get DialogResult() As Boolean
    ' Return
    DialogResult = valueOfDialogResult
End Property
Public Property Let DialogResult(ByVal vNewValue As Boolean)
    valueOfDialogResult = vNewValue
End Property

' [キャンセル] ボタンが押されたときの処理
' フォームを閉じて、DialogResult プロパティを False に変更します。
Private Sub cancelButton_Click()
    valueOfDialogResult = False
    Me.Hide
End Sub

' [置換を実行] ボタンが押されたときの処理を実行
' フォームを閉じて、DialogResult プロパティを True に変更します。
Private Sub executeButon_Click()
    If Me.FindText = "" Then Exit Sub
    
    valueOfDialogResult = True
    Me.Hide
End Sub


