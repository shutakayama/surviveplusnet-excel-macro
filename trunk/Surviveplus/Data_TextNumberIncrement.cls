VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Data_TextNumberIncrement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------
'   Data_TextNumberIncrement クラス
'------------------------------------------------------------------------
'   目的：      テキスト内の数値のインクリメントを司るクラスです。
'   注釈：      Data.末尾の数値をインクリメント 用のクラスです。
'   使用方法：
'   変更履歴：  SHIN-ICHI 2009/3/23 - オリジナル（プロシージャ内の変数やロジックから分離）
'   Copyright(C) 2009 SHIN-ICHI all rights reservd.
'------------------------------------------------------------------------

'
' パターン検出に使用する正規表現オブジェクトです。
' インスタンス内にキャッシュして、再作成によるパフォーマンスの低下を防ぎます。
'
Dim regNotNumber As Object
Dim regEndNumber As Object
Dim matches As Object
   
'------------------------------------------------------------------------
'   目的：      インスタンスを初期化します。
'   注釈：
'   変更履歴：  SHIN-ICHI 2009/3/23 - オリジナル
'------------------------------------------------------------------------
Private Sub Class_Initialize()
    
    Set regNotNumber = CreateObject("VBScript.RegExp")
    regNotNumber.Pattern = "[^0-9]"
    
    Set regEndNumber = CreateObject("VBScript.RegExp")
    regEndNumber.Pattern = "^(.*)([^0-9])([0-9]+)([^0-9]?)(.*?)$"
    
End Sub

'------------------------------------------------------------------------
'   目的：      インスタンスの破棄を行います。
'   注釈：
'   変更履歴：  SHIN-ICHI 2009/3/23 - オリジナル
'------------------------------------------------------------------------
Private Sub Class_Terminate()
    On Error Resume Next
    Set matches = Nothing
    Set regNotNumber = Nothing
    Set regEndNumber = Nothing
End Sub

'------------------------------------------------------------------------
'   目的：      指定されたテキストの末尾の数値をインクリメントして返します。
'   入力：      引数 text - 数値をインクリメントするテキストを指定します。
'   戻り値：    末尾の数値がインクリメントされたテキストを返します。インクリメント出来ないときはそのまま返します。
'   注釈：
'   使用方法：newText = textNumberIncrement.Increment( targetText )
'   変更履歴：  SHIN-ICHI 2009/3/23 - オリジナル
'------------------------------------------------------------------------
Public Function Increment(ByVal targetText As String) As String

    If targetText = "" Then Exit Function
    
    If regNotNumber.Test(targetText) Then
        Set matches = regEndNumber.Execute(targetText)
        
        If matches.Count > 0 Then
            With matches(0)
                Dim oldNumber As String: oldNumber = .SubMatches(2)
                Dim newNumber As String: newNumber = CInt(oldNumber) + 1
                If Len(newNumber) < Len(oldNumber) Then
                    newNumber = String(Len(oldNumber) - Len(newNumber), "0") & newNumber
                End If
                
                'Return
                Increment = .SubMatches(0) & .SubMatches(1) & newNumber & .SubMatches(3) & .SubMatches(4)
            End With
        Else
            'Return
            Increment = targetText
        End If
    
    Else
        'Return
        Increment = CInt(targetText) + 1
    End If
End Function

