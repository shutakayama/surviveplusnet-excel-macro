VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FuncDelegate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'------------------------------------------------------------------------
' FuncDelegate クラス
'------------------------------------------------------------------------
'   目的：      匿名関数デリゲート を司るクラスです。
'   注釈：      このクラスを使用するには、Microsoft Script Control 1.0 への参照設定が必要です。
'   使用方法：  FuncModule の Func メソッドを使って初期化して使用してください。
'                  ＜例＞If Func("a,b", "a=a*10:Return a+b").Run(1,1) = 11 Then MsgBox "OK"
'   変更履歴：  SHIN-ICHI 2010/10/9 - 新規作成
'   MIT License / Copyright(C) 2010 SHIN-ICHI all rights reserverd. ( http:// surviveplus.net )
'------------------------------------------------------------------------

Private myScriptControl As New ScriptControl
Private valueOfArgs As String
Private valueOfStatements As String

Private Sub Class_Initialize()
    myScriptControl.Language = "VBScript"
End Sub

Private Sub Class_Terminate()
    Set myScriptControl = Nothing
End Sub

Public Sub SetStatement(args As String, statements As String)
    valueOfArgs = args
    valueOfStatements = statements
End Sub

'------------------------------------------------------------------------
'   目的：      匿名関数を実行して結果を返します。
'   入力：      引数 args - 関数に渡す引数（初期化時に指定した個数）指定します。
'                 数の数と順番は初期化時に指定した args に合わせてください。
'   戻り値：    実行した結果を返します。戻り値の内容は statement の内容によって異なります。
'                  例えば statement で評価式を Return している場合は Boolean を返します。
'   注釈：      このメソッド名が Invoke でなく Run なのは、Invoke が Visual Basic 6系 の予約語であるためです。
'                 また、このメソッドを既定のプロパティとして実装しなかったのは、FuncDelegate オブジェクトの変数への割り当てが面倒になるためです。
'   使用方法：  ＜例＞If Func("a,b", "a=a*10:Return a+b").Run(1,1) = 11 Then MsgBox "OK"
'   変更履歴：  SHIN-ICHI 2010/10/9 - 新規作成
'   MIT License / Copyright(C) 2010 SHIN-ICHI all rights reserverd. ( http:// surviveplus.net )
'------------------------------------------------------------------------
Public Property Get Run(ParamArray args() As Variant) As Variant
    myScriptControl.Reset
    myScriptControl.Language = "VBScript"
    
    Dim s As Variant
    Dim index As Integer: index = LBound(args)
    Dim argsNamesText As String
    Dim argsValuesText As String
    Dim notObjectIsFirst As Boolean: notObjectIsFirst = True
    For Each s In Split(valueOfArgs, ",")
        Dim argName  As String: argName = Trim(s)
        If IsObject(args(index)) Then
            Call myScriptControl.AddObject(argName, args(index), True)
        Else
            argsNamesText = argsNamesText & IIf(notObjectIsFirst, "", ",") & argName
            argsValuesText = argsValuesText & IIf(notObjectIsFirst, "", ",") & args(index)
            notObjectIsFirst = False
        End If
        index = index + 1
        If UBound(args) < index Then Exit For
    Next s
    
    myScriptControl.AddCode ("Function mydelegate(" & argsNamesText & "): " & Replace(valueOfStatements, "Return ", "mydelegate = ") & ":End Function")
    Run = myScriptControl.Eval("mydelegate(" & argsValuesText & ")")
End Property



