VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActionDelegate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------
' ActionDelegate クラス
'------------------------------------------------------------------------
'   目的：      匿名関数デリゲート を司るクラスです。
'   注釈：      このクラスを使用するには、Microsoft Script Control 1.0 への参照設定が必要です。
'   使用方法：  FuncModule の Action メソッドを使って初期化して使用してください。
'                  ＜例＞Action("a,b", "a=a*10:MsgBox a+b").Run(1,1)
'   変更履歴：  SHIN-ICHI 2010/10/21 - 新規作成
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
'   目的：      匿名関数を実行します。
'   入力：      引数 args - 関数に渡す引数（初期化時に指定した個数）指定します。
'                 数の数と順番は初期化時に指定した args に合わせてください。
'   注釈：      このメソッド名が Invoke でなく Run なのは、Invoke が Visual Basic 6系 の予約語であるためです。
'                 また、このメソッドを既定のプロパティとして実装しなかったのは、ActionDelegate オブジェクトの変数への割り当てが面倒になるためです。
'   使用方法：  ＜例＞Action("a,b", "a=a*10:MsgBox a+b").Run(1,1)
'   変更履歴：  SHIN-ICHI 2010/10/21 - 新規作成
'   MIT License / Copyright(C) 2010 SHIN-ICHI all rights reserverd. ( http:// surviveplus.net )
'------------------------------------------------------------------------
Public Function Run(ParamArray args() As Variant)
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
            argsValuesText = argsValuesText & IIf(notObjectIsFirst, "", ",") & IIf(TypeName(args(index)) = "String", """" & args(index) & """", args(index))
            notObjectIsFirst = False
        End If
        index = index + 1
        If UBound(args) < index Then Exit For
    Next s
    
    myScriptControl.AddCode ("Sub mydelegate(" & argsNamesText & "): " & replace(valueOfStatements, "Return ", "mydelegate = ") & ":End Sub")
    myScriptControl.ExecuteStatement ("mydelegate(" & argsValuesText & ")")
End Function




