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
' ActionDelegate �N���X
'------------------------------------------------------------------------
'   �ړI�F      �����֐��f���Q�[�g ���i��N���X�ł��B
'   ���߁F      ���̃N���X���g�p����ɂ́AMicrosoft Script Control 1.0 �ւ̎Q�Ɛݒ肪�K�v�ł��B
'   �g�p���@�F  FuncModule �� Action ���\�b�h���g���ď��������Ďg�p���Ă��������B
'                  ���၄Action("a,b", "a=a*10:MsgBox a+b").Run(1,1)
'   �ύX�����F  SHIN-ICHI 2010/10/21 - �V�K�쐬
'   MIT License / Copyright(C) 2010 SHIN-ICHI all rights reserverd. ( http:// surviveplus.net )
'------------------------------------------------------------------------

Private myScriptControl As New ScriptControl
Private valueOfArgs As String
Private valueOfStatements As String

Private valueOfClosureVariableName As String
Private valueOfClosureObject As Object

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
'   �ړI�F      �����֐������s���܂��B
'   ���́F      ���� args - �֐��ɓn�������i���������Ɏw�肵�����j�w�肵�܂��B
'                 ���̐��Ə��Ԃ͏��������Ɏw�肵�� args �ɍ��킹�Ă��������B
'   ���߁F      ���̃��\�b�h���� Invoke �łȂ� Run �Ȃ̂́AInvoke �� Visual Basic 6�n �̗\���ł��邽�߂ł��B
'                 �܂��A���̃��\�b�h������̃v���p�e�B�Ƃ��Ď������Ȃ������̂́AActionDelegate �I�u�W�F�N�g�̕ϐ��ւ̊��蓖�Ă��ʓ|�ɂȂ邽�߂ł��B
'   �g�p���@�F  ���၄Action("a,b", "a=a*10:MsgBox a+b").Run(1,1)
'   �ύX�����F  SHIN-ICHI 2010/10/21 - �V�K�쐬
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
    
    If valueOfClosureVariableName <> "" Then Call myScriptControl.AddObject(valueOfClosureVariableName, valueOfClosureObject)
    
    myScriptControl.AddCode ("Sub mydelegate(" & argsNamesText & "): " & valueOfStatements & ":End Sub")
    myScriptControl.ExecuteStatement ("mydelegate(" & argsValuesText & ")")
End Function

Public Function WithClosure(ByVal variableName As String, ByVal closure As Variant) As ActionDelegate
    Set WithClosure = Me
    valueOfClosureVariableName = variableName
    Set valueOfClosureObject = closure
End Function
