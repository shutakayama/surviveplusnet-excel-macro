VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReplaceForm 
   Caption         =   "�u��"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5370
   OleObjectBlob   =   "ReplaceForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "ReplaceForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' DialogResult �v���p�e�B�̃o�b�L���O�t�B�[���h�ł��B
Private valueOfDialogResult As Boolean

'------------------------------------------------------------------------
' FindText �v���p�e�B
'------------------------------------------------------------------------
'   �ړI�F      �������镶������擾�܂��͐ݒ肵�܂��B
'   ���߁F
'   �ύX�����F  s-koga 2010/06/23 - �I���W�i��
'------------------------------------------------------------------------
Public Property Get FindText() As String
    ' Return
    FindText = find.Text
End Property
Public Property Let FindText(ByVal vNewValue As String)
    find.Text = vNewValue
End Property

'------------------------------------------------------------------------
' ReplaceText �v���p�e�B
'------------------------------------------------------------------------
'   �ړI�F      �u����̕�������擾�܂��͐ݒ肵�܂��B
'   ���߁F
'   �ύX�����F  s-koga 2010/06/23 - �I���W�i��
'------------------------------------------------------------------------
Public Property Get ReplaceText() As String
    ' Return
    ReplaceText = Replace.Text
End Property
Public Property Let ReplaceText(ByVal vNewValue As String)
    Replace.Text = vNewValue
End Property

'------------------------------------------------------------------------
' DialogResult �v���p�e�B
'------------------------------------------------------------------------
'   �ړI�F      �L�����Z���̎��� False�A����ȊO�� True ���擾�܂��͐ݒ肵�܂��B
'   ���߁F
'   �ύX�����F  s-koga 2010/06/23 - �I���W�i��
'------------------------------------------------------------------------
Public Property Get DialogResult() As Boolean
    ' Return
    DialogResult = valueOfDialogResult
End Property
Public Property Let DialogResult(ByVal vNewValue As Boolean)
    valueOfDialogResult = vNewValue
End Property

' [�L�����Z��] �{�^���������ꂽ�Ƃ��̏���
' �t�H�[������āADialogResult �v���p�e�B�� False �ɕύX���܂��B
Private Sub cancelButton_Click()
    valueOfDialogResult = False
    Me.Hide
End Sub

' [�u�������s] �{�^���������ꂽ�Ƃ��̏��������s
' �t�H�[������āADialogResult �v���p�e�B�� True �ɕύX���܂��B
Private Sub executeButon_Click()
    If Me.FindText = "" Then Exit Sub
    
    valueOfDialogResult = True
    Me.Hide
End Sub


