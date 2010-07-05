VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CopyWorkSheetDialogForm 
   Caption         =   "���[�N�V�[�g�𕡐��R�s�["
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5205
   OleObjectBlob   =   "CopyWorkSheetDialogForm.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "CopyWorkSheetDialogForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' DialogResult �v���p�e�B�̃o�b�L���O�t�B�[���h�ł��B
Private valueOfDialogResult As Boolean

'------------------------------------------------------------------------
' NewNames �v���p�e�B
'------------------------------------------------------------------------
'   �ړI�F      ���[�U�[�ɂ���Ďw�肳�ꂽ�V�������[�N�V�[�g�̖��O�̔z����擾�܂��͐ݒ肵�܂��B
'   ���߁F
'   �ύX�����F  s-koga 2010/07/05 - �I���W�i��
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
' DialogResult �v���p�e�B
'------------------------------------------------------------------------
'   �ړI�F      ���[�U�[��OK�{�^�������������ǂ������擾�܂��͐ݒ肵�܂��B
'   ���߁F
'   �ύX�����F  s-koga 2010/07/05 - �I���W�i��
'------------------------------------------------------------------------
Public Property Get DialogResult() As Boolean
    ' Return
    DialogResult = valueOfDialogResult
End Property
Public Property Let DialogResult(ByVal vNewValue As Boolean)
    valueOfDialogResult = vNewValue
End Property

' �L�����Z���{�^���������ꂽ�Ƃ��̏��������s���܂��B
Private Sub cancelButton_Click()
    Me.DialogResult = False
    Call Me.Hide
End Sub

' OK�{�^���������ꂽ�Ƃ��̏��������s���܂��B
Private Sub okButton_Click()
    If newNamesBox.Text = "" Then
        newNamesBox.SetFocus
        Exit Sub
    End If
    
    Me.DialogResult = True
    Call Me.Hide
End Sub

'------------------------------------------------------------------------
'   �ړI�F      �_�C�A���O��\�����AOK�������ꂽ���ǂ�����Ԃ��܂�
'   �߂�l�F    ���[�U�[��OK�{�^�������������ǂ����� Boolean �Ƃ��ĕԂ��܂��B
'   ���߁F
'   �g�p���@�FIf dialog.ShowDialog() Then
'------------------------------------------------------------------------
Public Function ShowDialog() As Boolean
    Me.Show (vbModal)
    ShowDialog = Me.DialogResult
End Function

Private Sub UserForm_Click()

End Sub
