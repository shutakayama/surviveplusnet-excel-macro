VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEnumerable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public target As Variant 'Range
Public f As FuncDelegate

Public Function GetEnumerator() As IEnumerator

    Set GetEnumerator = New IEnumerator
    With GetEnumerator
        If IsObject(target) Then
            Set .targetList = target
        Else
            .targetList = target
        End If
        
        Select Case TypeName(.targetList)
        Case "IEnumerator"
            Set .f = f
        Case "Collection", "Range"
            .index = 0
        Case Else
            .index = LBound(.targetList) - 1
        End Select
        
    End With

End Function


Public Function Where(f As FuncDelegate) As IEnumerable

    Set Where = New IEnumerable
    With Where
        Set .target = Me.GetEnumerator
        Set .f = f
    End With
End Function

Public Sub ForEach(a As ActionDelegate)
    
    Dim enumerator As IEnumerator: Set enumerator = Me.GetEnumerator()
    While enumerator.MoveNext()
         a.Run (enumerator.Current)
    Wend

End Sub
