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

Public Function GetEnumerator() As IEnumerator

    Set GetEnumerator = New IEnumerator
    With GetEnumerator
        If IsObject(target) Then
            Set .targetList = target
        Else
            .targetList = target
        End If
        
            
        If TypeName(.targetList) = "Collection" Then
            .index = 0

        ElseIf TypeName(.targetList) = "Range" Then
            .index = 0
        Else
            .index = LBound(.targetList) - 1
        End If
        
    End With

End Function


