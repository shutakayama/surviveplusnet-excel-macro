VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IEnumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public targetList As Variant 'Range
Public Current As Variant 'Range
Public index As Integer
Public f As FuncDelegate


Public Function MoveNext() As Boolean

    Select Case TypeName(targetList)
    Case "IEnumerator"
        If f Is Nothing Then
            MoveNext = targetList.MoveNext
            If IsObject(targetList.Current) Then
                Set Me.Current = targetList.Current
            Else
                Me.Current = targetList.Current
            End If
            Exit Function
        Else
            While targetList.MoveNext
                If f.Run(targetList.Current) Then
                    MoveNext = True
                    If IsObject(targetList.Current) Then
                        Set Me.Current = targetList.Current
                    Else
                        Me.Current = targetList.Current
                    End If
                    Exit Function
                End If
            Wend
            MoveNext = False
            Exit Function
        End If
    
    Case Else
        index = index + 1
        Dim count As Integer
        If IsObject(targetList) Then
            count = targetList.count
        Else
            count = UBound(targetList)
        End If
        
        If count < index Then
            MoveNext = False
            Exit Function
        Else
            MoveNext = True
            If IsObject(targetList(index)) Then
                Set Current = targetList(index)
            Else
                Current = targetList(index)
            End If
            Exit Function
        End If

    End Select

End Function

'Public Function MoveNextWhere(w As FuncDelegate) As Boolean
'    While MoveNext()
'        If w.Run(Current) Then
'            MoveNextWhere = True
'            Exit Function
'        End If
'    Wend
'    MoveNextWhere = False
'    Exit Function
'End Function
