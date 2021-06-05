Option Explicit

Private start As StringNode
Private enumerator As Collection
Private ignore_case As Boolean


Property Get IgnoreCase() As Boolean
  IgnoreCase = ignore_case
End Property

Property Let IgnoreCase(Value As Boolean)
  ignore_case = Value
End Property


Private Function comp(lhs As String, rhs As String) As Boolean
  If ignore_case Then
    comp = UCase(lhs) > UCase(rhs)
  Else
    comp = lhs > rhs
  End If
End Function


Public Sub Add(str As String)
  Set enumerator = Nothing

  Dim new_node As StringNode
  Set new_node = New StringNode
  new_node.Value = str
  If start Is Nothing Then
    Set start = new_node
  Else
    Dim node As StringNode

    If comp(start.Value, str) Then
      Set new_node.NextNode = start
      Set start = new_node
      Exit Sub
    Else
      Set node = start
      Do Until node.NextNode Is Nothing
        If comp(node.NextNode.Value, str) Then
          Set new_node.NextNode = node.NextNode
          Set node.NextNode = new_node
          Exit Sub
        End If
        Set node = node.NextNode
      Loop
    End If
    Set node.NextNode = new_node
  End If
End Sub


Private Function work() As Collection
  If enumerator Is Nothing Then
    Set enumerator = New Collection
    Dim node As StringNode
    Set node = start
    Do Until node Is Nothing
      enumerator.Add node.Value
      Set node = node.NextNode
    Loop
  End If
  Set work = enumerator
End Function


'@Enumerator
Public Property Get NewEnum() As IUnknown
  Set NewEnum = work.[_NewEnum]
End Property


Private Sub Class_Terminate()
  Set enumerator = Nothing
  Clear
End Sub


Private Function nthNode(idx As Long) As StringNode
  Set nthNode = start
  Dim i As Long
  For i = 1 To idx
    Set nthNode = nthNode.NextNode
    If nthNode Is Nothing Then Err.Raise vbObjectError + 4, Application.Caller, "indexÇ™îÕàÕÇí¥Ç¶ÇƒÇ¢Ç‹Ç∑ÅD"
  Next
End Function


Public Property Get Item(idx As Long) As String
  Item = nthNode(idx).Value
End Property


Public Property Let Item(idx As Long, str As String)
  nthNode(idx).Value = str
  If Not enumerator Is Nothing Then
    enumerator(idx) = str
  End If
End Property

Sub Clear()
  Dim node As StringNode
  Do Until start Is Nothing
    Set node = start.NextNode
    Set start.NextNode = Nothing
    Set start = node
  Loop
End Sub