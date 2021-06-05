Option Explicit

Private start As VariantNode
Private enumerator As Collection


Public Sub Add(Value As Variant)
  Set enumerator = Nothing

  Dim new_node As VariantNode
  Set new_node = New VariantNode
  If VarType(Value) = vbObject Then
    Set new_node.Value = Value
  Else
    new_node.Value = Value
  End If

  If start Is Nothing Then
    Set start = new_node
  ElseIf start.Value > Value Then
    Set new_node.NextNode = start
    Set start = new_node
  Else
    Dim node As VariantNode
    Set node = start
    Do Until node.NextNode Is Nothing
      If node.NextNode.Value > Value Then
        Set new_node.NextNode = node.NextNode
        Set node.NextNode = new_node
        Exit Sub
      End If
      Set node = node.NextNode
    Loop
    Set node.NextNode = new_node
  End If
End Sub


Public Property Get Count() As Long
  Count = work.Count
End Property


Private Function work() As Collection
  If enumerator Is Nothing Then
    Set enumerator = New Collection
    Dim node As VariantNode
    Set node = start
    Do Until node Is Nothing
      enumerator.Add node.Value
      Set node = node.NextNode
    Loop
  End If
  Set work = enumerator
End Function


'@Enumerator
'@Ignore ProcedureNotUsed
Public Property Get NewEnum() As IEnumVARIANT
  Set NewEnum = work.[_NewEnum]
End Property


Private Sub Class_Terminate()
  Set enumerator = Nothing
  Clear
End Sub


Private Function nthNode(idx As Long) As VariantNode
  Set nthNode = start
  Dim i As Long
  For i = 1 To idx
    If nthNode.NextNode Is Nothing Then
      Err.Raise vbObjectError + 3, Application.Caller, "Index‚ª—Ìˆæ‚ð’´‚¦‚Ä‚¢‚Ü‚·"
    End If
    Set nthNode = nthNode.NextNode
  Next
End Function


Public Property Get Item(idx As Long) As Variant
  Dim node As VariantNode
  Set node = nthNode(idx)
  If VarType(node.Value) = vbObject Then
    Set Item = node.Value
  Else
    Item = node.Value
  End If
End Property


Public Property Let Item(idx As Long, Value As Variant)
  If VarType(Value) = vbObject Then
    Set nthNode(idx).Value = Value
    If Not enumerator Is Nothing Then Set enumerator(idx) = Value
  Else
    nthNode(idx).Value = Value
    If Not enumerator Is Nothing Then enumerator(idx) = Value
  End If
End Property

Sub Clear()
  Dim node As VariantNode
  Do Until start Is Nothing
    Set node = start
    Set start = node.NextNode
    Set node.NextNode = Nothing
  Loop
End Sub