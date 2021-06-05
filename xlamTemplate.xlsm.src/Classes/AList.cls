Option Explicit
'@Folder Tools

' arrが使えない環境があるので，その代替

Private arr() As Variant
Private tail As Long
Private cap As Long
Private enumerator As Collection

Private Sub Class_Initialize()
  'コンストラクタ
  cap = 20
  tail = -1
  ReDim arr(cap)
End Sub


'@Ignore ProcedureNotUsed
Private Sub Class_Terminated()
  'デストラクタ
  Clear
End Sub


'@Ignore FunctionReturnValueAlwaysDiscarded
Function Add(ByVal Value As Variant) As Long
  tail = tail + 1
  If tail > cap Then
    cap = cap + 20
    ReDim Preserve arr(0 To cap)
  End If
  arr(tail) = Value
  If Not enumerator Is Nothing Then enumerator.Add Value
  Add = tail
End Function


'@Ignore ProcedureNotUsed
Sub AddRange(c As ICollection)
  Dim i
  For Each i In c
    Add i
  Next
End Sub


' CIS: Case InSensitive
Private Function BSearchCIS(ByVal Value As Variant, lb As Long, ub As Long) As Long
  If ub = lb Then
    If Value = UCase(arr(ub)) Then
      BSearchCIS = ub
    Else
      BSearchCIS = -(ub + 1)
    End If
  Else
    Dim pivot As Long
    pivot = (ub + lb) / 2
    If Value < UCase(arr(pivot)) Then
      BSearchCIS = BSearchCIS(Value, lb, pivot)
    Else
      BSearchCIS = BSearchCIS(Value, pivot, ub)
    End If
  End If
End Function



Private Function BSearch(ByVal Value As Variant, lb As Long, ub As Long) As Long
  If ub = lb Then
    If Value = arr(ub) Then
      BSearch = ub
    Else
      BSearch = -(ub + 1)
    End If
  Else
    Dim pivot As Long
    pivot = (ub + lb) / 2
    If Value < arr(pivot) Then
      BSearch = BSearch(Value, lb, pivot)
    Else
      BSearch = BSearch(Value, pivot, ub)
    End If
  End If
End Function


'@Ignore ProcedureNotUsed
Public Sub testBsearch()
  ReDim arr(0 To 5)
  Dim i As Long
  For i = 0 To 5
    arr(i) = i * 2 + 1
  Next
  ' 1, 3, 5 ,7, 9, 11

  Debug.Print 6
  Debug.Print BSearch(6, 0, 5)
  Debug.Print 10
  Debug.Print BSearch(10, 0, 5)

  ReDim arr(0)
End Sub



'@Ignore ProcedureNotUsed
Function BinarySearch(ByVal Value As Variant, _
                      Optional ByVal comparer As Long = 0, _
                      Optional ByVal index As Long = 0, _
                      Optional ByVal Count As Long = -1) As Long
  Dim last As Long
  If Count < 0 Then
    Count = tail + 1
    last = tail
  Else
    last = index + Count - 1
  End If
  If VarType(Value) = vbString And comparer = 0 Then
    BinarySearch = BSearchCIS(UCase(Value), index, last)
  Else
    BinarySearch = BSearch(Value, index, last)
  End If
End Function


Sub Clear()
  Set enumerator = Nothing
  Dim i As Long
  For i = 0 To tail
    If VarType(arr(i)) = vbObject Then Set arr(i) = Nothing
  Next
  tail = -1
End Sub


'@Ignore ProcedureNotUsed
Function Contains(ByVal Item As Variant) As Boolean
  Dim v
  For Each v In arr
    If v = Item Then
      Contains = True
      Exit Function
    End If
  Next
End Function


'@Ignore ProcedureNotUsed
Sub CopyTo(ByRef Dest As Variant, Optional index As Long = 0)
  If VarType(Dest) = vbArray Then
    If UBound(Dest) - index < tail Then
      Err.Raise vbObjectError + 1, "AList.CopyTo(Array)", "コピー先の配列のサイズが不足しています"
    End If
    Dim i As Long
    For i = 0 To tail
      If VarType(i) = vbObject Then
        Set Dest(i) = arr(i)
      Else
        Dest(i) = arr(i)
      End If
    Next
  End If
End Sub


Function IndexOf(ByVal Value As Variant, _
                 Optional ByVal startIndex As Long = 0, _
                 Optional ByVal Count As Long = -1) As Long
  Dim last As Long
  If Count < 0 Then
    last = tail
  Else
    last = startIndex + Count - 1
  End If
  For IndexOf = startIndex To last
    If arr(IndexOf) = Value Then Exit Function
  Next
  IndexOf = -1
End Function


'@Ignore ProcedureNotUsed
Sub Insert(ByVal index As Long, ByVal Value As Variant)
  Set enumerator = Nothing

  Add 0 ' 適当に一つ追加してから，内容をずらす
  Dim i
  For i = tail To index + 1 Step -1
    If VarType(arr(i - 1)) = vbObject Then
      Set arr(i) = arr(i - 1)
    Else
      arr(i) = arr(i - 1)
    End If
  Next
  If VarType(Value) = vbObject Then
    Set arr(index) = Value
  Else
    arr(index) = Value
  End If
End Sub

'@Ignore ProcedureNotUsed
Sub InsertRange(ByVal index As Long, c)
  Set enumerator = Nothing

  Dim sz As Long
  Select Case VarType(c)
  Case vbArray
    sz = UBound(c) - LBound(c) + 1
  Case vbObject
    sz = c.Count
  Case Else
    Err.Raise vbObjectError + 2, "AList.InsertRange", "配列かオブジェクトが必要です"
  End Select

  tail = tail + sz
  If tail > cap Then
    cap = cap + sz
    ReDim Preserve arr(0 To cap)
  End If

  Dim i As Long
  For i = 0 To sz - 1
    If VarType(arr(tail - i - sz)) = vbObject Then
      Set arr(tail - i) = arr(tail - sz - i)
    Else
      arr(tail - i) = arr(tail - sz - i)
    End If
  Next

  i = index
  Dim v
  For Each v In c
    If VarType(v) = vbObject Then
      Set arr(i) = v
    Else
      arr(i) = v
    End If
    i = i + 1
  Next
End Sub


'@Ignore ProcedureNotUsed
Function LastIndexOf(ByVal Value As Variant, _
                     Optional ByVal startIndex As Long = 0, _
                     Optional ByVal Count As Long = -1) As Long
  Dim last
  If Count < 0 Then
    Count = tail + 1
  Else
    last = startIndex + Count - 1
  End If

  LastIndexOf = -1

  Dim i As Long
  For i = last To startIndex Step -1
    If arr(i) = Value Then
      LastIndexOf = i
      Exit Function
    End If
  Next
End Function


'@Ignore ProcedureNotUsed
Sub Remove(ByVal obj As Variant)
  Dim i As Long
  For i = 0 To UBound(arr)
    If obj = arr(i) Then
      RemoveAt i
      Exit Sub
    End If
  Next
End Sub


Sub RemoveAt(ByVal index As Long)
  RemoveRange index, 1
End Sub

Sub RemoveRange(ByVal index As Long, ByVal Count As Long)
  Set enumerator = Nothing

  tail = tail - Count
  Dim i As Long
  For i = index To tail
    arr(i) = arr(i + Count)
  Next
End Sub


Sub Reverse(Optional ByVal index As Long = 0, _
            Optional ByVal Count As Long = -1)
  Set enumerator = Nothing

  Dim last As Long
  If Count < 0 Then
    last = tail
    Count = tail + 1
  Else
    last = index + Count
  End If

  Dim i As Long
  For i = 0 To (Count - 1) / 2
    Dim tmp
    If VarType(arr(index + i)) = vbObject Then
      Set tmp = arr(index + i)
    Else
      tmp = arr(index + i)
    End If
    If VarType(arr(last - i)) = vbObject Then
      Set arr(index + i) = arr(last - i)
    Else
      arr(index + i) = arr(last - i)
    End If
    If VarType(tmp) = vbObject Then
      Set arr(last - i) = tmp
    Else
      arr(last - i) = tmp
    End If
  Next
End Sub

'@Ignore ProcedureNotUsed
Sub SetRange(ByVal index As Long, c)
  Set enumerator = Nothing
  Dim v
  For Each v In c
    If VarType(v) = vbObject Then
      Set arr(index) = v
    Else
      arr(index) = v
    End If
    index = index + 1
  Next
End Sub

Private Function isSmall(a, b, ic As Boolean) As Boolean
  If ic Then
    isSmall = UCase(CStr(a)) < UCase(CStr(b))
  Else
    isSmall = a < b
  End If
End Function


Private Function QS(nodes As VariantNode, ic As Boolean) As VariantNode
  Dim pivot As VariantNode
  Dim s_h As VariantNode
  Dim s_t As VariantNode
  Dim l_h As VariantNode
  Dim l_t As VariantNode
  Set pivot = nodes

  If nodes Is Nothing Then Exit Function

  Dim node As VariantNode
  Set node = pivot.NextNode
  Do Until node Is Nothing
    If isSmall(node.Value, pivot.Value, ic) Then
      'smaller
      If Not s_t Is Nothing Then Set s_t.NextNode = node
      Set s_t = node
      If s_h Is Nothing Then Set s_h = node
    Else
      'Larger or equal
      If Not l_t Is Nothing Then Set l_t.NextNode = node
      Set l_t = node
      If l_h Is Nothing Then Set l_h = node
    End If
  Loop
  Set s_h = QS(s_h, ic)
  Set l_h = QS(l_h, ic)
  If Not s_h Is Nothing Then
    Set QS = s_h
    Set node = s_h
    Do Until node.NextNode Is Nothing
      Set node = node.NextNode
    Loop
    Set node.NextNode = pivot
  Else
    Set QS = pivot
    Set node = pivot
  End If
  If Not l_h Is Nothing Then Set node.NextNode = l_h
End Function



Sub Sort(Optional ByVal IgnoreCase As Boolean = False, _
         Optional ByVal index As Long = 0, _
         Optional ByVal Count As Long = -1)
  Set enumerator = Nothing

  Dim nodes As VariantNode
  Dim node As VariantNode

  ' copy to node chain
  Set nodes = New VariantNode
  Set node = nodes
  node.Value = arr(index)
  Dim i As Long
  For i = 1 To Count - 1
    Set node.NextNode = New VariantNode
    Set node = node.NextNode
    node.Value = arr(index + i)
  Next

  ' Quick Sort
  Set nodes = QS(nodes, IgnoreCase)

  ' restore
  Set node = nodes
  For i = 0 To Count - 1
    arr(index + i) = node.Value
    Set node = node.NextNode
  Next

  ' free nodes
  Do Until nodes Is Nothing
    Set node = nodes
    Set nodes = node.NextNode
    Set node.NextNode = Nothing
  Loop
End Sub


Function ToArray() As Variant
  Dim toA()
  ReDim toA(0 To tail)
  Dim i As Long
  For i = 0 To tail
    If VarType(arr(i)) = vbObject Then
      Set toA(i) = arr(i)
    Else
      toA(i) = arr(i)
    End If
  Next
  ToArray = toA
End Function


'@Ignore ProcedureNotUsed
Function ToString() As String
  Dim sa() As String
  ReDim sa(0 To tail)
  Dim i As Long
  For i = 0 To tail
    sa(i) = CStr(arr(i))
  Next
  ToString = "[" & Join(sa, ", ") & "]"
End Function


'@Ignore ProcedureNotUsed
Sub TrimToSize()
  cap = tail
  ReDim Preserve arr(0 To cap)
End Sub


'@Ignore ProcedureNotUsed
Public Property Get Item(ByVal index As Long) As Variant
  If VarType(arr(index)) = vbObject Then
    Set Item = arr(index)
  Else
    Item = arr(index)
  End If
End Property


'@Ignore ProcedureNotUsed
Public Property Let Item(ByVal index As Long, ByVal Value As Variant)
  arr(index) = Value
  If Not enumerator Is Nothing Then enumerator(index) = Value
End Property


'@Ignore ProcedureNotUsed
Public Property Set Item(ByVal index As Long, ByVal Value As Object)
  Set arr(index) = Value
  If Not enumerator Is Nothing Then Set enumerator(index) = Value
End Property


'@Ignore ProcedureNotUsed
Public Property Get capacity() As Long
  capacity = cap + 1
End Property


'@Ignore ProcedureNotUsed
Public Property Let capacity(ByVal Value As Long)
  cap = Value - 1
  If tail > cap Then
    tail = cap
    Set enumerator = Nothing
  End If
  ReDim Preserve arr(0 To cap)
End Property


Public Property Get Count() As Long
  Count = tail + 1
End Property


'@Enumerator
'@Ignore ProcedureNotUsed
Public Property Get NewEnum() As IEnumVARIANT
  If enumerator Is Nothing Then
    Set enumerator = New Collection
    Dim i As Long
    For i = 0 To tail
      enumerator.Add arr(i)
    Next
  End If
  Set NewEnum = enumerator.[_NewEnum]
End Property


'@DefaultMember
'@Ignore ProcedureNotUsed
Public Property Get ItemObj(ByVal index As Long) As Variant
  Set ItemObj = arr(index)
End Property


'@DefaultMember
Public Property Get ItemVal(ByVal index As Long) As Variant
  ItemVal = arr(index)
End Property


'@Ignore ProcedureNotUsed
Public Property Set ItemObj(ByVal index As Long, ByVal obj As Object)
  Set arr(index) = obj
  If Not enumerator Is Nothing Then Set enumerator(index) = obj
End Property


'@Ignore ProcedureNotUsed
Public Property Let ItemVal(ByVal index As Long, ByVal val As Variant)
  arr(index) = val
  If Not enumerator Is Nothing Then enumerator(index) = val
End Property