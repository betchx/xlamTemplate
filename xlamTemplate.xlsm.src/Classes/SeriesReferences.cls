'@Folder("Personal")
Option Explicit

'@ModuleDescription("系列の文字列を処理するクラス．参照先の列を変更するなどが可能")

Private comps As Variant


'@Description("系列のFormulaに代入できる数式（A1形式）：Get, Let")
'@DefaultMember
Public Property Get Formula() As String
  Formula = "=" & comps(0) & "(" & comps(1)
  Dim i As Long
  For i = 2 To UBound(comps)
    Formula = Formula & "," & comps(i)
  Next
  Formula = Formula & ")"
End Property


Public Property Let Formula(ByVal Formula As String)
  comps = deFunc(Formula)
End Property


'@Description("系列のFormulaに代入できる数式(R1C1形式)：Get")
Public Property Get FormulaR1C1() As String
  FormulaR1C1 = "=" & comps(0) & "(" & toR1C1(comps(1))
  Dim i As Long
  For i = 2 To 3
    FormulaR1C1 = FormulaR1C1 & "," & toR1C1(comps(i))
  Next
  For i = 4 To UBound(comps)
    FormulaR1C1 = FormulaR1C1 & "," & comps(i)
  Next
  FormulaR1C1 = FormulaR1C1 & ")"
End Property


'@Description("X軸（カテゴリー軸）の参照先を表す文字列：Get, Let")
Public Property Get XAddress() As String
  XAddress = comps(2)
End Property


Public Property Let XAddress(ByVal val As String)
  comps(2) = val
End Property


'@Description("Y軸（値軸）の参照先を表す文字列：Get, Let")
Public Property Get YAddress() As String
  YAddress = comps(3)
End Property


Public Property Let YAddress(ByVal val As String)
  comps(3) = val
End Property


'@Description("系列のタイトルを表す文字列：Get, Let")
Public Property Get Title() As String
  Title = comps(1)
End Property


Public Property Let Title(ByVal val As String)
  comps(1) = val
End Property


'@Description("X軸（カテゴリー軸）の参照先を示すRangeオブジェクト：Get, Set")
Public Property Get x() As Range
  Set x = Range(XAddress)
End Property


Public Property Set x(ByVal rng As Range)
  Me.XAddress = FullAddress(rng)
End Property


'@Description("Y軸（値軸）の参照先を示すRangeオブジェクト：Get, Set")
Public Property Get Y() As Range
  Set Y = Range(YAddress)
End Property


Public Property Set Y(ByVal rng As Range)
  YAddress = FullAddress(rng)
End Property


'@Description("セル参照文字列をシート名付きのR1C1形式に変換する．")
Private Function toR1C1(ByVal adr As String) As String
  toR1C1 = FullAddressR1C1(Range(adr))
End Function


'@Description("系列の値とタイトルの参照先の列をCountだけ右にずらす．")
Public Sub OffsetColumn(Optional ByVal Count As Long = 1)
  YAddress = FullAddress(Y.Offset(0, Count))
  If Left$(Me.Title, 1) <> """" Then
    Me.Title = FullAddress(Range(Me.Title).Offset(0, Count))
  End If
End Sub

