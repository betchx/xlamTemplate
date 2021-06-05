Option Explicit
'@Folder Tools

Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal Hwnd As Long) As Long
Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
Declare PtrSafe Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" _
  (ByVal lpString As String) As Long
Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
  (Destination As Any, Source As Any, ByVal Length As Long)
Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long


'@Description("ActiveSheetへの参照を返す．インテリセンスが有効．")
Function ActSht() As Worksheet
  Set ActSht = ActiveSheet
End Function


'@Description("レジストリのAPP名．マクロのあるブックの名前を用いる．")
Function APP() As String
  APP = ThisWorkbook.Name
End Function


'@Description("シート名を含めたアドレスをA1形式で返す．")
Function FullAddress(ByVal rng As Range) As String
  FullAddress = getSheetName(rng) & "!" & rng.Address(True, True, xlA1)
End Function


'@Description("シート名を含めたアドレスをR1C1形式で返す．")
Function FullAddressR1C1(ByVal rng As Range) As String
  FullAddressR1C1 = getSheetName(rng) & "!" & rng.Address(True, True, xlR1C1)
End Function


'@Description("引数のRangeの親となるシートの名前を返す")
Function getSheetName(ByVal rng As Range) As String
  Dim sn As String
  sn = rng.Parent.Name
  With CreateObject("VBScript.RegExp")
    .Pattern = "^[.0-9]|^[A-Za-z]{1,3}[0-9]|[-&!""#$%=({}+~^|,;@`<>]"
    If .test(sn) Then
      getSheetName = "'" & rng.Parent.Name & "'"
    Else
      getSheetName = rng.Parent.Name
    End If
  End With
End Function


'@Description("A1形式のアドレスをR1C1形式に変換する．ConvertFormulaのラッパー．")
Function A1toR1C1(ByVal a1 As String) As String
  A1toR1C1 = Application.ConvertFormula(a1, xlA1, xlR1C1)
End Function


'@Description("A1形式のアドレスをR1C1形式に変換する．ConvertFormulaのラッパー．")
Function R1C1toA1(ByVal r1c1 As String) As String
  R1C1toA1 = Application.ConvertFormula(r1c1, xlR1C1, xlA1)
End Function


'@Description("ポイントをmmに変換する．")
Function pnt2mm(ByVal p As Single) As Single
  pnt2mm = CSng(p / Application.CentimetersToPoints(0.1))
End Function


'@Description("ポイントをcmに変換する．")
Function pnt2cm(ByVal pnt As Single) As Single
  pnt2cm = CSng(pnt / Application.CentimetersToPoints(1#))
End Function


'@Description("mmをポイントに変換する．")
Function mm2pnt(ByVal mm As Single) As Single
  mm2pnt = CSng(Application.CentimetersToPoints(mm * 0.1))
End Function


'@Description("cmをポイントに変換する．Application.CentimetersToPointsのほぼ別名")
Function cm2pnt(ByVal cm As Single) As Single
  cm2pnt = CSng(Application.CentimetersToPoints(cm))
End Function


'@Description("Rangeの参照先の値を整数（Long）で返す")
Function r2i(ByVal r As Range) As Long
  r2i = CInt(val(r.Value))
End Function


'@Description("Rangeの参照先の値を単精度浮動小数点数（Single）で返す")
Function r2s(ByVal r As Range) As Single
  r2s = CSng(val(r.Value))
End Function


'@Description("引数の名前の参照先の値を単精度浮動小数点数（Single）で返す．iはn2rに渡される．")
Function n2s(ByVal n As String, Optional ByVal i As Long = 0) As Single
  n2s = r2s(n2r(n, i))
End Function


'@Description("名前からレンジを取得する．iがある場合は範囲のi番目のセルを返す")
' まずはアクティブシートで探し，見つからない場合はワークブックから探すことになる．
Function n2r(ByVal n As String, Optional ByVal i As Long = 0) As Range
  If i > 0 Then
    Dim r As Range
    Set r = n2r(n)
    If Not r Is Nothing Then
      If r.Count >= i Then Set n2r = r.Cells(i)
    End If
    Exit Function
  End If

  On Error GoTo kkk:
  Set n2r = Range(n)
kkk:
  ' Note: 当初はループで探していたが，遅いことが判明したので，
  '       エラー上等で直接参照する形に変更した．
End Function


'@Description("名前オブジェクトを探して返す")
Function s2n(ByVal str As String) As Name
  On Error Resume Next
  Set s2n = ActSht.Names(str)
  If s2n Is Nothing Then
    Set s2n = ActiveWorkbook.Names(str)
  End If
  On Error GoTo 0
End Function


'@Description("範囲を表す文字列から範囲を返す．シート名があっても動作する．別ブックは無理．")
Function s2r(ByVal str As String) As Range
  On Error GoTo eee:
  If str Like "*!*" Then
    Dim a
    a = Split(str, "!")
    Set s2r = ActiveWorkbook.Worksheets(a(0)).Range(a(1))
  Else
    Set s2r = ActiveSheet.Range(str)
  End If
eee:
End Function


'@Description("関数を表す文字列を1段階だけ分解して，文字列の配列に変換して返す．index=0が関数名，以降が引数．")
Function deFunc(ByVal Formula As String) As String()
  Dim num_p        As Long: num_p = 0
  Dim ip           As Long: ip = 1
  Dim s            As String: s = ""
  Dim first_p      As Long: first_p = InStr(1, Formula, "(")

  Dim parts(0 To 30) As String
  If Left$(Formula, 1) = "=" Then
    parts(0) = Mid$(Formula, 2, first_p - 2)
  Else
    parts(0) = Left$(Formula, first_p - 1)
  End If

  ' 分解
  Dim idx          As Long
  For idx = first_p + 1 To Len(Formula) - 1
    Dim c            As String
    c = Mid$(Formula, idx, 1)
    If num_p = 0 And c = "," Then
      parts(ip) = s
      s = ""
      ip = ip + 1
    Else
      s = s & c
      If c = "(" Then num_p = num_p + 1
      If c = ")" Then num_p = num_p - 1
    End If
  Next
  parts(ip) = s

  Dim ans()        As String
  ReDim ans(0 To ip)
  For idx = 0 To ip
    ans(idx) = parts(idx)
  Next
  deFunc = ans
End Function

'@Description("deFuncで分解した配列を組み立てて関数の文字列にする．")
Function ReFunc(arr As Variant) As String
  If VarType(arr) > vbArray Then
    ReFunc = arr(0) & "("
    Dim i As Long
    For i = 1 To UBound(arr)
      If i > 1 Then ReFunc = ReFunc & ","
      ReFunc = ReFunc & arr(i)
    Next
    ReFunc = ReFunc & ")"
  ElseIf VarType(arr) = vbString Then
    ReFunc = arr & "()"
  Else
    ReFunc = CStr(arr) & "()"
  End If
End Function


'@Description("引数のOfficeオブジェクトとその親からChartObjectを探して返す．見つからない時はNothing.")
Function FindChartObject(ByVal start_from As Object) As ChartObject
  Dim obj As Object
  Set obj = start_from
  If obj Is Nothing Then Exit Function
  On Error GoTo eee:
  '@Ignore VariableNotUsed
  Dim i As Long
  For i = 1 To 20    '最大20階層まで（永久ループ防止）
    If TypeOf obj Is Excel.Application Then Exit Function
    Set obj = obj.Parent
    If TypeOf obj Is ChartObject Then
      Set FindChartObject = obj
      Exit Function
    End If
  Next
eee:
End Function


'@Description("共用するFileSystemObjectを返す．")
Function FSO() As FileSystemObject
  Static fso_ As FileSystemObject
  If fso_ Is Nothing Then Set fso_ = New FileSystemObject
  Set FSO = fso_
End Function


'@Description("With文とともに使用して，不要な処理を抑制してマクロの実行速度を増加させる．")
Function Boost(Optional ByVal イベント抑制 As Boolean = False, _
               Optional ByVal 再計算抑制 As Boolean = True, _
               Optional ByVal プリンタ通信抑制 As Boolean = True, _
               Optional ByVal 画面更新抑制 As Boolean = True) As MacroBooster
  Set Boost = New MacroBooster
  With Boost
    If イベント抑制 Then .DisableEvents
    If 再計算抑制 Then .NoCalc
    If プリンタ通信抑制 Then .NoPrinterComunication
    If 画面更新抑制 Then .NoScreenUpdating
    .Enter
  End With
End Function


'@Description("引数のオブジェクトの親をたどってチャートを返す．チャートがなかったらNothing")
Public Function GetChart(ByVal arg As Object) As Chart
  Dim o            As Object
  Set o = arg
  Do Until (TypeName(o) = "Application")
    If TypeName(o) = "Chart" Then
      Set GetChart = o
      Exit Function
    End If
    Set o = o.Parent
  Loop
End Function


'@Description("引数のオブジェクトの親をたどってチャートオブジェクトを返す．チャートオブジェクトがなかったらNothing")
Function GetChartObject(ByVal arg As Object) As ChartObject
  Dim o As Object
  Set o = arg
  Do Until (TypeName(o) = "Application")
    If TypeName(o) = "ChartObject" Then
      Set GetChartObject = o
      Exit Function
    End If
    Set o = o.Parent
  Loop
End Function


'@Description("Chartをフロート配置にするときの定数がわかりずらいために作成したもの")
Public Function フロート配置設定() As XlPlacement
  フロート配置設定 = xlMove
End Function


'@Description セルのフォーマット文字列を取得する関数．引数省略時には選択範囲を対象とする．
Function get_format(Optional ByRef r As Range = Nothing) As String
  If r Is Nothing Then Set r = Selection
  If IsNull(r.NumberFormat) Then
    Dim s          As String
    If r.Count > 1 Then
      s = CStr(r(1).Value)
    Else
      s = r.Value
    End If

    Dim u          As Variant
    For Each u In Array("1", "2", "3", "4", "5", "6", "7", "8", "9")
      s = Replace(s, u, "0")
    Next
    get_format = Replace(s, "-", "")
  Else
    get_format = r.NumberFormatLocal
  End If
End Function


Function Prof(Optional Enabled As Boolean = True) As IProfiler
  If Enabled Then
    Set Prof = New SimpleProfiler
  Else
    Set Prof = New DummyProfiler
  End If
End Function

' 使い方の例
Private Sub test_Prof()

  With Prof
    MsgBox "Press"
    .Mark "1"
    Dim s
    s = InputBox("Enter")
    .Mark "InputBox"
    MsgBox s
  End With

End Sub

' @Description("切り取りやコピーのモードの時に，破線で囲われている範囲(Range)を取得する．")
Function CutCopyOrigin() As Range
' 参照：  http://www2.aqua-r.tepm.jp/~hironobu/ke_m13.htm   E03M121(Excel2003)
  If Application.CutCopyMode = False Then Exit Function

  Dim mem&, sz&, lk&, vv As Variant, buf$
  OpenClipboard 0&
  mem = GetClipboardData(RegisterClipboardFormat("Link"))
  CloseClipboard

  If mem = 0 Then Exit Function

  sz = GlobalSize(mem)
  lk = GlobalLock(mem)
  buf = String(sz, vbNullChar)
  CopyMemory ByVal buf, ByVal lk, sz
  GlobalUnlock mem

  vv = Split(buf, vbNullChar)
  buf = "'" & vv(1) & "'!" & Application.ConvertFormula(vv(2), xlR1C1, xlA1)
  Set CutCopyOrigin = Range(buf)
End Function