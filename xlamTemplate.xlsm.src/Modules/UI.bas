Option Explicit
'@Folder RibbonUI
' moduleImageMso: UserRolesManage
' moduleTarget: Range

'@VariableDescription マクロのツリー情報を保持する
Private macroTree      As MacroProject

'@VariableDescription XML ログの出力先ファイル名
Const last_xml_file As String = "latest.xml"

'レジストリのセクション名
Const sec          As String = "UI"
Const SHORTCUT_SEC As String = "ShortCuts"
'
'=======================================================================================
'


Public Sub UI_onLoad(ByRef iRibbon As IRibbonUI)
  Set Ribbon.Item = iRibbon
End Sub


' idつきでxmlを作成する．
Public Sub XT_DynMenu_getContent(control As IRibbonControl, ByRef returnedVal As Variant)
  returnedVal = CreateXML(control.tag)
  Ribbon.InvalidateControl control.id
End Sub


Public Sub XT_SB_getScreentip(control As IRibbonControl, ByRef returnedVal As Variant)
  returnedVal = getInfo(control.tag).screenTip
End Sub


Public Sub XT_groupLabel(control As IRibbonControl, ByRef returnedVal As Variant)
  returnedVal = ThisWorkbook.Name
End Sub


Public Sub XT_SB_getLabel(control As IRibbonControl, ByRef returnedVal As Variant)
  Dim info         As MacroInfo
  Set info = getInfo(control.tag)
  If info.label <> "" Then
    returnedVal = info.label
  Else
    returnedVal = info.Name
  End If
End Sub


Public Sub XT_SB_getSupertip(control As IRibbonControl, ByRef returnedVal As Variant)
  returnedVal = Replace(getInfo(control.tag).superTip, "||", vbCrLf)
End Sub


Public Sub XT_SB_getImage(control As IRibbonControl, ByRef returnedVal As Variant)
  Dim info         As MacroInfo
  Set info = getInfo(control.tag)
  Select Case info.ImageType
    Case XT_ImageTypeMso
      ' SB内ボタンからの場合はこちらに来る可能性あり．
      returnedVal = info.Image
    Case XT_ImageTypeExternal
      If FSO.FileExists(info.Image) Then
        On Error GoTo nnn:
        Set returnedVal = LoadPicture(info.Image)
      Else
nnn:
        returnedVal = "MacroDefault"    ' pictureが見つからない場合
      End If
    'Case XT_IMageTypeNone
    Case Else
      returnedVal = "MacroDefault"    ' 設定がないのでデフォルト
  End Select
End Sub


' 未登録なら表示しない
Public Sub XT_SB_getVisible(control As IRibbonControl, ByRef returnedVal As Variant)
  Dim macro As String
  macro = GetSetting(APP, sec, control.tag, "")
  returnedVal = macro <> ""
End Sub


' 登録ボタンから実行する
Public Sub XT_SB_onAction(control As IRibbonControl)
  Dim macro As String
  macro = GetSetting(APP, sec, control.tag, "")
  If macro <> "" Then
    Application.Run macro
    '設定の順番変更
    If Right$(control.tag, 1) <> "0" Then
      Dim i As Long
      i = CLng(control.tag)
      Do
        SaveSetting APP, sec, CStr(i), GetSetting(APP, sec, CStr(i - 1), "")
        Ribbon.InvalidateControl "XT_SB" & CStr(i)
        DoEvents
        i = i - 1
      Loop While i Mod 10 > 0
      SaveSetting APP, sec, CStr(i), macro
      Ribbon.InvalidateControl "XT_SB" & CStr(i \ 10)
      Ribbon.InvalidateControl "XT_SB" & CStr(i)
    End If
  End If
End Sub


' 引数tagから MacroInfoを取得する．
' tagが数字の場合は，レジストリからマクロ名を取得する．
' tagの先頭1文字が数字の場合はそれを無視して検索する．
' 数字以外から始まる場合は，マクロ名とみなして検索する．
Private Function getInfo(ByVal tag As String) As MacroInfo
  If macroTree Is Nothing Then initTree

  Dim s            As String
  If tag = "" Then
    s = ""
  ElseIf tag Like "##" Then
    s = GetSetting(APP, sec, tag, "")
  ElseIf InStr("0123456789", Left$(tag, 1)) > 0 Then
    s = Mid$(tag, 2)
  Else
    s = tag
  End If

  If macroTree.hasMacro(s) Then
    Set getInfo = macroTree.macro(s)
  Else
    Set getInfo = New MacroInfo
    If 履歴のクリーンアップ() Then
      Call InvaditateMenu
    End If
  End If
End Function


' 呼び出し元のリボンコントロールのもつタグをもとにマクロを呼び出す
Sub CallByTag(control As IRibbonControl)
  Dim num          As String
  Dim cmd          As String
  If control.tag <> "" Then
    If control.tag Like "#*" Then
      num = Left$(control.tag, 1)
      cmd = Mid$(control.tag, 2)
    Else
      num = ""
      cmd = control.tag
    End If
    If macroTree.hasMacro(cmd) Then Application.Run cmd

    'numがあれば実行したマクロの保存
    If num <> "" Then
      If GetSetting(APP, sec, num & "0", "") <> cmd Then
        Dim k As Long
        For k = 0 To 9
          If GetSetting(APP, sec, num & CStr(k), "") = cmd Then Exit For
        Next
        Dim i As Long
        For i = k To 1 Step -1
          SaveSetting APP, sec, num & CStr(i), GetSetting(APP, sec, num & CStr(i - 1), "")
        Next
        SaveSetting APP, sec, num & "0", cmd
        Call InvaditateMenu
      End If
    End If
  End If
End Sub


'インターフェースの変換
Private Function CNode(ByVal ref As Variant) As MSXML2.IXMLDOMNode
  '@Ignore SetAssignmentWithIncompatibleObjectType
  Set CNode = ref
End Function


' 引数をもとに，ボタンのXMLを作成する．
Private Function Button(xmlDoc As DOMDocument30, ByVal btn_id As String, ByVal label As String, ByVal onAction As String, _
                        Optional ByVal screenTip As String = "", Optional ByVal superTip As String = "", Optional ByVal tag As String = "", _
                        Optional ByVal description As String = "", Optional ByVal imageMso As String = "") As IXMLDOMElement
  Set Button = xmlDoc.createElement("button")
  With Button
    .setAttribute "id", btn_id
    .setAttribute "label", label
    .setAttribute "onAction", onAction
    If screenTip <> "" Then .setAttribute "screentip", screenTip
    If superTip <> "" Then .setAttribute "supertip", superTip
    If tag <> "" Then .setAttribute "tag", tag
    If description <> "" Then .setAttribute "description", description
    If imageMso <> "" Then .setAttribute "imageMso", imageMso
  End With
End Function


' マクロ情報クラスの情報からボタンのXMLを作成する．
Private Function ButtonFromInfo(ByRef xmlDoc As DOMDocument30, ByRef info As MacroInfo, Optional ByVal tagPrefix As String = "") As IXMLDOMElement
  Set ButtonFromInfo = xmlDoc.createElement("button")
  With ButtonFromInfo
    .setAttribute "id", info.id
    Dim label As String
    If info.label <> "" Then
      label = info.label
    Else
      label = info.Name
    End If
    If info.onKey <> "" Then
      label = label & " (" & FormatShortcutKey(info.onKey) & ")"
    End If
    .setAttribute "label", label
    If info.Name <> "" Then .setAttribute "onAction", "CallByTag"
    If info.screenTip <> "" Then .setAttribute "screentip", info.screenTip
    If info.superTip <> "" Then .setAttribute "supertip", Replace(info.superTip, "||", "&#13;")
    If info.tag <> "" Then .setAttribute "tag", tagPrefix & info.tag
    If info.desc <> "" Then
      .setAttribute "size", "large"
      .setAttribute "description", info.desc
    End If

    Select Case info.ImageType
      Case XT_ImageTypeMso
        .setAttribute "imageMso", info.Image
      Case XT_ImageTypeEmbedded
        .setAttribute "image", info.Image
      Case XT_ImageTypeExternal
        .setAttribute "getImage", "XT_SB_getImage"
      Case XT_ImageTypeInternal, XT_ImageTypeInternalAutomatic
        .setAttribute "getImage", "XT_SB_getImage"
      Case XT_IMageTypeNone
        .setAttribute "imageMso", "MacroDefault"
    End Select
  End With
End Function


' xmlの構築
Private Function CreateXML(control_tag As String) As String
  Const ns         As String = "http://schemas.microsoft.com/office/2006/01/customui"
  Debug.Assert Len(control_tag) = 1
  If macroTree Is Nothing Then Call initTree
  'ショートカットキーを取得
  loadShortcutKeys

  ' MSXML v3.0を使用する．  v6.0を使うと xmlns="" が勝手に追加されて面倒．
  Dim xmlDoc As DOMDocument30
  Set xmlDoc = New DOMDocument30
  Dim menu As IXMLDOMElement
  Set menu = xmlDoc.createElement("menu")

  menu.setAttribute "xmlns", ns
  menu.setAttribute "itemSize", "normal"
  xmlDoc.appendChild CNode(menu)

  Dim s As String
  s = GetSetting(APP, sec, "00", "")
  If s <> "" Then
    ' 最近使ったマクロ
    Dim recent As IXMLDOMElement
    Set recent = xmlDoc.createElement("menu")
    With recent
      .setAttribute "id", ThisWorkbook.Name & "_recent"
      .setAttribute "label", "最近使用したマクロ"
      .setAttribute "imageMso", "RefreshIntervalMenu"
      '.setAttribute "supertip", "直近で使用した最大10種類のマクロ"
      Dim i As Long
      For i = 0 To 9
        Dim tag As String
        tag = "0" & CStr(i)
        s = GetSetting(APP, sec, tag, "")
        If s <> "" Then
          Dim mbtn As IXMLDOMElement
          Set mbtn = ButtonFromInfo(xmlDoc, getInfo(tag))
          ' idの重複対策
          mbtn.setAttribute "id", mbtn.getAttribute("id") + "_r"
          .appendChild CNode(mbtn)
        End If
      Next
    End With
    menu.appendChild recent
    '区切り線
    Dim sep As IXMLDOMElement
    Set sep = xmlDoc.createElement("menuSeparator")
    sep.setAttribute "id", ThisWorkbook.Name & "_" & "recent_sepa"
    menu.appendChild CNode(sep)
  End If

  Dim module As MacroModule
  Dim submenus As Dictionary
  Set submenus = New Dictionary
  Dim keys As SortedStringList
  Set keys = New SortedStringList
  For Each module In macroTree
    If module.isTarget(Selection) Then
      Dim submenu As IXMLDOMElement
      Set submenu = xmlDoc.createElement("menu")
      With submenu
        .setAttribute "id", ThisWorkbook.Name & "_" & module.Name
        .setAttribute "label", module.Name
        Select Case module.ImageType
          Case XT_IMageTypeNone
            .setAttribute "imageMso", "MacrosGallery"
          Case XT_ImageTypeInternal, XT_ImageTypeInternalAutomatic, XT_ImageTypeExternal
            .setAttribute "getImage", "XT_SB_getImage"
          Case XT_ImageTypeMso
            .setAttribute "imageMso", module.Image
          Case XT_ImageTypeEmbedded
            .setAttribute "image", module.Image
        End Select
        Dim info As MacroInfo
        For Each info In module.Macros
          If info.isTarget(Selection) Then
            submenu.appendChild CNode(ButtonFromInfo(xmlDoc, info, control_tag))
          End If
        Next
      End With
      If submenu.HasChildNodes() Then
        ' サブメニューがある場合のみ登録する
        submenus.Add module.Name, submenu
        keys.Add module.Name
      End If
    End If
  Next

  Dim key As Variant
  For Each key In keys
    menu.appendChild CNode(submenus(key))
  Next
  Set info = Nothing

  If menu.ChildNodes.Length = 1 Then
    ' サブメニューが一つしかない場合は展開してしまう
    Dim btn        As IXMLDOMElement
    For Each submenu In menu.ChildNodes
      For Each btn In submenu.ChildNodes
        menu.appendChild CNode(btn)
      Next
      '@Ignore ArgumentWithIncompatibleObjectType
      menu.RemoveChild CNode(submenu)
    Next
  End If

  If control_tag = "0" And Not ThisWorkbook.IsAddin Then
    ' アドインでなければ，メインのものにだけメニュー再構築を表示する．
    Set submenu = xmlDoc.createElement("menuSeparator")
    submenu.setAttribute "id", ThisWorkbook.Name & "_" & "sepa"
    menu.appendChild CNode(submenu)
    menu.appendChild CNode(Button(xmlDoc, "reset", "メニュー再構築", "RefreshMenu", "メニューをリセット", _
                                  "新たにマクロが追加された場合などにメニューに変更を反映させます"))
  End If

  CreateXML = Replace(xmlDoc.XML, "><", ">" & vbCrLf & "<")    'xmlDoc.xml

  If Not ThisWorkbook.IsAddin Then
    'アドインでなければ，デバッグの参考用にxmlを書き出す．
    Dim FSO        As New FileSystemObject
    With FSO.OpenTextFile(ThisWorkbook.path & "\" & last_xml_file, ForWriting, True)
      .WriteLine "control_tag: " & control_tag
      .WriteBlankLines 2
      .Write Replace(xmlDoc.XML, "><", ">" & vbCrLf & "<")
      .Close
    End With
  End If
  Set submenus = Nothing
  Set keys = Nothing
End Function


' スプリットボタンに登録された，最近のマクロについて，
' 対応するマクロが存在するかどうかを確認し，無ければ取り除く．
Private Function 履歴のクリーンアップ() As Boolean
  Dim APP As String: APP = ThisWorkbook.Name
  Dim sec As String: sec = "UI"
  Dim ten As Long
  For ten = 1 To 3
    Dim arr(0 To 9) As String    'マクロ名を保持する配列
    Dim n As Long: n = 0 ' 登録されているマクロ数
    Dim w As Long: w = 0 ' 登録されており，存在するマクロ数
    Dim r As Long ' インデクサ
    For r = 0 To 9    '登録され，実在するマクロ名のリストを取得
      Dim s        As String
      s = GetSetting(APP, sec, CStr(ten) & CStr(r), "")
      If s <> "" Then
        n = n + 1
        If macroTree.hasMacro(s) Then
          arr(w) = s
          w = w + 1
        End If
      End If
    Next

    If w < n Then    ' 存在しないマクロがあった場合は詰める．
      履歴のクリーンアップ = True
      Do While w < 10    ' 10個になるまで未登録を詰める
        arr(w) = ""
        w = w + 1
      Loop
      Dim i As Long
      For i = 0 To 9    'レジストリに保存する
        SaveSetting APP, sec, CStr(ten) & CStr(w), arr(w)
      Next
    End If
  Next
End Function


' interface
Public Sub RefreshMenu(control As IRibbonControl)
  Call InvaditateMenu
End Sub


' リボンを無効化して，メニューを再構築する．
Private Sub InvaditateMenu()
  If Ribbon.Item Is Nothing Then
    MsgBox "更新ができません．ソフトウェアを再起動してください．"
  Else
    Set macroTree = Nothing
    Ribbon.Invalidate
  End If
End Sub


' マクロをパースし，XML作成の準備をする．
' 合わせて，macrosコレクションも更新．
Private Sub initTree()
  Set macroTree = New MacroProject    ' 新しいものを割り当て
  macroTree.Parse ThisWorkbook.VBProject
End Sub


'@Description レジストリからショートカットの情報を読み込む
Private Sub loadShortcutKeys()
  Dim data As Variant
  data = GetAllSettings(ThisWorkbook.Name, SHORTCUT_SEC)

  If Not IsEmpty(data) Then
    Dim i As Long
    For i = LBound(data, 1) To UBound(data, 1)
      Dim key As String
      key = data(i, 0)
      Dim target As String
      target = data(i, 1)
      Dim info As MacroInfo
      Set info = getInfo(target)
      If info.Name = "" Then
        Debug.Print "マクロが見つかりませんでした．" & key & " => " & target
      Else
        info.onKey = key
        Application.onKey key, target
      End If
    Next
  End If
End Sub


'supertip: レジストリに登録されているショートカットの情報をカレントセル以降に書き出す．
Sub ショートカットの情報をダンプ()
  Dim data As Variant
  data = GetAllSettings(ThisWorkbook.Name, SHORTCUT_SEC)

  If IsEmpty(data) Then
    MsgBox "ショートカットは登録されていません．", vbOKOnly, "確認"
    ActiveCell.Value = "ショートカットは登録されていません"
  Else
    Dim rows As Long
    rows = UBound(data, 1) - LBound(data, 1) + 1

    Dim Dest As Range
    Set Dest = ActiveCell.Resize(rows + 3, 2)
    If WorksheetFunction.CountA(ActiveCell.Resize(rows + 3, 3)) > 0 Then
      If MsgBox("情報が挿入される範囲内(" & rows & "x3)にデータがあります．上書きして良いですか？", vbYesNo + vbQuestion, "上書き確認") = vbNo Then Exit Sub
      Dest.Clear
    End If
    ショートカットダンプのヘッダを書き出し
    ActiveCell.Offset(2, 0).Resize(rows, 2) = data
    ActiveCell.Offset(2, 2).FormulaR1C1 = "=IFError(VLookup(RC[-2],'[" & ThisWorkbook.Name & "]KeyList'!C1:C3, 3,False),""(なし)"")"
    ActiveCell.Offset(2, 2).Resize(rows, 1).FillDown
    ActiveCell.Offset(rows + 3, 1).Value = "(End)"
    ショートカット編集後作業のコメントを追加
  End If
End Sub


' label: ショートカットの情報をレジストリに保存
' supertip: カレントセルを基準に，マクロを割り当てをレジストリに保存する．
' supertip: カレントセルにショートカットキーの定義を，右隣のセルにマクロ名を入れる．
' supertip: 無効なデータの場合は，セルの背景が赤になる．
' supertip: 同様にして次の行以降もキー定義のセルが空になるまで順次繰り返す．
' supertip: 既存の情報は削除せず，追加となる．
' supertip: 割り当てを削除したい場合は，マクロ名を空欄にして実行する．
Sub ショートカットの情報を保存()
  Dim r As Range
  Dim key As String
  Dim macro As String
  Dim APP As String
  APP = ThisWorkbook.Name

  Set r = ActiveCell
  Do While r.Value <> ""
    key = r.Value
    If Left$(key, 1) = "'" Then key = Mid$(key, 2)
    macro = r.Offset(0, 1).Value
    If macro = "" Then
      On Error Resume Next
      DeleteSetting APP, SHORTCUT_SEC, key
      On Error GoTo 0
      r.Interior.color = RGB(0, 255, 0)
    Else
      If Left$(macro, 1) = "'" Then macro = Mid$(macro, 2)
      If isKeyValid(key) And macroTree.hasMacro(UCase$(macro)) Then
        SaveSetting APP, SHORTCUT_SEC, normalizedKey(key), macro
        r.Interior.ColorIndex = 24
      Else
        If Not isKeyValid(key) Then r.Interior.color = RGB(255, 0, 0)
        If Not macroTree.hasMacro(UCase$(macro)) Then r.Offset(0, 1).Interior.color = RGB(255, 0, 0)
      End If
    End If
    Set r = r.Offset(1, 0)
  Loop
End Sub


Private Function isKeyValid(ByVal key As String) As Boolean
  isKeyValid = False
  If Len(key) < 2 Then Exit Function
  If InStr("^%+", Left$(key, 1)) = 0 Then Exit Function
  Dim start As Long
  For start = 1 To Len(key)
    If InStr("^%+", Mid$(key, start, 1)) = 0 Then Exit For
  Next
  If start > Len(key) Then Exit Function
  Dim body As String
  body = Mid$(key, start)
  If Len(body) = 2 Then Exit Function
  If Len(body) > 2 Then
    isKeyValid = Left$(body, 1) = "{" And Right$(body, 1) = "}"
  ElseIf Len(body) = 1 Then
    isKeyValid = InStr("0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ.,:><*;\_-=~`?!""#$&'()|", body) > 0
  End If
End Function


Private Function normalizedKey(ByVal key As String) As String
  Dim ctrl As Boolean:   ctrl = False
  Dim alt As Boolean:     alt = False
  Dim shift As Boolean: shift = False

  Dim i As Long
  For i = 1 To Len(key)
    Select Case Mid$(key, i, 1)
      Case "^"
        ctrl = True
      Case "%"
        alt = True
      Case "+"
        shift = True
      Case Else
        Exit For
    End Select
  Next

  normalizedKey = ""
  If ctrl Then normalizedKey = "^"
  If alt Then normalizedKey = normalizedKey & "%"
  If shift Then normalizedKey = normalizedKey & "+"
  normalizedKey = normalizedKey & Mid$(key, i)
End Function


' ショートカットのキーをわかりやすく展開する
Private Function FormatShortcutKey(key As String) As String
  Dim rest As String
  rest = key
  If Left$(rest, 1) = "^" Then
    rest = Mid(key, 2)
    FormatShortcutKey = "Ctrl+"
  End If
  If Left$(rest, 1) = "%" Then
    FormatShortcutKey = FormatShortcutKey & "Alt+"
    rest = Mid(rest, 2)
  End If
  If Left$(rest, 1) = "+" Then
    FormatShortcutKey = FormatShortcutKey & "Shift+"
    rest = Mid(rest, 2)
  End If
  FormatShortcutKey = FormatShortcutKey & rest
End Function



' supertip: レジストリに保存されている情報をもとに，ショートカットキーをマクロに割り当てる．
Sub ショートカットの割り当て()
  Dim data As Variant
  data = GetAllSettings(ThisWorkbook.Name, SHORTCUT_SEC)
  If Not IsEmpty(data) Then
    Dim i As Long
    For i = LBound(data, 1) To UBound(data, 1)
      Application.onKey data(i, 0), data(i, 1)
    Next
  End If
End Sub


' supertip: レジストリから登録したショートカットキーを無効にし，エクセルのデフォルトの挙動に戻す．
Sub ショートカットの解除()
  Dim data As Variant
  data = GetAllSettings(ThisWorkbook.Name, SHORTCUT_SEC)
  If Not IsEmpty(data) Then
    Dim i As Long
    For i = LBound(data, 1) To UBound(data, 1)
      Application.onKey data(i, 0)    '第2引数を省略して，デフォルトに戻す．
    Next
  End If
End Sub


Private Sub ショートカットダンプのヘッダを書き出し()
  ActiveCell.Value = "キー"
  ActiveCell.AddComment "Ctrl: ^" & vbCrLf & _
                        "Alt: %" & vbCrLf & _
                        "Shift: +" & vbCrLf & _
                        "上記記号と特殊キーは{}で囲う．" & vbCrLf & _
                        "例：" & vbCrLf & _
                        "    ^{+}:  Ctrl++" & vbCrLf & _
                        "   {F10}: ファンクションキー F10" & vbCrLf & _
                        " %+{F12}: Alt+Shift+F12"
  ActiveCell.Offset(0, 1).Value = "マクロ名"
  ActiveCell.Offset(0, 2).Value = "上書きされるエクセル標準機能"
End Sub


Private Sub ショートカット編集後作業のコメントを追加()
  ActiveCell.Offset(2, 0).AddComment "編集後はここを選択してから，" & vbCrLf & _
                                     "  UI->ショートカットの情報を保存" & vbCrLf & _
                                     "  UI->ショートカットの登録" & vbCrLf & _
                                     "を順に実行してください．"
End Sub


Private Function ObtainSheet(ByVal sheet_name As String) As Worksheet
  On Error GoTo NoSheet:
  Set ObtainSheet = ActiveWorkbook.Worksheets(sheet_name)
  Exit Function
NoSheet:
  Set ObtainSheet = ActiveWorkbook.Worksheets.Add()
  ObtainSheet.Name = sheet_name
End Function


' supertip: メニューから呼び出すことのできる全マクロを一覧にして新しいシートに出力する．
'
Sub マクロ一覧をダンプ()
  If macroTree Is Nothing Then initTree

  Dim r As Range
  If ActiveWorkbook.Name = ThisWorkbook.Name Then
    Set r = ActiveCell
  Else
    With ObtainSheet("Macro一覧")
      .Activate
      Set r = .Range("A1")
    End With
  End If

  Dim num_macro As Long
  num_macro = macroTree.MacroCount

  Dim caption As Variant
  caption = Array("#", "name", "id", "project", "module", "tag", "label", "category", "screenTip", "superTip", "desc", "image", "imageType")

  Dim num_item As Long
  num_item = UBound(caption)

  Dim data() As Variant
  ReDim data(0 To num_macro, 0 To num_item)

  Dim col As Long
  For col = 0 To num_item
    data(0, col) = caption(col)
  Next

  Dim row As Long
  row = 0

  Dim key As Variant
  For Each key In macroTree.MacroNames
    Dim info As MacroInfo
    Set info = macroTree.macro(key)
    row = row + 1
    For col = 0 To num_item
      Select Case caption(col)
        Case "#"
          data(row, col) = row
        Case "name"
          data(row, col) = info.Name
        Case "id"
          data(row, col) = info.id
        Case "project"
          data(row, col) = info.ProjectName
        Case "module"
          data(row, col) = info.module
        Case "tag"
          data(row, col) = info.tag
        Case "label"
          data(row, col) = info.label
        Case "category"
          data(row, col) = info.category
        Case "screenTip"
          data(row, col) = info.screenTip
        Case "superTip"
          data(row, col) = info.superTip
        Case "desc"
          data(row, col) = info.desc
        Case "image"
          data(row, col) = info.Image
        Case "imageType"
          data(row, col) = info.imageTypeString
        Case Else
          Debug.Assert False
      End Select
    Next
  Next

  r.Resize(num_macro + 1, num_item) = data
End Sub



