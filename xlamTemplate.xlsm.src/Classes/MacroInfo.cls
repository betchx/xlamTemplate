Option Explicit
'@Folder RibbonUI
'@ModuleDescription マクロ情報を保持するクラス

' 内容的にはTypeでも良いが，コンテナ（ディクショナリ）に入れる為にクラスにしている．

Enum TM_ImageType
  TM_IMageTypeNone    '画像設定がない場合 （デフォルト）
  TM_ImageTypeMso    'MSOイメージのIDの場合
  TM_ImageTypeEmbedded    '内部に埋め込まれた画像（PNG等）のIDの場合
  TM_ImageTypeInternal    'シート内部OLEObjectの名前の場合
  TM_ImageTypeInternalAutomatic    'シート内部OLEObjectの名前から自動設定された場合
  TM_ImageTypeExternal    '外部ファイル名の場合
End Enum


'@Ignore EncapsulatePublicField
'@VariableDescription ボタンやメニューに登録されるときのID.
Public id          As String


'@Ignore EncapsulatePublicField
'@VariableDescription マクロが定義されているプロジェクト(オフィスファイル)の名前
Public ProjectName As String


'@Ignore EncapsulatePublicField
'@VariableDescription マクロが定義されているモジュールの名前
Public module      As String


'@Ignore EncapsulatePublicField
'@VariableDescription マウスをホバーした時に表示されるツールチップの上側に表示される太字の文字列．
Public screenTip   As String

'@Ignore EncapsulatePublicField
'@VariableDescription マウスをホバーした時に表示されるツールチップの下側の文字列．
Public superTip    As String

'@Ignore EncapsulatePublicField
'@VariableDescription 対象となるマクロ（サブルーチン）の名前
Public Name        As String

'@Ignore EncapsulatePublicField
'@VariableDescription タグ文字列．いろいろな情報が保持される．
Public tag         As String

'@Ignore EncapsulatePublicField
'@VariableDescription ラージボタンの場合にラベルの下に表示される文字列．
Public desc        As String

'@Ignore EncapsulatePublicField
'@VariableDescription 使用するアイコン画像の名前．
Public Image       As String

'@Ignore EncapsulatePublicField
'@VariableDescription アイコン画像のタイプ．
Public ImageType  As TM_ImageType

'@Ignore EncapsulatePublicField
'@VariableDescription ラベル文字列．省略された場合はnameが使用される．
Public label       As String


'@Ignore EncapsulatePublicField
'@VariableDescription メニューに表示する場合のカテゴリ．サブメニューとして使われる．
Public category    As String

'@Ignore EncapsulatePublicField
'@VariableDescription マクロを実行するショートカットの定義
Public onKey       As String

'@Ignore EncapsulatePublicField
'@VariableDescription メニュー項目を表示する対象
Public target       As String


'@Description 内容を空文字列 or Falseに初期化
Public Sub Clear()
  id = ""
  screenTip = ""
  superTip = ""
  Name = ""
  tag = ""
  Image = ""
  desc = ""
  ImageType = TM_IMageTypeNone
  label = ""
  category = ""
  onKey = ""
End Sub


Public Sub setShortcut(ByVal key As String, Optional ByVal control As Boolean = False, Optional ByVal alt As Boolean, Optional ByVal shift As Boolean)
  onKey = ""
  If key = "" Then Exit Sub    'クリアしたのと同じ
  If control Then onKey = "^"
  If alt Then onKey = onKey & "%"
  If shift Then onKey = onKey & "+"
  If onKey = "" Then
    ' 修飾キーがない場合は無効
    Exit Sub
  End If
  Select Case key
    Case "^", "%", "+", "~"
      onKey = onKey & "{" & key & "}"
    Case "~"
      onKey = onKey & key
    Case Else
      If Len(key) = 1 Then
        onKey = onKey & LCase$(key)    ' 小文字にしないと上手くいかない
      ElseIf Left$(key, 1) = "{" And Right$(key, 1) = "}" Then
        onKey = onKey & key
      Else
        ' 無効
        onKey = ""
      End If
  End Select
End Sub


Public Property Get shortcut() As String
  If onKey = "" Then Exit Property

  shortcut = "("
  Dim pos As Long
  For pos = 1 To Len(onKey)
    Select Case Mid$(onKey, pos, 1)
      Case "^"
        shortcut = shortcut & "Ctrl+"
      Case "%"
        shortcut = shortcut & "Alt+"
      Case "+"
        shortcut = shortcut & "Shift+"
      Case Else
        Exit For
    End Select
  Next

  Dim body As String
  If Right$(onKey, 1) = "}" Then
    body = Mid$(onKey, pos + 1, Len(onKey) - pos - 1)
  Else
    body = Mid$(onKey, pos)
  End If

  Select Case body
    Case "RIGHT"
      body = "→"
    Case "LEFT"
      body = "←"
    Case "UP"
      body = "↑"
    Case "DOWN"
      body = "↓"
    Case "~"
      body = "Enter"
    Case "ENTER"
      body = "テンキーEnter"
    Case "DEL", "DELETE"
      body = "Del"
    Case "END"
      body = "End"
    Case "ESCAPE", "ESC"
      body = "Esc"
    Case "HOME"
      body = "Home"
    Case "INSERT"
      body = "Ins"
    Case "PGDN"
      body = "PageDown"
    Case "PGUP"
      body = "PageUp"
    Case "RETURN"
      body = "Return"
    Case "TAB"
      body = "Tab"
  End Select

  shortcut = shortcut & body & ")"
End Property


Private Sub Class_Initialize()
  ImageType = TM_IMageTypeNone
End Sub


Public Property Get imageTypeString() As String
  Select Case ImageType
    Case TM_ImageTypeEmbedded
      imageTypeString = "埋め込み"
    Case TM_ImageTypeExternal
      imageTypeString = "外部ファイル"
    Case TM_ImageTypeInternal
      imageTypeString = "imageシート(手動設定)"
    Case TM_ImageTypeInternalAutomatic
      imageTypeString = "imageシート(自動設定)"
    Case TM_ImageTypeMso
      imageTypeString = "MSO"
    Case TM_IMageTypeNone
      imageTypeString = "未設定"
  End Select
End Property


Public Function isTarget(ByVal Value As Object) As Boolean
  Dim obj As Object
  '@Ignore VariableNotUsed
  Dim Count As Long
  isTarget = True
  On Error GoTo eee:

  If "" = Me.target Or "Application" = Me.target Or TypeName(Value) Like Me.target Then Exit Function

  If TypeName(Value) = "DrawingObjects" Then
    ' 複数選択されていた場合に対応
    For Each obj In Value
      If Me.isTarget(obj) Then Exit Function
    Next
  Else
    Set obj = Value
    For Count = 1 To 20 ' 永久ループ防止のため20階層まで．
      If TypeName(obj) = "Application" Then Exit For
      Set obj = obj.Parent
      If TypeName(obj) Like Me.target Then Exit Function
    Next
  End If
eee:
  isTarget = False
End Function