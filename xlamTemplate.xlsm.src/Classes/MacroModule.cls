Option Explicit
'@Folder RibbonUI


'@Ignore EncapsulatePublicField
'@VariableDescription モジュール名．メニューの名前としても用いられる．
Public Name        As String


'@VariableDescription マクロの情報を保持するMacroInfoのコレクション
Private m_Macros      As Collection


'@Ignore EncapsulatePublicField
'@VariableDescription サブメニューを表示する選択項目の設定．
Public target        As String


Private m_Image As String


Private m_ImageType As XT_ImageType


Private m_ProjectName As String


Private Sub Class_Initialize()
  Set m_Macros = New Collection
  m_Image = "MacrosGallery"
  m_ImageType = XT_IMageTypeNone
End Sub


Private Sub Class_Terminate()
  Call Clear
  Set m_Macros = Nothing
End Sub


'@Description 登録したマクロのコレクションを空にする．名前は保持される．
Public Sub Clear()
  Do While Macros.Count > 0
    Macros.Remove 1
  Loop
End Sub


'@Description MacroInfoを登録する．
Public Sub Add(ByVal macro As MacroInfo)
  Macros.Add macro, macro.Name
End Sub


'@Description "登録されているマクロの数"
Property Get Count() As Long
  Count = Macros.Count
End Property


'@Description "登録されているMacroInfoを返す読み取り専用プロパティ"
'@DefaultMember
Property Get Item(ByVal index As Variant) As MacroInfo
  Set Item = Macros(index)
End Property


'@Description "モジュールのイメージ情報を保存"
Public Property Get ImageType() As XT_ImageType
  ImageType = m_ImageType
End Property

Public Property Let ImageType(ByVal Value As XT_ImageType)
  m_ImageType = Value
End Property


'@Description("アイコンイメージの名前等を保持（Get, Let)")
Public Property Get Image() As String
  Image = m_Image
End Property


Public Property Let Image(ByVal Value As String)
  m_Image = Value
End Property


'@Description("マクロの情報を保持するMacroInfoのコレクション(Get)")
Public Property Get Macros() As Collection
  Set Macros = m_Macros
End Property



'@Enumerator
Public Function NewEnum() As IEnumVARIANT
  Set NewEnum = Macros.[_NewEnum]
End Function


'@Description("プロジェクトのファイル名(Get, Let)")
Public Property Get ProjectName() As String
  ProjectName = m_ProjectName
End Property


Public Property Let ProjectName(ByVal Value As String)
  m_ProjectName = Value
End Property


'@Description("与えたオブジェクトが対象に合致するか？")
Public Function isTarget(ByVal Value As Object) As Boolean
  isTarget = True
  Dim obj As Object
  On Error GoTo eee:
  If Me.target = "" Then Exit Function
  If target = "Application" Then Exit Function
  If TypeName(Value) Like target Then Exit Function

  If TypeName(Value) = "DrawingObjects" Then
    For Each obj In Value
      If Me.isTarget(obj) Then Exit Function
    Next
  Else
    Set obj = Value
    '@Ignore VariableNotUsed
    Dim i As Long
    For i = 1 To 20
      If TypeName(obj) = "Application" Then Exit For
      Set obj = obj.Parent
      If TypeName(obj) Like Me.target Then Exit Function
    Next
  End If
eee:
  isTarget = False
End Function

' マクロ定義の確認と取得． subなどは行頭からでないとヒットしない様にしてある．
Private Function MacroDef(ByVal line As String) As String
  If line Like "Sub *()*" Or line Like "Sub *(Optional *)*" Then
    MacroDef = Mid$(line, 5, InStr(line, "(") - 5)
  ElseIf line Like "Public Sub *()*" Or line Like "Pubilc Sub *(Optional *)*" Then
    MacroDef = Mid$(line, 12, InStr(line, "(") - 12)
  End If
End Function

Private Function IsEndOfDef(ByVal line As String) As Boolean
  IsEndOfDef = line Like "End Sub" Or line Like "End Function" Or line Like "End Type"
End Function


'@Description("引数のVBコンポーネントをパースして，情報を取り込む．")
Public Sub Parse(ByVal comp As VBComponent)
  Me.Name = comp.Name
  If comp.Type <> 1 And comp.Type <> 100 Then Exit Sub

  Dim 登録       As Boolean: If comp.Type = 1 Then 登録 = True Else 登録 = False
  Dim isModuleDefitinion As Boolean:  isModuleDefitinion = True

  Dim info     As MacroInfo:   Set info = New MacroInfo

  Dim iRow        As Long
  For iRow = 1 To comp.CodeModule.CountOfLines
    Dim line   As String:    line = comp.CodeModule.Lines(iRow, 1)

    Dim collon_pos As Long:    collon_pos = InStr(line, ":")
    Dim body   As String:    body = ""
    Dim notEmpty As Boolean
    If collon_pos > 0 Then body = Mid$(line, collon_pos + 1)
    notEmpty = Trim(body) <> ""

    'サブルーチン等の終わりにきたのでリセット （subやFunction内のキーワードを無効化）
    If IsEndOfDef(line) Then
      info.Clear
      If comp.Type = 1 Then 登録 = True Else 登録 = False
    End If

    ' Moduleの追加オプションのチェック
    If line Like "*Sub *" Or line Like "*Function *" Then isModuleDefitinion = False
    If notEmpty And isModuleDefitinion Then
      Select Case True
      Case line Like "'*moduleImage:*" And isModuleDefitinion
        Me.ImageType = XT_ImageTypeEmbedded
        Me.Image = Trim(body)
      Case line Like "'*moduleImageMso:*" And isModuleDefitinion
        Me.ImageType = XT_ImageTypeMso
        Me.Image = Trim(body)
      Case line Like "'*moduleInternalImage:*" And isModuleDefitinion
        Me.ImageType = XT_ImageTypeInternal
        Me.Image = Trim(body)
      Case line Like "'*moduleExternalImage:*" And isModuleDefitinion
        Me.ImageType = XT_ImageTypeExternal
        Me.Image = Trim(body)
      Case line Like "'*moduleTarget:*"
        Me.target = Trim(body)
      End Select
    End If

    '以降はMacroInfoの追加オプションの定義． (2020/5/8 仕様変更 see マクロ自動メニューの仕様.txt)
    If notEmpty Then
      Select Case True
      Case line Like "'*screentip:*"
        info.screenTip = Trim(body)
      Case line Like "'*supertip:*"
        info.superTip = info.superTip & RTrim$(body)
      Case line Like "'*desc:*"
        info.desc = info.desc & Trim(body)
      Case line Like "'*target:*"
        info.target = Trim(body)
      Case line Like "'*imageMso:*"
        info.Image = Trim(body)
        info.ImageType = XT_ImageTypeMso
      Case line Like "'*image:*"
        info.Image = Trim(body)
        info.ImageType = XT_ImageTypeEmbedded
      Case line Like "'*internalImage:*"
        info.Image = Trim(body)
        info.ImageType = XT_ImageTypeInternal
      Case line Like "'*externalImage:*"
        info.Image = Trim(body)
        info.ImageType = XT_ImageTypeExternal
      Case line Like "'*label:*"
        info.label = Trim(body)
      Case line Like "'*hidden:"
        Select Case UCase$(Trim(body))
        Case "ISADDIN"
          登録 = Not ThisWorkbook.IsAddin
        Case "FALSE"
          登録 = True
        Case Else
          登録 = False
        End Select
      Case line Like "'*export:"
        If UCase$(Trim(body)) <> "FALSE" Then 登録 = True
      End Select
    Else
      If line Like "'*hidden:" Then 登録 = False
      If line Like "'*export:" Then 登録 = True
    End If

    info.Name = MacroDef(line)
    If info.Name <> "" Then
      ' マクロの定義行なので，可能なら登録する．
      If 登録 Then
        info.module = comp.Name
        info.tag = info.Name
        If info.screenTip = "" Then info.screenTip = info.Name
        ' プロジェクトの情報を保存
        info.ProjectName = Me.ProjectName
        Me.Add info
      End If
      ' reset
      Set info = New MacroInfo
    End If
  Next
End Sub