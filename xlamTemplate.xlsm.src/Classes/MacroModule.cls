Option Explicit
'@Folder RibbonUI


'@Ignore EncapsulatePublicField
'@VariableDescription ���W���[�����D���j���[�̖��O�Ƃ��Ă��p������D
Public Name        As String


'@VariableDescription �}�N���̏���ێ�����MacroInfo�̃R���N�V����
Private m_Macros      As Collection


'@Ignore EncapsulatePublicField
'@VariableDescription �T�u���j���[��\������I�����ڂ̐ݒ�D
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


'@Description �o�^�����}�N���̃R���N�V��������ɂ���D���O�͕ێ������D
Public Sub Clear()
  Do While Macros.Count > 0
    Macros.Remove 1
  Loop
End Sub


'@Description MacroInfo��o�^����D
Public Sub Add(ByVal macro As MacroInfo)
  Macros.Add macro, macro.Name
End Sub


'@Description "�o�^����Ă���}�N���̐�"
Property Get Count() As Long
  Count = Macros.Count
End Property


'@Description "�o�^����Ă���MacroInfo��Ԃ��ǂݎ���p�v���p�e�B"
'@DefaultMember
Property Get Item(ByVal index As Variant) As MacroInfo
  Set Item = Macros(index)
End Property


'@Description "���W���[���̃C���[�W����ۑ�"
Public Property Get ImageType() As XT_ImageType
  ImageType = m_ImageType
End Property

Public Property Let ImageType(ByVal Value As XT_ImageType)
  m_ImageType = Value
End Property


'@Description("�A�C�R���C���[�W�̖��O����ێ��iGet, Let)")
Public Property Get Image() As String
  Image = m_Image
End Property


Public Property Let Image(ByVal Value As String)
  m_Image = Value
End Property


'@Description("�}�N���̏���ێ�����MacroInfo�̃R���N�V����(Get)")
Public Property Get Macros() As Collection
  Set Macros = m_Macros
End Property


'@Enumerator
Public Function NewEnum() As IEnumVARIANT
  Set NewEnum = Macros.[_NewEnum]
End Function


'@Description("�v���W�F�N�g�̃t�@�C����(Get, Let)")
Public Property Get ProjectName() As String
  ProjectName = m_ProjectName
End Property


Public Property Let ProjectName(ByVal Value As String)
  m_ProjectName = Value
End Property


'@Description("�^�����I�u�W�F�N�g���Ώۂɍ��v���邩�H")
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

' �}�N����`�̊m�F�Ǝ擾�D sub�Ȃǂ͍s������łȂ��ƃq�b�g���Ȃ��l�ɂ��Ă���D
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


'@Description("������VB�R���|�[�l���g���p�[�X���āC������荞�ށD")
Public Sub Parse(ByVal comp As VBComponent)
  Me.Name = comp.Name
  If comp.Type <> 1 And comp.Type <> 100 Then Exit Sub

  Dim �o�^       As Boolean: If comp.Type = 1 Then �o�^ = True Else �o�^ = False
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

    '�T�u���[�`�����̏I���ɂ����̂Ń��Z�b�g �isub��Function���̃L�[���[�h�𖳌����j
    If IsEndOfDef(line) Then
      info.Clear
      If comp.Type = 1 Then �o�^ = True Else �o�^ = False
    End If

    ' Module�̒ǉ��I�v�V�����̃`�F�b�N
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

    '�ȍ~��MacroInfo�̒ǉ��I�v�V�����̒�`�D (2020/5/8 �d�l�ύX see �}�N���������j���[�̎d�l.txt)
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
          �o�^ = Not ThisWorkbook.IsAddin
        Case "FALSE"
          �o�^ = True
        Case Else
          �o�^ = False
        End Select
      Case line Like "'*export:"
        If UCase$(Trim(body)) <> "FALSE" Then �o�^ = True
      End Select
    Else
      If line Like "'*hidden:" Then �o�^ = False
      If line Like "'*export:" Then �o�^ = True
    End If

    info.Name = MacroDef(line)
    If info.Name <> "" Then
      ' �}�N���̒�`�s�Ȃ̂ŁC�\�Ȃ�o�^����D
      If �o�^ Then
        info.module = comp.Name
        info.tag = info.Name
        If info.screenTip = "" Then info.screenTip = info.Name
        ' �v���W�F�N�g�̏���ۑ�
        info.ProjectName = Me.ProjectName
        Me.Add info
      End If
      ' reset
      Set info = New MacroInfo
    End If
  Next
End Sub