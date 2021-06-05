Option Explicit
'@Folder RibbonUI
'@ModuleDescription �}�N������ێ�����N���X

' ���e�I�ɂ�Type�ł��ǂ����C�R���e�i�i�f�B�N�V���i���j�ɓ����ׂɃN���X�ɂ��Ă���D

Enum XT_ImageType
  XT_IMageTypeNone    '�摜�ݒ肪�Ȃ��ꍇ �i�f�t�H���g�j
  XT_ImageTypeMso    'MSO�C���[�W��ID�̏ꍇ
  XT_ImageTypeEmbedded    '�����ɖ��ߍ��܂ꂽ�摜�iPNG���j��ID�̏ꍇ
  XT_ImageTypeInternal    '�V�[�g����OLEObject�̖��O�̏ꍇ
  XT_ImageTypeInternalAutomatic    '�V�[�g����OLEObject�̖��O���玩���ݒ肳�ꂽ�ꍇ
  XT_ImageTypeExternal    '�O���t�@�C�����̏ꍇ
End Enum

'@Ignore EncapsulatePublicField
'@VariableDescription �{�^���⃁�j���[�ɓo�^�����Ƃ���ID.
Public id          As String

'@Ignore EncapsulatePublicField
'@VariableDescription �}�N������`����Ă���v���W�F�N�g(�I�t�B�X�t�@�C��)�̖��O
Public ProjectName As String

'@Ignore EncapsulatePublicField
'@VariableDescription �}�N������`����Ă��郂�W���[���̖��O
Public module      As String

'@Ignore EncapsulatePublicField
'@VariableDescription �}�E�X���z�o�[�������ɕ\�������c�[���`�b�v�̏㑤�ɕ\������鑾���̕�����D
Public screenTip   As String

'@Ignore EncapsulatePublicField
'@VariableDescription �}�E�X���z�o�[�������ɕ\�������c�[���`�b�v�̉����̕�����D
Public superTip    As String

'@Ignore EncapsulatePublicField
'@VariableDescription �ΏۂƂȂ�}�N���i�T�u���[�`���j�̖��O
Public Name        As String

'@Ignore EncapsulatePublicField
'@VariableDescription �^�O������D���낢��ȏ�񂪕ێ������D
Public tag         As String

'@Ignore EncapsulatePublicField
'@VariableDescription ���[�W�{�^���̏ꍇ�Ƀ��x���̉��ɕ\������镶����D
Public desc        As String

'@Ignore EncapsulatePublicField
'@VariableDescription �g�p����A�C�R���摜�̖��O�D
Public Image       As String

'@Ignore EncapsulatePublicField
'@VariableDescription �A�C�R���摜�̃^�C�v�D
Public ImageType  As XT_ImageType

'@Ignore EncapsulatePublicField
'@VariableDescription ���x��������D�ȗ����ꂽ�ꍇ��name���g�p�����D
Public label       As String

'@Ignore EncapsulatePublicField
'@VariableDescription ���j���[�ɕ\������ꍇ�̃J�e�S���D�T�u���j���[�Ƃ��Ďg����D
Public category    As String

'@Ignore EncapsulatePublicField
'@VariableDescription �}�N�������s����V���[�g�J�b�g�̒�`
Public onKey       As String

'@Ignore EncapsulatePublicField
'@VariableDescription ���j���[���ڂ�\������Ώ�
Public target       As String


'@Description ���e���󕶎��� or False�ɏ�����
Public Sub Clear()
  id = ""
  screenTip = ""
  superTip = ""
  Name = ""
  tag = ""
  Image = ""
  desc = ""
  ImageType = XT_IMageTypeNone
  label = ""
  category = ""
  onKey = ""
End Sub


Public Sub setShortcut(ByVal key As String, Optional ByVal control As Boolean = False, Optional ByVal alt As Boolean, Optional ByVal shift As Boolean)
  onKey = ""
  If key = "" Then Exit Sub    '�N���A�����̂Ɠ���
  If control Then onKey = "^"
  If alt Then onKey = onKey & "%"
  If shift Then onKey = onKey & "+"
  If onKey = "" Then
    ' �C���L�[���Ȃ��ꍇ�͖���
    Exit Sub
  End If
  Select Case key
    Case "^", "%", "+", "~"
      onKey = onKey & "{" & key & "}"
    Case "~"
      onKey = onKey & key
    Case Else
      If Len(key) = 1 Then
        onKey = onKey & LCase$(key)    ' �������ɂ��Ȃ��Ə�肭�����Ȃ�
      ElseIf Left$(key, 1) = "{" And Right$(key, 1) = "}" Then
        onKey = onKey & key
      Else
        ' ����
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
      body = "��"
    Case "LEFT"
      body = "��"
    Case "UP"
      body = "��"
    Case "DOWN"
      body = "��"
    Case "~"
      body = "Enter"
    Case "ENTER"
      body = "�e���L�[Enter"
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
  ImageType = XT_IMageTypeNone
End Sub


Public Property Get imageTypeString() As String
  Select Case ImageType
    Case XT_ImageTypeEmbedded
      imageTypeString = "���ߍ���"
    Case XT_ImageTypeExternal
      imageTypeString = "�O���t�@�C��"
    Case XT_ImageTypeInternal
      imageTypeString = "image�V�[�g(�蓮�ݒ�)"
    Case XT_ImageTypeInternalAutomatic
      imageTypeString = "image�V�[�g(�����ݒ�)"
    Case XT_ImageTypeMso
      imageTypeString = "MSO"
    Case XT_IMageTypeNone
      imageTypeString = "���ݒ�"
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
    ' �����I������Ă����ꍇ�ɑΉ�
    For Each obj In Value
      If Me.isTarget(obj) Then Exit Function
    Next
  Else
    Set obj = Value
    For Count = 1 To 20 ' �i�v���[�v�h�~�̂���20�K�w�܂ŁD
      If TypeName(obj) = "Application" Then Exit For
      Set obj = obj.Parent
      If TypeName(obj) Like Me.target Then Exit Function
    Next
  End If
eee:
  isTarget = False
End Function