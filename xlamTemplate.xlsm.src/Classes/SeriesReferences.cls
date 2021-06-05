'@Folder("Personal")
Option Explicit

'@ModuleDescription("�n��̕��������������N���X�D�Q�Ɛ�̗��ύX����Ȃǂ��\")

Private comps As Variant


'@Description("�n���Formula�ɑ���ł��鐔���iA1�`���j�FGet, Let")
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


'@Description("�n���Formula�ɑ���ł��鐔��(R1C1�`��)�FGet")
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


'@Description("X���i�J�e�S���[���j�̎Q�Ɛ��\��������FGet, Let")
Public Property Get XAddress() As String
  XAddress = comps(2)
End Property


Public Property Let XAddress(ByVal val As String)
  comps(2) = val
End Property


'@Description("Y���i�l���j�̎Q�Ɛ��\��������FGet, Let")
Public Property Get YAddress() As String
  YAddress = comps(3)
End Property


Public Property Let YAddress(ByVal val As String)
  comps(3) = val
End Property


'@Description("�n��̃^�C�g����\��������FGet, Let")
Public Property Get Title() As String
  Title = comps(1)
End Property


Public Property Let Title(ByVal val As String)
  comps(1) = val
End Property


'@Description("X���i�J�e�S���[���j�̎Q�Ɛ������Range�I�u�W�F�N�g�FGet, Set")
Public Property Get x() As Range
  Set x = Range(XAddress)
End Property


Public Property Set x(ByVal rng As Range)
  Me.XAddress = FullAddress(rng)
End Property


'@Description("Y���i�l���j�̎Q�Ɛ������Range�I�u�W�F�N�g�FGet, Set")
Public Property Get Y() As Range
  Set Y = Range(YAddress)
End Property


Public Property Set Y(ByVal rng As Range)
  YAddress = FullAddress(rng)
End Property


'@Description("�Z���Q�ƕ�������V�[�g���t����R1C1�`���ɕϊ�����D")
Private Function toR1C1(ByVal adr As String) As String
  toR1C1 = FullAddressR1C1(Range(adr))
End Function


'@Description("�n��̒l�ƃ^�C�g���̎Q�Ɛ�̗��Count�����E�ɂ��炷�D")
Public Sub OffsetColumn(Optional ByVal Count As Long = 1)
  YAddress = FullAddress(Y.Offset(0, Count))
  If Left$(Me.Title, 1) <> """" Then
    Me.Title = FullAddress(Range(Me.Title).Offset(0, Count))
  End If
End Sub
