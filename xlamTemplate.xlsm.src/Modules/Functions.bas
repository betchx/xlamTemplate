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


'@Description("ActiveSheet�ւ̎Q�Ƃ�Ԃ��D�C���e���Z���X���L���D")
Function ActSht() As Worksheet
  Set ActSht = ActiveSheet
End Function


'@Description("���W�X�g����APP���D�}�N���̂���u�b�N�̖��O��p����D")
Function APP() As String
  APP = ThisWorkbook.Name
End Function


'@Description("�V�[�g�����܂߂��A�h���X��A1�`���ŕԂ��D")
Function FullAddress(ByVal rng As Range) As String
  FullAddress = getSheetName(rng) & "!" & rng.Address(True, True, xlA1)
End Function


'@Description("�V�[�g�����܂߂��A�h���X��R1C1�`���ŕԂ��D")
Function FullAddressR1C1(ByVal rng As Range) As String
  FullAddressR1C1 = getSheetName(rng) & "!" & rng.Address(True, True, xlR1C1)
End Function


'@Description("������Range�̐e�ƂȂ�V�[�g�̖��O��Ԃ�")
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


'@Description("A1�`���̃A�h���X��R1C1�`���ɕϊ�����DConvertFormula�̃��b�p�[�D")
Function A1toR1C1(ByVal a1 As String) As String
  A1toR1C1 = Application.ConvertFormula(a1, xlA1, xlR1C1)
End Function


'@Description("A1�`���̃A�h���X��R1C1�`���ɕϊ�����DConvertFormula�̃��b�p�[�D")
Function R1C1toA1(ByVal r1c1 As String) As String
  R1C1toA1 = Application.ConvertFormula(r1c1, xlR1C1, xlA1)
End Function


'@Description("�|�C���g��mm�ɕϊ�����D")
Function pnt2mm(ByVal p As Single) As Single
  pnt2mm = CSng(p / Application.CentimetersToPoints(0.1))
End Function


'@Description("�|�C���g��cm�ɕϊ�����D")
Function pnt2cm(ByVal pnt As Single) As Single
  pnt2cm = CSng(pnt / Application.CentimetersToPoints(1#))
End Function


'@Description("mm���|�C���g�ɕϊ�����D")
Function mm2pnt(ByVal mm As Single) As Single
  mm2pnt = CSng(Application.CentimetersToPoints(mm * 0.1))
End Function


'@Description("cm���|�C���g�ɕϊ�����DApplication.CentimetersToPoints�̂قڕʖ�")
Function cm2pnt(ByVal cm As Single) As Single
  cm2pnt = CSng(Application.CentimetersToPoints(cm))
End Function


'@Description("Range�̎Q�Ɛ�̒l�𐮐��iLong�j�ŕԂ�")
Function r2i(ByVal r As Range) As Long
  r2i = CInt(val(r.Value))
End Function


'@Description("Range�̎Q�Ɛ�̒l��P���x���������_���iSingle�j�ŕԂ�")
Function r2s(ByVal r As Range) As Single
  r2s = CSng(val(r.Value))
End Function


'@Description("�����̖��O�̎Q�Ɛ�̒l��P���x���������_���iSingle�j�ŕԂ��Di��n2r�ɓn�����D")
Function n2s(ByVal n As String, Optional ByVal i As Long = 0) As Single
  n2s = r2s(n2r(n, i))
End Function


'@Description("���O���烌���W���擾����Di������ꍇ�͔͈͂�i�Ԗڂ̃Z����Ԃ�")
' �܂��̓A�N�e�B�u�V�[�g�ŒT���C������Ȃ��ꍇ�̓��[�N�u�b�N����T�����ƂɂȂ�D
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
  ' Note: �����̓��[�v�ŒT���Ă������C�x�����Ƃ����������̂ŁC
  '       �G���[�㓙�Œ��ڎQ�Ƃ���`�ɕύX�����D
End Function


'@Description("���O�I�u�W�F�N�g��T���ĕԂ�")
Function s2n(ByVal str As String) As Name
  On Error Resume Next
  Set s2n = ActSht.Names(str)
  If s2n Is Nothing Then
    Set s2n = ActiveWorkbook.Names(str)
  End If
  On Error GoTo 0
End Function


'@Description("�͈͂�\�������񂩂�͈͂�Ԃ��D�V�[�g���������Ă����삷��D�ʃu�b�N�͖����D")
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


'@Description("�֐���\���������1�i�K�����������āC������̔z��ɕϊ����ĕԂ��Dindex=0���֐����C�ȍ~�������D")
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

  ' ����
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

'@Description("deFunc�ŕ��������z���g�ݗ��ĂĊ֐��̕�����ɂ���D")
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


'@Description("������Office�I�u�W�F�N�g�Ƃ��̐e����ChartObject��T���ĕԂ��D������Ȃ�����Nothing.")
Function FindChartObject(ByVal start_from As Object) As ChartObject
  Dim obj As Object
  Set obj = start_from
  If obj Is Nothing Then Exit Function
  On Error GoTo eee:
  '@Ignore VariableNotUsed
  Dim i As Long
  For i = 1 To 20    '�ő�20�K�w�܂Łi�i�v���[�v�h�~�j
    If TypeOf obj Is Excel.Application Then Exit Function
    Set obj = obj.Parent
    If TypeOf obj Is ChartObject Then
      Set FindChartObject = obj
      Exit Function
    End If
  Next
eee:
End Function


'@Description("���p����FileSystemObject��Ԃ��D")
Function FSO() As FileSystemObject
  Static fso_ As FileSystemObject
  If fso_ Is Nothing Then Set fso_ = New FileSystemObject
  Set FSO = fso_
End Function


'@Description("With���ƂƂ��Ɏg�p���āC�s�v�ȏ�����}�����ă}�N���̎��s���x�𑝉�������D")
Function Boost(Optional ByVal �C�x���g�}�� As Boolean = False, _
               Optional ByVal �Čv�Z�}�� As Boolean = True, _
               Optional ByVal �v�����^�ʐM�}�� As Boolean = True, _
               Optional ByVal ��ʍX�V�}�� As Boolean = True) As MacroBooster
  Set Boost = New MacroBooster
  With Boost
    If �C�x���g�}�� Then .DisableEvents
    If �Čv�Z�}�� Then .NoCalc
    If �v�����^�ʐM�}�� Then .NoPrinterComunication
    If ��ʍX�V�}�� Then .NoScreenUpdating
    .Enter
  End With
End Function


'@Description("�����̃I�u�W�F�N�g�̐e�����ǂ��ă`���[�g��Ԃ��D�`���[�g���Ȃ�������Nothing")
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


'@Description("�����̃I�u�W�F�N�g�̐e�����ǂ��ă`���[�g�I�u�W�F�N�g��Ԃ��D�`���[�g�I�u�W�F�N�g���Ȃ�������Nothing")
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


'@Description("Chart���t���[�g�z�u�ɂ���Ƃ��̒萔���킩�肸�炢���߂ɍ쐬��������")
Public Function �t���[�g�z�u�ݒ�() As XlPlacement
  �t���[�g�z�u�ݒ� = xlMove
End Function


'@Description �Z���̃t�H�[�}�b�g��������擾����֐��D�����ȗ����ɂ͑I��͈͂�ΏۂƂ���D
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

' �g�����̗�
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

' @Description("�؂����R�s�[�̃��[�h�̎��ɁC�j���ň͂��Ă���͈�(Range)���擾����D")
Function CutCopyOrigin() As Range
' �Q�ƁF  http://www2.aqua-r.tepm.jp/~hironobu/ke_m13.htm   E03M121(Excel2003)
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