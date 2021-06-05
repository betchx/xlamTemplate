'@ModuleDescription "OfficeRibon�̃|�C���^��ۑ�����@�\�D"
'@PredeclaredId
'@Folder RibbonUI
Option Explicit

Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDest As Any, ByRef pSrc As Any, ByVal cbLen As LongPtr)

Private gRibbon    As Office.IRibbonUI
Const sec          As String = "UI"
Const key          As String = "RibbonPointer"
'Note: ���{���̃|�C���^�ۑ��̎Q�Ɛ�   https://www.ka-net.org/ribbon/ri64.html


' ���{���ւ̎Q�Ƃ�Ԃ��D
' �}�N���̏C���ŁC�Q�Ƃ����Z�b�g���ꂽ�ꍇ�ւ̑Ή�����D
'@DefaultMember
'@Description "���{���ւ̎Q��(IRibbonUI)��Ԃ��D�����ɂȂ��Ă����ꍇ�͕ۑ����Ă����|�C���^����Ԃ��D"
Property Get Item() As Office.IRibbonUI
  If gRibbon Is Nothing Then
    Dim Ptr      As String
    ' ���W�X�g������|�C���^�̃A�h���X���擾����D
    Ptr = GetSetting(APP, sec, key)

    If Ptr <> "" Then

      #If VBA7 And Win64 Then
        Dim lRibbonPointer As LongPtr
        Dim pZero  As LongPtr: pZero = 0
        lRibbonPointer = CLngPtr(Ptr)
      #Else
        Dim lRibbonPointer As Long
        Dim pZero  As Long: pZero = 0
        lRibbonPointer = CLng(Ptr)
      #End If

      Dim ribbonObj As Object
      Set ribbonObj = Nothing
      '@Ignore ImplicitUnboundDefaultMemberAccess
      MoveMemory ribbonObj, lRibbonPointer, LenB(lRibbonPointer)
      Set gRibbon = ribbonObj
      '@Ignore ImplicitUnboundDefaultMemberAccess
      MoveMemory ribbonObj, pZero, LenB(pZero)    '��n�� (�[���N���A�j
    End If
  End If
  Set Item = gRibbon
End Property


Sub ActivateTab(ByVal ControlID As String)
  If Not Item Is Nothing Then gRibbon.ActivateTab ControlID
End Sub


Sub ActivateTabMso(ByVal ControlID As String)
  If Not Item Is Nothing Then gRibbon.ActivateTabMso ControlID
End Sub


Sub ActivateTabQ(ByVal ControlID As String, ByVal Namespace As String)
  If Not Item Is Nothing Then gRibbon.ActivateTabQ ControlID, Namespace
End Sub


Sub Invalidate()
  If Not Item Is Nothing Then gRibbon.Invalidate
End Sub


Sub InvalidateControl(ByVal ControlID As String)
  If Not Item Is Nothing Then gRibbon.InvalidateControl ControlID
End Sub


Sub InvalidateControlMso(ByVal ControlID As String)
  If Not Item Is Nothing Then gRibbon.InvalidateControlMso ControlID
End Sub


Public Sub Dispose()
  SaveSetting APP, sec, key, ""
  Set gRibbon = Nothing
End Sub


Property Set Item(ByVal pointer As IRibbonUI)
  Set gRibbon = pointer
  SaveSetting APP, sec, key, CStr(ObjPtr(pointer))    '���{���̃|�C���^�����W�X�g���ɋL�^
End Property
