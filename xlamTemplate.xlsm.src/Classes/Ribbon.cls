'@ModuleDescription "OfficeRibonのポインタを保存する機能．"
'@PredeclaredId
'@Folder RibbonUI
Option Explicit

#If VBA7 And Win64 Then
  Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDest As Any, ByRef pSrc As Any, ByVal cbLen As LongPtr)
#Else
  Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDest As Any, ByRef pSrc As Any, ByVal cbLen As Long)
#End If

Private gRibbon    As Office.IRibbonUI
Const sec          As String = "UI"
Const key          As String = "RibbonPointer"
'Note: リボンのポインタ保存の参照先   https://www.ka-net.org/ribbon/ri64.html


' リボンへの参照を返す．
' マクロの修正で，参照がリセットされた場合への対応あり．
'@DefaultMember
'@Description "リボンへの参照(IRibbonUI)を返す．無効になっていた場合は保存していたポインタから返す．"
Property Get Item() As Office.IRibbonUI
  If gRibbon Is Nothing Then
    Dim Ptr      As String
    ' レジストリからポインタのアドレスを取得する．
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
      MoveMemory ribbonObj, pZero, LenB(pZero)    '後始末 (ゼロクリア）
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
  SaveSetting APP, sec, key, CStr(ObjPtr(pointer))    'リボンのポインタをレジストリに記録
End Property

