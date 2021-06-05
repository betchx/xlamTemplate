'@IgnoreModule ModuleWithoutFolder
Option Explicit


Private Sub Workbook_AddinUninstall()
  ショートカットの解除
End Sub


Private Sub Workbook_AfterSave(ByVal Success As Boolean)
  If Me.IsAddin Then
    If MsgBox("編集後にメニューを表示すると落ちることがあるので，安全のため他のブックも保存してください．" & vbCrLf & _
              "今開いているすべてのブックを（上書き）保存しますか？", vbYesNo + vbCritical, "注意") = vbYes Then
      Dim book As Workbook
      For Each book In Application.Workbooks
        If Not book.Saved Then book.Save
      Next
    End If
  End If
End Sub


Private Sub Workbook_BeforeClose(Cancel As Boolean)
  Ribbon.Dispose    'リボンのアドレスを無効化
  If CBool(GetSetting(APP, "Addin", "Installed", "False")) Then
    GetAddin().Installed = True
  End If
End Sub


Private Function GetAddin() As AddIn
  On Error GoTo eee:
  Set GetAddin = AddIns(Me.Name)
eee:
End Function


Private Sub Workbook_Open()
  If Me.IsAddin Then
    ショートカットの割り当て
  Else
    Dim a As AddIn
    Set a = GetAddin()
    If a Is Nothing Then
      ' 登録されていない
      SaveSetting APP, "Addin", "Installed", "False"
    Else
      SaveSetting APP, "Addin", "Installed", CStr(a.Installed)
      If a.Installed Then a.Installed = False
    End If
  End If
End Sub

Private Sub 文字の統一()
  Dim Address ' Address
  Dim FSO ' FSO
  Dim Enabled ' Enabled
  Dim Parent ' Parent
  Dim Left ' Left
  Dim Right ' Right
  Dim Mid ' Mid
  Dim Hwnd ' Hwnd
  Dim Dest ' Dest
  Dim Item ' Item
  Dim Value ' Value
  Dim Values ' Values
  Dim Names ' Names
  Dim File  ' File
End Sub

