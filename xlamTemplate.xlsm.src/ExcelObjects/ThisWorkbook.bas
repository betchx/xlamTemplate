'@IgnoreModule ModuleWithoutFolder
Option Explicit


Private Sub Workbook_AddinUninstall()
  �V���[�g�J�b�g�̉���
End Sub


Private Sub Workbook_AfterSave(ByVal Success As Boolean)
  If Me.IsAddin Then
    If MsgBox("�ҏW��Ƀ��j���[��\������Ɨ����邱�Ƃ�����̂ŁC���S�̂��ߑ��̃u�b�N���ۑ����Ă��������D" & vbCrLf & _
              "���J���Ă��邷�ׂẴu�b�N���i�㏑���j�ۑ����܂����H", vbYesNo + vbCritical, "����") = vbYes Then
      Dim book As Workbook
      For Each book In Application.Workbooks
        If Not book.Saved Then book.Save
      Next
    End If
  End If
End Sub


Private Sub Workbook_BeforeClose(Cancel As Boolean)
  Ribbon.Dispose    '���{���̃A�h���X�𖳌���
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
    �V���[�g�J�b�g�̊��蓖��
  Else
    Dim a As AddIn
    Set a = GetAddin()
    If a Is Nothing Then
      ' �o�^����Ă��Ȃ�
      SaveSetting APP, "Addin", "Installed", "False"
    Else
      SaveSetting APP, "Addin", "Installed", CStr(a.Installed)
      If a.Installed Then a.Installed = False
    End If
  End If
End Sub

Private Sub �����̓���()
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

