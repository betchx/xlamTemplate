'@IgnoreModule ModuleWithoutFolder
Option Explicit


' hidden: True
Sub SaveAsAddin()
  With ThisWorkbook
    If �s���󔒃`�F�b�N() Then Exit Sub
    If RubberDuck�`�F�b�N() Then Exit Sub

    ' �܂�������ۑ�
    .Save

    Dim chkExport As MSForms.CheckBox

    If Me.shapes("chkCodeExport").ControlFormat.Value Then
      ' �}�N�����G�N�X�|�[�g
      If exportMacros() And ExtractCustomUI() Then

        '�G�N�X�|�[�g�ł����̂ŁC���̎��̃^�C���X�^���v���L�^����D
        With FSO.OpenTextFile(topDir & "\update-time.txt", ForWriting, True)
          .WriteLine Format(Now, "yyyy/mm/dd-HH:MM:SS")
          .Close
        End With
      Else
        '�G�N�X�|�[�g�Ɏ��s�����̂ő��s���邩�ǂ����₢���킹��
        If MsgBox("�}�N���̃G�N�X�|�[�g�Ɏ��s���܂������C�����܂����H", vbYesNo, "�m�F") = vbNo Then Exit Sub
      End If
    End If

    ' �A�h�C���Ƃ��ĕۑ�
    .SaveAs Filename:=AddinPath(), FileFormat:=xlOpenXMLAddIn, AddToMru:=False

    '�}�N���ŕ���ƁC�t�b�N�����܂������Ȃ��̂ŁC����͎̂蓮�ɂ���D
    '.Close SaveChanges:=False
  End With
End Sub


' hidden: True
Sub �}�N�����G�N�X�|�[�g()
  If exportMacros() And ExtractCustomUI() Then
    MsgBox "����"
  Else
    MsgBox "���s"
  End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function AddinPath() As String
  AddinPath = FSO.BuildPath(Application.UserLibraryPath, FSO.GetBaseName(ThisWorkbook.Name) & ".xlam")
End Function


Private Function topDir() As String
  Let topDir = ThisWorkbook.FullName & ".src"
End Function


Private Function �s���󔒃`�F�b�N() As Boolean
  Dim vbc As VBComponent
  For Each vbc In ThisWorkbook.VBProject.VBComponents
    Dim i As Long
    With vbc.CodeModule
      For i = 1 To .CountOfLines
        Dim s As String
        s = .Lines(i, 1)
        If s <> RTrim(s) Then
          If MsgBox("�s���󔒂�����܂����C���s���܂����H", vbYesNo, "�m�F") = vbNo Then �s���󔒃`�F�b�N = True
          Exit Function
        End If
      Next
    End With
  Next
End Function


Private Function RubberDuck�`�F�b�N() As Boolean
  Const RD As String = "Rubberduck"
  With ThisWorkbook.VBProject
    Dim r As Reference
    For Each r In .References
      If r.Name = RD Then
        If MsgBox("�Q�Ɛݒ��Rubberduck������܂��D�������܂����H", vbYesNo + vbDefaultButton1, "�Q�Ɖ����̊m�F") = vbYes Then
          .References.Remove r
          MsgBox "�Q�Ɛݒ��Rubberduck������܂����̂ŁC�Q�Ƃ��������܂����D" & vbCrLf & _
            "�G���[�h�~�̂��߂ɍĎ��s�����Ă��������D", vbOKOnly, "Rubbeduck����"
          RubberDuck�`�F�b�N = True
        End If
        Exit Function
      End If
    Next
  End With
End Function


Private Function Escape(s As String) As String
  Escape = Replace(s, " ", "` ")
End Function


Private Function ExtractCustomUI() As Boolean
  With New IWshRuntimeLibrary.WshShell
    Dim arc_path As String
    arc_path = FSO.BuildPath(ThisWorkbook.path, FSO.GetBaseName(ThisWorkbook.Name) & ".zip")
    FSO.CopyFile ThisWorkbook.FullName, arc_path, True


    Dim Dest As String
    Dest = FSO.BuildPath(ThisWorkbook.path, "melt")

    Dim melt_command As String
    melt_command = "Expand-Archive -LiteralPath " & Escape(arc_path) & " -DestinationPath " & Escape(Dest) & " -Force"

    Dim ps_command As String
    ps_command = "powershell -NoLogo -ExecutionPolicy RemoteSigned -Command "

    'Debug.Print ps_command & melt_command
    If .Run(ps_command & melt_command, 0, True) > 0 Then Exit Function

    ' copy
    Dim src_dir As String
    src_dir = ThisWorkbook.FullName & ".src"

    Dim copy_command As String
    copy_command = "Copy-Item -LiteralPath " & Escape(FSO.BuildPath(Dest, "customUI")) & _
      " -Destination " & Escape(src_dir) & " -Force -Recurse"

    If .Run(ps_command & copy_command, 0, True) > 0 Then Exit Function

    FSO.DeleteFolder Dest, True
    FSO.DeleteFile arc_path, True
  End With
  ExtractCustomUI = True
End Function


Private Function exportMacros() As Boolean
  Dim comp As VBComponent

  Dim out_dirs As New Dictionary
  out_dirs(vbext_ct_StdModule) = "Modules"
  out_dirs(vbext_ct_ClassModule) = "Classes"
  out_dirs(vbext_ct_MSForm) = "Forms"
  out_dirs(vbext_ct_Document) = "ExcelObjects"
  Dim out_dir As String

  Dim exts As New Dictionary
  exts(vbext_ct_StdModule) = ".bas"
  exts(vbext_ct_ClassModule) = ".cls"
  exts(vbext_ct_MSForm) = ".frm"
  exts(vbext_ct_Document) = ".bas"

  On Error GoTo eee:

  ' �o�͐�t�H���_���Ȃ���΍쐬
  If Not FSO.FolderExists(topDir) Then FSO.CreateFolder topDir

  ' ����ۂ̃t�H���_������  (�t�H���_������Γ��e���N���A�C�Ȃ���΍쐬�j
  Dim key
  For Each key In out_dirs
    out_dir = FSO.BuildPath(topDir, out_dirs(key))
    If FSO.FolderExists(out_dir) Then
      Dim d As Scripting.Folder
      Set d = FSO.GetFolder(out_dir)
      Dim f As Scripting.File
      For Each f In d.Files
        f.Delete
      Next
    Else
      FSO.CreateFolder out_dir
    End If
  Next
  
  ' �����o��
  ' �t�H�[���̏ꍇ�Cfrx�͏����o����Ȃ��̂ŁC���ӁD
  For Each comp In ThisWorkbook.VBProject.VBComponents
    If comp.CodeModule.CountOfLines > 0 Then
      out_dir = FSO.BuildPath(topDir, out_dirs(comp.Type))

      Dim out_path As String
      out_path = FSO.BuildPath(out_dir, comp.Name + exts(comp.Type))

      Dim ts As TextStream
      Set ts = FSO.OpenTextFile(out_path, ForWriting, Create:=True)
      ts.Write comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
      ts.Close
    End If
  Next

  Let exportMacros = True

eee:

End Function


