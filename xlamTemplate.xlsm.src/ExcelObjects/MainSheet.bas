'@IgnoreModule ModuleWithoutFolder
Option Explicit


' hidden: True
Sub SaveAsAddin()
  With ThisWorkbook
    If 行末空白チェック() Then Exit Sub
    If RubberDuckチェック() Then Exit Sub

    ' まず自分を保存
    .Save

    Dim chkExport As MSForms.CheckBox

    If Me.shapes("chkCodeExport").ControlFormat.Value Then
      ' マクロをエクスポート
      If exportMacros() And ExtractCustomUI() Then

        'エクスポートできたので，その時のタイムスタンプを記録する．
        With FSO.OpenTextFile(topDir & "\update-time.txt", ForWriting, True)
          .WriteLine Format(Now, "yyyy/mm/dd-HH:MM:SS")
          .Close
        End With
      Else
        'エクスポートに失敗したので続行するかどうか問い合わせる
        If MsgBox("マクロのエクスポートに失敗しましたが，続けますか？", vbYesNo, "確認") = vbNo Then Exit Sub
      End If
    End If

    ' アドインとして保存
    .SaveAs Filename:=AddinPath(), FileFormat:=xlOpenXMLAddIn, AddToMru:=False

    'マクロで閉じると，フックがうまく動かないので，閉じるのは手動にする．
    '.Close SaveChanges:=False
  End With
End Sub


' hidden: True
Sub マクロをエクスポート()
  If exportMacros() And ExtractCustomUI() Then
    MsgBox "成功"
  Else
    MsgBox "失敗"
  End If
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function AddinPath() As String
  AddinPath = FSO.BuildPath(Application.UserLibraryPath, FSO.GetBaseName(ThisWorkbook.Name) & ".xlam")
End Function


Private Function topDir() As String
  Let topDir = ThisWorkbook.FullName & ".src"
End Function


Private Function 行末空白チェック() As Boolean
  Dim vbc As VBComponent
  For Each vbc In ThisWorkbook.VBProject.VBComponents
    Dim i As Long
    With vbc.CodeModule
      For i = 1 To .CountOfLines
        Dim s As String
        s = .Lines(i, 1)
        If s <> RTrim(s) Then
          If MsgBox("行末空白がありますが，続行しますか？", vbYesNo, "確認") = vbNo Then 行末空白チェック = True
          Exit Function
        End If
      Next
    End With
  Next
End Function


Private Function RubberDuckチェック() As Boolean
  Const RD As String = "Rubberduck"
  With ThisWorkbook.VBProject
    Dim r As Reference
    For Each r In .References
      If r.Name = RD Then
        If MsgBox("参照設定にRubberduckがあります．解除しますか？", vbYesNo + vbDefaultButton1, "参照解除の確認") = vbYes Then
          .References.Remove r
          MsgBox "参照設定にRubberduckがありましたので，参照を解除しました．" & vbCrLf & _
            "エラー防止のために再実行をしてください．", vbOKOnly, "Rubbeduck解除"
          RubberDuckチェック = True
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

  ' 出力先フォルダがなければ作成
  If Not FSO.FolderExists(topDir) Then FSO.CreateFolder topDir

  ' 空っぽのフォルダを準備  (フォルダがあれば内容をクリア，なければ作成）
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
  
  ' 書き出し
  ' フォームの場合，frxは書き出されないので，注意．
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


