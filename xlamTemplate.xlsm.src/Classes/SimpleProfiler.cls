Option Explicit

Implements IProfiler

Dim start As Single
Dim prev As Single


Private Function �E��(�l As Single, Optional �� As Long = 10) As String
  If �� < 5 Then �� = 5
  �E�� = Right$(String(��, " ") & Format(�l, "0.000"), ��)
End Function

Private Sub Class_Initialize()
  Debug.Print "     Timer  �o�ߎ���  ���v����  �R�����g"
  start = Timer
  prev = start
End Sub


Private Sub Class_Terminate()
  IProfiler_Mark "End"
End Sub


Sub IProfiler_Mark(Optional comment As String = "")
  Dim t As Single
  t = Timer
  Debug.Print �E��(t) & �E��(t - start) & �E��(t - prev) & "  " & comment
  prev = t
End Sub