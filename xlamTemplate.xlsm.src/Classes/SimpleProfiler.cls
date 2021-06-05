Option Explicit

Implements IProfiler

Dim start As Single
Dim prev As Single


Private Function 右寄せ(値 As Single, Optional 幅 As Long = 10) As String
  If 幅 < 5 Then 幅 = 5
  右寄せ = Right$(String(幅, " ") & Format(値, "0.000"), 幅)
End Function

Private Sub Class_Initialize()
  Debug.Print "     Timer  経過時間  所要時間  コメント"
  start = Timer
  prev = start
End Sub


Private Sub Class_Terminate()
  IProfiler_Mark "End"
End Sub


Sub IProfiler_Mark(Optional comment As String = "")
  Dim t As Single
  t = Timer
  Debug.Print 右寄せ(t) & 右寄せ(t - start) & 右寄せ(t - prev) & "  " & comment
  prev = t
End Sub