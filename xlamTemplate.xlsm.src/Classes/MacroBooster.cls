Option Explicit

Private mCalc As XlCalculation
Private mScreenUpdate As Boolean
Private mEvent As Boolean
Private mPrintComm As Boolean


Private Sub Class_Initialize()
  mScreenUpdate = Application.ScreenUpdating
  mCalc = Application.Calculation
  mEvent = Application.EnableEvents
  mPrintComm = Application.PrintCommunication
End Sub


Private Sub Class_Terminate()
  Leave
End Sub


Sub Leave()
  With Application
    .Calculation = mCalc
    .ScreenUpdating = mScreenUpdate
    On Error Resume Next ' プリンタ設定をしていない場合にエラーになることがある模様
    .PrintCommunication = mPrintComm
    On Error GoTo 0
    .EnableEvents = mEvent
  End With
End Sub


Sub Enter(Optional ByVal EventsEnabled As Boolean = True)
  Application.Calculation = xlCalculationManual
  Application.ScreenUpdating = False
  Application.PrintCommunication = False
  Application.EnableEvents = EventsEnabled
End Sub


Sub NoCalc()
  Application.Calculation = xlCalculationManual
End Sub


Sub NoScreenUpdating()
  Application.ScreenUpdating = False
End Sub


Sub DisableEvents()
  Application.EnableEvents = False
End Sub


Sub NoPrinterComunication()
  Application.PrintCommunication = False
End Sub

