Option Explicit
'@Folder RibbonUI

Private Modules    As Collection
Private m_Name     As String

Private m_Macros   As Dictionary


Public Sub Add(ByVal module As MacroModule)
  Modules.Add module, module.Name
End Sub


Property Get Count() As Long
  Count = Modules.Count
End Property


Public Sub Clear()
  Do While Modules.Count > 0
    Modules(1).Clear
    Modules.Remove 1
  Loop
End Sub


'@DefaultMember
Property Get Item(ByVal index_or_key As Variant) As MacroModule
  Set Item = Modules.Item(index_or_key)
End Property


Public Property Get macro(ByVal MacroName As String) As MacroInfo
  If m_Macros.Exists(UCase$(MacroName)) Then Set macro = m_Macros(UCase$(MacroName))
End Property


Public Property Get Macros() As Variant()
  Macros = m_Macros.Items()
End Property


Public Property Get MacroNames() As Variant()
  MacroNames = m_Macros.keys()
End Property


Public Property Get MacroCount() As Long
  MacroCount = m_Macros.Count
End Property


Public Function hasMacro(ByVal MacroName As String) As Boolean
  hasMacro = m_Macros.Exists(UCase$(MacroName))
End Function


Private Sub Class_Initialize()
  Set Modules = New Collection
  Set m_Macros = New Dictionary
End Sub


Private Sub Class_Terminate()
  Call Clear
  Set Modules = Nothing
End Sub


'@Enumerator
Public Function NewEnum() As IEnumVARIANT
  Set NewEnum = Modules.[_NewEnum]
End Function


Public Function Exists(ByVal key As String) As Boolean
  Exists = False
  On Error GoTo eee:
  Exists = Modules.Item(key).Name <> ""
eee:
End Function


Public Property Get Name() As String
  Name = m_Name
End Property


'@Description("VBProjectをパースしてマクロ情報を取得")
Public Sub Parse(ByVal project As VBProject)
  Dim module As MacroModule
  Dim comp As VBComponent
  Dim iComp As Long
  Dim iMacro As Long
  Dim info As MacroInfo

  With New FileSystemObject
    m_Name = .GetFileName(project.Filename)
  End With
  For iComp = 1 To project.VBComponents.Count
    Set comp = project.VBComponents.Item(iComp)
    Set module = New MacroModule
    module.ProjectName = m_Name
    module.Parse comp
    If module.Count > 0 Then
      Modules.Add module, module.Name
      For iMacro = 1 To module.Count
        Set info = module.Item(iMacro)
        info.id = "Macro_" & m_Name & Format(iComp, "_00") & Format(iMacro, "_00")
        m_Macros.Add UCase$(info.Name), info
      Next
    End If
  Next
End Sub