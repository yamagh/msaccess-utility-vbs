Option Explicit

Class MSAccessUtility
  Public Application

  Private Sub Class_Initialize
    Set Application = CreateObject("Access.Application")
    Application.AutomationSecurity = 1
  End Sub

  Private Sub Class_Terminate
    Call Close
    Set Application = Nothing
  End Sub

  Public Default Property Get ToString
    ToString = "AccessUtil"
  End Property

  Public Property Get App
    Set App = Application
  End Property

  Public Sub Open(path)
    Call Application.OpenCurrentDatabase(path)
  End Sub

  Public Sub BypassOpen
    Dim xlApp
    Set xlApp = CreateObject("Excel.Application")

  End Sub

  Public Sub Close
    If Not Application.CurrentDb Is Nothing Then Application.CloseCurrentDatabase
  End Sub
End Class
