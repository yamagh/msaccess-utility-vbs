' @import ../lib/myutil.vbs
' @import ../lib/msaccess-utility.vbs

Option Explicit

Dim au

Sub SetUp
  Set au = New MSAccessUtility
End Sub

Sub TearDown
  Set au = Nothing
End Sub

' Initialize/Terminate
' ====================

Sub Test_Class_Initialize
  Dim acApp
  Set acApp = CreateObject("Access.Application")
  AssertEqual acApp.Name, au.Application.Name
  AssertEqual acApp.Version, au.Application.Version
  Set acApp = Nothing
End Sub

' DefaultProperty
' ===============

Sub Test_DefaultProperty
AssertEqual "AccessUtility", au
End Sub

' Property App
' ============

Sub Test_Property_App
AssertSame au.Application, au.App
End Sub
