' @import ../lib/myutil.vbs
' @import ../lib/msaccess-utility.vbs

Option Explicit

Dim au

Const tmp_dir     = "temp"
Const asserts_dir = "assets"

Sub SetUp
  Set au = New AccessUtility
  NewFSO.CreateFolder tmp_dir
End Sub

Sub TearDown
  NewFSO.DeleteFolder tmp_dir
  Set au = Nothing
End Sub

' Class_Initialize Test
' =====================

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
  AssertEqual "AccessUtil", au
End Sub

' Property App
' ============

Sub Test_Property_App
  AssertSame au.Application, au.App
End Sub

' Open
' ====

Sub Test_Open
  Call UnZip(NewFSO.GetFile("assets\access\sampledb.zip").Path)
  Dim current_dir : current_dir = NewFSO.GetFolder(".").Path
  Dim file_path : file_path = NewFSO.BuildPath(current_dir, "assets\access\sampledb\�C�x���g�Ǘ�.accdb")
  Call au.Open(file_path)
  Call au.Close
  NewFSO.DeleteFolder("assets\access\sampledb")
End Sub
