' @import ../lib/myutil.vbs
' @import ../lib/access-utility.vbs

Option Explicit

Dim au, cnt, dest_path

Const SOURCE = "assets\sampledb\sample01.mdb"
Const DEST   = "assets\sampledb\sample01-@@@.mdb"

Sub SetUp
  cnt = Replace(Space(3-Len(cnt)) & cnt, " ", 0)
  dest_path = Replace(DEST, "@@@", cnt)
  NewFSO.CopyFile SOURCE, dest_path
  Set au = New AccessUtility
End Sub

Sub TearDown
  Set au = Nothing
End Sub

Sub Test_Class_Initialize
  Dim acApp
  Set acApp = CreateObject("Access.Application")
  AssertEqual acApp.Name, au.Application.Name
  AssertEqual acApp.Version, au.Application.Version
  Set acApp = Nothing
End Sub
