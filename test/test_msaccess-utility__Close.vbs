' @import ../lib/myutil.vbs
' @import ../lib/msaccess-utility.vbs

Option Explicit

Dim au, file_path

Sub SetUp
  With NewFSO
    UnZip(.GetFile("assets\sampledb.zip").Path)
    file_path = .GetFile("assets\sampledb\sample01.mdb").Path
  End With
  Set au = New MSAccessUtility
  au.Open(file_path)
End Sub

Sub TearDown
  Set au = Nothing
  NewFSO.DeleteFolder("assets\sampledb")
End Sub

Sub Test_Close
  Dim current_path
  current_path = NewFSO.GetFolder(".").Path
  Dim db_path
  db_path = NewFSO.BuildPath(current_path, "assets\sampledb\sample01.mdb")
  AssertEqual db_path, au.Application.CurrentDb.Name
  au.Close
  AssertEqual True, au.Application.CurrentDb is Nothing
End Sub
