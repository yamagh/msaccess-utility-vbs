' @import ../lib/myutil.vbs

Option Explicit

Sub SetUp
  With NewFSO
    .CreateFolder("source")
    .CreateFolder("dest")
  End With
End Sub

Sub TearDown
  With NewFSO
    .DeleteFolder("source")
    .DeleteFolder("dest")
  End With
End Sub

Sub Test_CopyFile_WrongPath
  On Error Resume Next

  Err.Clear
  Call CopyFile("", "C:\Windows\Temp")
  AssertEqual 513, Err.Number

  Err.Clear
  Call CopyFile("C:\Windows\Temp", "")
  AssertEqual 513, Err.Number
End Sub

Sub Test_CopyFile_Success
  Touch "source\foo.txt"
  CopyFile "source\foo.txt", "dest\foo.txt"
  AssertEqual True, NewFSO.FileExists("dest\foo.txt")
End Sub
