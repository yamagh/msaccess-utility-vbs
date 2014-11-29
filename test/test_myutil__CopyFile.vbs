' @import ../lib/myutil.vbs

Option Explicit

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
  CopyFile "C:\tmp\s\source.txt", "C:\tmp\d\"
End Sub

Sub Test_CheckExistsPath
  AssertEqual True, CheckExistsPath("C:\Windows")
  AssertEqual True, CheckExistsPath("C:\Windows\System32\wscript.exe")

  AssertEqual False, CheckExistsPath("")
  AssertEqual False, CheckExistsPath("C:\Apple")
  AssertEqual False, CheckExistsPath("C:\Windows\System32\wscript.hoge")
End Sub
