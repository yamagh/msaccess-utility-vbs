' @import ../lib/myutil.vbs

Option Explicit

' File Control Functions
' ======================

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

' Zip Control Functions
' =====================

Sub Test_UnZip_Fail
  On Error Resume Next

  Call UnZip("")
  AssertEqual 514, Err.Number

  Err.Clear

  Call UnZip("C:\Windows")
  AssertEqual 514, Err.Number

  Err.Clear

  Call UnZip("C:\Windows\System32\Notepad.exe")
  AssertEqual 514, Err.Number
End Sub

Sub Test_UnZip_Success
  Dim fso : Set fso = NewFSO
  Dim path : path = fso.GetFile("assets\zip\SampleText.zip").Path
  Call UnZip(path)
  AssertEqual True, fso.FolderExists("assets\zip\SampleText")
  AssertEqual True, fso.FileExists("assets\zip\SampleText\SampleText.txt")
  fso.DeleteFolder "assets\zip\SampleText"
End Sub
