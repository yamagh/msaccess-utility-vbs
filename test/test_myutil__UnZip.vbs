' @import ../lib/myutil.vbs

Option Explicit

Sub Test_UnZip_Fail
  On Error Resume Next

  Err.Clear
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
