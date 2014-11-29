' @import ../lib/myutil.vbs

Option Explicit

Sub SetUp
  With NewFSO
    If .FolderExists("temp") = False Then .CreateFolder("temp")
  End With
End Sub

Sub TearDown
  NewFSO.DeleteFolder("temp")
End Sub

Sub Test_Create_01
  Const TEMP_FILE_PATH = "temp\tmp"
  Touch(TEMP_FILE_PATH)
  AssertEqual True, NewFSO.FileExists(TEMP_FILE_PATH)
  NewFSO.DeleteFile(TEMP_FILE_PATH)
End Sub

Sub Test_Create_02
  Const TEMP_FILE_PATH = "temp\sub\tmp"
  Touch(TEMP_FILE_PATH)
  AssertEqual False, NewFSO.FileExists(TEMP_FILE_PATH)
End Sub

Sub Test_Update_Timestamp
  NewFSO.CreateTextFile("temp\tmp")
  Dim before
  before = NewFSO.GetFile("temp\tmp").DateLastModified
  WScript.Sleep 1000
  Touch("temp\tmp")
  Dim after
  after = NewFSO.GetFile("temp\tmp").DateLastModified
  AssertWithMessage before < after, "Timestamp is same."
End Sub
