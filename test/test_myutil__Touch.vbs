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

Sub Test_Create
  Touch("temp\tmp")
  AssertEqual True, NewFSO.FileExists("temp\tmp")
  NewFSO.DeleteFile("temp\tmp")
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
