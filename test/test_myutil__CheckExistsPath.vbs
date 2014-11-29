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

Sub Test_File_Check_Exists
  Const TEMP_FILE_PATH = "temp\foo.txt"
  NewFSO.CreateTextFile(TEMP_FILE_PATH)
  AssertEqual True, CheckExistsPath(TEMP_FILE_PATH)
  AssertEqual True, CheckExistsPath("C:\Windows\System32\wscript.exe")
End Sub

Sub Test_File_Check_Not_Exists
  AssertEqual False, CheckExistsPath("foo\bar.txt")
  AssertEqual False, CheckExistsPath("C:\Windows\System32\wscript.foobar")
End Sub

Sub Test_Folder_Check_Exists
  Const TEMP_FOLDER_PATH = "temp\foo\"
  NewFSO.CreateFolder(TEMP_FOLDER_PATH)
  AssertEqual True, CheckExistsPath(TEMP_FOLDER_PATH)
End Sub

Sub Test_Folder_Check_Not_Exists
  AssertEqual False, CheckExistsPath("foo\bar\")
  AssertEqual False, CheckExistsPath("")
  AssertEqual False, CheckExistsPath("C:\Windows\Apple\Banana\Foo\Bar\")
End Sub
