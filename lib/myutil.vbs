Option Explicit

Const Err_WrongPath = 513
Const Err_WrongFile = 514

Function NewFSO
  Set NewFSO = CreateObject("Scripting.FileSystemObject")
End Function

Function NewShell
  Set NewShell = CreateObject("Shell.Application")
End Function

' File Control Functions
' ======================

Sub CopyFile(source, dest)
  If CheckExistsPath(source) = False Then
    Err.Raise Err_WrongPath, "myutil.CopyFile", "Copy source is not exist."
  End If

  If CheckExistsPath(dest) = False Then
    Err.Raise Err_WrongPath, "myutil.CopyFile", "Copy destination is not exist."
  End If

  NewFSO.CopyFile source, dest
End Sub

Function CheckExistsPath(path)
  With NewFSO
    If .FileExists(path) = True Or .FolderExists(path) = True Then
      CheckExistsPath = True
    End If
  End With
End Function

' Zip Control Functions
' =====================

Sub UnZip(path)
  Dim fso : Set fso = NewFSO
  If fso.GetExtensionName(path) <> "zip" Then
    Err.Raise Err_WrongFile, "myutil.UnZip", "Specified file is not zip file."
  End If

  Dim dir_name : dir_name = fso.GetBaseName(path)
  Dim dir_path : dir_path = fso.BuildPath(fso.GetParentFolderName(path), dir_name)
  If fso.FolderExists(dir_path) = False Then fso.CreateFolder dir_path

  Const FOF_SILENT            = &H04
  Const FOF_RENAMEONCOLLISION = &H08
  Const FOF_NOCONFIRMATION    = &H10
  Const FOF_ALLOWUNDO         = &H40
  Const FOF_FILESONLY         = &H80
  Const FOF_SIMPLEPROGRESS    = &H100
  Const FOF_NOCONFIRMMKDIR    = &H200
  Const FOF_NOERRORUI         = &H400
  Const FOF_NORECURSION       = &H1000

  NewShell.Namespace(dir_path).CopyHere NewShell.Namespace(path).Items, FOF_SILENT + FOF_NOCONFIRMATION
End Sub