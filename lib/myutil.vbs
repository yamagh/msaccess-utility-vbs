Option Explicit

Const Err_WrongPath  = 513
Const Err_WrongFile  = 514
Const Err_WrongParam = 515

Function NewFSO
  Set NewFSO = CreateObject("Scripting.FileSystemObject")
End Function

Function NewShell
  Set NewShell = CreateObject("Shell.Application")
End Function

Function NewDic
  Set NewDic = CreateObject("Scripting.Dictionary")
End Function

Function NewRegExp(pattern, opt)
  Dim re
  Set re = New RegExp
  re.Pattern = pattern
  If InStr(opt, "i") Then
    re.IgnoreCase = True
  End If
  If InStr(opt, "g") Then
    re.Global = True
  End If
  If InStr(opt, "m") Then
    re.MultiLine = True
  End If
  Set NewRegExp = re
End Function

' File Control Functions
' ======================

Sub CopyFile(source, dest)
  If CheckExistsPath(source) = False Then
    Err.Raise Err_WrongPath, "myutil.CopyFile", "Copy source is not exist."
  End If

  Dim parent
  parent = NewFSO.GetParentFolderName(dest)
  If CheckExistsPath(parent) = False Then
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

Sub Touch(path)
  Dim parent_path
  If CheckExistsPath(path) Then
    With NewFSO
      Dim file_path
      file_path = .GetFile(path).Path
      parent_path = .GetParentFolderName(file_path)
      Dim file_name
      file_name = .GetFileName(path)
    End With
    NewShell.NameSpace(parent_path).Items.Item(file_name).ModifyDate = Now
  Else
    With NewFSO
      parent_path = .GetParentFolderName(path)
      If CheckExistsPath(parent_path) Then
        .CreateTextFile(path)
      End If
    End With
  End If
End Sub

' Zip Control Functions
' =====================

Sub UnZip(path)
  Const FOF_SILENT            = &H04
  Const FOF_RENAMEONCOLLISION = &H08
  Const FOF_NOCONFIRMATION    = &H10
  Const FOF_ALLOWUNDO         = &H40
  Const FOF_FILESONLY         = &H80
  Const FOF_SIMPLEPROGRESS    = &H100
  Const FOF_NOCONFIRMMKDIR    = &H200
  Const FOF_NOERRORUI         = &H400
  Const FOF_NORECURSION       = &H1000

  With NewFSO
    If .GetExtensionName(path) <> "zip" Then
      Err.Raise Err_WrongFile, "myutil.UnZip", "Specified file is not zip file."
    End If
    Dim dir_name
    dir_name = .GetBaseName(path)
    Dim dir_path
    dir_path = .BuildPath(.GetParentFolderName(path), dir_name)
    If .FolderExists(dir_path) = False Then
      .CreateFolder dir_path
    End If
  End With

  With NewShell
    .Namespace(dir_path).CopyHere .Namespace(path).Items, FOF_SILENT + FOF_NOCONFIRMATION
  End With
End Sub

' String Control Functions
' ========================

Function CountInStr(str, find)
  If Len(str) = 0 Or Len(find) = 0 Then
    Err.Raise Err_WrongParam, "myutil.CountInStr", "Specified params is wrong."
  End If
  CountInStr = UBound(Split(str, find))
End Function

' # Args
' str
'     Source strings for search.
' find
'     Search query in `str`. You can use RegExp pattern string.
' opt
'     RegExp Option.
'     Ex) "igm"
'     Case: "i" IgnoreCase = True
'           "g" Global = True
'           "m" MultiLine = True
Function FindInStr(str, find, opt)
  Dim found
  Set found = NewDic
  Dim mt
  For Each mt In NewRegExp(find, opt).Execute(str)
    Dim dic
    Set dic = NewDic
    dic.Add "Row", CountInStr(Left(str, mt.FirstIndex + 1), vbNewLine) + 1
    dic.Add "Column", GetColumnInStr(str, mt.FirstIndex+1)
    dic.Add "Match", mt.Value
    dic.Add "Text", GetLinesInStr(str, dic("Row"), CountInStr(mt.Value, vbNewLine) + 1)
    found.Add found.Count + 1, dic
  Next
  Set FindInStr = found
End Function

' # Args
' str
'     Source strings.
' start
'     Start of lines to get.
' line_count
'     Count of lines to get.
Function GetLinesInStr(str, start, line_count)
  Dim text
  Dim str_arr
  str_arr = Split(str, vbNewLine)
  Dim max_lines
  max_lines = CountInStr(str, vbNewLine) + 1
  line_count = IIF(line_count = -1, max_lines - start + 1, line_count)
  Dim i
  For i = 1 To UBound(str_arr) + 1
    If start <= i And i < (start + line_count -1) Then
      text = text & str_arr(i-1) & vbNewLine
    ElseIf i = (start + line_count -1) Then
      text = text & str_arr(i-1)
    End If
  Next
  GetLinesInStr = text
End Function

Function GetColumnInStr(str, pos)
  If Len(str) < pos Or pos < 0 Then
    Err.Raise 513, "myutil.GetColumnInStr", "Specified position is wrong."
  End If
  Dim col
  col = -1
  Dim i
  For i = 1 To pos
    Dim c
    c = Mid(str, i, 1)
    If c = vbCr Or c = vbLf Then
      col = -1
    Else
      col = IIF(col=-1, 1, col + 1)
    End If
  Next
  GetColumnInStr = col
End Function

' Other Functions
' ===============

Function IIF(exp, t, f)
  If exp = True Then
    If IsObject(t) Then
      Set IIF = t
    Else
      IIF = t
    End If
  Else
    If IsObject(t) Then
      Set IIF = f
    Else
      IIF = f
    End If
  End If
End Function
