' @import ../lib/myutil.vbs

Option Explicit

Sub Test_Case_OK
  AssertEqual 0, CountInStr("abcdef", "z")
  AssertEqual 1, CountInStr("abcdef", "a")
  AssertEqual 1, CountInStr("abcdef", "c")
  AssertEqual 1, CountInStr("abcdef", "f")

  AssertEqual 0, CountInStr("abcdef", "az")
  AssertEqual 1, CountInStr("abcdef", "ab")
  AssertEqual 1, CountInStr("abcdef", "cd")
  AssertEqual 1, CountInStr("abcdef", "ef")
  AssertEqual 1, CountInStr("abcdef", "abcdef")

  AssertEqual 2, CountInStr("foobarbaz", "o")
  AssertEqual 2, CountInStr("foobarbaz", "ba")

  AssertEqual 0, CountInStr("����������", "��")
  AssertEqual 1, CountInStr("����������", "��")
  AssertEqual 1, CountInStr("����������", "��")
  AssertEqual 1, CountInStr("����������", "��")

  AssertEqual 0, CountInStr("����������", "����")
  AssertEqual 1, CountInStr("����������", "����")
  AssertEqual 1, CountInStr("����������", "����")
  AssertEqual 1, CountInStr("����������", "����")
  AssertEqual 1, CountInStr("����������", "����������")

  AssertEqual 2, CountInStr("�e�X�g�e�L�X�g", "�e")
  AssertEqual 2, CountInStr("�e�X�g�e�L�X�g", "�X�g")

  AssertEqual 1, CountInStr("abc" & vbNewLine, vbNewLine)
  AssertEqual 1, CountInStr("abc" & vbNewLine & "def", vbNewLine)
End Sub

Sub Test_Case_Ng
  Dim cnt
  On Error Resume Next

  Err.Clear
  cnt = CountInStr("", "")
  AssertEqual Err_WrongParam, Err.Number

  Err.Clear
  cnt = CountInStr("", "a")
  AssertEqual Err_WrongParam, Err.Number

  Err.Clear
  cnt = CountInStr("a", "")
  AssertEqual Err_WrongParam, Err.Number
End Sub
