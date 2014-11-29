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

  AssertEqual 0, CountInStr("あいうえお", "ん")
  AssertEqual 1, CountInStr("あいうえお", "あ")
  AssertEqual 1, CountInStr("あいうえお", "う")
  AssertEqual 1, CountInStr("あいうえお", "お")

  AssertEqual 0, CountInStr("あいうえお", "あん")
  AssertEqual 1, CountInStr("あいうえお", "あい")
  AssertEqual 1, CountInStr("あいうえお", "うえ")
  AssertEqual 1, CountInStr("あいうえお", "えお")
  AssertEqual 1, CountInStr("あいうえお", "あいうえお")

  AssertEqual 2, CountInStr("テストテキスト", "テ")
  AssertEqual 2, CountInStr("テストテキスト", "スト")

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
