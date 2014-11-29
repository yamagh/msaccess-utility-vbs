' @import ../lib/myutil.vbs
' @import ../lib/stdlib.vbs

Option Explicit

Dim found
Dim sample_text

sample_text = "" _
& vbNewLine & "This is sample text." _
& vbNewLine & "" _
& vbNewLine & "@TEST@TEST" _
& vbNewLine & "@tEST@teST" _
& vbNewLine & "@tesT@test" _
& vbNewLine & "@test@test" _
& vbNewLine & "1234567890-1234567890" _
& vbNewLine & "abcdefghijklmnopqrstuvwxyz-abcdefghijklmnopqrstuvwxyz" _
& vbNewLine & ""

Sub Test_Simple_Hit1
  Set found = FindInStr("foobar", "a", "")
  AssertEqual 1, found.Count

  AssertEqual 1, found(1)("Row")
  AssertEqual 5, found(1)("Column")
  AssertEqual "a", found(1)("Match")
  AssertEqual "foobar", found(1)("Text")
End Sub

Sub Test_Opt_Global_True
  Set found = FindInStr("foobar", "o", "g")
  AssertEqual 2, found.Count

  AssertEqual 1, found(1)("Row")
  AssertEqual 1, found(2)("Row")
  AssertEqual 2, found(1)("Column")
  AssertEqual 3, found(2)("Column")
  AssertEqual "o", found(1)("Match")
  AssertEqual "o", found(2)("Match")
  AssertEqual "foobar", found(1)("Text")
  AssertEqual "foobar", found(2)("Text")
End Sub

Sub Test_Opt_Global_False
  Set found = FindInStr("foobar", "o", "")
  AssertEqual 1, found.Count

  AssertEqual 1, found(1)("Row")
  AssertEqual 2, found(1)("Column")
  AssertEqual "o", found(1)("Match")
  AssertEqual "foobar", found(1)("Text")
End Sub
