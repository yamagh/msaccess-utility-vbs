' @import ../lib/myutil.vbs
' @import ../lib/stdlib.vbs

Option Explicit

Dim text
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

Sub Test_Taret_Text_Have_1Line
  AssertEqual "foobar", GetLinesInStr("foobar", 1, 1)
End Sub

Sub Test_Taret_Text_Have_AnyLines_And_Get_1Line
  AssertEqual "@TEST@TEST", GetLinesInStr(sample_text, 4, 1)
End Sub

Sub Test_Taret_Text_Have_AnyLines_And_Get_2lines
  text = "This is sample text."
  AssertEqual text, GetLinesInStr(sample_text, 2, 1)
End Sub

Sub Test_Taret_Text_Have_AnyLines_And_Get_AnyLines
  text = "@TEST@TEST" & vbNewLine & "@tEST@teST"
  AssertEqual text, GetLinesInStr(sample_text, 4, 2)
End Sub

Sub Test_Line_Get_Option_Minus1
  text = "1234567890-1234567890" _
       & vbNewLine & "abcdefghijklmnopqrstuvwxyz-abcdefghijklmnopqrstuvwxyz" _
       & vbNewLine & ""
  AssertEqual text, GetLinesInStr(sample_text, 8, -1)
End Sub
