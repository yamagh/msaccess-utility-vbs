' @import ../lib/myutil.vbs

Option Explicit

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

Sub Test_1Line_Text
  AssertEqual 1, GetColumnInStr("foobarbaz", 1)
  AssertEqual 2, GetColumnInStr("foobarbaz", 2)
  AssertEqual 3, GetColumnInStr("foobarbaz", 3)
  AssertEqual 4, GetColumnInStr("foobarbaz", 4)
  AssertEqual 5, GetColumnInStr("foobarbaz", 5)
  AssertEqual 6, GetColumnInStr("foobarbaz", 6)
  AssertEqual 7, GetColumnInStr("foobarbaz", 7)
  AssertEqual 8, GetColumnInStr("foobarbaz", 8)
  AssertEqual 9, GetColumnInStr("foobarbaz", 9)
End Sub

Sub Test_String
''  AssertEqual 1,  GetColumnInStr(sample_text, (2*0) + (1))
  AssertEqual 1,  GetColumnInStr(sample_text, (2*1) + (0+1))
''  AssertEqual 1,  GetColumnInStr(sample_text, (2*2) + (0+20+1))
  AssertEqual 1,  GetColumnInStr(sample_text, (2*3) + (0+20+0+1))
  AssertEqual 1,  GetColumnInStr(sample_text, (2*4) + (0+20+0+10+1))
  AssertEqual 1,  GetColumnInStr(sample_text, (2*5) + (0+20+0+10+10+1))
  AssertEqual 1,  GetColumnInStr(sample_text, (2*6) + (0+20+0+10+10+10+1))
  AssertEqual 1,  GetColumnInStr(sample_text, (2*7) + (0+20+0+10+10+10+10+1))
  AssertEqual 1,  GetColumnInStr(sample_text, (2*8) + (0+20+0+10+10+10+10+21+1))

''  AssertEqual 0,  GetColumnInStr(sample_text, (2*0) + (0))
  AssertEqual 20, GetColumnInStr(sample_text, (2*1) + (0+20))
''  AssertEqual 0,  GetColumnInStr(sample_text, (2*2) + (0+20+0))
  AssertEqual 10, GetColumnInStr(sample_text, (2*3) + (0+20+0+10))
  AssertEqual 10, GetColumnInStr(sample_text, (2*4) + (0+20+0+10+10))
  AssertEqual 10, GetColumnInStr(sample_text, (2*5) + (0+20+0+10+10+10))
  AssertEqual 10, GetColumnInStr(sample_text, (2*6) + (0+20+0+10+10+10+10))
  AssertEqual 21, GetColumnInStr(sample_text, (2*7) + (0+20+0+10+10+10+10+21))
  AssertEqual 53, GetColumnInStr(sample_text, (2*8) + (0+20+0+10+10+10+10+21+53))
End Sub

Sub Test_CrLf
  AssertEqual -1, GetColumnInStr(sample_text, (2*0) + (0))
  AssertEqual -1, GetColumnInStr(sample_text, (2*0) + (0+1))
  AssertEqual -1, GetColumnInStr(sample_text, (2*1) + (0+20 + 1))
  AssertEqual -1, GetColumnInStr(sample_text, (2*2) + (0+20+0 + 1))
  AssertEqual -1, GetColumnInStr(sample_text, (2*3) + (0+20+0+10 + 1))
  AssertEqual -1, GetColumnInStr(sample_text, (2*4) + (0+20+0+10+10 + 1))
  AssertEqual -1, GetColumnInStr(sample_text, (2*5) + (0+20+0+10+10+10 + 1))
  AssertEqual -1, GetColumnInStr(sample_text, (2*6) + (0+20+0+10+10+10+10 + 1))
  AssertEqual -1, GetColumnInStr(sample_text, (2*7) + (0+20+0+10+10+10+10+21 + 1))
  AssertEqual -1, GetColumnInStr(sample_text, (2*8) + (0+20+0+10+10+10+10+21+53 + 1))
End Sub

Sub Test_Param_str_Is_Blank
  AssertEqual -1, GetColumnInStr("", 0)

  Dim col
  On Error Resume Next
  col = GetColumnInStr("", 1)
  AssertEqual 513, Err.Number
End Sub

Sub Test_Param_pos
  Dim col

  On Error Resume Next
  col = GetColumnInStr("foobarbaz", -1)
  AssertEqual 513, Err.Number

  Err.Clear
  col = GetColumnInStr("foobarbaz", 10)
  AssertEqual 513, Err.Number
End Sub
