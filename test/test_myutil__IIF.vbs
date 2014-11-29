' @import ../lib/myutil.vbs

Option Explicit

Sub Test_String
  AssertEqual "foo", IIF(True, "foo", "bar")
  AssertEqual "bar", IIF(False, "foo", "bar")
End Sub

Sub Test_Integer
  AssertEqual 32768, IIF(True, 32768, -32768)
  AssertEqual -32768, IIF(False, 32768, -32768)
End Sub

Sub Test_Currency
  AssertEqual 922337203685477.5808, IIF(True, 922337203685477.5808, -922337203685477.5808)
  AssertEqual -922337203685477.5808, IIF(False, 922337203685477.5808, -922337203685477.5808)
End Sub

Sub Test_Double
  AssertEqual 4.94065645841247E-324, IIF(True, 4.94065645841247E-324, -4.94065645841247E-324)
  AssertEqual -4.94065645841247E-324, IIF(False, 4.94065645841247E-324, -4.94065645841247E-324)
End Sub

Sub Test_Date
  AssertEqual #2020/12/31#, IIF(True, #2020/12/31#, #2010/1/1#)
  AssertEqual #2010/1/1#, IIF(False, #2020/12/31#, #2010/1/1#)
End Sub

Sub Test_Bool
  AssertEqual False, IIF(True, False, True)
  AssertEqual True, IIF(False, False, True)
End Sub

Sub Test_Object
  Dim sh1
  Set sh1 = CreateObject("Shell.Application")
  Dim sh2
  Set sh2 = CreateObject("Shell.Application")
  AssertSame sh1, IIF(True, sh1, sh2)
  AssertSame sh2, IIF(False, sh1, sh2)
End Sub
