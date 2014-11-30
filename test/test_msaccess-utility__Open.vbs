' @import ../lib/myutil.vbs
' @import ../lib/msaccess-utility.vbs

Option Explicit

Dim au

Sub SetUp
  Set au = New MSAccessUtility
End Sub

Sub TearDown
  Set au = Nothing
End Sub

Sub Test_Open
  With NewFSO
    UnZip(.GetFile("assets\sampledb.zip").Path)

    Dim file_path
    file_path = .GetFile("assets\sampledb\sample01.mdb").Path

    au.Open(file_path)
    au.Close
    .DeleteFolder("assets\sampledb")
  End With
End Sub
