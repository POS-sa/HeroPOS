Attribute VB_Name = "StringConvert"
Dim i As Integer

Public Function UniStringToByte(UniStr As String, buffer() As Byte, Cnt As Integer) As Integer
    For i = 0 To LenB(StrConv(UniStr, vbFromUnicode)) - 1
        buffer(Cnt + i) = AscB(MidB(StrConv(UniStr, vbFromUnicode), i + 1, 1))
    Next i
    UniStringToByte = i
End Function


