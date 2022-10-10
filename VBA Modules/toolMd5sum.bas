Attribute VB_Name = "toolMd5sum"
' Module to enable MD5 hash within Excel formula over multiple cells
' RSF ©2022, version 1.0 2022-10-10

' From https://stackoverflow.com/a/72312483/4737365
Function StringToMD5Hex(ByVal s As String) As String
Dim enc As Object
Dim bytes() As Byte
Dim pos As Long
Dim outstr As String

Set enc = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")

bytes = StrConv(s, vbFromUnicode)
bytes = enc.ComputeHash_2(bytes)

For pos = LBound(bytes) To UBound(bytes)
   outstr = outstr & LCase(Right("0" & Hex(bytes(pos)), 2))
Next pos

StringToMD5Hex = outstr
Set enc = Nothing
End Function

Function Md5Hash(ByVal r As Range) As String
  Dim c As Variant, s As String, SEP_CHAR As String
  SEP_CHAR = "%"
  
  c = r.Value2
  
  If IsEmpty(c) Then
    'Debug.Print "empty"
    s = ""
  End If
  
  Dim isFirstVal As Boolean
  isFirstVal = True
  
  For Each v In c
    If isFirstVal Then
      s = v
      isFirstVal = False
    Else
      s = s & SEP_CHAR & v
    End If
  Next v
  
  'Debug.Print s
  Md5Hash = StringToMD5Hex(s)
  
End Function
