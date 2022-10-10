Sub VariantTest()

  Dim a As Variant
  Dim b As Variant
  Dim ar(0 To 3) As Variant
  Dim ws As Worksheet
  Set ws = ThisWorkbook.Sheets(1)
  
'  ws.Range("testRange").Select
  
  a = ws.Range("testRange").Value2
  

  ReDim b(1 To UBound(a, 1), 1 To 2)
  For i = 1 To UBound(a, 1)
    b(i, 1) = a(i, 1)
    b(i, 2) = a(i, 4)
  Next i
  
'
'  a.Select
  Stop

End Sub