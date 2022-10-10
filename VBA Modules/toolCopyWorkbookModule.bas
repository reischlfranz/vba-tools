Attribute VB_Name = "toolCopyWorkbookModule"
Function CopyWorkbook(Optional closeOriginal As Boolean = False, Optional activateNew As Boolean = False) As Workbook
  ' Generates a copy of the workbook including all VBA
  ' New File is then opened and old file closed
  Dim wbA As Workbook, wbB As Workbook
  Dim tempPath As String
  Dim fs As FileSystemObject
  
  Set wbA = ActiveWorkbook
      
  ' Save a temporary copy of original Workbook
  Set fs = CreateObject("Scripting.FileSystemObject")
  tempPath = fs.GetAbsolutePathName(Environ("temp") & "/" & fs.GetTempName() & "." & fs.GetExtensionName(wbA.FullName))
  wbA.SaveCopyAs (tempPath)
  
  ' Create a new workbook using the temporary file as template
  Set wbB = Workbooks.Add(tempPath)
  
  ' open new workbook
  If activateNew Then wbB.Activate
  
  ' Delete temporary file
  fs.DeleteFile tempPath
  
  ' close original file
  If closeOriginal Then wbA.Close
  
  ' Return new workbook object
  Set CopyWorkbook = wbB

End Function
