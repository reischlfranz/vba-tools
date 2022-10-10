Attribute VB_Name = "toolSystemTickCount"
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Function GetSystemTickCount() As Long
    GetSystemTickCount = GetTickCount
End Function

