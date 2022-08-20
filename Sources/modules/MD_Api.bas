Attribute VB_Name = "MD_Api"
'@IgnoreModule
'@Folder("Demo")
Option Compare Database
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As LongPtr)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
#End If
