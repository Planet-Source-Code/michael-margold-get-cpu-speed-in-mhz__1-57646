Attribute VB_Name = "MainModule"
Option Explicit

'Get the speed of the CPU in MHz
Public Function GetCPUSpeedMHz() As Long
  Dim Temp As Long
  Dim Reg As Object
  Set Reg = CreateObject("WScript.Shell")
  Temp = CLng(Reg.regread("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\CentralProcessor\0\~MHz"))
  Set Reg = Nothing
  GetCPUSpeedMHz = Temp
End Function

Sub main()
  MsgBox "CPU Speed is " & CStr(GetCPUSpeedMHz) & " MHz", vbInformation, "http://www.soft-collection.com"
End Sub

