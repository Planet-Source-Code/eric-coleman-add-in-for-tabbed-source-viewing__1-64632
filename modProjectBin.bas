Attribute VB_Name = "Module1"
Option Explicit

'C:\Program Files\Microsoft Visual Studio\MSDN98\98VSa\1033\SAMPLES\VB98\Taborder


'#####
'Purpose: This module has only one purpose and
'         is used only once.
'To Use:  To use this project, you must first
'         compile the project.  This registers
'         the activeX dll, then you must
'         call the AddToINI function.  Open
'         the Immediate window (Ctrl + G) in the Visual
'         Basic IDE and type AddToIni.
'         After you press Enter you should see
'         a MessageBox that tells you the Add-In
'         has successfully been installed.
'#####

Declare Function WritePrivateProfileString& Lib _
"kernel32" Alias "WritePrivateProfileStringA" _
(ByVal AppName$, ByVal KeyName$, ByVal _
keydefault$, ByVal FileName$)
Public Sub AddToINI()
Dim rc As Long
rc = WritePrivateProfileString("Add-Ins32", _
"VBgamerSourceBin.MainEOMClass", "0", "VBADDIN.INI")
MsgBox "Add-in is now entered in VBADDIN.INI file."
End Sub

Public Function IsCompiled() As Boolean
  On Local Error Resume Next
  Debug.Assert 1 / 0
  IsCompiled = Not (Err.Number <> 0)
  Err.Clear
End Function



