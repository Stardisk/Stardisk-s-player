Attribute VB_Name = "ini"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Function ReadINIKey(Section As String, KeyName As String, filename As String) As String
Dim RetVal As String
RetVal = String(1024, Chr(0))
ReadINIKey = Left(RetVal, GetPrivateProfileString(Section, KeyName, "", RetVal, Len(RetVal), filename))
End Function

Function WriteINIKey(Section As String, KeyName As String, KeyValue As String, filename As String)
        WritePrivateProfileString Section, KeyName, KeyValue, filename
End Function

Function DeleteSection(Section As String, filename As String)
    WritePrivateProfileString Section, 0&, 0&, filename
End Function





