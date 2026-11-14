Attribute VB_Name = "ini"
Option Explicit

Private Declare Function IniOpen Lib "kernel32.dll" Alias "CreateFileW" (ByVal path As LongPtr, ByVal access As Long, ByVal share As Long, ByVal security As LongPtr, ByVal disposition As Long, ByVal flags As Long, ByVal template As LongPtr) As LongPtr
Private Declare Function IniClose Lib "kernel32.dll" Alias "CloseHandle" (ByVal handle As LongPtr) As Long
Private Declare Function IniSize Lib "kernel32.dll" Alias "GetFileSizeEx" (ByVal handle As LongPtr, ByRef size As LongLong) As Long
Private Declare Function IniWrite Lib "kernel32.dll" Alias "WriteFile" (ByVal handle As LongPtr, ByVal buffer As LongPtr, ByVal count As Long, ByRef actual As Long, ByVal overlapped As LongPtr) As Long
Public Declare Function IniReadW Lib "kernel32.dll" Alias "GetPrivateProfileStringW" (ByVal section As LongPtr, ByVal key As LongPtr, ByVal defaultValue As LongPtr, ByVal result As LongPtr, ByVal count As Long, ByVal filename As LongPtr) As Long
Public Declare Function IniWriteW Lib "kernel32.dll" Alias "WritePrivateProfileStringW" (ByVal section As LongPtr, ByVal key As LongPtr, ByVal value As LongPtr, ByVal filename As LongPtr) As Long

Public Function ReadINIKey(ByVal Section As String, ByVal KeyName As String, ByVal filename As String) As String
    Dim buffer As String
    Dim length As Long
    buffer = String$(32768, vbNullChar)
    length = IniReadW(StrPtr(Section), StrPtr(KeyName), 0, StrPtr(buffer), Len(buffer), StrPtr(filename))
    If length >= Len(buffer) - 1 Then Err.Raise 5, "ReadINIKey", "Слишком длинная запись в INI-файле."
    ReadINIKey = Left$(buffer, length)
End Function

Public Sub WriteINIKey(ByVal Section As String, ByVal KeyName As String, ByVal KeyValue As String, ByVal filename As String)
    Dim valueBuffer As String
    EnsureINIFile filename
    valueBuffer = KeyValue & vbNullChar
    If IniWriteW(StrPtr(Section), StrPtr(KeyName), StrPtr(valueBuffer), StrPtr(filename)) = 0 Then RaisePlayerFileError Err.LastDllError, "Запись настроек"
End Sub

Public Sub DeleteSection(ByVal Section As String, ByVal filename As String)
    EnsureINIFile filename
    If IniWriteW(StrPtr(Section), 0, 0, StrPtr(filename)) = 0 Then RaisePlayerFileError Err.LastDllError, "Удаление раздела INI"
End Sub

' Create new/empty settings files as UTF-16LE. Existing content is untouched.
Public Sub EnsureINIFile(ByVal filename As String)
    Dim handle As LongPtr
    Dim size As LongLong
    Dim actual As Long
    Dim bom As String
    Dim savedNumber As Long
    Dim savedDescription As String
    handle = -1
    On Error GoTo failed
    handle = IniOpen(StrPtr(filename), &H40000000, 0, 0, 4, 0, 0) ' OPEN_ALWAYS
    If handle = -1 Then RaisePlayerFileError Err.LastDllError, "Открытие INI"
    If IniSize(handle, size) = 0 Then RaisePlayerFileError Err.LastDllError, "Размер INI"
    If size = 0 Then
        bom = ChrW$(&HFEFF&)
        If IniWrite(handle, StrPtr(bom), 2, actual, 0) = 0 Then RaisePlayerFileError Err.LastDllError, "Создание INI"
        If actual <> 2 Then Err.Raise 5, "EnsureINIFile", "Не удалось записать заголовок INI."
    End If
    IniClose handle
    Exit Sub
failed:
    savedNumber = Err.Number
    savedDescription = Err.Description
    On Error Resume Next
    If handle <> -1 Then IniClose handle
    On Error GoTo 0
    Err.Raise savedNumber, "EnsureINIFile", savedDescription
End Sub
