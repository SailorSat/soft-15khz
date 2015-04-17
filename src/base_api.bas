Attribute VB_Name = "base_api"
' -= Soft-15kHz - BASE - API
' -= © 2007-2009, Ariane 'SailorSat' Fugmann
Option Explicit


' -= API Types =-
Type VS_FIXEDFILEINFO
  dwSignature As Long
  dwStrucVersionl As Integer
  dwStrucVersionh As Integer
  dwFileVersionMSl As Integer
  dwFileVersionMSh As Integer
  dwFileVersionLSl As Integer
  dwFileVersionLSh As Integer
  dwProductVersionMSl As Integer
  dwProductVersionMSh As Integer
  dwProductVersionLSl As Integer
  dwProductVersionLSh As Integer
  dwFileFlagsMask As Long
  dwFileFlags As Long
  dwFileOS As Long
  dwFileType As Long
  dwFileSubtype As Long
  dwFileDateMS As Long
  dwFileDateLS As Long
End Type

Type OSVERSIONINFO
  dwOSVersionInfoSize As Long
  dwMajorVersion As Long
  dwMinorVersion As Long
  dwBuildNumber As Long
  dwPlatformId As Long
  szCSDVersion As String * 128
End Type

Type SHITEMID
  cb As Long
  abID As Byte
End Type

Type ITEMIDLIST
  mkid As SHITEMID
End Type


' -= API Includes =-
Declare Sub RtlMoveMemory Lib "kernel32.dll" (dest As Any, ByVal Source As Long, ByVal Length As Long)
Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
Declare Function GetVersionExA Lib "kernel32.dll" (lpVersionInformation As OSVERSIONINFO) As Integer

Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Declare Function SHGetPathFromIDListA Lib "shell32.dll" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function ShellExecuteA Lib "shell32.dll" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Function IsUserAnAdmin Lib "shell32.dll" () As Boolean

Declare Function RegOpenKeyA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKeyA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegSetValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Declare Function RegDeleteValueA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegDeleteKeyA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegQueryValueExA Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegEnumKeyA Lib "advapi32.dll" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Declare Function RegEnumValueA Lib "advapi32.dll" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long

Declare Function GetFileVersionInfoA Lib "version.dll" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, lpData As Any) As Long
Declare Function GetFileVersionInfoSizeA Lib "version.dll" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Declare Function VerQueryValueA Lib "version.dll" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long


' -= API Constants =-
Global Const HKEY_LOCAL_MACHINE = &H80000002
Global Const REG_SZ = 1
Global Const REG_BINARY = 3
Global Const REG_DWORD = 4
Global Const REG_MULTI_SZ = 7


' -= API Helper =-
Function GetSpecialFolder(ByVal CSIDL As Long) As String
  'CSIDL       Result
  '41          C:\WINNT\system32
  Dim sPath As String
  Dim IDL As ITEMIDLIST

  On Error Resume Next
  GetSpecialFolder = ""
  If SHGetSpecialFolderLocation(0, CSIDL, IDL) = 0 Then
    sPath = Space$(260)
    If SHGetPathFromIDListA(ByVal IDL.mkid.cb, ByVal sPath) Then
      GetSpecialFolder = Left$(sPath, InStr(sPath, vbNullChar) - 1) & ""
    End If
  End If
  If Err Then Err.Clear
  On Error GoTo 0
End Function

Function GetFileVersionInformation(ByVal FilePath As String) As String
  Dim BufferLength As Long
  Dim Buffer() As Byte
  Dim Pointer As Long
  'Dim PointerLength As Long
  Dim VInfo As VS_FIXEDFILEINFO
 
  BufferLength = GetFileVersionInfoSizeA(FilePath, 0&)
  If BufferLength < 1 Then Exit Function

  ReDim Buffer(BufferLength)
  GetFileVersionInfoA FilePath, 0&, BufferLength, Buffer(0)
  
  VerQueryValueA Buffer(0), "\", Pointer, BufferLength
  RtlMoveMemory VInfo, Pointer, Len(VInfo)

  GetFileVersionInformation = Format$(VInfo.dwFileVersionMSh) & "." & Format$(VInfo.dwFileVersionMSl) & "." & Format$(VInfo.dwFileVersionLSh) & "." & Format$(VInfo.dwFileVersionLSl)
End Function

Sub DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
  Dim hCurKey As Long
  Dim lRegResult As Long
  lRegResult = RegOpenKeyA(hKey, strPath, hCurKey)
  lRegResult = RegDeleteValueA(hCurKey, strValue)
  lRegResult = RegCloseKey(hCurKey)
End Sub

Sub DeleteKey(ByVal hKey As Long, ByVal strPath As String)
  Dim lRegResult As Long
  lRegResult = RegDeleteKeyA(hKey, strPath)
End Sub

Sub CreateKey(hKey As Long, strPath As String)
  Dim hCurKey As Long
  Dim lRegResult As Long
  lRegResult = RegCreateKeyA(hKey, strPath, hCurKey)
  If lRegResult <> 0 Then
    'Error? WTF!
  End If
  lRegResult = RegCloseKey(hKey)
End Sub

Function GetSettingLong(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, Optional Default As Long) As Long
  Dim lRegResult As Long
  Dim lValueType As Long
  Dim lBuffer As Long
  Dim lDataBufferSize As Long
  Dim hCurKey As Long
  If Not IsEmpty(Default) Then
    GetSettingLong = Default
  Else
    GetSettingLong = 0
  End If
  lRegResult = RegOpenKeyA(hKey, strPath, hCurKey)
  lDataBufferSize = 4
  lRegResult = RegQueryValueExA(hCurKey, strValue, 0&, lValueType, lBuffer, lDataBufferSize)
  If lRegResult = 0 Then
    If lValueType = REG_DWORD Then
      GetSettingLong = lBuffer
    End If
  Else
    'Error? WTF!
  End If
  lRegResult = RegCloseKey(hCurKey)
End Function

Sub SaveSettingLong(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String, ByVal lData As Long)
  Dim hCurKey As Long
  Dim lRegResult As Long
  lRegResult = RegCreateKeyA(hKey, strPath, hCurKey)
  lRegResult = RegSetValueExA(hCurKey, strValue, 0&, REG_DWORD, lData, 4)
  If lRegResult <> 0 Then
    'Error? WTF!
  End If
  lRegResult = RegCloseKey(hCurKey)
End Sub

Function GetSettingString(hKey As Long, strPath As String, strValue As String, Optional Default As String) As String
  Dim hCurKey As Long
  Dim lValueType As Long
  Dim strBuffer As String
  Dim lDataBufferSize As Long
  Dim intZeroPos As Integer
  Dim lRegResult As Long
  If Not IsEmpty(Default) Then
    GetSettingString = Default
  Else
    GetSettingString = ""
  End If
  lRegResult = RegOpenKeyA(hKey, strPath, hCurKey)
  lRegResult = RegQueryValueExA(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)
  If lRegResult = 0 Then
    If lValueType = REG_SZ Then
      strBuffer = String(lDataBufferSize, " ")
      lRegResult = RegQueryValueExA(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)
      GetSettingString = strBuffer
    ElseIf lValueType = REG_MULTI_SZ Then
      strBuffer = String(lDataBufferSize, " ")
      lRegResult = RegQueryValueExA(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)
      GetSettingString = strBuffer
    ElseIf lValueType = REG_BINARY Then
      strBuffer = String(lDataBufferSize, " ")
      lRegResult = RegQueryValueExA(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)
      GetSettingString = strBuffer
    End If
  Else
    'Error? WTF!
  End If
  lRegResult = RegCloseKey(hCurKey)
End Function

Sub SaveSettingBinary(hKey As Long, strPath As String, strValue As String, strData As String)
  Dim hCurKey As Long
  Dim lRegResult As Long
  
  lRegResult = RegCreateKeyA(hKey, strPath, hCurKey)
  lRegResult = RegSetValueExA(hCurKey, strValue, 0, REG_BINARY, ByVal strData, Len(strData))
  If lRegResult <> 0 Then
    'Error? WTF!
  End If
  lRegResult = RegCloseKey(hCurKey)
End Sub

Sub SaveSettingString(hKey As Long, strPath As String, strValue As String, strData As String)
  Dim hCurKey As Long
  Dim lRegResult As Long
  lRegResult = RegCreateKeyA(hKey, strPath, hCurKey)
  lRegResult = RegSetValueExA(hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData))
  If lRegResult <> 0 Then
    'Error? WTF!
  End If
  lRegResult = RegCloseKey(hCurKey)
End Sub

Sub SaveSettingMultiString(hKey As Long, strPath As String, strValue As String, strData As String)
  Dim hCurKey As Long
  Dim lRegResult As Long
  lRegResult = RegCreateKeyA(hKey, strPath, hCurKey)
  lRegResult = RegSetValueExA(hCurKey, strValue, 0, REG_MULTI_SZ, ByVal strData, Len(strData))
  If lRegResult <> 0 Then
    'Error? WTF!
  End If
  lRegResult = RegCloseKey(hCurKey)
End Sub

Function GetAllValues(hKey As Long, strPath As String) As Variant
  Dim lRegResult As Long
  Dim hCurKey As Long
  Dim lValueNameSize As Long
  Dim strValueName As String
  Dim lCounter As Long
  Dim byDataBuffer(4000) As Byte
  Dim lDataBufferSize As Long
  Dim lValueType As Long
  Dim strNames() As String
  Dim intZeroPos As Integer
  lRegResult = RegOpenKeyA(hKey, strPath, hCurKey)
  lCounter = 0
  Do
    lValueNameSize = 255
    strValueName = String$(lValueNameSize, " ")
    lDataBufferSize = 4000
    lRegResult = RegEnumValueA(hCurKey, lCounter, strValueName, lValueNameSize, 0&, lValueType, byDataBuffer(0), lDataBufferSize)
    If lRegResult = 0 Then
      ReDim Preserve strNames(lCounter) As String
      intZeroPos = InStr(strValueName, Chr$(0))
      If intZeroPos > 0 Then
        strNames(UBound(strNames)) = Left$(strValueName, intZeroPos - 1)
      Else
        strNames(UBound(strNames)) = strValueName
      End If
      lCounter = lCounter + 1
    Else
      Exit Do
    End If
  Loop
  GetAllValues = strNames
End Function

Function GetAllKeys(hKey As Long, strPath As String) As Variant
  Dim lRegResult As Long
  Dim lCounter As Long
  Dim hCurKey As Long
  Dim strBuffer As String
  Dim lDataBufferSize As Long
  Dim strNames() As String
  Dim intZeroPos As Integer
  lCounter = 0
  lRegResult = RegOpenKeyA(hKey, strPath, hCurKey)
  Do
    lDataBufferSize = 255
    strBuffer = String(lDataBufferSize, " ")
    lRegResult = RegEnumKeyA(hCurKey, lCounter, strBuffer, lDataBufferSize)
    If lRegResult = 0 Then
      ReDim Preserve strNames(lCounter) As String
      intZeroPos = InStr(strBuffer, Chr$(0))
      If intZeroPos > 0 Then
        strNames(UBound(strNames)) = Left$(strBuffer, intZeroPos - 1)
      Else
        strNames(UBound(strNames)) = strBuffer
      End If
      lCounter = lCounter + 1
    Else
      Exit Do
    End If
  Loop
  If lCounter = 0 Then
    ReDim strNames(0)
    strNames(0) = ""
  End If
  GetAllKeys = strNames
End Function
