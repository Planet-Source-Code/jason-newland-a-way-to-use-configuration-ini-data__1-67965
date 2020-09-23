VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cINIFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long

Private Const BufferSize As Long = 4096

Private strIniFile As String

Public Property Get File() As String
    File = strIniFile
End Property

Public Property Let File(Value As String)
    strIniFile = Value
End Property

Public Function GetValue(strSection As String, strKey As String) As Variant
    Dim strBuffer As String
    Dim lLength As Long

    strBuffer = Space(BufferSize)

    lLength = GetPrivateProfileString(strSection, strKey, vbNullString, strBuffer, BufferSize, strIniFile)
    
    GetValue = Left(strBuffer, lLength)
End Function

Public Sub WriteValue(strSection As String, strKey As String, vntValue As Variant)
    WritePrivateProfileString strSection, strKey, CStr(vntValue), strIniFile
End Sub

Public Function GetSection(strSection As String) As Variant
    Dim strBuffer As String
    Dim lLength As Long

    strBuffer = Space(BufferSize)

    lLength = GetPrivateProfileSection(strSection, strBuffer, BufferSize, strIniFile)
    
    GetSection = Split(Left(strBuffer, lLength), vbNullChar)
End Function

Public Function GetSectionKeys(strSection As String) As Variant
    Dim strBuffer As String
    Dim lLength As Long

    strBuffer = Space(BufferSize)

    lLength = GetPrivateProfileString(strSection, vbNullString, vbNullString, strBuffer, BufferSize, strIniFile)
    
    GetSectionKeys = Split(Left(strBuffer, lLength), vbNullChar)
End Function
