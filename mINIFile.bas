Attribute VB_Name = "mINIFile"

Option Explicit

Private Const MAX_PATH = 260
Private Const ERROR_NO_MORE_FILES = 18
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Public Function ReadINI(iniFile As String, iniSection As String, iniItem As String)
    On Error Resume Next
    'snippet and class by Robert Rowe
    Dim ini As cINIFile
    Dim strTemp As String
    Set ini = New cINIFile
    With ini
        .File = App.Path & "\" & iniFile
        strTemp = .GetValue(iniSection, iniItem)
        If Left$(strTemp, 2) = ":" & Chr$(2) Then strTemp = Mid$(strTemp, 2)
        If Left$(strTemp, 2) = ":" & Chr$(3) Then strTemp = Mid$(strTemp, 2)
        If Left$(strTemp, 3) = ":" & Chr$(31) Then strTemp = Mid$(strTemp, 2)
        ReadINI = strTemp
    End With
    Set ini = Nothing
End Function

Public Sub WriteINI(iniFile As String, iniSection As String, iniItem As String, iniData As String)
    On Error Resume Next
    'snippet and dll by Robert Rowe
    Dim ini As cINIFile
    Dim strTemp As String
    Set ini = New cINIFile
    strTemp = iniData
    If Left$(strTemp, 1) = Chr$(2) Then strTemp = ":" & strTemp
    If Left$(strTemp, 1) = Chr$(3) Then strTemp = ":" & strTemp
    If Left$(strTemp, 1) = Chr$(31) Then strTemp = ":" & strTemp
    With ini
        .File = App.Path & "\" & iniFile
        .WriteValue iniSection, iniItem, strTemp
    End With
    Set ini = Nothing
End Sub

Public Sub DelINISection(iniFile As String, iniSection As String)
    On Error Resume Next
    Dim ini As cINIFile
    Set ini = New cINIFile
    With ini
        .File = App.Path & "\" & iniFile
        .WriteValue iniSection, vbNullString, vbNullString
    End With
    Set ini = Nothing
End Sub

Function FileOrDirExists(Optional ByVal sFile As String = vbNullString, Optional ByVal sFolder As String = vbNullString) As Boolean
    On Error Resume Next
    Dim lpFindFileData As WIN32_FIND_DATA
    Dim lFileHdl  As Long
    Dim sTemp As String
    Dim sTemp2 As String
    Dim lRet As Long
    Dim iLastIndex  As Integer
    Dim strPath As String
    Dim sStartDir As String
    
    On Error Resume Next
    '// both params are empty
    If LenB(sFile) = 0 And LenB(sFolder) = 0 Then Exit Function
    '// both are full, empty folder param
    If LenB(sFile) <> 0 And LenB(sFolder) <> 0 Then sFolder = vbNullString
    If LenB(sFolder) <> 0 Then
        '// set start directory
        sStartDir = sFolder
    Else
        '// extract start directory from file path
        sStartDir = Left$(sFile, InStrRev(sFile, "\"))
        '// just get filename
        sFile = Right$(sFile, Len(sFile) - InStrRev(sFile, "\"))
    End If
    '// add trailing \ to start directory if required
    If Right$(sStartDir, 1) <> "\" Then sStartDir = sStartDir & "\"
    
    sStartDir = sStartDir & "*.*"
    
    '// get a file handle
    lFileHdl = FindFirstFile(sStartDir, lpFindFileData)
    
    If lFileHdl <> -1 Then
        If LenB(sFolder) <> 0 Then
            '// folder exists
            FileOrDirExists = True
        Else
            Do Until lRet = ERROR_NO_MORE_FILES
                strPath = Left$(sStartDir, Len(sStartDir) - 4) & "\"
                '// if it is a file
                If (lpFindFileData.dwFileAttributes And FILE_ATTRIBUTE_NORMAL) = vbNormal Then
                    sTemp = StrConv(StripTerminator(lpFindFileData.cFileName), vbProperCase)
                    '// remove LCase$ if you want the search to be case sensitive (unlikely!)
                    If LCase$(sTemp) = LCase$(sFile) Then
                        FileOrDirExists = True '// file found
                        Exit Do
                    End If
                End If
                '// based on the file handle iterate through all files and dirs
                lRet = FindNextFile(lFileHdl, lpFindFileData)
                If lRet = 0 Then Exit Do
            Loop
        End If
    End If
    '// close the file handle
    lRet = FindClose(lFileHdl)
End Function

Function StripTerminator(ByVal strString As String) As String
    On Error Resume Next
    Dim intZeroPos As Integer
    intZeroPos = InStr(strString, vbNullChar)
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function

