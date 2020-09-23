Attribute VB_Name = "mHash"

Option Explicit
'hash table module commands
'makes using a hashtable simpler
'
'©2006 - 2007 Jason James Newland
'
'usage: always HMAKE <tablename> first
'       then you can use HADD, HDEL, etc
Private NewHash() As cHashTable
Private HashMax As Long

' "hash" table functions (similar to mIRC)
' AIM: to make the organisation and sorting of data
'      easier
Public Sub hMake(ByVal sHashName As String)
    'hMake <table>
    On Error Resume Next
    Dim i As Long
    Dim HashCount As Long
    '
    'redimension the Arrays
    If HashMax = 0 Then
        HashMax = 5
        ReDim Preserve NewHash(0 To HashMax) As cHashTable
    End If
    '
    'check that the hash doesn't already exist
    'and also see if we need to increase the array size
    'saves two iterating loops
    For i = LBound(NewHash) To UBound(NewHash)
        If Not NewHash(i) Is Nothing Then
            If LCase(NewHash(i).Name) = LCase(sHashName) Then
                Exit Sub
            Else
                HashCount = HashCount + 1
            End If
        End If
    Next i
    '
    'increase hash limit if needed
    If HashCount + 1 > HashMax Then
        HashMax = HashMax + 3
        Debug.Print "Redimensioning hash tables"
        ReDim Preserve NewHash(0 To HashMax) As cHashTable
    End If
    '
    'search for an empty hash index to use
    For i = LBound(NewHash) To UBound(NewHash)
        If NewHash(i) Is Nothing Then
            Debug.Print "new hash " & i
            Set NewHash(i) = New cHashTable
            NewHash(i).Name = sHashName
            Exit For
        End If
    Next i
End Sub

Public Sub hFree(ByVal sHashName As String)
    'hFree <table>
    On Error Resume Next
    Dim i As Long
    For i = LBound(NewHash) To UBound(NewHash)
        If Not NewHash(i) Is Nothing Then
            If LCase(NewHash(i).Name) = LCase(sHashName) Then
                'clear hash table
                Debug.Print "Hash removed " & i
                Set NewHash(i) = Nothing
                Exit Sub
            End If
        End If
    Next i
End Sub

Public Sub hAdd(ByVal sHashName As String, sKey As String, sData As String)
    'hAdd <table>, <key>, <data>
    On Error Resume Next
    Dim i As Long
    i = hFind(sHashName)
    'now check that a key doesn't already exist then
    'add it
    If LenB(sKey) <> 0 Then
        If LenB(sData) <> 0 Then
            If LenB(hGet(sHashName, sKey)) <> 0 Then
                NewHash(i).Remove LCase(sKey)
            End If
            NewHash(i).Add LCase(sKey), sData
        End If
    End If
End Sub

Public Sub hDel(ByVal sHashName As String, sKey As String)
    'hDel <table>, <key>
    On Error Resume Next
    Dim i As Long
    i = hFind(sHashName)
    'check the key exists and remove it, else do nothing
    If i <> "-1" Then
        If LenB(hGet(sHashName, sKey)) <> 0 Then
            NewHash(i).Remove LCase(sKey)
        End If
    End If
End Sub

Public Function hFind(ByVal sHashName As String) As Long
    'hFind(<table>)
    'returns the table number index
    On Error Resume Next
    Dim i As Integer
    Dim hIndex As Long
    hIndex = -1
    'find the table index with the hashname matching sHashName
    For i = LBound(NewHash) To UBound(NewHash)
        If Not NewHash(i) Is Nothing Then
            If LCase(NewHash(i).Name) = LCase(sHashName) Then
                hIndex = i
                Exit For
            End If
        End If
    Next i
    hFind = hIndex
End Function

Public Function hGet(ByVal sHashName As String, sKey As String) As String
    'hGet(<table>, <key>)
    'returns the data of the given table and key value
    On Error Resume Next
    Dim i As Long
    For i = LBound(NewHash) To UBound(NewHash)
        If Not NewHash(i) Is Nothing Then
            If LCase(NewHash(i).Name) = LCase(sHashName) Then
                If NewHash(i).Exists(LCase(sKey)) = True Then
                    hGet = NewHash(i).Item(LCase(sKey))
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

'ini file subs
Public Sub LoadINIToHash(ByVal sHashName As String, sININame As String, sINISection As String)
    'uses the application path as its base reference
    'but can be easily modified to suit individual needs
    On Error GoTo continue
    'does the ini file actually exist?
    If FileOrDirExists(App.Path & "\" & sININame) = False Then Exit Sub
    '
    Dim strTemp As String
    Dim FNum As Integer
    Dim intTmp As Long
    '
    'first clear the hash
    hFree sHashName
    '
    'remake it
    hMake sHashName
    '
    FNum = FreeFile
    'open the inifile
    Open App.Path & "\" & sININame For Input As #FNum
       While Not EOF(FNum)
            Line Input #FNum, strTemp
            If LCase(strTemp) = "[" & LCase(sINISection) & "]" Then
                While Not EOF(FNum)
                    Line Input #FNum, strTemp
                    'if we are are not reaching the start of a new
                    'section or the end of file, add the data
                    If Left$(strTemp, 1) <> "[" And LenB(strTemp) <> 0 Then
                        intTmp = InStr(strTemp, "=")
                        If intTmp <> 0 Then
                            hAdd sHashName, Left$(strTemp, intTmp - 1), Mid$(strTemp, intTmp + 1)
                        End If
                    Else
                        GoTo continue
                    End If
                Wend
            End If
        Wend
continue:
    Close #FNum
End Sub

Public Sub SaveHashToINI(ByVal sHashName As String, sINIFile As String, sINISection As String)
    On Error Resume Next
    Dim Hash As Long
    Dim i As Long
    Dim key As Long
    Dim Keys() As String
    '
    Hash = hFind(sHashName)
    '
    If Hash <> -1 Then
        DelINISection sINIFile, sINISection
        Keys = NewHash(Hash).Keys
        For i = LBound(Keys) To UBound(Keys)
            WriteINI sINIFile, sINISection, Keys(i), NewHash(Hash).Item(Keys(i))
        Next i
    End If
End Sub
