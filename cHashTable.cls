VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cHashTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_HelpID = 2184
Attribute VB_Description = "This class implements a hashtable, a structure that offers many of the features of a collectior or dictionary, and is often even faster than the built-in collection."

Option Explicit

' ----------------------------------------------
' HASHTABLE class module
'
' This class implements a hashtable, a structure that offers many
' of the features of a collectior or dictionary, and is often
' even faster than the built-in collection.
'
' NOTE: must make Item the default member, using the Tools | Procedure
' Attributes dialog
'
' Usage:
'   Dim ht As New HashTable
'   ht.SetSize 10000           ' initial number of slots (the higher,
'  the better)
'
'   ' enforce case-insensitive key search
'   ht.IgnoreCase = True
'   ' add values
'   ht.Add "key", value        ' add a value associated to a key
'   ' count how many values are in the table
'   Print ht.Count
'   ' read/write a value
'   Print ht("key")
'   ht("key") = newValue
'
'   ' remove a value
'   ht.Remove "key"
'   ' remove all values
'   ht.RemoveAll
'   ' check whether a value exists
'   If ht.Exists("key") Then ...
'
'   ' get the array of keys and values
'   Dim keys() As String, values() As Variant
'   keys() = ht.Keys
'   values() = ht.Values
'
'----------------------------------------------

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, source As Any, ByVal bytes As Long)

' default values
Const DEFAULT_HASHSIZE = 1024
Const DEFAULT_LISTSIZE = 2048
Const DEFAULT_CHUNKSIZE = 1024

Private Type SlotType
    key As String
    Value As Variant
    nextItem As Long      ' 0 if last item
End Type

' for each hash code this array holds the first element
' in slotTable() with the corresponding hash code
Private hashTbl() As Long
' the array that holds the data
Private slotTable() As SlotType

' pointer to first free slot
Private FreeNdx As Long

Private m_Hashname As String
' size of hash table
Private m_HashSize As Long
' size of slot table
Private m_ListSize As Long
' chunk size
Private m_ChunkSize As Long
' items in the slot table
Private m_Count As Long

' member variable for IgnoreCase property
Private m_IgnoreCase As Boolean

' True if keys are searched in case-unsensitive mode
' this can be assigned to only when the hash table is empty

Property Get IgnoreCase() As Boolean
Attribute IgnoreCase.VB_Description = "True if keys are searched in case-unsensitive modeTrue if keys are searched in case-unsensitive mode"
Attribute IgnoreCase.VB_HelpID = 2185
    IgnoreCase = m_IgnoreCase
End Property

Property Let IgnoreCase(ByVal newValue As Boolean)
    '##BLOCK_DESCRIPTION True if keys are searched in case-unsensitive mode
    ' this can be assigned to only when the hash table is empty
    If m_Count Then
        Err.Raise 1001, , "The Hash Table isn't empty"
    End If
    m_IgnoreCase = newValue
End Property

' initialize the hash table
Sub SetSize(ByVal HashSize As Long, Optional ByVal ListSize As Long, _
    Optional ByVal ChunkSize As Long)
Attribute SetSize.VB_Description = "Initialize the hash table"
Attribute SetSize.VB_HelpID = 2186
    '##PARAMETER_DESCRIPTION HashSize The fixed size of the hash table
    '##PARAMETER_DESCRIPTION ListSize The initial size of the list
    '##PARAMETER_DESCRIPTION ChunkSize The automatic grow block size of the list
    '##BLOCK_DESCRIPTION Initialize the hash table
    ' provide defaults
    If ListSize <= 0 Then ListSize = m_ListSize
    If ChunkSize <= 0 Then ChunkSize = m_ChunkSize
    ' save size values
    m_HashSize = HashSize
    m_ListSize = ListSize
    m_ChunkSize = ChunkSize
    m_Count = 0
    ' rebuild tables
    FreeNdx = 0
    ReDim hashTbl(0 To HashSize - 1) As Long
    ReDim slotTable(0) As SlotType
    ExpandSlotTable m_ListSize
End Sub

' check whether an item is in the hash table
Function Exists(ByVal key As String) As Boolean
Attribute Exists.VB_Description = "Check whether an item is in the hash table"
Attribute Exists.VB_HelpID = 2187
    '##BLOCK_DESCRIPTION Check whether an item is in the hash table
    Exists = GetSlotIndex(key) <> 0
End Function

' add a new element to the hash table
Sub Add(ByVal key As String, ByVal Value As Variant)
Attribute Add.VB_Description = "Add a new element to the hash table"
Attribute Add.VB_HelpID = 2188
    '##BLOCK_DESCRIPTION Add a new element to the hash table
    Dim ndx As Long, Create As Boolean
    ' get the index to the slot where the value is
    ' (allocate a new slot if necessary)
    Create = True
    
    ndx = GetSlotIndex(key, Create)
    
    If Create Then
        ' the item was actually added
        If IsObject(Value) Then
            Set slotTable(ndx).Value = Value
        Else
            slotTable(ndx).Value = Value
        End If
    Else
        ' raise error "This key is already associated with an item of this
        ' collection"
        Err.Raise 457
    End If
End Sub

' the value associated to a key
' (empty if not found)
Property Get Item(ByVal key As String) As Variant
Attribute Item.VB_Description = "The value associated to a key (empty if not found)    "
Attribute Item.VB_HelpID = 2189
    '##BLOCK_DESCRIPTION The value associated to a key _
    (empty if not found)
    Dim ndx As Long
    ' get the index to the slot where the value is
    ndx = GetSlotIndex(key)
    If ndx = 0 Then
        ' return Empty if not found
    ElseIf IsObject(slotTable(ndx).Value) Then
        Set Item = slotTable(ndx).Value
    Else
        Item = slotTable(ndx).Value
    End If
End Property

Property Let Item(ByVal key As String, ByVal Value As Variant)
    Dim ndx As Long
    ' get the index to the slot where the value is
    ' (allocate a new slot if necessary)
    ndx = GetSlotIndex(key, True)
    ' store the value
    slotTable(ndx).Value = Value
End Property

Property Set Item(ByVal key As String, ByVal Value As Object)
    Dim ndx As Long
    ' get the index to the slot where the value is
    ' (allocate a new slot if necessary)
    ndx = GetSlotIndex(key, True)
    ' store the value
    Set slotTable(ndx).Value = Value
End Property

' remove an item from the hash table
Sub Remove(ByVal key As String)
Attribute Remove.VB_Description = "Remove an item from the hash table"
Attribute Remove.VB_HelpID = 2190
    '##BLOCK_DESCRIPTION Remove an item from the hash table
    Dim ndx As Long, HCode As Long, LastNdx As Long
    ndx = GetSlotIndex(key, False, HCode, LastNdx)
    ' raise error if no such element
    If ndx = 0 Then Err.Raise 5
    
    If LastNdx Then
        ' this isn't the first item in the slotTable() array
        slotTable(LastNdx).nextItem = slotTable(ndx).nextItem
    ElseIf slotTable(ndx).nextItem Then
        ' this is the first item in the slotTable() array
        ' and is followed by one or more items
        hashTbl(HCode) = slotTable(ndx).nextItem
    Else
        ' this is the only item in the slotTable() array
        ' for this hash code
        hashTbl(HCode) = 0
    End If
    
    ' put the element back in the free list
    With slotTable(ndx)
        .Value = Empty
        .nextItem = FreeNdx
    End With
    
    FreeNdx = ndx
    ' we have deleted an item
    m_Count = m_Count - 1
    
End Sub

' remove all items from the hash table
Sub RemoveAll()
Attribute RemoveAll.VB_Description = "Remove all items from the hash table"
Attribute RemoveAll.VB_HelpID = 2191
    '##BLOCK_DESCRIPTION Remove all items from the hash table
    SetSize m_HashSize, m_ListSize, m_ChunkSize
End Sub

' the number of items in the hash table

Property Get Count() As Long
Attribute Count.VB_Description = "The number of items in the hash table"
Attribute Count.VB_HelpID = 2192
    '##BLOCK_DESCRIPTION The number of items in the hash table
    Count = m_Count
End Property

' the array of all keys
' (VB5 users: convert return type to Variant)
Property Get Keys() As String()
Attribute Keys.VB_Description = "The array of all keys"
Attribute Keys.VB_HelpID = 2193
    '##BLOCK_DESCRIPTION The array of all keys
    Dim i As Long, ndx As Long
    Dim n As Long
    
    If m_Count < 1 Then Exit Function
    
    ReDim res(0 To m_Count - 1) As String
    
    For i = 0 To m_HashSize - 1
        ' take the pointer from the hash table
        ndx = hashTbl(i)
        ' walk the slottable() array
        Do While ndx
            res(n) = slotTable(ndx).key
            n = n + 1
            ndx = slotTable(ndx).nextItem
        Loop
    Next
    ' assign to the result
    Keys = res()
End Property

' the array of all values
' (VB5 users: convert return type to Variant)

Property Get Values() As Variant()
Attribute Values.VB_Description = "The array of all values"
Attribute Values.VB_HelpID = 2194
    '##BLOCK_DESCRIPTION The array of all values
    Dim i As Long, ndx As Long
    Dim n As Long, res() As Variant
    
    If m_Count > 0 Then
        ReDim res(0 To m_Count - 1) As Variant
        
        For i = 0 To m_HashSize - 1
            ' take the pointer from the hash table
            ndx = hashTbl(i)
            ' walk the slottable() array
            Do While ndx
                If IsObject(slotTable(ndx).Value) Then
                    Set res(n) = slotTable(ndx).Value
                Else
                    res(n) = slotTable(ndx).Value
                End If
                n = n + 1
                ndx = slotTable(ndx).nextItem
            Loop
        Next
    End If
    ' assign to the result
    Values = res()
End Property

'-----------------------------------------
' Private procedures
'-----------------------------------------
Private Sub Class_Initialize()
    ' initialize the tables at default size
    SetSize DEFAULT_HASHSIZE, DEFAULT_LISTSIZE, DEFAULT_CHUNKSIZE
End Sub

' expand the slotTable() array
Private Sub ExpandSlotTable(ByVal numEls As Long)
    Dim newFreeNdx As Long, i As Long
    newFreeNdx = UBound(slotTable) + 1
    ReDim Preserve slotTable(0 To UBound(slotTable) + numEls) As SlotType
    ' create the linked list of free items
    For i = newFreeNdx To UBound(slotTable)
        slotTable(i).nextItem = i + 1
    Next
    ' overwrite the last (wrong) value
    slotTable(UBound(slotTable)).nextItem = FreeNdx
    ' we now know where to pick the first free item
    FreeNdx = newFreeNdx
End Sub

' return the hash code of a string
Private Function HashCode(key As String) As Long
    Dim lastEl As Long, i As Long
    ' copy ansi codes into an array of long
    lastEl = (Len(key) - 1) \ 4
    ReDim codes(lastEl) As Long
    ' this also converts from Unicode to ANSI
    CopyMemory codes(0), ByVal key, Len(key)
    
    ' XOR the ANSI codes of all characters
    For i = 0 To lastEl
        HashCode = HashCode Xor codes(i)
    Next
    
End Function

' get the index where an item is stored or 0 if not found
' if Create = True the item is created
'
' on exit Create=True only if a slot has been actually created
Private Function GetSlotIndex(ByVal key As String, Optional Create As Boolean, _
    Optional HCode As Long, Optional LastNdx As Long) As Long
    Dim ndx As Long
    On Error Resume Next
    ' raise error if invalid key
    If Len(key) = 0 Then Err.Raise 1001, , "Invalid key"
    ' keep case-unsensitiveness into account
    If m_IgnoreCase Then key = UCase$(key)
    ' get the index in the hashTbl() array
    HCode = HashCode(key) Mod m_HashSize
    ' get the pointer to the slotTable() array
    ndx = hashTbl(HCode)
    
    ' exit if there is no item with that hash code
    Do While ndx
        ' compare key with actual value
        If slotTable(ndx).key = key Then Exit Do
        ' remember last pointer
        LastNdx = ndx
        ' check the next item
        ndx = slotTable(ndx).nextItem
    Loop
    
    ' create a new item if not there
    If ndx = 0 And Create Then
        ndx = GetFreeSlot()
        PrepareSlot ndx, key, HCode, LastNdx
    Else
        ' signal that no item has been created
        Create = False
    End If
    ' this is the return value
    GetSlotIndex = ndx

End Function

' return the first free slot
Private Function GetFreeSlot() As Long
    ' allocate new memory if necessary
    If FreeNdx = 0 Then ExpandSlotTable m_ChunkSize
    ' use the first slot
    GetFreeSlot = FreeNdx
    ' update the pointer to the first slot
    FreeNdx = slotTable(GetFreeSlot).nextItem
    ' signal this as the end of the linked list
    slotTable(GetFreeSlot).nextItem = 0
    ' we have one more item
    m_Count = m_Count + 1
End Function

' assign a key and value to a given slot
Private Sub PrepareSlot(ByVal Index As Long, ByVal key As String, _
    ByVal HCode As Long, ByVal LastNdx As Long)
    ' assign the key
    ' keep case-sensitiveness into account
    If m_IgnoreCase Then key = UCase$(key)
    slotTable(Index).key = key
    
    If LastNdx Then
        ' this is the successor of another slot
        slotTable(LastNdx).nextItem = Index
    Else
        ' this is the first slot for a given hash code
        hashTbl(HCode) = Index
    End If
End Sub

Public Property Let Name(ByVal sName As String)
    m_Hashname = sName
End Property

Public Property Get Name() As String
    Name = m_Hashname
End Property
