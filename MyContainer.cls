VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MyContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pDict As Scripting.Dictionary ' Strongly typed
Private pKeyOrder As Collection ' Preserves insertion/sort order

Private Sub Class_Initialize()
    Set pDict = New Scripting.Dictionary
    Set pKeyOrder = New Collection
End Sub

' Add item with key
Public Sub Add(Item As Variant, key As String)
    If pDict.Exists(key) Then
        Err.Raise vbObjectError + 1, , "Key already exists."
    End If
    pDict.Add key, Item
    pKeyOrder.Add key
End Sub

' Check key existence
Public Function Exists(key As Variant) As Boolean
    Exists = pDict.Exists(key)
End Function

' Retrieve value type item by key
Public Property Get Item(key As Variant) As Variant
    If Not pDict.Exists(key) Then
        Err.Raise vbObjectError + 2, , "Key not found."
    End If
    If IsObject(pDict(key)) Then
        Set Item = pDict(key)
    Else
        Item = pDict(key)
    End If
End Property

' Value by index (1-based)
Public Property Get ItemByIndex(Index As Long) As Variant
    If Index < 1 Or Index > pKeyOrder.Count Then
        Err.Raise vbObjectError + 6, , "Index out of bounds."
    End If
    If IsObject(pDict(pKeyOrder(Index))) Then
        Set ItemByIndex = pDict(pKeyOrder(Index))
    Else
        ItemByIndex = pDict(pKeyOrder(Index))
    End If
End Property

' Sort keys alphabetically and reorder key collection
Public Sub OrderByKey()
    Dim keys() As String, i As Long
    ReDim keys(0 To pKeyOrder.Count - 1)
    For i = 1 To pKeyOrder.Count
        keys(i - 1) = pKeyOrder(i)
    Next i
    QuickSortArray keys, LBound(keys), UBound(keys)

    Set pKeyOrder = New Collection
    For i = LBound(keys) To UBound(keys)
        pKeyOrder.Add keys(i)
    Next i
End Sub

' Remove by key (returns value or object)
Public Function Remove(key As Variant) As Variant
    If Not pDict.Exists(key) Then
        Err.Raise vbObjectError + 10, , "Key not found."
    End If
    If IsObject(pDict(key)) Then
        Set Remove = pDict(key)
    Else
        Remove = pDict(key)
    End If
    pDict.Remove key

    Dim i As Long
    For i = 1 To pKeyOrder.Count
        If pKeyOrder(i) = key Then
            pKeyOrder.Remove i
            Exit For
        End If
    Next i
End Function

' Remove by index (returns value or object)
Public Function RemoveByIndex(Index As Long) As Variant
    If Index < 1 Or Index > pKeyOrder.Count Then
        Err.Raise vbObjectError + 11, , "Index out of bounds."
    End If
    Dim key As String
    key = pKeyOrder(Index)
    If IsObject(pDict(key)) Then
        Set RemoveByIndex = pDict(key)
    Else
        RemoveByIndex = pDict(key)
    End If
    pDict.Remove key
    pKeyOrder.Remove Index
    Exit Function
End Function

' Count of items
Public Property Get Count() As Long
    Count = pDict.Count
End Property

' Clear all items
Public Sub Clear()
    Set pDict = New Scripting.Dictionary
    Set pKeyOrder = New Collection
End Sub

' Return array of keys in current order
Public Function keys() As Variant
    Dim arr() As Variant, i As Long
    ReDim arr(0 To pKeyOrder.Count - 1)
    For i = 1 To pKeyOrder.Count
        arr(i - 1) = pKeyOrder(i)
    Next i
    keys = arr
End Function

' Return array of values in current order
Public Function Values() As Variant
    Dim arr() As Variant, i As Long
    ReDim arr(0 To pKeyOrder.Count - 1)
    For i = 1 To pKeyOrder.Count
        If IsObject(pDict(pKeyOrder(i))) Then
            Set arr(i - 1) = pDict(pKeyOrder(i))
        Else
            arr(i - 1) = pDict(pKeyOrder(i))
        End If
    Next i
    Values = arr
    Debug.Print "Passei"

End Function

' Return 2D array: Key | Value
Public Function ToArray() As Variant
    Dim i As Long
    Dim result() As Variant
    ReDim result(1 To pKeyOrder.Count, 1 To 2)
    
    For i = 1 To pKeyOrder.Count
        result(i, 1) = pKeyOrder(i)
        If IsObject(pDict(pKeyOrder(i))) Then
            Set result(i, 2) = pDict(pKeyOrder(i))
        Else
            result(i, 2) = pDict(pKeyOrder(i))
        End If
    Next i
    ToArray = result
End Function

' Return the key at a given 1-based index
Public Function KeyByIndex(Index As Long) As String
    If Index < 1 Or Index > pKeyOrder.Count Then
        Err.Raise vbObjectError + 12, , "Index out of bounds."
    End If
    KeyByIndex = pKeyOrder(Index)
End Function

