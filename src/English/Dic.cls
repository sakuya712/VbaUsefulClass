VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'''<classSummary> Dictionary upper class, extended Dictionary
'''<updateDate> 2021/02/19
'''<version> Ver.100

'Mode setting of processing when the key of AddIf is duplicated
Enum DuplicateMode
    dmNothing = 1000                        'Do nothing without adding
    dmOverwrite = 1001                      'Overwrite the element
    dmSetDuplicateCollection = 1100         'Add to duplicate collection without adding
    dmSetBothDuplicateCollection = 1101     'Add to duplicate collection without adding. Add the duplicated side that has already been set
End Enum

Private Dic_ As Object
Private DuplicateCollection_ As Collection

'''<summary> To return myself
'''<param> None
'''<returns> Self-reference
Public Property Get self() As Dic
    Set self = Me
End Property

'''<summary> Set object 
'''<param> None
'''<returns> None
Private Sub Class_Initialize()
    Set Dic_ = CreateObject("Scripting.Dictionary")
    Set DuplicateCollection_ = New Collection
    Me.SetCompareMode (vbTextCompare)
End Sub

'''<summary> Add new key / item pair 
'''<param> Key, element
'''<returns> None
Public Function Add(ByVal Key As String, ByVal Item As Variant)
    Dic_.Add Key, Item
End Function

'''<summary> Add new key / item pair, However, if it already exists, the set process is performed. 
'''<param> Key, element, settings 
'''<returns> None
Public Function AddIf(ByVal Key As String, ByVal Item As Variant, Optional ByVal Mode As DuplicateMode = dmNothing)
    If Me.Exists(Key) = False Then
        Me.Add Key, Item
        Exit Function
    End If
    Select Case Mode
        Case DuplicateMode.dmNothing
            Exit Function
        Case DuplicateMode.dmOverwrite
            Dic_.Item(Key) = Item
            Exit Function
        Case DuplicateMode.dmSetDuplicateCollection
            DuplicateCollection_.Add Item
            Exit Function
        Case DuplicateMode.dmSetBothDuplicateCollection
            DuplicateCollection_.Add Me.Item(Key)
            DuplicateCollection_.Add Item
            Exit Function
    End Select
End Function

'''<summary> Set the comparison mode 
'''<param> VbCompareMethod enum
'''<returns> None
'''<remarks> At the time of the constructor, it is vbTextCompare. 
Public Function SetCompareMode(ByVal Mode As VbCompareMethod)
    Select Case Mode
        Case VbCompareMethod.vbBinaryCompare
            Dic_.CompareMode = 0
        Case VbCompareMethod.vbTextCompare
            Dic_.CompareMode = 1
        Case VbCompareMethod.vbDatabaseCompare
            Dic_.CompareMode = 2
    End Select
End Function

'''<summary> Returns the number of elements 
'''<param> None
'''<returns> the number of elements 
Public Property Get Count() As Long
    Count = Dic_.Count
End Property

'''<summary> Returns whether it exists as a boolean value 
'''<param> None
'''<returns> boolean value
Public Function Exists(ByVal Key As String) As Boolean
    Exists = Dic_.Exists(Key)
End Function

'''<summary> Returns the contents of the element 
'''<param> Keys
'''<returns> element
'''<remarks> Read-only, default member, unlike the original dictionary 
Public Function Item(ByVal Key As String) As Variant
Attribute Item.VB_UserMemId = 0
    If VBA.IsObject(Dic_.Item(Key)) Then
        Set Item = Dic_.Item(Key)
    Else
        Item = Dic_.Item(Key)
    End If
End Function

'''<summary> Returns the entire contents of the element 
'''<param> None
'''<returns> Array of elements 
Public Property Get Items() As Variant
    Items = Dic_.Items()
End Property

'''<summary> Set a new key 
'''<param> Key to change, new key 
'''<returns> None
Public Function Key(ByVal ChangeKey As String, ByVal NewKey As String)
    Dic_.Key(ChangeKey) = NewKey
End Function

'''<summary> Returns all keys 
'''<param> None
'''<returns> Array of keys 
Public Property Get Keys() As Variant
    Keys = Dic_.Keys()
End Property

'''<summary> Delete element
'''<param> Key to delete
'''<returns> None
Public Function Remove(ByVal Key As String)
    Dic_.Remove (Key)
End Function

'''<summary> AddIf returns duplicates
'''<param> None
'''<returns> Returns duplicate elements
Public Property Get DuplicateCollection() As Collection
    Set DuplicateCollection = DuplicateCollection_
End Property
