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
'''<classSummary> Dictionaryのアッパークラス、拡張したDictionary
'''<updateDate> 2021/02/19
'''<version> Ver.100

'AddIfのキーが重複した場合の処理のモード設定
Enum DuplicateMode
    dmNothing = 1000                        'Addせずになにもしない
    dmOverwrite = 1001                      '要素を上書きする
    dmSetDuplicateCollection = 1100         'Addせずに重複コレクションに追加する
    dmSetBothDuplicateCollection = 1101     'Addせずに重複コレクションに追加する。既にセットした重複された側も追加する
End Enum

Private Dic_ As Object
Private DuplicateCollection_ As Collection

'''<summary> 自分を返すためのもの
'''<param> なし
'''<returns> 自己参照
Public Property Get self() As Dic
    Set self = Me
End Property

'''<summary> オブジェクトをセット
'''<param> なし
'''<returns> なし
Private Sub Class_Initialize()
    Set Dic_ = CreateObject("Scripting.Dictionary")
    Set DuplicateCollection_ = New Collection
    Me.SetCompareMode (vbTextCompare)
End Sub

'''<summary> 新しいキー/アイテムのペアを追加
'''<param> キー、要素
'''<returns> なし
Public Function Add(ByVal Key As String, ByVal Item As Variant)
    Dic_.Add Key, Item
End Function

'''<summary> 新しいキー/アイテムのペアを追加、但し、すでに存在している場合、設定した処理を行う
'''<param> キー、要素、設定
'''<returns> なし
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

'''<summary> 比較モードを設定します
'''<param> VbCompareMethodの列挙型
'''<returns> なし
'''<remarks> コンストラクタ時点ではvbTextCompare(テキスト比較)になっています。
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

'''<summary> 要素数を返す
'''<param> なし
'''<returns> 要素数
Public Property Get Count() As Long
    Count = Dic_.Count
End Property

'''<summary> 存在するかどうかを真偽値で返す
'''<param> なし
'''<returns> 真偽値
Public Function Exists(ByVal Key As String) As Boolean
    Exists = Dic_.Exists(Key)
End Function

'''<summary> 要素の内容を返す
'''<param> キー
'''<returns> 要素
'''<remarks> 元のdictionaryと異なり読み取り専用 、 既定メンバ
Public Function Item(ByVal Key As String) As Variant
Attribute Item.VB_UserMemId = 0
    If VBA.IsObject(Dic_.Item(Key)) Then
        Set Item = Dic_.Item(Key)
    Else
        Item = Dic_.Item(Key)
    End If
End Function

'''<summary> 要素の内容を全て返す
'''<param> なし
'''<returns> 要素の配列
Public Property Get Items() As Variant
    Items = Dic_.Items()
End Property

'''<summary> キーを新しく設定する
'''<param> 変更するキー、新しいキー
'''<returns> なし
Public Function Key(ByVal ChangeKey As String, ByVal NewKey As String)
    Dic_.Key(ChangeKey) = NewKey
End Function

'''<summary> キーを全て返す
'''<param> なし
'''<returns> キーの配列
Public Property Get Keys() As Variant
    Keys = Dic_.Keys()
End Property

'''<summary> 要素を削除する
'''<param> 削除するキー
'''<returns> なし
Public Function Remove(ByVal Key As String)
    Dic_.Remove (Key)
End Function

'''<summary> AddIfで重複したものを返す
'''<param> なし
'''<returns> 重複した要素を返す
Public Property Get DuplicateCollection() As Collection
    Set DuplicateCollection = DuplicateCollection_
End Property
