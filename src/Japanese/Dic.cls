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
'''<classSummary> Dictionary�̃A�b�p�[�N���X�A�g������Dictionary
'''<updateDate> 2021/02/19
'''<version> Ver.100

'AddIf�̃L�[���d�������ꍇ�̏����̃��[�h�ݒ�
Enum DuplicateMode
    dmNothing = 1000                        'Add�����ɂȂɂ����Ȃ�
    dmOverwrite = 1001                      '�v�f���㏑������
    dmSetDuplicateCollection = 1100         'Add�����ɏd���R���N�V�����ɒǉ�����
    dmSetBothDuplicateCollection = 1101     'Add�����ɏd���R���N�V�����ɒǉ�����B���ɃZ�b�g�����d�����ꂽ�����ǉ�����
End Enum

Private Dic_ As Object
Private DuplicateCollection_ As Collection

'''<summary> ������Ԃ����߂̂���
'''<param> �Ȃ�
'''<returns> ���ȎQ��
Public Property Get self() As Dic
    Set self = Me
End Property

'''<summary> �I�u�W�F�N�g���Z�b�g
'''<param> �Ȃ�
'''<returns> �Ȃ�
Private Sub Class_Initialize()
    Set Dic_ = CreateObject("Scripting.Dictionary")
    Set DuplicateCollection_ = New Collection
    Me.SetCompareMode (vbTextCompare)
End Sub

'''<summary> �V�����L�[/�A�C�e���̃y�A��ǉ�
'''<param> �L�[�A�v�f
'''<returns> �Ȃ�
Public Function Add(ByVal Key As String, ByVal Item As Variant)
    Dic_.Add Key, Item
End Function

'''<summary> �V�����L�[/�A�C�e���̃y�A��ǉ��A�A���A���łɑ��݂��Ă���ꍇ�A�ݒ肵���������s��
'''<param> �L�[�A�v�f�A�ݒ�
'''<returns> �Ȃ�
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

'''<summary> ��r���[�h��ݒ肵�܂�
'''<param> VbCompareMethod�̗񋓌^
'''<returns> �Ȃ�
'''<remarks> �R���X�g���N�^���_�ł�vbTextCompare(�e�L�X�g��r)�ɂȂ��Ă��܂��B
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

'''<summary> �v�f����Ԃ�
'''<param> �Ȃ�
'''<returns> �v�f��
Public Property Get Count() As Long
    Count = Dic_.Count
End Property

'''<summary> ���݂��邩�ǂ�����^�U�l�ŕԂ�
'''<param> �Ȃ�
'''<returns> �^�U�l
Public Function Exists(ByVal Key As String) As Boolean
    Exists = Dic_.Exists(Key)
End Function

'''<summary> �v�f�̓��e��Ԃ�
'''<param> �L�[
'''<returns> �v�f
'''<remarks> ����dictionary�ƈقȂ�ǂݎ���p �A ���胁���o
Public Function Item(ByVal Key As String) As Variant
Attribute Item.VB_UserMemId = 0
    If VBA.IsObject(Dic_.Item(Key)) Then
        Set Item = Dic_.Item(Key)
    Else
        Item = Dic_.Item(Key)
    End If
End Function

'''<summary> �v�f�̓��e��S�ĕԂ�
'''<param> �Ȃ�
'''<returns> �v�f�̔z��
Public Property Get Items() As Variant
    Items = Dic_.Items()
End Property

'''<summary> �L�[��V�����ݒ肷��
'''<param> �ύX����L�[�A�V�����L�[
'''<returns> �Ȃ�
Public Function Key(ByVal ChangeKey As String, ByVal NewKey As String)
    Dic_.Key(ChangeKey) = NewKey
End Function

'''<summary> �L�[��S�ĕԂ�
'''<param> �Ȃ�
'''<returns> �L�[�̔z��
Public Property Get Keys() As Variant
    Keys = Dic_.Keys()
End Property

'''<summary> �v�f���폜����
'''<param> �폜����L�[
'''<returns> �Ȃ�
Public Function Remove(ByVal Key As String)
    Dic_.Remove (Key)
End Function

'''<summary> AddIf�ŏd���������̂�Ԃ�
'''<param> �Ȃ�
'''<returns> �d�������v�f��Ԃ�
Public Property Get DuplicateCollection() As Collection
    Set DuplicateCollection = DuplicateCollection_
End Property