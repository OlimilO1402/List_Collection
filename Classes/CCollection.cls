VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit 'om 2025-11-30 lines: 410; 2025-12-04 lines: 464;
'the Collection is 1-based
'class CCollection : replaces and extends the vb intrinsic-Collection with new additional functions
'Implements Collection 'nope
Private m_Col       As Collection 'Of Object/Variant
Private m_IsHashed  As Boolean
Private m_Version   As Variant 'Decimal 'Incremented in every Add, Insert or Delete

Private Sub Class_Initialize()
    Set m_Col = New Collection
End Sub

Friend Sub New_(IsHashed As Boolean, Optional Col As Collection = Nothing)
    m_IsHashed = IsHashed: If Not Col Is Nothing Then Set m_Col = Col
End Sub

'Copy to MNew
'Public Function CCollection(ByVal IsHashed As Boolean, Optional Col As Collection = Nothing) As CCollection
'    Set CCollection = New CCollection: CCollection.New_ IsHashed, Col
'End Function

Public Sub NewC(other As CCollection)
    Dim i As Long
    With other
        For i = 1 To other.Count
            m_Col.Add other.Item(i)
        Next
    End With
End Sub

'Optional includingObjects As Boolean = False
Public Function Clone() As CCollection
    Set Clone = New CCollection: Clone.NewC Me
End Function

Public Property Get DataType() As VbVarType
    If m_Col.Count = 0 Then Exit Property
    DataType = VarType(m_Col.Item(1))
End Property

Public Function Add(Item, Optional Key, Optional Before, Optional After)
    If IsObject(Item) Then
        Dim Obj As Object: Set Obj = Item
        Set Add = Obj
        If m_IsHashed Then
            If IsMissing(Key) Then Key = Obj.Key
        End If
    End If
    If IsMissing(Key) Then
        If IsMissing(Before) Then
            If IsMissing(After) Then
                m_Col.Add Item
            Else
                m_Col.Add Item, , , After
            End If
        Else
            m_Col.Add Item, , Before 'before togehter with after makes no sense
        End If
    Else
        If IsMissing(Before) Then
            If IsMissing(After) Then
                m_Col.Add Item, Key
            Else
                m_Col.Add Item, Key, , After
            End If
        Else
            m_Col.Add Item, Key, Before
        End If
    End If
    m_Version = m_Version + CDec(1)
End Function

Public Function AddKey(Obj As Object) As Object
    ' the object needs to have a Public Function/PropertyGet Key As String
    Set AddKey = Obj: m_Col.Add Obj, Obj.Key
    m_Version = m_Version + CDec(1)
End Function

Public Function AddOrGet(Obj As Object) As Object
    Dim Key As String: Key = Obj.Key ' the object needs to have a Public Function/PropertyGet Key As String
    If Me.Contains(Key) Then
        Set AddOrGet = m_Col.Item(Key)
    Else
        Set AddOrGet = Obj
        m_Col.Add Obj, Key
        m_Version = m_Version + CDec(1)
    End If
End Function

Public Function TryAddObject(Obj As Object, ByVal Key As String) As Boolean
Try: On Error GoTo Catch
    m_Col.Add Obj, Key
    TryAddObject = True
    m_Version = m_Version + CDec(1)
Catch: On Error GoTo 0
End Function

Public Sub ChangeKey(OldIndexKey, NewKey As String)
    'for this function all credits go to the incredible www.vb-tec.de alias Jost Schwider
    'you can find the original version of the function "CollectionChangeKey" here: https://vb-tec.de/collctns.htm
    Dim Value As Variant
    If IsObject(m_Col(OldIndexKey)) Then
        Set Value = m_Col.Item(OldIndexKey)
    Else
        Value = m_Col.Item(OldIndexKey)
    End If
    m_Col.Add Value, NewKey  'first add
    m_Col.Remove OldIndexKey 'then remove
    m_Version = m_Version + CDec(1)
End Sub

Public Sub Clear()
    Class_Initialize
    m_Version = m_Version + CDec(1)
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = m_Col.[_NewEnum]
End Function

Public Function Contains(ByVal Key As String) As Boolean
    'for this function all credits go to the incredible www.vb-tec.de alias Jost Schwider
    'you can find the original version of the function "IsInCollection" here: https://vb-tec.de/collctns.htm
    On Error Resume Next
    '"Extras->Optionen->Allgemein->Unterbrechen bei Fehlern->Bei nicht verarbeiteten Fehlern"
    If IsEmpty(m_Col(Key)) Then: 'DoNothing
    Contains = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Property Get Count() As Long
    Count = m_Col.Count
End Property

Public Function IndexFromObject(Obj As Object) As Long
    Dim i As Long, v, o As Object
    For Each v In m_Col
        Set o = v
        If o.Key = Obj.Key Then
            IndexFromObject = i
            Exit Function
        End If
        i = i + 1
    Next
End Function

Public Property Get IsHashed() As Boolean
    IsHashed = m_IsHashed
End Property

Public Sub Insert(ByVal Index As Long, ValueOrObj)
    'insert always means "Befor" the Index
    'you can insert before the first element to get a new first element
    'if you want to insert after the last element just use Add
    m_Col.Add ValueOrObj, , Index
    m_Version = m_Version + CDec(1)
End Sub

Public Property Get Item(IndexOrKey)
    If IsObject(m_Col.Item(IndexOrKey)) Then
        Set Item = m_Col.Item(IndexOrKey)
    Else
        Item = m_Col.Item(IndexOrKey)
    End If
End Property

Public Property Let Item(IndexOrKey, ValueOrObj)
    Dim vtik As VbVarType: vtik = VarType(IndexOrKey)
    If vtik = vbString Then 'its a key
        m_Col.Remove IndexOrKey
        m_Col.Add ValueOrObj, IndexOrKey
    ElseIf vtik = vbLong Or vtik = vbInteger Then
        m_Col.Remove IndexOrKey
        m_Col.Add ValueOrObj, , IndexOrKey
    End If
    m_Version = m_Version + CDec(1)
End Property

Public Property Set Item(IndexOrKey, ValueOrObj As Object)
    Dim vtik As VbVarType: vtik = VarType(IndexOrKey)
    If vtik = vbString Then 'its a key
        m_Col.Remove IndexOrKey
        m_Col.Add ValueOrObj, IndexOrKey
    ElseIf vtik = vbLong Or vtik = vbInteger Then
        m_Col.Remove IndexOrKey
        m_Col.Add ValueOrObj, , IndexOrKey
    End If
    m_Version = m_Version + CDec(1)
End Property

Public Sub MoveUp(ByVal i As Long)
    Dim c As Long: c = m_Col.Count
    If i <= 1 Or c < i Then Exit Sub
    SwapItems i, i - 1
    m_Version = m_Version + CDec(1)
End Sub

Public Sub MoveUpKey(ByVal i As Long)
    Dim c As Long: c = m_Col.Count
    If i <= 1 Or c < i Then Exit Sub
    Dim i1 As Long: i1 = i
    Dim i2 As Long: i2 = i - 1
    Dim Obj1 As Object: Set Obj1 = m_Col.Item(i1)
    Dim Obj2 As Object: Set Obj2 = m_Col.Item(i2)
    Dim Key1 As String: Key1 = Obj1.Key
    Dim Key2 As String: Key2 = Obj2.Key
    SwapItemsKey i1, Key1, i2, Key2
    m_Version = m_Version + CDec(1)
End Sub

Public Sub MoveDown(ByVal i As Long)
    Dim c As Long: c = m_Col.Count
    If i < 1 Or c <= i Then Exit Sub
    SwapItems i, i + 1
End Sub

Public Sub MoveDownKey(ByVal i As Long)
    Dim c As Long: c = m_Col.Count
    If i < 1 Or c <= i Then Exit Sub
    Dim i1 As Long: i1 = i
    Dim i2 As Long: i2 = i + 1
    Dim Obj1 As Object: Set Obj1 = m_Col.Item(i1)
    Dim Obj2 As Object: Set Obj2 = m_Col.Item(i2)
    Dim Key1 As String: Key1 = Obj1.Key
    Dim Key2 As String: Key2 = Obj2.Key
    SwapItemsKey i1, Key1, i2, Key2
End Sub

Public Function Remove(IndexOrKey)
    Set Remove = Me.Item(IndexOrKey)
    m_Col.Remove IndexOrKey
    m_Version = m_Version + CDec(1)
End Function

Public Sub SwapItems(ByVal i1 As Long, ByVal i2 As Long)
    Dim c As Long: c = m_Col.Count
    If c = 0 Then Exit Sub
    If i2 < i1 Then: Dim i_tmp As Long: i_tmp = i1: i1 = i2: i2 = i_tmp
    If i1 <= 0 Or c <= i1 Then Exit Sub
    If i2 <= 0 Or c < i2 Then Exit Sub
    If i1 = i2 Then Exit Sub
    Dim Obj1: If IsObject(m_Col.Item(i1)) Then Set Obj1 = m_Col.Item(i1) Else Obj1 = m_Col.Item(i1)
    Dim Obj2: If IsObject(m_Col.Item(i2)) Then Set Obj2 = m_Col.Item(i2) Else Obj2 = m_Col.Item(i2)
    m_Col.Remove i1: m_Col.Add Obj2, , i1:   m_Col.Remove i2
    If i2 < c Then m_Col.Add Obj1, , i2 Else m_Col.Add Obj1
    m_Version = m_Version + CDec(1)
End Sub

Public Sub SwapItemsKey(ByVal i1 As Long, ByVal Key1 As String, _
                        ByVal i2 As Long, ByVal Key2 As String)
    Dim c As Long: c = m_Col.Count
    If c = 0 Then Exit Sub
    If i2 < i1 Then: Dim i_tmp As Long: i_tmp = i1: i1 = i2: i2 = i_tmp
    If i1 <= 0 Or c <= i1 Then Exit Sub
    If i2 <= 0 Or c < i2 Then Exit Sub
    If i1 = i2 Then Exit Sub
    Dim Obj1: If IsObject(m_Col.Item(Key1)) Then Set Obj1 = m_Col.Item(Key1) Else Obj1 = m_Col.Item(Key1)
    Dim Obj2: If IsObject(m_Col.Item(Key2)) Then Set Obj2 = m_Col.Item(Key2) Else Obj2 = m_Col.Item(Key2)
    m_Col.Remove Key2: m_Col.Add Obj2, Key2, i1
    m_Col.Remove Key1: m_Col.Add Obj1, Key1, i1
    m_Version = m_Version + CDec(1)
End Sub

Public Sub ToListBox(aLB As ListBox, Optional ByVal addEmptyLineFirst As Boolean = False, Optional ByVal doPtrToItemData As Boolean = False)
    ToListCtrl aLB, addEmptyLineFirst
End Sub

Public Sub ToComboBox(aCB As ComboBox, Optional ByVal addEmptyLineFirst As Boolean = False, Optional ByVal doPtrToItemData As Boolean = False)
    ToListCtrl aCB, addEmptyLineFirst
End Sub

Public Sub ToListCtrl(ComboBoxOrListBox, Optional ByVal addEmptyLineFirst As Boolean = False, Optional ByVal doPtrToItemData As Boolean = False)
    Dim i As Long, c As Long: c = m_Col.Count: If c = 0 Then Exit Sub
    Dim vt As VbVarType: vt = VarType(m_Col.Item(1))
    Dim v, Obj As Object
    With ComboBoxOrListBox
        If .ListCount Then .Clear
        If addEmptyLineFirst Then .AddItem vbNullString
        Select Case vt
        Case vbByte, vbInteger, vbLong, vbCurrency, vbDate, vbSingle, vbDouble, vbDecimal, vbString
            For i = 1 To c
                .AddItem m_Col.Item(i)
            Next
        Case vbObject
            For i = 1 To c
                Set Obj = m_Col.Item(i)
                .AddItem Obj.ToStr ' the object needs to have a Public Function ToStr As String
                If doPtrToItemData Then .ItemData(i - 1) = Obj.Ptr ': Debug.Print Obj.Ptr ' and a Public Function Ptr As LongPtr
                ' ItemData can only be of type Long no String
            Next
        End Select
    End With
End Sub

Public Property Get ObjectFromListCtrl(ComboBoxOrListBox, ByRef i_out As Long) As Object
    Dim li As Long: li = ComboBoxOrListBox.ListIndex
    If i_out < 0 Then i_out = li
    If i_out < 0 Then Exit Property
    Dim Key As String: Key = ComboBoxOrListBox.ItemData(i_out)
    If Contains(Key) Then Set ObjectFromListCtrl = m_Col.Item(Key)
End Property

Public Function ToStr(Col As Collection) As String
    Dim s As String, v, o As Object
    For Each v In Col
        If IsObject(v) Then
            Set o = v
            s = s & o.ToStr & vbCrLf
        Else
            s = s & CStr(v) & vbCrLf
        End If
    Next
    ToStr = s
End Function

Public Property Get Version() As Variant 'As Decimal
    Version = m_Version
End Property

' Sort:
Public Sub Sort()
    'Set m_Col = Col
    Dim c As Long: c = m_Col.Count
    'If c = 0 Then: Set m_Col = Nothing: Exit Sub
    If c = 0 Then Exit Sub
    Dim vt As VbVarType: vt = VarType(m_Col.Item(1))
    Select Case vt
    Case vbByte, vbInteger, vbLong, vbCurrency, vbDate, vbSingle, vbDouble, vbDecimal
        QuickSortVar 1, c
    Case vbString
        QuickSortStr 1, c
    Case vbObject
        QuickSortObj 1, c
    End Select
    'Set m_Col = Nothing
End Sub
    
' The recursive data-independent QuickSort for primitive data-variables
Private Sub QuickSortVar(ByVal i1 As Long, ByVal i2 As Long)
    Dim t As Long
    If i2 > i1 Then
        t = DivideVar(i1, i2)
        QuickSortVar i1, t - 1
        QuickSortVar t + 1, i2
    End If
End Sub

Private Function DivideVar(ByVal i1 As Long, ByVal i2 As Long) As Long
    Dim i As Long: i = i1 - 1
    Dim j As Long: j = i2
    Dim p As Long: p = j
    Do
        Do
            i = i + 1
        Loop While (CompareVar(i, p) < 0)
        Do
            j = j - 1
        Loop While ((i1 < j) And (CompareVar(p, j) < 0))
        If i < j Then SwapVar i, j
    Loop While (i < j)
    SwapVar i, p
    DivideVar = i
End Function

Private Function CompareVar(ByVal i1 As Long, ByVal i2 As Long) As Variant
    CompareVar = m_Col.Item(i1) - m_Col.Item(i2)
End Function

Private Sub SwapVar(ByVal i1 As Long, ByVal i2 As Long)
    If i1 = i2 Then Exit Sub
    Dim c As Long: c = m_Col.Count
    If i2 < i1 Then: Dim i_tmp As Long: i_tmp = i1: i1 = i2: i2 = i_tmp
    Dim Var1: Var1 = m_Col.Item(i1)
    Dim Var2: Var2 = m_Col.Item(i2)
    m_Col.Remove i1: m_Col.Add Var2, , i1:   m_Col.Remove i2
    If i2 < c Then m_Col.Add Var1, , i2 Else m_Col.Add Var1
End Sub

' The recursive data-independent QuickSort for strings
Private Sub QuickSortStr(ByVal i1 As Long, ByVal i2 As Long)
    Dim t As Long
    If i1 < i2 Then
        t = DivideStr(i1, i2)
        QuickSortStr i1, t - 1
        QuickSortStr t + 1, i2
    End If
End Sub

Private Function DivideStr(ByVal i1 As Long, ByVal i2 As Long) As Long
    Dim i As Long: i = i1 - 1
    Dim j As Long: j = i2
    Dim p As Long: p = j
    Do
        Do
            i = i + 1
        Loop While (CompareStr(i, p) < 0)
        Do
            j = j - 1
        Loop While ((i1 < j) And (CompareStr(p, j) < 0))
        If i < j Then SwapStr i, j
    Loop While (i < j)
    SwapStr i, p
    DivideStr = i
End Function

Private Function CompareStr(ByVal i1 As Long, ByVal i2 As Long)
    CompareStr = StrComp(m_Col.Item(i1), m_Col.Item(i2))
    'Dim Str1 As String: Str1 = m_col.Item(i1)
    'Dim Str2 As String: Str2 = m_col.Item(i2)
    'CompareStr = StrComp(Str1, Str2)
End Function

Private Sub SwapStr(ByVal i1 As Long, ByVal i2 As Long)
    If i1 = i2 Then Exit Sub
    Dim c As Long: c = m_Col.Count
    If i2 < i1 Then: Dim i_tmp As Long: i_tmp = i1: i1 = i2: i2 = i_tmp
    Dim Str1 As String: Str1 = m_Col.Item(i1)
    Dim Str2 As String: Str2 = m_Col.Item(i2)
    m_Col.Remove i1: m_Col.Add Str2, , i1:   m_Col.Remove i2
    If i2 < c Then m_Col.Add Str1, , i2 Else m_Col.Add Str1
End Sub

' The recursive data-independent QuickSort for objects
Private Sub QuickSortObj(ByVal i1 As Long, ByVal i2 As Long)
    Dim t As Long
    If i2 > i1 Then
        t = DivideObj(i1, i2)
        QuickSortObj i1, t - 1
        QuickSortObj t + 1, i2
    End If
End Sub

Private Function DivideObj(ByVal i1 As Long, ByVal i2 As Long) As Long
    Dim i As Long: i = i1 - 1
    Dim j As Long: j = i2
    Dim p As Long: p = j
    Do
        Do
            i = i + 1
        Loop While (CompareObj(i, p) < 0)
        Do
            j = j - 1
        Loop While ((i1 < j) And (CompareObj(p, j) < 0))
        If i < j Then SwapObj i, j
    Loop While (i < j)
    SwapObj i, p
    DivideObj = i
End Function

Private Function CompareObj(ByVal i1 As Long, ByVal i2 As Long) As Long
    Dim Obj1 As Object: Set Obj1 = m_Col.Item(i1)
    Dim Obj2 As Object: Set Obj2 = m_Col.Item(i2)
    CompareObj = Obj1.Compare(Obj2)
End Function

Private Sub SwapObj(ByVal i1 As Long, ByVal i2 As Long)
    If i1 = i2 Then Exit Sub
    Dim c As Long: c = m_Col.Count
    If i2 < i1 Then: Dim i_tmp As Long: i_tmp = i1: i1 = i2: i2 = i_tmp
    Dim Obj1 As Object: Set Obj1 = m_Col.Item(i1)
    Dim Obj2 As Object: Set Obj2 = m_Col.Item(i2)
    m_Col.Remove i1: m_Col.Add Obj2, , i1:   m_Col.Remove i2
    If i2 < c Then m_Col.Add Obj1, , i2 Else m_Col.Add Obj1
End Sub
