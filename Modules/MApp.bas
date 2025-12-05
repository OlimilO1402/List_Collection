Attribute VB_Name = "MApp"
Option Explicit

Sub Main()
    FMain.Show
End Sub

Public Function VbVarType_ToStr(vt As VbVarType) As String
    Dim s As String
    Select Case vt
    Case VbVarType.vbEmpty:           s = "Empty"      ' 0
    Case VbVarType.vbNull:            s = "Null"       ' 1
    Case VbVarType.vbInteger:         s = "Integer"    ' 2
    Case VbVarType.vbLong:            s = "Long"       ' 3
    Case VbVarType.vbSingle:          s = "Single"     ' 4
    Case VbVarType.vbDouble:          s = "Double"     ' 5
    Case VbVarType.vbCurrency:        s = "Currency"   ' 6
    Case VbVarType.vbDate:            s = "Date"       ' 7
    Case VbVarType.vbString:          s = "String"     ' 8
    Case VbVarType.vbObject:          s = "Object"     ' 9
    Case VbVarType.vbError:           s = "Error"      '10
    Case VbVarType.vbBoolean:         s = "Boolean"    '11
    Case VbVarType.vbVariant:         s = "Variant"    '12
    Case VbVarType.vbDataObject:      s = "DataObject" '13
    Case VbVarType.vbDecimal:         s = "Decimal"    '14
                                                       '15
                                                       '16
    Case VbVarType.vbByte:            s = "Byte"       '17
    Case VbVarType.vbUserDefinedType: s = "UserDefinedType" '36
    Case VbVarType.vbArray:           s = "Array"    '8192
    Case Else: s = "undefined": Debug.Print vt
    End Select
    VbVarType_ToStr = s
End Function
    
Public Function GetRandomName() As String
    Dim n1 As Long: n1 = 20 * Rnd: Dim c1 As Long: c1 = 65# + Rnd * 25#
    Dim n2 As Long: n2 = 20 * Rnd: Dim c2 As Long: c2 = 65# + Rnd * 25#
    Dim nam1 As String: nam1 = ChrW(c1)
    Dim nam2 As String: nam2 = ChrW(c2)
    Dim i As Long
    For i = 0 To n1
        c1 = 97# + 25# * VBA.Math.Rnd
        nam1 = nam1 & ChrW(c1)
    Next
    For i = 0 To n2
        c2 = 97# + 25# * VBA.Math.Rnd
        nam2 = nam2 & ChrW(c2)
    Next
    GetRandomName = nam1 & " " & nam2
End Function

