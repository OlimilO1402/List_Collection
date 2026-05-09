Attribute VB_Name = "MNew"
Option Explicit

Public Function CCollection(ByVal IsHashed As Boolean, Optional col As Collection = Nothing, Optional Name As String = "", Optional ByVal OptionBaseLBound As Long = 1) As CCollection
    Set CCollection = New CCollection: CCollection.New_ IsHashed, col, Name, OptionBaseLBound
End Function

