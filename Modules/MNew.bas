Attribute VB_Name = "MNew"
Option Explicit

Public Function CCollection(ByVal IsHashed As Boolean, Optional Col As Collection = Nothing) As CCollection
    Set CCollection = New CCollection: CCollection.New_ IsHashed, Col
End Function

