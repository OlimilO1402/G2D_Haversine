VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Shop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Name As String
Public GeoPos As New GeoPos

Public Sub Shop()
End Sub

Public Sub New_(nam As String, pos As GeoPos)
    Name = nam
    Set GeoPos = pos
End Sub
'Public Sub Shop(nam As String, lat As Double, lon As Double)
'    Name = nam
'    GeoPos = New GeoPos(lat, lon)
'End Sub
'
Public Function hashCode() As Long
    Dim jsnam As String: jsnam = Name
    hashCode = jsnam Xor GeoPos.latitude Xor GeoPos.longitude
End Function

Public Function ToString() As String
    ToString = Name & GeoPos.ToString '& " lat: " & CStr(GeoPos.latitude) & " lon: " & CStr(GeoPos.longitude)
End Function

'Public Sub Draw(g As Graphics)
Public Sub Draw(g As PictureBox)
    GeoPos.Draw g
    g.Print Name ', GeoPos.latitude * 4, GeoPos.longitude * 4
End Sub

Public Static Function CreateRandomShop() As Shop
    Set CreateRandomShop = New Shop: CreateRandomShop.New_ CreateRandomName(3), GeoPos.CreateRandomGeoPos
End Function

Function CreateRandomName(namlen As Integer) As String
    Dim n As String, i As Integer
    For i = 1 To namlen: n = n & Chr(Asc("A") + Rnd() * (Asc("Z") - Asc("A"))): Next
    CreateRandomName = n
End Function
