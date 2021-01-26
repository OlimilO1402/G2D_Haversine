VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GeoPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Implements Comparator

Public latitude  As Double
Public longitude As Double

Public Sub GeoPos()
End Sub
Public Sub New_(lat As Double, lon As Double)
   latitude = lat
   longitude = lon
End Sub

'Public Sub Draw(g As Graphics)
Public Sub Draw(g As PictureBox)
   'g.fillRect(latitude*4, longitude*4, 4, 4)
   Dim la As Double: la = Me.latitude * 4
   Dim lo As Double: lo = Me.longitude * 4
   g.Line (la - 2, lo - 2)-(la + 4, lo + 4), , B 'F
End Sub

Public Function HaverSineDistanceTo(other As GeoPos) As Double
   HaverSineDistanceTo = HaverSineDistance(Me, other)
   'HaverSineDistanceTo = PythagorasDistance(Me, other)
End Function

Private Function HaverSineDistance(p1 As GeoPos, p2 As GeoPos) As Double

'http://rosettacode.org/wiki/Haversine_formula
'https://www.movable-type.co.uk/scripts/latlong.html



'const R = 6371000 '; // metres
'const phi1 = lat1 * Math.PI/180 '; // f, ? in radians
'const phi2 = lat2 * Math.PI/180 ';
'const Delta_phi = (lat2-lat1) * Math.PI/180 ';
'const Delta_lam = (lon2-lon1) * Math.PI/180 ';
'
'const a = Math.sin(Delta_phi / 2) * Math.sin(Delta_phi / 2) +
'          Math.cos(phi_1) * Math.cos(phi_2) *
'          Math.sin(Delta_lam / 2) * Math.sin(Delta_lam / 2) ';
'const c = 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a)) ';
'
'const d = R * c; // in metres



   Dim dlon As Double: dlon = Math_toRadians(Abs(p1.longitude - p2.longitude)) 'longitude and latitude are angles
   Dim dlat As Double: dlat = Math_toRadians(Abs(p1.latitude - p2.latitude))   'longitude and latitude are angles
   Dim sindlat2 As Double: sindlat2 = Sin(dlat / 2)
   Dim sindlon2 As Double: sindlon2 = Sin(dlon / 2)
   'haversin(theta) = sin^2(theta/2) = (1-cos(theta))/2
   Dim a As Double: a = (sindlat2 * sindlat2) + Cos(Math_toRadians(p1.latitude)) * Cos(Math_toRadians(p2.latitude)) * (sindlon2 * sindlon2)
   Dim sq As Double: sq = Math.Sqr(a)
   Dim mn As Double: mn = minD(1#, sq)
   Dim c As Double: c = 2 * Math_ASin(mn)
   Dim R As Double: R = 6371008.767 'earth radius in m
   HaverSineDistance = R * c
End Function

Function minD(v1 As Double, v2 As Double) As Double
    If v1 < v2 Then minD = v1 Else minD = v2
End Function

Private Function PythagorasDistance(p1 As GeoPos, p2 As GeoPos) As Double
   Dim dx As Double: dx = (p1.latitude - p2.latitude)
   Dim dy As Double: dy = (p1.longitude - p2.longitude)
   PythagorasDistance = Math.Sqr(dx * dx + dy * dy)
End Function

'Public Static Function CreateRandomGeoPos() As GeoPos
Public Function CreateRandomGeoPos() As GeoPos
   Randomize
   Dim l1 As Double: l1 = Rnd * 90
   Dim l2 As Double: l2 = Rnd * 90
   Set CreateRandomGeoPos = New GeoPos: CreateRandomGeoPos.New_ l1, l2
End Function

Private Function Math_toRadians(a As Double) As Double
    'Const pi As Double = 3.14159265358979
    Dim pi: pi = 4 * Math.Atn(1) '3.14159265358979
    Math_toRadians = a * pi / 180
End Function

Private Function Math_ASin(ByVal y As Double) As Double   ' aka arcsin
    Dim pi: pi = 4 * Math.Atn(1) '3.14159265358979
    Select Case y
        Case 1
            Math_ASin = 0.5 * pi
        Case -1
            Math_ASin = -0.5 * pi
        Case Else
            Math_ASin = VBA.Math.Atn(y / Sqr(1 - y * y))
    End Select
End Function

Friend Function ToString() As String
    Dim sfm As String: sfm = "0.000"
    ToString = "{lat:" & Format(Me.latitude, sfm) & "; lon: " & Format(Me.longitude, sfm) & "}"
End Function
