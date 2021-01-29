VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7395
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   10830
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   10830
   StartUpPosition =   3  'Windows-Standard
   Begin VB.ListBox List2 
      Height          =   1425
      Left            =   120
      TabIndex        =   5
      Top             =   5400
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "zeige die 5 nähesten Shops"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Neu"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.PictureBox g 
      BackColor       =   &H80000005&
      FillColor       =   &H80000012&
      Height          =   6255
      Left            =   3000
      ScaleHeight     =   6195
      ScaleWidth      =   7635
      TabIndex        =   0
      Top             =   600
      Width           =   7695
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   2535
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnuPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuSetAsNewPos 
         Caption         =   "set as new position"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GeoPos As New GeoPos
Dim Shop   As New Shop
'Dim ShopDB       As java#util#ArrayList 'Proofs that this code must be around since Jabaco-time maybe about 2010
Dim ShopDB       As Collection

'Dim nearestshops As java#util#TreeSet
Dim nearestshops As Collection
 
Dim mypos As GeoPos


Public Sub Form_Load()
   '
   g.ScaleMode = vbPixels
End Sub

Public Sub Command1_Click()
    
    'ShopDB = New java#util#ArrayList
    Set ShopDB = New Collection
    Set nearestshops = New Collection
    Set mypos = GeoPos.CreateRandomGeoPos
    Label1.Caption = mypos.ToString
    
    List1.Clear
    List2.Clear
    Dim i As Integer
    For i = 0 To 19
        Dim ashop As Shop: Set ashop = Shop.CreateRandomShop()
        ShopDB.Add ashop
        List1.AddItem ashop.ToString
    Next
    Me.Refresh
    
End Sub

Public Sub Command2_Click()
    Set nearestshops = CalcNearest5(mypos)
    Dim ashop As Shop
    List2.Clear
    For Each ashop In nearestshops
        List2.AddItem ashop.ToString
    Next
    Me.Refresh
End Sub

Function minD(v1 As Double, v2 As Double) As Double
    If v1 < v2 Then minD = v1 Else minD = v2
End Function
Function MaxD(v1 As Double, v2 As Double) As Double
    If v1 > v2 Then MaxD = v1 Else MaxD = v2
End Function

'http://rosettacode.org/wiki/Haversine_formula

'https://www.movable-type.co.uk/scripts/latlong.html
'a = sin²(Delta_phi / 2) + cos(phi_1) · cos(phi_2) · sin²(Delta_lam / 2)
'c = 2 · atan2( WURZEL(a), WURZEL(1-a) )
'd = R · c
'phi is latitude, lam is longitude, R is earth’s radius (mean radius = 6371km)
'note that angles need to be in radians to pass to trig functions!
'
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

'Function CalcNearest5(pos As GeoPos) As java#util#TreeSet
Function CalcNearest5(pos As GeoPos) As Collection
'OK, so ist der Algo ein Mist!!
'alle Shops durchlaufen,
'und zur Liste der nächstens Pos hinzufügen
'falls die Liste die maximale Anzahl (hier 5) übersteigt
'dann den mit der weitesten Entfernung durch den neuen ersetzen, wenn der neue weiter entfernt ist dann den neuen nicht hinzufügen
'
    'nearestshops = New java#util#LinkedList
    'nein ein TreeSet ist das Mittel der Wahl
    'nearestshops = New java#util#TreeSet
    Set nearestshops = New Collection
    Dim newD As Double
    Dim newMinD As Double, oldMaxD As Double, d As Double
    Dim ashop As Shop, bshop As Shop
    Dim hsd As Double
    Dim i As Long 'Integer
    Dim max_i As Long
    For Each ashop In ShopDB
        newD = ashop.GeoPos.HaverSineDistanceTo(pos)
        
        If nearestshops.Count > 4 Then
            'zuerst die maximale Distanz der Bereits in der liste befindlichen Shops herausfinden
            'und zugleich dessen index
            For i = 1 To nearestshops.Count
                Set bshop = nearestshops.Item(i)
                hsd = bshop.GeoPos.HaverSineDistanceTo(pos)
                oldMaxD = MaxD(oldMaxD, hsd)
                If oldMaxD = hsd Then
                    max_i = i
                End If
            Next
            'wenn die neue Distanz kleiner ist als die maximale distanz der bereits inder Liste befindlichen Shops
            If newD < oldMaxD Then
                'OK die neue distanz ist kleiner
                'dann den mit der größten distanz löschen
                nearestshops.Remove max_i
                'und das neue Element hinzufügen:
                nearestshops.Add ashop
            End If
            oldMaxD = 0
            max_i = 0
        Else
            nearestshops.Add ashop
        End If
        
        '????
        'For i = 0 To nearestshops.Count - 1
        '    bshop = nearestshops.iterator.Next
        'Next
        
        
        'If newD < oldD Then
        '   nearestshops.addFirst ashop
        '   If nearestshops.size = 6 Then nearestshops.removeLast
        'End If
        'oldD = newD
    Next
    
    Set CalcNearest5 = nearestshops
End Function

'Public Sub Form_Paint(g As Graphics)
Public Sub Form_Paint()
    g.Cls
    Dim ashop As Shop
    'If ShopDB <> Nothing Then
    If Not ShopDB Is Nothing Then
        'g.setColor (Color.black)
        g.ForeColor = vbBlack
        For Each ashop In ShopDB: ashop.Draw g: Next
    End If
    'If nearestshops <> Nothing Then
    If Not nearestshops Is Nothing Then
        'g.setColor (Color.red)
        g.ForeColor = vbRed
        For Each ashop In nearestshops: ashop.Draw Me.g: Next
    End If
    'If mypos <> Nothing Then
    If Not mypos Is Nothing Then
        'g.setColor (Color.blue)
        g.ForeColor = vbBlue
        mypos.Draw g
    End If
End Sub

Public Sub List1_Click()
    Dim ashop As Shop: Set ashop = ShopDB.Item(List1.ListIndex + 1)
    'Dim g As Graphics = Me.getGraphics
    'Dim g As PictureBox: Set g = Me.PictureBox1 'Me.getGraphics
    'g.setColor Color.red
    'g.Cls
    Form_Paint
    g.ForeColor = vbGreen 'Red
    ashop.Draw g
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = MouseButtonConstants.vbRightButton Then
        PopupMenu Me.mnuPopup
    End If
End Sub

Private Sub mnuSetAsNewPos_Click()
    Dim i As Long: i = List1.ListIndex
    Dim gp As GeoPos: Set gp = ShopDB.Item(i + 1).GeoPos
    Dim tmp As GeoPos: Set tmp = mypos
    If Not gp Is Nothing Then Set mypos = gp
    Set ShopDB.Item(i + 1).GeoPos = tmp
    
    Me.Refresh

End Sub
