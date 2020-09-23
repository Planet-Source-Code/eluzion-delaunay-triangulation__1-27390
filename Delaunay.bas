Attribute VB_Name = "Module1"
'Credit to Paul Bourke (pbourke@swin.edu.au) for the original Fortran 77 Program :))
'Conversion by EluZioN (EluZioN@casesladder.com)
'You can use this code however you like providing the above credits remain in tact

Option Explicit

'Points (Vertices)
Public Type dVertex
    x As Long
    y As Long
    z As Long
End Type

'Created Triangles, vv# are the vertex pointers
Public Type dTriangle
    vv0 As Long
    vv1 As Long
    vv2 As Long
End Type

'Set these as applicable
Public Const MaxVertices = 500
Public Const MaxTriangles = 1000

'Our points
Public Vertex(MaxVertices) As dVertex

'Our Created Triangles
Public Triangle(MaxTriangles) As dTriangle

Private Function InCircle(xp As Long, yp As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long, x3 As Long, y3 As Long, ByRef xc, ByRef yc, ByRef r) As Boolean
'Return TRUE if the point (xp,yp) lies inside the circumcircle
'made up by points (x1,y1) (x2,y2) (x3,y3)
'The circumcircle centre is returned in (xc,yc) and the radius r
'NOTE: A point on the edge is inside the circumcircle
     
Dim eps As Double
Dim m1 As Double
Dim m2 As Double
Dim mx1 As Double
Dim mx2 As Double
Dim my1 As Double
Dim my2 As Double
Dim dx As Double
Dim dy As Double
Dim rsqr As Double
Dim drsqr As Double

eps = 0.000001

InCircle = False
      
If Abs(y1 - y2) < eps And Abs(y2 - y3) < eps Then
    MsgBox "INCIRCUM - F - Points are coincident !!"
    Exit Function
End If

If Abs(y2 - y1) < eps Then
    m2 = -(x3 - x2) / (y3 - y2)
    mx2 = (x2 + x3) / 2
    my2 = (y2 + y3) / 2
    xc = (x2 + x1) / 2
    yc = m2 * (xc - mx2) + my2
ElseIf Abs(y3 - y2) < eps Then
    m1 = -(x2 - x1) / (y2 - y1)
    mx1 = (x1 + x2) / 2
    my1 = (y1 + y2) / 2
    xc = (x3 + x2) / 2
    yc = m1 * (xc - mx1) + my1
Else
    m1 = -(x2 - x1) / (y2 - y1)
    m2 = -(x3 - x2) / (y3 - y2)
    mx1 = (x1 + x2) / 2
    mx2 = (x2 + x3) / 2
    my1 = (y1 + y2) / 2
    my2 = (y2 + y3) / 2
    xc = (m1 * mx1 - m2 * mx2 + my2 - my1) / (m1 - m2)
    yc = m1 * (xc - mx1) + my1
End If
      
dx = x2 - xc
dy = y2 - yc
rsqr = dx * dx + dy * dy
r = Sqr(rsqr)
dx = xp - xc
dy = yp - yc
drsqr = dx * dx + dy * dy

If drsqr <= rsqr Then InCircle = True
        
End Function
Private Function WhichSide(xp As Long, yp As Long, x1 As Long, y1 As Long, x2 As Long, y2 As Long) As Integer
'Determines which side of a line the point (xp,yp) lies.
'The line goes from (x1,y1) to (x2,y2)
'Returns -1 for a point to the left
'         0 for a point on the line
'        +1 for a point to the right
 
Dim equation As Double

equation = ((yp - y1) * (x2 - x1)) - ((y2 - y1) * (xp - x1))

If equation > 0 Then
    WhichSide = -1
ElseIf equation = 0 Then
    WhichSide = 0
Else
    WhichSide = 1
End If

End Function

Public Function Triangulate(nvert As Integer) As Integer
'Takes as input NVERT vertices in arrays Vertex()
'Returned is a list of NTRI triangular faces in the array
'Triangle(). These triangles are arranged in clockwise order.

Dim Complete(MaxTriangles) As Boolean
Dim Edges(2, MaxTriangles * 3) As Long
Dim Nedge As Long

'For Super Triangle
Dim xmin As Long
Dim xmax As Long
Dim ymin As Long
Dim ymax As Long
Dim xmid As Long
Dim ymid As Long
Dim dx As Double
Dim dy As Double
Dim dmax As Double

'General Variables
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim ntri As Integer
Dim xc As Double
Dim yc As Double
Dim r As Double
Dim inc As Boolean

'Find the maximum and minimum vertex bounds.
'This is to allow calculation of the bounding triangle
xmin = Vertex(1).x
ymin = Vertex(1).y
xmax = xmin
ymax = ymin
For i = 2 To nvert
    If Vertex(i).x < xmin Then xmin = Vertex(i).x
    If Vertex(i).x > xmax Then xmax = Vertex(i).x
    If Vertex(i).y < ymin Then ymin = Vertex(i).y
    If Vertex(i).y > ymax Then ymax = Vertex(i).y
Next i
dx = xmax - xmin
dy = ymax - ymin
If dx > dy Then
    dmax = dx
Else
    dmax = dy
End If
xmid = (xmax + xmin) / 2
ymid = (ymax + ymin) / 2

'Set up the supertriangle
'This is a triangle which encompasses all the sample points.
'The supertriangle coordinates are added to the end of the
'vertex list. The supertriangle is the first triangle in
'the triangle list.

Vertex(nvert + 1).x = xmid - 2 * dmax
Vertex(nvert + 1).y = ymid - dmax
Vertex(nvert + 2).x = xmid
Vertex(nvert + 2).y = ymid + 2 * dmax
Vertex(nvert + 3).x = xmid + 2 * dmax
Vertex(nvert + 3).y = ymid - dmax
Triangle(1).vv0 = nvert + 1
Triangle(1).vv1 = nvert + 2
Triangle(1).vv2 = nvert + 3
Complete(1) = False
ntri = 1

'Include each point one at a time into the existing mesh
For i = 1 To nvert
    Nedge = 0
    'Set up the edge buffer.
    'If the point (Vertex(i).x,Vertex(i).y) lies inside the circumcircle then the
    'three edges of that triangle are added to the edge buffer.
    j = 0
    Do
        j = j + 1
        If Complete(j) <> True Then
            inc = InCircle(Vertex(i).x, Vertex(i).y, Vertex(Triangle(j).vv0).x, Vertex(Triangle(j).vv0).y, Vertex(Triangle(j).vv1).x, Vertex(Triangle(j).vv1).y, Vertex(Triangle(j).vv2).x, Vertex(Triangle(j).vv2).y, xc, yc, r)
            'Include this if points are sorted by X
            'If (xc + r) < Vertex(i).x Then
                'complete(j) = True
            'Else
            If inc Then
                Edges(1, Nedge + 1) = Triangle(j).vv0
                Edges(2, Nedge + 1) = Triangle(j).vv1
                Edges(1, Nedge + 2) = Triangle(j).vv1
                Edges(2, Nedge + 2) = Triangle(j).vv2
                Edges(1, Nedge + 3) = Triangle(j).vv2
                Edges(2, Nedge + 3) = Triangle(j).vv0
                Nedge = Nedge + 3
                Triangle(j).vv0 = Triangle(ntri).vv0
                Triangle(j).vv1 = Triangle(ntri).vv1
                Triangle(j).vv2 = Triangle(ntri).vv2
                Complete(j) = Complete(ntri)
                j = j - 1
                ntri = ntri - 1
            End If
            'End If
        End If
    Loop While j < ntri

    'Tag multiple edges
    'Note: if all triangles are specified anticlockwise then all
    'interior edges are opposite pointing in direction.
    For j = 1 To Nedge - 1
        If Not Edges(1, j) = 0 And Not Edges(2, j) = 0 Then
            For k = j + 1 To Nedge
                If Not Edges(1, k) = 0 And Not Edges(2, k) = 0 Then
                    If Edges(1, j) = Edges(2, k) Then
                        If Edges(2, j) = Edges(1, k) Then
                            Edges(1, j) = 0
                            Edges(2, j) = 0
                            Edges(1, k) = 0
                            Edges(2, k) = 0
                         End If
                     End If
               End If
             Next k
        End If
    Next j
    
    'Form new triangles for the current point
    'Skipping over any tagged edges.
    'All edges are arranged in clockwise order.
    For j = 1 To Nedge
            If Not Edges(1, j) = 0 And Not Edges(2, j) = 0 Then
                ntri = ntri + 1
                Triangle(ntri).vv0 = Edges(1, j)
                Triangle(ntri).vv1 = Edges(2, j)
                Triangle(ntri).vv2 = i
                Complete(ntri) = False
            End If
    Next j
Next i

'Remove triangles with supertriangle vertices
'These are triangles which have a vertex number greater than NVERT
i = 0
Do
    i = i + 1
    If Triangle(i).vv0 > nvert Or Triangle(i).vv1 > nvert Or Triangle(i).vv2 > nvert Then
        Triangle(i).vv0 = Triangle(ntri).vv0
        Triangle(i).vv1 = Triangle(ntri).vv1
        Triangle(i).vv2 = Triangle(ntri).vv2
        i = i - 1
        ntri = ntri - 1
    End If
Loop While i < ntri

Triangulate = ntri
End Function

