VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Triangulate"
   ClientHeight    =   5685
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4935
      Left            =   120
      ScaleHeight     =   4875
      ScaleWidth      =   6675
      TabIndex        =   0
      Top             =   600
      Width           =   6735
   End
   Begin VB.Label Label1 
      Caption         =   "Click The Pic Box to draw triangles"
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   120
      Width           =   2655
   End
   Begin VB.Label lblTris 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblPoints 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tPoints As Integer 'Variable for total number of points (vertices)

Private Sub Form_Load()
'Initiate total points to 1, using base 0 causes problems in the functions
tPoints = 1
End Sub

Private Sub mnuabout_Click()
frmAbout.Show
End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

'variable to hold how many triangles are created by the triangulate function
Dim HowMany As Integer

'Set Vertex coordinates where you clicked the pic box
Vertex(tPoints).x = x
Vertex(tPoints).y = y

'Perform Triangulation Function if there are more than 2 points
If tPoints > 2 Then
    'Clear the Picture Box
    Picture1.Cls
    'Returns number of triangles created.
    HowMany = Triangulate(tPoints)
Else
    'Draw a circle where you clicked so it does something
    Picture1.Circle (Vertex(tPoints).x, Vertex(tPoints).y), 20, vbBlack
End If

'Increment the total number of points
tPoints = tPoints + 1

'Display the total points and total triangles
lblPoints.Caption = "Points: " & tPoints
lblTris.Caption = "Triangles: " & HowMany

'Draw the created triangles
For i = 1 To HowMany
    Picture1.Line (Vertex(Triangle(i).vv0).x, Vertex(Triangle(i).vv0).y)-(Vertex(Triangle(i).vv1).x, Vertex(Triangle(i).vv1).y)
    Picture1.Line (Vertex(Triangle(i).vv1).x, Vertex(Triangle(i).vv1).y)-(Vertex(Triangle(i).vv2).x, Vertex(Triangle(i).vv2).y)
    Picture1.Line (Vertex(Triangle(i).vv0).x, Vertex(Triangle(i).vv0).y)-(Vertex(Triangle(i).vv2).x, Vertex(Triangle(i).vv2).y)
Next i

End Sub
