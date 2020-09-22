VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Contour"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9705
   LinkTopic       =   "Form1"
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   647
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Contour By Somsak Thumsatiwinai ThaiLand (somsak2@ksc.th.com)"
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   6120
      Width           =   9735
      Begin VB.CommandButton Command4 
         Caption         =   "Quit"
         Height          =   615
         Left            =   7080
         TabIndex        =   9
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Draw Triangle"
         Height          =   615
         Left            =   5400
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Draw Contour"
         Height          =   615
         Left            =   4200
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Random Point"
         Height          =   615
         Left            =   3000
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Text            =   "50"
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Text            =   "2"
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Div. Height"
         Height          =   375
         Left            =   720
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "No.Point (X,Y)"
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.PictureBox draw 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   0
      ScaleHeight     =   399
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   645
      TabIndex        =   0
      Top             =   0
      Width           =   9705
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Dim hdcdraw As Long
Private Type XYZ
x As Double
y As Double
z As Double
End Type
Private Type TRI
p1 As Integer
p2 As Integer
p3 As Integer
End Type
Private Type SEGS
x1 As Double
y1 As Double
x2 As Double
y2 As Double
End Type
Private Type VERT
x As Double
y As Double
End Type
Dim api As POINTAPI
Dim seg(1000) As SEGS
Dim ver(1000) As VERT
Dim t(2000) As TRI
Dim pt(1000) As XYZ
Dim noTri As Integer
Dim nopt As Integer
Dim p(2) As XYZ
Dim xc As Double, yc As Double, r As Double
Dim ht As Double
Dim noVert As Integer
Dim noSeg As Integer
'Dim noCont As Integer
Private Sub Form_Load()
hdcdraw = draw.hdc
nopt = 50
ht = 2
noTri = 0
End Sub


Private Sub Form_Paint()
draw.height = Form1.ScaleHeight - 70
Frame1.Top = draw.height
Frame1.Left = 0
Frame1.height = 70
Frame1.Width = Form1.ScaleWidth

Command1_Click
End Sub

Private Sub Form_Resize()
Command1_Click
End Sub


Private Sub Command2_Click()
'draw contour
Dim i As Integer
Dim noContour As Integer
Dim z As Double
Dim div1 As Double
Dim div2 As Double
Dim div3 As Double
Dim no As Integer
Dim mu As Double
Dim noline As Integer
Dim nseg As Integer
Dim color As Integer
Dim found As Boolean
Dim j As Integer, k As Integer
Dim tx As Double, ty As Double
Dim maxheight As Double
Dim height As Double

maxheight = 20#
noContour = CInt(maxheight / ht)
height = 0# 'start height
draw.Cls
'draw.DrawWidth = 2
For i = 0 To noContour - 1
noSeg = 0
    
    For j = 0 To noTri - 1
    div1 = pt(t(j).p2).z - pt(t(j).p1).z
    div2 = pt(t(j).p2).z - pt(t(j).p3).z
    div3 = pt(t(j).p3).z - pt(t(j).p1).z
    no = 0
    If Abs(div1) > 0.0001 Then
    mu = -(pt(t(j).p1).z - height) / div1
        If mu >= 0 And mu <= 1 Then
        p(no).x = pt(t(j).p1).x + mu * (pt(t(j).p2).x - pt(t(j).p1).x)
        p(no).y = pt(t(j).p1).y + mu * (pt(t(j).p2).y - pt(t(j).p1).y)
        p(no).z = height
        no = no + 1 'start line /end line
        End If
    End If
    If Abs(div2) > 0.0001 Then
    mu = -(pt(t(j).p3).z - height) / div2
        If mu >= 0 And mu <= 1 Then
        p(no).x = pt(t(j).p3).x + mu * (pt(t(j).p2).x - pt(t(j).p3).x)
        p(no).y = pt(t(j).p3).y + mu * (pt(t(j).p2).y - pt(t(j).p3).y)
        p(no).z = height
        no = no + 1 'start line /end line
        End If
    End If
    If Abs(div3) > 0.0001 Then
    mu = -(pt(t(j).p1).z - height) / div3
        If mu >= 0 And mu <= 1 Then
        p(no).x = pt(t(j).p1).x + mu * (pt(t(j).p3).x - pt(t(j).p1).x)
        p(no).y = pt(t(j).p1).y + mu * (pt(t(j).p3).y - pt(t(j).p1).y)
        p(no).z = height
        no = no + 1 'start line /end line
        End If
    End If
        
    If no = 2 Then  'line OK
    
    seg(noSeg).x1 = p(0).x
    seg(noSeg).y1 = p(0).y
    seg(noSeg).x2 = p(1).x
    seg(noSeg).y2 = p(1).y
    noSeg = noSeg + 1
    End If


    Next j
'---------find polyline contour
        Do While noSeg > 0
        noVert = 0
        ver(noVert).x = seg(0).x1
        ver(noVert).y = seg(0).y1
        noVert = noVert + 1
        ver(noVert).x = seg(0).x2
        ver(noVert).y = seg(0).y2
        
            For k = 0 To noSeg - 2
            seg(k) = seg(k + 1)
            Next k
            noSeg = noSeg - 1

            nseg = 0
      
            Do While nseg < noSeg
            found = False
            If (Abs(ver(0).x - seg(nseg).x1) < 1#) And (Abs(ver(0).y - seg(nseg).y1) < 1#) Then
        
                For k = 0 To noVert
                ver(noVert + 1 - k) = ver(noVert - k) '+1
                Next k
        
                ver(0).x = seg(nseg).x2
                ver(0).y = seg(nseg).y2
        
                noVert = noVert + 1
                found = True
            
                For k = nseg To noSeg - 2
                seg(k) = seg(k + 1)
                Next k
            
                noSeg = noSeg - 1

            ElseIf (Abs(ver(0).x - seg(nseg).x2) < 1#) And (Abs(ver(0).y - seg(nseg).y2) < 1#) Then
        
                For k = 0 To noVert
                ver(noVert + 1 - k) = ver(noVert - k)
                Next k
                
                ver(0).x = seg(nseg).x1
                ver(0).y = seg(nseg).y1
        
                noVert = noVert + 1
                found = True
                
                For k = nseg To noSeg - 2
                seg(k) = seg(k + 1)
                Next k
                noSeg = noSeg - 1

            ElseIf (Abs(ver(noVert).x - seg(nseg).x1) < 1#) And (Abs(ver(noVert).y - seg(nseg).y1) < 1#) Then
        
                noVert = noVert + 1
                ver(noVert).x = seg(nseg).x2
                ver(noVert).y = seg(nseg).y2
                found = True
                
                For k = nseg To noSeg - 2
                seg(k) = seg(k + 1)
                Next k
                noSeg = noSeg - 1

            ElseIf (Abs(ver(noVert).x - seg(nseg).x2) < 1#) And (Abs(ver(noVert).y - seg(nseg).y2) < 1#) Then
        
                noVert = noVert + 1
                ver(noVert).x = seg(nseg).x1
                ver(noVert).y = seg(nseg).y1
        
                found = True
                For k = nseg To noSeg - 2
                seg(k) = seg(k + 1)
                Next k
                noSeg = noSeg - 1
            End If
        
        If found Then
        nseg = 0
        Else
        nseg = nseg + 1
        End If
        
        Loop
  
        DrawLineContour 5
        Loop
    height = height + ht
Next i
'draw.DrawWidth = 1

End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Command1_Click()
Dim i As Integer, j As Integer
Dim no As Integer
Dim x As Double, y As Double
    
    no = CInt(Sqr(nopt))
    x = (draw.Width - 10) / no
    y = (draw.height - 10) / no
    
    Randomize Timer
    
    For i = 0 To no - 1
    For j = 0 To no - 1

    pt(i * no + j).x = Format(Rnd * x + i * x, "0.000")
    pt(i * no + j).y = Format(Rnd * y + j * y, "0.000")
    pt(i * no + j).z = Format(Rnd * 20 + 1, "0.000")

    Next j
    Next i
    noTri = 0
    nopt = no * no
    CalTriangle
    Command2_Click
End Sub


Private Sub Command3_Click()
Dim i As Integer
    'draw triangle
    
    draw.ForeColor = QBColor(1)
    For i = 0 To noTri - 1
    
    MoveToEx hdcdraw, pt(t(i).p1).x, pt(t(i).p1).y, api
    LineTo hdcdraw, pt(t(i).p2).x, pt(t(i).p2).y
    LineTo hdcdraw, pt(t(i).p3).x, pt(t(i).p3).y
    LineTo hdcdraw, pt(t(i).p1).x, pt(t(i).p1).y
    
    Next i
    
    For i = 0 To nopt - 1
    draw.Circle (pt(i).x, pt(i).y), 2, QBColor(4)
    Next i
    
End Sub

Private Function inCircle(ByVal xp As Double, ByVal yp As Double, ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, ByVal x3 As Double, ByVal y3 As Double) As Boolean
Dim m1 As Double, m2 As Double, mx1 As Double, mx2 As Double, my1 As Double, my2 As Double
Dim dx As Double, dy As Double, rsqr As Double, drsqr As Double
   
   If Abs(y1 - y2) < 0.0001 And Abs(y2 - y3) < 0.0001 And Abs(y1 - y3) < 0.0001 Then
   inCircle = False
   r = 100000#
   xc = 0
   yc = 0
   Exit Function
   End If
       
    If Abs(x1 - x2) < 0.0001 And Abs(x2 - x3) < 0.0001 And Abs(x1 - x3) < 0.0001 Then
    inCircle = False
    r = 100000#
    xc = 0
    yc = 0
    Exit Function
    End If

   If Abs(y2 - y1) < 0.0001 Then
      m2 = -(x3 - x2) / (y3 - y2)
      mx2 = (x2 + x3) / 2#
      my2 = (y2 + y3) / 2#
      xc = (x2 + x1) / 2#
      yc = m2 * (xc - mx2) + my2
    ElseIf Abs(y3 - y2) < 0.0001 Then
      m1 = -(x2 - x1) / (y2 - y1)
      mx1 = (x1 + x2) / 2#
      my1 = (y1 + y2) / 2#
      xc = (x3 + x2) / 2#
      yc = m1 * (xc - mx1) + my1
    Else
      m1 = -(x2 - x1) / (y2 - y1)
      m2 = -(x3 - x2) / (y3 - y2)
      mx1 = (x1 + x2) / 2#
      mx2 = (x2 + x3) / 2#
      my1 = (y1 + y2) / 2#
      my2 = (y2 + y3) / 2#
      xc = (m1 * mx1 - m2 * mx2 + my2 - my1) / (m1 - m2)
      yc = m1 * (xc - mx1) + my1
   End If

   dx = x2 - xc
   dy = y2 - yc
   rsqr = dx * dx + dy * dy
   r = Sqr(rsqr)    'radius

   dx = xp - xc
   dy = yp - yc
   drsqr = dx * dx + dy * dy + 1
   
   If drsqr < rsqr Then
   inCircle = True
   Else
   inCircle = False
   End If

End Function

Private Sub Text1_Change()

ht = CInt(0 & Text1.Text)
If ht < 1 Then ht = 1
'draw.Cls
Command2_Click
End Sub

Private Sub Text2_Change()
nopt = CInt(0 & Text2.Text)
If nopt < 3 Then nopt = 3
Command1_Click
End Sub


Private Sub DrawLineContour(segment As Integer)
Dim i As Integer, j As Integer, last As Integer
Dim x As Double, y As Double
Dim u As Double, nc1 As Double, nc2 As Double, nc3 As Double, nc4 As Double
Dim xp(2000) As Double, yp(2000) As Double
Dim nv As Integer
Dim check As Boolean

    check = False
    If Abs(ver(0).x - ver(noVert).x) < 1# And Abs(ver(0).y - ver(noVert).y) < 1# Then
    ver(0).x = (ver(0).x + ver(1).x) / 2#
    ver(0).y = (ver(0).y + ver(1).y) / 2#
    ver(noVert + 1) = ver(noVert)
    ver(noVert).x = (ver(noVert - 1).x + ver(noVert).x) / 2#
    ver(noVert).y = (ver(noVert - 1).y + ver(noVert).y) / 2#
    noVert = noVert + 1
    check = True
    End If
nv = 0

    xp(0) = ver(0).x
    yp(0) = ver(0).y
    xp(1) = ver(0).x
    yp(1) = ver(0).y
    xp(2) = 0.5 * ver(1).x + 0.5 * ver(0).x
    yp(2) = 0.5 * ver(1).y + 0.5 * ver(0).y
nv = nv + 3
    For i = 1 To noVert - 1
    xp(nv + 1) = ver(i).x
    yp(nv + 1) = ver(i).y
    xp(nv) = 0.5 * ver(i).x + 0.5 * ver(i - 1).x
    yp(nv) = 0.5 * ver(i).y + 0.5 * ver(i - 1).y
    nv = nv + 2
    Next i

    If check Then
    xp(nv) = ver(noVert).x
    yp(nv) = ver(noVert).y
    xp(nv + 1) = ver(0).x
    yp(nv + 1) = ver(0).y
    nv = nv + 2
    Else
    xp(nv + 1) = ver(noVert).x
    yp(nv + 1) = ver(noVert).y
    xp(nv) = 0.5 * ver(noVert).x + 0.5 * ver(noVert - 1).x
    yp(nv) = 0.5 * ver(noVert).y + 0.5 * ver(noVert - 1).y
    nv = nv + 2
    End If

    i = nv

    xp(i) = xp(i - 1)
    yp(i) = yp(i - 1)
    xp(i + 1) = xp(i)
    yp(i + 1) = yp(i)
    last = i + 2
    
    draw.ForeColor = QBColor(0)
    MoveToEx hdcdraw, xp(0), yp(0), api
    
    For i = 1 To last - 3
    u = 0
    Do While u <= 1
    nc1 = -(u * u * u / 6) + u * u / 2 - u / 2 + 1 / 6
    nc2 = u * u * u / 2 - u * u + 2 / 3
    nc3 = (-u * u * u + u * u + u) / 2 + 1 / 6
    nc4 = u * u * u / 6
    x = nc1 * xp(i - 1) + nc2 * xp(i) + nc3 * xp(i + 1) + nc4 * xp(i + 2)
    y = nc1 * yp(i - 1) + nc2 * yp(i) + nc3 * yp(i + 1) + nc4 * yp(i + 2)
    
    LineTo hdcdraw, x, y
    u = u + 1 / segment
    Loop
    Next i
    
End Sub

Private Sub CalTriangle()
Dim no As Integer, index As Integer
Dim inside As Boolean, nopoint As Integer
Dim i As Integer, j As Integer, k As Integer
Dim xp As Double, yp As Double, x1 As Double, y1 As Double, x2 As Double, y2 As Double
Dim x3 As Double, y3 As Double, z3 As Double, xt As Double, yt As Double
Dim starttri As Integer, newtri As Integer
Dim rmin As Double, r2 As Double, r3 As Double, rt As Double
Dim xc1 As Double, xc2 As Double, xc3 As Double, xct As Double
Dim yc1 As Double, yc2 As Double, yc3 As Double, yct As Double
Dim rt1 As Double, rt2 As Double, rt3 As Double
Dim ok(1000) As Boolean
Dim usepoint As Boolean
    
    noTri = 0
    
    For i = 0 To nopt
    ok(i) = False
    Next i
'----------------------------find 1st tri
    i = 0
       For j = i + 1 To nopt - 2
            For k = j + 1 To nopt - 1
         
            x1 = pt(i).x
            y1 = pt(i).y
            x2 = pt(j).x
            y2 = pt(j).y
            x3 = pt(k).x
            y3 = pt(k).y
       
            inside = inCircle(x1, y1, x1, y1, x2, y2, x3, y3)  'find r,xc,yc
                           
             inside = False
             For no = 0 To nopt - 1
             r2 = (xc - pt(no).x) * (xc - pt(no).x) + (yc - pt(no).y) * (yc - pt(no).y) + 1
             
             If r2 < r * r Then
             inside = True
             Exit For
             End If
             Next no
             
            If Not inside Then
                t(noTri).p1 = j
                t(noTri).p2 = k
                t(noTri).p3 = i
                ok(j) = True
                ok(k) = True
                ok(i) = True
                noTri = noTri + 1
            
            End If
            
            If noTri = 1 Then
            Exit For
            End If
            
            Next k
            If noTri = 1 Then
            Exit For
            End If
        Next j

'-----------------find next tri
            newtri = 0

Do While newtri < noTri
      newtri = noTri
      
      For i = 1 To nopt - 1
        If Not ok(i) Then
      
            xp = pt(i).x
            yp = pt(i).y
            
            For no = 0 To noTri - 1
            x1 = pt(t(no).p1).x
            y1 = pt(t(no).p1).y
            x2 = pt(t(no).p2).x
            y2 = pt(t(no).p2).y
            x3 = pt(t(no).p3).x
            y3 = pt(t(no).p3).y
' side 1-------------------
      
        inside = inCircle(x2, y2, x1, y1, x3, y3, xp, yp)
        If Not inside Then
             inside = False
             For k = 0 To nopt - 1
             r2 = (xc - pt(k).x) * (xc - pt(k).x) + (yc - pt(k).y) * (yc - pt(k).y) + 1
             
             If r2 < r * r Then
             inside = True
             Exit For
             End If
             Next k
             
             If Not inside And r < 1000# Then
             
                    t(noTri).p1 = t(no).p1
                    t(noTri).p2 = t(no).p3
                    t(noTri).p3 = i
                    noTri = noTri + 1
                    ok(i) = True
             End If
        End If
' side 2-------------------
      
        inside = inCircle(x1, y1, x2, y2, x3, y3, xp, yp)
        If Not inside Then
             inside = False
             For k = 0 To nopt - 1
             r2 = (xc - pt(k).x) * (xc - pt(k).x) + (yc - pt(k).y) * (yc - pt(k).y) + 1
             
             If r2 < r * r Then
             inside = True
             Exit For
             End If
             Next k
             
             If Not inside And r < 1000# Then
             
                    t(noTri).p1 = t(no).p2
                    t(noTri).p2 = t(no).p3
                    t(noTri).p3 = i
                    noTri = noTri + 1
                    ok(i) = True
             End If
        End If
' side 3-------------------
        inside = inCircle(x3, y3, x1, y1, x2, y2, xp, yp)
        If Not inside Then
             inside = False
             For k = 0 To nopt - 1
             r2 = (xc - pt(k).x) * (xc - pt(k).x) + (yc - pt(k).y) * (yc - pt(k).y) + 1
             
             If r2 < r * r Then
             inside = True
             Exit For
             End If
             Next k
             
             If Not inside And r < 1000# Then
             
                    t(noTri).p1 = t(no).p1
                    t(noTri).p2 = t(no).p2
                    t(noTri).p3 = i
                    noTri = noTri + 1
                    ok(i) = True
             End If
        End If
        
        Next no
        
        If ok(i) Then
        Exit For
        End If
        
    End If
     
Next i
Loop

End Sub


