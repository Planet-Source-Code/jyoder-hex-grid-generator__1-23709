VERSION 5.00
Begin VB.Form frmHexGrid 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000C000&
   Caption         =   "Hex Grid"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   -537.74
   ScaleMode       =   0  'User
   ScaleTop        =   625
   ScaleWidth      =   806.788
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDraw 
      BackColor       =   &H0000C000&
      Caption         =   "Change Hex Gen Specs"
      ForeColor       =   &H00FFFFFF&
      Height          =   2415
      Left            =   9240
      TabIndex        =   5
      Top             =   1800
      Width           =   1935
      Begin VB.TextBox txtColumns 
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox txtRows 
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtRadius 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   360
         Width           =   615
      End
      Begin VB.CommandButton pbRedraw 
         Caption         =   "&Redraw Grid"
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Columns"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Rows"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Radius"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame fraMouse 
      BackColor       =   &H0000C000&
      Caption         =   "To Change Hex Color"
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   9240
      TabIndex        =   0
      Top             =   120
      Width           =   1935
      Begin VB.OptionButton optClick 
         BackColor       =   &H0000C000&
         Caption         =   "Click Mouse"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.OptionButton optMove 
         BackColor       =   &H0000C000&
         Caption         =   "Move Mouse"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   840
         Width           =   1335
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Hope someone can find it useful!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9120
      TabIndex        =   4
      Top             =   6720
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"Main.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   9120
      TabIndex        =   3
      Top             =   4440
      Width           =   2175
   End
End
Attribute VB_Name = "frmHexGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' This Hex Grid code is the beginning of a game I'm working on but I thought I'd
'       share it at this stage for anyone who can use it for their own games/ideas.
'       Also note that the settings in the Scale properties for this form make the
'       lower left corner 0,0 and incrementing X moves right, while Y moves up.

Private Type tCoor
    xCoor As Double     ' xCoor at center of hex
    yCoor As Double     ' yCoor at center of hex
    p1x As Double       ' x coor vertice at 1 o' clock position
    p1y As Double       ' x coor vertice at 1 o' clock position
    p2x As Double       ' (the rest are vertices moving around the hex clockwise)
    p2y As Double
    p3x As Double
    p3y As Double
    p4x As Double
    p4y As Double
    p5x As Double
    p5y As Double
    p6x As Double
    p6y As Double
End Type

Dim mCoor() As tCoor    ' Will contain vertices of all the hexes
Dim r As Long           ' This is radius of circle which would surround the hexes
Dim mxColumns As Long   ' # of columns of hexes to generate
Dim myRows As Long      ' # of rows of hexes to generate
Dim xmStart As Integer  ' Pixel X point to start at
Dim ymStart As Integer  ' Pixel Y point to start at

Private Sub SetupGrid()
Dim X As Integer
Dim Y As Integer
Dim xIncr As Double
Dim yIncr As Double
Dim xFactor As Integer
Dim yFactor As Integer
    
    ReDim mCoor(1 To mxColumns, 1 To myRows)
    
    xIncr = r * 1.5
    yIncr = (r * Sqr(3))
    
    For X = 1 To mxColumns
        For Y = 1 To myRows
            
            xFactor = X - 1
            yFactor = Y - 1
            
            mCoor(X, Y).xCoor = xmStart + (xFactor * xIncr)
            mCoor(X, Y).yCoor = ymStart + (yFactor * yIncr)
            
            If (X / 2) = (X \ 2) Then ' if it's an even column
                ' Up more by half of yIncr cuz it's an even column
                mCoor(X, Y).yCoor = mCoor(X, Y).yCoor + (0.5 * yIncr)
            End If
            
            ' p1 = point at the 1 o'clock position, then move around clockwise
            mCoor(X, Y).p1x = mCoor(X, Y).xCoor + (r * 0.5)
            mCoor(X, Y).p1y = mCoor(X, Y).yCoor + (yIncr * 0.5)
            
            mCoor(X, Y).p2x = mCoor(X, Y).xCoor + r
            mCoor(X, Y).p2y = mCoor(X, Y).yCoor
            
            mCoor(X, Y).p3x = mCoor(X, Y).xCoor + (r * 0.5)
            mCoor(X, Y).p3y = mCoor(X, Y).yCoor - (yIncr * 0.5)
            
            mCoor(X, Y).p4x = mCoor(X, Y).xCoor - (r * 0.5)
            mCoor(X, Y).p4y = mCoor(X, Y).yCoor - (yIncr * 0.5)
            
            mCoor(X, Y).p5x = mCoor(X, Y).xCoor - r
            mCoor(X, Y).p5y = mCoor(X, Y).yCoor
            
            mCoor(X, Y).p6x = mCoor(X, Y).xCoor - (r * 0.5)
            mCoor(X, Y).p6y = mCoor(X, Y).yCoor + (yIncr * 0.5)
            
        Next Y
    Next X
    
End Sub

Private Sub DrawHex(xGrid As Integer, yGrid As Integer, Optional vColor As Variant)
Dim lColor As Long
    
    If (IsMissing(vColor)) Then lColor = vbBlack Else lColor = vColor
    
    ' Draw single hex
    With mCoor(xGrid, yGrid)
        
        ' Draw overlapping circles if you want instead
        'Circle (.xCoor, .yCoor), r, lColor
        
        ' Draws a hex
        Line (.p1x, .p1y)-(.p2x, .p2y), lColor
        Line (.p2x, .p2y)-(.p3x, .p3y), lColor
        Line (.p3x, .p3y)-(.p4x, .p4y), lColor
        Line (.p4x, .p4y)-(.p5x, .p5y), lColor
        Line (.p5x, .p5y)-(.p6x, .p6y), lColor
        Line (.p6x, .p6y)-(.p1x, .p1y), lColor
        
    End With
    
End Sub

Private Sub DrawGrid()
Dim X As Integer
Dim Y As Integer
    Cls
    For X = 1 To mxColumns
        For Y = 1 To myRows
            DrawHex X, Y
        Next Y
    Next X
End Sub

Private Sub Form_Load()

    ' Change various values here to get different results
    Me.BackColor = RGB(50, 135, 50)
    optClick.BackColor = Me.BackColor
    optMove.BackColor = Me.BackColor
    fraMouse.BackColor = Me.BackColor
    fraDraw.BackColor = Me.BackColor
    
    optClick = True
    
    Me.DrawWidth = 1
    
    xmStart = 40
    ymStart = 125
    mxColumns = 14
    myRows = 10
    r = 28
    
    txtRadius = r
    txtColumns = mxColumns
    txtRows = myRows
    
    SetupGrid
    DrawGrid
    
End Sub


Private Function Distance(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double) As Double
    Distance = Sqr(((x2 - x1) ^ 2) + ((y2 - y1) ^ 2))
End Function


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim xx As Integer
Dim yy As Integer
Dim dDist As Double
Dim dClosest As Double
Dim xHold As Integer
Dim yHold As Integer
    
    ' This proc will detect which center of which hex is closest to the point
    ' on the form you clicked which will be whatever hex you clicked on
    
    dClosest = 10000
    
    For xx = 1 To mxColumns
        For yy = 1 To myRows
            dDist = Distance(X, Y, mCoor(xx, yy).xCoor, mCoor(xx, yy).yCoor)
            If dDist < dClosest Then
                xHold = xx
                yHold = yy
                dClosest = dDist
            End If
        Next yy
    Next xx
    
    ' This is outside the grid coordinates, so ignore it
    If (dClosest > r) Then Exit Sub
    
    'MsgBox xHold & ", " & yHold
    DrawHex xHold, yHold, vbYellow
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If optMove Then
        Form_MouseDown Button, Shift, X, Y
    End If
End Sub

Private Sub pbRedraw_Click()
    r = txtRadius
    mxColumns = txtColumns
    myRows = txtRows
    SetupGrid
    DrawGrid
End Sub
