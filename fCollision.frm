VERSION 5.00
Begin VB.Form fCollision 
   Caption         =   "Collision Test"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10995
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   548
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   733
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   35
      Left            =   10080
      Top             =   2280
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7935
      Left            =   120
      ScaleHeight     =   529
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   617
      TabIndex        =   0
      Top             =   120
      Width           =   9255
   End
   Begin VB.Label Label1 
      Caption         =   "Click Image to Restart"
      Height          =   495
      Left            =   9600
      TabIndex        =   1
      ToolTipText     =   "Click Image to Restart"
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "fCollision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim BI()          As cBiARC
Dim NB            As Long
Dim P()           As tVec2
Dim NP            As Long

Dim SRF2          As cCairoSurface
Dim CC2           As cCairoContext

Dim BiarcPath     As cBiArcPath

Dim T             As Double

Dim TS            As Long

Dim BallPos       As tVec2
Dim BallVel       As tVec2
Dim BallRadius    As Double
Dim BallAngVel    As Double
Dim BallAng       As Double

Dim LeftSide      As Boolean

Private Sub Form_Activate()
    TEST
End Sub

Private Sub PIC_Click()
    TEST

End Sub

Public Sub TEST()
    Dim tp        As tVec2

    Dim I&
    With CC2: .SetSourceColor 0: .Paint: End With

    Set BiarcPath = New cBiArcPath

    With BiarcPath

        .AddPointAndTangDirection Vec2(20, 20), Vec2(0, 1)

        .AddPointAndTangDirection Vec2(50, 300), Vec2(0, 1)

        For I = 1 To 4
            .AddPointAndTangDirection Vec2(I * 105, 550 - I * 80 - (Rnd * 2 - 1) * 100), Vec2(1, 1)
        Next
        .AddPointAndTangDirection Vec2(600, 20), Vec2(0, -1)


        If Rnd < 0.5 Then                        'Flip X to ( to check other side)
            For I = 1 To .Npoints
                tp = .Point(I)
                tp.X = 620 - tp.X
                .SetPoint I, tp
            Next
            LeftSide = False
            Me.Caption = "TEST Collison RIGHT Side"
        Else
            LeftSide = True
            Me.Caption = "TEST Collison LEFT Side"
        End If


        .AutoSetTangents
    End With
    BallPos = Vec2(50 + Rnd * 500, 20)
    BallVel = Vec2(0, 0)
    BallRadius = 15

    Timer1.Enabled = True



    For I = 1 To NP
        If I Mod 2 = 0 Then CC2.SetSourceColor vbRed Else: CC2.SetSourceColor vbYellow
        CC2.Arc P(I).X, P(I).Y, 8
        CC2.Fill
        CC2.TextOut P(I).X - 4, P(I).Y - 7, CStr(I)
    Next


    SRF2.DrawToDC PIC.hDC

End Sub

Private Sub Form_Load()
    Randomize Timer

    Set SRF2 = Cairo.CreateSurface(Me.PIC.Width, Me.PIC.Height, ImageSurface)    'size of our rendering-area in Pixels
    Set CC2 = SRF2.CreateContext                 'create a Drawing-Context from the PixelSurface above
    CC2.AntiAlias = CAIRO_ANTIALIAS_BEST
    CC2.SetLineCap CAIRO_LINE_CAP_ROUND
    CC2.SetLineJoin CAIRO_LINE_JOIN_ROUND
End Sub





Private Sub Timer1_Timer()
    Dim P         As tVec2
    Dim I         As Long
    Dim R         As tVec2

    T = T + 0.0071
    If T > 1 Then T = T - 1

    With CC2: .SetSourceColor 0: .Paint: End With

    CC2.SetSourceColor vbCyan, 0.75
    CC2.Arc BallPos.X, BallPos.Y, BallRadius
    CC2.Stroke

    CC2.MoveTo BallPos.X, BallPos.Y
    CC2.RelLineTo BallRadius * Cos(BallAng), BallRadius * Sin(BallAng)
    CC2.Stroke


    BiarcPath.DRAW CC2, vbYellow, 1, 3           ', True
    If BiarcPath.Closed Then
        CC2.SetSourceColor vbCyan, 0.33
        BiarcPath.DrawOnlyCairoArcs CC2
        CC2.Fill
    End If

    CC2.SetSourceColor vbGreen, 0.5
    For I = 1 To BiarcPath.Npoints
        P = BiarcPath.Point(I)
        CC2.Arc P.X, P.Y, 8
        CC2.Fill
        CC2.TextOut P.X - 4, P.Y - 7, CStr(I)
    Next

    CC2.SetSourceColor vbWhite
    P = BiarcPath.InterpolatedPointAt(T)
    CC2.Arc P.X, P.Y, 7.5
    CC2.Fill

    SRF2.DrawToDC PIC.hDC
    DoEvents


    '----------------- BALL
    For I = 1 To 5
        BallVel = MUL2(BallVel, 0.999)
        BallAngVel = BallAngVel * 0.999
        BallVel.Y = BallVel.Y + 0.05             ' Gravity
        BallPos = SUM2(BallPos, BallVel)
        BallAng = BallAng + BallAngVel

        'BiarcPath.CircleCollisionAndResponse BallPos, BallVel, BallRadius, 0.95
        BiarcPath.CircleCollisionAndResponse2 BallPos, BallVel, BallRadius, BallAngVel, 0.95

    Next
    '-------------------------

End Sub




