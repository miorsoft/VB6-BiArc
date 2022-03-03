VERSION 5.00
Begin VB.Form fImage 
   Caption         =   "Image Biarc"
   ClientHeight    =   8415
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   ScaleHeight     =   561
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   694
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   30
      Left            =   9720
      Top             =   6240
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
      ToolTipText     =   "Play / Pause animation and change image"
      Top             =   120
      Width           =   9255
   End
End
Attribute VB_Name = "fImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim BiArc         As cBiARC
Dim SRF2          As cCairoSurface
Dim CC2           As cCairoContext

Dim iSRF          As cCairoSurface

Dim pW            As Long
Dim pH            As Long


Private Sub Form_Load()

    Set BiArc = New cBiARC


    pW = Me.PIC.Width - 1
    pH = Me.PIC.Height - 1
    Set SRF2 = Cairo.CreateSurface(pW + 1, pH + 1, ImageSurface)    'size of our rendering-area in Pixels


    Set CC2 = SRF2.CreateContext                 'create a Drawing-Context from the PixelSurface above
    CC2.AntiAlias = CAIRO_ANTIALIAS_BEST
    CC2.SetLineCap CAIRO_LINE_CAP_ROUND
    CC2.SetLineJoin CAIRO_LINE_JOIN_ROUND


    Set iSRF = Cairo.ImageList.AddImage("Loaded", App.Path & "\res\Rope.png")



    With BiArc
        .Point1 = Vec2(70, 70)
        .ControlPt1 = Vec2(300, 100)
        .Point2 = Vec2(460, 370)
        .ControlPt2 = Vec2(550, 377)
        .CALC
        .DRAW CC2, vbRed, 1, 4
        SRF2.DrawToDC PIC.hDC
    End With

End Sub


Private Sub Form_Activate()
    PIC_Click
    Timer1.Enabled = True

End Sub

Private Sub PIC_Click()
    Static I&
    Randomize Timer



    DrawImage True


    Timer1.Enabled = Not (Timer1.Enabled)

    If Timer1.Enabled Then
        I = (I + 1) Mod 3
        Select Case I
        Case 0
            Set iSRF = Cairo.ImageList.AddImage("Loaded", App.Path & "\res\Barbwire.png")
        Case 1
            Set iSRF = Cairo.ImageList.AddImage("Loaded", App.Path & "\res\Chain.png")
        Case 2
            Set iSRF = Cairo.ImageList.AddImage("Loaded", App.Path & "\res\Rope.png")
        'Case 3
        '    Set iSRF = Cairo.ImageList.AddImage("Loaded", App.Path & "\res\Smoke3.png")
        End Select
    End If



End Sub

Private Sub DrawImage(ClearScreen As Boolean)
    Dim ArrSource() As Long
    Dim ArrSrf2() As Long

    Dim W         As Long
    Dim H         As Long
    Dim X         As Double
    Dim Y         As Double
    Dim tx#, ty#
    Dim A#, R#, T As tVec2, wA&
    Dim cen       As tVec2
    Dim CosA#, SinA#
    Dim Segno#
    Dim currR#
    Dim StepX#
    Dim Repeat    As Double

    Dim ScaleH    As Double


    If ClearScreen Then CC2.SetSourceColor 0: CC2.Paint


    iSRF.BindToArrayLong ArrSource
    SRF2.BindToArrayLong ArrSrf2

    W = UBound(ArrSource, 1)
    H = UBound(ArrSource, 2)


    Repeat = 3.5
    ScaleH = 0.75


    With BiArc

        'Repeat = .LengthTot / W 'AutoRepeat

        StepX = 0.9 * (W * Repeat) / .LengthTot


        For X = 0 To W * Repeat Step StepX

            T = .InterpolatedPointAt((X / W) / Repeat, wA)
            If wA = 1 Then
                cen = .Center1
                R = .Radius1
                Segno = IIf(.ClockWise1, 1, -1)
            Else
                cen = .Center2
                R = .Radius2
                Segno = IIf(.ClockWise2, 1, -1)
            End If

            A = Atan2(T.X - cen.X, T.Y - cen.Y)

            CosA = Cos(A)
            SinA = Sin(A)

            For Y = 0 To H Step 0.5
                currR = (Segno * ScaleH * (Y - H * 0.5) + R)

                tx = cen.X + CosA * currR
                ty = cen.Y + SinA * currR
                If tx >= 0 Then
                    If ty >= 0 Then
                        If tx <= pW Then
                            If ty <= pH Then
                                ArrSrf2(tx, ty) = ArrSource(X Mod W, Y)
                            End If
                        End If
                    End If
                End If
            Next
        Next

        CC2.SetSourceColor vbCyan, 0.25
        CC2.Arc .ControlPt1.X, .ControlPt1.Y, 7
        CC2.MoveTo .ControlPt1.X, .ControlPt1.Y: CC2.LineTo .Point1.X, .Point1.Y
        CC2.Stroke
        CC2.Arc .ControlPt2.X, .ControlPt2.Y, 7
        CC2.MoveTo .ControlPt2.X, .ControlPt2.Y: CC2.LineTo .Point2.X, .Point2.Y
        CC2.Stroke

    End With

    iSRF.ReleaseArrayLong ArrSource
    SRF2.ReleaseArrayLong ArrSrf2



    SRF2.DrawToDC PIC.hDC

End Sub

Private Sub Timer1_Timer()
    Dim V         As tVec2

    With BiArc

        V = .ControlPt1
        V.X = V.X + (Rnd * 2 - 1) * 5
        V.Y = V.Y + (Rnd * 2 - 1) * 5
        .ControlPt1 = V


        V = .ControlPt2
        V.X = V.X + (Rnd * 2 - 1) * 5
        V.Y = V.Y + (Rnd * 2 - 1) * 5
        .ControlPt2 = V

        V = .Point2
        V.X = V.X + (Rnd * 2 - 1) * 5
        V.Y = V.Y + (Rnd * 2 - 1) * 5
        .Point2 = V
        .CALC

        DrawImage True


    End With


End Sub
