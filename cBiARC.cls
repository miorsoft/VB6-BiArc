VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cBiARC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

'http://www.ryanjuckett.com/programming/biarc-interpolation/
'/******************************************************************************
'  Copyright (c) 2014 Ryan Juckett
'  http://www.ryanjuckett.com/
'
'  This software Is provided 'as-is', without any express or implied
'  warranty. In no event will the authors be held liable for any damages
'  arising from the use of this software.
'
'  Permission is granted to anyone to use this software for any purpose,
'  including commercial applications, and to alter it and redistribute it
'  freely, subject to the following restrictions:
'
'  1. The origin of this software must not be misrepresented; you must not
'     claim that you wrote the original software. If you use this software
'     in a product, an acknowledgment in the product documentation would be
'     appreciated but is not required.
'
'  2. Altered source versions must be plainly marked as such, and must not be
'     misrepresented as being the original software.
'
'  3. This notice may not be removed or altered from any source
'     distribution.
'******************************************************************************/

' USAGE:
' -SetPointsAndControlPts
' -CALC
' -DRAW

'******************************************************************************/


Private Const boundsWidth As Double = 1.2
Private Const boundsColor As Long = vbGreen
'Private Const tangentWidth As Double = 1
'Private Const pointColor1 As Long = vbGreen
'Private Const pointColor2 As Long = vbRed
Private Const arcColor1 As Long = vbGreen
Private Const arcColor2 As Long = vbRed


Private mPoint1   As tVec2
Private mPoint2   As tVec2
Private mControlP1 As tVec2                      ' Tangent control Point
Private mControlP2 As tVec2
Private mTangDir1 As tVec2                       'Normalized tangent Direction
Private mTangDir2 As tVec2

Private mCenter1  As tVec2
Private mCenter2  As tVec2
Private mRadius1  As Double
Private mRadius2  As Double
Private mA11      As Double
Private mA12      As Double
Private mA21      As Double
Private mA22      As Double

Private d1        As Double
Private d2        As Double
Private mJoint    As tVec2

Private mLengthArc1#
Private mLengthArc2#
Private mLengthTot#

Private AngArc1#
Private AngArc2#
Private mClockWise1 As Boolean
Private mClockWise2 As Boolean

Private mBound1   As tVec2
Private mBound2   As tVec2
Private mBound3   As tVec2
Private mBound4   As tVec2

Private SourceSRF As cCairoSurface
Private SRF2      As cCairoSurface
Private pW        As Long
Private pH        As Long
Private ImaScaleX As Double
Private ImaScaleY As Double
Private InvScaleX As Double
Private InvScaleY As Double
Private ImaXOffset As Double


Friend Property Let Point1(P As tVec2)
    mPoint1 = P
End Property
Friend Property Let Point2(P As tVec2)
    mPoint2 = P
End Property
Friend Property Let ControlPt1(P As tVec2)
    mControlP1 = P
End Property
Friend Property Let ControlPt2(P As tVec2)
    mControlP2 = P
End Property

Friend Property Get Point1() As tVec2
    Point1 = mPoint1
End Property
Friend Property Get Point2() As tVec2
    Point2 = mPoint2
End Property
Friend Property Get ControlPt1() As tVec2
    ControlPt1 = mControlP1
End Property
Friend Property Get ControlPt2() As tVec2
    ControlPt2 = mControlP2
End Property
Friend Property Get TangentDir1() As tVec2
    TangentDir1 = mTangDir1
End Property
Friend Property Get TangentDir2() As tVec2
    TangentDir2 = mTangDir2
End Property
Friend Property Get Center1() As tVec2
    Center1 = mCenter1
End Property
Friend Property Get Center2() As tVec2
    Center2 = mCenter2
End Property

Friend Property Get Radius1() As Double
    Radius1 = mRadius1
End Property
Friend Property Get Radius2() As Double
    Radius2 = mRadius2
End Property

Friend Property Get A11() As Double
    A11 = mA11
End Property
Friend Property Get A12() As Double
    A12 = mA12
End Property
Friend Property Get A21() As Double
    A21 = mA21
End Property
Friend Property Get A22() As Double
    A22 = mA22
End Property



Friend Property Get LengthArc1() As Double
    LengthArc1 = mLengthArc1
End Property
Friend Property Get LengthArc2() As Double
    LengthArc2 = mLengthArc2
End Property

Friend Property Get LengthTot() As Double
    LengthTot = mLengthTot
End Property

Friend Property Get ClockWise1() As Boolean
    ClockWise1 = mClockWise1
End Property
Friend Property Get ClockWise2() As Boolean
    ClockWise2 = mClockWise2
End Property

Friend Sub GetBounds(A As tVec2, b As tVec2)
    A.X = min(min(mBound1.X, mBound2.X), min(mBound3.X, mBound4.X))
    A.Y = min(min(mBound1.Y, mBound2.Y), min(mBound3.Y, mBound4.Y))
    b.X = max(max(mBound1.X, mBound2.X), max(mBound3.X, mBound4.X))
    b.Y = max(max(mBound1.Y, mBound2.Y), max(mBound3.Y, mBound4.Y))

End Sub


Private Sub CCDrawLine(CC As cCairoContext, mPoint1 As tVec2, mPoint2 As tVec2, Lwidth As Double, Color As Long)
    CC.DrawLine mPoint1.X, mPoint1.Y, mPoint2.X, mPoint2.Y, , Lwidth, Color
End Sub
'''Function DrawArcFromEdge(context, mPoint1, mControlP1, mPoint2, arcWidth, color, fromP1)
'''{
'''    var chord = DIFF2(mPoint2, mPoint1);
'''    var n1 = new Vec2(-mControlP1.y, mControlP1.x);
'''    var chordDotN1 = DOT2(chord, n1);
'''
'''    if (IsEqualEps(chordDotN1,0))
'''    {
'''        // straight line
'''        DrawLine(context, mPoint1, mPoint2, arcWidth, color);
'''    }
'''    Else
'''    {
'''        var radius = Vec2_MagSqr(chord) / (2*chordDotN1);
'''        var center = Vec2_AddScaled(mPoint1, n1, radius);
'''
'''        var p1Offset = DIFF2(mPoint1, center);
'''        var p2Offset = DIFF2(mPoint2, center);
'''
'''        var p1Ang1 = Math.atan2(p1Offset.y,p1Offset.x);
'''        var p2Ang1 = Math.atan2(p2Offset.y,p2Offset.x);
'''        if ( p1Offset.x*mControlP1.y - p1Offset.y*mControlP1.x > 0)
'''            DrawArc(context, center, Math.abs(radius), p1Ang1, p2Ang1, arcWidth, color, !fromP1);
'''        Else
'''            DrawArc(context, center, Math.abs(radius), p1Ang1, p2Ang1, arcWidth, color, fromP1);
'''
'''        context.globalAlpha = 0.05;
'''        DrawCircle(context, center, Math.abs(radius), color);
'''        context.globalAlpha = 1.0;
'''    }
'''}
''Private Sub DrawArcFromEdge(mPoint1 As tVec2, mControlP1 As tVec2, mPoint2 As tVec2, _
 ''                           arcWidth As Double, Color As Long, fromP1 As Boolean)
''    Dim chord     As tVec2
''    Dim N1        As tVec2
''    Dim chordDotN1 As Double
''    Dim Radius    As Double
''    Dim Center    As tVec2
''    Dim p1Offset  As tVec2
''    Dim p2Offset  As tVec2
''    Dim p1Ang1    As Double
''    Dim p2Ang1    As Double
''    Dim TA        As Double
''
''    chord = DIFF2(mPoint2, mPoint1)
''    N1 = Vec2(-mControlP1.Y, mControlP1.x)
''    chordDotN1 = DOT2(chord, N1)
''
''    If (IsEqualEps(chordDotN1, 0)) Then
''
''        '// straight line
''        CCDrawLine mPoint1, mPoint2, arcWidth, Color
''
''    Else
''
''        Radius = LengthSq2(chord) / (2 * chordDotN1)
''        Center = ADDScaled2(mPoint1, N1, Radius)
''
''        p1Offset = DIFF2(mPoint1, Center)
''        p2Offset = DIFF2(mPoint2, Center)
''
''        p1Ang1 = Atan2(p1Offset.x, p1Offset.Y)
''        p2Ang1 = Atan2(p2Offset.x, p2Offset.Y)
''
''
''        If (p1Offset.x * mControlP1.Y - p1Offset.Y * mControlP1.x > 0) Then
''            '    DrawArc(context, center, Math.abs(radius), p1Ang1, p2Ang1, arcWidth, color, !fromP1);
''            '''''            With CC
''            '''''                .SetSourceColor cOLOR
''            '''''                .SetLineWidth arcWidth
''            '''''                If (fromP1) Then TA = p1Ang1: p1Ang1 = p2Ang1: p2Ang1 = TA
''            '''''                .Arc Center.x, Center.Y, Abs(radius), p2Ang1, p1Ang1
''            '''''                .Stroke
''            '''''            End With
''            'If color = vbGreen Then
''            'Debug.Print "side1 " & fromP1
''        Else
''            '    DrawArc(context, center, Math.abs(radius), p1Ang1, p2Ang1, arcWidth, color, fromP1);
''            '''''            With CC
''            '''''                .SetSourceColor cOLOR
''            '''''                .SetLineWidth arcWidth
''            '''''                If Not (fromP1) Then TA = p1Ang1: p1Ang1 = p2Ang1: p2Ang1 = TA
''            '''''                .Arc Center.x, Center.Y, Abs(radius), p2Ang1, p1Ang1
''            '''''                .Stroke
''            '''''            End With
''            'If color = vbGreen Then
''            'Debug.Print "side2 " & fromP1
''        End If
''
''        With CC
''            .SetSourceColor Color, 0.3
''            .Arc Center.x, Center.Y, Abs(Radius)
''            .Fill
''            .SetSourceColor vbYellow
''            .Arc mPoint2.x, mPoint2.Y, 3
''            .Fill
''        End With
''
''
''        '        context.globalAlpha = 0.05;
''        '        DrawCircle(context, center, Math.abs(radius), color);
''        '        context.globalAlpha = 1.0;
''    End If
''
''
''End Sub




''Friend Sub CalcAndDRAWBiARC(mPoint1 As tVec2, mPoint2 As tVec2, _
 ''                            mTangDir1 As tVec2, mTangDir2 As tVec2, _
 ''                            Optional CustomDistance As Double = 0)
''
''
''    Dim d1        As Double
''    Dim d2        As Double
''
''    Dim v         As tVec2
''    Dim vMagSqr   As Double
''    Dim vDotT1    As Double
''
''    Dim vDotT2    As Double
''    Dim t1DotT2   As Double
''    Dim denominator As Double
''
''    Dim mJoint     As tVec2
''    Dim invlen    As Double
''
''    Dim t         As tVec2
''    Dim tMagSqr   As Double
''    Dim equalTangents As Boolean
''    Dim perpT1    As Boolean
''
''    Dim vDotT     As Double
''
''    Dim discriminant As Double
''
''
''    Dim angle     As Double
''    Dim Center1   As tVec2
''    Dim Center2   As tVec2
''    Dim Radius    As Double
''    Dim cross     As Double
''
''
''    d1 = CustomDistance
''
''    '    With CC: .SetSourceColor 0: .Paint: End With
''
''
''    v = DIFF2(mPoint2, mPoint1)
''    vMagSqr = LengthSq2(v)
''
''    vDotT1 = DOT2(v, mTangDir1)
''
''
''    '// if we are using a custom value for d1
''    If (CustomDistance) Then
''
''        vDotT2 = DOT2(v, mTangDir2)
''        t1DotT2 = DOT2(mTangDir1, mTangDir2)
''        denominator = (vDotT2 - d1 * (t1DotT2 - 1))
''
''        If (IsEqualEps(denominator, 0#)) Then
''
''            'document.getElementById('d1').value = d1;
''            'document.getElementById('d2').value = 'Infinity';
''            d1 = CustomDistance
''            d2 = MAX_VALUE
''
''            '// the second arc is a semicircle
''            mJoint = ADDScaled2(mPoint1, mTangDir1, d1)
''            mJoint = ADDScaled2(mJoint, mTangDir2, vDotT2 - d1 * t1DotT2)
''
''            '// draw bounds
''            CCDrawLine mPoint1, ADDScaled2(mPoint1, mTangDir1, d1), boundsWidth, boundsColor
''            CCDrawLine ADDScaled2(mPoint1, mTangDir1, d1), mJoint, boundsWidth, boundsColor
''
''            '// draw arcs
''            DrawArcFromEdge mPoint1, mTangDir1, mJoint, arcWidth, arcColor1, True
''            DrawArcFromEdge mPoint2, mTangDir2, mJoint, arcWidth, arcColor2, False
''
''        Else
''
''            d2 = (0.5 * vMagSqr - d1 * vDotT1) / denominator
''
''            invlen = 1# / (d1 + d2)
''
''            mJoint = MUL2(DIFF2(mTangDir1, mTangDir2), d1 * d2)
''            mJoint = ADDScaled2(mJoint, mPoint2, d1)
''            mJoint = ADDScaled2(mJoint, mPoint1, d2)
''            mJoint = MUL2(mJoint, invlen)
''
''            'document.getElementById('d1').value = d1;
''            'document.getElementById('d2').value = d2;
''
''            '// draw bounds
''            CCDrawLine mPoint1, ADDScaled2(mPoint1, mTangDir1, d1), boundsWidth, boundsColor
''            CCDrawLine ADDScaled2(mPoint1, mTangDir1, d1), mJoint, boundsWidth, boundsColor
''            CCDrawLine mJoint, ADDScaled2(mPoint2, mTangDir2, -d2), boundsWidth, boundsColor
''            CCDrawLine ADDScaled2(mPoint2, mTangDir2, -d2), mPoint2, boundsWidth, boundsColor
''
''            '// draw arcs
''            DrawArcFromEdge mPoint1, mTangDir1, mJoint, arcWidth * 2, arcColor1, True
''            DrawArcFromEdge mPoint2, mTangDir2, mJoint, arcWidth, arcColor2, False
''        End If
''
''        '// else set d1 equal to d2
''    Else
''
''
''        t = Vec2ADD(mTangDir1, mTangDir2)
''        tMagSqr = LengthSq2(t)
''
''        equalTangents = IsEqualEps(tMagSqr, 4#)
''
''        perpT1 = IsEqualEps(vDotT1, 0#)
''        If (equalTangents And perpT1) Then
''            '// we have two semicircles
''            mJoint = ADDScaled2(mPoint1, v, 0.5)
''
''            '                document.getElementById('d1').value = 'Infinity';
''            '                document.getElementById('d2').value = 'Infinity';
''            d1 = MAX_VALUE
''            d2 = MAX_VALUE
''
''            '// draw arcs
''
''
''            angle = Atan2(v.x, v.Y)
''            Center1 = ADDScaled2(mPoint1, v, 0.25)
''            Center2 = ADDScaled2(mPoint1, v, 0.75)
''            Radius = Sqr(vMagSqr) * 0.25
''            cross = v.x * mTangDir1.Y - v.Y * mTangDir1.x
''
''            'DrawArc(context, center1, radius, angle, angle+Math.PI, arcWidth, arcColor1, cross < 0);
''            'DrawArc(context, center2, radius, angle, angle+Math.PI, arcWidth, arcColor2, cross > 0);
''
''            'context.globalAlpha = 0.05;
''            'DrawCircle(context, center1, radius, arcColor1);
''            'DrawCircle(context, center2, radius, arcColor2);
''            'context.globalAlpha = 1.0;
''
''        Else
''
''
''            'Stop
''            vDotT = DOT2(v, t)
''
''            perpT1 = IsEqualEps(vDotT1, 0#)
''
''            If (equalTangents) Then
''
''                d1 = vMagSqr / (4 * vDotT1)
''
''            Else
''
''                denominator = 2 - 2 * DOT2(mTangDir1, mTangDir2)
''                discriminant = vDotT * vDotT + denominator * vMagSqr
''                d1 = (Sqr(discriminant) - vDotT) / denominator
''            End If
''
''            mJoint = MUL2(DIFF2(mTangDir1, mTangDir2), d1)
''            mJoint = Vec2ADD(mJoint, mPoint1)
''            mJoint = Vec2ADD(mJoint, mPoint2)
''            mJoint = MUL2(mJoint, 0.5)
''
''            'document.getElementById('d1').value = d1;
''            'document.getElementById('d2').value = d1;
''
''            '// draw bounds
''            CCDrawLine mPoint1, ADDScaled2(mPoint1, mTangDir1, d1), boundsWidth, boundsColor
''            CCDrawLine ADDScaled2(mPoint1, mTangDir1, d1), mJoint, boundsWidth, boundsColor
''            CCDrawLine mJoint, ADDScaled2(mPoint2, mTangDir2, -d1), boundsWidth, boundsColor
''            CCDrawLine ADDScaled2(mPoint2, mTangDir2, -d1), mPoint2, boundsWidth, boundsColor
''
''            '// draw arcs
''            DrawArcFromEdge mPoint1, mTangDir1, mJoint, arcWidth, arcColor1, True
''            DrawArcFromEdge mPoint2, mTangDir2, mJoint, arcWidth, arcColor2, False
''        End If
''    End If
''
''
''
''
''    '    // draw points
''    '    DrawCircle(context, mPoint1, c_Point_Radius, pointColor1);
''    '    DrawCircle(context, mPoint2, c_Point_Radius, pointColor2);
''
''    '    // draw tangents
''    '    DrawLine(context, mPoint1, ADDScaled2(mPoint1, mTangDir1, c_Tangent_Length), tangentWidth, pointColor1);
''    '    DrawLine(context, mPoint2, ADDScaled2(mPoint2, mTangDir2, c_Tangent_Length), tangentWidth, pointColor2);
''    '---------------------------------------------------------------------------------------------
''
''
''
''    '''''    With CC
''    '''''        .SetSourceColor arcColor1
''    '''''        .Arc mPoint1.x, mPoint1.Y, 5
''    '''''        .Fill
''    '''''        .MoveTo mPoint1.x, mPoint1.Y
''    '''''        .LineTo mControlP1.x, mControlP1.Y
''    '''''        .Stroke
''    '''''
''    '''''        .SetSourceColor arcColor2
''    '''''        .Arc mPoint2.x, mPoint2.Y, 5
''    '''''        .Fill
''    '''''        .MoveTo mPoint2.x, mPoint2.Y
''    '''''        .LineTo mControlP2.x, mControlP2.Y
''    '''''        .Stroke
''    '''''    End With
''
''    '''''    cPTS.Draw CC
''    '    srf.DrawToDC PicHDC
''
''
''End Sub

Friend Sub SetPointsAndControlPts(Point1 As tVec2, Cntrl1 As tVec2, _
                                  Point2 As tVec2, Cntrl2 As tVec2)
    mPoint1 = Point1
    mPoint2 = Point2
    mControlP1 = Cntrl1
    mControlP2 = Cntrl2
    CALC
End Sub
Friend Sub SetPointsAndTangDirs(Point1 As tVec2, TangDir1 As tVec2, _
                                Point2 As tVec2, TangDir2 As tVec2)
    mPoint1 = Point1
    mPoint2 = Point2
    mControlP1 = SUM2(Point1, TangDir1)
    mControlP2 = SUM2(Point2, TangDir2)
    CALC
End Sub



'Friend Sub CalcBiARC(mPoint1 As tVec2, mPoint2 As tVec2, _
 '                     mTangDir1 As tVec2, mTangDir2 As tVec2, _
 '                      ByRef OutCenter1 As tVec2, _
 '                      ByRef OutRadius1 As Double, _
 '                      ByRef OutAng11 As Double, _
 '                      ByRef OutAng12 As Double, _
 '                      ByRef OutCenter2 As tVec2, _
 '                      ByRef OutRadius2 As Double, _
 '                      ByRef OutAng21 As Double, _
 '                      ByRef OutAng22 As Double, _
 '                      Optional ByVal CustomDistance = 0)
'
'
'    Dim d1        As Double
'    Dim d2        As Double
'
'    Dim v         As tVec2
'    Dim vMagSqr   As Double
'    Dim vDotT1    As Double
'
'    Dim vDotT2    As Double
'    Dim t1DotT2   As Double
'    Dim denominator As Double
'
'    Dim mJoint     As tVec2
'    Dim invlen    As Double
'
'    Dim t         As tVec2
'    Dim tMagSqr   As Double
'    Dim equalTangents As Boolean
'    Dim perpT1    As Boolean
'
'    Dim vDotT     As Double
'
'    Dim discriminant As Double
'
'
'    Dim angle     As Double
'    Dim Center1   As tVec2
'    Dim Center2   As tVec2
'    Dim Radius    As Double
'    Dim cross     As Double
'
'
'    d1 = CustomDistance
'
'    '    With CC
'    '        .SetSourceColor 0
'    '        .Paint
'    '    End With
'
'
'    v = DIFF2(mPoint2, mPoint1)
'    vMagSqr = LengthSq2(v)
'
'    vDotT1 = DOT2(v, mTangDir1)
'
'
'    '// if we are using a custom value for d1
'    If (CustomDistance) Then
'
'        vDotT2 = DOT2(v, mTangDir2)
'        t1DotT2 = DOT2(mTangDir1, mTangDir2)
'        denominator = (vDotT2 - d1 * (t1DotT2 - 1))
'
'        If (IsEqualEps(denominator, 0#)) Then
'            d1 = CustomDistance
'            d2 = MAX_VALUE
'            '// the second arc is a semicircle
'            mJoint = ADDScaled2(mPoint1, mTangDir1, d1)
'            mJoint = ADDScaled2(mJoint, mTangDir2, vDotT2 - d1 * t1DotT2)
'
'        Else
'
'            d2 = (0.5 * vMagSqr - d1 * vDotT1) / denominator
'
'            invlen = 1# / (d1 + d2)
'
'            mJoint = MUL2(DIFF2(mTangDir1, mTangDir2), d1 * d2)
'            mJoint = ADDScaled2(mJoint, mPoint2, d1)
'            mJoint = ADDScaled2(mJoint, mPoint1, d2)
'            mJoint = MUL2(mJoint, invlen)
'
'        End If
'
'        '// else set d1 equal to d2
'    Else
'
'
'        t = Vec2ADD(mTangDir1, mTangDir2)
'        tMagSqr = LengthSq2(t)
'
'        equalTangents = IsEqualEps(tMagSqr, 4#)
'
'        perpT1 = IsEqualEps(vDotT1, 0#)
'        If (equalTangents And perpT1) Then
'            '// we have two semicircles
'            mJoint = ADDScaled2(mPoint1, v, 0.5)
'
'            d1 = MAX_VALUE
'            d2 = MAX_VALUE
'            '// draw arcs
'            angle = Atan2(v.x, v.Y)
'            Center1 = ADDScaled2(mPoint1, v, 0.25)
'            Center2 = ADDScaled2(mPoint1, v, 0.75)
'            Radius = Sqr(vMagSqr) * 0.25
'            cross = v.x * mTangDir1.Y - v.Y * mTangDir1.x
'
'            'DrawArc(context, center1, radius, angle, angle+Math.PI, arcWidth, arcColor1, cross < 0);
'            'DrawArc(context, center2, radius, angle, angle+Math.PI, arcWidth, arcColor2, cross > 0);
'            'context.globalAlpha = 0.05;
'            'DrawCircle(context, center1, radius, arcColor1);
'            'DrawCircle(context, center2, radius, arcColor2);
'            'context.globalAlpha = 1.0;
'
'        Else
'
'
'            'Stop
'            vDotT = DOT2(v, t)
'
'            perpT1 = IsEqualEps(vDotT1, 0#)
'
'            If (equalTangents) Then
'
'                d1 = vMagSqr / (4 * vDotT1)
'
'            Else
'
'                denominator = 2 - 2 * DOT2(mTangDir1, mTangDir2)
'                discriminant = vDotT * vDotT + denominator * vMagSqr
'                d1 = (Sqr(discriminant) - vDotT) / denominator
'            End If
'
'            mJoint = MUL2(DIFF2(mTangDir1, mTangDir2), d1)
'            mJoint = Vec2ADD(mJoint, mPoint1)
'            mJoint = Vec2ADD(mJoint, mPoint2)
'            mJoint = MUL2(mJoint, 0.5)
'
'
'        End If
'    End If
'
'
'    '**********************************************************************
'    '**********************************************************************
'    '**********************************************************************
'    ArcFromEdge mPoint1, mTangDir1, mJoint, True, OutCenter1, OutRadius1, OutAng11, OutAng12
'    ArcFromEdge mPoint2, mTangDir2, mJoint, False, OutCenter2, OutRadius2, OutAng21, OutAng22
'
'    '****************************************************************
'    '****************************************************************
'    '****************************************************************
'
'
'End Sub

Private Sub ArcFromEdge(Point As tVec2, tangentDir As tVec2, pJoint As tVec2, fromP1 As Boolean, _
                        ByRef outCenter As tVec2, _
                        ByRef outRadius As Double, _
                        ByRef outAng1 As Double, _
                        ByRef outAng2 As Double)

    Dim chord     As tVec2
    Dim NtangentDir As tVec2
    Dim chordDotN1 As Double
    Dim StartPoint As tVec2
    Dim EndPoint  As tVec2
    Dim P1ANG     As Double
    Dim P2ANG     As Double

    chord = DIFF2(pJoint, Point)
    NtangentDir = Vec2(-tangentDir.Y, tangentDir.X)
    chordDotN1 = DOT2(chord, NtangentDir)

    If (IsEqualEps(chordDotN1, 0)) Then
        MsgBox "If (IsEqualEps(chordDotN1, 0)) Then  - Not Implemented Yet"

        '// straight line
        ''        CCDrawLine CC, Point, pJoint, 5, Color
    Else

        outRadius = LengthSq2(chord) / (2 * chordDotN1)
        outCenter = ADDScaled2(Point, NtangentDir, outRadius)

        StartPoint = DIFF2(Point, outCenter)
        EndPoint = DIFF2(pJoint, outCenter)

        P1ANG = Atan2(StartPoint.X, StartPoint.Y)
        P2ANG = Atan2(EndPoint.X, EndPoint.Y)

        If (StartPoint.X * tangentDir.Y - StartPoint.Y * tangentDir.X) > 0 Then
            If (fromP1) Then
                outAng1 = P1ANG
                outAng2 = P2ANG
            Else
                outAng1 = P2ANG
                outAng2 = P1ANG
            End If
        Else
            If Not (fromP1) Then
                outAng1 = P1ANG
                outAng2 = P2ANG
            Else
                outAng1 = P2ANG
                outAng2 = P1ANG
            End If
        End If
    End If

    outRadius = Abs(outRadius)

End Sub




Public Sub CALC(Optional CustomDistance As Double = 0)


    Dim V         As tVec2
    Dim vMagSqr   As Double
    Dim vDotT1    As Double

    Dim vDotT2    As Double
    Dim T1DotT2   As Double
    Dim denominator As Double

    Dim invlen    As Double

    Dim T         As tVec2
    Dim tMagSqr   As Double
    Dim equalTangents As Boolean
    Dim perpT1    As Boolean

    Dim vDotT     As Double

    Dim discriminant As Double

    Dim angle     As Double
    Dim Center1   As tVec2
    Dim Center2   As tVec2
    Dim Radius    As Double
    Dim cross     As Double

    mTangDir1 = Normalize2(DIFF2(mControlP1, mPoint1))
    mTangDir2 = Normalize2(DIFF2(mControlP2, mPoint2))

    d1 = CustomDistance

    V = DIFF2(mPoint2, mPoint1)
    vMagSqr = LengthSq2(V)

    vDotT1 = DOT2(V, mTangDir1)

    '// if we are using a custom value for d1
    If (CustomDistance) Then

        vDotT2 = DOT2(V, mTangDir2)
        T1DotT2 = DOT2(mTangDir1, mTangDir2)
        denominator = (vDotT2 - d1 * (T1DotT2 - 1))

        If (IsEqualEps(denominator, 0#)) Then
            d1 = CustomDistance
            d2 = MAX_VALUE
            '// the second arc is a semicircle
            mJoint = ADDScaled2(mPoint1, mTangDir1, d1)
            mJoint = ADDScaled2(mJoint, mTangDir2, vDotT2 - d1 * T1DotT2)

        Else

            d2 = (0.5 * vMagSqr - d1 * vDotT1) / denominator

            invlen = 1# / (d1 + d2)

            mJoint = MUL2(DIFF2(mTangDir1, mTangDir2), d1 * d2)
            mJoint = ADDScaled2(mJoint, mPoint2, d1)
            mJoint = ADDScaled2(mJoint, mPoint1, d2)
            mJoint = MUL2(mJoint, invlen)

        End If

        '// else set d1 equal to d2
    Else


        T = Vec2ADD(mTangDir1, mTangDir2)
        tMagSqr = LengthSq2(T)

        equalTangents = IsEqualEps(tMagSqr, 4#)

        perpT1 = IsEqualEps(vDotT1, 0#)
        If (equalTangents And perpT1) Then
            Stop
            '// we have two semicircles
            mJoint = ADDScaled2(mPoint1, V, 0.5)
            d1 = MAX_VALUE
            d2 = MAX_VALUE
            angle = Atan2(V.X, V.Y)
            Center1 = ADDScaled2(mPoint1, V, 0.25)
            Center2 = ADDScaled2(mPoint1, V, 0.75)
            Radius = Sqr(vMagSqr) * 0.25
            cross = V.X * mTangDir1.Y - V.Y * mTangDir1.X
        Else

            vDotT = DOT2(V, T)
            perpT1 = IsEqualEps(vDotT1, 0#)
            If (equalTangents) Then
                d1 = vMagSqr / (4 * vDotT1)
            Else
                denominator = 2 - 2 * DOT2(mTangDir1, mTangDir2)
                discriminant = vDotT * vDotT + denominator * vMagSqr
                d1 = (Sqr(discriminant) - vDotT) / denominator
            End If

            mJoint = MUL2(DIFF2(mTangDir1, mTangDir2), d1)
            mJoint = Vec2ADD(mJoint, mPoint1)
            mJoint = Vec2ADD(mJoint, mPoint2)
            mJoint = MUL2(mJoint, 0.5)

        End If
    End If

    '****************************************************************
    ArcFromEdge mPoint1, mTangDir1, mJoint, True, mCenter1, mRadius1, mA11, mA12
    ArcFromEdge mPoint2, mTangDir2, mJoint, False, mCenter2, mRadius2, mA21, mA22
    '****************************************************************




    '-------------- BOUNDS
    mBound1 = mPoint1
    mBound2 = ADDScaled2(mPoint1, mTangDir1, d1)
    mBound3 = mPoint2
    mBound4 = ADDScaled2(mPoint2, mTangDir2, -d1)

    '-------------------------



    ' | Author miorsoft - reexre
    ' |
    ' V

    AngArc1 = AngleDIFF(mA12, mA11)
    AngArc2 = AngleDIFF(mA22, mA21)
    mLengthArc1 = mRadius1 * Abs(AngArc1)
    mLengthArc2 = mRadius2 * Abs(AngArc2)
    mLengthTot = mLengthArc1 + mLengthArc2

    mClockWise1 = DOT2(Vec2(-mTangDir1.Y, mTangDir1.X), DIFF2(mJoint, SUM2(mPoint1, mTangDir1))) > 0
    mClockWise2 = DOT2(Vec2(-mTangDir2.Y, mTangDir2.X), DIFF2(mJoint, SUM2(mPoint2, mTangDir2))) > 0



End Sub

Public Sub DRAW(CC As cCairoContext, Color As Long, Alpha As Double, Width As Double, Optional Lines As Boolean = False, Optional FullCircles As Boolean = False)


    If Lines Then
        '// draw bounds
        CCDrawLine CC, mPoint1, ADDScaled2(mPoint1, mTangDir1, d1), boundsWidth, boundsColor
        CCDrawLine CC, ADDScaled2(mPoint1, mTangDir1, d1), mJoint, boundsWidth, boundsColor
        CCDrawLine CC, mJoint, ADDScaled2(mPoint2, mTangDir2, -d1), boundsWidth, boundsColor
        CCDrawLine CC, ADDScaled2(mPoint2, mTangDir2, -d1), mPoint2, boundsWidth, boundsColor
    End If
    If FullCircles Then
        With CC
            .SetSourceColor arcColor1, 0.3
            .Arc mCenter1.X, mCenter1.Y, mRadius1
            .Fill
            .SetSourceColor arcColor2, 0.3
            .Arc mCenter2.X, mCenter2.Y, mRadius2
            .Fill
        End With
    End If

    ' BIARC
    With CC
        .SetSourceColor Color, Alpha
        .SetLineWidth Width
        .Arc mCenter1.X, mCenter1.Y, mRadius1, mA11, mA12
        .Stroke
        .Arc mCenter2.X, mCenter2.Y, mRadius2, mA21, mA22
        .Stroke
    End With
End Sub




Public Sub DrawArcsReadyToFillStroke(CC As cCairoContext)
    With CC
        If mClockWise1 Then
            .Arc mCenter1.X, mCenter1.Y, mRadius1, mA11, mA12
        Else
            .ArcNegative mCenter1.X, mCenter1.Y, mRadius1, mA12, mA11
        End If
        If mClockWise2 Then
            .Arc mCenter2.X, mCenter2.Y, mRadius2, mA21, mA22
        Else
            .ArcNegative mCenter2.X, mCenter2.Y, mRadius2, mA22, mA21
        End If
    End With
End Sub
Friend Function InterpolatedPointAt(ByVal T As Double, Optional WArc As Long = 0) As tVec2
' Author miorsoft - reexre
    Dim Perc1#, Perc2#
    Dim iPerc1#, iPerc2#


    Dim Out       As tVec2

    Perc1 = mLengthArc1 / mLengthTot
    Perc2 = mLengthArc2 / mLengthTot
    iPerc1 = 1# / Perc1
    iPerc2 = 1# / Perc2

    If T < Perc1 Then                            ' First arc
        If mClockWise1 Then
            Out.X = mCenter1.X + Cos(mA12 - AngArc1 * (1# - T * iPerc1)) * mRadius1
            Out.Y = mCenter1.Y + Sin(mA12 - AngArc1 * (1# - T * iPerc1)) * mRadius1
        Else
            Out.X = mCenter1.X + Cos(mA11 + AngArc1 * (1# - T * iPerc1)) * mRadius1
            Out.Y = mCenter1.Y + Sin(mA11 + AngArc1 * (1# - T * iPerc1)) * mRadius1
        End If
        WArc = 1
    Else                                         ' Second arc
        T = T - Perc1                            '<---
        If mClockWise2 Then
            Out.X = mCenter2.X + Cos(mA22 - AngArc2 * (1# - T * iPerc2)) * mRadius2
            Out.Y = mCenter2.Y + Sin(mA22 - AngArc2 * (1# - T * iPerc2)) * mRadius2
        Else
            Out.X = mCenter2.X + Cos(mA21 + AngArc2 * (1# - T * iPerc2)) * mRadius2
            Out.Y = mCenter2.Y + Sin(mA21 + AngArc2 * (1# - T * iPerc2)) * mRadius2
        End If
        WArc = 2
    End If

    InterpolatedPointAt = Out

End Function

'Friend Function CircleDistanceOLD(CircleCenter As tVec2, CircleRadius As Double, OutArc1or2 As Long, Optional LeftSide As Boolean = True) As Double
'    Dim D         As Double
'    Dim A         As Double
'    Dim B         As Boolean
'
'        A = Atan2(CircleCenter.X - mCenter1.X, CircleCenter.Y - mCenter1.Y)
'        If IsAngBetween(A, mA11, mA12) Then
'            '        CC.SetSourceColor vbRed
'            '        CC.Arc mCenter1.X, mCenter1.Y, mRadius1: CC.Stroke
'            CircleDistanceOLD = Length2(DIFF2(Center1, CircleCenter)) - mRadius1    '- CircleRadius
'            If LeftSide Then
'                CircleDistanceOLD = CircleDistanceOLD - CircleRadius
'            Else
'                CircleDistanceOLD = -CircleDistanceOLD - CircleRadius
'            End If
'            If Not (mClockWise1) Then CircleDistanceOLD = -CircleDistanceOLD - CircleRadius * 2
'            'If CircleDistanceOLD < 0 Then Stop
'            OutArc1or2 = 1
'        Else
'            A = Atan2(CircleCenter.X - mCenter2.X, CircleCenter.Y - mCenter2.Y)
'            If IsAngBetween(A, mA21, mA22) Then
'                '        CC.SetSourceColor vbRed
'                '        CC.Arc mCenter2.X, mCenter2.Y, mRadius2: CC.Stroke
'                CircleDistanceOLD = Length2(DIFF2(Center2, CircleCenter)) - mRadius2    '- CircleRadius
'                If LeftSide Then
'                    CircleDistanceOLD = CircleDistanceOLD - CircleRadius
'                Else
'                    CircleDistanceOLD = -CircleDistanceOLD - CircleRadius
'                End If
'                If Not (mClockWise2) Then CircleDistanceOLD = -CircleDistanceOLD - CircleRadius * 2
'                'If CircleDistanceOLD < 0 Then Stop
'                OutArc1or2 = 2
'            End If
'        End If
'End Function
Friend Function CircleDistance(CircleCenter As tVec2, CircleRadius As Double, OutArc1or2 As Long, Optional LeftSide As Boolean = True) As Double
' returns signed distance of a circle

    Dim DistFromArc As Double
    If IsAngBetween(Atan2(CircleCenter.X - mCenter1.X, CircleCenter.Y - mCenter1.Y), _
                    mA11, mA12) Then
        DistFromArc = Length2(DIFF2(Center1, CircleCenter)) - mRadius1
        If Not (LeftSide) Then DistFromArc = -DistFromArc
        If Not (mClockWise1) Then DistFromArc = -DistFromArc
        CircleDistance = DistFromArc - CircleRadius
        OutArc1or2 = 1
    ElseIf IsAngBetween(Atan2(CircleCenter.X - mCenter2.X, CircleCenter.Y - mCenter2.Y), _
                        mA21, mA22) Then
        DistFromArc = Length2(DIFF2(Center2, CircleCenter)) - mRadius2
        If Not (LeftSide) Then DistFromArc = -DistFromArc
        If Not (mClockWise2) Then DistFromArc = -DistFromArc
        CircleDistance = DistFromArc - CircleRadius
        OutArc1or2 = 2
    End If
    ' Brute Force an error sometimes occour (Must investigate)
    ' - Very high negative CircleDistance -
    ' It works, but Circle Speed cannot be bigger than Radius at each timestep
    If CircleDistance < -CircleRadius Then CircleDistance = -CircleDistance
End Function

Friend Function PointDistance(CircleCenter As tVec2, OutArc1or2 As Long) As Double
' Circle Speed cannot be bigger than Radius at each timestep
    If IsAngBetween(Atan2(CircleCenter.X - mCenter1.X, CircleCenter.Y - mCenter1.Y), _
                    mA11, mA12) Then
        PointDistance = Abs(Length2(DIFF2(Center1, CircleCenter)) - mRadius1)
        OutArc1or2 = 1
    ElseIf IsAngBetween(Atan2(CircleCenter.X - mCenter2.X, CircleCenter.Y - mCenter2.Y), _
                        mA21, mA22) Then
        PointDistance = Abs(Length2(DIFF2(Center2, CircleCenter)) - mRadius2)
        OutArc1or2 = 2
    End If
End Function



Friend Sub SetImage(FileName As String, TargetSrf As cCairoSurface, Optional ImageScaleX As Double = 1, Optional ImageScaleY As Double = 1)

    Set SourceSRF = Cairo.ImageList.AddImage("sourceImage", FileName)
    Set SRF2 = TargetSrf

    pW = SRF2.Width - 1
    pH = SRF2.Height - 1

    ImaScaleX = ImageScaleX
    ImaScaleY = ImageScaleY

    InvScaleX = 1 / ImaScaleX
    InvScaleY = 1 / ImaScaleY
End Sub



Friend Sub DrawIMAGE(CC As cCairoContext, ClearScreen As Boolean, Optional Length As Double = 0)
'    Dim ArrSource() As Long
'    Dim ArrSrf2() As Long
    Dim ArrSource() As Byte
    Dim ArrSrf2() As Byte

    Dim H2&, H22#
    Dim DX#, DY#
    Dim X&, Y&
    Dim xF&, xT&
    Dim yF&, yT&
    Dim D#
    Dim W&, H&
    Dim A#
    Dim GetX#, GetY#
    Dim bA        As tVec2
    Dim bB        As tVec2
    Dim SgnClockWise1 As Double
    Dim SgnClockWise2 As Double
    Dim X4        As Long
    Dim GetX4     As Long
    Dim Alpha#
    Dim Klen#


    Dim dmin1#, dmax1#
    Dim dmin2#, dmax2#

    If ClearScreen Then CC.SetSourceRGB 0.33, 0.33, 0.33: CC.Paint

    '    SourceSRF.BindToArrayLong ArrSource
    '    SRF2.BindToArrayLong ArrSrf2
    SourceSRF.BindToArray ArrSource
    SRF2.BindToArray ArrSrf2

    W = UBound(ArrSource, 1) \ 4
    H = UBound(ArrSource, 2)

    H2 = H / 2
    H22 = H2 * H2
    If Length Then Klen = Length / mLengthTot Else: Klen = 1

    GetBounds bA, bB

    xF = bA.X - H2 * ImaScaleY
    yF = bA.Y - H2 * ImaScaleY
    xT = bB.X + H2 * ImaScaleY
    yT = bB.Y + H2 * ImaScaleY

    If xF < 0 Then xF = 0
    If yF < 0 Then yF = 0
    If xT > pW Then xT = pW
    If yT > pH Then yT = pH

    dmin1 = mRadius1 - H2 * ImaScaleY
    dmin1 = Sgn(dmin1) * dmin1 * dmin1
    dmax1 = mRadius1 + H2 * ImaScaleY
    dmax1 = dmax1 * dmax1

    dmin2 = mRadius2 - H2 * ImaScaleY
    dmin2 = Sgn(dmin2) * dmin2 * dmin2
    dmax2 = mRadius2 + H2 * ImaScaleY
    dmax2 = dmax2 * dmax2

    If mClockWise1 Then SgnClockWise1 = 1 Else: SgnClockWise1 = -1
    If mClockWise2 Then SgnClockWise2 = 1 Else: SgnClockWise2 = -1

    For X = xF To xT
        X4 = X * 4
        For Y = yF To yT
            DX = X - mCenter1.X
            DY = Y - mCenter1.Y
            'D = Sqr(DX * DX + DY * DY) - mRadius1
            'If Abs(D) < H2 Then
            D = DX * DX + DY * DY
            If D > dmin1 And D < dmax1 Then
                A = Atan2(DX, DY)
                If IsAngBetween(A, mA11, mA12) Then
                    D = Sqr(D) - mRadius1
                    GetY = H2 - D * SgnClockWise1 * InvScaleY
                    If mClockWise1 Then
                        GetX = AngleDIFF(A, mA11) * mRadius1
                    Else
                        GetX = -AngleDIFF(A, mA12) * mRadius1
                    End If
                    GetX = (GetX * InvScaleX * Klen) Mod W
                    GetY = GetY Mod H
                    '                    ArrSrf2(X, Y) = ArrSource(GetX, GetY)
                    GetX4 = GetX * 4&
                    '                            ArrSrf2(X4, Y) = ArrSource(GetX4, GetY)

                    Alpha = ArrSource(GetX4 + 3&, GetY) * 3.92156862745098E-03    '/255
                    If Alpha Then
                        ColorMixByAlpha ArrSource(GetX4 + 2&, GetY), ArrSource(GetX4 + 1&, GetY), ArrSource(GetX4, GetY), Alpha, _
                                        ArrSrf2(X4 + 2&, Y), ArrSrf2(X4 + 1&, Y), ArrSrf2(X4, Y)
                    End If
                End If
            End If
            DX = X - mCenter2.X
            DY = Y - mCenter2.Y
            'D = Sqr(DX * DX + DY * DY) - mRadius2
            'If Abs(D) < H2 Then
            D = DX * DX + DY * DY
            If D > dmin2 And D < dmax2 Then
                A = Atan2(DX, DY)
                If IsAngBetween(A, mA21, mA22) Then
                    D = Sqr(D) - mRadius2
                    GetY = H2 - D * SgnClockWise2 * InvScaleY

                    If mClockWise2 Then
                        GetX = AngleDIFF(A, mA21) * mRadius2
                    Else
                        GetX = -AngleDIFF(A, mA22) * mRadius2
                    End If
                    GetX = GetX + mLengthArc1
                    GetX = (GetX * InvScaleX * Klen) Mod W
                    GetY = GetY Mod H
                    '                    ArrSrf2(X, Y) = ArrSource(GetX, GetY)
                    GetX4 = GetX * 4&
                    '                                                ArrSrf2(X4, Y) = ArrSource(GetX4, GetY)

                    Alpha = ArrSource(GetX4 + 3&, GetY) * 3.92156862745098E-03
                    If Alpha Then
                        ColorMixByAlpha ArrSource(GetX4 + 2&, GetY), ArrSource(GetX4 + 1&, GetY), ArrSource(GetX4, GetY), Alpha, _
                                        ArrSrf2(X4 + 2&, Y), ArrSrf2(X4 + 1&, Y), ArrSrf2(X4, Y)
                    End If
                End If
            End If

        Next
    Next




    '    SourceSRF.ReleaseArrayLong ArrSource
    '    SRF2.ReleaseArrayLong ArrSrf2

    SourceSRF.ReleaseArray ArrSource
    SRF2.ReleaseArray ArrSrf2



End Sub
Private Sub ColorMixByAlpha(iR As Byte, iG As Byte, iB As Byte, iAlpha As Double, outR As Byte, outG As Byte, outB As Byte)
    outR = iR * (iAlpha) + outR * (1# - iAlpha)
    outG = iG * (iAlpha) + outG * (1# - iAlpha)
    outB = iB * (iAlpha) + outB * (1# - iAlpha)
End Sub

