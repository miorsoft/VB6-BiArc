Attribute VB_Name = "mBiARC2"

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


Private Const boundsWidth As Double = 0.5
Private Const boundsColor As Long = vbGreen

Private Const arcWidth As Double = 2

Private Const tangentWidth As Double = 1
Private Const pointColor1 As Long = vbGreen
Private Const pointColor2 As Long = vbRed

Private Const arcColor1 As Long = vbGreen
Private Const arcColor2 As Long = vbRed



''''Function DrawLine(context, p1, p2, width, color)
''''{
''''    context.beginPath();
''''    context.moveTo(p1.x, p1.y);
''''    context.lineTo(p2.x, p2.y);
''''    context.lineWidth = width;
''''    context.strokeStyle = color;
''''    context.stroke();
''''}
Public Sub CCDrawLine(p1 As tVec2, p2 As tVec2, Lwidth As Double, Color As Long)
    With CC
        .SetLineWidth Lwidth
        .SetSourceColor Color
        .MoveTo p1.x, p1.Y
        .LineTo p2.x, p2.Y
        .Stroke
    End With
End Sub
'''Function DrawArcFromEdge(context, p1, t1, p2, arcWidth, color, fromP1)
'''{
'''    var chord = Vec2Sub(p2, p1);
'''    var n1 = new Vec2(-t1.y, t1.x);
'''    var chordDotN1 = Vec2Dot(chord, n1);
'''
'''    if (IsEqualEps(chordDotN1,0))
'''    {
'''        // straight line
'''        DrawLine(context, p1, p2, arcWidth, color);
'''    }
'''    Else
'''    {
'''        var radius = Vec2_MagSqr(chord) / (2*chordDotN1);
'''        var center = Vec2_AddScaled(p1, n1, radius);
'''
'''        var p1Offset = Vec2Sub(p1, center);
'''        var p2Offset = Vec2Sub(p2, center);
'''
'''        var p1Ang1 = Math.atan2(p1Offset.y,p1Offset.x);
'''        var p2Ang1 = Math.atan2(p2Offset.y,p2Offset.x);
'''        if ( p1Offset.x*t1.y - p1Offset.y*t1.x > 0)
'''            DrawArc(context, center, Math.abs(radius), p1Ang1, p2Ang1, arcWidth, color, !fromP1);
'''        Else
'''            DrawArc(context, center, Math.abs(radius), p1Ang1, p2Ang1, arcWidth, color, fromP1);
'''
'''        context.globalAlpha = 0.05;
'''        DrawCircle(context, center, Math.abs(radius), color);
'''        context.globalAlpha = 1.0;
'''    }
'''}
Public Sub DrawArcFromEdge(p1 As tVec2, t1 As tVec2, p2 As tVec2, _
                           arcWidth As Double, Color As Long, fromP1 As Boolean)
    Dim chord     As tVec2
    Dim N1        As tVec2
    Dim chordDotN1 As Double
    Dim radius    As Double
    Dim Center    As tVec2
    Dim p1Offset  As tVec2
    Dim p2Offset  As tVec2
    Dim p1Ang1    As Double
    Dim p2Ang1    As Double
    Dim TA        As Double

    chord = Vec2SUB(p2, p1)
    N1 = Vec2(-t1.Y, t1.x)
    chordDotN1 = Vec2DOT(chord, N1)

    If (IsEqualEps(chordDotN1, 0)) Then

        '// straight line
        CCDrawLine p1, p2, arcWidth, Color

    Else

        radius = Vec2LengthSq(chord) / (2 * chordDotN1)
        Center = Vec2ADDScaled(p1, N1, radius)

        p1Offset = Vec2SUB(p1, Center)
        p2Offset = Vec2SUB(p2, Center)

        p1Ang1 = Atan2(p1Offset.x, p1Offset.Y)
        p2Ang1 = Atan2(p2Offset.x, p2Offset.Y)


        If (p1Offset.x * t1.Y - p1Offset.Y * t1.x > 0) Then
            '    DrawArc(context, center, Math.abs(radius), p1Ang1, p2Ang1, arcWidth, color, !fromP1);
            '''''            With CC
            '''''                .SetSourceColor cOLOR
            '''''                .SetLineWidth arcWidth
            '''''                If (fromP1) Then TA = p1Ang1: p1Ang1 = p2Ang1: p2Ang1 = TA
            '''''                .Arc Center.x, Center.Y, Abs(radius), p2Ang1, p1Ang1
            '''''                .Stroke
            '''''            End With
            'If color = vbGreen Then
            'Debug.Print "side1 " & fromP1
        Else
            '    DrawArc(context, center, Math.abs(radius), p1Ang1, p2Ang1, arcWidth, color, fromP1);
            '''''            With CC
            '''''                .SetSourceColor cOLOR
            '''''                .SetLineWidth arcWidth
            '''''                If Not (fromP1) Then TA = p1Ang1: p1Ang1 = p2Ang1: p2Ang1 = TA
            '''''                .Arc Center.x, Center.Y, Abs(radius), p2Ang1, p1Ang1
            '''''                .Stroke
            '''''            End With
            'If color = vbGreen Then
            'Debug.Print "side2 " & fromP1
        End If

        With CC
            .SetSourceColor Color, 0.3
            .Arc Center.x, Center.Y, Abs(radius)
            .Fill
            .SetSourceColor vbYellow
            .Arc p2.x, p2.Y, 3
            .Fill
        End With


        '        context.globalAlpha = 0.05;
        '        DrawCircle(context, center, Math.abs(radius), color);
        '        context.globalAlpha = 1.0;
    End If


End Sub




Public Sub CalcAndDRAWBiARC(p1 As tVec2, p2 As tVec2, _
                            NT1 As tVec2, NT2 As tVec2, _
                            CustomDistance As Double)


    Dim d1        As Double
    Dim d2        As Double

    Dim v         As tVec2
    Dim vMagSqr   As Double
    Dim vDotT1    As Double

    Dim vDotT2    As Double
    Dim t1DotT2   As Double
    Dim denominator As Double

    Dim joint     As tVec2
    Dim invlen    As Double

    Dim t         As tVec2
    Dim tMagSqr   As Double
    Dim equalTangents As Boolean
    Dim perpT1    As Boolean

    Dim vDotT     As Double

    Dim discriminant As Double


    Dim angle     As Double
    Dim center1   As tVec2
    Dim center2   As tVec2
    Dim radius    As Double
    Dim cross     As Double


    d1 = CustomDistance

    With CC: .SetSourceColor 0: .Paint: End With


    v = Vec2SUB(p2, p1)
    vMagSqr = Vec2LengthSq(v)

    vDotT1 = Vec2DOT(v, NT1)


    '// if we are using a custom value for d1
    If (CustomDistance) Then

        vDotT2 = Vec2DOT(v, NT2)
        t1DotT2 = Vec2DOT(NT1, NT2)
        denominator = (vDotT2 - d1 * (t1DotT2 - 1))

        If (IsEqualEps(denominator, 0#)) Then

            'document.getElementById('d1').value = d1;
            'document.getElementById('d2').value = 'Infinity';
            d1 = CustomDistance
            d2 = MAX_VALUE

            '// the second arc is a semicircle
            joint = Vec2ADDScaled(p1, NT1, d1)
            joint = Vec2ADDScaled(joint, NT2, vDotT2 - d1 * t1DotT2)

            '// draw bounds
            CCDrawLine p1, Vec2ADDScaled(p1, NT1, d1), boundsWidth, boundsColor
            CCDrawLine Vec2ADDScaled(p1, NT1, d1), joint, boundsWidth, boundsColor

            '// draw arcs
            DrawArcFromEdge p1, NT1, joint, arcWidth, arcColor1, True
            DrawArcFromEdge p2, NT2, joint, arcWidth, arcColor2, False

        Else

            d2 = (0.5 * vMagSqr - d1 * vDotT1) / denominator

            invlen = 1# / (d1 + d2)

            joint = Vec2MUL(Vec2SUB(NT1, NT2), d1 * d2)
            joint = Vec2ADDScaled(joint, p2, d1)
            joint = Vec2ADDScaled(joint, p1, d2)
            joint = Vec2MUL(joint, invlen)

            'document.getElementById('d1').value = d1;
            'document.getElementById('d2').value = d2;

            '// draw bounds
            CCDrawLine p1, Vec2ADDScaled(p1, NT1, d1), boundsWidth, boundsColor
            CCDrawLine Vec2ADDScaled(p1, NT1, d1), joint, boundsWidth, boundsColor
            CCDrawLine joint, Vec2ADDScaled(p2, NT2, -d2), boundsWidth, boundsColor
            CCDrawLine Vec2ADDScaled(p2, NT2, -d2), p2, boundsWidth, boundsColor

            '// draw arcs
            DrawArcFromEdge p1, NT1, joint, arcWidth * 2, arcColor1, True
            DrawArcFromEdge p2, NT2, joint, arcWidth, arcColor2, False
        End If

        '// else set d1 equal to d2
    Else


        t = Vec2ADD(NT1, NT2)
        tMagSqr = Vec2LengthSq(t)

        equalTangents = IsEqualEps(tMagSqr, 4#)

        perpT1 = IsEqualEps(vDotT1, 0#)
        If (equalTangents And perpT1) Then
            '// we have two semicircles
            joint = Vec2ADDScaled(p1, v, 0.5)

            '                document.getElementById('d1').value = 'Infinity';
            '                document.getElementById('d2').value = 'Infinity';
            d1 = MAX_VALUE
            d2 = MAX_VALUE

            '// draw arcs


            angle = Atan2(v.x, v.Y)
            center1 = Vec2ADDScaled(p1, v, 0.25)
            center2 = Vec2ADDScaled(p1, v, 0.75)
            radius = Sqr(vMagSqr) * 0.25
            cross = v.x * NT1.Y - v.Y * NT1.x

            'DrawArc(context, center1, radius, angle, angle+Math.PI, arcWidth, arcColor1, cross < 0);
            'DrawArc(context, center2, radius, angle, angle+Math.PI, arcWidth, arcColor2, cross > 0);

            'context.globalAlpha = 0.05;
            'DrawCircle(context, center1, radius, arcColor1);
            'DrawCircle(context, center2, radius, arcColor2);
            'context.globalAlpha = 1.0;

        Else


            'Stop
            vDotT = Vec2DOT(v, t)

            perpT1 = IsEqualEps(vDotT1, 0#)

            If (equalTangents) Then

                d1 = vMagSqr / (4 * vDotT1)

            Else

                denominator = 2 - 2 * Vec2DOT(NT1, NT2)
                discriminant = vDotT * vDotT + denominator * vMagSqr
                d1 = (Sqr(discriminant) - vDotT) / denominator
            End If

            joint = Vec2MUL(Vec2SUB(NT1, NT2), d1)
            joint = Vec2ADD(joint, p1)
            joint = Vec2ADD(joint, p2)
            joint = Vec2MUL(joint, 0.5)

            'document.getElementById('d1').value = d1;
            'document.getElementById('d2').value = d1;

            '// draw bounds
            CCDrawLine p1, Vec2ADDScaled(p1, NT1, d1), boundsWidth, boundsColor
            CCDrawLine Vec2ADDScaled(p1, NT1, d1), joint, boundsWidth, boundsColor
            CCDrawLine joint, Vec2ADDScaled(p2, NT2, -d1), boundsWidth, boundsColor
            CCDrawLine Vec2ADDScaled(p2, NT2, -d1), p2, boundsWidth, boundsColor

            '// draw arcs
            DrawArcFromEdge p1, NT1, joint, arcWidth, arcColor1, True
            DrawArcFromEdge p2, NT2, joint, arcWidth, arcColor2, False
        End If
    End If




    '    // draw points
    '    DrawCircle(context, p1, c_Point_Radius, pointColor1);
    '    DrawCircle(context, p2, c_Point_Radius, pointColor2);

    '    // draw tangents
    '    DrawLine(context, p1, Vec2AddScaled(p1, NT1, c_Tangent_Length), tangentWidth, pointColor1);
    '    DrawLine(context, p2, Vec2AddScaled(p2, NT2, c_Tangent_Length), tangentWidth, pointColor2);
    '---------------------------------------------------------------------------------------------



    '''''    With CC
    '''''        .SetSourceColor arcColor1
    '''''        .Arc p1.x, p1.Y, 5
    '''''        .Fill
    '''''        .MoveTo p1.x, p1.Y
    '''''        .LineTo t1.x, t1.Y
    '''''        .Stroke
    '''''
    '''''        .SetSourceColor arcColor2
    '''''        .Arc p2.x, p2.Y, 5
    '''''        .Fill
    '''''        .MoveTo p2.x, p2.Y
    '''''        .LineTo t2.x, t2.Y
    '''''        .Stroke
    '''''    End With

    '''''    cPTS.Draw CC
    '    srf.DrawToDC PicHDC


End Sub




Public Sub CalcBiARC(p1 As tVec2, p2 As tVec2, _
                     NT1 As tVec2, NT2 As tVec2, _
                     ByRef OutCenter1 As tVec2, _
                     ByRef outRadius1 As Double, _
                     ByRef OutAng11 As Double, _
                     ByRef OutAng12 As Double, _
                     ByRef OutCenter2 As tVec2, _
                     ByRef outRadius2 As Double, _
                     ByRef OutAng21 As Double, _
                     ByRef OutAng22 As Double, _
                     Optional ByVal CustomDistance = 0)


    Dim d1        As Double
    Dim d2        As Double

    Dim v         As tVec2
    Dim vMagSqr   As Double
    Dim vDotT1    As Double

    Dim vDotT2    As Double
    Dim t1DotT2   As Double
    Dim denominator As Double

    Dim joint     As tVec2
    Dim invlen    As Double

    Dim t         As tVec2
    Dim tMagSqr   As Double
    Dim equalTangents As Boolean
    Dim perpT1    As Boolean

    Dim vDotT     As Double

    Dim discriminant As Double


    Dim angle     As Double
    Dim center1   As tVec2
    Dim center2   As tVec2
    Dim radius    As Double
    Dim cross     As Double


    d1 = CustomDistance

    '    With CC
    '        .SetSourceColor 0
    '        .Paint
    '    End With


    v = Vec2SUB(p2, p1)
    vMagSqr = Vec2LengthSq(v)

    vDotT1 = Vec2DOT(v, NT1)


    '// if we are using a custom value for d1
    If (CustomDistance) Then

        vDotT2 = Vec2DOT(v, NT2)
        t1DotT2 = Vec2DOT(NT1, NT2)
        denominator = (vDotT2 - d1 * (t1DotT2 - 1))

        If (IsEqualEps(denominator, 0#)) Then
            d1 = CustomDistance
            d2 = MAX_VALUE
            '// the second arc is a semicircle
            joint = Vec2ADDScaled(p1, NT1, d1)
            joint = Vec2ADDScaled(joint, NT2, vDotT2 - d1 * t1DotT2)

        Else

            d2 = (0.5 * vMagSqr - d1 * vDotT1) / denominator

            invlen = 1# / (d1 + d2)

            joint = Vec2MUL(Vec2SUB(NT1, NT2), d1 * d2)
            joint = Vec2ADDScaled(joint, p2, d1)
            joint = Vec2ADDScaled(joint, p1, d2)
            joint = Vec2MUL(joint, invlen)

        End If

        '// else set d1 equal to d2
    Else


        t = Vec2ADD(NT1, NT2)
        tMagSqr = Vec2LengthSq(t)

        equalTangents = IsEqualEps(tMagSqr, 4#)

        perpT1 = IsEqualEps(vDotT1, 0#)
        If (equalTangents And perpT1) Then
            '// we have two semicircles
            joint = Vec2ADDScaled(p1, v, 0.5)

            d1 = MAX_VALUE
            d2 = MAX_VALUE
            '// draw arcs
            angle = Atan2(v.x, v.Y)
            center1 = Vec2ADDScaled(p1, v, 0.25)
            center2 = Vec2ADDScaled(p1, v, 0.75)
            radius = Sqr(vMagSqr) * 0.25
            cross = v.x * NT1.Y - v.Y * NT1.x

            'DrawArc(context, center1, radius, angle, angle+Math.PI, arcWidth, arcColor1, cross < 0);
            'DrawArc(context, center2, radius, angle, angle+Math.PI, arcWidth, arcColor2, cross > 0);
            'context.globalAlpha = 0.05;
            'DrawCircle(context, center1, radius, arcColor1);
            'DrawCircle(context, center2, radius, arcColor2);
            'context.globalAlpha = 1.0;

        Else


            'Stop
            vDotT = Vec2DOT(v, t)

            perpT1 = IsEqualEps(vDotT1, 0#)

            If (equalTangents) Then

                d1 = vMagSqr / (4 * vDotT1)

            Else

                denominator = 2 - 2 * Vec2DOT(NT1, NT2)
                discriminant = vDotT * vDotT + denominator * vMagSqr
                d1 = (Sqr(discriminant) - vDotT) / denominator
            End If

            joint = Vec2MUL(Vec2SUB(NT1, NT2), d1)
            joint = Vec2ADD(joint, p1)
            joint = Vec2ADD(joint, p2)
            joint = Vec2MUL(joint, 0.5)


        End If
    End If


    '**********************************************************************
    '**********************************************************************
    '**********************************************************************
    ArcFromEdge p1, NT1, joint, True, OutCenter1, outRadius1, OutAng11, OutAng12
    ArcFromEdge p2, NT2, joint, False, OutCenter2, outRadius2, OutAng21, OutAng22

    '****************************************************************
    '****************************************************************
    '****************************************************************


End Sub

Private Sub ArcFromEdge(p1 As tVec2, t1 As tVec2, p2Joint As tVec2, fromP1 As Boolean, _
                        ByRef Center As tVec2, _
                        ByRef radius As Double, _
                        ByRef Ang1 As Double, _
                        ByRef Ang2 As Double)

    Dim chord     As tVec2
    Dim N1        As tVec2
    Dim chordDotN1 As Double
    Dim p1Offset  As tVec2
    Dim p2Offset  As tVec2
    Dim P1ANG     As Double
    Dim P2ANG     As Double

    chord = Vec2SUB(p2Joint, p1)
    N1 = Vec2(-t1.Y, t1.x)
    chordDotN1 = Vec2DOT(chord, N1)

    If (IsEqualEps(chordDotN1, 0)) Then
        '// straight line
        '     CCDrawLine p1, p2Joint, arcWidth, color
    Else

        radius = Vec2LengthSq(chord) / (2 * chordDotN1)
        Center = Vec2ADDScaled(p1, N1, radius)

        p1Offset = Vec2SUB(p1, Center)
        p2Offset = Vec2SUB(p2Joint, Center)

        P1ANG = Atan2(p1Offset.x, p1Offset.Y)
        P2ANG = Atan2(p2Offset.x, p2Offset.Y)

        If (p1Offset.x * t1.Y - p1Offset.Y * t1.x > 0) Then
            '    DrawArc(context, center, Math.abs(radius), P1ANG, P2ANG, arcWidth, color, !fromP1);
            ' If (fromP1) Then TA = P1ANG: P1ANG = P2ANG: P2ANG = TA
            If (fromP1) Then
                Ang1 = P1ANG
                Ang2 = P2ANG
            Else
                Ang1 = P2ANG
                Ang2 = P1ANG
            End If
        Else
            '    DrawArc(context, center, Math.abs(radius), P1ANG, P2ANG, arcWidth, color, fromP1);
            'If Not (fromP1) Then TA = P1ANG: P1ANG = P2ANG: P2ANG = TA
            '.ARC Center.X, Center.Y, Abs(Radius), P2ANG, P1ANG
            If Not (fromP1) Then
                Ang1 = P1ANG
                Ang2 = P2ANG
            Else
                Ang1 = P2ANG
                Ang2 = P1ANG
            End If
        End If
    End If

    radius = Abs(radius)

End Sub


