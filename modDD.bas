Attribute VB_Name = "modDD"
' =================================================================================
' getAngle - CALCULATE THE ANGLE BASED ON DP1 AND DP2 RETURNS ANGLE IN RADIANS
'
' 12-11-01 RAMSEY SYSTEMS, INC.
' =================================================================================
Public Function getAngle(startPt As Point3d, endPt As Point3d, flipped As Boolean) As Double
    Dim vec1 As Point3d
    Dim vec2 As Point3d
    Dim angle As Double
    
    flipped = False
    
    If (startPt.X <= endPt.X And startPt.Y < endPt.Y) Or (startPt.X > endPt.X And startPt.Y < endPt.Y) Then
        vec1.X = 1
    Else
        vec1.X = -1
        flipped = True
    End If
    
    vec1.Y = 0
    vec1.Z = 0
        
    vec2.X = endPt.X - startPt.X
    vec2.Y = endPt.Y - startPt.Y
    vec2.Z = endPt.Z - startPt.Z
    
    angle = Point3dAngleBetweenVectors(vec1, vec2)
    
    getAngle = angle
End Function
' =================================================================================
' ic_handleMirrorTrans - Mirror Cell
'
' 12-11-01 RAMSEY SYSTEMS, INC.
' =================================================================================
Public Function ic_handleMirrorTrans(angle As Double, origin As Point3d, sx As Double, sy As Double, view As view) As Transform3d
    Dim et As Transform3d, cMtx As Matrix3d, cInvMtx As Matrix3d, rInvMtx As Matrix3d, mMtx As Matrix3d, aMtx As Matrix3d, rMtx As Matrix3d

    cMtx = view.Rotation
    rMtx = Matrix3dFromAxisAndRotationAngle(2, angle)
    cInvMtx = Matrix3dInverse(cMtx)
    rInvMtx = Matrix3dInverse(rMtx)
    mMtx = Matrix3dFromScaleFactors(sx, sy, 1)
    aMtx = Matrix3dFromMatrix3dTimesMatrix3dTimesMatrix3d(rMtx, mMtx, rInvMtx)
    aMtx = Matrix3dFromMatrix3dTimesMatrix3dTimesMatrix3d(cMtx, aMtx, cInvMtx)
    et = Transform3dFromMatrix3dAndFixedPoint3d(aMtx, origin)
    ic_handleMirrorTrans = et
    
End Function

' =================================================================================
' pf_mirrorFitting - MIRROR CELL ABOUT AXIS
'
' Ramsey Systems, Inc. March 18, 2002
' =================================================================================
Public Function pf_mirrorFitting(oCell As CellElement, mPt As Point3d, origin As Point3d, view As view) As Integer
        Dim isFlipped As Boolean, mAngle As Double, Eltrans As Transform3d
        
        mAngle = getAngle(origin, mPt, isFlipped)
        'Eltrans = ic_handleMirrorTrans(ActiveSettings.angle, origin, 1, 1, view) ' - - - no mirror
        
       If Not isFlipped And mAngle >= 0 And mAngle <= Pi / 2 Then
            Eltrans = ic_handleMirrorTrans(ActiveSettings.angle, origin, 1, 1, view) ' - - - no mirror
        ElseIf Not isFlipped And mAngle > Pi / 2 And mAngle <= Pi Then
            Eltrans = ic_handleMirrorTrans(ActiveSettings.angle, origin, -1, 1, view) ' - - - mirror across yz-plane
        ElseIf isFlipped And mAngle > 0 And mAngle <= Pi / 2 Then
            Eltrans = ic_handleMirrorTrans(ActiveSettings.angle, origin, -1, -1, view)  ' - - - mirror across xz-plane and yz-plane
        ElseIf isFlipped And mAngle > Pi / 2 And mAngle <= Pi Then
            Eltrans = ic_handleMirrorTrans(ActiveSettings.angle, origin, 1, -1, view)  ' - - - mirror across xz-plane
        End If
        
        oCell.Transform Eltrans
        
        pf_mirrorFitting = 0
End Function
Sub drawDoor()
    Load formDD
    formDD.Show
End Sub
