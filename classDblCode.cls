VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "classDblCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IPrimitiveCommandEvents

Dim ptNum As Integer
Dim ptP As Point3d

Private Sub IPrimitiveCommandEvents_Cleanup()
    
End Sub
' ===============================================================================
' DATAPOINT EVENT HANDLER.
'
' Ramsey Systems, Inc. March 15, 2002
' FREE ROTATE - March 26, 2002 - requires rPt (3rd point) for rotation.
' ===============================================================================
Private Sub IPrimitiveCommandEvents_DataPoint(point As Point3d, ByVal view As view)
    ' NO MIRROR AND NO FREE ROTATE
    If Not formDD.chkMirror And Not formDD.chkRotate Then
        pf_dynDrawDoor point, point, point, view, msdDrawingModeNormal
        CommandState.StartPrimitive Me
    End If
    
    ' MIRROR ONLY or FREE ROTATE ONLY
    If (Not formDD.chkMirror And formDD.chkRotate) Or (formDD.chkMirror And Not formDD.chkRotate) Then
        If ptNum = 0 Then
            ptNum = 1
            ptP = point
        Else
            pf_dynDrawDoor point, ptP, point, view, msdDrawingModeNormal
            CommandState.StartPrimitive Me
        End If
    End If

    ' MIRROR AND FREE ROTATE BOTH
    If formDD.chkMirror And formDD.chkRotate Then
        If ptNum = 0 Then
            ptNum = 1
            ptP = point
        ElseIf ptNum = 1 Then
            ptNum = 2
        Else
            pf_dynDrawDoor point, ptP, point, view, msdDrawingModeNormal
            CommandState.StartPrimitive Me
        End If
    End If
End Sub
' ===============================================================================
' DYNAMIC HANDLER. - CALLS ROUTINE
'
' Ramsey Systems, Inc. March 15, 2002
' FREE ROTATE - March 26, 2002
' ===============================================================================
Private Sub IPrimitiveCommandEvents_Dynamics(point As Point3d, ByVal view As view, ByVal DrawMode As MsdDrawingMode)
    ' NO MIRROR AND NO FREE ROTATE
    If Not formDD.chkMirror And Not formDD.chkRotate Then
        pf_dynDrawDoor point, point, point, view, DrawMode
    End If
    
    ' MIRROR ONLY or FREE ROTATE ONLY
    If (Not formDD.chkMirror And formDD.chkRotate) Or (formDD.chkMirror And Not formDD.chkRotate) Then
        If ptNum = 0 Then
            pf_dynDrawDoor point, point, point, view, DrawMode
        Else
            pf_dynDrawDoor point, ptP, point, view, DrawMode
        End If
    End If
    
    ' MIRROR AND FREE ROTATE BOTH
    If formDD.chkMirror And formDD.chkRotate Then
        If ptNum = 0 Then
            pf_dynDrawDoor point, point, point, view, DrawMode
        ElseIf ptNum = 1 Then
            pf_dynDrawDoor point, ptP, point, view, DrawMode
        Else
            pf_dynDrawDoor point, ptP, point, view, DrawMode
        End If
    End If
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal KeyIn As String)

End Sub
' ===============================================================================
' RESET EVENT HANDLER.
'
' Ramsey Systems, Inc. March 15, 2002
' ===============================================================================
Private Sub IPrimitiveCommandEvents_Reset()
    CommandState.StartDefaultCommand
End Sub
' ===============================================================================
' PRIMITIVE INITIALIZATION.
'
' Ramsey Systems, Inc. April 18, 2002
' ===============================================================================
Private Sub IPrimitiveCommandEvents_Start()
    ' INITIALIZE DATAPOINT COUNTER
    ptNum = 0
        
    CommandState.EnableAccuSnap
    CommandState.StartDynamics
End Sub
' ===============================================================================
' pf_dynDrawDoor - DRAW ROUTINE : : : DOOR
'
' Ramsey Systems, Inc. April 18, 2002
' ===============================================================================
Private Sub pf_dynDrawDoor(mPt As Point3d, point As Point3d, rPt As Point3d, view As view, DrawMode As MsdDrawingMode)
    Dim linePts() As Point3d, constructPt As Point3d, origin As Point3d, jamVertices() As Point3d
    Dim oLine() As LineElement
    Dim oArc As ArcElement
    Dim oArc2 As ArcElement
    Dim oCell As CellElement
    Dim oElts() As Element
    Dim i As Integer
    
    ' DRAW ABOUT 0,0,0
    constructPt.X = 0
    constructPt.Y = 0
    constructPt.Z = 0
          
    ' DRAW LINES FOR THE TAP FROM EDGE
    ReDim linePts(0 To 1)
    ReDim oLine(0 To 3)
    
    ' DRAW THE RIGHT DOOR
    linePts(0) = constructPt
    linePts(0).X = constructPt.X + formDD.txtDoorSize / 2
    linePts(1) = linePts(0)
    linePts(1).Y = linePts(0).Y + formDD.txtDoorSize / 2
    Set oLine(0) = CreateLineElement1(Nothing, linePts)
    
    ' DRAW THE LEFT DOOR
    linePts(0) = constructPt
    linePts(0).X = constructPt.X - formDD.txtDoorSize / 2
    linePts(1) = linePts(0)
    linePts(1).Y = linePts(0).Y + formDD.txtDoorSize / 2
    Set oLine(1) = CreateLineElement1(Nothing, linePts)
    
    ' DRAW FIRST JAM
    ReDim jamVertices(0 To 4)
    
    jamVertices(0) = oLine(0).StartPoint
    jamVertices(0).X = jamVertices(0).X + formDD.txtJamSize / 2   ' Developer's Whim - a detail that needs to be defined by the spec
    jamVertices(0).Y = formDD.txtJamSize / 2
    jamVertices(1) = jamVertices(0)
    jamVertices(1).X = jamVertices(0).X - formDD.txtJamSize / 2
    jamVertices(2) = jamVertices(1)
    jamVertices(2).Y = jamVertices(1).Y - formDD.txtJamSize
    jamVertices(3) = jamVertices(2)
    jamVertices(3).X = jamVertices(2).X + formDD.txtJamSize / 2
    jamVertices(4) = jamVertices(0)
    Set oLine(2) = CreateLineElement1(Nothing, jamVertices)
    
    ' DRAW SECOND JAM
    jamVertices(0) = constructPt
    jamVertices(0).X = constructPt.X - formDD.txtDoorSize / 2 - formDD.txtJamSize / 2
    jamVertices(0).Y = formDD.txtJamSize / 2
    jamVertices(1) = jamVertices(0)
    jamVertices(1).X = jamVertices(0).X + formDD.txtJamSize / 2
    jamVertices(2) = jamVertices(1)
    jamVertices(2).Y = jamVertices(1).Y - formDD.txtJamSize
    jamVertices(3) = jamVertices(2)
    jamVertices(3).X = jamVertices(2).X - formDD.txtJamSize / 2
    jamVertices(4) = jamVertices(0)
    Set oLine(3) = CreateLineElement1(Nothing, jamVertices)
    
    ' DRAW DOOR SWEEP
    Set oArc = CreateArcElement2(Nothing, oLine(0).StartPoint, formDD.txtDoorSize / 2, formDD.txtDoorSize / 2, view.Rotation, Pi / 2, Pi / 2)
    Set oArc2 = CreateArcElement2(Nothing, oLine(1).StartPoint, formDD.txtDoorSize / 2, formDD.txtDoorSize / 2, view.Rotation, 0, Pi / 2)
    
    ' FORCE THE ARC WEIGHT TO BE 0
    oArc.LineWeight = 0
    oArc2.LineWeight = 0
    
    ' ADD ELEMENTS TO CELL
    ReDim oElts(0 To 5)
    
    For i = 0 To 3
        Set oElts(i) = oLine(i)
    Next
    
    Set oElts(4) = oArc
    Set oElts(5) = oArc2
    
    Set oCell = CreateCellElement1("DEMODOOR", oElts, constructPt, False)
    
    origin = point
    
    Dim Eltrans As Transform3d
    Eltrans = Transform3dFromXYZ(origin.X, origin.Y, origin.Z)
    oCell.Transform Eltrans
    
    ' HANDLE ACTIVE ROTATION
    Eltrans = Transform3dFromMatrix3dAndFixedPoint3d(Matrix3dFromAxisAndRotationAngle(2, ActiveSettings.angle), point)
    oCell.Transform Eltrans
        
    ' HANDLE SCALE REQUEST
    Dim skale As Point3d
    skale = ActiveSettings.Scale
    Eltrans = Transform3dFromMatrix3dAndFixedPoint3d(Matrix3dFromScaleFactors(skale.X, skale.Y, skale.Z), point)
    oCell.Transform Eltrans
    
    ' HANDLE MIRROR REQUESTS
    If ptNum = 1 And formDD.chkMirror Then
        Dim stat As Integer
        stat = pf_mirrorFitting(oCell, mPt, point, view)
    End If

    ' HANDLE FREE ROTATION - AFTER MIRROR REQUEST
    If (formDD.chkRotate And formDD.chkMirror And ptNum = 2) Or (formDD.chkRotate And Not formDD.chkMirror And ptNum = 1) Then
        Dim rAngle As Double, isFlipped As Boolean
        
        rAngle = getAngle(point, rPt, isFlipped)
        If Not isFlipped Then rAngle = rAngle + Pi
        Eltrans = Transform3dFromMatrix3dAndFixedPoint3d(Matrix3dFromAxisAndRotationAngle(2, rAngle), point)
        oCell.Transform Eltrans
    End If
    
    ' DRAW THE CELL
    If DrawMode = msdDrawingModeNormal Then
        ActiveModelReference.AddElement oCell
        oCell.Redraw DrawMode
    Else
        oCell.Redraw DrawMode
    End If
    
    For i = 0 To 3
        Set oLine(i) = Nothing
        Set oElts(i) = Nothing
    Next
    
    For i = 0 To 5
    Set oElts(i) = Nothing
    Next
    Set oCell = Nothing
End Sub

