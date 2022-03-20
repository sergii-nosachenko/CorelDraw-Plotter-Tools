Attribute VB_Name = "ReDefineReferencePoints"
Sub Start()
    Dim s As Shape
    Dim p As Page
    If ActiveDocument Is Nothing Then
        answer = MsgBox("No Active Document opened, create new one?", vbOKCancel, "Warning")
        If answer = 1 Then
            ' Open New document dialog by clicking the New button on panel
            ' There's no legal method for calling this function from the code
            Application.FrameWork.Automation.InvokeItem "fa65d0c1-879b-4ef5-9465-af09e00e91ab"
        End If
    Else
        ActiveDocument.Unit = cdrMillimeter
        Set s = ActiveDocument.Selection
        If s Is Nothing Or s.Shapes.Count = 0 Then
            Set p = ActiveDocument.ActivePage
            answer = MsgBox("No Shape selected (as bounds source). Use document bounds instead?", vbOKCancel, "Warning")
            If answer = 1 Then
                With FRAME_SIZE
                    .LeftX = p.LeftX
                    .TopY = p.TopY
                    .RightX = p.RightX
                    .BottomY = p.BottomY
                End With
            Else
                Exit Sub
            End If
            DoJob
        Else
            With FRAME_SIZE
                .LeftX = s.LeftX
                .TopY = s.TopY
                .RightX = s.RightX
                .BottomY = s.BottomY
            End With
            DoJob
        End If
    End If
End Sub

Private Sub DoJob()
    Dim guide, refPointBL, refPointTR As Shape
    Dim l, originalLayer As Layer
    
    ' Create Undo action for whole sequence (for performance)
    ActiveDocument.BeginCommandGroup "ReDefine frame and reference points"

    ' Set default units to mm
    ActiveDocument.Unit = cdrMillimeter
    
    ' Disable redrawing of active view
    Application.Optimization = True
    ActiveDocument.Selection.Delete
    Set l = ActiveDocument.ActivePage.GuidesLayer
    Set originalLayer = ActiveDocument.ActivePage.ActiveLayer
    If l.Editable = False Then l.Editable = True
    l.FindShapes("FrameGuide").Delete
    l.FindShapes("refPointBL").Delete
    l.FindShapes("refPointTR").Delete
    ' Frame guide
    Set guide = l.CreateRectangle(FRAME_SIZE.LeftX, FRAME_SIZE.TopY, FRAME_SIZE.RightX, FRAME_SIZE.BottomY)
    guide.Name = "FrameGuide"
    ' Bottom left reference point
    Set refPointBL = l.CreateEllipse2(FRAME_SIZE.LeftX, FRAME_SIZE.BottomY, 0.5, 0.5)
    refPointBL.Name = "refPointBL"
    ' Top right reference point
    Set refPointTR = l.CreateEllipse2(FRAME_SIZE.RightX, FRAME_SIZE.TopY, 0.5, 0.5)
    refPointTR.Name = "refPointTR"
    
    originalLayer.Activate
    ActiveDocument.ClearSelection
    
    'l.MoveAbove ActiveDocument.ActivePage.AllLayers.Top
    l.Editable = False
    
    ' Enable redrawing of active view
    Application.Optimization = False
    Application.Refresh
    
    ActiveDocument.EndCommandGroup
    
End Sub
