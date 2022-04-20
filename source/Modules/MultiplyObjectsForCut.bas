Attribute VB_Name = "MultiplyObjectsForCut"
Sub Start()
    Dim s, refPointBL, refPointTR As Shape
    Dim p As Page
    
    If ActiveDocument Is Nothing Then
        answer = MsgBox("No Active Document opened, create new?", vbOKCancel, "Warning")
        If answer = 1 Then
            ' Open New document dialog by clicking the New button on panel
            ' There's no legal method for calling this function from the code
            Application.FrameWork.Automation.InvokeItem "fa65d0c1-879b-4ef5-9465-af09e00e91ab"
        End If
    Else
    
        ActiveDocument.Unit = cdrMillimeter
        Set s = ActiveDocument.Selection
        If s.Shapes.Count = 0 Then
            MsgBox "No Object(s) selected as source", vbCritical, "Warning"
            Exit Sub
        Else
            Set p = ActiveDocument.ActivePage
            Set refPointBL = p.FindShape(Name:="refPointBL")
            Set refPointTR = p.FindShape(Name:="refPointTR")
            If refPointBL Is Nothing Or refPointTR Is Nothing Then
                answer = MsgBox("No reference points found. Use document bounds instead?", vbOKCancel, "Warning")
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
            Else
                With FRAME_SIZE
                    .LeftX = refPointBL.CenterX
                    .TopY = refPointTR.CenterY
                    .RightX = refPointTR.CenterX
                    .BottomY = refPointBL.CenterY
                End With
            End If
            MultiplyObjectsForCut_Options.Show
        End If
    End If
End Sub

Public Sub DoJob()
    ' Create Undo action for whole sequence (for performance)
    ActiveDocument.BeginCommandGroup "Multiply objects/contours"
    
    MACRO_STATUS = 1
    Optimization = True
    ActiveDocument.Unit = cdrMillimeter
    ActiveDocument.ReferencePoint = cdrCenter
    
    Dim r, c, k
    Dim sel As Shape
    Dim items As ShapeRange
    Dim pctCompl As Single, TOTAL As Single, DONE As Single
    Set sel = ActiveDocument.Selection.group
    TOTAL = MultiplyObjectsForCut_Options.copiesTop.Text * MultiplyObjectsForCut_Options.copiesLeft.Text
    sel.Name = "ITEM"
    sel.OrderToFront
    DONE = 1
    'Starting
    For r = 1 To MultiplyObjectsForCut_Options.copiesTop.Text
        If r > 1 Then
            sel.Duplicate 0, 0
            sel.OrderToFront
            sel.Move 0, MultiplyObjectsForCut_Options.distanceTop.Value
            DONE = DONE + 1
        End If
        For c = 1 To MultiplyObjectsForCut_Options.copiesLeft.Text
            If (r Mod 2) = 1 Then 'is Odd
                k = -1
            Else
                k = 1
            End If
            If c > 1 Then
                sel.Duplicate 0, 0
                sel.OrderToFront
                sel.Move k * MultiplyObjectsForCut_Options.distanceLeft.Value, 0
                DONE = DONE + 1
            End If
        Next
        pctCompl = DONE / TOTAL
        ProgressWindow.Progress pctCompl * 100, ProgressWindow.FRAME.width * pctCompl
    Next
    
    Set items = ActiveLayer.FindShapes("ITEM")
    Set sel = items.group

    sel.CenterX = (FRAME_SIZE.RightX - FRAME_SIZE.LeftX) / 2
    sel.CenterY = (FRAME_SIZE.TopY - FRAME_SIZE.BottomY) / 2
    
    Optimization = False
    ActiveWindow.Refresh
    If Not sel Is Nothing Then sel.Ungroup
    ActiveDocument.ClearSelection
    Unload ProgressWindow
    Unload MultiplyObjectsForCut_Options
    MACRO_STATUS = 0
    ActiveDocument.EndCommandGroup
End Sub

Public Sub ResetAfterError()
    Optimization = False
    ActiveWindow.Refresh
    ActiveSelection.Ungroup
End Sub
