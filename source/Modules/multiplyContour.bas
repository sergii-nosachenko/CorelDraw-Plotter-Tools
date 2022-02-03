Attribute VB_Name = "multiplyContour"
Sub DuplicateContour()
    Dim s As Shape
    If ActiveDocument Is Nothing Then
        MsgBox "No Active Document opened"
        Exit Sub
    Else
        Set s = ActiveDocument.Selection
        If s.Shapes.Count = 0 Then
            MsgBox "No Object Selected"
            Exit Sub
        Else
            multiplyOptions.Show
        End If
    End If
End Sub

Public Sub DoJob()
    ' Create Undo action for whole sequence (for performance)
    ActiveDocument.BeginCommandGroup "Multiply Contour"
    
    MACRO_STATUS = 1
    Optimization = True
    ActiveDocument.Unit = cdrMillimeter
    Dim r, c, k
    Dim sel, copie As Shape
    Dim pctCompl As Single, TOTAL As Single, DONE As Single
    Set sel = ActiveDocument.Selection
    TOTAL = multiplyOptions.copiesTop.Text * multiplyOptions.copiesLeft.Text
    ActiveDocument.ActiveShape.Name = "CUT"
    sel.OrderToFront
    DONE = 1
    'Starting
    For r = 1 To multiplyOptions.copiesTop.Text
        If r > 1 Then
            sel.Duplicate 0, 0
            sel.OrderToFront
            sel.Move 0, multiplyOptions.distanceTop.Value
            DONE = DONE + 1
        End If
        For c = 1 To multiplyOptions.copiesLeft.Text
            If (r Mod 2) = 1 Then 'is Odd
                k = -1
            Else
                k = 1
            End If
            If c > 1 Then
                sel.Duplicate 0, 0
                sel.OrderToFront
                sel.Move k * multiplyOptions.distanceLeft.Value, 0
                DONE = DONE + 1
            End If
        Next
        pctCompl = DONE / TOTAL
        ProgressWindow.Progress pctCompl * 100, ProgressWindow.Frame.width * pctCompl
    Next
    ActiveLayer.FindShapes(Name:="CUT").CreateSelection
    ActiveSelection.Group
    Optimization = False
    ActiveWindow.Refresh
    ActiveSelection.Ungroup
    ActiveDocument.ClearSelection
    Unload ProgressWindow
    Unload multiplyOptions
    MACRO_STATUS = 0
    ActiveDocument.EndCommandGroup
End Sub
