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
    Optimization = True
    ActiveDocument.Unit = cdrMillimeter
    Dim r, c, k
    Dim sel, copie As Shape
    Dim pctCompl As Single, total As Single, done As Single
    Set sel = ActiveDocument.Selection
    total = multiplyOptions.copiesTop.Text * multiplyOptions.copiesLeft.Text
    ActiveDocument.ActiveShape.name = "CUT"
    sel.OrderToFront
    done = 1
    'Starting
    For r = 1 To multiplyOptions.copiesTop.Text
        If r > 1 Then
            sel.Duplicate 0, 0
            sel.OrderToFront
            sel.Move 0, multiplyOptions.distanceTop.Value
            done = done + 1
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
                done = done + 1
            End If
        Next
        pctCompl = done / total
        Progress pctCompl * 100, ProgressWindow.Frame.width * pctCompl
    Next
    ActiveLayer.FindShapes(name:="CUT").CreateSelection
    ActiveSelection.Group
    Optimization = False
    ActiveWindow.Refresh
    ActiveSelection.Ungroup
    ActiveDocument.ClearSelection
    Unload ProgressWindow
    Unload multiplyOptions
End Sub

Private Sub Progress(pctCompl As Single, width As Single)

    ProgressWindow.Text.Caption = Round(pctCompl, 0) & "% complete"
    ProgressWindow.Bar.width = width
    DoEvents

End Sub
