Public SmoothnessLevel As Integer, PassesCount As Integer
Public FilletValue As Double
Public Otimization As Boolean
Public pctCompl As Single, total As Single, done As Single
Sub Start()
    Dim sel As Shape
    If ActiveDocument Is Nothing Then
        MsgBox "No Active Document opened"
        Exit Sub
    Else
        Set sel = ActiveDocument.Selection
        If sel.Shapes.Count = 0 Then
            MsgBox "No Object Selected"
            Exit Sub
        Else
            PrepareForCutOptions.Show
        End If
    End If
End Sub
Public Sub DoJob()
    'Starting
    ActiveDocument.Unit = cdrMillimeter
    Optimization = True
    Dim k
    
    Dim TempLayer As Layer, OriginalLayer As Layer
    Set TempLayer = ActiveDocument.ActivePage.AllLayers.Find("Temporary!!!")
    
    If TempLayer Is Nothing Then
        Set TempLayer = ActiveDocument.ActivePage.CreateLayer("Temporary!!!")
    End If
        
    Dim OrigSelection As Shape
    Set OrigSelection = ActiveDocument.Selection
    Set OriginalLayer = OrigSelection.Shapes.First.Layer
    
    OrigSelection.UngroupAll
    OrigSelection.MoveToLayer TempLayer
    
    Dim AllShapes As Shapes
    Set AllShapes = TempLayer.Shapes
    If Otimization Then
        total = AllShapes.Count * (PassesCount + 1)
    Else
        total = AllShapes.Count * PassesCount
    End If
    done = 0

    For k = 1 To AllShapes.Count

        ProcessShape AllShapes(k)
        AllShapes(k).Outline.Color.CMYKAssign 0, 100, 0, 0
    
    Next k
    
    Set AllShapes = TempLayer.FindShapes("CUT").Shapes
    
    AllShapes.All.CreateSelection
    AllShapes.All.MoveToLayer OriginalLayer
    TempLayer.Delete
      
    Optimization = False
    ActiveWindow.Refresh
    Unload ProgressWindow
    Unload PrepareForCutOptions
End Sub

Private Sub Progress(pctCompl As Single, width As Single)

    ProgressWindow.Text.Caption = Round(pctCompl, 0) & "% complete"
    ProgressWindow.Bar.width = width
    DoEvents

End Sub

Private Sub ProcessShape(curShape As Shape)

    Dim i As Integer
    
    Dim AllNodes As NodeRange
    
    Dim FilletValueLocal As Double
    
    Dim SmoothnessLevelLocal As Integer
    
    FilletValueLocal = FilletValue
    
    SmoothnessLevelLocal = SmoothnessLevel
    
    curShape.ConvertToCurves
    
    curShape.Curve.JoinTouchingSubpaths False, 0.2
    
    If Otimization Then
        'Contour tool magic
        
        Dim newShape As Shape
        Dim outside As Effect
        Dim outsideShape As ShapeRange
        
        Set outside = curShape.CreateContour(cdrContourOutside, 0.05, 1, , CreateCMYKColor(100, 0, 0, 0))
        
        If Not outside.Contour Is Nothing Then
            Set outsideShape = outside.Contour.ContourGroup.BreakApartEx
            Set newShape = outsideShape(1)
        End If
        
        If Not newShape Is Nothing Then
            curShape.Delete
            Set curShape = newShape
        End If
        
        'Progress
        done = done + 1
        pctCompl = done / total
        Progress pctCompl * 100, ProgressWindow.Frame.width * pctCompl
        
    End If
    
    curShape.Curve.Segments.All.SetType cdrCurveSegment
    
    curShape.CreateSelection
    
    i = 0
    
    Do Until i = PassesCount
        If Otimization Then SimplifyCloseupNodes curShape.Curve.Nodes.All, FilletValueLocal
        Set AllNodes = curShape.Curve.Nodes.All
        If SmoothnessLevelLocal > 0 Then AllNodes.Smoothen SmoothnessLevelLocal
        AllNodes.SetType cdrCuspNode
        If FilletValueLocal > 0 Then AllNodes.Fillet FilletValueLocal, True
        FilletValueLocal = FilletValueLocal - 0.25
        SmoothnessLevelLocal = SmoothnessLevelLocal - 3
        If FilletValueLocal < 0.25 Then FilletValueLocal = 0
        If SmoothnessLevelLocal < 1 Then SmoothnessLevelLocal = 0
        i = i + 1
        'Progress
        done = done + 1
        pctCompl = done / total
        Progress pctCompl * 100, ProgressWindow.Frame.width * pctCompl
    Loop
    
    curShape.Name = "CUT"
    
End Sub

Private Sub SimplifyCloseupNodes(AllNodes As NodeRange, FilletValueLocal As Double)
    
    Dim Angle As Double
    
    Dim AllSegments As SegmentRange
    Dim prevSegment As Segment, curSegment As Segment, nextSegment As Segment
    Dim startNode As Node, endNode As Node
    Set AllSegments = AllNodes.SegmentRange
    Set curSegment = AllSegments.FirstSegment
    
    If FilletValueLocal < 0.25 Then FilletValueLocal = 0.25
    
    Do
        If curSegment.Length <= FilletValueLocal Then
            Set prevSegment = curSegment.previous
            Set nextSegment = curSegment.Next
            If prevSegment.Length > 1 Then prevSegment.AddNodeAt 0.5, cdrRelativeSegmentOffset
            If nextSegment.Length > 1 Then nextSegment.AddNodeAt 0.5, cdrRelativeSegmentOffset
        End If
        Set curSegment = curSegment.Next
    Loop Until curSegment.index = AllSegments.FirstSegment.index

    Set AllNodes = ActiveDocument.Selection.Shapes.First.Curve.Nodes.All
    Set AllSegments = AllNodes.SegmentRange
    Set curSegment = AllSegments.FirstSegment
    
    Do
        On Error Resume Next

        Set startNode = curSegment.startNode
        Set endNode = curSegment.endNode
        Set prevSegment = curSegment.previous
        Set nextSegment = curSegment.Next
        Angle = Abs(curSegment.EndingControlPointAngle - nextSegment.StartingControlPointAngle)
        If curSegment.Length <= FilletValueLocal And Angle > 45 Then
            endNode.Delete
        End If
        Angle = Abs(prevSegment.EndingControlPointAngle - curSegment.StartingControlPointAngle)
        If curSegment.Length <= FilletValueLocal And Angle > 45 Then
            startNode.Delete
        End If
        Set curSegment = curSegment.Next
    Loop Until curSegment.index = AllSegments.FirstSegment.index
End Sub
