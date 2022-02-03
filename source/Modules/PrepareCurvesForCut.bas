Attribute VB_Name = "PrepareCurvesForCut"
Dim PCT_COMPL As Single, TOTAL As Single, DONE As Single, MESSAGE As String
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
    ' Error handling
    On Error GoTo EndErrorJob
    ' Create Undo action for whole sequence (for performance)
    ActiveDocument.BeginCommandGroup "Prepare Curves For Cut"
    MACRO_STATUS = 1
    ' Set default units to mm
    ActiveDocument.Unit = cdrMillimeter
    
    ' Disable redrawing of active view
    Application.Optimization = True
    
    ' Creating of temp layer (or reusing existing)
    Dim TempLayer As Layer, OriginalLayer As Layer
    Set TempLayer = ActiveDocument.ActivePage.AllLayers.Find("Temporary!!!")
    If TempLayer Is Nothing Then
        Set TempLayer = ActiveDocument.ActivePage.CreateLayer("Temporary!!!")
    End If
        
    ' Moving shapes to temp layer
    Dim OrigSelection As Shape
    Set OrigSelection = ActiveDocument.Selection
    Set OriginalLayer = OrigSelection.Shapes.First.Layer
    OrigSelection.UngroupAll
    OrigSelection.MoveToLayer TempLayer
    
    Dim AllShapes As Shapes
    Set AllShapes = TempLayer.Shapes
    
    ' Set total value for progress bar
    If OPTIMIZE_ADVANCED Then
        TOTAL = AllShapes.Count * (PASSES_COUNT * 3 + 1)
    Else
        TOTAL = AllShapes.Count * PASSES_COUNT
    End If
    
    ' Set done value for progress bar
    DONE = 0
    PCT_COMPL = 0
    
    ' Looping through all shapes
    Dim totalShapes As Integer
    totalShapes = AllShapes.Count
    Dim k As Integer
    For k = 1 To totalShapes
    
        ProgressWindow.CheckStatus
        ' Call processing procedure
        ProcessShape AllShapes(k), k, totalShapes
    Next k
    
    ' Moving completed shapes to original layer
    Set AllShapes = TempLayer.FindShapes("CUT").Shapes
    AllShapes.All.CreateSelection
    AllShapes.All.MoveToLayer OriginalLayer
    TempLayer.Delete
    GoTo EndJob
    
EndErrorJob:
    ' Message on error
    MsgBox "Error occured on shape #" & k
EndJob:
    ' Unload forms and refresh window
    Unload ProgressWindow
    'Unload PrepareForCutOptions
    Application.Optimization = False
    ActiveWindow.Refresh
    Application.Refresh
    MACRO_STATUS = 0
    ActiveDocument.EndCommandGroup
End Sub

Private Sub ProcessShape(curShape As Shape, curIndex As Integer, totalShapes As Integer)

    Dim i As Integer
    Dim AllNodes As NodeRange
    Dim FilletValueLocal As Double
    Dim SmoothnessLevelLocal As Integer
    
    ' Basic message for current shape progress message
    MESSAGE = "(shape #" & curIndex & " of " & totalShapes & ")"
    
    ' Define local copies of global variables from form
    FilletValueLocal = FILLET_VALUE
    SmoothnessLevelLocal = SMOOTHNESS_LEVEL
    
    ' Converting shape to curve
    curShape.ConvertToCurves
    ' Close curve if endings are touching
    curShape.Curve.JoinTouchingSubpaths False, 0.2
    
    ' Save original width and height for final correction
    Dim origWidth, origHeight As Double
    origWidth = curShape.SizeWidth
    origHeight = curShape.SizeHeight
    
    If OPTIMIZE_ADVANCED Then
       
        'Progress message update
        ProgressWindow.Progress PCT_COMPL * 100, ProgressWindow.Frame.width * PCT_COMPL, MESSAGE & " | Reshaping..."
        
        Dim newShape As Shape
        Dim outside As Effect
        Dim outsideShape As ShapeRange
        
        curShape.CreateSelection
        
        ' Close curve
        curShape.Curve.Closed = True
        
        ' Reducing nodes for better performance and preventing bugs
        reduceNodes_cheating curShape.Curve.Nodes.All
        
        ' Add new contour outside current curve by 0.05 mm
        Set outside = curShape.CreateContour(cdrContourOutside, 0.05, 1, , CreateCMYKColor(100, 0, 0, 0))
        
        ' Break contour group apart
        If Not outside.Contour Is Nothing Then
            Set outsideShape = outside.Contour.ContourGroup.BreakApartEx
            Set newShape = outsideShape(1)
        End If
        
        ' Remove original shape
        If Not newShape Is Nothing Then
            curShape.Delete
            Set curShape = newShape
        End If
        
        'Progress update
        DONE = DONE + 1
        PCT_COMPL = DONE / TOTAL
        ProgressWindow.Progress PCT_COMPL * 100, ProgressWindow.Frame.width * PCT_COMPL, MESSAGE
        
    End If
    
    ' Set curve type to Curve Segment (prevents bugs)
    curShape.Curve.Segments.All.SetType cdrCurveSegment
    curShape.CreateSelection
    
    ' Loop in passes defined by user in form
    i = 0
    Do Until i = PASSES_COUNT
    
        ProgressWindow.CheckStatus
    
        If OPTIMIZE_ADVANCED Then
            ' Call the procedure to simplify and cleanup nodes
            SimplifyCloseupNodes curShape.Curve.Nodes.All, FilletValueLocal
        End If
        
        'Progress message update
        ProgressWindow.Progress PCT_COMPL * 100, ProgressWindow.Frame.width * PCT_COMPL, MESSAGE & " | Smoothing..."
        
        Set AllNodes = curShape.Curve.Nodes.All
        ' Smoothing curve
        If SmoothnessLevelLocal > 0 Then AllNodes.Smoothen SmoothnessLevelLocal
        
        ' Set nodes type to CuspNode (for better fillets) and set fillet value
        AllNodes.SetType cdrCuspNode
        If FilletValueLocal > 0 Then AllNodes.Fillet FilletValueLocal, True
        
        ' Adjust smoothness and fillet values for next step
        FilletValueLocal = FilletValueLocal - 0.25
        SmoothnessLevelLocal = SmoothnessLevelLocal - 3
        If FilletValueLocal < 0.25 Then FilletValueLocal = 0
        If SmoothnessLevelLocal < 1 Then SmoothnessLevelLocal = 0
        
        ' Reducing nodes for better performance and preventing bugs
        reduceNodes_cheating curShape.Curve.Nodes.All
        
        'Progress update
        DONE = DONE + 1
        PCT_COMPL = DONE / TOTAL
        ProgressWindow.Progress PCT_COMPL * 100, ProgressWindow.Frame.width * PCT_COMPL, MESSAGE
        
        i = i + 1
    Loop
    
    ' Restore original size
    ActiveDocument.ReferencePoint = cdrCenter
    curShape.SizeWidth = origWidth
    curShape.SizeHeight = origHeight
    curShape.Name = "CUT"
    ' Set name for completed shape
    curShape.Name = "CUT"
    ' Recolor otline of completed shape
    curShape.Outline.Color.CMYKAssign 0, 100, 0, 0
    
End Sub

Private Sub SimplifyCloseupNodes(AllNodes As NodeRange, FilletValueLocal As Double)
    
    ' Error handling (skip shape)
    On Error GoTo EndError
    
    Dim doneLocal As Single
    doneLocal = DONE
    
    Dim Angle As Double
    Dim AllSegments As SegmentRange
    Dim prevSegment, curSegment, nextSegment As Segment
    Dim startNode, endNode As Node
    
    Set AllSegments = AllNodes.SegmentRange
    Set curSegment = AllSegments.FirstSegment
    
    ' Adjust fillet value for simplifying purposes
    If FilletValueLocal < 0.25 Then FilletValueLocal = 0.25
    
    Do
        On Error Resume Next
        
        ProgressWindow.CheckStatus
        
        'Progress message update
        ProgressWindow.Progress PCT_COMPL * 100, ProgressWindow.Frame.width * PCT_COMPL, MESSAGE & " | Optimizing segment #" & curSegment.index & "..."
        
        ' If segment length <= FilletValueLocal add center points to neighbor segments
        If curSegment.Length <= FilletValueLocal Then
            Set prevSegment = curSegment.Previous
            Set nextSegment = curSegment.Next
            If prevSegment.Length > 1 Then prevSegment.AddNodeAt 0.5, cdrRelativeSegmentOffset
            If nextSegment.Length > 1 Then nextSegment.AddNodeAt 0.5, cdrRelativeSegmentOffset
        End If
        Set curSegment = curSegment.Next
    Loop Until curSegment.index = AllSegments.FirstSegment.index

    'Progress update
    DONE = DONE + 1
    PCT_COMPL = DONE / TOTAL
    ProgressWindow.Progress PCT_COMPL * 100, ProgressWindow.Frame.width * PCT_COMPL, MESSAGE

    Set AllNodes = ActiveDocument.Selection.Shapes.First.Curve.Nodes.All
    Set AllSegments = AllNodes.SegmentRange
    Set curSegment = AllSegments.FirstSegment
    
    Do
        On Error Resume Next
        
        ProgressWindow.CheckStatus
        
        'Progress message update
        ProgressWindow.Progress PCT_COMPL * 100, ProgressWindow.Frame.width * PCT_COMPL, MESSAGE & " | Cleaning segment #" & curSegment.index & "..."
        
        ' Clear unnecessary nodes on short segments (but only if angle > 90)
        ' Angle is important - node can hold sharp angle and removing one can cause curve reshaping
        Set startNode = curSegment.startNode
        Set endNode = curSegment.endNode
        Set prevSegment = curSegment.Previous
        Set nextSegment = curSegment.Next
        ' Calculate angle between control points of the start node
        Angle = Abs(curSegment.StartingControlPointAngle - prevSegment.EndingControlPointAngle)
        If curSegment.Length <= FilletValueLocal And Angle > 90 Then
            curSegment.StartingControlPointLength = 0
            curSegment.EndingControlPointLength = 0
            Set newNode = prevSegment.AddNodeAt(0.75, cdrRelativeSegmentOffset)
            Set curSegment = newNode.Segment.Next.Next
            Set startNode = curSegment.startNode
            startNode.Delete
            Set curSegment = endNode.Segment
            Set prevSegment = curSegment.Previous
            Set nextSegment = curSegment.Next
            Set startNode = curSegment.startNode
            Set endNode = curSegment.endNode
        End If
        
        ' Calculate angle between control points of the end node
        Angle = Abs(curSegment.EndingControlPointAngle - nextSegment.StartingControlPointAngle)
        If curSegment.Length <= FilletValueLocal And Angle > 90 Then
            curSegment.StartingControlPointLength = 0
            curSegment.EndingControlPointLength = 0
            Set newNode = nextSegment.AddNodeAt(0.25, cdrRelativeSegmentOffset)
            Set curSegment = newNode.Segment.Previous
            Set endNode = curSegment.endNode
            endNode.Delete
            Set curSegment = startNode.Segment.Next
        End If
        Set curSegment = curSegment.Next
    Loop Until curSegment.index = AllSegments.FirstSegment.index
    
    'Progress update
    DONE = DONE + 1
    PCT_COMPL = DONE / TOTAL
    ProgressWindow.Progress PCT_COMPL * 100, ProgressWindow.Frame.width * PCT_COMPL, MESSAGE
    
    GoTo EndSimplify
    
EndError:
    DONE = doneLocal + 2
EndSimplify:
End Sub

Sub reduceNodes_cheating(SelNodes As NodeRange)
    ' Reduce curve nodes by clicking the button on Node Edit Panel
    ' There's no legal method for calling this function from the code
    SelNodes.CreateSelection
    ActiveWindow.Activate
    ActiveTool = cdrToolNodeEdit
    Application.FrameWork.Automation.InvokeItem "b714bc06-7325-4d33-b330-4f4efec22c91"
    ' Wait until Reduce Nodes command completed
    DoEvents
End Sub
