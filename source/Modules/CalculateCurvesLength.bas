Attribute VB_Name = "CalculateCurvesLength"
Sub Start()
    Dim s As Shape
    If ActiveDocument Is Nothing Then
        answer = MsgBox("No Active Document opened, create new?", vbOKCancel, "Warning")
        If answer = 1 Then
            ' Open New document dialog by clicking the New button on panel
            ' There's no legal method for calling this function from the code
            Application.FrameWork.Automation.InvokeItem "fa65d0c1-879b-4ef5-9465-af09e00e91ab"
        End If
    Else
        Set s = ActiveDocument.Selection
        If s.Shapes.Count = 0 Then
            MsgBox "No Object(s) selected. Nothing to measure.", vbCritical, "Warning"
            Exit Sub
        Else
            DoJob
        End If
    End If
End Sub

Sub DoJob(Optional Msg As String)

    On Error GoTo ErrorHandler

    Optimization = True
    
    Dim AllShapes As Shapes
    Set AllShapes = ActiveDocument.Selection.Shapes
    
    Dim cutLength As Single
    cutLength = 0
    
    ActiveDocument.Unit = cdrMillimeter
    
    Dim k As Integer
    For k = 1 To AllShapes.Count
        ' Call processing procedure
        cutLength = cutLength + MeasureShape(AllShapes(k))
    Next k
    
    Optimization = False
    ActiveWindow.Refresh
    
    If IsMissing(Msg) Then Msg = ""
    
    answer = MsgBox(Msg & "Total shapes = " & AllShapes.Count & ". Contour length = " & cutLength & "mm. Copy length to clipboard?", vbOKCancel, "Completed")
    
    If answer = 1 Then
        Dim obj As New MSForms.DataObject
        obj.SetText cutLength
        obj.PutInClipboard
    End If
    Exit Sub
    
ErrorHandler:
    ' Message on error
    MsgBox "Error occured on shape #" & k, vbCritical, "Critical error"
    ResetAfterError
    
End Sub

Private Function MeasureShape(currentShape As Shape) As Single
On Error GoTo ErrorHandler
    Dim AllSegments As SegmentRange
    Set currentShape = currentShape.Duplicate
    currentShape.ConvertToCurves
    Set AllSegments = currentShape.Curve.Nodes.All.SegmentRange
    Dim k As Integer
    MeasureShape = 0
    For k = 1 To AllSegments.Count
        MeasureShape = MeasureShape + AllSegments(k).Length
    Next k
    currentShape.Delete
    Exit Function
    
ErrorHandler:
    currentShape.Delete
    Err.Raise 0
End Function

Public Sub ResetAfterError()
    Optimization = False
    ActiveWindow.Refresh
    ActiveSelection.Ungroup
End Sub
