Attribute VB_Name = "correctTable"
Function IsEven(varParam)
    IsEven = ((varParam Mod 2) = 0)
End Function

Function IsOdd(varParam)
    IsOdd = ((varParam Mod 2) = 1)
End Function

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
            ProgressWindow.Tag = "correctTable.DoJob"
            ProgressWindow.Caption = "Processing table"
            ProgressWindow.Show
        End If
    End If
End Sub

Public Sub DoJob()
    ' Create Undo action for whole sequence (for performance)
    ActiveDocument.BeginCommandGroup "Correct Table For Cut"
    
    MACRO_STATUS = 1
    ActiveDocument.Unit = cdrMillimeter
    Optimization = True
    Dim k
    Dim pctCompl As Single, TOTAL As Single, DONE As Single
    Dim EvenCount As Boolean
    Dim sel As Shape
    Dim lines As Shapes
    Dim horLines As ShapeRange
    Dim vertLines As ShapeRange
    Set sel = ActiveDocument.Selection
    EvenCount = False
    'Starting
    
      '5%
      With sel
          .Flip cdrFlipHorizontal
          .Ungroup
      End With
      pctCompl = 5 / 100
      ProgressWindow.Progress pctCompl * 100, ProgressWindow.Frame.width * pctCompl
      
      '15%
      Set lines = ActiveSelection.Shapes
      For Each l In lines
          If l.SizeWidth > l.SizeHeight Then
              l.Outline.Color.CMYKAssign 0, 100, 100, 0
          Else
              l.Outline.Color.CMYKAssign 100, 0, 100, 0
          End If
      Next l
      pctCompl = 20 / 100
      ProgressWindow.Progress pctCompl * 100, ProgressWindow.Frame.width * pctCompl
      
      '40%
      Set vertLines = lines.FindShapes(Query:="@outline.color = cmyk(100,0,100,0)")
      TOTAL = vertLines.Count
      DONE = 0
      For k = 1 To vertLines.Count
          vertLines.Range(k).OrderToFront
          If IsOdd(k) Then
              vertLines.Range(k).Flip cdrFlipVertical
          End If
          DONE = DONE + 1
          pctCompl = (20 + (40 / TOTAL * DONE)) / 100
          ProgressWindow.Progress pctCompl * 100, ProgressWindow.Frame.width * pctCompl
      Next k
    
      '40%
      Set horLines = lines.FindShapes(Query:="@outline.color = cmyk(0,100,100,0)")
      EvenCount = IsEven(horLines.Count)
      TOTAL = horLines.Count
      DONE = 0
      For k = 1 To horLines.Count
          If EvenCount Then
              If IsEven(k) Then
                  horLines.Range(k).Flip cdrFlipHorizontal
              End If
          Else
              If IsOdd(k) Then
                  horLines.Range(k).Flip cdrFlipHorizontal
              End If
          End If
          DONE = DONE + 1
          pctCompl = (60 + (40 / TOTAL * DONE)) / 100
          ProgressWindow.Progress pctCompl * 100, ProgressWindow.Frame.width * pctCompl
      Next k

    ActiveLayer.Shapes.All.CreateSelection
    ActiveSelection.Group
    Optimization = False
    ActiveWindow.Refresh
    ActiveSelection.Ungroup
    Unload ProgressWindow
    MACRO_STATUS = 0
    ActiveDocument.EndCommandGroup
End Sub
