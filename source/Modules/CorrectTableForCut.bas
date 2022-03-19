Attribute VB_Name = "CorrectTableForCut"
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
            ProgressWindow.Tag = "CorrectTableForCut.DoJob"
            ProgressWindow.Caption = "Processing table"
            ProgressWindow.Show
        End If
    End If
End Sub

Sub DoJob(Optional Overcut)

    ' Create Undo action for whole sequence (for performance)
    ActiveDocument.BeginCommandGroup "Correct Table For Cut"
    
    MACRO_STATUS = 1
    ActiveDocument.Unit = cdrMillimeter
    Optimization = True
    Dim k
    Dim pctCompl, TOTAL, DONE As Single
    Dim EvenCount As Boolean
    Dim sel, lines As Shapes
    Dim horLines, vertLines As ShapeRange
    Set sel = ActiveDocument.Selection
    
    If IsMissing(Overcut) Then Overcut = 0

    ActiveDocument.ReferencePoint = cdrCenter
    
    EvenCount = False
    'Starting
    
      '+10%

      ' Creating of temp layer (or reusing existing)
      Dim TempLayer As Layer, OriginalLayer As Layer
      Set TempLayer = ActiveDocument.ActivePage.AllLayers.Find("Temporary!!!")
      If TempLayer Is Nothing Then
         Set TempLayer = ActiveDocument.ActivePage.CreateLayer("Temporary!!!")
      End If
      
      ' Moving shapes to temp layer
      Set OriginalLayer = sel.Shapes.First.Layer
      sel.ConvertToCurves
      sel.MoveToLayer TempLayer
      TempLayer.Shapes.All.UngroupAll
      
      Set lines = TempLayer.Shapes
      
      pctCompl = 10 / 100
      ProgressWindow.Progress pctCompl * 100, ProgressWindow.FRAME.width * pctCompl
        
      '+10%
      For Each l In lines
          If l.SizeWidth > l.SizeHeight Then
              l.Name = "CUT-HORIZONTAL"
          Else
              l.Name = "CUT-VERTICAL"
          End If
      Next l
      
      pctCompl = 20 / 100
      ProgressWindow.Progress pctCompl * 100, ProgressWindow.FRAME.width * pctCompl
      
      '+40%
      Set vertLines = lines.FindShapes("CUT-VERTICAL")
      TOTAL = vertLines.Count
      DONE = 0
      For k = 1 To vertLines.Count
          vertLines.Range(k).OrderToFront
          If IsOdd(k) Then
              vertLines.Range(k).Flip cdrFlipVertical
          End If
          vertLines.Range(k).SizeHeight = vertLines.Range(k).SizeHeight + Overcut
          DONE = DONE + 1
          pctCompl = (20 + (40 / TOTAL * DONE)) / 100
          ProgressWindow.Progress pctCompl * 100, ProgressWindow.FRAME.width * pctCompl
      Next k
    
      '+40%
      Set horLines = lines.FindShapes("CUT-HORIZONTAL")
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
          horLines.Range(k).SizeWidth = horLines.Range(k).SizeWidth + Overcut
          DONE = DONE + 1
          pctCompl = (60 + (40 / TOTAL * DONE)) / 100
          ProgressWindow.Progress pctCompl * 100, ProgressWindow.FRAME.width * pctCompl
      Next k
      
    Optimization = False
      
    ' Moving completed shapes to original layer
    TempLayer.Shapes.All.SetOutlineProperties 0.2, , CreateCMYKColor(0, 100, 0, 0)
    TempLayer.Shapes.All.group
    TempLayer.Shapes.All.Ungroup
    TempLayer.Shapes.All.MoveToLayer OriginalLayer
    TempLayer.Delete

    ActiveWindow.Refresh
    Unload ProgressWindow
    MACRO_STATUS = 0
    ActiveDocument.EndCommandGroup
End Sub

Public Sub ResetAfterError()
    Optimization = False
    ActiveWindow.Refresh
    ActiveSelection.Ungroup
End Sub
