Attribute VB_Name = "CreateTableForCut"
Sub Start()
    Dim s, refPointBL, refPointTR As Shape
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
            Set refPointBL = p.FindShape(Name:="refPointBL")
            Set refPointTR = p.FindShape(Name:="refPointTR")
            If refPointBL Is Nothing Or refPointTR Is Nothing Then
                answer = MsgBox("No Shape selected (as bounds source) / No reference points found. Use document bounds instead?", vbOKCancel, "Warning")
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
            CreateTableForCut_Options.Show
        Else
            With FRAME_SIZE
                .LeftX = s.LeftX
                .TopY = s.TopY
                .RightX = s.RightX
                .BottomY = s.BottomY
            End With
            CreateTableForCut_Options.Show
        End If
    End If
End Sub

Public Sub DoJob()
    ' Create Undo action for whole sequence (for performance)
    ActiveDocument.BeginCommandGroup "Create table for cut"
    
    MACRO_STATUS = 1
    ActiveDocument.Unit = cdrMillimeter
    Optimization = True
    
    Dim Cols, Rows, CellWidth, CellHeight, Overcut As Single
    Dim pctCompl, TOTAL, DONE As Single
    Dim width, Height As Single
    Dim ExpectedWidth, ExpectedHeight As Double
    Dim PosX, PosY As Double
    Dim table As Shape
    
    Cols = CreateTableForCut_Options.ColumnsCount.Value
    Rows = CreateTableForCut_Options.RowsCount.Value
    CellWidth = CreateTableForCut_Options.CellWidth.Value
    CellHeight = CreateTableForCut_Options.CellHeight.Value
    Overcut = CreateTableForCut_Options.Overcut.Value * 2
    
    '5%
    width = FRAME_SIZE.RightX - FRAME_SIZE.LeftX
    Height = FRAME_SIZE.TopY - FRAME_SIZE.BottomY
    ExpectedWidth = CellWidth * Cols
    ExpectedHeight = CellHeight * Rows
    PosX = FRAME_SIZE.LeftX + width / 2
    PosY = FRAME_SIZE.BottomY + Height / 2
    pctCompl = 5 / 100
    ProgressWindow.Progress pctCompl * 100, ProgressWindow.FRAME.width * pctCompl
    'Starting
    '25%
    Set table = ActiveLayer.CreateCustomShape("Table", 1, 10, 5, 7, Cols, Rows)
    pctCompl = pctCompl + 25 / 100
    ProgressWindow.Progress pctCompl * 100, ProgressWindow.FRAME.width * pctCompl
    '20%
    table.SetSize width, Height
    pctCompl = pctCompl + 20 / 100
    ProgressWindow.Progress pctCompl * 100, ProgressWindow.FRAME.width * pctCompl
    '20%
    table.SetPosition PosX, PosY
    pctCompl = pctCompl + 20 / 100
    ProgressWindow.Progress pctCompl * 100, ProgressWindow.FRAME.width * pctCompl
    '20%
    table.ConvertToCurves
    pctCompl = pctCompl + 20 / 100
    ProgressWindow.Progress pctCompl * 100, ProgressWindow.FRAME.width * pctCompl
    '5%
    table.SetSizeEx table.CenterX, table.CenterY, ExpectedWidth, ExpectedHeight
    pctCompl = pctCompl + 5 / 100
    ProgressWindow.Progress pctCompl * 100, ProgressWindow.FRAME.width * pctCompl
    '5%
    ActiveDocument.Selection.group
    ActiveDocument.ActiveShape.Name = "CUT"
    ActiveLayer.FindShapes("CUT").CreateSelection
    pctCompl = pctCompl + 5 / 100
    ProgressWindow.Progress pctCompl * 100, ProgressWindow.FRAME.width * pctCompl
    
    Unload ProgressWindow
    
    CorrectTableForCut.DoJob Overcut
    
    Optimization = False
    ActiveWindow.Refresh
    If Not ActiveSelection Is Nothing Then ActiveSelection.Ungroup
    Unload CreateTableForCut_Options
    ActiveDocument.EndCommandGroup
    
    CalculateCurvesLength.DoJob "Total cells = " & (Cols * Rows) & ". "
    
    MACRO_STATUS = 0
    
End Sub

Public Sub ResetAfterError()
    Optimization = False
    ActiveWindow.Refresh
    ActiveSelection.Ungroup
End Sub
