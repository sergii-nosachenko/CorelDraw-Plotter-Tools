Attribute VB_Name = "createTableCut"
Dim Cols As String, Rows As String, CellWidth As Single, CellHeight As Single

Sub CreateTableCut()
    ActiveDocument.Unit = cdrMillimeter
    Dim s As Shape, refPointBL As Shape, refPointTR As Shape, l As Layer
    If ActiveDocument Is Nothing Then
        MsgBox "No Active Document opened"
        Exit Sub
    Else
        Set s = ActiveDocument.Selection
        If s Is Nothing Or s.Shapes.Count = 0 Then
            Set refPointBL = ActiveDocument.ActivePage.FindShape(Name:="refPointBL")
            Set refPointTR = ActiveDocument.ActivePage.FindShape(Name:="refPointTR")
            Set l = ActiveDocument.ActivePage.Layers.Find("Cut")
            If refPointBL Is Nothing Or refPointTR Is Nothing Or l Is Nothing Then
                MsgBox "No Object Selected And No Reference Points"
                Exit Sub
            Else
                l.CreateRectangle refPointBL.CenterX, refPointTR.CenterY, refPointTR.CenterX, refPointBL.CenterY
                l.Shapes.First.CreateSelection
                Set s = ActiveDocument.Selection
                CreateTable.Show
            End If
        Else
            CreateTable.Show
        End If
    End If
End Sub
Sub Start(Cols_ As String, Rows_ As String, CellWidth_ As Single, CellHeight_ As Single)
    Cols = Cols_
    Rows = Rows_
    CellWidth = CellWidth_
    CellHeight = CellHeight_
    ProgressWindow.Tag = "createTableCut.DoJob"
    ProgressWindow.Caption = "Creating table"
    ProgressWindow.Show
End Sub

Public Sub DoJob()

    ActiveDocument.Unit = cdrMillimeter
    Optimization = True
    Dim pctCompl As Single, total As Single, done As Single
    Dim width As Single, Height As Single, ExpectedWidth As Double, ExpectedHeight As Double
    Dim PosX As Double, PosY As Double
    Dim sel As Shape, table As Shape
    Set sel = ActiveDocument.Selection
    '5%
    ActiveDocument.ActiveShape.Name = "FRAME"
    width = sel.SizeWidth + 0.2
    Height = sel.SizeHeight + 0.2
    ExpectedWidth = CellWidth * Cols
    ExpectedHeight = CellHeight * Rows
    PosX = sel.PositionX - 0.2
    PosY = sel.PositionY + 0.2
    pctCompl = 5 / 100
    ProgressWindow.Progress pctCompl * 100, ProgressWindow.Frame.width * pctCompl
    'Starting
    '25%
    Set table = ActiveLayer.CreateCustomShape("Table", 1, 10, 5, 7, Cols, Rows)
    pctCompl = pctCompl + 25 / 100
    ProgressWindow.Progress pctCompl * 100, ProgressWindow.Frame.width * pctCompl
    '20%
    table.SetSize width, Height
    pctCompl = pctCompl + 20 / 100
    ProgressWindow.Progress pctCompl * 100, ProgressWindow.Frame.width * pctCompl
    '20%
    table.SetPosition PosX, PosY
    pctCompl = pctCompl + 20 / 100
    ProgressWindow.Progress pctCompl * 100, ProgressWindow.Frame.width * pctCompl
    '20%
    table.ConvertToCurves
    pctCompl = pctCompl + 20 / 100
    ProgressWindow.Progress pctCompl * 100, ProgressWindow.Frame.width * pctCompl
    '5%
    table.SetSizeEx table.CenterX, table.CenterY, ExpectedWidth + 0.2, ExpectedHeight + 0.2
    pctCompl = pctCompl + 5 / 100
    ProgressWindow.Progress pctCompl * 100, ProgressWindow.Frame.width * pctCompl
    '5%
    ActiveDocument.Selection.Group
    ActiveDocument.ActiveShape.Name = "CUT"
    ActiveLayer.FindShapes(Name:="FRAME").CreateSelection
    ActiveDocument.Selection.Delete
    ActiveLayer.FindShapes(Name:="CUT").CreateSelection
    pctCompl = pctCompl + 5 / 100
    ProgressWindow.Progress pctCompl * 100, ProgressWindow.Frame.width * pctCompl
    Unload ProgressWindow
    
    correctTable.Start
    
    Dim cutLength As Double
    
    cutLength = (Cols * CellWidth + 0.2) * (Rows + 1) + (Rows * CellHeight + 0.2) * (Cols + 1)
    
    Dim obj As New MSForms.DataObject
    obj.SetText CellWidth & "x" & CellHeight & "mm_CUT=" & Round(cutLength) & "mm_" & (Cols * Rows) & " sht"
    obj.PutInClipboard
    
    Optimization = False
    ActiveWindow.Refresh
    ActiveSelection.Ungroup
    Unload CreateTable
    
End Sub
