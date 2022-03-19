VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreateTableForCut_Options 
   Caption         =   "Create cutting table"
   ClientHeight    =   5928
   ClientLeft      =   120
   ClientTop       =   468
   ClientWidth     =   3840
   OleObjectBlob   =   "CreateTableForCut_Options.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CreateTableForCut_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ColsCountMem, RowsCountMem As String
Public CellW, CellH, OverC As String
Public frameWidth, frameHeight As Double
Public ResChk As Boolean

Private Sub ColumnsCount_Change()
  If IsNumeric(ColumnsCount.Text) Then
    If ColumnsCount.Text < 1 Then
        ColumnsCount.Text = 1
    End If
    ColsCountMem = ColumnsCount.Text
  ElseIf ColumnsCount.Text = "" Then
    ColumnsCount.Text = 1
  Else
    ColumnsCount.Text = ColsCountMem
  End If
  adjustColumnsCount.Value = 0
  ResChk = ParamsCheck
End Sub

Private Sub ColumnsCount_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ResChk = ParamsCheck
    If KeyCode = 13 And ResChk Then
        Call CreateButton_Click
    End If
End Sub

Private Sub RowsCount_Change()
  If IsNumeric(RowsCount.Text) Then
    If RowsCount.Text < 1 Then
        RowsCount.Text = 1
    End If
    RowsCountMem = RowsCount.Value
  ElseIf RowsCount.Text = "" Then
    RowsCount.Text = 1
  Else
    RowsCount.Text = RowsCountMem
  End If
  adjustRowsCount.Value = 0
  ResChk = ParamsCheck
End Sub

Private Sub RowsCount_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ResChk = ParamsCheck
    If KeyCode = 13 And ResChk Then
        Call CreateButton_Click
    End If
End Sub

Private Sub CellWidth_Change()
  Dim textval As String
  Dim oRE, bMatch
  Set oRE = New RegExp
  oRE.Pattern = "^(\d+)([\.,]?\d*)?$"
  textval = CellWidth.Value
  bMatch = oRE.test(textval)
  If bMatch Then
    CellW = textval
  Else
    CellWidth.Text = CellW
  End If
  adjustCellWidth.Value = 0
  If CellWidth.Text <> 0 Then ColumnsCount.Text = ItemsPerSize(CellWidth.Text, True)
  ResChk = ParamsCheck
End Sub

Private Sub CellWidth_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ResChk = ParamsCheck
    If KeyCode = 13 And ResChk Then 'Enter
        Call CreateButton_Click
    End If
End Sub

Private Sub CellHeight_Change()
  Dim textval As String
  Dim oRE, bMatch
  Set oRE = New RegExp
  oRE.Pattern = "^(\d+)([\.,]?\d*)?$"
  textval = CellHeight.Value
  bMatch = oRE.test(textval)
  If bMatch Then
    CellH = textval
  Else
    CellHeight.Text = CellH
  End If
  adjustCellHeight.Value = 0
  If CellHeight.Text <> 0 Then RowsCount.Text = ItemsPerSize(CellHeight.Text, False)
  ResChk = ParamsCheck
End Sub
Private Sub CellHeight_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ResChk = ParamsCheck
    If KeyCode = 13 And ResChk Then
        Call CreateButton_Click
    End If
End Sub
Private Sub Overcut_Change()
  Dim textval As String
  Dim oRE, bMatch
  Set oRE = New RegExp
  oRE.Pattern = "^(\d+)([\.,]?\d*)?$"
  If Overcut.Value = "" Then Overcut.Value = 0
  textval = Overcut.Value
  bMatch = oRE.test(textval)
  If bMatch And textval <= 10 Then
    OverC = textval
  Else
    Overcut.Text = OverC
  End If
  adjustOvercut.Value = 0
  If Overcut.Text <> "" Then
    RowsCount.Text = ItemsPerSize(CellHeight.Text, False)
    ColumnsCount.Text = ItemsPerSize(CellWidth.Text, True)
  End If
  ResChk = ParamsCheck
End Sub
Private Sub Overcut_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ResChk = ParamsCheck
    If KeyCode = 13 And ResChk Then
        Call CreateButton_Click
    End If
End Sub
Private Function ItemsPerSize(itemSize, width As Boolean) As Double
    Dim sel As Shape
    Dim SelSize As Double
    Dim SelWidth As Double, SelHeight As Double
    SelWidth = frameWidth - Overcut.Value * 2
    SelHeight = frameHeight - Overcut.Value * 2
    
    If width Then
        SelSize = SelWidth - Overcut.Value * 2
    Else
        SelSize = SelHeight - Overcut.Value * 2
    End If
    
    If itemSize <> 0 Then ItemsPerSize = Int(SelSize / itemSize)
End Function
Private Sub adjustRowsCount_Change()
    RowsCount.Text = RowsCount.Value + adjustRowsCount.Value
End Sub

Private Sub adjustColumnsCount_Change()
    ColumnsCount.Text = ColumnsCount.Value + adjustColumnsCount.Value
End Sub

Private Sub adjustCellWidth_Change()
    CellWidth.Text = CellWidth.Value + adjustCellWidth.Value
End Sub

Private Sub adjustCellHeight_Change()
    CellHeight.Text = CellHeight.Value + adjustCellHeight.Value
End Sub
Private Sub adjustOvercut_Change()
    Overcut.Text = Overcut.Value + adjustOvercut.Value / 10
End Sub
Private Function ParamsCheck() As Boolean
    Dim sel As Shape
    Dim SelWidth As Double, SelHeight As Double
    Dim ExpectedWidth As Double, ExpectedHeight As Double

    SelWidth = frameWidth - Overcut.Value * 2
    SelHeight = frameHeight - Overcut.Value * 2
    ExpectedWidth = CellWidth.Text * ColumnsCount.Text
    ExpectedHeight = CellHeight.Text * RowsCount.Text
    If ExpectedWidth > SelWidth And ExpectedHeight > SelHeight Then
        Info.Caption = "Expected table size is larger, than selected area"
        Info.ForeColor = &HFF
        Info.Visible = True
        CreateButton.Enabled = False
        ParamsCheck = False
    ElseIf ExpectedWidth > SelWidth Then
        Info.Caption = "Expected table width is larger, than selected area"
        Info.ForeColor = &HFF
        Info.Visible = True
        CreateButton.Enabled = False
        ParamsCheck = False
    ElseIf ExpectedHeight > SelHeight Then
        Info.Caption = "Expected table height is larger, than selected area"
        Info.ForeColor = &HFF
        Info.Visible = True
        CreateButton.Enabled = False
        ParamsCheck = False
    Else
        If CellWidth.Text > 0 And CellHeight.Text > 0 Then
            Info.Caption = "Cell count = " & (RowsCount.Text * ColumnsCount.Text)
            Info.ForeColor = &H0
            Info.Visible = True
            CreateButton.Enabled = True
            ParamsCheck = True
        Else
            Info.Caption = "Not enough data"
            Info.ForeColor = &HC0C0C0
            Info.Visible = True
            CreateButton.Enabled = False
            ParamsCheck = False
        End If
    End If
End Function
Private Sub CreateButton_Click()
    CreateButton.Enabled = False
    CreateButton.Caption = "Processing..."
    ProgressWindow.Tag = "CreateTableForCut.DoJob"
    ProgressWindow.Caption = "Creating table"
    ProgressWindow.Show
End Sub

Private Sub UserForm_Initialize()
    ActiveDocument.Unit = cdrMillimeter
    With ColumnsCount
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    frameWidth = FRAME_SIZE.RightX - FRAME_SIZE.LeftX
    frameHeight = FRAME_SIZE.TopY - FRAME_SIZE.BottomY
    CellW = CellWidth.Text
    CellH = CellHeight.Text
    OverC = Overcut.Text
    ColsCountMem = ColumnsCount.Text
    RowsCountMem = RowsCount.Text
    CellWidth.SetFocus
    CellWidth.SelStart = 0
    CellWidth.SelLength = Len(CellW)
End Sub

