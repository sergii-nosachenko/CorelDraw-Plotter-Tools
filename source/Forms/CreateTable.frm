VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CreateTable 
   Caption         =   "Create cutting table"
   ClientHeight    =   5160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3840
   OleObjectBlob   =   "CreateTable.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CreateTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public ColsCountMem, RowsCountMem As String
Public CellW, CellH As String
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
    RowsCountMem = RowsCount.Text
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
  oRE.Pattern = "^(\d+)([\.,]\d*)?$"
  textval = CellWidth.Text
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
  oRE.Pattern = "^(\d+)([\.,]\d*)?$"
  textval = CellHeight.Text
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
Private Function ItemsPerSize(itemSize, width As Boolean) As Double
    Dim sel As Shape
    Dim SelSize As Double
    Set sel = ActiveDocument.Selection
    
    If width Then
        SelSize = sel.SizeWidth + 0.2
    Else
        SelSize = sel.SizeHeight + 0.2
    End If
    
    ItemsPerSize = Int(SelSize / itemSize)
End Function
Private Sub adjustRowsCount_Change()
    RowsCount.Text = RowsCount.Text + adjustRowsCount.Value
    ResChk = ParamsCheck
End Sub

Private Sub adjustColumnsCount_Change()
    ColumnsCount.Text = ColumnsCount.Text + adjustColumnsCount.Value
    ResChk = ParamsCheck
End Sub

Private Sub adjustCellWidth_Change()
    CellWidth.Text = CellWidth.Text + adjustCellWidth.Value
    ResChk = ParamsCheck
End Sub

Private Sub adjustCellHeight_Change()
    CellHeight.Text = CellHeight.Text + adjustCellHeight.Value
    ResChk = ParamsCheck
End Sub
Private Function ParamsCheck() As Boolean
    Dim sel As Shape
    Dim SelWidth As Double, SelHeight As Double
    Dim ExpectedWidth As Double, ExpectedHeight As Double
    Set sel = ActiveDocument.Selection
    
    SelWidth = sel.SizeWidth + 0.2
    SelHeight = sel.SizeHeight + 0.2
    ExpectedWidth = CellWidth.Text * ColumnsCount.Text
    ExpectedHeight = CellHeight.Text * RowsCount.Text
    If ExpectedWidth > SelWidth And ExpectedHeight > SelHeight Then
        Info.Caption = "Ўирина та висота таблиц≥ б≥льш≥, н≥ж вибрана область!"
        Info.ForeColor = &HFF
        Info.Visible = True
        CreateButton.Enabled = False
        ParamsCheck = False
    ElseIf ExpectedWidth > SelWidth Then
        Info.Caption = "Ўирина таблиц≥ б≥льша, н≥ж вибрана область!"
        Info.ForeColor = &HFF
        Info.Visible = True
        CreateButton.Enabled = False
        ParamsCheck = False
    ElseIf ExpectedHeight > SelHeight Then
        Info.Caption = "¬исота таблиц≥ б≥льша, н≥ж вибрана область!"
        Info.ForeColor = &HFF
        Info.Visible = True
        CreateButton.Enabled = False
        ParamsCheck = False
    Else
        If CellWidth.Text > 0 And CellHeight.Text > 0 Then
            Info.Caption = " ≥льк≥сть €чейок =" & (RowsCount.Text * ColumnsCount.Text)
            Info.ForeColor = &H0
            Info.Visible = True
            CreateButton.Enabled = True
            ParamsCheck = True
        Else
            Info.Caption = ""
            Info.Visible = False
            CreateButton.Enabled = False
            ParamsCheck = False
        End If
    End If
End Function
Private Sub CreateButton_Click()
    CreateButton.Enabled = False
    CreateButton.Caption = "Processing..."
    CreateTableCut.Start ColumnsCount.Text, RowsCount.Text, CellWidth.Text, CellHeight.Text
End Sub

Private Sub UserForm_Initialize()
    With ColumnsCount
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    ActiveDocument.Unit = cdrMillimeter
    CellW = CellWidth.Text
    CellH = CellHeight.Text
    ColsCountMem = ColumnsCount.Text
    RowsCountMem = RowsCount.Text
    CellWidth.SetFocus
    CellWidth.SelStart = 0
    CellWidth.SelLength = Len(CellW)
End Sub

