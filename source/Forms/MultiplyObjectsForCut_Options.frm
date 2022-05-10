VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MultiplyObjectsForCut_Options 
   Caption         =   "Multiply selected contour"
   ClientHeight    =   5895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3750
   OleObjectBlob   =   "MultiplyObjectsForCut_Options.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MultiplyObjectsForCut_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public copiesL, copiesT As String
Public distanceL, distanceT As String
Public frameWidth, frameHeight As Double
Public curveLength As Double

Private Sub copiesLeft_Change()
  If IsNumeric(copiesLeft.Text) Then
    If copiesLeft.Text < 1 Then
        copiesLeft.Text = 1
    End If
    copiesL = copiesLeft.Text
  ElseIf copiesLeft.Text = "" Then
    copiesLeft.Text = 1
  Else
    copiesLeft.Text = copiesL
  End If
  adjustCopiesLeftButton.Value = 0
  Call ParamsCheck
End Sub

Private Sub copiesLeft_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        Call MultiplyButton_Click
    End If
End Sub

Private Sub copiesTop_Change()
  If IsNumeric(copiesTop.Text) Then
    If copiesTop.Text < 1 Then
        copiesTop.Text = 1
    End If
    copiesT = copiesTop.Text
  ElseIf copiesTop.Text = "" Then
    copiesTop.Text = 1
  Else
    copiesTop.Text = copiesT
  End If
  adjustCopiesTopButton.Value = 0
  Call ParamsCheck
End Sub

Private Sub copiesTop_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        Call MultiplyButton_Click
    End If
End Sub

Private Sub distanceLeft_Change()
  Dim textval As String
  Dim oRE, bMatch
  Set oRE = New RegExp
  oRE.Pattern = "^(\d+)([\.,]\d+)?$"
  textval = distanceLeft.Text
  bMatch = oRE.test(textval)
  If bMatch Then
    distanceL = textval
  Else
    distanceLeft.Text = distanceL
  End If
  adjustDistanceLeftButton.Value = 0
  If frameWidth > 0 Then
    If distanceLeft.Text > 0 Then
        copiesLeft.Text = frameWidth \ CDbl(distanceLeft.Text)
    Else
        copiesLeft.Text = 1
    End If
  End If
  Call ParamsCheck
End Sub

Private Sub distanceLeft_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        Call MultiplyButton_Click
    End If
End Sub

Private Sub distanceTop_Change()
  Dim textval As String
  Dim oRE, bMatch
  Set oRE = New RegExp
  oRE.Pattern = "^(\d+)([\.,]\d+)?$"
  textval = distanceTop.Text
  bMatch = oRE.test(textval)
  If bMatch Then
    distanceT = textval
  Else
    distanceTop.Text = distanceT
  End If
  If frameHeight > 0 Then
    If distanceTop.Text > 0 Then
        copiesTop.Text = frameHeight \ CDbl(distanceTop.Text)
    Else
        copiesTop.Text = 1
    End If
  End If
  adjustDistanceTopButton.Value = 0
  Call ParamsCheck
End Sub

Private Sub distanceTop_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        Call MultiplyButton_Click
    End If
End Sub

Private Sub adjustCopiesLeftButton_Change()
    copiesLeft.Text = copiesLeft.Text + adjustCopiesLeftButton.Value
    Call ParamsCheck
End Sub

Private Sub adjustCopiesTopButton_Change()
    copiesTop.Text = copiesTop.Text + adjustCopiesTopButton.Value
    Call ParamsCheck
End Sub

Private Sub adjustDistanceLeftButton_Change()
    distanceLeft.Text = distanceLeft.Text + adjustDistanceLeftButton.Value
    Call ParamsCheck
End Sub

Private Sub adjustDistanceTopButton_Change()
    distanceTop.Text = distanceTop.Text + adjustDistanceTopButton.Value
    Call ParamsCheck
End Sub
Private Sub ParamsCheck()
    If distanceLeft.Text > 0 And distanceTop.Text > 0 Then
        Info.Caption = "Total count = " & (copiesTop.Text * copiesLeft.Text) & vbNewLine & _
        "Art bounds size = " & copiesLeft.Text * distanceLeft.Text - (distanceLeft.Text - Round(ActiveShape.SizeWidth, 1)) & "x" & _
        copiesTop.Text * distanceTop.Text - (distanceTop.Text - Round(ActiveShape.SizeHeight, 1)) & " mm"
    Else
        Info.Caption = ""
    End If
End Sub

Private Sub MultiplyButton_Click()
    MultiplyButton.Enabled = False
    MultiplyButton.Caption = "Processing..."
    ProgressWindow.Tag = "MultiplyObjectsForCut.DoJob"
    ProgressWindow.Caption = "Processing shapes"
    ProgressWindow.Show
End Sub

Private Sub UserForm_Initialize()
    With copiesLeft
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    ActiveDocument.Unit = cdrMillimeter
    If Not FRAME_SIZE Is Nothing Then
        frameWidth = FRAME_SIZE.RightX - FRAME_SIZE.LeftX
        frameHeight = FRAME_SIZE.TopY - FRAME_SIZE.BottomY
    End If
    distanceL = 0
    distanceT = 0
    copiesL = copiesLeft.Text
    copiesT = copiesTop.Text
    distanceLeft.Text = Round(ActiveShape.SizeWidth, 1)
    distanceTop.Text = Round(ActiveShape.SizeHeight, 1)
    distanceLeft.SetFocus
    distanceLeft.SelStart = 0
    distanceLeft.SelLength = Len(distanceLeft.Text)
    Call ParamsCheck
End Sub
