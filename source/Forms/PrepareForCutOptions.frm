VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PrepareForCutOptions 
   Caption         =   "Prepare Curves for Cutting"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5145
   OleObjectBlob   =   "PrepareForCutOptions.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PrepareForCutOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Start_Click()
    If Smoothness.ListIndex = 0 Then SMOOTHNESS_LEVEL = 0
    If Smoothness.ListIndex = 2 Then SMOOTHNESS_LEVEL = 1
    If Smoothness.ListIndex = 3 Then SMOOTHNESS_LEVEL = 4
    If Smoothness.ListIndex = 4 Then SMOOTHNESS_LEVEL = 7
    If Fillet.ListIndex = 0 Then FILLET_VALUE = 0
    If Fillet.ListIndex = 1 Then FILLET_VALUE = 0.5
    If Fillet.ListIndex = 2 Then FILLET_VALUE = 0.75
    If Fillet.ListIndex = 3 Then FILLET_VALUE = 1
    If AdvancedOtimizationCheck.Value = True Then
        OPTIMIZE_ADVANCED = True
    Else
        OPTIMIZE_ADVANCED = False
    End If
    PASSES_COUNT = Passes.ListIndex + 1
    Start.Enabled = False
    Start.Caption = "Processing..."
    Unload PrepareForCutOptions
    ProgressWindow.Caption = "Processing curves"
    ProgressWindow.Tag = "PrepareCurvesForCut.DoJob"
    ProgressWindow.Show
End Sub
Private Sub UserForm_Initialize()
    With Smoothness
        .AddItem "Без згладжування", 0
        .AddItem "Легка", 1
        .AddItem "Середня", 2
        .AddItem "Сильна", 3
        .Value = "Легка"
    End With
    With Fillet
        .AddItem "Без заокруглення", 0
        .AddItem "0,5", 1
        .AddItem "0,75", 2
        .AddItem "1,0", 3
        .Value = "0,5"
    End With
    With Passes
        .AddItem "1", 0
        .AddItem "2", 1
        .AddItem "3", 2
        .Value = "1"
    End With
End Sub
