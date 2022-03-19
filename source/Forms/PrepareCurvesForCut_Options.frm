VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} PrepareCurvesForCut_Options 
   Caption         =   "Prepare shapes for cutting"
   ClientHeight    =   3696
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   5148
   OleObjectBlob   =   "PrepareCurvesForCut_Options.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "PrepareCurvesForCut_Options"
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
    OPTIMIZE_ADVANCED = AdvancedOtimizationCheck.Value
    PASSES_COUNT = Passes.ListIndex + 1
    Start.Enabled = False
    Start.Caption = "Processing..."
    Unload PrepareCurvesForCut_Options
    ProgressWindow.Caption = "Processing curves"
    ProgressWindow.Tag = "PrepareCurvesForCut.DoJob"
    ProgressWindow.Show
End Sub
Private Sub UserForm_Initialize()
    With Smoothness
        .AddItem "None", 0
        .AddItem "Low", 1
        .AddItem "Medium", 2
        .AddItem "High", 3
        .Value = "Low"
    End With
    With Fillet
        .AddItem "None", 0
        .AddItem "0,5 mm", 1
        .AddItem "0,75 mm", 2
        .AddItem "1,0 mm", 3
        .Value = "0,5 mm"
    End With
    With Passes
        .AddItem "1", 0
        .AddItem "2", 1
        .AddItem "3", 2
        .Value = "1"
    End With
End Sub
