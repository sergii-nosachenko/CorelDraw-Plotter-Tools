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
    If Smoothness.ListIndex = 0 Then SmoothnessLevel = 0
    If Smoothness.ListIndex = 2 Then SmoothnessLevel = 1
    If Smoothness.ListIndex = 3 Then SmoothnessLevel = 4
    If Smoothness.ListIndex = 4 Then SmoothnessLevel = 7
    If Fillet.ListIndex = 0 Then FilletValue = 0
    If Fillet.ListIndex = 1 Then FilletValue = 0.5
    If Fillet.ListIndex = 2 Then FilletValue = 0.75
    If Fillet.ListIndex = 3 Then FilletValue = 1
    If AdvancedOtimizationCheck.Value = True Then
        Otimization = True
    Else
        Otimization = False
    End If
    PassesCount = Passes.ListIndex + 1
    Start.Enabled = False
    Start.Caption = "Processing..."
    ProgressWindow.Caption = "Processing curves"
    ProgressWindow.Tag = "PrepareCurvesForCut.DoJob"
    ProgressWindow.Show
End Sub
Private Sub UserForm_Initialize()
    With Smoothness
        .AddItem "��� ������������", 0
        .AddItem "�����", 1
        .AddItem "�������", 2
        .AddItem "������", 3
        .Value = "�����"
    End With
    With Fillet
        .AddItem "��� ������������", 0
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