VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressWindow 
   Caption         =   "Processing"
   ClientHeight    =   1155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7035
   OleObjectBlob   =   "ProgressWindow.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ProgressWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()
    If Me.Tag <> "" Then Application.GMSManager.RunMacro "Overprint", Me.Tag
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    'Prevent user from closing with the Close box in the title bar.
    If CloseMode = vbFormControlMenu Then
        Dim result As VbMsgBoxResult
        MACRO_STATUS = 2
        result = MsgBox("Точно хочеш перервати процес?", vbOKCancel, "Зупинити оптимізацію")
        If result = vbOK Then
            MACRO_STATUS = 3
            Application.Optimization = False
            Cancel = False
        Else
            MACRO_STATUS = 1
            Cancel = True
        End If
    End If
End Sub
Sub Progress(pctCompl As Single, width As Single, Optional msg As String = "")
    ProgressWindow.Text.Caption = Round(pctCompl, 0) & "% complete " & msg
    ProgressWindow.Bar.width = width
    DoEvents
End Sub
Sub CheckStatus()
    ' Wait until user accepts macro termination
    Do Until MACRO_STATUS = 1
        If MACRO_STATUS = 3 Then
            MACRO_STATUS = 0
            End
        End If
        DoEvents
    Loop
End Sub
