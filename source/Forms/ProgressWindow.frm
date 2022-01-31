VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ProgressWindow 
   Caption         =   "Processing"
   ClientHeight    =   1155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5790
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
