VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formDD 
   Caption         =   "Draw A Door"
   ClientHeight    =   2250
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3900
   OleObjectBlob   =   "formDD.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formDD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    CommandState.StartPrimitive New classDD
End Sub
