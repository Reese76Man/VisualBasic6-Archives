VERSION 5.00
Begin VB.MDIForm FrmMainMenu 
   BackColor       =   &H8000000C&
   Caption         =   "Quizzard One"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu MNUFile 
      Caption         =   "&File"
      Begin VB.Menu MNUExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu MNUHelp 
      Caption         =   "&Help"
      Begin VB.Menu MNUAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "FrmMainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MNUAbout_Click()
' LOADFORM FRMAbout
End Sub

Private Sub MNUExit_Click()
' EndProgram
End Sub
