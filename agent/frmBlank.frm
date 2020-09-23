VERSION 5.00
Begin VB.Form frmBlank 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrExitProgram 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4080
      Top             =   2640
   End
End
Attribute VB_Name = "frmBlank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmMain.Show , frmBlank
End Sub

Private Sub tmrExitProgram_Timer()
Unload Me
End
End Sub
