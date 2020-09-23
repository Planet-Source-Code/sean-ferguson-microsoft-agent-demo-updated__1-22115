VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "AGENTCTL.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Microsoft Agent Demo"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstMoveItems 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton cmdDown 
      Caption         =   "Down"
      Height          =   255
      Left            =   3120
      TabIndex        =   5
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "Up"
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&Remove"
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   2880
      Width           =   2055
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   2055
   End
   Begin VB.ListBox lstActions 
      Height          =   2400
      ItemData        =   "frmMain.frx":0000
      Left            =   120
      List            =   "frmMain.frx":0002
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton cmdExec 
      Caption         =   "&Execute"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   3360
      Width           =   4455
   End
   Begin MSComDlg.CommonDialog commonDialog 
      Left            =   720
      Top             =   3600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Agent Demo Files (*.ada)|*.ada|"
   End
   Begin AgentObjectsCtl.Agent msAgent 
      Left            =   120
      Top             =   3600
      _cx             =   847
      _cy             =   847
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileClear 
         Caption         =   "&Clear All Actions"
      End
      Begin VB.Menu mnuSeperator0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open Actions"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save Actions"
      End
      Begin VB.Menu mnuSeperator1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Merlin As IAgentCtlCharacterEx
Private Peedy As IAgentCtlCharacterEx
Private Robby As IAgentCtlCharacterEx
Private Genie As IAgentCtlCharacterEx
Private addEvent As Boolean
Private isModified As Boolean

Private Sub cmdAdd_Click()
addEvent = True
frmAdd.Show , frmMain
Do Until frmAdd.Visible = False
    DoEvents
Loop
addEvent = False
Unload frmAdd
isModified = True
End Sub

Private Sub cmdDown_Click()
Dim LIndex As Integer
Dim X As Integer
If lstActions.ListIndex = lstActions.ListCount - 1 Then Exit Sub
lstMoveItems.AddItem lstActions.List(lstActions.ListIndex)
For X = (lstActions.ListIndex + 2) To (lstActions.ListCount - 1)
lstMoveItems.AddItem lstActions.List(X)
Next X
Do Until (lstActions.ListCount - 1 <= lstActions.ListIndex + 1)
lstActions.RemoveItem lstActions.ListCount - 1
Loop
LIndex = lstActions.ListIndex
lstActions.RemoveItem lstActions.ListIndex
For X = 0 To lstMoveItems.ListCount - 1
lstActions.AddItem lstMoveItems.List(X)
Next X
lstMoveItems.Clear
lstActions.ListIndex = LIndex + 1
isModified = True
End Sub

Private Sub cmdExec_Click()
Dim actionRequest, lastAgent As String
If lstActions.ListCount < 1 Then Exit Sub
Me.Hide
For X = 0 To lstActions.ListCount - 1
aD = Split(lstActions.List(X), "|")
If X > 0 And lastAgent <> aD(0) Then msAgent.Characters(aD(0)).Wait actionRequest
lastAgent = aD(0)
If aD(1) = "Action" Then
    Set actionRequest = msAgent.Characters(aD(0)).Play(aD(2))
ElseIf aD(1) = "Move" Then
    mC = Split(aD(2), "x")
    Set actionRequest = msAgent.Characters(aD(0)).MoveTo(mC(0), mC(1))
Else
    Set actionRequest = msAgent.Characters(aD(0)).Speak(aD(2))
End If
Next X
Me.Show
End Sub

Private Sub cmdRemove_Click()
If lstActions.SelCount < 1 Then Exit Sub
lstActions.RemoveItem lstActions.ListIndex
isModified = True
End Sub

Private Sub cmdUp_Click()
Dim X As Integer
If lstActions.ListIndex = 0 Then Exit Sub
lstMoveItems.AddItem lstActions.List(lstActions.ListIndex - 1)
lstActions.RemoveItem lstActions.ListIndex - 1
For X = lstActions.ListIndex + 1 To lstActions.ListCount - 1
lstMoveItems.AddItem lstActions.List(X)
Next X
Do Until (lstActions.ListCount - 1 <= lstActions.ListIndex)
lstActions.RemoveItem lstActions.ListCount - 1
Loop
For X = 0 To lstMoveItems.ListCount - 1
lstActions.AddItem lstMoveItems.List(X)
Next X
lstMoveItems.Clear
isModified = True
End Sub

Private Sub Form_Load()
msAgent.Characters.Load "Merlin", "Merlin.acs"
msAgent.Characters.Load "Peedy", "Peedy.acs"
msAgent.Characters.Load "Robby", "Robby.acs"
msAgent.Characters.Load "Genie", "Genie.acs"
Set Merlin = Me.msAgent.Characters("Merlin")
Set Peedy = Me.msAgent.Characters("Peedy")
Set Robby = Me.msAgent.Characters("Robby")
Set Genie = Me.msAgent.Characters("Genie")
Merlin.MoveTo 1, 1
Peedy.MoveTo 600, 1
Robby.MoveTo 1, 400
Genie.MoveTo 600, 400
Merlin.Show
Peedy.Show
Robby.Show
Genie.Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If lstActions.ListCount > 0 And isModified = True Then
    If MsgBox("The actions list has not been saved. Do you want to exit anyway?", vbYesNo + vbQuestion, "Exit") = vbNo Then
        Cancel = 1
    End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmBlank.tmrExitProgram.Enabled = True
End Sub

Private Sub mnuFileClear_Click()
If MsgBox("Do you want to clear all actions?", vbYesNo + vbQuestion, "Clear") = vbYes Then lstActions.Clear
End Sub

Private Sub mnuFileExit_Click()
Unload Me
End
End Sub

Private Sub mnuFileOpen_Click()
If isModified = True Then If MsgBox("The actions list has not been saved. Do you wish to continue?", vbYesNo + vbQuestion, "Exit") = vbNo Then Exit Sub
commonDialog.ShowOpen
If commonDialog.FileName = "" Then Exit Sub
On Error GoTo Skip
Open commonDialog.FileName For Input As #1
commonDialog.FileName = ""
Input #1, fileData
If fileData <> "[Agent Demo Actions]" Then MsgBox "The selected file does not contain Microsoft Agent Demo actions.": Exit Sub
isModified = False
lstActions.Clear
Do Until (EOF(1))
Input #1, fileData
lstActions.AddItem fileData
Loop
Close #1
Exit Sub

Skip:
MsgBox "An error has occured while opening the file."
End Sub

Private Sub mnuFileSave_Click()
If lstActions.ListCount < 1 Then MsgBox "There are no actions to save.": Exit Sub
commonDialog.ShowSave
If commonDialog.FileName = "" Then Exit Sub
Open commonDialog.FileName For Output As #1
commonDialog.FileName = ""
Print #1, "[Agent Demo Actions]"
For X = 0 To lstActions.ListCount - 1
Print #1, lstActions.List(X)
Next X
Close #1
isModified = False
End Sub

Private Sub msAgent_Move(ByVal CharacterID As String, ByVal X As Integer, ByVal y As Integer, ByVal Cause As Integer)
If addEvent = True Then
    If CharacterID = frmAdd.lstAgent.List(frmAdd.lstAgent.ListIndex) Then
          frmAdd.mtX.Text = X
          frmAdd.mtY.Text = y
    End If
End If
End Sub
