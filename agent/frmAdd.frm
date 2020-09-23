VERSION 5.00
Begin VB.Form frmAdd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Enabled         =   0   'False
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   5280
      Width           =   5535
   End
   Begin VB.TextBox txtSpeak 
      Height          =   1335
      Left            =   1320
      TabIndex        =   10
      Top             =   3720
      Width           =   4335
   End
   Begin VB.OptionButton optAction 
      Caption         =   "Speak:"
      Enabled         =   0   'False
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox mtY 
      Height          =   285
      Left            =   2760
      TabIndex        =   8
      Top             =   3225
      Width           =   615
   End
   Begin VB.TextBox mtX 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1680
      TabIndex        =   6
      Top             =   3210
      Width           =   615
   End
   Begin VB.OptionButton optAction 
      Caption         =   "Move To:"
      Enabled         =   0   'False
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   1095
   End
   Begin VB.ComboBox lstAgent 
      Height          =   315
      ItemData        =   "frmAdd.frx":0000
      Left            =   1320
      List            =   "frmAdd.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   4335
   End
   Begin VB.ListBox lstAction 
      Enabled         =   0   'False
      Height          =   2400
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   4335
   End
   Begin VB.OptionButton optAction 
      Caption         =   "Do Action:"
      Enabled         =   0   'False
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Y:"
      Height          =   255
      Left            =   2520
      TabIndex        =   7
      Top             =   3255
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "X:"
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Choose Agent:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
If optAction(0).Value = True Then
    If lstAction.SelCount < 1 Then Exit Sub
    frmMain.lstActions.AddItem lstAgent.List(lstAgent.ListIndex) & "|Action|" & lstAction.List(lstAction.ListIndex)
ElseIf optAction(1).Value = True Then
    If mtX.Text = "" Or mtY.Text = "" Then Exit Sub
    If Trim(Str(Val(mtX.Text))) <> mtX.Text Or Trim(Str(Val(mtY.Text))) <> mtY.Text Then Exit Sub
    frmMain.lstActions.AddItem lstAgent.List(lstAgent.ListIndex) & "|Move|" & mtX.Text & "x" & mtY.Text
Else
    If txtSpeak.Text = "" Then Exit Sub
    frmMain.lstActions.AddItem lstAgent.List(lstAgent.ListIndex) & "|Speak|" & txtSpeak.Text
End If
Me.Hide
End Sub

Private Sub lstAgent_Click()
If lstAgent.ListIndex < 0 Then Exit Sub
lstAction.Clear
For Each AnimationName In frmMain.msAgent.Characters(lstAgent.List(lstAgent.ListIndex)).AnimationNames
    lstAction.AddItem AnimationName
Next
optAction(0).Enabled = True
optAction(1).Enabled = True
optAction(2).Enabled = True
lstAction.Enabled = True
cmdAdd.Enabled = True
End Sub

Private Sub optAction_Click(Index As Integer)
If Index = 0 Then
    mtX.Enabled = False
    mtY.Enabled = False
    txtSpeak.Enabled = False
    Label2.Enabled = False
    Label3.Enabled = False
    lstAction.Enabled = True
ElseIf Index = 1 Then
    mtX.Enabled = True
    mtY.Enabled = True
    txtSpeak.Enabled = False
    Label2.Enabled = True
    Label3.Enabled = True
    lstAction.Enabled = False
ElseIf Index = 2 Then
    mtX.Enabled = False
    mtY.Enabled = False
    txtSpeak.Enabled = True
    Label2.Enabled = False
    Label3.Enabled = False
    lstAction.Enabled = False
End If
End Sub
