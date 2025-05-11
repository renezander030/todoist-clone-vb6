VERSION 5.00
Begin VB.Form TaskDetail 
   Appearance      =   0  '2D
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Task Details"
   ClientHeight    =   9315
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10110
   Icon            =   "TaskDetail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox taskPriority 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   465
      TabIndex        =   4
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox taskDescriptionTxt 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7335
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   3
      Top             =   1080
      Width           =   6735
   End
   Begin VB.TextBox taskContentTxt 
      Appearance      =   0  '2D
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   2
      Top             =   120
      Width           =   6735
   End
   Begin VB.CommandButton discardbtn 
      Appearance      =   0  '2D
      BackColor       =   &H00FF8080&
      Caption         =   "Discard"
      Height          =   375
      Left            =   5040
      Style           =   1  'Grafisch
      TabIndex        =   1
      Top             =   8640
      Width           =   1215
   End
   Begin VB.CommandButton savebtn 
      Appearance      =   0  '2D
      BackColor       =   &H00FF8080&
      Caption         =   "Save"
      Height          =   375
      Left            =   6360
      Style           =   1  'Grafisch
      TabIndex        =   0
      Top             =   8640
      Width           =   1215
   End
End
Attribute VB_Name = "TaskDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public TaskContent As String, TaskDescription As String

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    If KeyCode = vbKeyReturn Then
        Call savebtn_Click
    End If
End Sub

Private Sub Form_Load()
    taskContentTxt.Text = TaskContent
    taskDescriptionTxt.Text = TaskDescription
    KeyPreview = True
End Sub

Private Sub savebtn_Click()
    ' save
    MsgBox "Saving"
End Sub
