VERSION 5.00
Begin VB.Form NewTask 
   Appearance      =   0  '2D
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "New Task"
   ClientHeight    =   1410
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6975
   Icon            =   "NewTask.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1410
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox prioritySel 
      Appearance      =   0  '2D
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "NewTask.frx":1CCA
      Left            =   3480
      List            =   "NewTask.frx":1CDA
      Style           =   2  'Dropdown-Liste
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox dueDateTxt 
      Appearance      =   0  '2D
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Top             =   960
      Width           =   1815
   End
   Begin VB.TextBox newTaskContenttxt 
      Appearance      =   0  '2D
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   6735
   End
   Begin VB.CommandButton CancelButton 
      Appearance      =   0  '2D
      Caption         =   "Discard"
      Height          =   375
      Left            =   4320
      Style           =   1  'Grafisch
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Appearance      =   0  '2D
      BackColor       =   &H8000000A&
      Caption         =   "Add"
      Height          =   375
      Left            =   5640
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Grafisch
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label priorityCap 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      Caption         =   "Priority"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      Caption         =   "Due"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   495
   End
End
Attribute VB_Name = "NewTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    newTaskContenttxt.SetFocus
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    If KeyCode = vbKeyReturn Then
        Call OKButton_Click
    End If
End Sub

Private Sub Form_Load()
    KeyPreview = True
    prioritySel.Text = 4
End Sub

Private Sub OKButton_Click()

    Dim priority As Integer
    
    'map priorities
    priority = 1 'default lowest priority, 1, P4
    If prioritySel.Text = 2 Then
        priority = 3
    ElseIf prioritySel.Text = 3 Then
        priority = 2
    ElseIf prioritySel.Text = 1 Then
        priority = 4
    End If
    
    AddNewTask newTaskContenttxt.Text, dueDateTxt.Text, priority
    Unload Me
    frmMain.Sync
End Sub
