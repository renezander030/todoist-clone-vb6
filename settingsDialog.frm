VERSION 5.00
Begin VB.Form Settings 
   Appearance      =   0  '2D
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Settings"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "settingsDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox proxyTxt 
      Appearance      =   0  '2D
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1320
      Width           =   3855
   End
   Begin VB.TextBox apiTokenTxt 
      Appearance      =   0  '2D
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   3855
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Discard"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "Save"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label proxyCap 
      Appearance      =   0  '2D
      BackColor       =   &H80000009&
      Caption         =   "Proxy IP Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   960
      Width           =   2295
   End
   Begin VB.Label apiTokenCap 
      Appearance      =   0  '2D
      BackColor       =   &H80000009&
      Caption         =   "API Token"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    apiTokenTxt.Text = APITOKEN
    proxyTxt.Text = PROXY
End Sub

Private Sub OKButton_Click()
    'Saves
    'updates global variables
    APITOKEN = apiTokenTxt.Text
    PROXY = proxyTxt.Text
    Dim F As Integer
    F = FreeFile(0)
    Open "APIToken.txt" For Output As #F
    Print #F, apiTokenTxt.Text;
    Close #F
    Open "Proxy.txt" For Output As #F
    Print #F, proxyTxt.Text;
    Close #F
    Unload Me
End Sub
