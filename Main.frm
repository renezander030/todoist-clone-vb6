VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmMain 
   Appearance      =   0  '2D
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Todoist"
   ClientHeight    =   10050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15645
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10050
   ScaleWidth      =   15645
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox ShapeMinimize 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   14640
      ScaleHeight     =   465
      ScaleWidth      =   495
      TabIndex        =   13
      Top             =   0
      Width           =   525
   End
   Begin VB.PictureBox ShapeClose 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   15240
      ScaleHeight     =   465
      ScaleWidth      =   495
      TabIndex        =   12
      Top             =   0
      Width           =   525
   End
   Begin VB.PictureBox picIconTemplate 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   400
      Index           =   0
      Left            =   12240
      ScaleHeight     =   27
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   27
      TabIndex        =   9
      Top             =   1560
      Visible         =   0   'False
      Width           =   400
   End
   Begin VB.PictureBox picOverlay 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'Kein
      Height          =   495
      Left            =   12240
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   8
      Top             =   840
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.PictureBox picMenu 
      Appearance      =   0  '2D
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'Kein
      ForeColor       =   &H80000008&
      Height          =   9615
      Left            =   120
      ScaleHeight     =   9615
      ScaleWidth      =   2655
      TabIndex        =   3
      Top             =   480
      Width           =   2655
      Begin VB.Timer tmrSlide 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   600
         Top             =   3120
      End
      Begin VB.Timer SyncTimer 
         Left            =   120
         Top             =   3120
      End
      Begin VB.CommandButton newTaskBtn 
         Caption         =   "New Task"
         Height          =   615
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Width           =   2295
      End
      Begin VB.CommandButton syncBtn 
         Caption         =   "Sync"
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CommandButton settingsBtn 
         Caption         =   "Settings"
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   2400
         Width           =   2295
      End
      Begin VB.Label syncState 
         Appearance      =   0  '2D
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Sync: Unknown"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   120
         Width           =   2175
      End
   End
   Begin MSFlexGridLib.MSFlexGrid TaskList 
      Height          =   8295
      Left            =   3240
      TabIndex        =   2
      Top             =   1440
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   14631
      _Version        =   393216
      BorderStyle     =   0
      Appearance      =   0
   End
   Begin VB.CommandButton triggerMenuBtn 
      Appearance      =   0  '2D
      BackColor       =   &H00FFFFFF&
      Caption         =   ">"
      Height          =   375
      Left            =   2880
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Grafisch
      TabIndex        =   1
      Top             =   600
      Width           =   375
   End
   Begin VB.PictureBox picTitleBar 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   15585
      TabIndex        =   11
      Top             =   0
      Width           =   15615
   End
   Begin VB.Label totalTasksCap 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label title 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Today"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SYNC_CAPTION_PREFIX = "Last Sync: "
Dim oldTime As Date
Dim newTime As Date
Dim diffTime As Date

Dim secondsElapsed As Long

Private menuExpanded As Boolean
Private Const MENU_FULL_WIDTH As Integer = 1300
Private Const MENU_MIN_WIDTH As Integer = 1
Private Const SLIDE_STEP As Integer = 2560
Private slidingOut As Boolean
Private totalShift As Integer

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
    ByVal lParam As Long) As Long

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private Const GWL_STYLE = -16
Private Const WS_CAPTION = &HC00000
Private Const WS_SYSMENU = &H80000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_THICKFRAME = &H40000

Private Const GWL_EXSTYLE = -20
Private Const WS_EX_APPWINDOW = &H40000
Private Const WS_EX_TOOLWINDOW = &H80

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
    ByVal hwnd As Long, ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function RoundRect Lib "gdi32" ( _
    ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long, _
    ByVal X3 As Long, ByVal Y3 As Long) As Long

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim i As Integer
    For i = Forms.Count - 1 To 0 Step -1
        Unload Forms(i)
    Next i
End Sub

Private Sub picTitleBar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub

Public Function Sync()
    syncState.Caption = SYNC_CAPTION_PREFIX & "Syncing..."
    
    On Error Resume Next

    With TaskList
        .Cols = 2
        .ColWidth(0) = 500
        .ColWidth(1) = TaskList.Width - 500
        .Rows = 1
        .WordWrap = True
        .BackColorBkg = vbWhite
    End With

    Dim tCollection As Collection
    Set tCollection = GetTasks()
    
    ' sort by date
    Dim i As Long
    
    ReDim taskCollection(1 To tCollection.Count)
    For i = 1 To tCollection.Count
        Set taskCollection(i) = tCollection(i)
    Next i
        
    Dim j As Long
    Dim temp As Task
    
    For i = 1 To UBound(taskCollection)
        For j = i + 1 To UBound(taskCollection)
            If taskCollection(i).due.dueDate > taskCollection(j).due.dueDate And taskCollection(j).Content <> "" Then
                Set temp = taskCollection(i)
                Set taskCollection(i) = taskCollection(j)
                Set taskCollection(j) = temp
            End If
        Next j
    Next i
    
    ' updates row count
    With TaskList
        .Rows = UBound(taskCollection)
    End With
    
    TaskList.Width = TaskList.Width - 20
    
    Dim t As Task
    
    For i = 1 To UBound(taskCollection) - 1
        
        ' row configuration
        TaskList.RowHeight(i) = 800
        
        ' Priority
        With TaskList
            .Row = i
            .Col = 0
            .CellForeColor = vbWhite
            .CellBackColor = vbWhite
            .CellFontSize = 12
            .ColAlignment(0) = flexAlignCenterCenter
        End With
    
        ' Content column
        With TaskList
            .Row = i
            .Col = 1
            .CellForeColor = vbBlack
            .CellBackColor = vbWhite
            .CellFontSize = 12
            .ColAlignment(1) = flexAlignLeftTop
        End With
        
        TaskList.TextMatrix(i, 1) = taskCollection(i).Content & vbCrLf & Format(CDate(Replace(Replace(taskCollection(i).due.dueDate, "T", " "), "Z", " ")), "hh:mm")
        
        TaskList.TextMatrix(i, 0) = taskCollection(i).priority
        TaskList.GridLinesFixed = flexGridNone

        ' draw header background
        With TaskList
            .Row = 0
            .Col = 0
            .CellBackColor = vbWhite
            .Col = 1
            .CellBackColor = vbWhite
        End With
        
        DrawIcons
        
    Next
    
    ' update sync state
    syncState.Caption = SYNC_CAPTION_PREFIX & Format(Now, "hh:mm")

    ' scroll to top
    TaskList.TopRow = 1

    ' update total tasks value
    totalTasksCap.Caption = UBound(taskCollection) - 1 & " Tasks"

End Function

Private Sub DrawIcons()
    Dim i As Integer
    Dim iconBox As PictureBox
    Dim cx As Single, cy As Single, r As Single
    Dim cellLeft As Single, cellTop As Single
    Dim priorityValue As Integer
    Dim colorValue As Long
    
    ' Remove old icon overlays
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        If TypeOf ctrl Is PictureBox Then
            If ctrl.Name Like "picIcon*" Then
                If ctrl.Index > 0 Then
                    Unload ctrl
                End If
            End If
        End If
    Next
    
    ' Loop over each row
    For i = 1 To TaskList.Rows - 1
        ' Create new PictureBox from template
        Load picIconTemplate(i)
        Set iconBox = picIconTemplate(i)
        With iconBox
            .Visible = True
            .AutoRedraw = True
            .ScaleMode = vbPixels
            .BorderStyle = 0
            .Width = 400
            .Height = 400
            
            ' Get position of cell (row i, col 0)
            TaskList.Row = i
            TaskList.Col = 0
            cellLeft = TaskList.cellLeft
            cellTop = TaskList.cellTop
            
            .Move TaskList.Left + cellLeft + 30, TaskList.Top + cellTop + 30
            
            priorityValue = TaskList.TextMatrix(i, 0)
                        
            Select Case priorityValue
                Case 1
                    colorValue = vbBlack ' red
                Case 2
                    colorValue = vbBlue
                Case 3
                    colorValue = RGB(255, 165, 0) ' orange
                Case 4
                    colorValue = RGB(200, 0, 0)
                Case Else
                    colorValue = vbBlack ' black for unknown
            End Select
            
            ' Draw the circle
            .ForeColor = colorValue
            .Cls
            .ZOrder (0)
            r = 10
            cx = .ScaleWidth / 2
            cy = .ScaleHeight / 2
        End With
        
        iconBox.Circle (cx, cy), r
    Next i
End Sub

Private Sub picIconTemplate_Click(Index As Integer)
    CloseTask taskCollection(Index).id
    Sync
End Sub

Private Sub Form_Activate()
'    SetForeGroundWindow Me.hwnd
    'Me.Visible = False
    'Me.Visible = True
End Sub

Private Sub DrawRoundedButton(pic As PictureBox, Text As String)
    
    Dim radiusX As Long, radiusY As Long
    radiusX = 0
    radiusY = 0
    
    pic.Cls ' Clear previous drawings
    pic.FillStyle = vbFSSolid
    pic.FillColor = RGB(246, 248, 249) ' Background color
    pic.ForeColor = pic.FillColor

    ' Draw the rounded rectangle
    RoundRect pic.hdc, 0, 0, pic.ScaleWidth, pic.ScaleHeight, radiusX, radiusY

    
    pic.ForeColor = vbBlack   ' Text color
    pic.Font.Size = 12

    ' Draw the text centered
    pic.CurrentX = (pic.ScaleWidth - pic.TextWidth(Text)) \ 2
    pic.CurrentY = (pic.ScaleHeight - pic.TextHeight(Text)) \ 2
    pic.Print Text
End Sub

Private Sub Form_Load()
       
    ' update color for side bar
    picMenu.BackColor = RGB(246, 248, 249)
    picMenu.Left = 0
    picMenu.Height = Me.Height
    
    ' draw buttons
    ShapeClose.AutoRedraw = True
    ShapeClose.BorderStyle = 0
    ShapeClose.ScaleMode = 3
    ShapeClose.BackColor = vbWhite
    ShapeClose.Visible = True

    ShapeMinimize.AutoRedraw = True
    ShapeMinimize.BorderStyle = 0
    ShapeMinimize.ScaleMode = 3
    ShapeMinimize.BackColor = vbWhite
    ShapeMinimize.Visible = True


    DrawRoundedButton ShapeClose, "X"
    DrawRoundedButton ShapeMinimize, "-"


    'frmOwner.Show
    frmMain.Show vbModeless, frmOwner
    
    On Error Resume Next
    
    ' hide title bar
    Dim lStyle As Long, lExStyle As Long
    lStyle = GetWindowLong(Me.hwnd, GWL_STYLE)
    
    lStyle = lStyle And Not WS_CAPTION
    lStyle = lStyle And Not WS_SYSMENU
    lStyle = lStyle And Not WS_THICKFRAME
    
    SetWindowLong Me.hwnd, GWL_STYLE, lStyle
    
    ' extended style to show in taskbar
    lExStyle = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    lExStyle = (lExStyle Or WS_EX_APPWINDOW) And Not WS_EX_TOOLWINDOW
    
    SetWindowLong Me.hwnd, GWL_EXSTYLE, lExStyle

    DrawMenuBar Me.hwnd
    
    App.TaskVisible = True
        
    ' draw own
    picTitleBar.BackColor = RGB(246, 248, 249)
    picTitleBar.Height = 500
    picTitleBar.Width = frmMain.Width
    picTitleBar.BorderStyle = 0
    
    KeyPreview = True
    Dim F As Integer
    F = FreeFile(0)
    Open "APIToken.txt" For Input As #F
    APITOKEN = Input$(LOF(F), #F)
    Close #F
    
    Open "Proxy.txt" For Input As #F
    PROXY = Input$(LOF(F), #F)
    Close #F
        
    With TaskList
        .Cols = 2
        .ColWidth(0) = 500
        .ColWidth(1) = TaskList.Width - 500
        .Rows = 1
        .WordWrap = True
    End With
    
    With picOverlay
        .Move 100, 100, 1000, 1000
        .Visible = False
        .AutoRedraw = True
        .ScaleMode = vbPixels
        .BackColor = vbWhite
        .BorderStyle = 0 ' 0 for no border
        .Width = 1000
        .Height = 1000
        .Top = 100
        .Left = 100
    End With
    
    'initial sync
    If APITOKEN = "" Or PROXY = "" Then
        MsgBox "API Token or Proxy not set"
    Else
        Sync
    End If
  
    'sync on timer
    SyncTimer.Enabled = True
    SyncTimer.Interval = 1000 ' 1 second
    secondsElapsed = 0
  
    ' side bar
    menuExpanded = True
    picMenu.Width = SLIDE_STEP
    picMenu.Visible = True
  
End Sub

Private Sub ShapeClose_Click()
    Unload frmOwner
    Unload frmMain
End Sub

Private Sub ShapeClose_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "Mouse Moved!"
End Sub

Private Sub ShapeMinimize_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub ShapeMinimize_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Debug.Print "Mouse Moved Minimized!"
End Sub

Private Sub newTaskBtn_Click()
    NewTask.Show
End Sub

Private Sub settingsBtn_Click()
    Settings.Show
End Sub


Private Sub syncBtn_Click()
  Call Sync
End Sub

Private Sub SyncTimer_Timer()
    secondsElapsed = secondsElapsed + 1
    If secondsElapsed >= 600 Then ' 10 minutes
        Call Sync
        secondsElapsed = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyM Then
        Call triggerMenuBtn_Click
    End If

    If KeyCode = vbKeyQ Then
        Call newTaskBtn_Click
    End If
    
    If KeyCode = vbKeyS Then
        Call syncBtn_Click
    End If

    If (Shift And vbAltMask) And KeyCode = vbKeyF4 Then
        KeyCode = 0 ' Cancel default behavior
        Unload frmOwner
        Unload frmMain
    End If
End Sub

Private Sub TaskList_Click()
    Dim clickedRow As Integer, rawText As String
    clickedRow = TaskList.Row
    TaskDetail.TaskContent = taskCollection(clickedRow).Content
    
    ' fix line break
    rawText = Replace(taskCollection(clickedRow).Description, vbLf, vbCrLf)
    TaskDetail.TaskDescription = rawText
    
    TaskDetail.Show
End Sub

Private Sub tmrSlide_Timer()
    Dim i As Integer
    Dim Shift As Integer
    
    If slidingOut Then
        ' slide in (hide menu)
        picMenu.Width = 1
        ShiftControls -SLIDE_STEP
        If picMenu.Width > SLIDE_STEP Then
            Shift = -SLIDE_STEP
            picMenu.Width = MENU_MIN_WIDTH
            triggerMenuBtn.Left = picMenu.Left + picMenu.Width
            'ShiftControls -totalShift
        Else
            picMenu.Visible = False
            tmrSlide.Enabled = False
            menuExpanded = False
            
            ' Reverse all the shift
            'ShiftControls -totalShift
            totalShift = 0
        End If
        triggerMenuBtn.Left = picMenu.Left + picMenu.Width
    Else
        ' slide out (show menu)
        If Not picMenu.Visible Then
            picMenu.Width = 1 ' Start small, but not 0
            triggerMenuBtn.Left = picMenu.Left + picMenu.Width
            picMenu.Visible = True
        End If
        
        If picMenu.Width < MENU_FULL_WIDTH Then
            Shift = SLIDE_STEP
            picMenu.Width = picMenu.Width + Shift
            triggerMenuBtn.Left = picMenu.Left + picMenu.Width
            ShiftControls SLIDE_STEP
        Else
            tmrSlide.Enabled = False
            totalShift = 0
        End If
        menuExpanded = True
    End If
End Sub

Private Sub ShiftControls(offset As Integer)
    Dim ctrl As Control
    For Each ctrl In Me.Controls
        On Error Resume Next
        If ctrl.Name <> "picMenu" _
            And ctrl.Name <> "triggerMenuBtn" _
            And ctrl.Name <> "ShapeClose" _
            And ctrl.Name <> "ShapeMinimize" _
            Then
                ctrl.Left = ctrl.Left + offset
        End If
    Next ctrl
    
    totalShift = totalShift + offset
End Sub

Private Sub triggerMenuBtn_Click()
    slidingOut = menuExpanded
    tmrSlide.Enabled = True
End Sub
