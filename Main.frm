VERSION 5.00
Begin VB.Form KumikoChan 
   Caption         =   "I Love Kumiko Chan"
   ClientHeight    =   1560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4770
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Main.frx":048A
   ScaleHeight     =   1560
   ScaleWidth      =   4770
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtPosition 
      Height          =   2175
      Left            =   5640
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "Main.frx":07CC
      Top             =   960
      Width           =   4215
   End
   Begin VB.ComboBox ComboMode 
      Height          =   315
      Left            =   5640
      TabIndex        =   3
      Text            =   "Combo1"
      Top             =   480
      Width           =   4215
   End
   Begin VB.ComboBox SpeedCombo 
      Height          =   315
      Left            =   2160
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   240
      Width           =   2535
   End
   Begin VB.Timer Timer1 
      Left            =   2160
      Top             =   840
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   615
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "KumikoChan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Form:
'   Timer1
'   CommandButton: cmdStart
'   CommandButton: cmdStop
'   ComboBox: SpeedCombo
'   ComboBox: ComboMode
'   TextBox: txtPosition (MultiLine = True, ScrollBars = 2 for vertical)
'
' Mode 1 = Simple (single spot auto click)
' Mode 2 = Record (record path + timing between clicks, then playback loop)
' Mark spot = "A" key
' Finish recording/start playback = Enter
' Stop playback = Enter

Option Explicit

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, _
    ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const VK_RETURN = &HD
Private Const VK_A = &H41

Private Const TIMER_INTERVAL_RECORD = 50  ' 50ms for recording
Private Const TIMER_INTERVAL_PLAY = 10    ' 10ms for playback

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type ClickStep
    x As Long
    y As Long
    delay As Long  ' delay before next click
End Type

Private ClickX As Long
Private ClickY As Long

Private Steps() As ClickStep
Private StepCount As Long
Private CurrentStep As Long
Private Recording As Boolean
Private LastTick As Long
Private Playing As Boolean

Private Sub Form_Load()
    Dim i As Integer
    ' Fill speeds
    For i = 1 To 10
        SpeedCombo.AddItem i & "s"
    Next i
    SpeedCombo.AddItem "500ms"
    SpeedCombo.AddItem "100ms"
    SpeedCombo.AddItem "50ms"
    SpeedCombo.AddItem "10ms"
    SpeedCombo.AddItem "1ms"
    SpeedCombo.ListIndex = 0

    ' Fill modes
    ComboMode.AddItem "Simple Mode"
    ComboMode.AddItem "Record Mode"
    ComboMode.ListIndex = 0

    txtPosition.Text = ""
End Sub

Private Function GetIntervalFromCombo() As Long
    Dim txt As String
    txt = SpeedCombo.Text
    If InStr(txt, "ms") > 0 Then
        GetIntervalFromCombo = Val(txt)   ' already ms
    ElseIf InStr(txt, "s") > 0 Then
        GetIntervalFromCombo = Val(txt) * 1000
    Else
        GetIntervalFromCombo = 1000
    End If
End Function

Private Function APressed() As Boolean
    Static LastAPress As Long
    Dim currentTime As Long
    
    currentTime = GetTickCount()
    
    ' Debounce check - prevent multiple detections within 200ms
    If currentTime - LastAPress < 200 Then
        APressed = False
        Exit Function
    End If
    
    If (GetAsyncKeyState(VK_A) And &H8000) <> 0 Then
        APressed = True
        LastAPress = currentTime
    Else
        APressed = False
    End If
End Function

Private Sub cmdStart_Click()
    If ComboMode.ListIndex = 0 Then
        ' Simple Mode
        Dim pt As POINTAPI
        MsgBox "Move your mouse to the target spot and press ENTER."
        Do
            DoEvents
            GetCursorPos pt
            If (GetAsyncKeyState(VK_RETURN) And &H8000) <> 0 Then
                ClickX = pt.x
                ClickY = pt.y
                Exit Do
            End If
        Loop
        Timer1.Interval = GetIntervalFromCombo()
        Timer1.Enabled = True
    Else
        ' Record Mode
        MsgBox "Record Mode: Move mouse. Press 'A' to mark clicks. Press ENTER to finish and start playback."
        ReDim Steps(1 To 1)
        StepCount = 0
        Recording = True
        Playing = False
        LastTick = GetTickCount()
        txtPosition.Text = ""
        Timer1.Interval = TIMER_INTERVAL_RECORD  ' Use appropriate interval
        Timer1.Enabled = True
    End If
End Sub

Private Sub cmdStop_Click()
    Timer1.Enabled = False
    Recording = False
    Playing = False
End Sub

Private Sub Timer1_Timer()
    Dim pt As POINTAPI
    Dim t As Long

    If ComboMode.ListIndex = 0 Then
        ' --- Simple Mode ---
        GetCursorPos pt
        If pt.x <> ClickX Or pt.y <> ClickY Then
            Timer1.Enabled = False
            Exit Sub
        End If
        SetCursorPos ClickX, ClickY
        mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
        mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    Else
        ' --- Record Mode ---
        If Recording Then
            If APressed() Then
                GetCursorPos pt
                t = GetTickCount()
                StepCount = StepCount + 1
                ReDim Preserve Steps(1 To StepCount)
                Steps(StepCount).x = pt.x
                Steps(StepCount).y = pt.y
                
                ' Calculate delay from previous step (not from start)
                If StepCount = 1 Then
                    Steps(StepCount).delay = 1000  ' 1 second initial delay
                Else
                    Steps(StepCount).delay = t - LastTick
                End If
                
                LastTick = t

                ' Append to txtPosition
                txtPosition.Text = txtPosition.Text & _
                    "Step " & StepCount & ": (" & pt.x & "," & pt.y & _
                    ") Delay=" & Steps(StepCount).delay & "ms" & vbCrLf

                ' wait for release
                Do While APressed()
                    DoEvents
                Loop
            End If
            If (GetAsyncKeyState(VK_RETURN) And &H8000) <> 0 Then
                Recording = False
                Playing = True
                CurrentStep = 1
                LastTick = GetTickCount()
                Timer1.Interval = 10
            End If
        ElseIf Playing Then
            If StepCount = 0 Then Exit Sub
            If (GetAsyncKeyState(VK_RETURN) And &H8000) <> 0 Then
                Playing = False
                Timer1.Enabled = False
                Exit Sub
            End If
            t = GetTickCount()
            If t - LastTick >= Steps(CurrentStep).delay Then
                SetCursorPos Steps(CurrentStep).x, Steps(CurrentStep).y
                mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
                mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
                CurrentStep = CurrentStep + 1
                If CurrentStep > StepCount Then
                    CurrentStep = 1
                End If
                LastTick = t
            End If
        End If
    End If
End Sub


