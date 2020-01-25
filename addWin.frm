VERSION 5.00
Begin VB.Form addWin 
   Caption         =   "Add Task"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   375
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer captureTimer 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   4140
      Top             =   4830
   End
   Begin VB.Frame optionFrame 
      Caption         =   "Options: "
      Height          =   1845
      Left            =   300
      TabIndex        =   3
      Top             =   2775
      Width           =   4140
      Begin VB.OptionButton bothBtn 
         Caption         =   "ID based on BOTH"
         Height          =   285
         Left            =   795
         TabIndex        =   6
         Top             =   1320
         Width           =   2955
      End
      Begin VB.OptionButton winClassBtn 
         Caption         =   "ID based on Class Name"
         Height          =   285
         Left            =   795
         TabIndex        =   5
         Top             =   330
         Value           =   -1  'True
         Width           =   2955
      End
      Begin VB.OptionButton winTextBtn 
         Caption         =   "ID based on Window Text"
         Height          =   285
         Left            =   795
         TabIndex        =   4
         Top             =   825
         Width           =   2955
      End
   End
   Begin VB.CommandButton captureBtn 
      Caption         =   "Capture"
      Height          =   435
      Left            =   870
      TabIndex        =   2
      Top             =   1830
      Width           =   3285
   End
   Begin VB.CommandButton okBtn 
      Caption         =   "OK"
      Height          =   495
      Left            =   855
      TabIndex        =   1
      Top             =   4980
      Width           =   1440
   End
   Begin VB.CommandButton cancelBtn 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2460
      TabIndex        =   0
      Top             =   4980
      Width           =   1440
   End
   Begin VB.Frame infoFrame 
      Caption         =   "Task Info: "
      Height          =   1605
      Left            =   150
      TabIndex        =   7
      Top             =   150
      Width           =   4470
      Begin VB.TextBox winClassEdit 
         Height          =   300
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   345
         Width           =   2850
      End
      Begin VB.TextBox winTextEdit 
         Height          =   300
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1050
         Width           =   2850
      End
      Begin VB.Label Label1 
         Caption         =   "Class Name:"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   375
         Width           =   945
      End
      Begin VB.Label Label2 
         Caption         =   "Window Text:"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1080
         Width           =   990
      End
   End
End
Attribute VB_Name = "addWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cancelBtn_Click()

mainWin.Show
Unload Me

End Sub

Private Sub captureBtn_Click()

If captureTimer.Enabled Then
    okBtn.Enabled = True
    cancelBtn.Enabled = True
    infoFrame.Enabled = True
    optionFrame.Enabled = True
    captureBtn.Caption = "Capture"
    captureTimer.Enabled = False
Else
    okBtn.Enabled = False
    cancelBtn.Enabled = False
    infoFrame.Enabled = False
    optionFrame.Enabled = False
    captureBtn.Caption = "Stop"
    captureBtn.SetFocus
    captureTimer.Enabled = True
End If

End Sub


Private Sub captureTimer_Timer()

Dim hWin As Long
Dim curPos As POINTAPI
Dim cName As String
Dim cnLen As Long
Dim wText As String
Dim wtLen As Long

hWin = 0

'Get window under cursor
GetCursorPos curPos
hWin = WindowFromPoint(curPos.x, curPos.y)
If hWin = 0 Then Exit Sub

'Get class name
cName = Space(255)
cnLen = GetClassName(hWin, cName, Len(cName))
cName = Left(cName, cnLen)
winClassEdit.Text = cName

'Get window text
wtLen = GetWindowTextLength(hWin)
wText = Space(wtLen)
GetWindowText hWin, wText, wtLen + 1
winTextEdit.Text = wText

End Sub

Private Sub okBtn_Click()

Dim n As Long

If Trim(winClassEdit.Text) = "" Then
    MsgBox "Need class name!", vbCritical
    Exit Sub
End If

'Create new task object
taskList.Add New TaskClass
With taskList.Item(taskList.Count)
    .className = winClassEdit.Text
    .winText = winTextEdit.Text
    If winClassBtn.Value Then
        .idMode = ID_CLASS
    ElseIf winTextBtn.Value Then
        .idMode = ID_TEXT
    ElseIf bothBtn.Value Then
        .idMode = ID_BOTH
    End If
End With

mainWin.taskListBox.AddItem winClassEdit.Text

SaveTaskData

mainWin.Show
Unload Me

End Sub

