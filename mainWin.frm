VERSION 5.00
Begin VB.Form mainWin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gleet TaskFilter"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5295
   Icon            =   "mainWin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   279
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   353
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer checkTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3855
      Top             =   2415
   End
   Begin VB.CommandButton startBtn 
      Caption         =   "Start Monitor"
      Height          =   495
      Left            =   3330
      TabIndex        =   4
      Top             =   3015
      Width           =   1620
   End
   Begin VB.CommandButton removeBtn 
      Caption         =   "Remove Task"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3330
      TabIndex        =   2
      Top             =   1635
      Width           =   1620
   End
   Begin VB.ListBox taskListBox 
      Appearance      =   0  'Flat
      Height          =   3540
      Left            =   210
      TabIndex        =   1
      Top             =   420
      Width           =   2865
   End
   Begin VB.CommandButton addBtn 
      Caption         =   "Add Task"
      Height          =   495
      Left            =   3330
      TabIndex        =   0
      Top             =   825
      Width           =   1620
   End
   Begin VB.Label Label1 
      Caption         =   "Tasks:"
      Height          =   240
      Left            =   225
      TabIndex        =   3
      Top             =   150
      Width           =   750
   End
End
Attribute VB_Name = "mainWin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub addBtn_Click()

Load addWin
addWin.Show
mainWin.Hide

End Sub

Private Sub checkTimer_Timer()

Dim cName As String
Dim cnLen As Long
Dim wText As String
Dim wtLen As Long
Dim hWin As Long
Dim n As Long
Dim passStr As String

'Pause timer
'checkTimer.Enabled = False

'Get active window
hWin = GetForegroundWindow()

'Get class name
cName = Space(255)
cnLen = GetClassName(hWin, cName, Len(cName))
cName = Left(cName, cnLen)

'Get window text
wtLen = GetWindowTextLength(hWin)
wText = Space(wtLen)
GetWindowText hWin, wText, wtLen + 1

'Check for exit condition
If Not passing Then
    If wText = "Stop Zee Monitor" Then
        passing = True
        passStr = InputBox("Enter deactivation password:", "Gleet TaskFilter")
        passing = False
        If passStr = "please" Then
            App.TaskVisible = True
            mainWin.Show
            checkTimer.Enabled = False
            Exit Sub
        Else
            MsgBox "Incorrect password!", vbCritical, "Gleet TaskFilter"
        End If
    End If
End If

'Check against selected tasked
For n = 1 To taskList.Count
    With taskList.Item(n)
    Select Case .idMode
        Case ID_CLASS
            If cName = .className Then
                PostMessage hWin, WM_CLOSE, 0, 0
                MsgBox "UNAUTHORIZED TASK DETECTED!!", vbCritical, "Gleet TaskFilter"
                Exit For
            End If
        Case ID_TEXT
            If wText = .winText Then
                PostMessage hWin, WM_CLOSE, 0, 0
                MsgBox "UNAUTHORIZED TASK DETECTED!!", vbCritical, "Gleet TaskFilter"
                Exit For
            End If
        Case ID_BOTH
            If (cName = .className) And (wText = .winText) Then
                PostMessage hWin, WM_CLOSE, 0, 0
                MsgBox "UNAUTHORIZED TASK DETECTED!!", vbCritical, "Gleet TaskFilter"
                Exit For
            End If
    End Select
    End With
Next n

'Restart timer
'checkTimer.Enabled = True

End Sub

Private Sub Form_Load()

LoadTaskData

End Sub





Private Sub removeBtn_Click()

If List1.ListIndex < 0 Then Exit Sub

If MsgBox("Delete task '" & taskListBox.List(taskListBox.ListIndex) & "' ?", vbYesNo) = vbNo Then
    Exit Sub
End If

taskList.Remove taskListBox.ListIndex + 1
taskListBox.RemoveItem taskListBox.ListIndex

SaveTaskData

removeBtn.Enabled = False

End Sub

Private Sub startBtn_Click()

passing = False
App.TaskVisible = False
checkTimer.Enabled = True
mainWin.Hide

End Sub


Private Sub taskListBox_Click()
    removeBtn.Enabled = True
End Sub


