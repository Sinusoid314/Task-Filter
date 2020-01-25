Attribute VB_Name = "Module1"
Public Const ID_CLASS = 1
Public Const ID_TEXT = 2
Public Const ID_BOTH = 3

Public passing As Boolean

Public taskList As New Collection

Public Const WM_DESTROY = &H2
Public Const WM_CLOSE = &H10

Public Type POINTAPI
        x As Long
        y As Long
End Type

Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long


Public Sub LoadTaskData()

Dim tmpNum As Long
Dim tmpCN As String
Dim tmpWT As String
Dim tmpMode As Integer
Dim n As Long

'Load task list data from file
If Dir(App.Path & "\task.dat") <> "" Then
    Open App.Path & "\task.dat" For Input As #1
        Input #1, tmpNum
        For n = 1 To tmpNum
            Input #1, tmpCN
            Input #1, tmpWT
            Input #1, tmpMode
            taskList.Add New TaskClass
            With taskList.Item(taskList.Count)
                .className = tmpCN
                .winText = tmpWT
                .idMode = tmpMode
            End With
            mainWin.taskListBox.AddItem tmpCN
        Next n
    Close #1
End If

End Sub

Public Sub Main()

Load mainWin

If Command = "start" Then
'Start the monitor and hide the program
    passing = False
    App.TaskVisible = False
    mainWin.checkTimer.Enabled = True
    mainWin.Hide
Else
'Start the program normally
    mainWin.Show
End If

End Sub


Public Sub SaveTaskData()

'Write task list data to file
Open App.Path & "\task.dat" For Output As #1
    Print #1, taskList.Count
    For n = 1 To taskList.Count
        Print #1, taskList.Item(n).className
        Print #1, taskList.Item(n).winText
        Print #1, taskList.Item(n).idMode
    Next n
Close #1

End Sub


