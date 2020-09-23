VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   7455
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox mainpic 
      Height          =   1695
      Left            =   0
      ScaleHeight     =   1635
      ScaleWidth      =   6195
      TabIndex        =   1
      Top             =   0
      Width           =   6255
      Begin VB.CommandButton cmdRun 
         Caption         =   "Insert"
         Height          =   495
         Left            =   3600
         TabIndex        =   14
         Top             =   0
         Width           =   1215
      End
      Begin VB.TextBox txtProgram 
         Height          =   285
         Left            =   720
         TabIndex        =   13
         Text            =   "C:\WINNT\system32\notepad.exe"
         Top             =   120
         Width           =   2775
      End
      Begin VB.CommandButton cmdFree 
         Caption         =   "Free"
         Height          =   495
         Left            =   4920
         TabIndex        =   12
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdTerminate 
         Caption         =   "Terminate"
         Height          =   315
         Left            =   3960
         TabIndex        =   11
         Top             =   960
         Width           =   1155
      End
      Begin VB.CommandButton cmdShow 
         Caption         =   "Show"
         Height          =   315
         Left            =   2640
         TabIndex        =   10
         Top             =   1320
         Width           =   1155
      End
      Begin VB.CommandButton cmdHide 
         Caption         =   "Hide"
         Height          =   315
         Left            =   2640
         TabIndex        =   9
         Top             =   960
         Width           =   1155
      End
      Begin VB.CommandButton cmdSetTitle 
         Caption         =   "Set Title"
         Height          =   315
         Left            =   2640
         TabIndex        =   8
         Top             =   600
         Width           =   1155
      End
      Begin VB.CommandButton cmdDisable 
         Caption         =   "Disable"
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Top             =   1320
         Width           =   1155
      End
      Begin VB.CommandButton cmdEnable 
         Caption         =   "Enable"
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Top             =   960
         Width           =   1155
      End
      Begin VB.CommandButton cmdFlash 
         Caption         =   "Flash"
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Top             =   600
         Width           =   1155
      End
      Begin VB.CommandButton cmdMaximize 
         Caption         =   "Maximize"
         Height          =   315
         Left            =   0
         TabIndex        =   4
         Top             =   600
         Width           =   1155
      End
      Begin VB.CommandButton cmdMinimize 
         Caption         =   "Minimize"
         Height          =   315
         Left            =   0
         TabIndex        =   3
         Top             =   960
         Width           =   1155
      End
      Begin VB.CommandButton cmdNormal 
         Caption         =   "Normal"
         Height          =   315
         Left            =   0
         TabIndex        =   2
         Top             =   1320
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Program"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   15
         Top             =   120
         Width           =   615
      End
   End
   Begin VB.PictureBox picChild 
      Height          =   1695
      Left            =   360
      ScaleHeight     =   1635
      ScaleWidth      =   5715
      TabIndex        =   0
      Top             =   2040
      Width           =   5775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'97% of the credit goes to www.vb-helper.com, allapi.net, freevbcode.com, pscode.com, elitespy, ulli, and of course everyone who submits vb code.... thank you very much
'i basically just slapped together some things i found, i hope this helps

Private Sub cmdDisable_Click()

    EnableWindow picChild.hwnd, 0

End Sub

Private Sub cmdEnable_Click()

    EnableWindow picChild.hwnd, 1

End Sub

Private Sub cmdFlash_Click()

    FlashWindow picChild.hwnd, 3

End Sub

Private Sub cmdFree_Click()

    SetParent child_hwnd, old_parent

    cmdRun.Enabled = True
    cmdFree.Enabled = False

End Sub

Private Sub cmdHide_Click()

    ShowWindow picChild.hwnd, SW_HIDE

End Sub

Private Sub cmdMaximize_Click()

    ShowWindow picChild.hwnd, SW_MAXIMIZE

End Sub

Private Sub cmdMinimize_Click()

    ShowWindow picChild.hwnd, SW_MINIMIZE

End Sub

Private Sub cmdNormal_Click()

    ShowWindow picChild.hwnd, SW_NORMAL

End Sub

Private Sub cmdRun_Click()

  Dim pid As Long
  Dim buf As String
  Dim buf_len As Long

  ' Start the program.

    pid = Shell(txtProgram.Text, vbNormalFocus)
    If pid = 0 Then
        MsgBox "Error starting program"
        Exit Sub '>---> Bottom
    End If

    ' Get the window handle.
    child_hwnd = InstanceToWnd(pid)

    ' Reparent the program so it lies inside
    ' the PictureBox.
    old_parent = SetParent(child_hwnd, picChild.hwnd)

    cmdRun.Enabled = False
    cmdFree.Enabled = True

End Sub

Private Sub cmdSetTitle_Click()

  Dim sTitle As String
  ' Ask user for new window title

    sTitle = InputBox("Enter new window title:", "This program")
    ' Set new window title
    SetWindowText picChild.hwnd, sTitle

End Sub

Private Sub cmdShow_Click()

    ShowWindow picChild.hwnd, SW_SHOW

End Sub

Private Sub cmdTerminate_Click()

    SendMessage picChild.hwnd, WM_CLOSE, 0, 0

End Sub

Private Sub Form_Load()

    SetScrollBar Me.hwnd, vbBoth
    SetScrollBar picChild.hwnd, vbBoth, True
    ControlSize picChild, True
    ControlCaption picChild, True
    ControlSysMenu picChild, True
    ControlMinBox picChild, True
    SetWindowText picChild.hwnd, "Yes, this is a picturebox!"

End Sub

Private Sub Form_Resize()

    AdjustScrollInfo Me.hwnd

End Sub

Private Sub Form_Unload(Cancel As Integer)

    cmdFree_Click
    Set Form1 = Nothing

End Sub

Private Sub picChild_Resize()

    AdjustScrollInfo picChild.hwnd

End Sub

':) Ulli's VB Code Formatter V2.16.6 (2005-Jul-24 04:28 AM) 1 + 133 = 134 Lines
