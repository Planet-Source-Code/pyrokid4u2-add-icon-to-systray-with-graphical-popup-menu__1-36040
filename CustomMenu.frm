VERSION 5.00
Begin VB.Form CustomMenu 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   0
   ClientLeft      =   4950
   ClientTop       =   2790
   ClientWidth     =   1830
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "CustomMenu.frx":0000
   ScaleHeight     =   0
   ScaleWidth      =   1830
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1320
      Top             =   3000
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   360
      Top             =   0
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H000000FF&
      Caption         =   "Close Menu"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000040C0&
      Caption         =   "Open Notepad"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H000040C0&
      Caption         =   "Open URL"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1200
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H000040C0&
      Caption         =   "Close CD Tray"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   840
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H000040C0&
      Caption         =   "Open CD Tray"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H000000FF&
      Caption         =   "Exit "
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Line Line16 
      BorderColor     =   &H00E0E0E0&
      X1              =   1680
      X2              =   120
      Y1              =   2895
      Y2              =   2895
   End
   Begin VB.Line Line15 
      BorderColor     =   &H00808080&
      X1              =   1680
      X2              =   120
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line14 
      BorderColor     =   &H00E0E0E0&
      X1              =   1695
      X2              =   1695
      Y1              =   2880
      Y2              =   2040
   End
   Begin VB.Line Line13 
      BorderColor     =   &H00808080&
      X1              =   1680
      X2              =   1680
      Y1              =   2880
      Y2              =   2040
   End
   Begin VB.Line Line12 
      BorderColor     =   &H00808080&
      X1              =   135
      X2              =   135
      Y1              =   2040
      Y2              =   2880
   End
   Begin VB.Line Line11 
      BorderColor     =   &H00E0E0E0&
      X1              =   120
      X2              =   120
      Y1              =   2040
      Y2              =   2880
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   1680
      Y1              =   2055
      Y2              =   2055
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00E0E0E0&
      X1              =   120
      X2              =   1680
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   135
      X2              =   135
      Y1              =   360
      Y2              =   1950
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00E0E0E0&
      X1              =   1680
      X2              =   120
      Y1              =   1935
      Y2              =   1935
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00E0E0E0&
      X1              =   1695
      X2              =   1695
      Y1              =   1920
      Y2              =   350
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00808080&
      X1              =   1680
      X2              =   1680
      Y1              =   1920
      Y2              =   350
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   120
      X2              =   1680
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00E0E0E0&
      X1              =   120
      X2              =   120
      Y1              =   360
      Y2              =   1950
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Option:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   1860
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   1680
      Y1              =   375
      Y2              =   375
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00808080&
      X1              =   1680
      X2              =   120
      Y1              =   1920
      Y2              =   1920
   End
End
Attribute VB_Name = "CustomMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetActiveWindow Lib "user32" (ByVal hwnd As Long) As Long


Private Declare Function GetActiveWindow Lib "user32" () As Integer



Dim MyHandleOnThings As Integer




Private Sub Command1_Click()

Call mciSendString("set CDAudio Door Open Wait", 0&, 0&, 0&)

End Sub

Private Sub Command2_Click()
    Call mciSendString("set CDAudio Door Closed Wait", 0&, 0&, 0&)
End Sub

Private Sub Command3_Click()
Openurl InputBox("What URL would you like to open?", "Specify URL")
End Sub

Private Sub Command4_Click()
Shell "Notepad.exe", vbNormalFocus

End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Command6_Click()
Unload Me
'Me.Height = 0
End Sub

Private Sub Form_Load()

Me.Show
MyHandleOnThings = GetActiveWindow
SetActiveWindow (MyHandleOnThings)
Timer1.Enabled = True

End Sub





Private Sub Form_Unload(Cancel As Integer)

    Call mciSendString("set CDAudio Door Closed Wait", 0&, 0&, 0&)
       
End Sub


Private Sub Timer1_Timer()
If Me.Height >= 3030 Then
Timer1.Enabled = False
Timer2.Enabled = True
Exit Sub
End If
Me.Top = Screen.Height - CustomMenu.Height - 650                'Y '- CustomMenu.Height

Me.Height = Me.Height + 200

End Sub

Private Sub Timer2_Timer()
    If GetActiveWindow() <> MyHandleOnThings Then

        Unload Me
    End If
End Sub


Private Function SendMCIString(cmd As String, fShowError As Boolean) As Boolean
    Static rc As Long
    Static errStr As String * 200
    rc = mciSendString(cmd, 0, 0, hwnd)


    If (fShowError And rc <> 0) Then
        mciGetErrorString rc, errStr, Len(errStr)
        MsgBox errStr
    End If
    SendMCIString = (rc = 0)
End Function

