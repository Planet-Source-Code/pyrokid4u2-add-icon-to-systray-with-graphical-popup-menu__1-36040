VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Systray With Custom Popup Menu"
   ClientHeight    =   1605
   ClientLeft      =   2640
   ClientTop       =   3555
   ClientWidth     =   4470
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   4470
   Begin VB.Timer RemClip 
      Interval        =   500
      Left            =   3600
      Top             =   1200
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   3480
      Picture         =   "frmMain.frx":0442
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   480
      Picture         =   "frmMain.frx":0884
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   3
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send To Systray!"
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Click to goto planetsourcecode.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   4455
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MakeTopMost Me.hwnd
    AddToTray Me, "Click For Custom Menu", Me.Icon  'Picture1.Picture
    SetClipVars 4095, 8295
        'SetMenuIcon Me.hwnd, 0, 2, 0, Picture1.Picture, Picture1.Picture
RemClip.Enabled = True
    SetMenuIcon Me.hwnd, 0, 2, 0, Picture1.Picture, Picture1.Picture
MsgBox "Now just left click on the american flag that appeared in the systray.", vbExclamation, "Icon Added"
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Mousemove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim Message As Long
   On Error Resume Next
    Message = X / Screen.TwipsPerPixelX

    Select Case Message
        Case 513
            temp = GetY
            If temp > (Screen.Height / Screen.TwipsPerPixelY) - 30 Then
                CustomMenu.Left = X + CustomMenu.Width ' - (CustomMenu.Width / 2)
                CustomMenu.Top = Screen.Height - CustomMenu.Height - 360 'Y '- CustomMenu.Height
                CustomMenu.Show
            End If
    End Select
End Sub




Private Sub Form_Unload(Cancel As Integer)
    RemoveClipping
    RemoveFromTray
End Sub

Private Sub Label1_Click()
Openurl "http://www.planetsourcecode.com"
End Sub

Private Sub RemClip_Timer()
    RemoveClipping
    RemClip.Enabled = False
End Sub



