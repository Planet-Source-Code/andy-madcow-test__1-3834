VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H80000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3480
   ClientLeft      =   3045
   ClientTop       =   3240
   ClientWidth     =   4590
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdNoBut 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   4335
   End
   Begin VB.Timer tmrFollow 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3480
      Top             =   1800
   End
   Begin VB.CommandButton cmdYesBut 
      Caption         =   "Take the mad cow test"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   4335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "P.s I dare you to click exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   4335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "The mad cow test.Have you had yours lately?"
      BeginProperty Font 
         Name            =   "Child's Play"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    'Sets the position of the window
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
    'Set the parent of ANY object (can be lots of fun! ;-)
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
    'Get the hWnd of the object's parent
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    'Get the current cursor Hot-Spot position

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Const a_Radius = 30 'Acceptable Radius the cursor can be
                'within for the button to 'grab' the cursor
Const HWND_TOPMOST = -1
Dim XnY As POINTAPI, ExitDo As Boolean

Private Sub cmdNoBut_Click()
    cmdYesBut.ZOrder 0  'Set the follower button to infront
    tmrFollow.Enabled = True  'Start the button moving!
End Sub

Private Sub cmdYesBut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'The Click event doesn't work when the button's parent is set to None
    
    ExitDo = True
    'Stop the Do..Loop from running, though you don't need
    'this if you're going to unload the form like this
    
    If GetParent(cmdYesBut.hwnd) <> Me.hwnd Then cmdYesBut.Visible = False
    'If the parent was set to anything other than the form
    'then make it invisible, so it wont get infront of the
    'message box
    start.Show
    Unload frmmain
End Sub

Private Sub tmrFollow_Timer()
    GetCursorPos XnY
    XnY.X = ScaleX(XnY.X, vbPixels, vbTwips) 'Change the dimensions from Pixels
    XnY.Y = ScaleY(XnY.Y, vbPixels, vbTwips) 'to Twips

    'Movement in X
    If cmdYesBut.Left < 0 Then
        cmdYesBut.Left = 0
        Me.Left = Me.Left - 15  'push window
    ElseIf cmdYesBut.Left + cmdYesBut.Width > Me.Width Then
        cmdYesBut.Left = Me.Width - cmdYesBut.Width
        Me.Left = Me.Left + 15  'push window
    Else:
        If cmdYesBut.Left + cmdYesBut.Width / 2 + Me.Left < XnY.X Then cmdYesBut.Left = cmdYesBut.Left + 30 Else cmdYesBut.Left = cmdYesBut.Left - 30
    End If

    'Movement in Y
    If cmdYesBut.Top < 0 Then
        cmdYesBut.Top = 0
        Me.Top = Me.Top - 15
    ElseIf cmdYesBut.Top + cmdYesBut.Height > Me.Height Then
        cmdYesBut.Top = Me.Height - cmdYesBut.Height
        Me.Top = Me.Top + 15
    Else:
        If cmdYesBut.Top + cmdYesBut.Height / 2 + Me.Top < XnY.Y Then cmdYesBut.Top = cmdYesBut.Top + 30 Else cmdYesBut.Top = cmdYesBut.Top - 30
    End If

    If (cmdYesBut.Left + cmdYesBut.Width / 2 + Me.Left < XnY.X + a_Radius) _
        And (cmdYesBut.Left + cmdYesBut.Width / 2 + Me.Left > XnY.X - a_Radius) _
        And (cmdYesBut.Top + cmdYesBut.Height / 2 + Me.Top > XnY.Y - a_Radius) _
        And (cmdYesBut.Top + cmdYesBut.Height / 2 + Me.Top < XnY.Y + a_Radius) Then
        'Within a_Radius twips of the center
        '(pretty long IF statement huh?!)
        tmrFollow.Enabled = False
        Call StickButton(Me, cmdYesBut, cmdYesBut.Width / 2, cmdYesBut.Height / 2)
    End If
End Sub

Private Sub StickButton(ByVal Form As Form, ByVal Button As CommandButton, DpX As Long, DpY As Long)
    SetParent Button.hwnd, 0    'Sets the button's parent to none
    SetWindowPos Button.hwnd, HWND_TOPMOST, 0, 0, 0, 0, 3 'Sets the button to be always on top
    Button.Move Button.Left + Form.Left, Button.Top + Form.Top 'Make sure it's in the same position
    Do
        DoEvents    'So it doesn't 'Hang' the program
        GetCursorPos XnY
        XnY.X = ScaleX(XnY.X, vbPixels, vbTwips)
        XnY.Y = ScaleY(XnY.Y, vbPixels, vbTwips)
        Button.Left = XnY.X - DpX
        Button.Top = XnY.Y - DpY
        If ExitDo Then Exit Do
    Loop  'Stick the Button to the cursor until ExitDo is true
    'And they wont be able to click anything else!! hehe!
    '...why not disable CTRL+ALT+DELETE? hehe!
End Sub
