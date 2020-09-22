VERSION 5.00
Begin VB.Form start 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "                      Mad cow test"
   ClientHeight    =   4110
   ClientLeft      =   4305
   ClientTop       =   2775
   ClientWidth     =   4215
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "madcowtest2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "About"
      Height          =   735
      Left            =   3360
      TabIndex        =   9
      Top             =   3240
      Width           =   735
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "norm2.wav"
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Play mad cow sound"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   2640
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Play normal cow sound"
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   1080
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit - The mad cow test"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Exit"
      Top             =   3240
      Width           =   3135
   End
   Begin VB.PictureBox Picture2 
      Height          =   1455
      Left            =   120
      Picture         =   "madcowtest2.frx":030A
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   1
      Top             =   1680
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   120
      Picture         =   "madcowtest2.frx":21AC
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "If your cow sounds like this Then fire up the BBQ"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1680
      TabIndex        =   8
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "If your cow sounds like this then you better have fish instaed"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1680
      TabIndex        =   7
      Top             =   1680
      Width           =   2415
   End
End
Attribute VB_Name = "start"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call MsgBox("Mad cow test - By: Andy" & Chr(13) & Chr(13) & "This is a simple little app that has sounds, pictures, and a cool thing at start")
End Sub

Private Sub Command2_Click()
Call MsgBox("Mooooooooooooooooooo" & Chr(13) & Chr(13) & "Mo Mooo Moooooo Moo Mo, Mooooo, Moo Mo Mooooooo Moo Mo Mooo")
Unload Me
End
End Sub

Private Sub Command3_Click()
Text1.Text = "norm2.wav"
sndStartM
End Sub

Public Sub sndPlayW(Filename As String)
    Call sndPlaySound(Filename, SND_ASYNC Or SND_NODEFAULT Or SND_NOSTOP)
End Sub
Public Sub sndPlayM(Filename As String)
    Call mciSendString("Open " & Filename & " Alias MM", 0, 0, 0)
    Call mciSendString("Play MM", 0, 0, 0)
End Sub

Public Sub sndPauseM()
    Call mciSendString("Stop MM", 0, 0, 0)
End Sub
Public Sub sndStartM()
    If Mid(Text1.Text, (Len(Text1.Text) - 3), 4) = ".mid" Then
        sndPlayM (File1.Path & "\" & Text1.Text)
    ElseIf Mid(Text1.Text, (Len(Text1.Text) - 3), 4) = ".mp3" Then
        sndPlayM (File1.Path & "\" & Text1.Text)
    ElseIf Mid(Text1.Text, (Len(Text1.Text) - 4), 5) = ".mpeg" Then
        sndPlayM (File1.Path & "\" & Text1.Text)
    ElseIf Mid(Text1.Text, (Len(Text1.Text) - 3), 4) = ".wav" Then
        sndPlayW (File1.Path & "\" & Text1.Text)
    End If
End Sub

Private Sub Command4_Click()
Text1.Text = "mad.wav"
sndStartM
End Sub


Private Sub Form_Load()
File1.Path = App.Path
End Sub

