VERSION 5.00
Begin VB.Form mainform2 
   BorderStyle     =   0  'None
   Caption         =   "mainform2"
   ClientHeight    =   14070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   26235
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "mainform2.frx":0000
   ScaleHeight     =   14070
   ScaleMode       =   0  'User
   ScaleWidth      =   26235
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Reports"
      BeginProperty Font 
         Name            =   "Mongolian Baiti"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   10560
      TabIndex        =   2
      Top             =   8190
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill"
      BeginProperty Font 
         Name            =   "Mongolian Baiti"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   6360
      TabIndex        =   1
      Top             =   8190
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Logout"
      BeginProperty Font 
         Name            =   "Mongolian Baiti"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   14880
      TabIndex        =   0
      Top             =   8190
      Width           =   1935
   End
   Begin VB.Image Image3 
      Height          =   2505
      Left            =   14685
      Picture         =   "mainform2.frx":1711F
      Stretch         =   -1  'True
      Top             =   5385
      Width           =   2505
   End
   Begin VB.Image Image2 
      Height          =   2505
      Left            =   10365
      Picture         =   "mainform2.frx":1D7B7
      Stretch         =   -1  'True
      Top             =   5385
      Width           =   2505
   End
   Begin VB.Image Image1 
      Height          =   2505
      Left            =   6165
      Picture         =   "mainform2.frx":3794E
      Stretch         =   -1  'True
      Top             =   5385
      Width           =   2505
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   2505
      Left            =   5805
      Top             =   5025
      Width           =   2625
   End
   Begin VB.Shape Shape8 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF80&
      FillStyle       =   0  'Solid
      Height          =   2505
      Left            =   10005
      Top             =   5025
      Width           =   2505
   End
   Begin VB.Shape Shape10 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFC0&
      FillStyle       =   0  'Solid
      Height          =   2505
      Left            =   14325
      Top             =   5025
      Width           =   2505
   End
End
Attribute VB_Name = "mainform2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image1.Height <> 2505 Or Image2.Height <> 2505 Or Image3.Height <> 2505 Then
Image1.Height = 2505
Image1.Width = 2505
Image2.Height = 2505
Image2.Width = 2505
Image3.Height = 2505
Image3.Width = 2505
End If
End Sub


Private Sub Image1_Click()
bill1.Top = 0
bill1.Left = 0
bill1.Show
Unload Me
End Sub
     
Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image1.Height <> 3000 Then
Image1.Height = 3000
Image1.Width = 3000
Image2.Height = 2505
Image2.Width = 2505
Image3.Height = 2505
Image3.Width = 2505
End If
End Sub

Private Sub Image2_Click()
report.Top = 0
report.Left = 0
report.Show
Unload Me
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image2.Height <> 3000 Then
Image2.Height = 3000
Image2.Width = 3000
Image1.Height = 2505
Image1.Width = 2505
Image3.Height = 2505
Image3.Width = 2505
End If
End Sub

Private Sub Image3_Click()
LoginForm.Top = 0
LoginForm.Left = 0
Unload Me
LoginForm.Show
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image3.Height <> 3000 Then
Image3.Height = 3000
Image3.Width = 3000
Image2.Height = 2505
Image2.Width = 2505
Image1.Height = 2505
Image1.Width = 2505
End If
End Sub

