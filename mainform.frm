VERSION 5.00
Begin VB.Form mainform 
   BorderStyle     =   0  'None
   Caption         =   "mainform"
   ClientHeight    =   12195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   23040
   ControlBox      =   0   'False
   FillColor       =   &H00C000C0&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "mainform.frx":0000
   ScaleHeight     =   12195
   ScaleMode       =   0  'User
   ScaleWidth      =   23040
   ShowInTaskbar   =   0   'False
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
      Left            =   17880
      TabIndex        =   7
      Top             =   10080
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Batch"
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
      Left            =   3000
      TabIndex        =   6
      Top             =   10080
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee"
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
      Left            =   8400
      TabIndex        =   5
      Top             =   5640
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Supplier"
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
      Left            =   13680
      TabIndex        =   4
      Top             =   5640
      Width           =   2175
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
      Left            =   8400
      TabIndex        =   3
      Top             =   10080
      Width           =   1935
   End
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
      Left            =   13560
      TabIndex        =   2
      Top             =   10080
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Medcine"
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
      Left            =   18000
      TabIndex        =   1
      Top             =   5640
      Width           =   2295
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Doctor"
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
      Left            =   3120
      TabIndex        =   0
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Image Image8 
      Height          =   2505
      Left            =   17880
      Picture         =   "mainform.frx":1711F
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   2505
   End
   Begin VB.Image Image7 
      Height          =   2505
      Left            =   13560
      Picture         =   "mainform.frx":1D7B7
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   2505
   End
   Begin VB.Image Image4 
      Height          =   2505
      Left            =   18000
      Picture         =   "mainform.frx":20D76
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   2505
   End
   Begin VB.Image Image1 
      Height          =   2505
      Left            =   3000
      Picture         =   "mainform.frx":238F5
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   2505
   End
   Begin VB.Image Image3 
      Height          =   2505
      Left            =   13680
      Picture         =   "mainform.frx":2C393
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   2505
   End
   Begin VB.Image Image2 
      Height          =   2505
      Left            =   8400
      Picture         =   "mainform.frx":31338
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   2505
   End
   Begin VB.Image Image6 
      Height          =   2505
      Left            =   8400
      Picture         =   "mainform.frx":33F21
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   2505
   End
   Begin VB.Image Image5 
      Height          =   2505
      Left            =   3000
      Picture         =   "mainform.frx":350BD
      Stretch         =   -1  'True
      Top             =   7320
      Width           =   2505
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   2505
      Left            =   2520
      Top             =   2520
      Width           =   2625
   End
   Begin VB.Shape Shape7 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00C0C000&
      FillStyle       =   0  'Solid
      Height          =   2505
      Left            =   2520
      Top             =   6960
      Width           =   2625
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   2505
      Left            =   7920
      Top             =   2520
      Width           =   2505
   End
   Begin VB.Shape Shape5 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF80&
      FillStyle       =   0  'Solid
      Height          =   2505
      Left            =   13200
      Top             =   2520
      Width           =   2505
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF00&
      FillStyle       =   0  'Solid
      Height          =   2505
      Left            =   7920
      Top             =   6960
      Width           =   2625
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFC0&
      FillStyle       =   0  'Solid
      Height          =   2505
      Left            =   17520
      Top             =   2520
      Width           =   2505
   End
   Begin VB.Shape Shape10 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFC0&
      FillStyle       =   0  'Solid
      Height          =   2505
      Left            =   17400
      Top             =   6960
      Width           =   2505
   End
   Begin VB.Shape Shape8 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFF80&
      FillStyle       =   0  'Solid
      Height          =   2505
      Left            =   13080
      Top             =   6960
      Width           =   2505
   End
End
Attribute VB_Name = "mainform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image1.Height <> 2505 Or Image2.Height <> 2505 Or Image3.Height <> 2505 Or Image4.Height <> 2505 Or Image5.Height <> 2505 Or Image6.Height <> 2505 Or Image7.Height <> 2505 Or Image8.Height <> 2505 Then
Image1.Height = 2505
Image1.Width = 2505
Image2.Height = 2505
Image2.Width = 2505
Image3.Height = 2505
Image3.Width = 2505
Image4.Height = 2505
Image4.Width = 2505
Image5.Height = 2505
Image5.Width = 2505
Image6.Height = 2505
Image6.Width = 2505
Image7.Height = 2505
Image7.Width = 2505
Image8.Height = 2505
Image8.Width = 2505
End If
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image1.Height <> 2800 Then
Image1.Height = 2800
Image1.Width = 2800
Image2.Height = 2405
Image2.Width = 2405
Image3.Height = 2405
Image3.Width = 2405
Image4.Height = 2405
Image4.Width = 2405
Image5.Height = 2405
Image5.Width = 2405
Image6.Height = 2405
Image6.Width = 2405
Image7.Height = 2405
Image7.Width = 2405
Image8.Height = 2405
Image8.Width = 2405
End If
End Sub


Private Sub Image1_Click()
doctor1.Top = 0
doctor1.Left = 0
Unload Me
doctor1.Show
End Sub


Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image2.Height <> 2800 Then
Image2.Height = 2800
Image2.Width = 2800
Image1.Height = 2405
Image1.Width = 2405
Image3.Height = 2405
Image3.Width = 2405
Image4.Height = 2405
Image4.Width = 2405
Image5.Height = 2405
Image5.Width = 2405
Image6.Height = 2405
Image6.Width = 2405
Image7.Height = 2405
Image7.Width = 2405
Image8.Height = 2405
Image8.Width = 2405
End If
End Sub

Private Sub Image2_Click()

employee1.Top = 0
employee1.Left = 0
Unload Me
employee1.Show
End Sub

Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image3.Height <> 2800 Then
Image3.Height = 2800
Image3.Width = 2800
Image1.Height = 2405
Image1.Width = 2405
Image2.Height = 2405
Image2.Width = 2405
Image4.Height = 2405
Image4.Width = 2405
Image5.Height = 2405
Image5.Width = 2405
Image6.Height = 2405
Image6.Width = 2405
Image7.Height = 2405
Image7.Width = 2405
Image8.Height = 2405
Image8.Width = 2405
End If
End Sub
Private Sub Image3_Click()
supplier1.Left = 0
supplier1.Top = 0
supplier1.Show
Unload Me
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image4.Height <> 2800 Then
Image4.Height = 2800
Image4.Width = 2800
Image1.Height = 2405
Image1.Width = 2405
Image2.Height = 2405
Image2.Width = 2405
Image3.Height = 2405
Image3.Width = 2405
Image5.Height = 2405
Image5.Width = 2405
Image6.Height = 2405
Image6.Width = 2405
Image7.Height = 2405
Image7.Width = 2405
Image8.Height = 2405
Image8.Width = 2405
End If
End Sub
Private Sub Image4_Click()
medicine1.Top = 0
medicine1.Left = 0
medicine1.Show
Unload Me
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image5.Height <> 2800 Then
Image5.Height = 2800
Image5.Width = 2800
Image1.Height = 2405
Image1.Width = 2405
Image2.Height = 2405
Image2.Width = 2405
Image3.Height = 2405
Image3.Width = 2405
Image4.Height = 2405
Image4.Width = 2405
Image6.Height = 2405
Image6.Width = 2405
Image7.Height = 2405
Image7.Width = 2405
Image8.Height = 2405
Image8.Width = 2405
End If
End Sub
Private Sub Image5_Click()
batch1.Top = 0
batch1.Left = 0
batch1.Show
Unload Me
End Sub

Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image6.Height <> 2800 Then
Image6.Height = 2800
Image6.Width = 2800
Image1.Height = 2405
Image1.Width = 2405
Image2.Height = 2405
Image2.Width = 2405
Image3.Height = 2405
Image3.Width = 2405
Image4.Height = 2405
Image4.Width = 2405
Image5.Height = 2405
Image5.Width = 2405
Image7.Height = 2405
Image7.Width = 2405
Image8.Height = 2405
Image8.Width = 2405
End If
End Sub
Private Sub image6_click()
bill1.Top = 0
bill1.Left = 0
bill1.Show
Unload Me
End Sub
Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image7.Height <> 2800 Then
Image7.Height = 2800
Image7.Width = 2800
Image1.Height = 2405
Image1.Width = 2405
Image2.Height = 2405
Image2.Width = 2405
Image3.Height = 2405
Image3.Width = 2405
Image4.Height = 2405
Image4.Width = 2405
Image5.Height = 2405
Image5.Width = 2405
Image6.Height = 2405
Image6.Width = 2405
Image8.Height = 2405
Image8.Width = 2405
End If
End Sub
Private Sub image7_click()
report.Top = 0
report.Left = 0
report.Show
Unload Me
End Sub
Private Sub Image8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image8.Height <> 2800 Then
Image8.Height = 2800
Image8.Width = 2800
Image1.Height = 2405
Image1.Width = 2405
Image2.Height = 2405
Image2.Width = 2405
Image3.Height = 2405
Image3.Width = 2405
Image4.Height = 2405
Image4.Width = 2405
Image5.Height = 2405
Image5.Width = 2405
Image6.Height = 2405
Image6.Width = 2405
Image7.Height = 2405
Image7.Width = 2405
End If
End Sub
Private Sub image8_click()
Unload Me
LoginForm.Top = 0
LoginForm.Left = 0
LoginForm.Show
End Sub



