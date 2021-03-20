VERSION 5.00
Begin VB.Form LoginForm 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "LoginForm"
   ClientHeight    =   14070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   26235
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "LoginForm.frx":0000
   ScaleHeight     =   14070
   ScaleWidth      =   26235
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H00000000&
      Height          =   6735
      Left            =   5400
      TabIndex        =   15
      Top             =   3840
      Visible         =   0   'False
      Width           =   11055
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   23.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   4200
         TabIndex        =   0
         Top             =   2040
         Width           =   6015
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   23.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         IMEMode         =   3  'DISABLE
         Left            =   4200
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   3720
         Width           =   6015
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   7680
         TabIndex        =   26
         Top             =   5520
         Width           =   3135
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   240
         TabIndex        =   25
         Top             =   5520
         Width           =   3135
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   735
         Left            =   3840
         TabIndex        =   16
         Top             =   5520
         Width           =   3495
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "P A S S W O R D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   720
         TabIndex        =   18
         Top             =   3720
         Width           =   3255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "L O G I N - I D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   840
         TabIndex        =   17
         Top             =   2160
         Width           =   2535
      End
      Begin VB.Shape Shape5 
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   1800
         Width           =   10215
      End
      Begin VB.Shape Shape6 
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   3480
         Width           =   10215
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Admin Login"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   480
         TabIndex        =   19
         Top             =   360
         Width           =   6735
      End
      Begin VB.Shape login 
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   5400
         Width           =   3375
      End
      Begin VB.Shape cancel 
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   3840
         Shape           =   4  'Rounded Rectangle
         Top             =   5400
         Width           =   3495
      End
      Begin VB.Shape forget 
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   7560
         Shape           =   4  'Rounded Rectangle
         Top             =   5400
         Width           =   3375
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H00000000&
      Height          =   6735
      Left            =   5400
      TabIndex        =   2
      Top             =   3840
      Visible         =   0   'False
      Width           =   11055
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   3240
         TabIndex        =   28
         Top             =   1920
         Width           =   6975
      End
      Begin VB.TextBox Text6 
         Height          =   495
         Left            =   3240
         TabIndex        =   3
         Top             =   3600
         Width           =   6975
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   6720
         TabIndex        =   27
         Top             =   5400
         Width           =   2175
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Show password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   1800
         TabIndex        =   4
         Top             =   5400
         Width           =   3615
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Answer"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   720
         TabIndex        =   5
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Question"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   720
         TabIndex        =   6
         Top             =   2040
         Width           =   2535
      End
      Begin VB.Shape Shape11 
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   1680
         Width           =   10215
      End
      Begin VB.Shape Shape12 
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   360
         Shape           =   4  'Rounded Rectangle
         Top             =   3360
         Width           =   10215
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "forget password"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   480
         TabIndex        =   7
         Top             =   240
         Width           =   6735
      End
      Begin VB.Shape Shape13 
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   1800
         Shape           =   4  'Rounded Rectangle
         Top             =   5280
         Width           =   3615
      End
      Begin VB.Shape Shape15 
         FillStyle       =   0  'Solid
         Height          =   735
         Left            =   6600
         Shape           =   4  'Rounded Rectangle
         Top             =   5280
         Width           =   2415
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H00000000&
      Height          =   6735
      Left            =   5160
      TabIndex        =   8
      Top             =   3840
      Visible         =   0   'False
      Width           =   11055
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000FFFF&
         Caption         =   "Show Password"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   24
         Top             =   4680
         Width           =   2175
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         IMEMode         =   3  'DISABLE
         Left            =   4200
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   3720
         Width           =   6015
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   4200
         TabIndex        =   9
         Top             =   2160
         Width           =   6015
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   7560
         TabIndex        =   23
         Top             =   5520
         Width           =   3135
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   3960
         TabIndex        =   22
         Top             =   5520
         Width           =   3135
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Forget Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   375
         Left            =   7680
         TabIndex        =   21
         Top             =   4680
         Width           =   2055
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   22.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   615
         Left            =   360
         TabIndex        =   11
         Top             =   5520
         Width           =   3135
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "P A S S W O R D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   720
         TabIndex        =   13
         Top             =   3840
         Width           =   3255
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "L O G I N - I D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   17.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   720
         TabIndex        =   12
         Top             =   2280
         Width           =   2535
      End
      Begin VB.Shape Shape3 
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   7440
         Shape           =   4  'Rounded Rectangle
         Top             =   5400
         Width           =   3375
      End
      Begin VB.Shape Shape7 
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   3840
         Shape           =   4  'Rounded Rectangle
         Top             =   5400
         Width           =   3495
      End
      Begin VB.Shape Shape8 
         FillStyle       =   0  'Solid
         Height          =   855
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   5400
         Width           =   3375
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Login"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   48
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   6735
      End
      Begin VB.Shape Shape9 
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   3480
         Width           =   10215
      End
      Begin VB.Shape Shape10 
         FillStyle       =   0  'Solid
         Height          =   1095
         Left            =   240
         Shape           =   4  'Rounded Rectangle
         Top             =   1920
         Width           =   10215
      End
   End
   Begin VB.Image Image2 
      Height          =   3375
      Left            =   12720
      Picture         =   "LoginForm.frx":BB65
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   3270
   End
   Begin VB.Image Image1 
      Height          =   3375
      Left            =   7080
      Picture         =   "LoginForm.frx":12EFC
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   3270
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   2535
      Left            =   14280
      Top             =   240
      Width           =   135
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "HELLO HAVE A GREAT DAY.."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   3375
      Left            =   14640
      TabIndex        =   20
      Top             =   0
      Width           =   8895
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   3855
      Left            =   6480
      Top             =   5640
      Width           =   3855
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0000FFFF&
      FillStyle       =   0  'Solid
      Height          =   3855
      Left            =   12120
      Top             =   5640
      Width           =   3855
   End
End
Attribute VB_Name = "LoginForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Text3.PasswordChar = ""
Else
    Text3.PasswordChar = "*"
End If
End Sub

Private Sub Form_Load()
Frame1.Visible = False
Frame2.Visible = False
Frame3.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image1.Height <> 3720 Or Image2.Height <> 3720 Then
Image1.Height = 3720
Image1.Width = 3720
Image2.Height = 3720
Image2.Width = 3720
End If
End Sub


Private Sub Image1_Click()
Image1.Visible = False
Image2.Visible = False
Shape2.Visible = False
Shape4.Visible = False
Frame1.Visible = True
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image1.Height <> 4000 Then
Image1.Height = 4000
Image1.Width = 4000
Image2.Height = 3500
Image2.Width = 3500
End If
End Sub

Private Sub Image2_Click()
Image1.Visible = False
Image2.Visible = False
Shape2.Visible = False
Shape4.Visible = False
Frame2.Visible = True
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Image2.Height <> 4000 Then
Image2.Height = 4000
Image2.Width = 4000
Image1.Height = 3500
Image1.Width = 3500
End If
End Sub

Private Sub Label12_Click()
Frame2.Visible = True
Frame3.Visible = False
Frame1.Visible = False

End Sub

Private Sub Label13_Click()
End
End Sub

Private Sub Label14_Click()
Frame3.Visible = False
Frame1.Visible = False
Frame2.Visible = False
Image1.Visible = True
Image2.Visible = True
Shape2.Visible = True
Shape4.Visible = True
End Sub

Private Sub Label17_Click()
If Text1.Text = "admin" And Text2.Text = "admin" Then
Text1.Text = ""
Text2.Text = ""
mainform.Top = 0
mainform.Left = 0
Unload Me
mainform.Show
Else
MsgBox "USERNAME or PASSWORD is incorect", vbCritical
Text1.Text = ""
Text2.Text = ""
End If
flag = False
End Sub

Private Sub Label18_Click()
If Text6.Text = remployee!E_answer Then
    MsgBox "username= " & remployee!E_username & " password = " & remployee!E_password, vbInformation, "INFO"
Else
    MsgBox "INCORRECT ANSWER"
    Text6.Text = ""
    Text6.SetFocus
End If
End Sub


Private Sub Label2_Click()
If remployee.State = adStateOpen Then
    remployee.Close
End If
remployee.Open "select * from [Employee] where(E_username='" & Text4.Text & "')", conn, adOpenDynamic, adLockPessimistic
If remployee.RecordCount <> 1 Then
    MsgBox "PLEASE ENTER CORRECT USERNAME!!", vbCritical, "INFO"
    Text4.Text = ""
    Text4.SetFocus
Else
Text5.Text = remployee!E_question1
Frame3.Visible = True
Frame2.Visible = False
Frame1.Visible = False

End If
End Sub

Private Sub Label3_Click()
Frame2.Visible = False
Image1.Visible = True
Image2.Visible = True
Shape2.Visible = True
Shape4.Visible = True
End Sub

Private Sub Label7_Click()
End
End Sub

Private Sub Picture1_Click()
Frame1.Visible = True
Picture1.Visible = False
Picture2.Visible = False
Shape4.Visible = False
Shape2.Visible = False

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape4.Visible = True
End Sub

Private Sub Picture2_Click()
Frame2.Visible = True
Picture1.Visible = False
Picture2.Visible = False
Shape4.Visible = False
Shape2.Visible = False

End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape2.Visible = True
End Sub


Private Sub Label8_Click()
If remployee.State = adStateOpen Then
remployee.Close
End If
remployee.Open "select * from [Employee] where(E_username='" & Text4.Text & "')", conn, adOpenDynamic, adLockPessimistic
If remployee.RecordCount <> 1 Then
    MsgBox "PLEASE ENTER CORRECT USERNAME!!", vbCritical, "INFO"
    Text4.Text = ""
Else
    If Text3.Text = remployee!E_password Then
        eid = remployee!E_id
        mainform2.Top = 0
        mainform2.Left = 0
        mainform2.Show
        Unload Me
    Else
        MsgBox "INCORRECT PASSWORD, PLEASE TRY AGAIN", vbCritical, "ERROR"
        Text3.Text = ""
        Text3.SetFocus
    End If
End If
flag = True
End Sub

