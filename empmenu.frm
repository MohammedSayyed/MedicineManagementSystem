VERSION 5.00
Begin VB.Form empmenu 
   BorderStyle     =   0  'None
   Caption         =   "Form4"
   ClientHeight    =   12690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   22695
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   Picture         =   "empmenu.frx":0000
   ScaleHeight     =   12690
   ScaleWidth      =   22695
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image3 
      Height          =   1305
      Left            =   12075
      Picture         =   "empmenu.frx":39BB7
      Stretch         =   -1  'True
      Top             =   6405
      Width           =   1305
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   45
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   855
      Left            =   10155
      TabIndex        =   2
      Top             =   7965
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   45
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   855
      Left            =   4875
      TabIndex        =   1
      Top             =   7965
      Width           =   1815
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   10155
      Top             =   8805
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   4275
      Top             =   8805
      Width           =   3255
   End
   Begin VB.Image Image2 
      Height          =   1305
      Left            =   6195
      Picture         =   "empmenu.frx":3BB63
      Stretch         =   -1  'True
      Top             =   6405
      Width           =   1305
   End
   Begin VB.Image Image5 
      Height          =   2985
      Left            =   15315
      Picture         =   "empmenu.frx":41080
      Stretch         =   -1  'True
      Top             =   4725
      Width           =   3105
   End
   Begin VB.Image Image1 
      Height          =   2985
      Left            =   4395
      Picture         =   "empmenu.frx":52A1A
      Stretch         =   -1  'True
      Top             =   4725
      Width           =   3105
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "EMPLOYEE MENU"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   45
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   1335
      Left            =   480
      TabIndex        =   0
      Top             =   720
      Width           =   8775
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   240
      Top             =   720
      Width           =   135
   End
   Begin VB.Image Image4 
      Height          =   2985
      Left            =   10275
      Picture         =   "empmenu.frx":55603
      Stretch         =   -1  'True
      Top             =   4725
      Width           =   3105
   End
End
Attribute VB_Name = "empmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
employee1.Left = 0
employee1.Top = 0
employee1.Show
Unload Me

End Sub

Private Sub Image2_Click()
employee1.Left = 0
employee1.Top = 0
employee1.Show
Unload Me
End Sub

Private Sub Image3_Click()
employee2.Left = 0
employee2.Top = 0
employee2.Show
Unload Me
End Sub

Private Sub Image4_Click()
employee2.Left = 0
employee2.Top = 0
employee2.Show
Unload Me
End Sub

Private Sub Image5_Click()
mainform.Left = 0
mainform.Top = 0
mainform.Show
Unload Me
End Sub
