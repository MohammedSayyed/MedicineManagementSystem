VERSION 5.00
Begin VB.Form batchmenu 
   BorderStyle     =   0  'None
   Caption         =   "Form7"
   ClientHeight    =   12690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   22695
   ControlBox      =   0   'False
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "batchmenu.frx":0000
   ScaleHeight     =   12690
   ScaleWidth      =   22695
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image3 
      Height          =   1305
      Left            =   12075
      Picture         =   "batchmenu.frx":39BB7
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
      Picture         =   "batchmenu.frx":3BB63
      Stretch         =   -1  'True
      Top             =   6405
      Width           =   1305
   End
   Begin VB.Image Image5 
      Height          =   2985
      Left            =   15315
      Picture         =   "batchmenu.frx":41080
      Stretch         =   -1  'True
      Top             =   4725
      Width           =   3105
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00404000&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   360
      Top             =   960
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "BATCH MENU"
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
      Left            =   720
      TabIndex        =   0
      Top             =   960
      Width           =   8775
   End
   Begin VB.Image Image4 
      Height          =   2985
      Left            =   4395
      Picture         =   "batchmenu.frx":52A1A
      Stretch         =   -1  'True
      Top             =   4725
      Width           =   3105
   End
   Begin VB.Image Image1 
      Height          =   2985
      Left            =   10275
      Picture         =   "batchmenu.frx":5B5AA
      Stretch         =   -1  'True
      Top             =   4725
      Width           =   3105
   End
End
Attribute VB_Name = "batchmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()
batch2.Top = 0
batch2.Left = 0
batch2.Show
Unload Me
End Sub

Private Sub Image4_Click()
batch1.Top = 0
batch1.Left = 0
batch1.Show
Unload Me
End Sub

Private Sub Image5_Click()
mainform.Left = 0
mainform.Top = 0
mainform.Show
Unload Me
End Sub
