VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ADD medicine"
   ClientHeight    =   8445
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15870
   LinkTopic       =   "Form1"
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "CHECK"
      Height          =   495
      Left            =   5640
      TabIndex        =   13
      Top             =   1800
      Width           =   2895
   End
   Begin VB.Frame Frame1 
      Caption         =   "ADD medicine"
      Height          =   6855
      Left            =   2520
      TabIndex        =   2
      Top             =   3120
      Width           =   10215
      Begin VB.TextBox Text6 
         Height          =   1095
         Left            =   3480
         TabIndex        =   12
         Top             =   3720
         Width           =   4695
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   3480
         TabIndex        =   11
         Top             =   3120
         Width           =   4695
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   3480
         TabIndex        =   10
         Top             =   2520
         Width           =   4695
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   3480
         TabIndex        =   9
         Top             =   1920
         Width           =   4695
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   3480
         TabIndex        =   8
         Top             =   1320
         Width           =   4695
      End
      Begin VB.Label Label10 
         Caption         =   "CANCEL"
         Height          =   615
         Left            =   7680
         TabIndex        =   17
         Top             =   5520
         Width           =   2415
      End
      Begin VB.Label Label9 
         Caption         =   "ADD"
         Height          =   615
         Left            =   5640
         TabIndex        =   16
         Top             =   5520
         Width           =   2415
      End
      Begin VB.Label Label7 
         Caption         =   "BACK"
         Height          =   615
         Left            =   3600
         TabIndex        =   14
         Top             =   5520
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "Medcine Description"
         Height          =   615
         Left            =   480
         TabIndex        =   7
         Top             =   3840
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "Medicine color"
         Height          =   615
         Left            =   480
         TabIndex        =   6
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "Medicine Type"
         Height          =   615
         Left            =   480
         TabIndex        =   5
         Top             =   2640
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Medicine Name "
         Height          =   615
         Left            =   480
         TabIndex        =   4
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Medicine ID"
         Height          =   615
         Left            =   480
         TabIndex        =   3
         Top             =   1440
         Width           =   2895
      End
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Left            =   7560
      TabIndex        =   0
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label Label8 
      Caption         =   "BACK"
      Height          =   615
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Please enter medicine name"
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   840
      Width           =   5655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Frame2_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Label3_Click()
End Sub

Private Sub Label4_Click()

End Sub
