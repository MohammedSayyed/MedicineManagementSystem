VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   8820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15735
   LinkTopic       =   "Form4"
   ScaleHeight     =   8820
   ScaleWidth      =   15735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "CHECK"
      Height          =   495
      Left            =   7080
      TabIndex        =   16
      Top             =   1680
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   7800
      TabIndex        =   15
      Top             =   360
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Caption         =   "UPDATE batch"
      Height          =   6855
      Left            =   3000
      TabIndex        =   0
      Top             =   2400
      Width           =   10215
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   3480
         TabIndex        =   5
         Top             =   1320
         Width           =   4695
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   3480
         TabIndex        =   4
         Top             =   1920
         Width           =   4695
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   3480
         TabIndex        =   3
         Top             =   2520
         Width           =   4695
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   3480
         TabIndex        =   2
         Top             =   3120
         Width           =   4695
      End
      Begin VB.TextBox Text6 
         Height          =   1095
         Left            =   3480
         TabIndex        =   1
         Top             =   3720
         Width           =   4695
      End
      Begin VB.Label Label2 
         Caption         =   "Medicine ID"
         Height          =   615
         Left            =   480
         TabIndex        =   13
         Top             =   1440
         Width           =   2895
      End
      Begin VB.Label Label3 
         Caption         =   "Batch ID"
         Height          =   615
         Left            =   480
         TabIndex        =   12
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "Expiry Date"
         Height          =   615
         Left            =   480
         TabIndex        =   11
         Top             =   2640
         Width           =   2655
      End
      Begin VB.Label Label5 
         Caption         =   "Manufacturing Date"
         Height          =   615
         Left            =   480
         TabIndex        =   10
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label Label6 
         Caption         =   "Medicine Quantity"
         Height          =   615
         Left            =   480
         TabIndex        =   9
         Top             =   3840
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "BACK"
         Height          =   615
         Left            =   3600
         TabIndex        =   8
         Top             =   5520
         Width           =   855
      End
      Begin VB.Label Label7 
         Caption         =   "UPDATE"
         Height          =   615
         Left            =   5640
         TabIndex        =   7
         Top             =   5520
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "CANCEL"
         Height          =   615
         Left            =   7200
         TabIndex        =   6
         Top             =   5520
         Width           =   1095
      End
   End
   Begin VB.Label Label9 
      Caption         =   "Please Enter Batch ID"
      Height          =   735
      Left            =   3120
      TabIndex        =   14
      Top             =   480
      Width           =   4215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub
