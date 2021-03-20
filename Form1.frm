VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000005&
   Caption         =   "Form1"
   ClientHeight    =   10500
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   21060
   LinkTopic       =   "Form1"
   ScaleHeight     =   10500
   ScaleWidth      =   21060
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "ADD medicine"
      Height          =   9135
      Left            =   5160
      TabIndex        =   5
      Top             =   2880
      Width           =   17175
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   840
         TabIndex        =   10
         Top             =   1560
         Width           =   4695
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   840
         TabIndex        =   9
         Top             =   3600
         Width           =   4695
      End
      Begin VB.TextBox Text4 
         Height          =   615
         Left            =   840
         TabIndex        =   8
         Top             =   5520
         Width           =   4695
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   8280
         TabIndex        =   7
         Top             =   1560
         Width           =   4695
      End
      Begin VB.TextBox Text6 
         Height          =   1215
         Left            =   8280
         TabIndex        =   6
         Top             =   3600
         Width           =   4695
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "BACK"
         BeginProperty Font 
            Name            =   "MS UI Gothic"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6720
         TabIndex        =   13
         Top             =   7320
         Width           =   2415
      End
      Begin VB.Shape Shape4 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   6600
         Shape           =   4  'Rounded Rectangle
         Top             =   7200
         Width           =   2535
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Medicine ID"
         BeginProperty Font 
            Name            =   "Yu Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         TabIndex        =   18
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Medicine Name "
         BeginProperty Font 
            Name            =   "Yu Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         TabIndex        =   17
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Medicine Type"
         BeginProperty Font 
            Name            =   "Yu Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         TabIndex        =   16
         Top             =   4920
         Width           =   2655
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Medicine color"
         BeginProperty Font 
            Name            =   "Yu Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8280
         TabIndex        =   15
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Medcine Description"
         BeginProperty Font 
            Name            =   "Yu Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8160
         TabIndex        =   14
         Top             =   3000
         Width           =   3735
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ADD"
         BeginProperty Font 
            Name            =   "MS UI Gothic"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1680
         TabIndex        =   12
         Top             =   7320
         Width           =   2415
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "MS UI Gothic"
            Size            =   15
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11760
         TabIndex        =   11
         Top             =   7320
         Width           =   2415
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   1680
         Shape           =   4  'Rounded Rectangle
         Top             =   7200
         Width           =   2535
      End
      Begin VB.Shape Shape5 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   11640
         Shape           =   4  'Rounded Rectangle
         Top             =   7200
         Width           =   2535
      End
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   10080
      TabIndex        =   3
      Top             =   1800
      Width           =   5655
   End
   Begin VB.PictureBox Picture1 
      Height          =   12735
      Left            =   0
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   12675
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   -240
      Width           =   4455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CHECK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   16320
      TabIndex        =   4
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   16200
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Enter Medicine Name"
      BeginProperty Font 
         Name            =   "Mongolian Baiti"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   2
      Top             =   1800
      Width           =   4935
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000D&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   4680
      Top             =   1080
      Width           =   18015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "MEDICINE INFO.."
      BeginProperty Font 
         Name            =   "Mongolian Baiti"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   615
      Left            =   4920
      TabIndex        =   1
      Top             =   240
      Width           =   6135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
