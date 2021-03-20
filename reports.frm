VERSION 5.00
Begin VB.Form report 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   12195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   22605
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   12195
   ScaleWidth      =   22605
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   12735
      Left            =   0
      Picture         =   "reports.frx":0000
      ScaleHeight     =   12675
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BACK"
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
      Left            =   10440
      TabIndex        =   8
      Top             =   8880
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "                     SUPPLIER REPORT"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1635
      Left            =   14520
      TabIndex        =   7
      Top             =   6120
      Width           =   1995
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "                         BILL   REPORT"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1635
      Left            =   10320
      TabIndex        =   6
      Top             =   6120
      Width           =   1995
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "                     BATCH REPORT"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1035
      Left            =   6000
      TabIndex        =   5
      Top             =   6240
      Width           =   1995
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "                     MEDICINE REPORT"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1035
      Left            =   14520
      TabIndex        =   4
      Top             =   3120
      Width           =   1995
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "                     DOCTOR REPORT"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1035
      Left            =   10320
      TabIndex        =   3
      Top             =   3120
      Width           =   1995
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "                     EMPLOYEE REPORT"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1035
      Left            =   6000
      TabIndex        =   2
      Top             =   3120
      Width           =   1995
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Reports..."
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
      Left            =   4800
      TabIndex        =   1
      Top             =   120
      Width           =   6135
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000D&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   4680
      Top             =   840
      Width           =   18015
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1935
      Index           =   0
      Left            =   6000
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Width           =   1995
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1935
      Left            =   14520
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Width           =   1995
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1935
      Left            =   10320
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Width           =   1995
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1935
      Index           =   1
      Left            =   6000
      Shape           =   4  'Rounded Rectangle
      Top             =   5880
      Width           =   1995
   End
   Begin VB.Shape Shape5 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1935
      Left            =   10320
      Shape           =   4  'Rounded Rectangle
      Top             =   5880
      Width           =   1995
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1935
      Left            =   14520
      Shape           =   4  'Rounded Rectangle
      Top             =   5880
      Width           =   1995
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   3
      Left            =   10320
      Shape           =   4  'Rounded Rectangle
      Top             =   8760
      Width           =   2055
   End
End
Attribute VB_Name = "report"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Label2_Click()
'If remployee.State = adStateOpen Then
'    remployee.Close
'End If
'remployee.Open "Select * from employee", conn, adOpenDynamic, adLockPessimistic
'Set DataGrid1.DataSource = remployee
'Frame1.Visible = True
DataReport1.Show
End Sub

Private Sub Label3_Click()
'If rdoctor.State = adStateOpen Then
'    rdoctor.Close
'End If
'rdoctor.Open "Select * from doctor", conn, adOpenDynamic, adLockPessimistic
'Set DataGrid1.DataSource = rdoctor
'Frame1.Visible = True
DataReport3.Show
End Sub

Private Sub Label4_Click()
'If rmedicine.State = adStateOpen Then
'    rmedicine.Close
'End If
'rmedicine.Open "Select * from medicine", conn, adOpenDynamic, adLockPessimistic
'Set DataGrid1.DataSource = rmedicine
'Frame1.Visible = True
DataReport4.Show
End Sub

Private Sub Label5_Click()
'If rbatch.State = adStateOpen Then
'    rbatch.Close
'End If
'rbatch.Open "Select * from batch", conn, adOpenDynamic, adLockPessimistic
'Set DataGrid1.DataSource = rbatch
'Frame1.Visible = True
DataReport2.Show
End Sub

Private Sub Label6_Click()
billReport.Top = 0
billReport.Left = 0
billReport.Show
Unload Me
End Sub

Private Sub Label7_Click()
'If rsupplier.State = adStateOpen Then
'    rsupplier.Close
'End If
'rsupplier.Open "Select * from supplier", conn, adOpenDynamic, adLockPessimistic
'Set DataGrid1.DataSource = rsupplier
'Frame1.Visible = True
DataReport5.Show
End Sub

Private Sub Label8_Click()
Frame1.Visible = False
End Sub

Private Sub Label9_Click()
If flag = False Then
    mainform.Left = 0
    mainform.Top = 0
    mainform.Show
Else
    mainform2.Top = 0
    mainform2.Left = 0
    mainform2.Show
End If
Unload Me
End Sub
