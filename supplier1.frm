VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form supplier1 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "suppier1"
   ClientHeight    =   12165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   22710
   ForeColor       =   &H8000000C&
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   12165
   ScaleWidth      =   22710
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "ADD medicine"
      Height          =   7095
      Left            =   4800
      TabIndex        =   1
      Top             =   1320
      Visible         =   0   'False
      Width           =   15135
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Yu Gothic Medium"
            Size            =   12
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8280
         TabIndex        =   16
         Top             =   3480
         Width           =   4695
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Yu Gothic Medium"
            Size            =   12
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         TabIndex        =   4
         Top             =   1560
         Width           =   4695
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Yu Gothic Medium"
            Size            =   12
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   840
         TabIndex        =   3
         Top             =   3480
         Width           =   4695
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Yu Gothic Medium"
            Size            =   12
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8280
         TabIndex        =   2
         Top             =   1560
         Width           =   4695
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier E-mail"
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
         TabIndex        =   17
         Top             =   2880
         Width           =   3735
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Name "
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
         TabIndex        =   9
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Address"
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
         TabIndex        =   8
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier PH_NO"
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
         TabIndex        =   7
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ADD"
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
         Height          =   375
         Left            =   3600
         TabIndex        =   6
         Top             =   6120
         Width           =   1575
      End
      Begin VB.Label Label12 
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
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   7560
         TabIndex        =   5
         Top             =   6120
         Width           =   1575
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   3480
         Shape           =   4  'Rounded Rectangle
         Top             =   6000
         Width           =   1815
      End
      Begin VB.Shape Shape4 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   7320
         Shape           =   4  'Rounded Rectangle
         Top             =   6000
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   12735
      Left            =   0
      Picture         =   "supplier1.frx":0000
      ScaleHeight     =   12675
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      CausesValidation=   0   'False
      Height          =   2355
      Left            =   4800
      Negotiate       =   -1  'True
      TabIndex        =   15
      Top             =   1320
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   4154
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   24
      RowDividerStyle =   1
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   9600
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UPDATE"
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
      Height          =   375
      Left            =   11280
      TabIndex        =   19
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ADD"
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
      Height          =   375
      Left            =   8640
      TabIndex        =   18
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   14160
      TabIndex        =   14
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label Label14 
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
      Height          =   375
      Left            =   14400
      TabIndex        =   13
      Top             =   1440
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Shape Shape7 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   14400
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Enter Supplier's Name"
      BeginProperty Font 
         Name            =   "Mongolian Baiti"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4920
      TabIndex        =   12
      Top             =   1320
      Visible         =   0   'False
      Width           =   4215
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SUPPLIER INFO.."
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
      TabIndex        =   11
      Top             =   120
      Width           =   6135
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   14040
      Shape           =   4  'Rounded Rectangle
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   11160
      Shape           =   4  'Rounded Rectangle
      Top             =   4200
      Width           =   1935
   End
   Begin VB.Shape Shape5 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   8520
      Shape           =   4  'Rounded Rectangle
      Top             =   4200
      Width           =   1935
   End
End
Attribute VB_Name = "supplier1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
If rsupplier.State = adStateOpen Then
    rsupplier.Close
End If
rsupplier.Open "select S_name,S_phno,S_address,S_email from supplier", conn, adOpenDynamic, adLockPessimistic
Set DataGrid1.DataSource = rsupplier

End Sub

Private Sub Label10_Click()
supplier2.Top = 0
supplier2.Left = 0
supplier2.Show
Unload Me
End Sub

Private Sub Label12_Click()
Frame1.Visible = False
End Sub

Private Sub Label14_Click()
If rsupplier.State = adStateOpen Then
rsupplier.Close
End If
rsupplier.Open "select * from [Supplier] where(S_name='" & Text1.Text & "')", conn, adOpenDynamic, adLockPessimistic
If rsupplier.RecordCount <> 0 Then
MsgBox "Supplier with that Name Exists!.", vbInformation, "ALERT"
Set DataGrid1.DataSource = rsupplier
DataGrid1.Visible = True
Frame1.Visible = False
Else
Frame1.Visible = True
End If
End Sub

Private Sub Label3_Click()
mainform.Left = 0
mainform.Top = 0
mainform.Show
Unload Me
End Sub

Private Sub Label4_Click()
If Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
    MsgBox "Some of your fields are incomplete. Please complete it", vbCritical, "Warning"
Else
    If rsupplier.State = adStateOpen Then
    rsupplier.Close
    End If
    rsupplier.Open "select * from [Supplier]", conn, adOpenDynamic, adLockPessimistic
    rsupplier.AddNew
    rsupplier!S_name = Text2.Text
    rsupplier!S_address = Text3.Text
    rsupplier!S_phno = Text4.Text
    rsupplier!S_email = Text5.Text
    rsupplier.Update
    MsgBox "SAVED SUCCESSFULLY...", vbInformation, "INFORMATION"
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Frame1.Visible = False
    If rsupplier.State = adStateOpen Then
        rsupplier.Close
    End If
    rsupplier.Open "select * from [Supplier]", conn, adOpenDynamic, adLockPessimistic
    Set DataGrid1.DataSource = rsupplier
End If
End Sub

Private Sub Label6_Click()
Frame1.Visible = True
End Sub

Private Sub Text2_LostFocus()
Call validateName(Text2.Text)
End Sub

Private Sub Text4_LostFocus()
Call validatePhone(Text4.Text)
End Sub

Private Sub Text5_LostFocus()
Call validateMail(Text5.Text)
End Sub
