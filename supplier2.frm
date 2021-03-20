VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form supplier2 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "supplier2"
   ClientHeight    =   12090
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   22770
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   12090
   ScaleWidth      =   22770
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "ADD medicine"
      Height          =   6375
      Left            =   5400
      TabIndex        =   1
      Top             =   3480
      Visible         =   0   'False
      Width           =   13215
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
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   3600
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
         TabIndex        =   3
         Top             =   3600
         Width           =   4695
      End
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
         TabIndex        =   2
         Top             =   1560
         Width           =   4695
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Name"
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
         TabIndex        =   11
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label8 
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
         TabIndex        =   10
         Top             =   960
         Width           =   3615
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
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Label Label4 
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
         Left            =   3360
         TabIndex        =   7
         Top             =   5400
         Width           =   2175
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
         Left            =   8040
         TabIndex        =   6
         Top             =   5400
         Width           =   2535
      End
      Begin VB.Shape Shape4 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   8040
         Shape           =   4  'Rounded Rectangle
         Top             =   5280
         Width           =   2535
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   3240
         Shape           =   4  'Rounded Rectangle
         Top             =   5280
         Width           =   2535
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   12735
      Left            =   0
      Picture         =   "supplier2.frx":0000
      ScaleHeight     =   12675
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      CausesValidation=   0   'False
      Height          =   2355
      Left            =   5400
      Negotiate       =   -1  'True
      TabIndex        =   17
      Top             =   1080
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
      Left            =   11640
      TabIndex        =   12
      Top             =   1200
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label Label9 
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
      Left            =   19320
      TabIndex        =   16
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label3 
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
      TabIndex        =   15
      Top             =   0
      Width           =   6135
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000D&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   4680
      Top             =   720
      Width           =   18015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Enter Supplier's ID"
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
      Left            =   6000
      TabIndex        =   14
      Top             =   1200
      Visible         =   0   'False
      Width           =   4935
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
      Height          =   495
      Left            =   16800
      TabIndex        =   13
      Top             =   1320
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape Shape7 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   16800
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   19320
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   2535
   End
End
Attribute VB_Name = "supplier2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DataGrid1_GotFocus()
 Text2.Text = DataGrid1.Columns(0)
    Text3.Text = DataGrid1.Columns(2)
    Text4.Text = DataGrid1.Columns(1)
    Text5.Text = DataGrid1.Columns(3)
    Frame1.Visible = True
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Text2.Text = DataGrid1.Columns(0)
    Text3.Text = DataGrid1.Columns(2)
    Text4.Text = DataGrid1.Columns(1)
    Text5.Text = DataGrid1.Columns(3)
    Frame1.Visible = True
End Sub

Private Sub Form_Load()
If rsupplier.State = adStateOpen Then
    rsupplier.Close
End If
rsupplier.Open "select S_name,S_phno,S_address,S_email from supplier", conn, adOpenDynamic, adLockPessimistic
Set DataGrid1.DataSource = rsupplier

End Sub

Private Sub Label12_Click()
Frame1.Visible = False
End Sub

Private Sub Label14_Click()
If rsupplier.State = adStateOpen Then
rsupplier.Close
End If
rsupplier.Open "select * from [Supplier] where(S_name='" & Text1.Text & "')", conn, adOpenDynamic, adLockPessimistic
If rsupplier.RecordCount = 1 Then
    Text2.Text = rsupplier!S_name
    Text3.Text = rsupplier!S_address
    Text4.Text = rsupplier!S_phno
    Text7.Text = rsupplier!S_amount
    Frame1.Visible = True
Else
MsgBox "INVALID SUPPLIER NAME", vbCritical, "ERRROR"
End If
End Sub

Private Sub Label4_Click()
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
End Sub

Private Sub Label9_Click()
supplier1.Left = 0
supplier1.Top = 0
supplier1.Show
Unload Me
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
