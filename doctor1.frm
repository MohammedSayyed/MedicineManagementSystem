VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form doctor1 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "doctor1"
   ClientHeight    =   14370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   23100
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   14370
   ScaleMode       =   0  'User
   ScaleWidth      =   23100
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "UPDATE batch"
      Height          =   9495
      Left            =   4920
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
      Width           =   17535
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000005&
         Caption         =   "FEMALE"
         BeginProperty Font 
            Name            =   "@Malgun Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9480
         TabIndex        =   15
         Top             =   2040
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000005&
         Caption         =   "MALE"
         BeginProperty Font 
            Name            =   "@Malgun Gothic"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   8040
         TabIndex        =   14
         Top             =   2040
         Width           =   1215
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
         Height          =   495
         Left            =   1560
         TabIndex        =   6
         Top             =   1920
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
         Height          =   495
         Left            =   1560
         TabIndex        =   5
         Top             =   3840
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
         Height          =   495
         Left            =   8040
         TabIndex        =   4
         Top             =   3840
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
         Height          =   495
         Left            =   8040
         TabIndex        =   3
         Top             =   5880
         Width           =   4695
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Yu Gothic Medium"
            Size            =   12
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   2
         Top             =   5880
         Width           =   4695
      End
      Begin VB.Label Label10 
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
         Left            =   8040
         TabIndex        =   21
         Top             =   7440
         Width           =   1935
      End
      Begin VB.Label Label104 
         BackStyle       =   0  'Transparent
         Caption         =   "Gender"
         BeginProperty Font 
            Name            =   "Yu Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8040
         TabIndex        =   13
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SAVE"
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
         Left            =   4440
         TabIndex        =   12
         Top             =   7440
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor Reg_no"
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
         Left            =   1560
         TabIndex        =   11
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor name"
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
         Left            =   1560
         TabIndex        =   10
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor PH_NO"
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
         Left            =   8040
         TabIndex        =   9
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor Degree"
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
         Left            =   8040
         TabIndex        =   8
         Top             =   5280
         Width           =   2655
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor Address"
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
         Left            =   1560
         TabIndex        =   7
         Top             =   5280
         Width           =   2655
      End
      Begin VB.Shape Shape5 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   4320
         Shape           =   4  'Rounded Rectangle
         Top             =   7320
         Width           =   1935
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   8040
         Shape           =   4  'Rounded Rectangle
         Top             =   7320
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   20415
      Left            =   -120
      Picture         =   "doctor1.frx":0000
      ScaleHeight     =   20355
      ScaleWidth      =   4755
      TabIndex        =   0
      Top             =   0
      Width           =   4815
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      CausesValidation=   0   'False
      Height          =   2355
      Left            =   5880
      Negotiate       =   -1  'True
      TabIndex        =   20
      Top             =   1920
      Width           =   14295
      _ExtentX        =   25215
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
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   11160
      TabIndex        =   22
      Top             =   1320
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Label Label14 
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
      Left            =   14640
      TabIndex        =   25
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ADD"
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
      Left            =   9120
      TabIndex        =   24
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UPDATE"
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
      Left            =   11880
      TabIndex        =   23
      Top             =   5160
      Width           =   1935
   End
   Begin VB.Label Label1 
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
      Left            =   18120
      TabIndex        =   19
      Top             =   1320
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "DOCTOR INFO.."
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
      Left            =   4680
      TabIndex        =   18
      Top             =   120
      Width           =   10215
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Available Doctors"
      BeginProperty Font 
         Name            =   "Mongolian Baiti"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5880
      TabIndex        =   17
      Top             =   1200
      Visible         =   0   'False
      Width           =   5295
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
      Left            =   15360
      TabIndex        =   16
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   15360
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   18120
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   9120
      Shape           =   4  'Rounded Rectangle
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Shape Shape7 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   11880
      Shape           =   4  'Rounded Rectangle
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Shape Shape8 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   14640
      Shape           =   4  'Rounded Rectangle
      Top             =   5040
      Width           =   1935
   End
End
Attribute VB_Name = "doctor1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flg As Boolean


Private Sub Form_Load()
If rdoctor.State = adStateOpen Then
    rdoctor.Close
    End If
rdoctor.Open "select * from [Doctor]", conn, adOpenDynamic, adLockPessimistic
Set DataGrid1.DataSource = rdoctor

End Sub

Private Sub Label1_Click()
docmenu.Left = 0
docmenu.Top = 0
docmenu.Show
Unload Me
End Sub

Private Sub Label10_Click()
Frame1.Visible = False

End Sub

Private Sub Label11_Click()
If Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or (Option1.Value Or Option2.Value) <> True Then
    MsgBox "Some of your fields are incomplete. Please complete it", vbCritical, "Warning"
Else
    If rdoctor.State = adStateOpen Then
    rdoctor.Close
    End If
    rdoctor.Open "select * from [Doctor]", conn, adOpenDynamic, adLockPessimistic
    rdoctor.AddNew
    rdoctor!D_regno = Text2.Text
    rdoctor!D_name = Text3.Text
    rdoctor!D_phno = Val(Text4.Text)
    rdoctor!D_degree = Text5.Text
    rdoctor!D_address = Text6.Text
    If Option1.Value = True Then
        rdoctor!D_gender = "Male"
    Else
        rdoctor!D_gender = "Female"
    End If
    rdoctor.Update
    MsgBox "SAVED SUCCESSFULLY...", vbInformation, "INFORMATION"
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    
End If
End Sub

Private Sub Label12_Click()
Frame1.Visible = True
End Sub

Private Sub Label14_Click()
mainform.Left = 0
mainform.Top = 0
mainform.Show
Unload Me

End Sub

Private Sub Label3_Click()
If rdoctor.State = adStateOpen Then
rdoctor.Close
End If
rdoctor.Open "select * from [Doctor] where(D_regno='" & Text1.Text & "')", conn, adOpenDynamic, adLockPessimistic
If rdoctor.RecordCount <> 0 Then
MsgBox "Doctor with that Registration number Exists!.", vbInformation, "ALERT"
Set DataGrid1.DataSource = rdoctor
DataGrid1.Visible = True
Frame1.Visible = False
Else
MsgBox "Doctor with that reg no does not exist!.", vbInformation, "INFO"
Text2.Text = Text1.Text
Text2.Enabled = False
Shape3.Visible = True
Label10.Visible = True
Frame1.Top = 1560
Frame1.Left = 4680
Frame1.Visible = True
End If
End Sub

Private Sub Label7_Click()
doctor2.Left = 0
doctor2.Top = 0
doctor2.Show
Unload Me
End Sub

Private Sub Option1_Click()
flg = True
End Sub

Private Sub Option2_Click()
flg = True
End Sub

Private Sub Text2_LostFocus()
If rdoctor.State = adStateOpen Then
    rdoctor.Close
End If
rdoctor.Open "select * from doctor where D_regno='" & Text2.Text & "'", conn, adOpenDynamic, adLockPessimistic
If rdoctor.RecordCount = 1 Then
    MsgBox "Doctor with that Registration Number Exists!!,please enter another Registration number!", vbInformation
End If
End Sub

Private Sub Text3_LostFocus()
Call validateName(Text3.Text)
End Sub

Private Sub Text4_LostFocus()
Call validatePhone(Text4.Text)
End Sub
