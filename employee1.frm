VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form employee1 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "employee1"
   ClientHeight    =   13620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   22695
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   13620
   ScaleMode       =   0  'User
   ScaleWidth      =   22695
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "UPDATE batch"
      Height          =   11415
      Left            =   5040
      TabIndex        =   1
      Top             =   1080
      Visible         =   0   'False
      Width           =   13575
      Begin MSACAL.Calendar Calendar1 
         Height          =   3015
         Left            =   840
         TabIndex        =   35
         Top             =   8280
         Visible         =   0   'False
         Width           =   4695
         _Version        =   524288
         _ExtentX        =   8281
         _ExtentY        =   5318
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2019
         Month           =   10
         Day             =   2
         DayLength       =   1
         MonthLength     =   1
         DayFontColor    =   0
         FirstDay        =   7
         GridCellEffect  =   1
         GridFontColor   =   10485760
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   -1  'True
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   10485760
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000007&
         Height          =   375
         Left            =   5640
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   7800
         Width           =   375
      End
      Begin VB.TextBox Text10 
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
         Left            =   840
         TabIndex        =   32
         Top             =   7800
         Width           =   4695
      End
      Begin VB.TextBox Text9 
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
         Left            =   840
         TabIndex        =   23
         Top             =   6120
         Width           =   5655
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Yu Gothic Medium"
            Size            =   12
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   840
         TabIndex        =   22
         Top             =   4800
         Width           =   5655
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
         Left            =   840
         TabIndex        =   20
         Top             =   3480
         Width           =   4695
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Yu Gothic Medium"
            Size            =   12
            Charset         =   0
            Weight          =   500
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   8520
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   6120
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
         Left            =   8520
         TabIndex        =   8
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
         Height          =   495
         Left            =   840
         TabIndex        =   7
         Top             =   2160
         Width           =   4695
      End
      Begin VB.TextBox Text7 
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
         Left            =   8520
         TabIndex        =   6
         Top             =   4800
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
         Height          =   525
         Left            =   840
         TabIndex        =   5
         Top             =   840
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
         Left            =   8520
         TabIndex        =   4
         Top             =   840
         Width           =   4695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000005&
         Caption         =   "MALE"
         BeginProperty Font 
            Name            =   "@Malgun Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8520
         TabIndex        =   3
         Top             =   2400
         Width           =   1455
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000005&
         Caption         =   "FEMALE"
         BeginProperty Font 
            Name            =   "@Malgun Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   10080
         TabIndex        =   2
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Date Of Joining"
         BeginProperty Font 
            Name            =   "Yu Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   840
         TabIndex        =   33
         Top             =   6960
         Width           =   2655
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   10920
         TabIndex        =   30
         Top             =   9960
         Width           =   1935
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Addresss"
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
         Left            =   8520
         TabIndex        =   24
         Top             =   5640
         Width           =   2895
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Answer"
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
         TabIndex        =   19
         Top             =   5640
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Select a Question"
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
         Top             =   4320
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Confirm Password"
         BeginProperty Font 
            Name            =   "Yu Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8520
         TabIndex        =   17
         Top             =   4320
         Width           =   2655
      End
      Begin VB.Label Label16 
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
         Left            =   8760
         TabIndex        =   15
         Top             =   9960
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Username"
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
         TabIndex        =   14
         Top             =   3000
         Width           =   3735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee e-mail ID"
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
         TabIndex        =   13
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Yu Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8520
         TabIndex        =   12
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee PH_NO"
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
         Left            =   8520
         TabIndex        =   11
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee Name"
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
         TabIndex        =   10
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label10 
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
         Left            =   8520
         TabIndex        =   9
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   8760
         Shape           =   4  'Rounded Rectangle
         Top             =   9840
         Width           =   1935
      End
      Begin VB.Shape Shape6 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   10920
         Shape           =   4  'Rounded Rectangle
         Top             =   9840
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   14295
      Left            =   -120
      Picture         =   "employee1.frx":0000
      ScaleHeight     =   14235
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   -240
      Width           =   4695
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   360
      Top             =   120
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      CausesValidation=   0   'False
      Height          =   4275
      Left            =   5040
      Negotiate       =   -1  'True
      TabIndex        =   29
      Top             =   1560
      Width           =   16575
      _ExtentX        =   29236
      _ExtentY        =   7541
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
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   12360
      TabIndex        =   31
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   15000
      TabIndex        =   28
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ADD NEW"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9720
      TabIndex        =   27
      Top             =   6480
      Width           =   1935
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Employees:"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   26
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   11160
      TabIndex        =   25
      Top             =   15960
      Width           =   1935
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
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "EMPLOYEE INFO.."
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
      TabIndex        =   21
      Top             =   120
      Width           =   6015
   End
   Begin VB.Shape Shape5 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   11160
      Shape           =   4  'Rounded Rectangle
      Top             =   15840
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   9720
      Shape           =   4  'Rounded Rectangle
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   15000
      Shape           =   4  'Rounded Rectangle
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Shape Shape7 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   12360
      Shape           =   4  'Rounded Rectangle
      Top             =   6360
      Width           =   1935
   End
End
Attribute VB_Name = "employee1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Calendar1_Click()
Text10.Text = Calendar1.Value
Calendar1.Visible = False
End Sub

Private Sub Command1_Click()
Calendar1.Visible = True
End Sub

Private Sub Form_Load()
Timer1.Enabled = True
Combo1.AddItem ("WHAT IS YOUR PET NAME?")
Combo1.AddItem ("WHAT IS YOUR BEST FRIENDS NAME?")
Combo1.AddItem ("WHAT IS YOUR FAVOURITE DISH?")
If remployee.State = adStateOpen Then
remployee.Close
End If
remployee.Open "select E_name,E_email,E_address,E_question,E_question1,E_answer,E_phno,E_username,E_password,E_gender,E_dateOfJoining from [Employee]", conn, adOpenDynamic, adLockPessimistic
Set DataGrid1.DataSource = remployee
DataGrid1.Refresh


End Sub


Private Sub Label13_Click()
employee2.Left = 0
employee2.Top = 0
employee2.Show
Unload Me
End Sub

Private Sub Label14_Click()
Frame1.Visible = True
DataGrid1.Visible = False
End Sub

Private Sub Label15_Click()
mainform.Top = 0
mainform.Left = 0
mainform.Show
Unload Me
End Sub

Private Sub Label16_Click()
If Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Or (Option1.Value Or Option2.Value) <> True Then
MsgBox "SOME FIELDS ARE INCOMPLETE PLEASE COMPLETE THEM"
Else
If remployee.State = adStateOpen Then
remployee.Close
End If
remployee.Open "select * from [Employee]", conn, adOpenDynamic, adLockPessimistic
remployee.AddNew
remployee!E_name = Text2.Text
remployee!E_phno = Text3.Text
remployee!E_email = Text4.Text
remployee!E_username = Text5.Text
remployee!E_password = Text6.Text
remployee!E_address = Text8.Text
remployee!E_question = Combo1.ListIndex
remployee!E_question1 = Combo1.Text
remployee!E_answer = Text9.Text
If Option1.Value = True Then
    remployee!E_gender = "Male"
Else
    remployee!E_gender = "Female"
End If
remployee!E_dateOfJoining = Text10.Text
remployee.Update
MsgBox "SAVED SUCCESSFULLY...", vbInformation, "INFORMATION"
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text10.Text = ""
    Combo1.Text = ""
    Option1.Value = False
    Option2.Value = False
 

End If



End Sub



Private Sub Label18_Click()
Frame1.Visible = False
If remployee.State = adStateOpen Then
remployee.Close
End If
remployee.Open "select E_name,E_email,E_address,E_question,E_question1,E_answer,E_phno,E_username,E_password,E_gender,E_dateOfJoining from [Employee]", conn, adOpenDynamic, adLockPessimistic
Set DataGrid1.DataSource = remployee
DataGrid1.Refresh
DataGrid1.Visible = True
End Sub

Private Sub Text1_LostFocus()

If remployee.State = adStateOpen Then
remployee.Close
End If
remployee.Open "select * from [Employee] where(E_id=" & Val(Text1.Text) & ")", conn, adOpenDynamic, adLockPessimistic
If remployee.RecordCount <> 0 Then
    MsgBox "EMPLOYEE ID ALREADY IN USE BY SOME OTHER EMPLOYEE, PLEASE CHOOSE ANOTHER ID", vbCritical, "ERROR"
   Text1.Text = ""
    Text1.SetFocus
End If
If Text1.Text = "" Then
MsgBox "PLEASE ENTER EMPLOYEE ID FIRST", vbInformation, "INFO"
 Text1.Text = ""
    Text1.SetFocus
End If

End Sub

Private Sub Text2_LostFocus()
Call validateName(Text2.Text)
End Sub

Private Sub Text3_LostFocus()
Call validatePhone(Text3.Text)
End Sub

Private Sub Text4_LostFocus()
Call validateMail(Text4.Text)
End Sub

Private Sub Text5_LostFocus()
If remployee.State = adStateOpen Then
remployee.Close
End If
remployee.Open "select * from [Employee] where(E_username='" & Text5.Text & "')", conn, adOpenDynamic, adLockPessimistic
If remployee.RecordCount <> 0 Then
    MsgBox "EMPLOYEE USERNAME ALREADY IN USE BY SOME OTHER EMPLOYEE, PLEASE CHOOSE ANOTHER USERNAME", vbCritical, "ERROR"
    Text5.Text = ""
    Text5.SetFocus

End If
End Sub

Private Sub Text7_LostFocus()
If Text6.Text <> Text7.Text Then
    MsgBox "PASSWORDS DO NOT MATCH,PLEASE ENTER THE PASSWORD AGAIN", vbCritical, "ERROR"
    Text7.Text = ""
    Text7.SetFocus
End If

End Sub

Private Sub Timer1_Timer()
'Text1.SetFocus
Timer1.Enabled = False
End Sub

