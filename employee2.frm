VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form employee2 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "employee2"
   ClientHeight    =   14055
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   21945
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   14055
   ScaleMode       =   0  'User
   ScaleWidth      =   21945
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "UPDATE batch"
      Height          =   11175
      Left            =   4920
      TabIndex        =   2
      Top             =   3600
      Visible         =   0   'False
      Width           =   13335
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
         Left            =   1200
         TabIndex        =   30
         Top             =   7560
         Width           =   4695
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000007&
         Height          =   375
         Left            =   6000
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   7560
         Width           =   375
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
         TabIndex        =   12
         Top             =   2280
         Width           =   1935
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
         TabIndex        =   11
         Top             =   2400
         Width           =   1455
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
         Left            =   8520
         TabIndex        =   10
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
         Height          =   525
         Left            =   1200
         TabIndex        =   9
         Top             =   840
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
         Left            =   1200
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   4800
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
         Left            =   1200
         TabIndex        =   5
         Top             =   3480
         Width           =   4695
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
         Left            =   1200
         TabIndex        =   4
         Top             =   4800
         Width           =   5655
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
         Left            =   1200
         TabIndex        =   3
         Top             =   6120
         Width           =   5655
      End
      Begin MSACAL.Calendar Calendar1 
         Height          =   3015
         Left            =   1200
         TabIndex        =   28
         Top             =   8160
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
      Begin VB.Label Label4 
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
         Left            =   1200
         TabIndex        =   31
         Top             =   6840
         Width           =   2655
      End
      Begin VB.Label Label16 
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
         Height          =   375
         Left            =   8760
         TabIndex        =   24
         Top             =   7680
         Width           =   1695
      End
      Begin VB.Label Label19 
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
         TabIndex        =   22
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label15 
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
         Left            =   1200
         TabIndex        =   21
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label10 
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
         TabIndex        =   20
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label9 
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
         TabIndex        =   19
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Label Label8 
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
         Left            =   1200
         TabIndex        =   18
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label7 
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
         Left            =   1200
         TabIndex        =   17
         Top             =   3000
         Width           =   3735
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
         Left            =   1200
         TabIndex        =   16
         Top             =   4320
         Width           =   2655
      End
      Begin VB.Label Label1 
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
         Left            =   1200
         TabIndex        =   15
         Top             =   5640
         Width           =   2655
      End
      Begin VB.Label Label161 
         BackStyle       =   0  'Transparent
         Caption         =   "Employee ID"
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
         Left            =   1200
         TabIndex        =   14
         Top             =   -960
         Width           =   2895
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
         TabIndex        =   13
         Top             =   4320
         Width           =   2895
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   8640
         Shape           =   4  'Rounded Rectangle
         Top             =   7560
         Width           =   1935
      End
      Begin VB.Label Label11 
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
         Left            =   11160
         TabIndex        =   26
         Top             =   7680
         Width           =   1815
      End
      Begin VB.Shape Shape5 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   11040
         Shape           =   4  'Rounded Rectangle
         Top             =   7560
         Width           =   1935
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   360
      Top             =   120
   End
   Begin VB.PictureBox Picture2 
      Height          =   15375
      Left            =   0
      Picture         =   "employee2.frx":0000
      ScaleHeight     =   15315
      ScaleWidth      =   4395
      TabIndex        =   1
      Top             =   0
      Width           =   4455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      CausesValidation=   0   'False
      Height          =   2355
      Left            =   4920
      Negotiate       =   -1  'True
      TabIndex        =   27
      Top             =   1200
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
   Begin VB.Label Label5 
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
      Left            =   19080
      TabIndex        =   25
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "EMPLOYEE UPDATE.."
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
      TabIndex        =   23
      Top             =   0
      Width           =   7215
   End
   Begin VB.Shape Shape10 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000D&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   4680
      Top             =   720
      Width           =   18015
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
      Left            =   17880
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   17880
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
      Left            =   19080
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   1935
   End
End
Attribute VB_Name = "employee2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Private Sub DataGrid1_GotFocus()
Text10.Text = DataGrid1.Columns(10)
    Text3.Text = DataGrid1.Columns(0)
    Text4.Text = DataGrid1.Columns(6)
    Text5.Text = DataGrid1.Columns(1)
    Text6.Text = DataGrid1.Columns(7)
    Text7.Text = DataGrid1.Columns(8)
    Text8.Text = DataGrid1.Columns(2)
    Text9.Text = DataGrid1.Columns(5)
    If DataGrid1.Columns(9) = "Male" Then
        Option1.Value = True
    Else
        Option2.Value = True
    End If
    Combo1.ListIndex = DataGrid1.Columns(3)
    
Frame1.Visible = True
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Text10.Text = DataGrid1.Columns(10)
    Text3.Text = DataGrid1.Columns(0)
    Text4.Text = DataGrid1.Columns(6)
    Text5.Text = DataGrid1.Columns(1)
    Text6.Text = DataGrid1.Columns(7)
    Text7.Text = DataGrid1.Columns(8)
    Text8.Text = DataGrid1.Columns(2)
    Text9.Text = DataGrid1.Columns(5)
    If DataGrid1.Columns(9) = "Male" Then
        Option1.Value = True
    Else
        Option2.Value = True
    End If
    Combo1.ListIndex = DataGrid1.Columns(3)
Frame1.Visible = True
End Sub


Private Sub Form_Load()
Combo1.AddItem ("WHAT IS YOUR PET NAME?")
Combo1.AddItem ("WHAT IS YOUR BEST FRIENDS NAME?")
Combo1.AddItem ("WHAT IS YOUR FAVOURITE DISH?")

If remployee.State = adStateOpen Then
remployee.Close
End If
remployee.Open "select E_name,E_email,E_address,E_question,E_question1,E_answer,E_phno,E_username,E_password,E_gender,E_dateOfJoining from [Employee]", conn, adOpenDynamic, adLockPessimistic
Set DataGrid1.DataSource = remployee
End Sub

Private Sub Label11_Click()
Frame1.Visible = False
End Sub

Private Sub Label16_Click()
remployee!E_name = Text3.Text
remployee!E_phno = Val(Text4.Text)
remployee!E_email = Text5.Text
remployee!E_username = Text6.Text
remployee!E_password = Text7.Text
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
MsgBox "UPDATED SUCCESSFULLY...", vbInformation, "INFORMATION"
Text10.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    'Text1.Text = ""
    Combo1.Text = ""
    Option1.Value = False
    Option2.Value = False
Frame1.Visible = False
'Label11.Visible = False
'Shape5.Visible = False


End Sub

Private Sub Label3_Click()
If remployee.State = adStateOpen Then
remployee.Close
End If
remployee.Open "select * from [Employee] where(E_id = " & Val(Text1.Text) & ")", conn, adOpenDynamic, adLockPessimistic
If remployee.RecordCount <> 1 Then
    MsgBox "PLEASE ENTER A VALID EMPLOYEE ID OF AN EXISTING EMPLOYEE!!", vbCritical, "ERROR"
Else
    Text2.Text = remployee!E_id
    Text3.Text = remployee!E_name
    Text4.Text = remployee!E_phno
    Text5.Text = remployee!E_email
    Text6.Text = remployee!E_username
    Text7.Text = remployee!E_password
    Text8.Text = remployee!E_address
    Text9.Text = remployee!E_answer
    If remployee!E_gender = "Male" Then
        Option1.Value = True
    Else
        Option2.Value = True
    End If
    Combo1.ListIndex = remployee!E_question
    Frame1.Top = 1320
    Frame1.Left = 4800
    Frame1.Visible = True
    Label11.Visible = True
    Shape5.Visible = True
End If
End Sub

Private Sub Label4_Click()
If remployee.State = adStateOpen Then
   remployee.Close
End If
remployee.Open "delete from [Employee] where(E_id=" & Val(Text2.Text) & ")", conn, adOpenDynamic, adLockPessimistic
remployee.Open "select * from [Employee] where(E_id=" & Val(Text2.Text) & ")", conn, adOpenDynamic, adLockPessimistic
If remployee.RecordCount = 0 Then
    MsgBox "EMPLOYEE DELETED SUCCESSFULLY!!", vbInformation, "INFO"
End If
Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Text1.Text = ""
    Combo1.Text = ""
    Option1.Value = False
    Option2.Value = False
Frame1.Visible = False
Label11.Visible = False
Shape5.Visible = False
End Sub

Private Sub Label5_Click()
employee1.Left = 0
employee1.Top = 0
employee1.Show
Unload Me
End Sub

Private Sub Text3_LostFocus()
Call validateName(Text3.Text)
End Sub

Private Sub Text4_LostFocus()
Call validatePhone(Text4.Text)
End Sub

Private Sub Text5_LostFocus()
Call validateMail(Text5.Text)
End Sub

Private Sub Text6_LostFocus()
If rs.State = adStateOpen Then
rs.Close
End If
rs.Open "select * from [Employee] where(E_username='" & Text6.Text & "')", conn, adOpenDynamic, adLockPessimistic
If rs.RecordCount <> 0 Then
    MsgBox "EMPLOYEE USERNAME ALREADY IN USE BY SOME OTHER EMPLOYEE, PLEASE CHOOSE ANOTHER USERNAME", vbCritical, "ERROR"
    Text6.Text = ""
    Text6.SetFocus

End If
End Sub
