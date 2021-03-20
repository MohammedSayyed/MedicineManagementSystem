VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form doctor2 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "doctor2"
   ClientHeight    =   13560
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   22695
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   678
   ScaleMode       =   0  'User
   ScaleWidth      =   1134.75
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "UPDATE batch"
      Height          =   9255
      Left            =   5400
      TabIndex        =   1
      Top             =   4200
      Visible         =   0   'False
      Width           =   13215
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
         Left            =   1560
         TabIndex        =   24
         Top             =   7320
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
         TabIndex        =   8
         Top             =   5880
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
         Left            =   7920
         TabIndex        =   7
         Top             =   6000
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
         Left            =   7920
         TabIndex        =   6
         Top             =   3840
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
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
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
         TabIndex        =   4
         Top             =   1920
         Width           =   4695
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
         TabIndex        =   3
         Top             =   2040
         Width           =   1215
      End
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
         TabIndex        =   2
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor Total Sales"
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
         TabIndex        =   25
         Top             =   6720
         Width           =   2655
      End
      Begin VB.Label Label7 
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
         Left            =   7560
         TabIndex        =   21
         Top             =   8520
         Width           =   1815
      End
      Begin VB.Label Label12 
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
         Left            =   4200
         TabIndex        =   15
         Top             =   8520
         Width           =   1695
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor Clinic Address"
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
         TabIndex        =   14
         Top             =   5280
         Width           =   3375
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
         Left            =   7920
         TabIndex        =   13
         Top             =   5280
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
         Left            =   7920
         TabIndex        =   12
         Top             =   3240
         Width           =   2655
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
         TabIndex        =   11
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor Reg_NO"
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
         Top             =   1320
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
         Left            =   8040
         TabIndex        =   9
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   4080
         Shape           =   4  'Rounded Rectangle
         Top             =   8400
         Width           =   1935
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   7560
         Shape           =   4  'Rounded Rectangle
         Top             =   8400
         Width           =   1935
      End
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   10440
      TabIndex        =   16
      Top             =   960
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.PictureBox Picture1 
      Height          =   13455
      Left            =   0
      Picture         =   "doctor2.frx":0000
      ScaleHeight     =   13395
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      CausesValidation=   0   'False
      Height          =   2355
      Left            =   5400
      Negotiate       =   -1  'True
      TabIndex        =   22
      Top             =   1800
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
            ColumnWidth     =   1515.005
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   1515.005
         EndProperty
      EndProperty
   End
   Begin VB.Label Label11z 
      BackStyle       =   0  'Transparent
      Caption         =   "Click on the record to update"
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
      Left            =   5400
      TabIndex        =   23
      Top             =   1320
      Width           =   5295
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
      Left            =   19080
      TabIndex        =   20
      Top             =   2760
      Width           =   1815
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
      Left            =   15000
      TabIndex        =   19
      Top             =   1080
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please enter doctors reg_no"
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
      Left            =   5400
      TabIndex        =   18
      Top             =   960
      Visible         =   0   'False
      Width           =   5295
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
   Begin VB.Label Label32 
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
      TabIndex        =   17
      Top             =   0
      Width           =   10215
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   15000
      Shape           =   4  'Rounded Rectangle
      Top             =   960
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
      Top             =   2640
      Width           =   1935
   End
End
Attribute VB_Name = "doctor2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DataGrid1_GotFocus()
Text2.Text = DataGrid1.Columns(0)
Text3.Text = DataGrid1.Columns(1)
Text4.Text = DataGrid1.Columns(2)
Text5.Text = DataGrid1.Columns(3)
Text6.Text = DataGrid1.Columns(4)
If DataGrid1.Columns(5) = "Male" Then
    Option1.Value = True
Else
    Option2.Value = True
End If
'Text7.Text = DataGrid1.Columns(6)
Frame1.Visible = True
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Text2.Text = DataGrid1.Columns(0)
Text3.Text = DataGrid1.Columns(1)
Text4.Text = DataGrid1.Columns(2)
Text5.Text = DataGrid1.Columns(3)
Text6.Text = DataGrid1.Columns(4)
If DataGrid1.Columns(5) = "Male" Then
    Option1.Value = True
Else
    Option2.Value = True
End If
'Text7.Text = DataGrid1.Columns(6)
Frame1.Visible = True
End Sub

Private Sub Form_Load()
If rdoctor.State = adStateOpen Then
rdoctor.Close
End If
rdoctor.Open "select D_regno,D_name,D_phno,D_degree,D_address,D_gender from [Doctor]", conn, adOpenDynamic, adLockPessimistic
Set DataGrid1.DataSource = rdoctor
DataGrid1.Visible = True
End Sub

Private Sub Label1_Click()
doctor1.Left = 0
doctor1.Top = 0
doctor1.Show
Unload Me
End Sub

Private Sub Label11_Click()
If rdoctor.State = adStateOpen Then
   rdoctor.Close
End If
rdoctor.Open "delete from [Doctor] where(D_regno='" & Text2.Text & "')", conn, adOpenDynamic, adLockPessimistic
rdoctor.Open "select * from [Doctor] where(D_regno='" & Text2.Text & "')", conn, adOpenDynamic, adLockPessimistic
If rdoctor.RecordCount = 0 Then
    MsgBox "DOCTOR DELETED SUCCESSFULLY!!", vbInformation, "INFO"
End If
Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
End Sub

Private Sub Label12_Click()
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
MsgBox "UPDATED SUCCESSFULLY...", vbInformation, "INFORMATION"
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text1.Text = ""
Option1.Value = False
Option2.Value = False
Frame1.Visible = False
End Sub

Private Sub Label14_Click()
Frame1.Visible = False
Label14.Visible = False
Shape5.Visible = False
End Sub

Private Sub Label3_Click()
If rdoctor.State = adStateOpen Then
rdoctor.Close
End If
rdoctor.Open "select * from [Doctor] where(D_regno = " & Val(Text1.Text) & ")", conn, adOpenDynamic, adLockPessimistic
If rdoctor.RecordCount <> 1 Then
    MsgBox "PLEASE ENTER A VALID REG NO OF AN EXISTING DOCTOR", vbCritical, "ERROR"
Else
    Text2.Text = rdoctor!D_regno
    Text3.Text = rdoctor!D_name
    Text4.Text = rdoctor!D_phno
    Text5.Text = rdoctor!D_degree
    Text6.Text = rdoctor!D_address
    If rdoctor!D_gender = "Male" Then
        Option1.Value = True
    Else
        Option2.Value = True
    End If
    Shape5.Visible = True
    Label14.Visible = True
    Frame1.Visible = True
End If
End Sub

Private Sub Label7_Click()
Frame1.Visible = False
If remployee.State = adStateOpen Then
    remployee.Close
End If
remployee.Open "select * from e"
DataGrid1.Refresh
End Sub

Private Sub Text3_LostFocus()
Call validateName(Text3.Text)
End Sub

Private Sub Text4_LostFocus()
Call validatePhone(Text4.Text)

End Sub
