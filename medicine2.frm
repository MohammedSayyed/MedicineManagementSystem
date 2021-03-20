VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form medicine2 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "medicine2"
   ClientHeight    =   14430
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   22905
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   14430
   ScaleWidth      =   22905
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "ADD medicine"
      Height          =   10095
      Left            =   5040
      TabIndex        =   5
      Top             =   4320
      Visible         =   0   'False
      Width           =   13095
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
         TabIndex        =   14
         Text            =   "20"
         Top             =   8400
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
         Height          =   495
         Left            =   8280
         TabIndex        =   13
         Top             =   7200
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
         Left            =   840
         TabIndex        =   12
         Top             =   6960
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
         Left            =   8280
         TabIndex        =   11
         Top             =   5520
         Width           =   4695
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   840
         TabIndex        =   10
         Top             =   5520
         Width           =   4815
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
         Left            =   840
         TabIndex        =   9
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
         Height          =   495
         Left            =   840
         TabIndex        =   8
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
         Height          =   495
         Left            =   8280
         TabIndex        =   7
         Top             =   1560
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
         Left            =   8280
         TabIndex        =   6
         Top             =   3600
         Width           =   4695
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Medcine Profit%"
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
         TabIndex        =   25
         Top             =   7920
         Width           =   3735
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Medicine SP"
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
         TabIndex        =   24
         Top             =   6600
         Width           =   3735
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Medicine CP"
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
         TabIndex        =   23
         Top             =   6480
         Width           =   3735
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Medicine MRP"
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
         TabIndex        =   22
         Top             =   4920
         Width           =   3735
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Medicine Name"
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
         TabIndex        =   21
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Medicine Company"
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Medcine Packing"
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
         Top             =   3000
         Width           =   3735
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "UPDATE"
         BeginProperty Font 
            Name            =   "MS UI Gothic"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8400
         TabIndex        =   16
         Top             =   8760
         Width           =   1935
      End
      Begin VB.Label Label16 
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
         Height          =   495
         Left            =   10920
         TabIndex        =   15
         Top             =   8760
         Width           =   1815
      End
      Begin VB.Shape Shape5 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   10920
         Shape           =   4  'Rounded Rectangle
         Top             =   8640
         Width           =   1815
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   8400
         Shape           =   4  'Rounded Rectangle
         Top             =   8640
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture1 
      FillStyle       =   0  'Solid
      Height          =   14415
      Left            =   0
      Picture         =   "medicine2.frx":0000
      ScaleHeight     =   14355
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      CausesValidation=   0   'False
      Height          =   2835
      Left            =   5040
      Negotiate       =   -1  'True
      TabIndex        =   27
      Top             =   1320
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   5001
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
         Size            =   8.25
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
         Weight          =   700
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
      Left            =   12240
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label Label18 
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
      Height          =   495
      Left            =   20280
      TabIndex        =   26
      Top             =   2160
      Width           =   2175
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
      Left            =   17520
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Shape Shape7 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   17400
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label7 
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
      Left            =   5040
      TabIndex        =   3
      Top             =   240
      Width           =   6135
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000D&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   4800
      Top             =   960
      Width           =   18015
   End
   Begin VB.Label Label1 
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
      Left            =   6600
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   20280
      Shape           =   4  'Rounded Rectangle
      Top             =   2040
      Width           =   2175
   End
End
Attribute VB_Name = "medicine2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DataGrid1_Click()
Text2.Text = DataGrid1.Columns(0)
    Text3.Text = DataGrid1.Columns(4)
    Text4.Text = DataGrid1.Columns(2)
    Text5.Text = DataGrid1.Columns(3)
    Text6.Text = DataGrid1.Columns(5)
    Text7.Text = DataGrid1.Columns(6)
    Text8.Text = DataGrid1.Columns(7)
    Text9.Text = DataGrid1.Columns(8)
    Combo1.Text = DataGrid1.Columns(1)
Frame1.Visible = True
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Text2.Text = DataGrid1.Columns(0)
    Text3.Text = DataGrid1.Columns(4)
    Text4.Text = DataGrid1.Columns(2)
    Text5.Text = DataGrid1.Columns(3)
    Text6.Text = DataGrid1.Columns(5)
    Text7.Text = DataGrid1.Columns(6)
    Text8.Text = DataGrid1.Columns(7)
    Text9.Text = DataGrid1.Columns(8)
    Combo1.Text = DataGrid1.Columns(1)
Frame1.Visible = True
End Sub

Private Sub Form_Load()
Combo1.AddItem "syp"
Combo1.AddItem "IV"
Combo1.AddItem "tab"
Combo1.AddItem "cap"

If rmedicine.State = adStateOpen Then
    rmedicine.Close
End If
rmedicine.Open "select m_name,m_type,m_color,m_packing,m_company,m_mrp,m_cp,m_sp,m_pf from [Medicine]", conn, adOpenDynamic, adLockPessimistic
Set DataGrid1.DataSource = rmedicine
DataGrid1.Refresh
End Sub

Private Sub Label10_Click()
    rmedicine!M_name = Text2.Text
    rmedicine!M_company = Text3.Text
    rmedicine!M_color = Text4.Text
    rmedicine!M_packing = Text5.Text
    rmedicine!M_type = Combo1.Text
    rmedicine!M_mrp = Val(Text6.Text)
    rmedicine!M_cp = Val(Text7.Text)
    rmedicine!M_sp = Val(Text8.Text)
    rmedicine!M_pf = Val(Text9.Text)
    rmedicine.Update
    MsgBox "SAVED SUCCESSFULLY...", vbInformation, "INFORMATION"
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Text7.Text = ""
    Text8.Text = ""
    Text9.Text = ""
    Combo1.Text = ""
    Frame1.Visible = False
End Sub

Private Sub Label12_Click()
Unload Me
mainform.Top = 0
mainform.Left = 0
mainform.Show
End Sub

Private Sub Label14_Click()
If rmedicine.State = adStateOpen Then
rmedicine.Close
End If
rmedicine.Open "select * from [Medicine] where(M_name = '" & Text1.Text & "')", conn, adOpenDynamic, adLockPessimistic
If rmedicine.RecordCount <> 1 Then
    MsgBox "PLEASE ENTER A VALID MEDICINE NAME", vbCritical, "ERROR"
Else
    Text2.Text = rmedicine!M_name
    Text3.Text = rmedicine!M_company
    Text4.Text = rmedicine!M_color
    Text5.Text = rmedicine!M_packing
    Text6.Text = rmedicine!M_mrp
    Text7.Text = rmedicine!M_cp
    Text8.Text = rmedicine!M_sp
    Text9.Text = rmedicine!M_pf
    Combo1.Text = rmedicine!M_type
    Frame1.Visible = True
End If
End Sub

Private Sub Label16_Click()
Frame1.Visible = False
End Sub


Private Sub Label17_Click()
If rmedicine.State = adStateOpen Then
   rmedicine.Close
End If
rmedicine.Open "delete from [Medicine] where(M_name='" & Text2.Text & "')", conn, adOpenDynamic, adLockPessimistic
rmedicine.Open "select * from [Medicine] where(M_name='" & Text2.Text & "')", conn, adOpenDynamic, adLockPessimistic
If rmedicine.RecordCount = 0 Then
    MsgBox "Medicine DELETED SUCCESSFULLY!!", vbInformation, "INFO"
    Frame1.Visible = False
End If
End Sub

Private Sub Label18_Click()
medicine1.Left = 0
medicine1.Top = 0
medicine1.Show
Unload Me
End Sub

Private Sub Label19_Click()

End Sub

Private Sub Text2_LostFocus()
Call validateName(Text2.Text)
End Sub

Private Sub Text3_Change()
Call validateName(Text3.Text)
End Sub

Private Sub Text9_LostFocus()
Text8.Text = Val(Text7.Text) + (Val(Text7.Text) * (Val(Text9.Text) / 100))
End Sub


