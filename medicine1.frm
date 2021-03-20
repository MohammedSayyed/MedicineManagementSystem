VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form medicine1 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "0"
   ClientHeight    =   14130
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   23370
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   14130
   ScaleWidth      =   23370
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "ADD medicine"
      Height          =   10935
      Left            =   4800
      TabIndex        =   15
      Top             =   1560
      Visible         =   0   'False
      Width           =   17175
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
         TabIndex        =   5
         Text            =   "20"
         Top             =   8640
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
         TabIndex        =   25
         Top             =   7200
         Width           =   4695
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   4
         Top             =   7200
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
         TabIndex        =   8
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
         TabIndex        =   3
         Top             =   5520
         Width           =   4815
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
         Left            =   840
         TabIndex        =   1
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
         TabIndex        =   2
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
         TabIndex        =   6
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
         TabIndex        =   7
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
         TabIndex        =   27
         Top             =   8160
         Width           =   3735
      End
      Begin VB.Label Label14 
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
         TabIndex        =   26
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
         TabIndex        =   24
         Top             =   6720
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
         TabIndex        =   23
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         Height          =   375
         Left            =   8400
         TabIndex        =   9
         Top             =   8760
         Width           =   1935
      End
      Begin VB.Label Label10 
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
         Left            =   11400
         TabIndex        =   10
         Top             =   8760
         Width           =   2415
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
      Begin VB.Shape Shape5 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   11280
         Shape           =   4  'Rounded Rectangle
         Top             =   8640
         Width           =   2535
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   21975
      Left            =   0
      Picture         =   "medicine1.frx":0000
      ScaleHeight     =   21915
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   -1080
      Width           =   4455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      CausesValidation=   0   'False
      Height          =   2835
      Left            =   4800
      Negotiate       =   -1  'True
      TabIndex        =   21
      Top             =   1680
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   5001
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   29
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
         Size            =   13.5
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
      Left            =   10080
      TabIndex        =   13
      Top             =   1800
      Visible         =   0   'False
      Width           =   5655
   End
   Begin VB.Label Label17 
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
      Height          =   495
      Left            =   11880
      TabIndex        =   29
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Label Label16 
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
      Height          =   495
      Left            =   8880
      TabIndex        =   28
      Top             =   5520
      Width           =   2175
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
      Height          =   495
      Left            =   14880
      TabIndex        =   22
      Top             =   5520
      Width           =   2175
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   14880
      Shape           =   4  'Rounded Rectangle
      Top             =   5400
      Width           =   2175
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
      Left            =   16200
      TabIndex        =   14
      Top             =   1920
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   16200
      Shape           =   4  'Rounded Rectangle
      Top             =   1800
      Visible         =   0   'False
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
      TabIndex        =   12
      Top             =   1800
      Visible         =   0   'False
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
      Caption         =   "MEDICINE INFO..."
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
      TabIndex        =   11
      Top             =   240
      Width           =   6135
   End
   Begin VB.Shape Shape7 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   11880
      Shape           =   4  'Rounded Rectangle
      Top             =   5400
      Width           =   2175
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   8880
      Shape           =   4  'Rounded Rectangle
      Top             =   5400
      Width           =   2175
   End
End
Attribute VB_Name = "medicine1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = AutoMatchCBBox(Combo1, KeyAscii)
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
If rmedicine.State = adStateOpen Then
    rmedicine.Close
End If
rmedicine.Open "select * from [Medicine]", conn, adOpenDynamic, adLockPessimistic
Set DataGrid1.DataSource = rmedicine
DataGrid1.Refresh
Frame1.Visible = False
End Sub

Private Sub Label16_Click()
Frame1.Visible = True
End Sub

Private Sub Label17_Click()
medicine2.Top = 0
medicine2.Left = 0
medicine2.Show
Unload Me
End Sub

Private Sub Label3_Click()
If rmedicine.State = adStateOpen Then
rmedicine.Close
End If
rmedicine.Open "select * from [Medicine] where(M_name='" & Text1.Text & "')", conn, adOpenDynamic, adLockPessimistic
If rmedicine.RecordCount <> 0 Then
    MsgBox "MEDICINE ALREADY EXISTS"
    Set DataGrid1.DataSource = rmedicine
    DataGrid1.Refresh
    DataGrid1.Visible = True
Else
    DataGrid1.Visible = False
    Frame1.Visible = True
End If
End Sub

Private Sub Label7_Click()
mainform.Left = 0
mainform.Top = 0
mainform.Show
Unload Me
End Sub

Private Sub Label9_Click()
If Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Combo1.Text = "" Then
    MsgBox "Some of your fields are incomplete. Please complete it", vbCritical, "Warning"
Else
    If rmedicine.State = adStateOpen Then
    rmedicine.Close
    End If
    rmedicine.Open "select * from [Medicine]", conn, adOpenDynamic, adLockPessimistic
    rmedicine.AddNew
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
    If rmedicine.State = adStateOpen Then
    rmedicine.Close
End If
rmedicine.Open "select * from [Medicine]", conn, adOpenDynamic, adLockPessimistic
Set DataGrid1.DataSource = rmedicine
DataGrid1.Refresh
    Frame1.Visible = False
End If

End Sub

Private Sub Text2_LostFocus()
Call validateName(Text2.Text)
If rmedicine.State = adStateOpen Then
rmedicine.Close
End If
rmedicine.Open "select * from medicine where M_name='" & Text2.Text & "'"
If rmedicine.RecordCount <> 0 Then
    MsgBox "Medicine with that name exists!!", vbInformation
End If

End Sub

Private Sub Text3_LostFocus()
Call validateName(Text3.Text)
End Sub

Private Sub Text9_LostFocus()
Text8.Text = Val(Text7.Text) + (Val(Text7.Text) * (Val(Text9.Text) / 100))
End Sub


