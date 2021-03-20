VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form batch2 
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "batch2"
   ClientHeight    =   14370
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   23040
   ControlBox      =   0   'False
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   718.5
   ScaleMode       =   0  'User
   ScaleWidth      =   1152
   ShowInTaskbar   =   0   'False
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
      Left            =   9720
      TabIndex        =   23
      Top             =   1560
      Width           =   3615
   End
   Begin VB.PictureBox Picture1 
      Height          =   15615
      Left            =   0
      Picture         =   "batch2.frx":0000
      ScaleHeight     =   15555
      ScaleWidth      =   4395
      TabIndex        =   12
      Top             =   0
      Width           =   4455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "UPDATE batch"
      Height          =   8535
      Left            =   5400
      TabIndex        =   0
      Top             =   5160
      Visible         =   0   'False
      Width           =   17535
      Begin MSACAL.Calendar Calendar2 
         Height          =   2775
         Left            =   8040
         TabIndex        =   19
         Top             =   4440
         Visible         =   0   'False
         Width           =   5535
         _Version        =   524288
         _ExtentX        =   9763
         _ExtentY        =   4895
         _StockProps     =   1
         BackColor       =   -2147483633
         Year            =   2019
         Month           =   8
         Day             =   30
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
      Begin MSACAL.Calendar Calendar1 
         Height          =   2775
         Left            =   8040
         TabIndex        =   18
         Top             =   2400
         Visible         =   0   'False
         Width           =   5535
         _Version        =   524288
         _ExtentX        =   9763
         _ExtentY        =   4895
         _StockProps     =   1
         BackColor       =   16777215
         Year            =   2019
         Month           =   8
         Day             =   30
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
      Begin VB.ComboBox Combo3 
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
         Left            =   8160
         TabIndex        =   25
         Top             =   5880
         Width           =   4695
      End
      Begin VB.ComboBox Combo2 
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
         Left            =   1560
         TabIndex        =   24
         Top             =   1920
         Width           =   4695
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000007&
         Caption         =   "Command2"
         Height          =   375
         Left            =   12840
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1920
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000007&
         Caption         =   "Command2"
         Height          =   375
         Left            =   12840
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3840
         Width           =   375
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
         TabIndex        =   4
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
         TabIndex        =   3
         Top             =   1920
         Width           =   4695
      End
      Begin VB.TextBox Text5 
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
         Left            =   8040
         TabIndex        =   2
         Top             =   3840
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
         TabIndex        =   1
         Top             =   5880
         Width           =   4695
      End
      Begin VB.Label Label12 
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
         Left            =   8160
         TabIndex        =   26
         Top             =   5280
         Width           =   2895
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
         Height          =   375
         Left            =   10680
         TabIndex        =   22
         Top             =   7680
         Width           =   1695
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DELETE"
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
         Left            =   6000
         TabIndex        =   14
         Top             =   7680
         Width           =   1815
      End
      Begin VB.Label Label2 
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
         Left            =   1560
         TabIndex        =   10
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Batch ID"
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
         TabIndex        =   9
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Expiry Date"
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
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Manufacturing Date"
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
         TabIndex        =   7
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Medicine Quantity"
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
         TabIndex        =   6
         Top             =   5280
         Width           =   2655
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
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   7680
         Width           =   1815
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   1680
         Shape           =   4  'Rounded Rectangle
         Top             =   7560
         Width           =   1935
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   5880
         Shape           =   4  'Rounded Rectangle
         Top             =   7560
         Width           =   1935
      End
      Begin VB.Shape Shape5 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   10560
         Shape           =   4  'Rounded Rectangle
         Top             =   7560
         Width           =   1935
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      CausesValidation=   0   'False
      Height          =   2355
      Left            =   5640
      Negotiate       =   -1  'True
      TabIndex        =   21
      Top             =   2640
      Visible         =   0   'False
      Width           =   14055
      _ExtentX        =   24791
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
   Begin VB.Label Label8 
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
      Height          =   375
      Left            =   17760
      TabIndex        =   20
      Top             =   1560
      Width           =   1695
   End
   Begin VB.Label Label10 
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
      TabIndex        =   15
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label100 
      BackStyle       =   0  'Transparent
      Caption         =   "MEDICINE BATCH INFO.."
      BeginProperty Font 
         Name            =   "Mongolian Baiti"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   615
      Left            =   4680
      TabIndex        =   13
      Top             =   120
      Width           =   10215
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   4680
      Top             =   960
      Width           =   18015
   End
   Begin VB.Label Label9 
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
      Height          =   975
      Left            =   5640
      TabIndex        =   11
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   15000
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   17640
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   1935
   End
End
Attribute VB_Name = "batch2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Calendar1_Click()
If Format(Calendar1.Value, "yyyy/mm/dd") <= Format(Now, "yyyy/mm/dd") Then
    MsgBox "EXPIRY DATE IS LESS THAN CURRENT DATE!!!, PLEASE ENTER A VALID EXPIRY DATE", vbCritical, "ERROR"
    Text4.Text = ""
    Calendar1.Visible = True
Else
    Text4.Text = Format(Calendar1.Value, "dd-mm-yyyy")
    Calendar1.Visible = False
End If
End Sub

Private Sub Calendar2_Click()
If Format(Calendar2.Value, "yyyy/mm/dd") > Format(Now, "yyyy/mm/dd") Then
    MsgBox "MFG DATE IS GREATER THAN CURRENT DATE!!!, PLEASE ENTER A VALID MFG DATE", vbCritical, "ERROR"
    Text5.Text = ""
    Calendar2.Visible = True
Else
    Text5.Text = Calendar2.Value
    Calendar2.Visible = False
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = AutoMatchCBBox(Combo1, KeyAscii)
End Sub

Private Sub Combo1_LostFocus()
If Combo1.Text = "" Then
MsgBox "Please enter medicine name!!", vbCritical, "INFO"
Else
    If rbatch1.State = adStateOpen Then
        rbatch1.Close
    End If
    
    If rmedicine.State = adStateOpen Then
        rmedicine.Close
    End If
    
    rbatch1.Open "select * from [Batch] where(M_name='" & Combo1.Text & "')", conn, adOpenDynamic, adLockPessimistic
    rmedicine.Open "select * from [Medicine] where(M_name='" & Combo1.Text & "')", conn, adOpenDynamic, adLockPessimistic
    
    If rmedicine.RecordCount = 1 And rbatch1.RecordCount <> 0 Then
        Set DataGrid1.DataSource = rbatch1
        DataGrid1.Visible = True
        DataGrid1.Enabled = True
    ElseIf rmedicine.RecordCount = 1 And rbatch1.RecordCount = 0 Then
        DataGrid1.Visible = False
        MsgBox "THIS MEDICINE DOES NOT HAS ANY BATCHES!!!", vbInformation, "INFO"
        
    End If
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = AutoMatchCBBox(Combo2, KeyAscii)
End Sub

Private Sub Command1_Click()
Calendar1.Visible = True
End Sub

Private Sub Command2_Click()
Calendar2.Visible = True
End Sub

Private Sub DataGrid1_GotFocus()
Combo2.Text = DataGrid1.Columns(5)
Text3.Text = DataGrid1.Columns(1)
Text4.Text = DataGrid1.Columns(2)
Text5.Text = DataGrid1.Columns(3)
Text6.Text = DataGrid1.Columns(4)
Combo3.Text = DataGrid1.Columns(6)


Frame1.Visible = True
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
Combo2.Text = DataGrid1.Columns(5)
Text3.Text = DataGrid1.Columns(1)
Text4.Text = DataGrid1.Columns(2)
Text5.Text = DataGrid1.Columns(3)
Text6.Text = DataGrid1.Columns(4)
Combo3.Text = DataGrid1.Columns(6)
Frame1.Visible = True
End Sub

Private Sub Form_Load()
If rmedicine.State = adStateOpen Then
    rmedicine.Close
End If
rmedicine.Open "select M_name from [Medicine]", conn, adOpenDynamic, adLockPessimistic
Call populate(rmedicine, Combo1, "M_name")
If rmedicine.State = adStateOpen Then
    rmedicine.Close
End If
rmedicine.Open "select M_name from [Medicine]", conn, adOpenDynamic, adLockPessimistic
Call populate(rmedicine, Combo2, "M_name")
If rsupplier.State = adStateOpen Then
    rsupplier.Close
End If
rsupplier.Open "select S_name from supplier", conn, adOpenDynamic, adLockPessimistic
Call populate(rsupplier, Combo3, "S_name")
End Sub

Private Sub Label1_Click()
Frame1.Visible = False
End Sub

Private Sub Label10_Click()
If Combo1.Text = "" Then
MsgBox "Please enter medicine name!!", vbCritical, "INFO"
Else
    If rbatch1.State = adStateOpen Then
        rbatch1.Close
    End If
    
    If rmedicine.State = adStateOpen Then
        rmedicine.Close
    End If
    
    rbatch1.Open "select * from [Batch] where(M_name='" & Combo1.Text & "')", conn, adOpenDynamic, adLockPessimistic
    rmedicine.Open "select * from [Medicine] where(M_name='" & Combo1.Text & "')", conn, adOpenDynamic, adLockPessimistic
    
    If rmedicine.RecordCount = 1 And rbatch1.RecordCount <> 0 Then
        Set DataGrid1.DataSource = rbatch1
        DataGrid1.Visible = True
        DataGrid1.Enabled = True
    ElseIf rmedicine.RecordCount = 1 And rbatch1.RecordCount = 0 Then
        DataGrid1.Visible = False
        MsgBox "THIS MEDICINE DOES NOT HAS ANY BATCHES!!!", vbInformation, "INFO"
        
    End If
End If


End Sub

Private Sub Label11_Click()
If Combo2.Text = "" Then
MsgBox "No Batch Selscted!", vbCritical, "ALERT"
Else
    If rbatch.State = adStateOpen Then
    rbatch.Close
    End If
    rbatch.Open "delete from [Batch] where(B_no='" & Text3.Text & "')", conn, adOpenDynamic, adLockPessimistic
    MsgBox "DELETED!!", vbInformation
    Combo2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    DataGrid1.Enabled = False
    Frame1.Visible = False
    DataGrid1.Visible = False
    
End If

End Sub

Private Sub Label7_Click()
    rbatch1!M_name = Combo2.Text
    rbatch1!B_no = Text3.Text
    rbatch1!B_exp_date = Text4.Text
    rbatch1!B_mfg_date = Text5.Text
    rbatch1!B_qty = Text6.Text
    rbatch1!B_sup = Combo3.Text
    rbatch1.Update
    'rbatch.Close
    MsgBox "BATCH UPDATED SUCCESSFULLY...", vbInformation, "ALERT"
    DataGrid1.Refresh
    Combo2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Frame1.Visible = False
End Sub

Private Sub Label8_Click()
batch1.Left = 0
batch1.Top = 0
batch1.Show
Unload Me
End Sub

Private Sub Text2_LostFocus()
If Text2.Text = "" Then
    MsgBox "Please enter a medicine name!!", vbInformation
Else
If rmedicine.State = adStateOpen Then
    rmedicine.Close
End If
rmedicine.Open "select * from [Medicine] where(M_name='" & Text2.Text & "')", conn, adOpenDynamic, adLockPessimistic
If rmedicine.RecordCount <> 1 Then
    MsgBox "NO SUCH MEDICINE EXISTS, PLEASE ENTER AN EXISTING MEDICINE OR ADD NEW MEDICINE FIRST!!!", vbCritical, "INFO"
End If
rmedicine.Close
End If
End Sub



Private Sub Text6_LostFocus()
If Text6.Text = "" Then
    MsgBox "PLEASE ENTER QUANTITY", vbInformation
ElseIf num(Text6.Text) = False Then
    MsgBox "Please enter only numbers!", vbInformation
    Text6.Text = ""
End If

End Sub
