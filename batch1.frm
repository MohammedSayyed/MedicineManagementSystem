VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form batch1 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "batch1"
   ClientHeight    =   14985
   ClientLeft      =   60
   ClientTop       =   0
   ClientWidth     =   22080
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   14985
   ScaleMode       =   0  'User
   ScaleWidth      =   22080
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "UPDATE batch"
      Height          =   8895
      Left            =   4800
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   17535
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
         Left            =   1440
         TabIndex        =   26
         Top             =   2040
         Width           =   4695
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000007&
         Caption         =   "Command2"
         Height          =   375
         Left            =   12840
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3840
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000007&
         Caption         =   "Command2"
         Height          =   375
         Left            =   12840
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1920
         Width           =   375
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
         TabIndex        =   7
         Top             =   3840
         Width           =   4695
      End
      Begin VB.TextBox Text4 
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
      Begin MSACAL.Calendar Calendar1 
         Height          =   2895
         Left            =   8400
         TabIndex        =   15
         Top             =   4440
         Visible         =   0   'False
         Width           =   4335
         _Version        =   524288
         _ExtentX        =   7646
         _ExtentY        =   5106
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
      Begin MSACAL.Calendar Calendar2 
         Height          =   2895
         Left            =   8400
         TabIndex        =   16
         Top             =   4440
         Visible         =   0   'False
         Width           =   4335
         _Version        =   524288
         _ExtentX        =   7646
         _ExtentY        =   5106
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
         Left            =   8040
         TabIndex        =   27
         Top             =   5880
         Width           =   4695
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
         Left            =   8640
         TabIndex        =   21
         Top             =   7800
         Width           =   1815
      End
      Begin VB.Label Label7 
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
         Height          =   375
         Left            =   4560
         TabIndex        =   14
         Top             =   7800
         Width           =   1695
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
         TabIndex        =   13
         Top             =   5280
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
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   1320
         Width           =   2655
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
         TabIndex        =   10
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Medicine ID"
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
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   4440
         Shape           =   4  'Rounded Rectangle
         Top             =   7680
         Width           =   1935
      End
      Begin VB.Shape Shape5 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   8640
         Shape           =   4  'Rounded Rectangle
         Top             =   7680
         Width           =   1935
      End
      Begin VB.Label Label15 
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
         Left            =   8040
         TabIndex        =   28
         Top             =   5280
         Width           =   2895
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   15615
      Left            =   0
      Picture         =   "batch1.frx":0000
      ScaleHeight     =   15555
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      CausesValidation=   0   'False
      Height          =   2955
      Left            =   5520
      Negotiate       =   -1  'True
      TabIndex        =   20
      Top             =   2640
      Visible         =   0   'False
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   5212
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
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      TabIndex        =   22
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Label Label14 
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
      Left            =   13830
      TabIndex        =   25
      Top             =   6360
      Visible         =   0   'False
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
      Height          =   495
      Left            =   11430
      TabIndex        =   24
      Top             =   6360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CREATE NEW"
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
      Left            =   9030
      TabIndex        =   23
      Top             =   6360
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label13 
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
      Left            =   16680
      TabIndex        =   19
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label1 
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
      Left            =   14400
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Select Medicine name"
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
      TabIndex        =   2
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808000&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   4800
      Top             =   840
      Width           =   18015
   End
   Begin VB.Label Label10 
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
      Left            =   4800
      TabIndex        =   1
      Top             =   0
      Width           =   10215
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   14400
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   16680
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   9030
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Shape Shape7 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   11430
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Shape Shape8 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   13830
      Shape           =   4  'Rounded Rectangle
      Top             =   6240
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "batch1"
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
    Text4.Text = Calendar1.Value
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
If rmedicine.State = adStateOpen Then
    rmedicine.Close
End If
rmedicine.Open "select * from [Medicine] where(M_name='" & Combo1.Text & "')", conn, adOpenDynamic, adLockPessimistic
If rmedicine.RecordCount = 1 Then
    If rbatch.State = adStateOpen Then
        rbatch.Close
    End If
    rbatch.Open "select * from [Batch] where(M_name= '" & Combo1.Text & "' ) ", conn, adOpenDynamic, adLockPessimistic
    If rbatch.RecordCount <> 0 Then
        Set DataGrid1.DataSource = rbatch
        DataGrid1.Refresh
        DataGrid1.Visible = True
        Shape6.Visible = True
        Label8.Visible = True
        Label12.Visible = True
        Label14.Visible = True
        Shape7.Visible = True
        Shape8.Visible = True
    Else
        Frame1.Top = 1080
        Frame1.Left = 4680

        Frame1.Visible = True
        Label11.Visible = True
        Shape5.Visible = True
        Label12.Visible = False
        Label14.Visible = False
        Shape7.Visible = False
        Shape8.Visible = False
        Text2.Text = Combo1.Text
        
    End If
Else
    MsgBox "NO SUCH MEDICINE EXISTS!!, PLEASE ENTER A VALID MEDICINE NAME!!", vbCritical
End If

End Sub

Private Sub Combo2_Change()

End Sub

Private Sub Command1_Click()
Calendar2.Visible = True
End Sub

Private Sub Command2_Click()
Calendar1.Visible = True

End Sub

Private Sub Form_Load()
If rmedicine.State = adStateOpen Then
    rmedicine.Close
End If
rmedicine.Open "select M_name from [Medicine]", conn, adOpenDynamic, adLockPessimistic
Call populate(rmedicine, Combo1, "M_name")
If rsupplier.State = adStateOpen Then
    rsupplier.Close
End If
rsupplier.Open "select S_name from supplier", conn, adOpenDynamic, adLockPessimistic
Call populate(rsupplier, Combo3, "S_name")
End Sub

Private Sub Label1_Click()
DataGrid1.Visible = False
Shape6.Visible = False
Label8.Visible = False

        
If Combo1.Text = "" Then
    MsgBox "Please enter a medicine name!!", vbCritical, "ERROR"
Else
If rmedicine.State = adStateOpen Then
    rmedicine.Close
End If
rmedicine.Open "select * from [Medicine] where(M_name='" & Combo1.Text & "')", conn, adOpenDynamic, adLockPessimistic
If rmedicine.RecordCount = 1 Then
    If rbatch.State = adStateOpen Then
        rbatch.Close
    End If
    rbatch.Open "select * from [Batch] where(M_name= '" & Combo1.Text & "' ) ", conn, adOpenDynamic, adLockPessimistic
    If rbatch.RecordCount <> 0 Then
        Set DataGrid1.DataSource = rbatch
        DataGrid1.Refresh
        DataGrid1.Visible = True
        Shape6.Visible = True
        Label8.Visible = True
        Label12.Visible = True
        Label14.Visible = True
        Shape7.Visible = True
        Shape8.Visible = True
    Else
        Frame1.Top = 1080
        Frame1.Left = 4680

        Frame1.Visible = True
        Text2.Text = Combo1.Text
    End If
Else
    MsgBox "NO SUCH MEDICINE EXISTS!!, PLEASE ENTER A VALID MEDICINE NAME!!", vbCritical
End If
End If






End Sub

Private Sub Label11_Click()
Frame1.Visible = False
        DataGrid1.Visible = False
        Shape6.Visible = False
        Label8.Visible = False
        Label12.Visible = False
        Label14.Visible = False
        Shape7.Visible = False
        Shape8.Visible = False
End Sub

Private Sub Label12_Click()
Unload Me
batch2.Top = 0
batch2.Left = 0
batch2.Show
End Sub

Private Sub Label13_Click()
mainform.Left = 0
mainform.Top = 0
mainform.Show
Unload Me
End Sub

Private Sub Label14_Click()
Unload Me
batch2.Top = 0
batch2.Left = 0
batch2.Show
End Sub

Private Sub Label7_Click()
If Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Then
    MsgBox "Some of your fields are incomplete. Please complete it", vbCritical, "Warning"
Else
    If rbatch.State = adStateOpen Then
    rbatch.Close
    End If
    rbatch.Open "select * from [Batch]", conn, adOpenDynamic, adLockPessimistic
    rbatch.AddNew
    rbatch!B_no = Text3.Text
    rbatch!B_exp_date = Text4.Text
    rbatch!B_mfg_date = Text5.Text
    rbatch!B_qty = Val(Text6.Text)
    rbatch!M_name = Text2.Text
    rbatch!B_sup = Combo3.Text
    rbatch.Update
    Set DataGrid1.DataSource = rbatch
    DataGrid1.Refresh
    MsgBox "SAVED SUCCESSFULLY...", vbInformation, "INFORMATION"
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text5.Text = ""
    Text6.Text = ""
    Frame1.Visible = False
    Label12.Visible = False
    Label14.Visible = False
    Shape7.Visible = False
    Shape8.Visible = False
    DataGrid1.Visible = False
    Label8.Visible = False
    Shape6.Visible = False
End If
End Sub

Private Sub Label8_Click()
Frame1.Top = 1080
Frame1.Left = 4680

Frame1.Visible = True
Text2.Text = Combo1.Text
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



Private Sub Text3_LostFocus()
If Text3.Text = "" Then
    MsgBox "Please enter a BATCH number", vbInformation
Else
If rbatch.State = adStateOpen Then
    rbatch.Close
End If
rbatch.Open "select * from [Batch] where(B_no='" & Text3.Text & "')", conn, adOpenDynamic, adLockPessimistic
If rbatch.RecordCount = 1 Then
    MsgBox "BATCH NUMBER ALREADY EXISTS!!, PLEASE ENTER A NEW BATCH NUMBER", vbCritical, "INFO"
End If
rbatch.Close
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
