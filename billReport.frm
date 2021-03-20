VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form billReport 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   22605
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   12195
   ScaleWidth      =   22605
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   3975
      Left            =   5400
      TabIndex        =   5
      Top             =   4920
      Visible         =   0   'False
      Width           =   13215
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
         Left            =   3720
         TabIndex        =   8
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Doctor Name"
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
         Left            =   720
         TabIndex        =   9
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label3 
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
         Left            =   9240
         TabIndex        =   7
         Top             =   240
         Width           =   1815
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Index           =   2
         Left            =   9240
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   12735
      Left            =   0
      Picture         =   "billReport.frx":0000
      ScaleHeight     =   12675
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   3975
      Left            =   5400
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   13215
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000007&
         Height          =   375
         Left            =   9720
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H80000007&
         Height          =   375
         Left            =   6000
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   375
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
         Left            =   7320
         TabIndex        =   15
         Top             =   360
         Width           =   2295
      End
      Begin VB.TextBox Text1 
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
         Left            =   3600
         TabIndex        =   14
         Top             =   360
         Width           =   2295
      End
      Begin MSACAL.Calendar Calendar1 
         Height          =   3015
         Left            =   1920
         TabIndex        =   10
         Top             =   960
         Visible         =   0   'False
         Width           =   4455
         _Version        =   524288
         _ExtentX        =   7858
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
      Begin MSACAL.Calendar Calendar2 
         Height          =   3015
         Left            =   7320
         TabIndex        =   11
         Top             =   960
         Visible         =   0   'False
         Width           =   4455
         _Version        =   524288
         _ExtentX        =   7858
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
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SHOW"
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
         Left            =   10560
         TabIndex        =   19
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label9 
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
         Left            =   5880
         TabIndex        =   18
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "to"
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
         Left            =   6720
         TabIndex        =   13
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Date        from:  "
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
         Left            =   600
         TabIndex        =   12
         Top             =   360
         Width           =   2895
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Index           =   3
         Left            =   5880
         Shape           =   4  'Rounded Rectangle
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Index           =   4
         Left            =   10560
         Shape           =   4  'Rounded Rectangle
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Label Label12 
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
      Left            =   11040
      TabIndex        =   20
      Top             =   5160
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "                     DOCTOR-WISE   BILL  REPORT"
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
      Height          =   1275
      Left            =   15240
      TabIndex        =   4
      Top             =   2880
      Width           =   1995
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  DATE-WISE           BILL        REPORT"
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
      Height          =   1155
      Left            =   11040
      TabIndex        =   3
      Top             =   3120
      Width           =   1995
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "        ALL                BILL         REPORT"
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
      Height          =   1035
      Left            =   6720
      TabIndex        =   2
      Top             =   3120
      Width           =   1995
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "BILL Reports..."
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
      TabIndex        =   1
      Top             =   120
      Width           =   6135
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
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1935
      Index           =   1
      Left            =   6720
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   1995
   End
   Begin VB.Shape Shape5 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1935
      Left            =   11040
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   1995
   End
   Begin VB.Shape Shape6 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   1935
      Left            =   15240
      Shape           =   4  'Rounded Rectangle
      Top             =   2640
      Width           =   1995
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   5
      Left            =   11040
      Shape           =   4  'Rounded Rectangle
      Top             =   5040
      Width           =   1935
   End
End
Attribute VB_Name = "billReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Calendar1_Click()
Text1.Text = Format(Calendar1.Value, "yyyy-mm-dd")
Calendar1.Visible = False
End Sub

Private Sub Calendar2_Click()
Text2.Text = Format(Calendar2.Value, "yyyy-mm-dd")
Calendar2.Visible = False
End Sub

Private Sub Combo1_Click()
'If rbill_1.State = adStateOpen Then
 '   rbill_1.Close
'End If
'rbill_1.Open "select * from [Bill1] where(D_regno=(select D_regno from [Doctor]where(D_name='" & Combo1.Text & "')))", conn, adOpenDynamic, adLockPessimistic
'Set DataGrid2.DataSource = rbill_1
DataEnvironment1.Command6 Combo1.Text
DataReport6.Refresh
DataReport6.Sections("header").Controls("lblname").Caption = "DOCTOR: " & Combo1.Text
DataReport6.Refresh
DataReport6.Show
DataEnvironment1.rsCommand6.Close

End Sub

Private Sub Command1_Click()
Calendar1.Visible = True
End Sub

Private Sub Command2_Click()
Calendar2.Visible = True
End Sub

Private Sub Label10_Click()
If Text1.Text = "" Then
    Text1.Text = Format(Now, "yyyy-mm-dd")
    Text2.Text = Format(Now, "yyyy-mm-dd")
ElseIf Text2.Text = "" Then
    Text2.Text = Format(Now, "yyyy-mm-dd")
End If
'If Text1.Text <> "" And Text2.Text <> "" Then
 '   If rbill_1.State = adStateOpen Then
 '   rbill_1.Close
 '   End If
'End If
'rbill_1.Open "select * from [BIll1] where B_date between #" & Format(Text1.Text, "yyyy/mm/dd") & "# and #" & Format(Text2.Text, "yyyy/mm/dd") & "#", conn, adOpenDynamic, adLockPessimistic
'rbill_1.Open "SELECT * FROM [bill1] WHERE NOT (From_date > @RangeTill OR To_date < @RangeFrom)"
'Set DataGrid3.DataSource = rbill_1
'DataGrid3.Visible = True
DataEnvironment1.Command7 Text1.Text, Text2.Text
DataReport7.Refresh
DataReport7.Show
DataEnvironment1.rsCommand7.Close
End Sub

Private Sub Label12_Click()
report.Top = 0
report.Left = 0
report.Show
Unload Me
End Sub

Private Sub Label2_Click()
Frame1.Visible = False
End Sub

Private Sub Label3_Click()
Frame2.Visible = False
End Sub

Private Sub Label5_Click()
'If rbill_1.State = adStateOpen Then
'    rbill_1.Close
'End If
'rbill_1.Open "Select * from [BIll1]", conn, adOpenDynamic, adLockPessimistic
'Set DataGrid1.DataSource = rbill_1
'DataGrid1.Refresh
'Frame1.Visible = True
DataReport8.Show
End Sub

Private Sub Label6_Click()
Frame3.Visible = True
End Sub

Private Sub Label7_Click()
If rdoctor.State = adStateOpen Then
    rdoctor.Close
End If
rdoctor.Open "select D_name from [Doctor]", conn, adOpenDynamic, adLockPessimistic
Call populate(rdoctor, Combo1, "D_name")
Frame2.Visible = True
End Sub

Private Sub Label9_Click()
Frame3.Visible = False
End Sub
