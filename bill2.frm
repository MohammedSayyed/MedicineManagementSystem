VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form bill2 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   23175
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   12765
   ScaleWidth      =   23175
   ShowInTaskbar   =   0   'False
   Begin MSACAL.Calendar Calendar2 
      Height          =   3015
      Left            =   11520
      TabIndex        =   5
      Top             =   2760
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
   Begin MSACAL.Calendar Calendar1 
      Height          =   3015
      Left            =   6120
      TabIndex        =   4
      Top             =   2760
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
   Begin VB.PictureBox Picture1 
      Height          =   12735
      Left            =   0
      Picture         =   "bill2.frx":0000
      ScaleHeight     =   12675
      ScaleWidth      =   4395
      TabIndex        =   12
      Top             =   0
      Width           =   4455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "UPDATE batch"
      Height          =   4455
      Left            =   5280
      TabIndex        =   9
      Top             =   3240
      Visible         =   0   'False
      Width           =   13215
      Begin MSDataGridLib.DataGrid DataGrid1 
         CausesValidation=   0   'False
         Height          =   2955
         Left            =   0
         Negotiate       =   -1  'True
         TabIndex        =   11
         Top             =   0
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
      Begin VB.Label Label5 
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
         Left            =   4440
         TabIndex        =   15
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "VIEW"
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
         Left            =   6840
         TabIndex        =   10
         Top             =   3600
         Width           =   1695
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Index           =   0
         Left            =   6720
         Shape           =   4  'Rounded Rectangle
         Top             =   3480
         Width           =   1935
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Index           =   2
         Left            =   4320
         Shape           =   4  'Rounded Rectangle
         Top             =   3480
         Width           =   1935
      End
   End
   Begin VB.TextBox Text1 
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
      Left            =   7800
      TabIndex        =   3
      Top             =   2160
      Width           =   2295
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
      Left            =   11520
      TabIndex        =   2
      Top             =   2160
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000007&
      Height          =   375
      Left            =   10200
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H80000007&
      Height          =   375
      Left            =   13920
      MaskColor       =   &H00000000&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2160
      Width           =   375
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
      Left            =   16800
      TabIndex        =   14
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000D&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   4920
      Top             =   720
      Width           =   18015
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "BILL VIEW.. ."
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
      Height          =   495
      Left            =   4920
      TabIndex        =   13
      Top             =   0
      Width           =   10215
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
      Left            =   4800
      TabIndex        =   8
      Top             =   2160
      Width           =   2895
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
      Left            =   10920
      TabIndex        =   7
      Top             =   2160
      Width           =   375
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
      Left            =   14760
      TabIndex        =   6
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   4
      Left            =   14760
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   1
      Left            =   16800
      Shape           =   4  'Rounded Rectangle
      Top             =   2160
      Width           =   1935
   End
End
Attribute VB_Name = "bill2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Calendar1.Visible = True
End Sub

Private Sub Command2_Click()
Calendar2.Visible = True
End Sub

Private Sub Calendar1_Click()
Text1.Text = Calendar1.Value
Calendar1.Visible = False
End Sub

Private Sub Calendar2_Click()
Text2.Text = Calendar2.Value
Calendar2.Visible = False
End Sub

Private Sub DataGrid1_Click()
billID = Val(DataGrid1.Columns(0))
End Sub

'Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
'billID = Val(DataGrid1.Columns(0))
'End Sub

Private Sub Label1_Click()
'rbill_1.Close
'DataGrid1.Enabled = False
billPrint.Top = 0
billPrint.Left = 0
billPrint.Show
Unload Me
End Sub

Private Sub Label10_Click()
If Text1.Text = "" Then
    Text1.Text = Format(Now, "dd-mm-yyyy")
    Text2.Text = Format(Now, "dd-mm-yyyy")
ElseIf Text2.Text = "" Then
    Text2.Text = Format(Now, "dd-mm-yyyy")
End If
If Text1.Text <> "" And Text2.Text <> "" Then
    If rbill_1.State = adStateOpen Then
    rbill_1.Close
    End If
    rbill_1.Open "select B_id,Doctor.D_name,B_amt,E_id,B_date from Bill1 INNER JOIN Doctor on Bill1.D_regno=Doctor.D_regno where B_date between #" & Format(Text1.Text, "yyyy/mm/dd") & "# and #" & Format(Text2.Text, "yyyy/mm/dd") & "#", conn, adOpenKeyset, adLockPessimistic
    'rbill_1.Open "SELECT * FROM [bill1] WHERE NOT (From_date > @RangeTill OR To_date < @RangeFrom)"
    If rbill_1.RecordCount <> 0 Then
        Set DataGrid1.DataSource = rbill_1
        billID = Val(DataGrid1.Columns(0))
        Frame1.Visible = True
    Else
        Frame1.Visible = False
    MsgBox "NO BILLS between these Dates!!", vbInformation
    End If
End If
End Sub

Private Sub Label3_Click()
bill1.Top = 0
bill1.Left = 0
bill1.Show
Unload Me
End Sub

Private Sub Label5_Click()
bid = billID
If rbill_2.State = adStateOpen Then
    rbill_2.Close
End If
rbill_2.Open "select * from bill2 where B_ID=" & bid, conn, adOpenDynamic, adLockPessimistic
Set bill1.DataGrid1.DataSource = rbill_2
If rdoctor.State = adStateOpen Then
    rdoctor.Close
End If
rdoctor.Open "select D_name from Doctor where D_regno=(select D_regno from bill1 where B_ID=" & bid & ")", conn, adOpenDynamic, adLockPessimistic
bill1.Combo1.Text = rdoctor!D_name
bill1.Combo1.Enabled = False
bill1.Combo2.Enabled = True
bill1.Shape3.Visible = False
bill1.Label7.Visible = False
bill1.Label12.Visible = True
bill1.Label12.Enabled = True
bill1.Shape7.Visible = True
bill1.Label1.Enabled = True
bill1.Label11.Enabled = True
bill1.Top = 0
bill1.Left = 0
bill1.Show
Unload Me
End Sub
