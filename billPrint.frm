VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form billPrint 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   15285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   16365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   15285
   ScaleWidth      =   16365
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox Discount 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13920
      TabIndex        =   1
      Top             =   9000
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   5055
      Left            =   480
      LinkItem        =   "&H80000005&"
      TabIndex        =   5
      Top             =   3120
      Width           =   255
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      CausesValidation=   0   'False
      Height          =   5055
      Left            =   480
      Negotiate       =   -1  'True
      TabIndex        =   6
      Top             =   3120
      Visible         =   0   'False
      Width           =   15435
      _ExtentX        =   27226
      _ExtentY        =   8916
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   26
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
         Name            =   "Yu Gothic Medium"
         Size            =   12
         Charset         =   0
         Weight          =   500
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PRINT"
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
      Left            =   6000
      TabIndex        =   20
      Top             =   12600
      Width           =   1695
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
      Height          =   375
      Left            =   8880
      TabIndex        =   19
      Top             =   12600
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Authorized Signatory"
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
      Left            =   11400
      TabIndex        =   18
      Top             =   11400
      Width           =   2775
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "For MJ MEDICAL DISTRIBUTORS"
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
      Left            =   11400
      TabIndex        =   17
      Top             =   10320
      Width           =   4455
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Email:joelsasane211@gmail.com                                    mohammedsayyed1010@gmail.com"
      BeginProperty Font 
         Name            =   "Yu Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   720
      TabIndex        =   16
      Top             =   10440
      Width           =   6135
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   4
      X1              =   0
      X2              =   16320
      Y1              =   10080
      Y2              =   10080
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   5
      X1              =   0
      X2              =   16320
      Y1              =   8280
      Y2              =   8280
   End
   Begin VB.Label Total 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Left            =   13800
      TabIndex        =   2
      Top             =   9480
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Net:"
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
      Left            =   13080
      TabIndex        =   15
      Top             =   9480
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "DISCOUNT AMT:"
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
      Left            =   11400
      TabIndex        =   14
      Top             =   9000
      Width           =   2295
   End
   Begin VB.Label amount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      Left            =   13800
      TabIndex        =   0
      Top             =   8400
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Gross:"
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
      Left            =   12840
      TabIndex        =   13
      Top             =   8400
      Width           =   855
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   5
      X1              =   0
      X2              =   16320
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00000000&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   5
      X1              =   0
      X2              =   16320
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label doctorREG 
      BackStyle       =   0  'Transparent
      Caption         =   "REG-NO: "
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
      Left            =   11160
      TabIndex        =   12
      Top             =   2280
      Width           =   5175
   End
   Begin VB.Label doctorMobile 
      BackStyle       =   0  'Transparent
      Caption         =   "MOB: "
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
      Left            =   11160
      TabIndex        =   11
      Top             =   1800
      Width           =   5175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      DrawMode        =   1  'Blackness
      X1              =   0
      X2              =   16320
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderStyle     =   5  'Dash-Dot-Dot
      BorderWidth     =   3
      DrawMode        =   7  'Invert
      X1              =   8160
      X2              =   8160
      Y1              =   120
      Y2              =   2880
   End
   Begin VB.Label doctorAddress 
      BackStyle       =   0  'Transparent
      Caption         =   " "
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
      Left            =   1320
      TabIndex        =   10
      Top             =   2160
      Width           =   6735
   End
   Begin VB.Label DoctorName 
      BackStyle       =   0  'Transparent
      Caption         =   "To, "
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
      Left            =   720
      TabIndex        =   9
      Top             =   1800
      Width           =   7335
   End
   Begin VB.Label billDate 
      BackStyle       =   0  'Transparent
      Caption         =   "BILL-DATE: "
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
      Left            =   11160
      TabIndex        =   8
      Top             =   840
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "MOB:9860264614,9579049848"
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
      TabIndex        =   7
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "MJ MEDICAL DISTRIBUTORS"
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
      TabIndex        =   4
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label billN0 
      BackStyle       =   0  'Transparent
      Caption         =   "BILL .No.: "
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
      Left            =   11160
      TabIndex        =   3
      Top             =   360
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   1
      Left            =   8760
      Shape           =   4  'Rounded Rectangle
      Top             =   12480
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   615
      Index           =   0
      Left            =   5880
      Shape           =   4  'Rounded Rectangle
      Top             =   12480
      Width           =   1935
   End
End
Attribute VB_Name = "billPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Discount_Change()
If Val(Discount) <= amount Then
    Total = amount - Val(Discount.Text)
Else
    MsgBox "Discount amount is greater than the Total amount!!", vbInformation
    Discount = Val(Discount) \ 10
End If
End Sub



Private Sub Form_Load()
If rmedicine.State = adStateOpen Then
    rmedicine.Close
End If

If rbill_1.State = adStateOpen Then
    rbill_1.Close
End If

If rbill_2.State = adStateOpen Then
    rbill_2.Close
End If
rbill_1.Open "select * from bill1 where B_id=" & billID & "", conn, adOpenDynamic, adLockPessimistic
rbill_2.Open "select BILL2.M_name as PRODUCT ,Medicine.M_type As TYPE,Batch_no AS BATCH,Batch_exp as EXP,Medicine.M_packing AS PACKING,Bill2.M_qty as QTY ,Medicine.M_mrp AS MRP ,Medicine.M_sp as RATE,M_price as AMOUNT from Bill2 INNER JOIN Medicine ON Bill2.M_name = Medicine.M_name where B_id=" & billID, conn, adOpenDynamic, adLockPessimistic
Set DataGrid1.DataSource = rbill_2
DataGrid1.Visible = True
If rbatch.State = adStateOpen Then
    rbatch.Close
End If
If rdoctor.State = adStateOpen Then
    rdoctor.Close
End If
rdoctor.Open "select D_name,D_address,D_phno from Doctor where D_regno='" & rbill_1!D_regno & "'", conn, adOpenDynamic, adLockPessimistic
DoctorName.Caption = DoctorName.Caption + rdoctor!D_name
doctorAddress.Caption = rdoctor!D_address
doctorMobile.Caption = doctorMobile.Caption + rdoctor!D_phno
doctorREG.Caption = doctorREG.Caption + rbill_1!D_regno
amount = rbill_1!B_Amt
billN0.Caption = billN0.Caption + str(billID)
billDate.Caption = billDate.Caption + str(rbill_1!B_date)
With DataGrid1
    .Columns(0).Width = 2000
    .Columns(1).Width = 1500
    .Columns(2).Width = 1500
    .Columns(3).Width = 2370
    .Columns(4).Width = 1500
    .Columns(5).Width = 1500
    .Columns(6).Width = 1570
    .Columns(7).Width = 1280
    .Columns(8).Width = 1400
End With
DataGrid1.Enabled = False
End Sub
'rbill_1.Open "SELECT B_id,Doctor.D_name,B_amt,E_id,B_date From Bill1 INNER JOIN Doctor ON Bill1.D_regno=Doctor.D_regno where B_date=#" & Format(Text1.Text, "yyyy/mm/dd") & "#", conn, adOpenDynamic, adLockPessimistic

Private Sub Label1_Click()
Label1.Visible = False
Label10.Visible = False
Shape1(0).Visible = False
Shape1(1).Visible = False

PrintForm
Label1.Visible = True
Label10.Visible = True
Shape1(0).Visible = True
Shape1(1).Visible = True

End Sub

Private Sub Label10_Click()
bill2.Top = 0
bill2.Left = 0
bill2.Show
Unload Me
End Sub
