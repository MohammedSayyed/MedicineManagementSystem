VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form bill1 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   12495
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   22920
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   12495
   ScaleWidth      =   22920
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      Height          =   12735
      Left            =   -1200
      Picture         =   "bill1.frx":0000
      ScaleHeight     =   12675
      ScaleWidth      =   5835
      TabIndex        =   12
      Top             =   -240
      Width           =   5895
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "UPDATE batch"
      Height          =   7095
      Left            =   5520
      TabIndex        =   4
      Top             =   4680
      Width           =   12975
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   1695
         Left            =   1200
         TabIndex        =   17
         Top             =   2160
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   2990
         _Version        =   393216
         Enabled         =   0   'False
         HeadLines       =   1
         RowHeight       =   27
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
            Name            =   "Arial Narrow"
            Size            =   14.25
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
      Begin VB.ComboBox Combo2 
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
         Height          =   435
         Left            =   8040
         TabIndex        =   1
         Top             =   1080
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
         TabIndex        =   0
         Top             =   840
         Width           =   4695
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
         Left            =   8040
         TabIndex        =   2
         Top             =   3000
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
         Height          =   465
         Left            =   1200
         TabIndex        =   3
         Top             =   4800
         Width           =   4575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "VIEW BILLS"
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
         Left            =   10200
         TabIndex        =   18
         Top             =   6000
         Width           =   1815
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "SAVE"
         Enabled         =   0   'False
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
         Left            =   3960
         TabIndex        =   15
         Top             =   6000
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
         Left            =   9000
         TabIndex        =   6
         Top             =   4920
         Width           =   1695
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   8880
         Shape           =   4  'Rounded Rectangle
         Top             =   4800
         Width           =   1935
      End
      Begin VB.Label Label12 
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
         Left            =   9000
         TabIndex        =   14
         Top             =   4920
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Doctor Name"
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
         TabIndex        =   11
         Top             =   240
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
         Left            =   1200
         TabIndex        =   10
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label4 
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
         Left            =   8040
         TabIndex        =   9
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label5 
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
         Left            =   8040
         TabIndex        =   8
         Top             =   2400
         Width           =   2655
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Price"
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
         Left            =   1200
         TabIndex        =   7
         Top             =   4200
         Width           =   2535
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
         Left            =   7080
         TabIndex        =   5
         Top             =   6000
         Width           =   1815
      End
      Begin VB.Shape Shape5 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   6960
         Shape           =   4  'Rounded Rectangle
         Top             =   5880
         Width           =   1935
      End
      Begin VB.Shape Shape7 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   8880
         Shape           =   4  'Rounded Rectangle
         Top             =   4800
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Shape Shape8 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   3960
         Shape           =   4  'Rounded Rectangle
         Top             =   5880
         Width           =   1935
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00E0E0E0&
         FillStyle       =   0  'Solid
         Height          =   615
         Left            =   10080
         Shape           =   4  'Rounded Rectangle
         Top             =   5880
         Width           =   1935
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      CausesValidation=   0   'False
      Height          =   4035
      Left            =   5520
      Negotiate       =   -1  'True
      TabIndex        =   16
      Top             =   1200
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   7117
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Appearance      =   0
      BorderStyle     =   0
      Enabled         =   0   'False
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
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "BILL ADD NEW. . ."
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
      Left            =   5280
      TabIndex        =   13
      Top             =   120
      Width           =   10215
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H8000000D&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   5400
      Top             =   960
      Width           =   18015
   End
End
Attribute VB_Name = "bill1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tamt As Double
Dim qty As Integer
Dim bno As String
Dim exp As Date

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = AutoMatchCBBox(Combo1, KeyAscii)
End Sub

Private Sub Combo1_LostFocus()
If rdoctor.State = adStateOpen Then
    rdoctor.Close
End If
rdoctor.Open "select * from [Doctor] where(d_name='" & Combo1.Text & "')", conn, adOpenDynamic, adLockPessimistic
If rdoctor.RecordCount <> 1 Then
    MsgBox "PLEASE ENTER A VALID DOCTOR", vbInformation, "INFO"
    Combo1.Text = ""
    Combo1.SetFocus
Else
    Combo2.Enabled = True
    Combo2.SetFocus
End If

End Sub



Private Sub Combo2_KeyPress(KeyAscii As Integer)
KeyAscii = AutoMatchCBBox(Combo2, KeyAscii)
End Sub

Private Sub DataGrid2_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
bno = DataGrid2.Columns(0)
qty = DataGrid2.Columns(2)
exp = DataGrid2.Columns(1)

End Sub

Private Sub Combo2_LostFocus()
If rmedicine.State = adStateOpen Then
    rmedicine.Close
End If
    rmedicine.Open "select * from [Medicine] where(M_name='" & Combo2.Text & "')", conn, adOpenDynamic, adLockPessimistic
If rmedicine.RecordCount <> 1 Then
    MsgBox "PLEASE ENTER AN EXISTING MEDICINE!!", vbInformation, "INFO"
    rmedicine.Close
Else
    If rbatch.State = adStateOpen Then
        rbatch.Close
    End If
    rbatch.Open "select B_no,B_exp_date,B_qty from [Batch] where(M_name='" & Combo2.Text & "')", conn, adOpenDynamicm, adLockPessimistic
    If rbatch.RecordCount <> 0 Then
        Set DataGrid2.DataSource = rbatch
        bno = DataGrid2.Columns(0)
        qty = DataGrid2.Columns(2)
        exp = DataGrid2.Columns(1)
        DataGrid2.Enabled = True
        Text1.Enabled = True
        Text2.Enabled = False
    Else
        MsgBox "This medicine does not have any Batch!!,please select a different medicine or add medicine for that medicine", vbInformation
    End If
End If

End Sub

Private Sub Form_Load()
Dim name As String
If rdoctor.State = adStateOpen Then
rdoctor.Close
End If
rdoctor.Open "select D_name from [Doctor]", conn, adOpenDynamic, adLockPessimistic
name = "D_name"
Call populate(rdoctor, Combo1, name)

If rmedicine.State = adStateOpen Then
rmedicine.Close
End If
rmedicine.Open "select M_name from [Medicine]", conn, adOpenDynamic, adLockPessimistic
name = "M_name"
Call populate(rmedicine, Combo2, name)
End Sub

Private Sub Label1_Click()
bill2.Top = 0
bill2.Left = 0
bill2.Show
Unload Me
End Sub

Private Sub Label11_Click()
If rbill_2.State = adStateOpen Then
    If rbill_2.RecordCount = 0 Then
        rbill_1.Open "delete from bill1 where B_id= " & bid & " "
    End If
End If
If flag = False Then
    mainform.Left = 0
    mainform.Top = 0
    mainform.Show
Else
    mainform2.Top = 0
    mainform2.Left = 0
    mainform2.Show
End If
Unload Me
End Sub

Private Sub Label12_Click()
If Text1.Text = "" Or Text2.Text = "" Or Combo1.Text = "" Or Combo2.Text = "" Then
    MsgBox "SOME FIELDS ARE INCOMPLETE,PLEASE COMPLETE THEM", vbInformation, "INFO"
Else
Label1.Enabled = False
Label11.Enabled = False
If rbill_2.State = adStateOpen Then
    rbill_2.Close
End If
rbill_2.Open "select * from bill2 where M_name='" & Combo2.Text & "' and B_id=" & bid & " and Batch_No='" & bno & "'", conn, adOpenDynamic, adLockPessimistic
If rbill_2.RecordCount <> 0 Then
    If Format(DataGrid2.Columns(1), "yyyy-mm-dd") < Format(Now, "yyyy-mm-dd") Then
        MsgBox "THE MEDICINE HAS EXPIRED!!!, PLEASE SELECT A DIFFERENT BATCH", vbInformation
    ElseIf Val(Text1.Text) > 0 Or Abs(Val(Text1.Text)) < rbill_2!M_qty Then
        If rbill_2.State = adStateOpen Then
            rbill_2.Close
        End If
        rbill_2.Open "update bill2 set m_qty=m_qty+ " & Text1.Text & ",m_price=m_price+ " & Text2.Text & " where M_name='" & Combo2.Text & "' and Batch_no='" & bno & "'", conn, adOpenDynamic, adLockPessimistic
        If rbatch.State = adStateOpen Then
            rbatch.Close
        End If
        rbatch.Open "update [Batch] set B_qty=B_qty-" & Text1.Text & " where(B_no='" & bno & "')", conn, adOpenDynamic, adLockPessimistic
        
    ElseIf Abs(Val(Text1.Text)) = rbill_2!M_qty Then
        If rbill_2.State = adStateOpen Then
            rbill_2.Close
        End If
        rbill_2.Open "delete from bill2 where M_name='" & Combo2.Text & "' and B_id=" & bid & " and Batch_no='" & bno & "'"
        rbill_2.Open "select * from [Bill2] where(B_ID=" & bid & ")", conn, adOpenDynamic, adLockPessimistic
        If rbill_2.RecordCount <> 0 Then
            Set DataGrid1.DataSource = rbill_2
        Else
            Label11.Enabled = True
            Label1.Enabled = True
        End If
        Set DataGrid1.DataSource = rbill_2
        If rbatch.State = adStateOpen Then
            rbatch.Close
        End If
        rbatch.Open "update [Batch] set B_qty=B_qty-" & Text1.Text & " where(B_no='" & bno & "')", conn, adOpenDynamic, adLockPessimistic
    ElseIf Abs(Val(Text1.Text)) > rbill_2!M_qty Then
        MsgBox "THE PROVIDED QUANTITY TO DELETE IS GREATER THAN THE ALREADY EXISTING QUANTITY", vbCritical, "INFO"
        Text1.Text = ""
        Text1.SetFocus
    End If
    If rbill_2.State = adStateOpen Then
        rbill_2.Close
    End If
    rbill_2.Open "select * from [Bill2] where(B_ID=" & bid & ")", conn, adOpenDynamic, adLockPessimistic
    Set DataGrid1.DataSource = rbill_2
    Combo1.Enabled = False
    Combo2.Text = ""
    Text1.Text = ""
    Text2.Text = ""
    Text1.Enabled = False
    Text2.Enabled = False
    Label15.Enabled = True
    DataGrid2.Refresh
    DataGrid2.Enabled = False
    Combo2.SetFocus

ElseIf Format(DataGrid2.Columns(1), "yyyy-mm-dd") < Format(Now, "yyyy-mm-dd") Then
        If rbill_2.State = adStateOpen Then
            rbill_2.Close
        End If
        rbill_2.Open "select * from [Bill2] where(B_ID=" & bid & ")", conn, adOpenDynamic, adLockPessimistic
        Set DataGrid1.DataSource = rbill_2
        MsgBox "THE MEDICINE HAS EXPIRED!!!, PLEASE SELECT A DIFFERENT BATCH", vbInformation
Else
    If Val(Text1.Text) < 0 Then
        MsgBox "CANT DELETE NON-EXISTING ITEMS!!", vbInformation, "INFO"
    Else
    If rbill_2.State = adStateOpen Then
    rbill_2.Close
    End If
    rbill_2.Open "select * from [Bill2] where(B_ID=" & bid & ")", conn, adOpenDynamic, adLockPessimistic
    rbill_2.AddNew
    rbill_2!B_ID = bid
    rbill_2!M_name = Combo2.Text
    rbill_2!Batch_no = bno
    rbill_2!M_qty = Val(Text1.Text)
    rbill_2!M_price = Val(Text2.Text)
    rbill_2!Batch_exp = exp
    rbill_2.Update
    If rbatch.State = adStateOpen Then
        rbatch.Close
    End If
    rbatch.Open "update [Batch] set B_qty=B_qty-" & Val(Text1.Text) & " where(B_no='" & bno & "')", conn, adOpenDynamic, adLockPessimistic
    Set DataGrid1.DataSource = rbill_2
    Combo1.Enabled = False
    Combo2.Text = ""
    Text1.Text = ""
    Text2.Text = ""
    Text1.Enabled = False
    Text2.Enabled = False
    DataGrid2.Enabled = False
    Combo2.SetFocus
    Label15.Enabled = True
    End If
End If
End If
End Sub

Private Sub Label15_Click()

If rdoctor.State = adStateOpen Then
    rdoctor.Close
End If
    
If rbill_1.State = adStateOpen Then
    rbill_1.Close
End If
If rbill_2.State = adStateOpen Then
    rbill_2.Close
End If
If rbatch.State = adStateOpen Then
    rbatch.Close
End If
If rmedicine.State = adStateOpen Then
    rmedicine.Close
End If


rbill_2.Open "select sum(M_price) as total from [Bill2] where(B_ID=" & bid & ")", conn, adOpenDynamic, adLockPessimistic
tamt = Val(rbill_2!Total)
rbill_2.Close
rbill_1.Open "update [Bill1] set B_amt=" & tamt & " where(B_ID=" & bid & ")", conn, adOpenDynamic, adLockPessimistic
MsgBox "BILL GENERATED!!", vbInformation, "INFO"

Label7.Visible = True
Shape3.Visible = True
Label12.Visible = False
Shape7.Visible = False
Combo1.Text = ""
Combo1.Enabled = True
Label15.Enabled = False
Label11.Enabled = True
Label1.Enabled = True
Combo2.Enabled = False
End Sub

Private Sub Label7_Click()
If Text1.Text = "" Or Text2.Text = "" Or Combo1.Text = "" Or Combo2.Text = "" Then
    MsgBox "SOME FIELDS ARE INCOMPLETE,PLEASE COMPLETE THEM", vbInformation, "INFO"
ElseIf Format(DataGrid2.Columns(1), "yyyy-mm-dd") < Format(Now, "yyyy-mm-dd") Then
    MsgBox "THE MEDICINE HAS EXPIRED!!!, PLEASE SELECT A DIFFERENT BATCH", vbInformation
ElseIf Val(Text1.Text) < 0 Then
    MsgBox "CANT DELETE NON-EXISTING ITEMS!!", vbInformation, "INFO"
Else
    
    If rdoctor.State = adStateOpen Then
        rdoctor.Close
    End If
    
    If rbill_1.State = adStateOpen Then
        rbill_1.Close
    End If
    rdoctor.Open "select D_regno from [Doctor] where(D_name='" & Combo1.Text & "')", conn, adOpenDynamic, adLockPessimistic
    rbill_1.Open "select * from [Bill1]", conn, adOpenDynamic, adLockPessimistic
    rbill_1.AddNew
    rbill_1!D_regno = rdoctor!D_regno
    rbill_1!B_date = Format(Now, "yyyy/mm/dd")
    rbill_1!E_id = eid
    rbill_1.Update
    bid = rbill_1!B_ID
    rbill_1.Close
    If rbill_2.State = adStateOpen Then
        rbill_2.Close
    End If
    rbill_2.Open "select * from [Bill2] where(B_ID=" & bid & ")", conn, adOpenDynamic, adLockPessimistic
    rbill_2.AddNew
    rbill_2!B_ID = bid
    rbill_2!M_name = Combo2.Text
    rbill_2!Batch_no = bno
    rbill_2!M_qty = Val(Text1.Text)
    rbill_2!M_price = Val(Text2.Text)
    rbill_2!Batch_exp = exp
    rbill_2.Update
    If rbatch.State = adStateOpen Then
        rbatch.Close
    End If
    'rbatch.Open "select B_qty from [Batch] where(B_no='" & bno & "')", conn, adOpenDynamic, adLockPessimistic
    rbatch.Open "update [Batch] set B_qty=B_qty-" & Val(Text1.Text) & " where(B_no='" & bno & "')", conn, adOpenDynamic, adLockPessimistic
    'rbatch!B_qty = rbatch!B_qty - Val(Text1.Text)
    'rbatch.Update
    'rbatch.Close
    Set DataGrid1.DataSource = rbill_2
    Label7.Visible = False
    Shape3.Visible = False
    Label12.Visible = True
    Shape7.Visible = True
    Label15.Enabled = True
    Combo1.Enabled = False
    Combo2.Text = ""
    Text1.Text = ""
    Text2.Text = ""
    Text1.Enabled = False
    Text2.Enabled = False
    Label11.Enabled = False
    DataGrid2.Enabled = False
    Combo2.SetFocus
    Label1.Enabled = False
    
End If

End Sub

Private Sub Text1_GotFocus()

    If Combo2.Text = "" Then
        MsgBox "PLEASE SELECT MEDICINE FIRST!!", vbInformation, "INFO"
        Combo2.SetFocus
    End If

End Sub

Private Sub Text1_LostFocus()
If Text1.Text = "" Then
    MsgBox "PLEASE ENTER QUANTITY!!", vbInformation, "INFO"
Else
    If num(Text1.Text) Then
    If Val(Text1.Text) > qty Then
        MsgBox "REQUIRED QUANTITY IS > THAN AVAILABLE QUANTITY", vbInformation, "INFO"
    Else
        Text2.Enabled = True
        Text2.SetFocus
    End If
    Else
        MsgBox "Please enter only numbers!!!", vbInformation
        Text1.SetFocus
        Text1.Text = ""
    End If
End If
If Val(Text1.Text) > 0 Then
    Label12.Caption = "ADD"
Else
    Label12.Caption = "SUB"
End If

End Sub

Private Sub Text2_GotFocus()
If rmedicine.State = adStateOpen Then
    rmedicine.Close
End If
rmedicine.Open "select M_sp from [Medicine] where(M_name='" & Combo2.Text & "')", conn, adOpenDynamic, adLockPessimistic
Text2.Text = Val(Text1.Text) * rmedicine!M_sp
rmedicine.Close
End Sub

Private Sub Text2_LostFocus()
If num(Text2.Text) <> True Then
    MsgBox "price can't be alphabets", vbInformation, "INFO"
    Text2.Text = ""
    Text2_GotFocus
End If

End Sub
