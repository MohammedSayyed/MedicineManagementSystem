Attribute VB_Name = "Module1"

Public conn As ADODB.Connection
Public rdoctor As ADODB.Recordset
Public remployee As ADODB.Recordset
Public rsupplier As ADODB.Recordset
Public rbatch As ADODB.Recordset
Public rbatch1 As ADODB.Recordset
Public rmedicine As ADODB.Recordset
Public rbill_1 As ADODB.Recordset
Public rbill_2 As ADODB.Recordset
Public billID As Integer
Public bid, eid As Integer
Public flag As Boolean


Public Sub main()
eid = 0
Set conn = New ADODB.Connection
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\medical.mdb;Persist Security Info=False"
conn.CursorLocation = adUseClient

Set rdoctor = New ADODB.Recordset
rdoctor.Open "select * from [Doctor]", conn, adOpenDynamic, adLockPessimistic

Set remployee = New ADODB.Recordset
remployee.Open "select * from [Employee]", conn, adOpenDynamic, adLockPessimistic

Set rsupplier = New ADODB.Recordset
rsupplier.Open "select * from [Supplier]", conn, adOpenDynamic, adLockPessimistic

Set rbatch = New ADODB.Recordset
rbatch.Open "select * from [Batch]", conn, adOpenDynamic, adLockPessimistic

Set rbatch1 = New ADODB.Recordset


Set rmedicine = New ADODB.Recordset
rmedicine.Open "select * from [Medicine]", conn, adOpenDynamic, adLockPessimistic

Set rbill_1 = New ADODB.Recordset
rbill_1.Open "select * from [Bill1]", conn, adOpenDynamic, adLockPessimistic

Set rbill_2 = New ADODB.Recordset
rbill_2.Open "select * from [Bill2]", conn, adOpenDynamic, adLockPessimistic

batch2.Show
End Sub
