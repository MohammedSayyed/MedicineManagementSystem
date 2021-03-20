Attribute VB_Name = "Module2"
Option Explicit

'
'Implement CComboBox:SelectString to search and select item
'Y. Huang <yinghsuan_h@yahoo.com>
'
'Copyright? Naaa! But Copyleft...
'http://www.gnu.org/copyleft/copyleft.html
'
'In KeyPress Enevt Method: KeyAscii = AutoMatchCBBox(ComBoBox, KeyAscii)
'
'Reference: WinUser.h
'VB ComboBox doesn't have SelectString(), so SendMessage to the Window Handle
'#define CB_SELECTSTRING     0x014D
'#define CB_SHOWDROPDOWN     0x014F
'#define CBN_SELENDOK        9
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
                ByVal hwnd As Long, _
                ByVal wMsg As Long, _
                ByVal wParam As Long, _
                lParam As Any) As Long
Private Const CB_ERR = -1, CB_SELECTSTRING = &H14D, CB_SHOWDROPDOWN = &H14F, CBN_SELENDOK = 9

'call this function in KeyPress event method
Public Function AutoMatchCBBox(ByRef cbBox As ComboBox, ByVal KeyAscii As Integer) As Integer
    
        
    Dim strFindThis As String, bContinueSearch As Boolean
    Dim lResult As Long, lStart As Long, lLength As Long
    AutoMatchCBBox = 0 ' block cbBox since we handle everything
    bContinueSearch = True
    lStart = cbBox.SelStart
    lLength = cbBox.SelLength

    On Error GoTo ErrHandle
        
    If KeyAscii < 32 Then 'control char
        bContinueSearch = False
        cbBox.SelLength = 0 'select nothing since we will delete/enter
        If KeyAscii = Asc(vbBack) Then 'take care BackSpace and Delete first
            If lLength = 0 Then 'delete last char
                If Len(cbBox) > 0 Then ' in case user delete empty cbBox
                    cbBox.Text = Left(cbBox.Text, Len(cbBox) - 1)
                End If
            Else 'leave unselected char(s) and delete rest of text
                cbBox.Text = Left(cbBox.Text, lStart)
            End If
            cbBox.SelStart = Len(cbBox) 'set insertion position @ the end of string
        ElseIf KeyAscii = vbKeyReturn Then  'user select this string
            cbBox.SelStart = Len(cbBox)
            lResult = SendMessage(cbBox.hwnd, CBN_SELENDOK, 0, 0)
            AutoMatchCBBox = KeyAscii 'let caller a chance to handle "Enter"
        End If
    Else 'generate searching string
        If lLength = 0 Then
            strFindThis = cbBox.Text & Chr(KeyAscii) 'No selection, append it
        Else
            strFindThis = Left(cbBox.Text, lStart) & Chr(KeyAscii)
        End If
    End If
    
    If bContinueSearch Then 'need to search
        Call VBComBoBoxDroppedDown(cbBox)  'open dropdown list
        lResult = SendMessage(cbBox.hwnd, CB_SELECTSTRING, -1, ByVal strFindThis)
        If lResult = CB_ERR Then 'not found
            cbBox.Text = strFindThis 'set cbBox as whatever it is
            cbBox.SelLength = 0 'no selected char(s) since not found
            cbBox.SelStart = Len(cbBox) 'set insertion position @ the end of string
        Else
            'found string, highlight rest of string for user
            cbBox.SelStart = Len(strFindThis)
            cbBox.SelLength = Len(cbBox) - cbBox.SelStart
        End If
    End If
    On Error GoTo 0
    Exit Function
    
ErrHandle:
    'got problem, simply return whatever pass in
    Debug.Print "Failed: AutoCompleteComboBox due to : " & Err.Description
    Debug.Assert False
    AutoMatchCBBox = KeyAscii
    On Error GoTo 0
End Function

'open dorpdown list
Private Sub VBComBoBoxDroppedDown(ByRef cbBox As ComboBox)
    Call SendMessage(cbBox.hwnd, CB_SHOWDROPDOWN, Abs(True), 0)
End Sub


Public Sub populate(ByRef rs As ADODB.Recordset, ByRef Combo1 As VB.ComboBox, name As String)
With rs
    Do While Not .EOF
        Combo1.AddItem rs.Fields(name)
        .MoveNext
    Loop
    .Close
End With
End Sub

Public Sub validateMail(ByVal s As String)
If s <> "" Then
If Mid(s, 1, 1) = "@" Or Mid(s, 1, 1) = "." Then
    MsgBox "INVALID MAIL", vbCritical
ElseIf Mid(s, Len(s), 1) = "@" Or Mid(s, Len(s), 1) = "." Then
    MsgBox "INVALID MAIL", vbCritical
ElseIf InStr(s, "@") = False Or InStr(s, ".") = False Then
    MsgBox "INVALID MAIL", vbCritical
End If
End If
End Sub

Public Sub validateName(ByVal s As String)
Dim i As Integer
For i = 1 To Len(s)
    Dim str As String
    str = Mid(s, i, 1)
    If IsNumeric(str) Then
        MsgBox "INVALID NAME!!", vbCritical
        Exit For
    End If
Next
End Sub
Public Sub validatePhone(ByVal s As String)
Dim i As Integer
If Len(s) <> 10 Then
MsgBox "INVALID PHONE NUMBER!!", vbCritical
Else
    For i = 1 To Len(s)
        Dim str As String
        str = Mid(s, i, 1)
        If IsNumeric(str) = False Then
            MsgBox "INVALID PHONE NUMBER!!", vbCritical
            Exit For
        End If
    Next
End If
End Sub
Public Function num(ByVal s As String) As Boolean
 Dim i As Integer
 num = True
 For i = 1 To Len(s)
        Dim str As String
        str = Mid(s, i, 1)
        If str <> "-" Then
        If IsNumeric(str) = False Then
            num = False
            Exit For
        End If
        End If
Next

End Function
