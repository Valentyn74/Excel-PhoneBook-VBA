VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "���������� �����"
   ClientHeight    =   4845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5445
   OleObjectBlob   =   "PhoneBookForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ���������� ���������� ��� �������� ��������� ������ ��������
Private tName_ As String
Private tPhon_ As String
Private IsProcessing As Boolean ' ��� �������������� ���������� �������

' ��������� ��� ��������� ��������
Private Const SHEET_NAME As String = "PhoneBook"
Private Const BUTTON_EDIT As String = "�������������"
Private Const BUTTON_SAVE As String = "���������"

' �������� ����� � ����������� �����
Private Sub cmdClose_Click()
    ThisWorkbook.Save
    Unload Me
End Sub

Private Sub ListBox1_Change()

 If cmdUpdate.Caption = BUTTON_EDIT Then
  If ListBox1.ListIndex = -1 Then
    cmdUpdate.Enabled = False
    Else
     cmdUpdate.Enabled = True
      End If
    End If
End Sub

' ������� ����� ��� ������� ����� �� txtName
Private Sub txtName_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ClearFields
    cmdDelete.Enabled = False
    LoadData
End Sub

Private Sub txtName_Exit(ByVal Cancel As MSForms.ReturnBoolean)
  '������ ������ ����� � TextBox �� ���������
     '������ � ���������� ��� ������ "Variant", ����� ���������� ����� ���� ���������� � VBA-������.
    Dim varText As Variant
    Dim i As Long
    '������ ������ �� ���������� � ���������� "varText".
    varText = Me.txtName.Text
    '��������� ������ �� ����� �� ��������.
    varText = Split(varText, " ")
    '��������� ������ ����� �� ���� ������.
        '��������� ��������� � ������� "varText" ���������� � ����.
    For i = 0 To UBound(varText) Step 1
        varText(i) = UCase(Left(varText(i), 1)) & Mid(varText(i), 2)
    Next i
    '�������� ���� � ���� ������ � ��������� ���������� �������
        '� ���������� "varText".
    varText = Join(varText, " ")
    '����� � ���������� "varText" ���������� ����� � ���������� ����,������� ����� �������� ����, ���� �����.
        txtName.Text = varText
End Sub

Private Sub txtPhone_Change()
Dim inputText As String
    Dim cleaned As String
    Dim i As Long
    
    inputText = txtPhone.Value
    ' ������� ��, ����� ���� � +
    cleaned = ""
    For i = 1 To Len(inputText)
        If IsNumeric(Mid(inputText, i, 1)) Or Mid(inputText, i, 1) = "+" Then
            cleaned = cleaned & Mid(inputText, i, 1)
        End If
    Next i
    
    ' ���� ���������� � 0, ��������� +380
    If Left(cleaned, 1) = "0" Then
        cleaned = "+380" & Mid(cleaned, 2)
    End If
    
    ' ������������ ����� �� +380 � 9 ����
    If Len(cleaned) > 13 Then
        cleaned = Left(cleaned, 13)
    End If
    
    ' ��������� ���� ��� ������ �������
    Application.EnableEvents = False
    txtPhone.Value = cleaned
    Application.EnableEvents = True
End Sub

Private Sub UserForm_Activate()
cmdAdd.Enabled = False
cmdDelete.Enabled = False
cmdUpdate.Enabled = False
End Sub

' ������������� �����
Private Sub UserForm_Initialize()
    ' ��������� ListBox
    With ListBox1
        .ColumnCount = 2
        .ColumnWidths = "160;70"
    End With
    ' �������� ������
    LoadData
End Sub

' ���������� ������ �� ����� PhoneBook �� �����
Private Sub SortData()
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub ' ��� ������ ��� ����������
    
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("A2:A" & lastRow), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range("A1:B" & lastRow)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
    ThisWorkbook.Save
    Exit Sub
ErrorHandler:
    MsgBox "������ ���������� ������: " & Err.Description, vbCritical
End Sub

' �������� ������ �� ����� � ListBox
Private Sub LoadData()
    On Error GoTo ErrorHandler
    ' ���������� ����� ���������
    SortData
    ListBox1.Clear
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    Dim data As Variant
    data = ws.Range("A2:B" & lastRow).Value
    ListBox1.List = data
    Exit Sub
ErrorHandler:
    MsgBox "������ �������� ������: " & Err.Description, vbCritical
End Sub

' ���������� ������ ��������
Private Sub cmdAdd_Click()
    On Error GoTo ErrorHandler
    txtName.Value = Trim(txtName.Value)
    txtPhone.Value = Trim(txtPhone.Value)
    
    ' �������� ������������� �����
    If txtName.Value = "" Or txtPhone.Value = "" Then
        MsgBox "������� ��� � ����� ��������!", vbExclamation, "���������� ��������"
        Exit Sub
    End If
    
    ' ������������ � �������� ������ ��������
    txtPhone.Value = NormalizePhoneNumber(txtPhone.Value)
    ' ���������� ����� (������ ����� ������������)
    Debug.Print "cmdAdd_Click: Normalized Phone=" & txtPhone.Value
    If Not IsValidPhone(txtPhone.Value) Then
        MsgBox "������������ ����� ��������! ����� ������ ��������� 9 ���� ����� +380 (��������, +380501234567).", vbExclamation, "���������� ��������"
        Exit Sub
    End If
    
    ' �������� �� ������������� ��������
    Dim r As Range
    Set r = FindContact(txtName.Value)
    If Not r Is Nothing Then
        MsgBox "������� (" & txtName.Value & ") ��� ����������! ����������� '�������������'.", vbInformation, "���������� ��������"
        Exit Sub
    End If
    
    ' ���������� ��������
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(lastRow, 1).Value = txtName.Value
    ws.Cells(lastRow, 2).Value = txtPhone.Value
    ThisWorkbook.Save
    
    MsgBox "������� (" & txtName.Value & ") ��������!", vbInformation, "���������� ��������"
    
    ' ���������� � �������
    ClearFields
    cmdDelete.Enabled = False
    LoadData
    Exit Sub
ErrorHandler:
    MsgBox "������ ���������� ��������: " & Err.Description, vbCritical
End Sub

' �������������� ��������
Private Sub cmdUpdate_Click()
    On Error GoTo ErrorHandler
    
    
    txtName.Value = Trim(txtName.Value)
    txtPhone.Value = Trim(txtPhone.Value)
    
    If cmdUpdate.Caption = BUTTON_EDIT Then

        ' ����� ��������������
        If txtName.Value = "" And txtPhone.Value = "" Then
            MsgBox "��������� ���� ��� ��������������!", vbExclamation, "�������������� ��������"
            Exit Sub
        End If
        cmdUpdate.Caption = BUTTON_SAVE
        tName_ = txtName.Value
        tPhon_ = txtPhone.Value
        txtName.BackColor = &HC0C0FF
        txtPhone.BackColor = &HC0C0FF
        cmdDelete.Enabled = False
        cmdAdd.Enabled = False
    Else
        ' ����� ����������
        If txtName.Value = "" Or txtPhone.Value = "" Then
            MsgBox "��������� ��� ����!", vbExclamation, "�������������� ��������"
            Exit Sub
        End If
        
       
        ' ������������ � �������� ������ ��������
        txtPhone.Value = NormalizePhoneNumber(txtPhone.Value)
        ' ���������� ����� (������ ����� ������������)
        Debug.Print "cmdUpdate_Click: Normalized Phone=" & txtPhone.Value
        If Not IsValidPhone(txtPhone.Value) Then
            MsgBox "������������ ����� ��������! ����� ������ ��������� 9 ���� ����� +380 (��������, +380501234567).", vbExclamation, "�������������� ��������"
            Exit Sub
        End If
        
        Dim r As Range
        Set r = FindContact(tName_)
        If Not r Is Nothing Then
            r.Value = txtName.Value
            r.Offset(0, 1).Value = txtPhone.Value
            ThisWorkbook.Save
            MsgBox "������� (" & tName_ & ") ������!", vbInformation, "�������������� ��������"
        Else
            MsgBox "������� (" & tName_ & ") �� ������!", vbInformation, "�������������� ��������"
        End If
        ' ����� �����
        cmdUpdate.Caption = BUTTON_EDIT
        txtName.BackColor = &H80000005
        txtPhone.BackColor = &H80000005
        cmdAdd.Enabled = True
         cmdUpdate.Enabled = False
        ClearFields
        LoadData
    End If
    Exit Sub
ErrorHandler:
    MsgBox "������ �������������� ��������: " & Err.Description, vbCritical
End Sub

' �������� ��������
Private Sub cmdDelete_Click()
 
    On Error GoTo ErrorHandler
    If ListBox1.ListIndex = -1 Then
        MsgBox "�������� ������� ��� ��������!", vbExclamation, "�������� ��������"
        Exit Sub
    End If

    Dim contactName As String
    contactName = ListBox1.List(ListBox1.ListIndex, 0)

    If MsgBox("�� ������������� ������ ������� �������" + Chr(13) + "(" & contactName & ")?", vbYesNo + vbQuestion, "�������� ��������") = vbNo Then
        Exit Sub
    End If

    Dim r As Range
    Set r = FindContact(contactName)
    If Not r Is Nothing Then
        r.EntireRow.Delete
        ThisWorkbook.Save
        MsgBox "������� (" & contactName & ") �����!", vbInformation, "�������� ��������"
        ClearFields
        cmdDelete.Enabled = False
        LoadData
    Else
        MsgBox "������� (" & contactName & ") �� ������!", vbInformation, "�������� ��������"
    End If
    Exit Sub
ErrorHandler:
    MsgBox "������ �������� ��������: " & Err.Description, vbCritical
End Sub

' ����� �������� �� �����
Private Function FindContact(name As String) As Range
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "���� '" & SHEET_NAME & "' �� ������!", vbCritical
        Exit Function
    End If
    Set FindContact = ws.Columns(1).Find(What:=name, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
End Function

' ����� �������� � ListBox
Private Sub ListBox1_Click()
    On Error GoTo ErrorHandler
    If ListBox1.ListIndex <> -1 And ListBox1.ListCount > 0 Then
        IsProcessing = True
        txtName.Value = ListBox1.List(ListBox1.ListIndex, 0)
        txtPhone.Value = ListBox1.List(ListBox1.ListIndex, 1)
        cmdDelete.Enabled = True
        IsProcessing = False
    End If
    Exit Sub
ErrorHandler:
    IsProcessing = False
    MsgBox "������ ������ ��������: " & Err.Description & vbCrLf & _
           "ListIndex: " & ListBox1.ListIndex & ", ListCount: " & ListBox1.ListCount, vbCritical
End Sub

' ����� �� ����� ��� ��������� txtName
Private Sub txtName_Change()
    If IsProcessing Then Exit Sub
    If txtName.Value <> "" And cmdUpdate.Caption = BUTTON_EDIT Then
        SearchData
        cmdAdd.Enabled = True
    Else
        LoadData
         cmdAdd.Enabled = False
    End If
End Sub

' ���������� ������ � ListBox �� �����
Private Sub SearchData()
    On Error GoTo ErrorHandler
    ListBox1.Clear
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    Dim i As Long
    Dim searchText As String
    searchText = LCase(txtName.Value)
    
    For i = 2 To lastRow
        If LCase(ws.Cells(i, 1).Value) Like "*" & searchText & "*" Then
            ListBox1.AddItem
            ListBox1.List(ListBox1.ListCount - 1, 0) = ws.Cells(i, 1).Value
            ListBox1.List(ListBox1.ListCount - 1, 1) = ws.Cells(i, 2).Value
        End If
    Next i
    Exit Sub
ErrorHandler:
    MsgBox "������ ������: " & Err.Description, vbCritical
End Sub

' ������������ ������ �������� (���������� +380, �������� ������ ��������)
Private Function NormalizePhoneNumber(phone As String) As String
    On Error GoTo ErrorHandler
    ' ������� �������, ������ � ��� ���������� �������, ����� +
    Dim cleaned As String
    cleaned = Trim(phone) ' ������� ��������� � �������� �������
    cleaned = Replace(cleaned, " ", "")
    cleaned = Replace(cleaned, "-", "")
    
    ' ���� ����� ���������� � +380 � �������� 9 ���� �����, ���������� ��� ����
    If Left(cleaned, 4) = "+380" And Len(cleaned) = 12 And IsNumeric(Mid(cleaned, 5)) Then
        NormalizePhoneNumber = cleaned
        Exit Function
    End If
    
    ' ������� ��� ���������� ������� ��� ��������
    Dim digitsOnly As String
    Dim i As Long
    For i = 1 To Len(cleaned)
        If IsNumeric(Mid(cleaned, i, 1)) Then
            digitsOnly = digitsOnly & Mid(cleaned, i, 1)
        End If
    Next i
    
    ' ���� ����� ���������� � 0, ������� 0
    If Left(digitsOnly, 1) = "0" Then
        digitsOnly = Mid(digitsOnly, 2)
    End If
    
    ' ���� ����� ������� �������� 9 ����, ��������� +380
    If Len(digitsOnly) = 9 And IsNumeric(digitsOnly) Then
        NormalizePhoneNumber = "+380" & digitsOnly
    Else
        NormalizePhoneNumber = cleaned ' ���������� ��������, ���� �� �������������
    End If
    Exit Function
ErrorHandler:
    NormalizePhoneNumber = phone ' ���������� �������� ����� ��� ������
End Function

' ��������� ������ ��������
Private Function IsValidPhone(phone As String) As Boolean
    On Error GoTo ErrorHandler
    ' ����������� �����
    Dim normalized As String
    normalized = NormalizePhoneNumber(phone)
    
    ' ��������: ������ ���������� � +380 � ��������� ����� 9 ���� �����
    If normalized Like "+380#########" Then
        IsValidPhone = True
    Else
        IsValidPhone = False
    End If
    
    ' ���������� ����� (������ ����� ������������)
    Debug.Print "IsValidPhone: Input=" & phone & ", Normalized=" & normalized & ", Valid=" & IsValidPhone
    Exit Function
ErrorHandler:
    IsValidPhone = False
End Function

' ������� ����� �����
Private Sub ClearFields()
    txtName.Value = ""
    txtPhone.Value = ""
End Sub

' ���������� ����� ��� �������� �����
Private Sub UserForm_Terminate()
    ThisWorkbook.Save
End Sub

