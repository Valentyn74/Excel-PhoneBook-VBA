VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Телефонная книга"
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
' Глобальные переменные для хранения временных данных контакта
Private tName_ As String
Private tPhon_ As String
Private IsProcessing As Boolean ' Для предотвращения конфликтов событий

' Константы для строковых значений
Private Const SHEET_NAME As String = "PhoneBook"
Private Const BUTTON_EDIT As String = "Редактировать"
Private Const BUTTON_SAVE As String = "Сохранить"

' Закрытие формы с сохранением книги
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

' Очистка полей при двойном клике на txtName
Private Sub txtName_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ClearFields
    cmdDelete.Enabled = False
    LoadData
End Sub

Private Sub txtName_Exit(ByVal Cancel As MSForms.ReturnBoolean)
  'Меняем первые буквы в TextBox на заглавные
     'Делаем у переменной тип данных "Variant", чтобы переменную можно было превратить в VBA-массив.
    Dim varText As Variant
    Dim i As Long
    'Взятие данных из текстбокса в переменную "varText".
    varText = Me.txtName.Text
    'Разбиваем данные на слова по пробелам.
    varText = Split(varText, " ")
    'Изменение первой буквы во всех словах.
        'Нумерация элементов в массиве "varText" начинается с нуля.
    For i = 0 To UBound(varText) Step 1
        varText(i) = UCase(Left(varText(i), 1)) & Mid(varText(i), 2)
    Next i
    'Соединие слов в одну строку и помещение результата обратно
        'в переменную "varText".
    varText = Join(varText, " ")
    'Здесь в переменной "varText" получается текст в правильном виде,который можно передать туда, куда нужно.
        txtName.Text = varText
End Sub

Private Sub txtPhone_Change()
Dim inputText As String
    Dim cleaned As String
    Dim i As Long
    
    inputText = txtPhone.Value
    ' Удаляем всё, кроме цифр и +
    cleaned = ""
    For i = 1 To Len(inputText)
        If IsNumeric(Mid(inputText, i, 1)) Or Mid(inputText, i, 1) = "+" Then
            cleaned = cleaned & Mid(inputText, i, 1)
        End If
    Next i
    
    ' Если начинается с 0, добавляем +380
    If Left(cleaned, 1) = "0" Then
        cleaned = "+380" & Mid(cleaned, 2)
    End If
    
    ' Ограничиваем длину до +380 и 9 цифр
    If Len(cleaned) > 13 Then
        cleaned = Left(cleaned, 13)
    End If
    
    ' Обновляем поле без вызова события
    Application.EnableEvents = False
    txtPhone.Value = cleaned
    Application.EnableEvents = True
End Sub

Private Sub UserForm_Activate()
cmdAdd.Enabled = False
cmdDelete.Enabled = False
cmdUpdate.Enabled = False
End Sub

' Инициализация формы
Private Sub UserForm_Initialize()
    ' Настройка ListBox
    With ListBox1
        .ColumnCount = 2
        .ColumnWidths = "160;70"
    End With
    ' Загрузка данных
    LoadData
End Sub

' Сортировка данных на листе PhoneBook по имени
Private Sub SortData()
    On Error GoTo ErrorHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub ' Нет данных для сортировки
    
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
    MsgBox "Ошибка сортировки данных: " & Err.Description, vbCritical
End Sub

' Загрузка данных из листа в ListBox
Private Sub LoadData()
    On Error GoTo ErrorHandler
    ' Сортировка перед загрузкой
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
    MsgBox "Ошибка загрузки данных: " & Err.Description, vbCritical
End Sub

' Добавление нового контакта
Private Sub cmdAdd_Click()
    On Error GoTo ErrorHandler
    txtName.Value = Trim(txtName.Value)
    txtPhone.Value = Trim(txtPhone.Value)
    
    ' Проверка заполненности полей
    If txtName.Value = "" Or txtPhone.Value = "" Then
        MsgBox "Введите имя и номер телефона!", vbExclamation, "Добавление контакта"
        Exit Sub
    End If
    
    ' Нормализация и проверка номера телефона
    txtPhone.Value = NormalizePhoneNumber(txtPhone.Value)
    ' Отладочный вывод (убрать после тестирования)
    Debug.Print "cmdAdd_Click: Normalized Phone=" & txtPhone.Value
    If Not IsValidPhone(txtPhone.Value) Then
        MsgBox "Некорректный номер телефона! Номер должен содержать 9 цифр после +380 (например, +380501234567).", vbExclamation, "Добавление контакта"
        Exit Sub
    End If
    
    ' Проверка на существование контакта
    Dim r As Range
    Set r = FindContact(txtName.Value)
    If Not r Is Nothing Then
        MsgBox "Контакт (" & txtName.Value & ") уже существует! Используйте 'Редактировать'.", vbInformation, "Добавление контакта"
        Exit Sub
    End If
    
    ' Добавление контакта
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(lastRow, 1).Value = txtName.Value
    ws.Cells(lastRow, 2).Value = txtPhone.Value
    ThisWorkbook.Save
    
    MsgBox "Контакт (" & txtName.Value & ") добавлен!", vbInformation, "Добавление контакта"
    
    ' Обновление и очистка
    ClearFields
    cmdDelete.Enabled = False
    LoadData
    Exit Sub
ErrorHandler:
    MsgBox "Ошибка добавления контакта: " & Err.Description, vbCritical
End Sub

' Редактирование контакта
Private Sub cmdUpdate_Click()
    On Error GoTo ErrorHandler
    
    
    txtName.Value = Trim(txtName.Value)
    txtPhone.Value = Trim(txtPhone.Value)
    
    If cmdUpdate.Caption = BUTTON_EDIT Then

        ' Режим редактирования
        If txtName.Value = "" And txtPhone.Value = "" Then
            MsgBox "Заполните поля для редактирования!", vbExclamation, "Редактирование контакта"
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
        ' Режим сохранения
        If txtName.Value = "" Or txtPhone.Value = "" Then
            MsgBox "Заполните оба поля!", vbExclamation, "Редактирование контакта"
            Exit Sub
        End If
        
       
        ' Нормализация и проверка номера телефона
        txtPhone.Value = NormalizePhoneNumber(txtPhone.Value)
        ' Отладочный вывод (убрать после тестирования)
        Debug.Print "cmdUpdate_Click: Normalized Phone=" & txtPhone.Value
        If Not IsValidPhone(txtPhone.Value) Then
            MsgBox "Некорректный номер телефона! Номер должен содержать 9 цифр после +380 (например, +380501234567).", vbExclamation, "Редактирование контакта"
            Exit Sub
        End If
        
        Dim r As Range
        Set r = FindContact(tName_)
        If Not r Is Nothing Then
            r.Value = txtName.Value
            r.Offset(0, 1).Value = txtPhone.Value
            ThisWorkbook.Save
            MsgBox "Контакт (" & tName_ & ") изменён!", vbInformation, "Редактирование контакта"
        Else
            MsgBox "Контакт (" & tName_ & ") не найден!", vbInformation, "Редактирование контакта"
        End If
        ' Сброс формы
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
    MsgBox "Ошибка редактирования контакта: " & Err.Description, vbCritical
End Sub

' Удаление контакта
Private Sub cmdDelete_Click()
 
    On Error GoTo ErrorHandler
    If ListBox1.ListIndex = -1 Then
        MsgBox "Выберите контакт для удаления!", vbExclamation, "Удаление контакта"
        Exit Sub
    End If

    Dim contactName As String
    contactName = ListBox1.List(ListBox1.ListIndex, 0)

    If MsgBox("Вы действительно хотите удалить контакт" + Chr(13) + "(" & contactName & ")?", vbYesNo + vbQuestion, "Удаление контакта") = vbNo Then
        Exit Sub
    End If

    Dim r As Range
    Set r = FindContact(contactName)
    If Not r Is Nothing Then
        r.EntireRow.Delete
        ThisWorkbook.Save
        MsgBox "Контакт (" & contactName & ") удалён!", vbInformation, "Удаление контакта"
        ClearFields
        cmdDelete.Enabled = False
        LoadData
    Else
        MsgBox "Контакт (" & contactName & ") не найден!", vbInformation, "Удаление контакта"
    End If
    Exit Sub
ErrorHandler:
    MsgBox "Ошибка удаления контакта: " & Err.Description, vbCritical
End Sub

' Поиск контакта по имени
Private Function FindContact(name As String) As Range
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "Лист '" & SHEET_NAME & "' не найден!", vbCritical
        Exit Function
    End If
    Set FindContact = ws.Columns(1).Find(What:=name, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
End Function

' Выбор контакта в ListBox
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
    MsgBox "Ошибка выбора контакта: " & Err.Description & vbCrLf & _
           "ListIndex: " & ListBox1.ListIndex & ", ListCount: " & ListBox1.ListCount, vbCritical
End Sub

' Поиск по имени при изменении txtName
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

' Фильтрация данных в ListBox по имени
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
    MsgBox "Ошибка поиска: " & Err.Description, vbCritical
End Sub

' Нормализация номера телефона (добавление +380, удаление лишних символов)
Private Function NormalizePhoneNumber(phone As String) As String
    On Error GoTo ErrorHandler
    ' Удаляем пробелы, дефисы и все нецифровые символы, кроме +
    Dim cleaned As String
    cleaned = Trim(phone) ' Удаляем начальные и конечные пробелы
    cleaned = Replace(cleaned, " ", "")
    cleaned = Replace(cleaned, "-", "")
    
    ' Если номер начинается с +380 и содержит 9 цифр после, возвращаем как есть
    If Left(cleaned, 4) = "+380" And Len(cleaned) = 12 And IsNumeric(Mid(cleaned, 5)) Then
        NormalizePhoneNumber = cleaned
        Exit Function
    End If
    
    ' Удаляем все нецифровые символы для проверки
    Dim digitsOnly As String
    Dim i As Long
    For i = 1 To Len(cleaned)
        If IsNumeric(Mid(cleaned, i, 1)) Then
            digitsOnly = digitsOnly & Mid(cleaned, i, 1)
        End If
    Next i
    
    ' Если номер начинается с 0, убираем 0
    If Left(digitsOnly, 1) = "0" Then
        digitsOnly = Mid(digitsOnly, 2)
    End If
    
    ' Если после очистки осталось 9 цифр, добавляем +380
    If Len(digitsOnly) = 9 And IsNumeric(digitsOnly) Then
        NormalizePhoneNumber = "+380" & digitsOnly
    Else
        NormalizePhoneNumber = cleaned ' Возвращаем исходный, если не соответствует
    End If
    Exit Function
ErrorHandler:
    NormalizePhoneNumber = phone ' Возвращаем исходный номер при ошибке
End Function

' Валидация номера телефона
Private Function IsValidPhone(phone As String) As Boolean
    On Error GoTo ErrorHandler
    ' Нормализуем номер
    Dim normalized As String
    normalized = NormalizePhoneNumber(phone)
    
    ' Проверка: должен начинаться с +380 и содержать ровно 9 цифр после
    If normalized Like "+380#########" Then
        IsValidPhone = True
    Else
        IsValidPhone = False
    End If
    
    ' Отладочный вывод (убрать после тестирования)
    Debug.Print "IsValidPhone: Input=" & phone & ", Normalized=" & normalized & ", Valid=" & IsValidPhone
    Exit Function
ErrorHandler:
    IsValidPhone = False
End Function

' Очистка полей формы
Private Sub ClearFields()
    txtName.Value = ""
    txtPhone.Value = ""
End Sub

' Сохранение книги при закрытии формы
Private Sub UserForm_Terminate()
    ThisWorkbook.Save
End Sub

