# Excel-VBA
У сучасному світі, де інформаційні технології швидко розвиваються, автоматизація рутинних процесів стає надзвичайно важливою задачею. Одним із таких процесів є відправка повідомлень студентам університету щодо виборчих дисциплін у навчальному процесі. Велика кількість студентів, різноманітні зміни в дисциплінах, перевибір дисциплін - все це вимагає великих зусиль та часу від відповідальних осіб.

Ця робота спрямована на розробку автоматизованого процесу відправки повідомлень студентам за допомогою макросу в Microsoft Excel. Головна мета полягає в ефективному використанні часу та ресурсів, швидкому і точному повідомленні студентів про вибір та зміни у виборчих дисциплінах.

Перед початком роботи було проведено детальний аналіз вимог та функціональних можливостей. На основі цього аналізу була визначена архітектура макросу та визначені основні етапи роботи. Відповідно до цих етапів було розроблено програмну частину макросу, яка включає в себе вибір студентів, формування та відправлення електронного листа, а також вибір студентів без електронної пошти та студентів, які перевибрали дисципліни.

Отримані результати дозволяють значно полегшити та прискорити процес відправки повідомлень студентам. Макрос дозволяє автоматично визначати адресатів та відправляти індивідуальні листи студентам з відповідними змінами у навчальному процесі. Крім того, розроблений макрос надійний та ефективний, забезпечуючи точність та швидкість виконання.

У файлі за посиланням в розділі About будуть детально розглянуті процес розробки макросу та його функціональні можливості. 

Фрагменти коду:

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim wsInput As Worksheet
    Dim wsChanges As Worksheet
    Dim outputRow As Long
    Dim student As String
    Dim group As String
    Dim subjectCode As String
    Dim subjectName As String
    Dim email As String
    
    ' Укажите листы, на которых хранятся данные
    Set wsInput = ThisWorkbook.Sheets("Sheet1") ' Лист1
    Set wsChanges = ThisWorkbook.Sheets("Sheet3") ' Лист3
    
    ' Проверяем, что двойной щелчок произошел в столбце C (код дисциплины)
    If Target.Column = 3 Then
        Application.EnableEvents = False ' Отключаем обработку событий, чтобы избежать рекурсивных изменений ячеек
        
        ' Получаем данные из щелкнутой ячейки и соответствующих столбцов
        student = Target.Offset(0, -2).Value ' ФИО студента (столбец A)
        group = Target.Offset(0, -1).Value ' Группа (столбец B)
        subjectCode = Target.Value ' Код дисциплины (столбец C)
        subjectName = Target.Offset(0, 1).Value ' Название дисциплины (столбец D)
        email = Target.Offset(0, 2).Value ' Электронная почта (столбец H)
        
        ' Проверяем, существует ли уже такой код дисциплины на "Листе3"
        If WorksheetFunction.CountIf(wsChanges.Columns(3), subjectCode) = 0 Then
            ' Код дисциплины не найден на "Листе3", выполняем запись
            outputRow = wsChanges.Cells(wsChanges.Rows.Count, 1).End(xlUp).Row + 1
            
            ' Записываем данные на "Лист3"
            wsChanges.Cells(outputRow, 1).Value = student ' ФИО студента
            wsChanges.Cells(outputRow, 2).Value = group ' Группа
            wsChanges.Cells(outputRow, 3).Value = subjectCode ' Код дисциплины
            wsChanges.Cells(outputRow, 4).Value = subjectName ' Название дисциплины
            wsChanges.Cells(outputRow, 9).Value = email ' Электронная почта
        End If
        
        Application.EnableEvents = True ' Включаем обработку событий
    End If
End Sub

Sub ОтправитьДисциплиныНаEmail()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim dict As Object
    Dim student As Variant
    Dim email As String
    Dim subject As String
    Dim message As String
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim wsOutput As Worksheet ' Лист для вывода данных, "Лист2"
    Dim outputRow As Long
    Dim studentsArray As Variant ' Вспомогательный массив для хранения уже записанных студентов
    Dim semester As String
    Dim nakaz As String
     
    ' Создаем словарь для хранения дисциплин каждого студента
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Указываем имя вашего листа
    Set ws = ThisWorkbook.Worksheets("Sheet1")
    
    ' Указываем лист для вывода данных
    Set wsOutput = ThisWorkbook.Worksheets("Sheet2")
    
    ' Находим последнюю заполненную строку в первом столбце
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    
    semester = ws.Range("G1").Value
    nakaz = ws.Range("H1").Value
    
    ' Проходимся по каждой строке и собираем дисциплины для каждого студента
    For i = 2 To lastRow
        student = ws.Cells(i, 1).Value
        subjectCode = ws.Cells(i, 3).Value
        subjectName = ws.Cells(i, 4).Value
        email = ws.Cells(i, 5).Value
        
        ' Проверяем, есть ли уже такой студент в словаре
        If dict.Exists(student) Then
            ' Если студент уже существует, добавляем новую дисциплину к списку
            dict(student) = dict(student) & vbCrLf & subjectCode & ": " & subjectName
        Else
            ' Если студента еще нет в словаре
            If email = "" Then ' Проверяем пустую ячейку с почтой
                ' Проверяем, был ли студент уже записан на "Лист2"
                If Not IsStudentAlreadyWritten(student, wsOutput) Then
                    ' Записываем данные студента на "Лист2"
                    outputRow = wsOutput.Cells(wsOutput.Rows.Count, 1).End(xlUp).Row + 1
                    wsOutput.Cells(outputRow, 1).Value = student
                    wsOutput.Cells(outputRow, 2).Value = ws.Cells(i, 2).Value ' Группа
                    wsOutput.Cells(outputRow, 3).Value = subjectCode ' Код дисциплины
                    wsOutput.Cells(outputRow, 4).Value = subjectName ' Название дисциплины
                End If
            Else
                ' Добавляем студента и его дисциплину в словарь
                dict(student) = subjectCode & ": " & subjectName
            End If
        End If
    Next i
    
    ' Инициализируем объект Outlook
    Set outlookApp = CreateObject("Outlook.Application")
    
    ' Отправляем письма с дисциплинами каждому студенту
    For Each student In dict.Keys
        ' Получаем email студента
        email = ws.Cells(Application.Match(student, ws.Columns(1), 0), 5).Value
        
        ' Создаем новое письмо
        Set outlookMail = outlookApp.CreateItem(0)
        
        ' Формируем тему письма
        subject = "Вибіркові дисципліни " & student
        
        ' Формируем текст письма
message = "Добрий день, " & student & "!" & vbCrLf & vbCrLf & _
               "Повідомляємо Вас, що на  " & semester & " відповідно до наказу " & nakaz & " до Вашого індивідуального навчального плану входять такі дисципліни за вибором:" & vbCrLf & vbCrLf & _
               dict(student) & vbCrLf & vbCrLf & _
               "Якщо у Вас виникли питання з обраних дисциплін за вибором необхідно звертатись до куратора Вашої академічної групи або завідувача кафедри (якщо Ви є студентом заочної або вечірньої форми навчання – до своїх методистів)." & vbCrLf & vbCrLf & _
               "Вашої відповіді цей лист не передбачає і розглядатись не буде." & vbCrLf & vbCrLf & _
               "З повагою, навчальний відділ."
        ' Заполняем поля письма
        With outlookMail
            .To = email
            .subject = subject
            .body = message
            .Send
        End With
        
        ' Освобождаем ресурсы объекта письма
        Set outlookMail = Nothing
    Next student
    
    ' Освобождаем ресурсы объекта Outlook
    Set outlookApp = Nothing
    
    MsgBox "Письма успешно отправлены!"
End Sub

Function IsStudentAlreadyWritten(ByVal student As String, ByVal wsOutput As Worksheet) As Boolean
    Dim lastRow As Long
    
    ' Находим последнюю заполненную строку в первом столбце на "Лист2"
    lastRow = wsOutput.Cells(wsOutput.Rows.Count, 1).End(xlUp).Row
    
    ' Проходимся по каждой строке на "Лист2" и проверяем, есть ли уже такой студент
    For i = 2 To lastRow
        If wsOutput.Cells(i, 1).Value = student Then
            IsStudentAlreadyWritten = True ' Студент уже записан
            Exit Function
        End If
    Next i
    
    IsStudentAlreadyWritten = False ' Студент еще не записан
End Function

Sub SendEmails()
    Dim wsChanges As Worksheet
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim lastRow As Long
    Dim recipient As Variant
    Dim subject As String
    Dim body As String
    Dim i As Long
    Dim semester As String
    Dim nakaz As String
    
    ' Указываем лист, на котором хранятся изменения
    Set wsChanges = ThisWorkbook.Sheets("Sheet3") ' Лист3
    
    semester = wsChanges.Range("J1").Value
    nakaz = wsChanges.Range("G2").Value
    
    ' Создаем объект Outlook
    Set outlookApp = CreateObject("Outlook.Application")
    
    ' Определяем последнюю строку на "Лист3"
    lastRow = wsChanges.Cells(wsChanges.Rows.Count, 1).End(xlUp).Row
    
    ' Словарь для хранения изменений по каждому студенту
    Dim changesDict As Object
    Set changesDict = CreateObject("Scripting.Dictionary")
    
    ' Цикл для сбора изменений по каждому студенту
    For i = 2 To lastRow ' Начинаем с 2-й строки, так как первая содержит заголовки столбцов
        ' Получаем данные для текущего студента
        recipient = wsChanges.Cells(i, 9).Value ' Email студента
        subject = "Изменение дисциплины" ' Тема письма

        ' Проверяем, есть ли уже изменения для этого студента
        If changesDict.Exists(recipient) Then
            ' Добавляем текущее изменение к существующим
            body = body & vbCrLf & vbCrLf & _
                   "Замість дисципліни " & wsChanges.Cells(i, 3).Value & " - " & wsChanges.Cells(i, 4).Value & vbCrLf & _
                   "повинна бути внесена така дисципліна за вибором - " & wsChanges.Cells(i, 5).Value & " - " & wsChanges.Cells(i, 6).Value & vbCrLf & _
                   "Зміна відбулась з причини – " & wsChanges.Cells(i, 8).Value
            ' Добавляем разделитель между изменениями
            changesDict(recipient) = changesDict(recipient) & vbCrLf & vbCrLf & body
        Else
            ' Создаем новую запись для студента
             body = "Добрий день, " & wsChanges.Cells(i, 1).Value & "!" & vbCrLf & vbCrLf & _
               "Повідомляємо Вас, що на " & semester & " навчального року відповідно до наказу №" & nakaz & " до Вашого індивідуального навчального плану:" & vbCrLf & vbCrLf & _
               "Замість дисципліни " & wsChanges.Cells(i, 3).Value & " - " & wsChanges.Cells(i, 4).Value & vbCrLf & vbCrLf & _
               "повинна бути внесена така дисципліна за вибором - " & wsChanges.Cells(i, 5).Value & " - " & wsChanges.Cells(i, 6).Value & vbCrLf & vbCrLf & _
               "Зміна відбулась з причини – " & wsChanges.Cells(i, 8).Value & vbCrLf & vbCrLf & _
               "Якщо у Вас виникли питання з обраних дисциплін за вибором, необхідно звертатись до куратора Вашої академічної групи або завідувача кафедри." & vbCrLf & vbCrLf & _
               "Вашої відповіді цей лист не передбачає і розглядатись не буде." & vbCrLf & vbCrLf & _
               "З повагою," & vbCrLf & _
               "Навчальний відділ."
            changesDict.Add recipient, body
        End If
    Next i
    
    ' Отправка писем каждому студенту
    For Each recipient In changesDict
        ' Создаем новое письмо
        Set outlookMail = outlookApp.CreateItem(0)
        
        ' Заполняем поля письма
        With outlookMail
            .To = recipient
            .subject = subject
            .body = changesDict(recipient)
            .Send ' Отправляем письмо без отображения для предварительного просмотра
        End With
        
        ' Освобождаем ресурсы для текущего письма
        Set outlookMail = Nothing
    Next recipient
    
    ' Освобождаем ресурсы Outlook
    Set outlookApp = Nothing
    
    MsgBox "Письма отправлены.", vbInformation
End Sub

