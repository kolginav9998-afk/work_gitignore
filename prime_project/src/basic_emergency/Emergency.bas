Option Explicit

' ПОКАТАК PRIME — EMERGENCY WORKING VERSION
' По прямому указанию пользователя: "Останови дальнейшее усложнение PRIME... не строй ради
' этого новый сложный transaction engine". Это НЕ часть основного движка PRIME_00..17 (тот
' заморожен и не тронут ни строкой в этом проходе) - отдельный, маленький, полностью
' независимый модуль на 6 листах, без скрытых DB_PRIME_*/SYS_PRIME_* таблиц, без транзакционного
' протокола PREPARED/COMMITTED, без событий OnChange (главный источник самых тяжёлых дефектов
' во всём предыдущем движке - buffered-array corruption, event guard, реентерабельность). Вся
' арифметика остатка - живые формулы Calc (SUMIFS), не движения в отдельной таблице. Каждая
' Sub здесь вызывается ТОЛЬКО явным кликом по кнопке - никогда из события листа.
'
' Формула остатка (везде одна и та же идея):
'   Остаток = Пришло (по этой строке/ЕИ) - SUMIFS(Выдачи!Количество, где ЕИ=этот, Статус="Проведено")
' Это значит: пока строка "Выдачи" не помечена явно как "Проведено" макросом Issues_ConductAllButton
' (после проверки остатка), она не уменьшает Остаток - "Провести всё" - единственное место,
' где может появиться "Проведено".

Public Const SH_ORDERS As String = "Заказы"
Public Const SH_OFFICE As String = "Приход — Офис"
Public Const SH_PROD As String = "Приход — Производство"
Public Const SH_PARTS As String = "Приход — Детали"
Public Const SH_ISSUE As String = "Выдачи"
Public Const SH_STOCK As String = "Наличие"
Public Const SH_SEQ As String = "_SEQ"

' Число предзаполненных строк с формулами на каждом листе - строки за этой границей нужно
' будет протянуть вручную (Правка -> Заполнить -> Вниз), как в обычной таблице Calc.
Public Const EMERGENCY_MAX_ROW As Long = 999

Private Function GetSheet(ByVal sheetName As String) As Object
    GetSheet = ThisComponent.Sheets.getByName(sheetName)
End Function

' === Генератор ЕИ-кода ==========================================================================
' Простой персистентный счётчик в скрытом листе "_SEQ" (ячейка A1) - НЕ PRIME_SequenceNext,
' никакого PREPARED/COMMITTED. Гарантия: код никогда не переиспользуется (счётчик только растёт).
' Не гарантирует отсутствие "дыр" при одновременном редактировании двумя людьми одновременно -
' для аварийной однопользовательской версии это осознанно приемлемый компромисс.
Private Function NextEiCode() As String
    Dim oSeq As Object
    oSeq = GetSheet(SH_SEQ)
    Dim n As Long
    n = CLng(oSeq.getCellByPosition(0, 0).getValue())
    n = n + 1
    oSeq.getCellByPosition(0, 0).setValue(n)
    NextEiCode = "ЕИ-" & Format(n, "00000000")
End Function

' === Провести всё: Заказы =======================================================================
' Колонки (0-индекс): 0 Номер заказа,1 Наименование,2 Количество,3 Факт. количество,4 Ед. изм.,
' 5 Место,6 Дата,7 Комментарий,8 ЕИ-код,9 Остаток(формула),10 Статус.
' Заказ сам по себе не приход: строка получает ЕИ ТОЛЬКО когда "Факт. количество" заполнено и
' положительно. Если один заказ пришёл двумя частями - пользователь добавляет ВТОРУЮ строку с
' тем же "Номер заказа" и своим "Факт. количество" - она получает свой отдельный ЕИ. Никакого
' отдельного механизма дочерних строк.
Public Sub Orders_ConductAllButton()
    Dim oSheet As Object
    oSheet = GetSheet(SH_ORDERS)
    Dim r As Long, n As Long
    n = 0
    For r = 1 To EMERGENCY_MAX_ROW
        Dim eiCell As Object, factCell As Object
        eiCell = oSheet.getCellByPosition(8, r)
        factCell = oSheet.getCellByPosition(3, r)
        If eiCell.getString() = "" And factCell.getValue() > 0 Then
            eiCell.setString(NextEiCode())
            oSheet.getCellByPosition(10, r).setString("Приход оформлен")
            n = n + 1
        End If
    Next r
    MsgBox "Заказы: оформлено новых приходов - " & n
End Sub

' === Провести всё: Приход — Офис / Производство / Детали =======================================
' Офис/Производство: 0 ЕИ-код,1 Наименование,2 Пришло,3 Ед. изм.,4 Остаток,5 Место,6 Дата,
' 7 Комментарий.
' Детали (доп. Артикул после Наименования): 0 ЕИ-код,1 Наименование,2 Артикул,3 Пришло,
' 4 Ед. изм.,5 Остаток,6 Место,7 Дата,8 Комментарий.
Private Sub ConductReceiptSheet(ByVal sheetName As String, ByVal colQty As Long)
    Dim oSheet As Object
    oSheet = GetSheet(sheetName)
    Dim r As Long, n As Long
    n = 0
    For r = 1 To EMERGENCY_MAX_ROW
        Dim eiCell As Object, qtyCell As Object
        eiCell = oSheet.getCellByPosition(0, r)
        qtyCell = oSheet.getCellByPosition(colQty, r)
        If eiCell.getString() = "" And qtyCell.getValue() > 0 Then
            eiCell.setString(NextEiCode())
            n = n + 1
        End If
    Next r
    MsgBox sheetName & ": оформлено новых приходов - " & n
End Sub

Public Sub Office_ConductAllButton()
    ConductReceiptSheet(SH_OFFICE, 2)
End Sub

Public Sub Production_ConductAllButton()
    ConductReceiptSheet(SH_PROD, 2)
End Sub

Public Sub Parts_ConductAllButton()
    ConductReceiptSheet(SH_PARTS, 3)
End Sub

' === Поиск ЕИ по всем 4 приходным листам ========================================================
Private Function FindEi(ByVal eiCode As String, ByRef outName As String, ByRef outUnit As String, _
        ByRef outBalance As Double, ByRef outSheet As String) As Boolean
    Dim sheets(3) As String
    sheets(0) = SH_ORDERS
    sheets(1) = SH_OFFICE
    sheets(2) = SH_PROD
    sheets(3) = SH_PARTS

    Dim i As Long
    For i = 0 To 3
        Dim oSheet As Object
        oSheet = GetSheet(sheets(i))
        Dim colEi As Long, colName As Long, colUnit As Long, colBalance As Long
        Select Case sheets(i)
            Case SH_ORDERS
                colEi = 8 : colName = 1 : colUnit = 4 : colBalance = 9
            Case SH_PARTS
                colEi = 0 : colName = 1 : colUnit = 4 : colBalance = 5
            Case Else ' Офис / Производство - одинаковая раскладка
                colEi = 0 : colName = 1 : colUnit = 3 : colBalance = 4
        End Select

        Dim r As Long
        For r = 1 To EMERGENCY_MAX_ROW
            If oSheet.getCellByPosition(colEi, r).getString() = eiCode Then
                outName = oSheet.getCellByPosition(colName, r).getString()
                outUnit = oSheet.getCellByPosition(colUnit, r).getString()
                outBalance = oSheet.getCellByPosition(colBalance, r).getValue()
                outSheet = sheets(i)
                FindEi = True
                Exit Function
            End If
        Next r
    Next i
    FindEi = False
End Function

' === Выдачи: автозаполнение + проверка остатка + проведение ====================================
' 0 Дата,1 ЕИ-код,2 Наименование,3 Количество,4 Ед. изм.,5 Доступно,6 Кто получил,7 Куда,
' 8 Возвратный,9 Комментарий,10 Статус.
' Единственное место во всей аварийной версии, где остаток может "уменьшиться" - выставляя
' Статус="Проведено" ПОСЛЕ проверки, что Количество не превышает текущий Остаток той приходной
' строки. Если не хватает - строка НЕ проводится, Статус получает конкретное сообщение с числом
' доступного остатка, остаток не меняется.
Public Sub Issues_ConductAllButton()
    Dim oSheet As Object
    oSheet = GetSheet(SH_ISSUE)
    Dim r As Long, posted As Long, blocked As Long
    posted = 0
    blocked = 0
    Dim blockMsg As String
    blockMsg = ""

    For r = 1 To EMERGENCY_MAX_ROW
        Dim eiCode As String
        eiCode = Trim(oSheet.getCellByPosition(1, r).getString())
        Dim status As String
        status = oSheet.getCellByPosition(10, r).getString()
        If eiCode <> "" And status <> "Проведено" Then
            Dim outName As String, outUnit As String, outSheet As String
            Dim outBalance As Double
            If Not FindEi(eiCode, outName, outUnit, outBalance, outSheet) Then
                oSheet.getCellByPosition(10, r).setString("Ошибка: ЕИ не найден")
                blocked = blocked + 1
                blockMsg = blockMsg & "Строка " & (r + 1) & ": ЕИ " & eiCode & " не найден ни на одном приходном листе." & Chr(10)
            Else
                If oSheet.getCellByPosition(2, r).getString() = "" Then oSheet.getCellByPosition(2, r).setString(outName)
                If oSheet.getCellByPosition(4, r).getString() = "" Then oSheet.getCellByPosition(4, r).setString(outUnit)
                oSheet.getCellByPosition(5, r).setValue(outBalance)

                Dim qty As Double
                qty = oSheet.getCellByPosition(3, r).getValue()
                If qty <= 0 Then
                    oSheet.getCellByPosition(10, r).setString("Ошибка: некорректное количество")
                    blocked = blocked + 1
                    blockMsg = blockMsg & "Строка " & (r + 1) & ": количество должно быть больше нуля." & Chr(10)
                ElseIf qty > outBalance + 0.0000005 Then
                    oSheet.getCellByPosition(10, r).setString("Недостаточно остатка по " & eiCode & ". Доступно: " & outBalance)
                    blocked = blocked + 1
                    blockMsg = blockMsg & "Строка " & (r + 1) & ": недостаточно остатка по " & eiCode & ". Доступно: " & outBalance & "." & Chr(10)
                Else
                    oSheet.getCellByPosition(10, r).setString("Проведено")
                    posted = posted + 1
                End If
            End If
        End If
    Next r

    Dim resultMsg As String
    resultMsg = "Выдачи: проведено - " & posted
    If blocked > 0 Then resultMsg = resultMsg & Chr(10) & "Заблокировано - " & blocked & ":" & Chr(10) & blockMsg
    MsgBox resultMsg
End Sub

' === Наличие: пересобрать общий обзор ===========================================================
' 0 ЕИ-код,1 Наименование,2 Откуда пришло,3 Пришло,4 Выдано,5 Остаток,6 Место.
' Только просмотр - не отдельный склад и ничего самостоятельно не считает, кроме данных
' четырёх приходных листов и "Выдачи" (Остаток здесь - значение уже посчитанной формулы на
' исходном приходном листе, не отдельный расчёт).
Public Sub Stock_RefreshButton()
    Dim oOut As Object
    oOut = GetSheet(SH_STOCK)

    Dim r As Long, c As Long
    For r = 1 To EMERGENCY_MAX_ROW
        For c = 0 To 6
            oOut.getCellByPosition(c, r).setString("")
        Next c
    Next r

    Dim sheets(3) As String
    sheets(0) = SH_ORDERS
    sheets(1) = SH_OFFICE
    sheets(2) = SH_PROD
    sheets(3) = SH_PARTS
    Dim outRow As Long
    outRow = 1

    Dim i As Long
    For i = 0 To 3
        Dim oSheet As Object
        oSheet = GetSheet(sheets(i))
        Dim colEi As Long, colName As Long, colQty As Long, colBalance As Long, colLoc As Long
        Select Case sheets(i)
            Case SH_ORDERS
                colEi = 8 : colName = 1 : colQty = 3 : colBalance = 9 : colLoc = 5
            Case SH_PARTS
                colEi = 0 : colName = 1 : colQty = 3 : colBalance = 5 : colLoc = 6
            Case Else
                colEi = 0 : colName = 1 : colQty = 2 : colBalance = 4 : colLoc = 5
        End Select

        Dim r2 As Long
        For r2 = 1 To EMERGENCY_MAX_ROW
            Dim eiCode As String
            eiCode = oSheet.getCellByPosition(colEi, r2).getString()
            If eiCode <> "" Then
                Dim qty As Double, balance As Double
                qty = oSheet.getCellByPosition(colQty, r2).getValue()
                balance = oSheet.getCellByPosition(colBalance, r2).getValue()
                oOut.getCellByPosition(0, outRow).setString(eiCode)
                oOut.getCellByPosition(1, outRow).setString(oSheet.getCellByPosition(colName, r2).getString())
                oOut.getCellByPosition(2, outRow).setString(sheets(i))
                oOut.getCellByPosition(3, outRow).setValue(qty)
                oOut.getCellByPosition(4, outRow).setValue(qty - balance)
                oOut.getCellByPosition(5, outRow).setValue(balance)
                oOut.getCellByPosition(6, outRow).setString(oSheet.getCellByPosition(colLoc, r2).getString())
                outRow = outRow + 1
            End If
        Next r2
    Next i

    MsgBox "Наличие обновлено: " & (outRow - 1) & " позиций."
End Sub
