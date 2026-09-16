Option Explicit

' Internal candidate. LibreOffice Basic/UNO, Linux offline.
' Owns only two new sheets; never posts warehouse movements.
Global Const WMSRPT_VERSION = "0.4.0"
Global gWMSRPT_RefreshSucceeded As Boolean
Global Const WMSRPT_INPUT = "Отчет — ввод"
Global Const WMSRPT_OUTPUT = "Отчет руководителю"
Global Const WMSRPT_MARKER = "POKATAK_MANAGER_REPORT_1"
Global Const WMSRPT_FIRST = 15

Sub WMSRPT_Install()
    Dim doc As Object,inp As Object,out As Object,nm As Variant,sh As Object
    Dim labels As Variant,actions As Variant,i As Long,formObj As Object
    doc=ThisComponent
    On Error GoTo EH
    ' Validate both destinations before touching either one.
    For Each nm In Array(WMSRPT_INPUT,WMSRPT_OUTPUT)
        If doc.Sheets.hasByName(CStr(nm)) Then
            sh=doc.Sheets.getByName(CStr(nm))
            If sh.getCellByPosition(25,0).String<>WMSRPT_MARKER Then
                MsgBox "Лист «" & CStr(nm) & "» уже существует и не принадлежит модулю отчёта. Данные сохранены.",48,"WMS":Exit Sub
            End If
        End If
    Next nm
    doc.lockControllers()
    For Each nm In Array(WMSRPT_INPUT,WMSRPT_OUTPUT)
        If Not doc.Sheets.hasByName(CStr(nm)) Then
            doc.Sheets.insertNewByName(CStr(nm),doc.Sheets.Count)
            sh=doc.Sheets.getByName(CStr(nm))
            sh.getCellByPosition(25,0).String=WMSRPT_MARKER
        End If
    Next nm
    inp=doc.Sheets.getByName(WMSRPT_INPUT):out=doc.Sheets.getByName(WMSRPT_OUTPUT)
    inp.getCellByPosition(0,0).String="РАБОЧИЙ ДЕНЬ И ОТЧЁТ"
    inp.getCellByPosition(0,2).String="Дата отчёта"
    inp.getCellByPosition(0,3).String="Автор"
    inp.getCellByPosition(0,4).String="Следующий рабочий день"
    inp.getCellByPosition(0,6).String="Состояние сводки"
    inp.getCellByPosition(0,7).String="Факты из базы"
    If inp.getCellByPosition(1,2).Type=0 Then inp.getCellByPosition(1,2).Value=CDbl(Date)
    If inp.getCellByPosition(1,3).Type=0 Then inp.getCellByPosition(1,3).String="Егор, кладовщик."
    If inp.getCellByPosition(1,4).Type=0 Then inp.getCellByPosition(1,4).Formula="=WORKDAY(B3;1)"
    If inp.getCellByPosition(1,6).Type=0 Then inp.getCellByPosition(1,6).String="Сводка ещё не загружена"
    inp.getCellByPosition(0,9).String="1. Проверьте дату. 2. Отметьте помощь, переговоры и задачи кнопками. 3. Нажмите «Сформировать отчёт». Приходы и выдачи программа соберёт сама."
    labels=Array("Дата","Включить","Событие","Кому / с кем","Заказ / предмет","Результат / действие","Свой текст")
    For i=0 To UBound(labels):inp.getCellByPosition(i,14).String=labels(i):Next i
    inp.getCellByPosition(0,5).String="Контроль заказов"
    If inp.getCellByPosition(1,5).Type=0 Then inp.getCellByPosition(1,5).String="Да"
    inp.getCellByPosition(0,8).String="Формат отчёта"
    If inp.getCellByPosition(1,8).Type=0 Then inp.getCellByPosition(1,8).String="Подробный"
    WMSRPT_Style doc,inp,False
    WMSRPT_Style doc,out,True
    out.getCellByPosition(0,0).String="ОТЧЁТ РУКОВОДИТЕЛЮ"
    out.getCellByPosition(0,1).String="Сформируйте текст на листе «Отчет — ввод». Текст можно выделить и скопировать обычным Ctrl+C."
    If out.getCellByPosition(0,3).Type=0 Then out.getCellByPosition(0,3).String="Отчёт ещё не сформирован."
    labels=Array("Принята поставка","Разгрузка погрузчиком","Помощь погрузчиком","Бухгалтерия","Закупки","Задача на завтра","В процессе","Обновить факты","Сформировать отчёт","Сохранить итог","Документы","Расхождения","Организация склада")
    actions=Array("WMSRPT_Delivery","WMSRPT_Unload","WMSRPT_Assist","WMSRPT_Accounts","WMSRPT_Procurement","WMSRPT_Task","WMSRPT_Ongoing","WMSRPT_RefreshFacts","WMSRPT_Build","WMSRPT_Save","WMSRPT_Documents","WMSRPT_Discrepancy","WMSRPT_Storage")
    formObj=WMSRPT_Form(doc,inp)
    For i=0 To UBound(labels)
        WMSRPT_Button doc,inp,formObj,i,CStr(labels(i)),CStr(actions(i))
    Next i
    inp.Columns.getByIndex(25).IsVisible=False
    out.Columns.getByIndex(25).IsVisible=False
Done:
    On Error Resume Next:doc.unlockControllers():On Error GoTo 0
    Exit Sub
EH:
    MsgBox "Настройка отчёта не завершена: " & CStr(Err) & " " & Error$,16,"WMS"
    Resume Done
End Sub

Sub WMSRPT_Style(doc As Object,sh As Object,isOutput As Boolean)
    Dim rng As Object,i As Long,widths As Variant,fmt As Long,loc As New com.sun.star.lang.Locale
    rng=sh.getCellRangeByPosition(0,0,6,214)
    rng.CharFontName="Liberation Sans":rng.CharHeight=10:rng.CellBackColor=16777215
    sh.getCellRangeByPosition(0,0,6,0).CellBackColor=2240061
    sh.getCellRangeByPosition(0,0,6,0).CharColor=16777215
    sh.getCellRangeByPosition(0,0,6,0).CharWeight=150
    sh.Rows.getByIndex(0).Height=1100
    sh.Rows.getByIndex(5).Height=1100:sh.Rows.getByIndex(8).Height=1100
    sh.Rows.getByIndex(9).Height=1100
    sh.getCellRangeByPosition(0,9,6,9).IsTextWrapped=True
    sh.Rows.getByIndex(11).Height=1100:sh.Rows.getByIndex(12).Height=1100:sh.Rows.getByIndex(13).Height=1100
    If isOutput Then
        sh.Columns.getByIndex(0).Width=19000
        sh.getCellByPosition(0,1).IsTextWrapped=True
        sh.Rows.getByIndex(1).Height=1300
        sh.getCellByPosition(0,3).IsTextWrapped=True
        sh.Rows.getByIndex(3).Height=18000
        sh.getCellByPosition(0,3).CharHeight=12
    Else
        widths=Array(3400,6400,4500,4500,4800,6000,6000)
        For i=0 To UBound(widths):sh.Columns.getByIndex(i).Width=CLng(widths(i)):Next i
        sh.getCellRangeByPosition(0,14,6,14).CellBackColor=2240061
        sh.getCellRangeByPosition(0,14,6,14).CharColor=16777215
        sh.getCellRangeByPosition(0,14,6,14).CharWeight=150
        sh.Rows.getByIndex(14).Height=850
        sh.getCellRangeByPosition(0,15,6,214).IsTextWrapped=True
        sh.getCellRangeByPosition(1,2,1,4).CellBackColor=15794160
        sh.getCellByPosition(1,7).IsTextWrapped=True
        sh.Rows.getByIndex(7).Height=5000
        loc.Language="ru":loc.Country="RU"
        fmt=doc.NumberFormats.queryKey("DD.MM.YYYY",loc,True)
        If fmt=-1 Then fmt=doc.NumberFormats.addNew("DD.MM.YYYY",loc)
        sh.getCellByPosition(1,2).NumberFormat=fmt
        sh.getCellByPosition(1,4).NumberFormat=fmt
        sh.getCellRangeByPosition(0,15,0,214).NumberFormat=fmt
    End If
End Sub

Function WMSRPT_Form(doc As Object,sh As Object) As Object
    Dim f As Object
    If sh.DrawPage.Forms.hasByName("WMSRPT_FORM") Then
        f=sh.DrawPage.Forms.getByName("WMSRPT_FORM")
    Else
        f=doc.createInstance("com.sun.star.form.component.Form")
        sh.DrawPage.Forms.insertByName("WMSRPT_FORM",f)
    End If
    WMSRPT_Form=f
End Function

Sub WMSRPT_Button(doc As Object,sh As Object,f As Object,i As Long,caption As String,proc As String)
    Dim ctl As Object,shape As Object,pos As Object,sz As Object,nm As String,idx As Long
    Dim ev As New com.sun.star.script.ScriptEventDescriptor
    nm="WMSRPT_BTN_" & CStr(i)
    If f.hasByName(nm) Then
        Dim j As Long,existingShape As Object
        For j=0 To sh.DrawPage.Count-1
            existingShape=sh.DrawPage.getByIndex(j)
            If existingShape.supportsService("com.sun.star.drawing.ControlShape") Then
                If existingShape.Control.Name=nm Then
                    pos=CreateUnoStruct("com.sun.star.awt.Point")
                    pos.X=100+(i Mod 5)*7000:pos.Y=sh.getCellByPosition(0,11+(i\5)).Position.Y+100
                    sz=CreateUnoStruct("com.sun.star.awt.Size"):sz.Width=6500:sz.Height=850
                    existingShape.Position=pos:existingShape.Size=sz
                    Exit Sub
                End If
            End If
        Next j
        Exit Sub
    End If
    ctl=doc.createInstance("com.sun.star.form.component.CommandButton")
    ctl.Name=nm:ctl.Label=caption
    idx=f.Count:f.insertByName(nm,ctl)
    shape=doc.createInstance("com.sun.star.drawing.ControlShape"):shape.Control=ctl
    pos=CreateUnoStruct("com.sun.star.awt.Point")
    pos.X=100+(i Mod 5)*7000:pos.Y=sh.getCellByPosition(0,11+(i\5)).Position.Y
    sz=CreateUnoStruct("com.sun.star.awt.Size"):sz.Width=6500:sz.Height=850
    shape.Position=pos:shape.Size=sz:sh.DrawPage.add(shape)
    ev.ListenerType="com.sun.star.awt.XActionListener":ev.EventMethod="actionPerformed"
    ev.ScriptType="Script":ev.ScriptCode="vnd.sun.star.script:Standard.WMS_34_ManagerReport." & proc & "?language=Basic&location=document"
    f.registerScriptEvent idx,ev
End Sub

Sub WMSRPT_Delivery()
    WMSRPT_AddEvent "Принята поставка"
End Sub
Sub WMSRPT_Unload()
    WMSRPT_AddEvent "Разгрузка погрузчиком"
End Sub
Sub WMSRPT_Assist()
    WMSRPT_AddEvent "Помощь погрузчиком"
End Sub
Sub WMSRPT_Accounts()
    WMSRPT_AddEvent "Бухгалтерия"
End Sub
Sub WMSRPT_Procurement()
    WMSRPT_AddEvent "Закупки"
End Sub
Sub WMSRPT_Task()
    WMSRPT_AddEvent "Задача на завтра"
End Sub
Sub WMSRPT_Ongoing()
    WMSRPT_AddEvent "В процессе"
End Sub

Sub WMSRPT_AddEvent(kind As String)
    Dim sh As Object,r As Long,who As String,subject As String,resultText As String
    On Error GoTo EH
    sh=ThisComponent.Sheets.getByName(WMSRPT_INPUT)
    If Not WMSRPT_ValidDay(sh.getCellByPosition(1,2).Value) Then
        MsgBox "Сначала укажите корректную дату отчёта.",48,"ПОКАТАК":Exit Sub
    End If
    who=InputBox("Кому помогли или с кем взаимодействовали? Можно оставить пустым.",kind)
    subject=InputBox("По какому вопросу, заказу или товару? Можно оставить пустым.",kind)
    resultText=InputBox("Что конкретно выполнено или согласовано? Пустой ответ отменяет добавление.",kind)
    If Trim(resultText)="" Then Exit Sub
    r=WMSCore_LastContentRow(sh,6)+1:If r<WMSRPT_FIRST Then r=WMSRPT_FIRST
    sh.getCellRangeByPosition(0,r,6,r).setDataArray(Array(Array(sh.getCellByPosition(1,2).Value,"Да",kind,who,subject,resultText,"")))
    sh.getCellByPosition(0,r).NumberFormat=sh.getCellByPosition(1,2).NumberFormat
    ThisComponent.CurrentController.setActiveSheet(sh)
    ThisComponent.CurrentController.select(sh.getCellByPosition(3,r))
    Exit Sub
EH:
    MsgBox "Событие не добавлено: " & CStr(Err) & " " & Error$,16,"WMS"
End Sub

Sub WMSRPT_RefreshFacts()
    Dim sh As Object,con As Object,st As Object,rs As Object,sErr As String,dayValue As Double,sql As String,summary As String,n As Long
    gWMSRPT_RefreshSucceeded=False
    If Not WMSDBX_TryEnter("MANAGER_REPORT") Then MsgBox "Дождитесь завершения складской операции.",48,"WMS":Exit Sub
    On Error GoTo EH
    sh=ThisComponent.Sheets.getByName(WMSRPT_INPUT)
    dayValue=sh.getCellByPosition(1,2).Value
    If Not WMSRPT_SettingsValid(sh,sErr) Then GoTo Failed
    If Not WMSRPT_ValidDay(dayValue) Then sErr="Укажите корректную дату отчёта без времени.":GoTo Failed
    sh.getCellByPosition(25,1).Value=0
    sh.getCellByPosition(1,6).String="Обновление не завершено"
    con=WMSDB_GetConnectionEx(ThisComponent,sErr):If sErr<>"" Then GoTo Failed
    summary=WMSRPT_OperationalSummary(con,dayValue) & WMSRPT_Recipients(con,dayValue)
    If UCase(Trim(sh.getCellByPosition(1,8).String))<>"КРАТКИЙ" Then
        summary=summary & Chr(10) & WMSD_ReportDetails(con,dayValue)
    End If
    sh.getCellByPosition(25,2).String=""
    sh.getCellByPosition(25,3).String=""
    If UCase(Trim(sh.getCellByPosition(1,5).String))<>"НЕТ" Then
        sh.getCellByPosition(25,2).String=WMSRPT_ControlSnapshot(sh)
    End If
    sh.getCellByPosition(1,7).String=summary
    sh.getCellByPosition(25,1).Value=dayValue
    gWMSRPT_RefreshSucceeded=True
    sh.getCellByPosition(1,6).String="Снимок на " & Format(Now,"DD.MM.YYYY HH:MM:SS") & ". Учитываются только проведённые движения в базе."
Done:
    WMSDBX_CloseRS rs:WMSDBX_CloseStmt st
    ' Read-only query uses the existing connection; do not commit unrelated work.
    WMSDBX_Leave
    Exit Sub
Failed:
    On Error Resume Next
    sh.getCellByPosition(25,1).Value=0
    sh.getCellByPosition(1,6).String="Сводка не обновлена: " & sErr
    On Error GoTo 0
    MsgBox sErr,48,"WMS — Отчёт"
    GoTo Done
EH:
    sErr=CStr(Err) & " " & Error$:Resume Failed
End Sub

Sub WMSRPT_Build()
    Dim inp As Object,out As Object,r As Long,lastR As Long,dayValue As Double,nextDay As Double
    Dim doneText As String,planText As String,ongoing As String,line As String,kind As String,who As String,obj As String,resultText As String,body As String
    On Error GoTo EH
    inp=ThisComponent.Sheets.getByName(WMSRPT_INPUT):out=ThisComponent.Sheets.getByName(WMSRPT_OUTPUT)
    dayValue=inp.getCellByPosition(1,2).Value:nextDay=inp.getCellByPosition(1,4).Value
    If Not WMSRPT_ValidDay(dayValue) Or Not WMSRPT_ValidDay(nextDay) Or nextDay<=dayValue Then MsgBox "Проверьте дату отчёта и следующего рабочего дня.",48,"WMS":Exit Sub
    If Trim(inp.getCellByPosition(1,3).String)="" Then MsgBox "Укажите автора отчёта.",48,"WMS":Exit Sub
    If out.getCellByPosition(25,2).String<>"" And out.getCellByPosition(0,3).String<>"" And out.getCellByPosition(0,3).String<>out.getCellByPosition(25,2).String Then
        MsgBox "Готовый текст изменён вручную. Сохраните его, перенесите правки в события и очистите поле готового текста перед новой сборкой. Перезапись остановлена.",48,"WMS":Exit Sub
    End If
    WMSRPT_RefreshFacts
    If Not gWMSRPT_RefreshSucceeded Then Exit Sub
    If inp.getCellByPosition(25,1).Value=dayValue Then
        doneText=inp.getCellByPosition(1,7).String & Chr(10)
    Else
        doneText="Сводка складских операций за эту дату не обновлена." & Chr(10)
    End If
    lastR=WMSCore_LastContentRow(inp,6)
    For r=WMSRPT_FIRST To lastR
        If inp.getCellByPosition(0,r).Value=dayValue And UCase(Trim(inp.getCellByPosition(1,r).String))="ДА" Then
            kind=Trim(inp.getCellByPosition(2,r).String)
            who=Trim(inp.getCellByPosition(3,r).String):obj=Trim(inp.getCellByPosition(4,r).String)
            resultText=Trim(inp.getCellByPosition(5,r).String):line=Trim(inp.getCellByPosition(6,r).String)
            If line="" Then
                If resultText="" Or kind="" Then MsgBox "Заполните событие и результат в строке " & CStr(r+1) & " или исключите её из отчёта.",48,"WMS":Exit Sub
                line=WMSRPT_EventPhrase(kind)
                If who<>"" Then line=line & " — " & who
                If obj<>"" Then line=line & " («" & obj & "»)"
                line=line & ": " & resultText
            End If
            Select Case kind
                Case "Задача на завтра":planText=planText & "• " & line & Chr(10)
                Case "В процессе":ongoing=ongoing & "• " & line & Chr(10)
                Case Else:doneText=doneText & "• " & line & Chr(10)
            End Select
        End If
    Next r
    If planText="" Then planText="Задачи не указаны." & Chr(10)
    If ongoing="" Then ongoing="Продолжающиеся работы не указаны." & Chr(10)
    body=Format(CDate(dayValue),"DD.MM.YYYY") & " г." & Chr(10) & inp.getCellByPosition(1,3).String & Chr(10) & Chr(10)
    body=body & "1. Завершено сегодня:" & Chr(10) & doneText & Chr(10)
    body=body & "2. Задачи на " & Format(CDate(nextDay),"DD.MM.YYYY") & ":" & Chr(10) & planText & Chr(10)
    body=body & "3. В процессе:" & Chr(10) & ongoing
    If inp.getCellByPosition(25,2).String<>"" Then
        body=body & Chr(10) & "4. Вопросы, требующие внимания (текущее состояние):" & Chr(10) & inp.getCellByPosition(25,2).String
    End If
    If inp.getCellByPosition(25,3).String<>"" Then
        body=body & Chr(10) & "5. Рекомендуемые действия (не подтверждённые задачи):" & Chr(10) & inp.getCellByPosition(25,3).String
    End If
    If Len(body)>30000 Then
        MsgBox "Отчёт длиннее 30 000 символов. Выберите краткий формат, отключите контроль заказов или сократите ручные события. Готовый текст не заменён.",48,"ПОКАТАК":Exit Sub
    End If
    out.getCellByPosition(0,3).String=body
    out.getCellByPosition(25,2).String=body
    out.getCellByPosition(25,1).Value=dayValue
    ThisComponent.CurrentController.setActiveSheet(out)
    ThisComponent.CurrentController.select(out.getCellByPosition(0,3))
    Exit Sub
EH:
    MsgBox "Отчёт не сформирован: " & CStr(Err) & " " & Error$,16,"WMS"
End Sub

Sub WMSRPT_Save()
    Dim sh As Object,sfa As Object,txt As Object,streamObj As Object,dirURL As String,fileURL As String,baseName As String,i As Long,body As String
    On Error GoTo EH
    sh=ThisComponent.Sheets.getByName(WMSRPT_OUTPUT)
    If sh.getCellByPosition(25,1).Value<=0 Then MsgBox "Сначала сформируйте отчёт.",48,"WMS":Exit Sub
    body=sh.getCellByPosition(0,3).String
    If Trim(body)="" Then MsgBox "Текст отчёта пуст.",48,"WMS":Exit Sub
    If ThisComponent.URL="" Then MsgBox "Сначала сохраните книгу в постоянную папку.",48,"WMS":Exit Sub
    sfa=CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If sh.getCellByPosition(25,3).String=body Then
        If sh.getCellByPosition(25,4).String<>"" Then
            If sfa.exists(sh.getCellByPosition(25,4).String) Then
                MsgBox "Этот текст уже сохранён: " & ConvertFromURL(sh.getCellByPosition(25,4).String),64,"WMS":Exit Sub
            End If
        End If
    End If
    dirURL=ConvertToURL(WMSACT_BaseFolder())
    Do While Right(dirURL,1)="/":dirURL=Left(dirURL,Len(dirURL)-1):Loop
    dirURL=dirURL & "/Reports"
    If Not sfa.exists(dirURL) Then sfa.createFolder(dirURL)
    baseName=Format(CDate(sh.getCellByPosition(25,1).Value),"YYYY-MM-DD") & "_report"
    i=1
    Do
        fileURL=dirURL & "/" & baseName & "_v" & CStr(i) & ".txt"
        If Not sfa.exists(fileURL) And Not sfa.exists(fileURL & ".pending") Then Exit Do
        i=i+1
    Loop
    streamObj=sfa.openFileWrite(fileURL & ".pending")
    txt=CreateUnoService("com.sun.star.io.TextOutputStream")
    txt.setOutputStream(streamObj):txt.setEncoding("UTF-8"):txt.writeString(body):txt.closeOutput()
    sfa.move(fileURL & ".pending",fileURL)
    sh.getCellByPosition(25,3).String=body
    sh.getCellByPosition(25,4).String=fileURL
    MsgBox "Итог сохранён: " & ConvertFromURL(fileURL),64,"WMS — Отчёт"
    Exit Sub
EH:
    On Error Resume Next:txt.closeOutput():On Error GoTo 0
    MsgBox "Сохранение отчёта не подтверждено. Проверьте папку Reports; прежние итоги не заменялись.",16,"WMS"
End Sub

Function WMSRPT_EventPhrase(kind As String) As String
    Select Case kind
        Case "Принята поставка":WMSRPT_EventPhrase="Выполнена приёмка поставки"
        Case "Разгрузка погрузчиком":WMSRPT_EventPhrase="Выполнена разгрузка поставки с использованием погрузчика"
        Case "Помощь погрузчиком":WMSRPT_EventPhrase="Оказана помощь с использованием погрузчика"
        Case "Бухгалтерия":WMSRPT_EventPhrase="Проведена рабочая коммуникация с бухгалтерией"
        Case "Закупки":WMSRPT_EventPhrase="Проведена рабочая коммуникация со специалистом отдела закупок"
        Case "Задача на завтра":WMSRPT_EventPhrase="Запланировано"
        Case "В процессе":WMSRPT_EventPhrase="Продолжается работа"
        Case "Документы":WMSRPT_EventPhrase="Выполнено документальное сопровождение складских операций"
        Case "Расхождения":WMSRPT_EventPhrase="Зафиксировано расхождение для дальнейшего урегулирования"
        Case "Организация склада":WMSRPT_EventPhrase="Выполнены работы по организации складского хранения"
        Case Else:WMSRPT_EventPhrase=kind
    End Select
End Function

Function WMSRPT_ValidDay(dayValue As Double) As Boolean
    WMSRPT_ValidDay=False
    If dayValue<=0 Or dayValue>2958465 Then Exit Function
    If dayValue<>Int(dayValue) Then Exit Function
    WMSRPT_ValidDay=True
End Function

Sub WMSRPT_Documents()
    WMSRPT_AddEvent "Документы"
End Sub
Sub WMSRPT_Discrepancy()
    WMSRPT_AddEvent "Расхождения"
End Sub
Sub WMSRPT_Storage()
    WMSRPT_AddEvent "Организация склада"
End Sub

Function WMSRPT_OperationalSummary(con As Object,dayValue As Double) As String
    Dim st As Object,rs As Object,sql As String,body As String,labelText As String,sourceType As String,movementType As String,n As Long,errNo As Long
    On Error GoTo EH
    sql="SELECT CASE WHEN UPPER(TRIM(SOURCE_TYPE))='RETURN' THEN 'RETURN' ELSE 'OTHER' END AS FLOW_KIND,MOVEMENT_TYPE,UNIT_NAME,COUNT(DISTINCT COALESCE(OPERATION_ID,SOURCE_ID)),COUNT(DISTINCT PRODUCT_CODE),CAST(SUM(QTY) AS DOUBLE PRECISION),CAST(SUM(ABS(QTY)) AS DOUBLE PRECISION) FROM WMS_STOCK_MOVEMENTS " & _
        "WHERE COALESCE(STATUS_NAME,'POSTED')='POSTED' AND MOVEMENT_DATE=DATE '" & Format(CDate(dayValue),"YYYY-MM-DD") & "' " & _
        "GROUP BY CASE WHEN UPPER(TRIM(SOURCE_TYPE))='RETURN' THEN 'RETURN' ELSE 'OTHER' END,MOVEMENT_TYPE,UNIT_NAME ORDER BY 1,2,3"
    st=con.createStatement():rs=st.executeQuery(sql):body="":n=0
    Do While rs.next()
        If n>=40 Then body=body & "• Показаны первые 40 групп операций; полный перечень — в журнале движений." & Chr(10):Exit Do
        sourceType=UCase(Trim(rs.getString(1))):movementType=UCase(Trim(rs.getString(2)))
        Select Case movementType
            Case "IN"
                If sourceType="RETURN" Then
                    labelText="Оформлены возвраты ТМЦ на склад"
                Else
                    labelText="Зарегистрирована приёмка ТМЦ"
                End If
            Case "OUT":labelText="Зарегистрирована выдача ТМЦ со склада"
            Case "ADJUST":labelText="Отражены корректирующие движения по складскому учёту"
            Case Else:labelText="Зарегистрированы движения типа " & movementType
        End Select
        body=body & "• " & labelText & ": операций — " & CStr(rs.getLong(4)) & ", товаров — " & CStr(rs.getLong(5)) & _
            ", количество — " & CStr(Abs(rs.getDouble(6))) & " " & rs.getString(3) & "." & Chr(10)
        If movementType="ADJUST" Then body=body & "  Сумма модулей корректировок: " & CStr(rs.getDouble(7)) & " " & rs.getString(3) & "." & Chr(10)
        n=n+1
    Loop
    If n=0 Then body="• Проведённых движений за выбранную дату не зарегистрировано." & Chr(10)
    WMSDBX_CloseRS rs:WMSDBX_CloseStmt st
    WMSRPT_OperationalSummary=body
    Exit Function
EH:
    errNo=Err
    WMSDBX_CloseRS rs:WMSDBX_CloseStmt st
    Error errNo
End Function

Function WMSRPT_ControlSnapshot(inp As Object) As String
    Dim alerts As Variant,total As Long,scanned As Long,sErr As String,i As Long,body As String,item As Variant
    Dim missingSupplier As Boolean,missingDoc As Boolean,lateDate As Boolean,suggestions As String
    On Error GoTo EH
    alerts=WMSD_Alerts(total,scanned,sErr)
    If sErr<>"" Then WMSRPT_ControlSnapshot="Контроль заказов недоступен: " & sErr & Chr(10):Exit Function
    body="Проверка строк листа «Заказы» на " & Format(Now,"DD.MM.YYYY HH:MM") & "; это не исторический снимок на дату отчёта." & Chr(10)
    If total=0 Then
        WMSRPT_ControlSnapshot=body & "• В рамках проверок поставщика, документа и срока замечаний не найдено." & Chr(10)
        Exit Function
    End If
    body=body & "• Требуют внимания строк: " & CStr(total) & ". Ниже — до 30 строк; это не количество уникальных заказов." & Chr(10)
    For i=0 To UBound(alerts)
        item=alerts(i)
        If CStr(item(0))<>"" Then
            body=body & "• " & CStr(item(0)) & ": " & CStr(item(1))
            If CStr(item(2))<>"" Then body=body & "; счёт " & CStr(item(2))
            If CStr(item(3))<>"" Then body=body & "; поставщик " & CStr(item(3))
            If CStr(item(4))<>"" Then body=body & "; назначение " & CStr(item(4))
            body=body & "; " & CStr(item(5)) & "." & Chr(10)
            If InStr(CStr(item(0)),"Нет поставщика")>0 Then missingSupplier=True
            If InStr(CStr(item(0)),"Факт без документа")>0 Then missingDoc=True
            If InStr(CStr(item(0)),"Срок прошёл")>0 Then lateDate=True
        End If
    Next i
    If missingSupplier Then suggestions=suggestions & "• Уточнить реквизиты поставщиков по отмеченным строкам для корректной идентификации поставок." & Chr(10)
    If missingDoc Then suggestions=suggestions & "• Запросить закрывающие документы по отмеченным поступлениям; согласовать комплектность с закупками и бухгалтерией." & Chr(10)
    If lateDate Then suggestions=suggestions & "• Уточнить у закупок сроки отмеченных поставок и актуализировать ожидаемые даты после подтверждения." & Chr(10)
    inp.getCellByPosition(25,3).String=suggestions
    WMSRPT_ControlSnapshot=body
    Exit Function
EH:
    WMSRPT_ControlSnapshot="Контроль заказов не выполнен: " & CStr(Err) & " " & Error$ & Chr(10)
End Function

Function WMSRPT_SettingsValid(sh As Object,ByRef sErr As String) As Boolean
    WMSRPT_SettingsValid=False:sErr=""
    Select Case UCase(Trim(sh.getCellByPosition(1,5).String))
        Case "","ДА","НЕТ"
        Case Else:sErr="В B6 укажите Да или Нет для контроля заказов.":Exit Function
    End Select
    Select Case UCase(Trim(sh.getCellByPosition(1,8).String))
        Case "","ПОДРОБНЫЙ","КРАТКИЙ"
        Case Else:sErr="В B9 укажите Подробный или Краткий формат.":Exit Function
    End Select
    WMSRPT_SettingsValid=True
End Function

Function WMSRPT_Recipients(con As Object,dayValue As Double) As String
    Dim st As Object,rs As Object,body As String,n As Long,errNo As Long
    On Error GoTo EH
    st=con.createStatement()
    rs=st.executeQuery("SELECT i.EMPLOYEE_NAME,COALESCE(i.ENTRY_UNIT,i.UNIT_NAME),COUNT(DISTINCT i.PRODUCT_CODE),CAST(SUM(COALESCE(i.ENTRY_QTY,i.ISSUE_QTY)) AS DOUBLE PRECISION) FROM WMS_ISSUES i WHERE i.ISSUE_DATE=DATE '" & Format(CDate(dayValue),"YYYY-MM-DD") & "' AND i.ROW_STATE='POSTED' GROUP BY i.EMPLOYEE_NAME,COALESCE(i.ENTRY_UNIT,i.UNIT_NAME) ORDER BY 1,2")
    Do While rs.next()
        n=n+1
        If n>30 Then body=body & "• Другие получатели — в журнале выдач." & Chr(10):Exit Do
        body=body & "• Обеспечен получатель «" & rs.getString(1) & "»: " & CStr(rs.getDouble(4)) & " " & rs.getString(2) & ", товаров — " & CStr(rs.getLong(3)) & "." & Chr(10)
    Loop
    WMSDBX_CloseRS rs:WMSDBX_CloseStmt st
    WMSRPT_Recipients=body
    Exit Function
EH:
    errNo=Err:WMSDBX_CloseRS rs:WMSDBX_CloseStmt st
    Error errNo
End Function
