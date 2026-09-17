Option Explicit

Global Const WMSIMP_VERSION = "0.1.0-DRAFT"
Global Const WMSIMP_LIMIT = 20000

' Import owns no stock movements and issues no database writes.
Sub WMSIMP_Import()
    Dim doc As Object,sourceDoc As Object,sourceSheet As Object,picker As Object,chosen As Variant
    Dim options(4) As New com.sun.star.beans.PropertyValue
    Dim sourceURL As String,sheetName As String,sErr As String,normalized As Variant,headerRow As Long,lastRow As Long
    Dim i As Long,names As String,cursorObj As Object,sourceNull As Variant,targetNull As Variant,dateOffset As Double
    If Not WMSDBX_TryEnter("ORDER_IMPORT") Then MsgBox "Дождитесь завершения операции WMS.",48,"WMS":Exit Sub
    On Error GoTo EH
    doc=ThisComponent
    picker=CreateUnoService("com.sun.star.ui.dialogs.FilePicker")
    picker.initialize(Array(com.sun.star.ui.dialogs.TemplateDescription.FILEOPEN_SIMPLE))
    picker.appendFilter("Таблицы заказов","*.ods;*.xlsx;*.xls")
    picker.setMultiSelectionMode(False)
    If picker.execute()<>1 Then GoTo Done
    chosen=picker.getFiles():sourceURL=CStr(chosen(0))
    If Left(sourceURL,5)<>"file:" Or sourceURL=doc.URL Then sErr="Выберите отдельный локальный файл закупщика.":GoTo Failed
    options(0).Name="Hidden":options(0).Value=True
    options(1).Name="ReadOnly":options(1).Value=True
    options(2).Name="MacroExecutionMode":options(2).Value=com.sun.star.document.MacroExecMode.NEVER_EXECUTE
    options(3).Name="UpdateDocMode":options(3).Value=com.sun.star.document.UpdateDocMode.NO_UPDATE
    options(4).Name="AsTemplate":options(4).Value=True
    sourceDoc=StarDesktop.loadComponentFromURL(sourceURL,"_blank",0,options())
    If Not sourceDoc.supportsService("com.sun.star.sheet.SpreadsheetDocument") Then sErr="Выбранный файл не является таблицей.":GoTo Failed
    If sourceDoc.Sheets.Count=1 Then
        sourceSheet=sourceDoc.Sheets.getByIndex(0)
    Else
        For i=0 To sourceDoc.Sheets.Count-1:names=names & sourceDoc.Sheets.getByIndex(i).Name & Chr(10):Next i
        sheetName=InputBox("Укажите лист закупщика:" & Chr(10) & names,"Импорт заказов","Заказы")
        If sheetName="" Then GoTo Done
        If Not sourceDoc.Sheets.hasByName(sheetName) Then sErr="Такого листа нет в файле.":GoTo Failed
        sourceSheet=sourceDoc.Sheets.getByName(sheetName)
    End If
    headerRow=WMSIMP_HeaderRow(sourceSheet)
    If headerRow<0 Then sErr="Не распознана привычная структура A–W: номер, наименование, артикул F, количество H и единица I.":GoTo Failed
    lastRow=WMSCore_LastContentRow(sourceSheet,22)
    If lastRow<=headerRow Then sErr="В выбранном листе нет позиций.":GoTo Failed
    If lastRow-headerRow>WMSIMP_LIMIT Then sErr="В файле больше 20000 строк после заголовка. Разделите импорт на части.":GoTo Failed
    sourceNull=sourceDoc.getNumberFormatSettings().NullDate
    targetNull=doc.getNumberFormatSettings().NullDate
    If DateSerial(targetNull.Year,targetNull.Month,targetNull.Day)<>DateSerial(1899,12,30) Then sErr="Нестандартная начальная дата рабочей книги. Импорт остановлен.":GoTo Failed
    dateOffset=CDbl(DateSerial(sourceNull.Year,sourceNull.Month,sourceNull.Day))-CDbl(DateSerial(targetNull.Year,targetNull.Month,targetNull.Day))
    If Not WMSIMP_Normalize(sourceSheet,headerRow,lastRow,dateOffset,normalized,sErr) Then GoTo Failed
    ' Product confirmation concerns the concrete parsed file, not development approval.
    If MsgBox("Прочитано позиций: " & CStr(UBound(normalized)+1) & Chr(10) & "Добавить их в Заказы для проверки?",36,"Импорт заказов")<>6 Then GoTo Done
    If Not WMSIMP_Append(doc,normalized,sErr) Then GoTo Failed
    MsgBox "Позиции добавлены в Заказы. Проверьте данные обычной кнопкой проверки. Приёмка и остатки ещё не изменялись. Сохраните книгу.",64,"WMS — Импорт"
Done:
    On Error Resume Next
    sourceDoc.close(True):picker.dispose()
    WMSDBX_Leave
    On Error GoTo 0
    Exit Sub
Failed:
    MsgBox sErr,16,"Импорт остановлен"
    GoTo Done
EH:
    sErr="Импорт: " & CStr(Err) & " " & Error$:Resume Failed
End Sub

Function WMSIMP_HeaderRow(sh As Object) As Long
    Dim r As Long,a As String,b As String,docHeader As String,invoiceHeader As String,supplierHeader As String
    WMSIMP_HeaderRow=-1
    For r=0 To 29
        a=LCase(Trim(sh.getCellByPosition(0,r).String)):b=LCase(sh.getCellByPosition(1,r).String)
        If (InStr(a,"номер")>0 Or InStr(a,"№")>0 Or InStr(a,"п/п")>0) And InStr(b,"наименован")>0 Then
            docHeader=LCase(sh.getCellByPosition(2,r).String)
            invoiceHeader=Replace(LCase(sh.getCellByPosition(3,r).String),"ё","е")
            supplierHeader=LCase(sh.getCellByPosition(11,r).String)
            If (InStr(docHeader,"док")>0 Or InStr(docHeader,"упд")>0) And InStr(invoiceHeader,"счет")>0 And (InStr(supplierHeader,"поставщик")>0 Or InStr(supplierHeader,"от кого")>0 Or InStr(supplierHeader,"площад")>0) Then
                If InStr(LCase(sh.getCellByPosition(5,r).String),"артикул")>0 And InStr(LCase(sh.getCellByPosition(7,r).String),"кол")>0 And InStr(LCase(sh.getCellByPosition(8,r).String),"ед")>0 Then WMSIMP_HeaderRow=r:Exit Function
            End If
        End If
    Next r
End Function

Function WMSIMP_Normalize(sh As Object,headerRow As Long,lastRow As Long,dateOffset As Double,ByRef result As Variant,ByRef sErr As String) As Boolean
    Dim raw As Variant,item As Variant,out() As Variant,groups() As Long,normalized As Variant
    Dim i As Long,c As Long,j As Long,n As Long,groupNo As Long,pos As Long,previousPos As Long,state As Integer
    Dim mergeCols As Variant,col As Variant,cur As Object,addr As Variant,first As Long,finish As Long,valueObj As Variant
    Dim cDestination As Long,cExpected As Long,d As Double,qty As Double,valid As Boolean
    WMSIMP_Normalize=False:sErr=""
    On Error GoTo EH
    raw=sh.getCellRangeByPosition(0,headerRow+1,22,lastRow).getDataArray()
    ReDim groups(0 To UBound(raw)):ReDim out(0 To UBound(raw))
    For i=0 To UBound(raw)
        item=raw(i)
        If Trim(CStr(item(1)))="" Then
            ' Reject data-only rows; do not silently discard unnamed goods.
            For c=2 To 22
                If Trim(CStr(item(c)))<>"" Then sErr="Пустое наименование в строке " & CStr(headerRow+i+2):Exit Function
            Next c
            previousPos=0:GoTo NextPosition
        End If
        state=Orders_BulkOrderPositionState(item(0),pos)
        If state<>1 Then sErr="Номер позиции должен быть целым положительным числом, строка " & CStr(headerRow+i+2):Exit Function
        If pos=1 Then
            groupNo=groupNo+1
        ElseIf previousPos=0 Or pos<>previousPos+1 Then
            sErr="Нарушена последовательность позиций в строке " & CStr(headerRow+i+2):Exit Function
        End If
        groups(i)=groupNo:previousPos=pos
NextPosition:
    Next i
    mergeCols=Array(2,3,11,14,15,17)
    For Each col In mergeCols
        i=0
        Do While i<=UBound(raw)
            cur=sh.createCursorByRange(sh.getCellByPosition(CLng(col),headerRow+1+i))
            cur.collapseToMergedArea():addr=cur.RangeAddress
            If addr.StartColumn<>CLng(col) Or addr.EndColumn<>CLng(col) Then sErr="Горизонтальное объединение в колонке " & CStr(CLng(col)+1) & " не поддерживается.":Exit Function
            first=addr.StartRow-headerRow-1:finish=addr.EndRow-headerRow-1
            If first<0 Or finish>UBound(raw) Then sErr="Объединение выходит за границы строк заказа.":Exit Function
            If finish>first Then
                If groups(first)=0 Then sErr="Объединение начинается вне товарной позиции.":Exit Function
                For j=first To finish
                    If groups(j)<>groups(first) Then sErr="Объединение пересекает границу заказа, строка " & CStr(headerRow+j+2):Exit Function
                Next j
                valueObj=raw(first)(CLng(col))
                For j=first To finish:item=raw(j):item(CLng(col))=valueObj:raw(j)=item:Next j
            End If
            i=finish+1
        Loop
    Next col
    cDestination=WMSIMP_FindSourceHeader(sh,headerRow,"Назначение")
    cExpected=WMSIMP_FindSourceHeader(sh,headerRow,"Ожидаемая дата поступления")
    For i=0 To UBound(raw)
        If groups(i)=0 Then GoTo NextOutput
        item=raw(i)
        normalized=Array("","","","","","","","","","","","","","","","","","","","","","","","","")
        For c=0 To 22:normalized(c)=item(c):Next c
        normalized(4)="":normalized(16)="":normalized(18)=""
        normalized(5)=sh.getCellByPosition(5,headerRow+1+i).String
        normalized(21)=sh.getCellByPosition(21,headerRow+1+i).String
        For Each col In Array(6,7,9,10)
            If Trim(CStr(item(CLng(col))))<>"" Then
                valid=Orders_BulkTryNumber(item(CLng(col)),qty)
                If Not valid Or qty<0 Then sErr="Неверное количество или цена в строке " & CStr(headerRow+i+2):Exit Function
                normalized(CLng(col))=qty
            End If
        Next col
        For Each col In Array(12,13,14)
            If Not WMSIMP_DateValue(item(CLng(col)),dateOffset,d) Then sErr="Неверная дата в строке " & CStr(headerRow+i+2):Exit Function
            normalized(CLng(col))=""
            If d>0 Then normalized(CLng(col))=d
        Next col
        If cDestination>=0 Then normalized(23)=sh.getCellByPosition(cDestination,headerRow+1+i).String
        If cExpected>=0 Then
            valueObj=sh.getCellRangeByPosition(cExpected,headerRow+1+i,cExpected,headerRow+1+i).getDataArray()
            valueObj=valueObj(0)(0)
            If Not WMSIMP_DateValue(valueObj,dateOffset,d) Then sErr="Неверная ожидаемая дата в строке " & CStr(headerRow+i+2):Exit Function
            If d>0 Then normalized(24)=d
        End If
        out(n)=normalized:n=n+1
NextOutput:
    Next i
    If n=0 Then sErr="Нет товарных позиций.":Exit Function
    ReDim Preserve out(0 To n-1):result=out()
    WMSIMP_Normalize=True
    Exit Function
EH:
    sErr="Чтение структуры: " & CStr(Err) & " " & Error$
End Function

Function WMSIMP_FindSourceHeader(sh As Object,r As Long,title As String) As Long
    Dim c As Long
    WMSIMP_FindSourceHeader=-1
    For c=0 To 255
        If Trim(sh.getCellByPosition(c,r).String)=title Then WMSIMP_FindSourceHeader=c:Exit Function
    Next c
End Function

Function WMSIMP_DateValue(valueObj As Variant,dateOffset As Double,ByRef d As Double) As Boolean
    d=0:WMSIMP_DateValue=True
    If Trim(CStr(valueObj))="" Then Exit Function
    If VarType(valueObj)=8 Then
        d=Orders_BulkDateValue(valueObj,True)
    ElseIf IsNumeric(valueObj) Then
        If CDbl(valueObj)=0 Then Exit Function
        d=CDbl(valueObj)+dateOffset
    End If
    WMSIMP_DateValue=(d>0 And d<=2958465 And d=Int(d))
End Function

Function WMSIMP_Payload(item As Variant) As String
    Dim c As Long,v As String,s As String
    For c=0 To UBound(item)
        v=CStr(item(c))
        If VarType(item(c))<>8 Then v=Replace(v,",",".")
        s=s & CStr(Len(v)) & ":" & v
    Next c
    WMSIMP_Payload=s
End Function

Function WMSIMP_Append(doc As Object,normalized As Variant,ByRef sErr As String) As Boolean
    Dim sh As Object,lastRow As Long,lastCol As Long,firstRow As Long,cToken As Long,cDest As Long,cDate As Long,cID As Long,cGroup As Long
    Dim cap As Long,keys() As String,found As Boolean,slot As Long,old As Variant,payload As String,i As Long,c As Long,n As Long
    Dim rows() As Variant,item As Variant,targetRow() As Variant,groupID As String,written As Boolean,locked As Boolean,oldBusy As Boolean,busySet As Boolean
    Dim targetCols(0 To 22) As Long,baseHeaders As Variant
    WMSIMP_Append=False:sErr=""
    On Error GoTo EH
    sh=doc.Sheets.getByName("Заказы")
    cToken=Orders_FindHeader(sh,"_WMS_ImportPayload"):cDest=Orders_FindHeader(sh,"Назначение"):cDate=Orders_FindHeader(sh,"Ожидаемая дата поступления")
    cID=Orders_FindHeader(sh,"_WMS_SourceID"):cGroup=Orders_FindHeader(sh,"_WMS_OrderGroupID")
    If cToken<0 Or cDest<0 Or cDate<0 Or cID<0 Or cGroup<0 Then sErr="Сначала установите обновление полей заказа.":Exit Function
    baseHeaders=Orders_BaseHeaders()
    For c=0 To 22
        targetCols(c)=Orders_FindHeader(sh,CStr(baseHeaders(c)))
        If targetCols(c)<0 Then sErr="В Заказах отсутствует поле «" & CStr(baseHeaders(c)) & "».":Exit Function
    Next c
    lastCol=Orders_LastHeaderCol(sh):lastRow=WMSCore_LastContentRow(sh,lastCol)
    n=UBound(normalized)+1:firstRow=lastRow+1
    If firstRow+n>sh.Rows.Count Then sErr="Недостаточно строк для импорта.":Exit Function
    cap=64:Do While cap<(lastRow+n+1)*2:cap=cap*2:Loop
    ReDim keys(0 To cap-1)
    If lastRow>=1 Then
        old=sh.getCellRangeByPosition(cToken,1,cToken,lastRow).getDataArray()
        For i=0 To UBound(old)
            payload=CStr(old(i)(0))
            If payload<>"" Then
                slot=Orders_GroupHashSlot(keys(),cap,payload,found)
                If slot>=0 Then keys(slot)=payload
            End If
        Next i
    End If
    ReDim rows(0 To n-1)
    For i=0 To n-1
        item=normalized(i):payload=WMSIMP_Payload(item)
        If Len(payload)>32767 Then sErr="Слишком длинная позиция для безопасного контроля повторов.":Exit Function
        slot=Orders_GroupHashSlot(keys(),cap,payload,found)
        If slot<0 Or found Then sErr="Позиция " & CStr(i+1) & " совпадает с ранее импортированной или повторяется в этом файле. Импорт остановлен для проверки.":Exit Function
        keys(slot)=payload
        ReDim targetRow(0 To lastCol)
        For c=0 To lastCol:targetRow(c)="":Next c
        For c=0 To 22:targetRow(targetCols(c))=item(c):Next c
        If CDbl(item(0))=1 Then groupID=Orders_NewOrderGroupID(firstRow+i)
        targetRow(cID)=Orders_NewSourceID(firstRow+i):targetRow(cGroup)=groupID
        targetRow(cDest)=item(23):targetRow(cDate)=item(24):targetRow(cToken)=payload
        rows(i)=targetRow()
    Next i
    oldBusy=gWMSORD_Busy:gWMSORD_Busy=True:busySet=True
    doc.lockControllers():locked=True
    written=True
    sh.getCellRangeByPosition(0,firstRow,lastCol,firstRow+n-1).setDataArray(rows())
    WMSIMP_Append=True
Done:
    On Error Resume Next
    If locked Then doc.unlockControllers()
    If busySet Then gWMSORD_Busy=oldBusy
    On Error GoTo 0
    Exit Function
EH:
    sErr="Добавление строк: " & CStr(Err) & " " & Error$
    If written Then
        On Error Resume Next:Err=0
        sh.getCellRangeByPosition(0,firstRow,lastCol,firstRow+n-1).clearContents(23)
        If Err<>0 Then sErr=sErr & ". Очистка частичного импорта не подтверждена: закройте книгу без сохранения."
        On Error GoTo 0
    End If
    GoTo Done
End Function

Sub WMSIMP_InstallUI()
    Dim doc As Object,sh As Object,f As Object,button As Object,shape As Object,pos As Object,sz As Object,r As Long
    Dim ev As New com.sun.star.script.ScriptEventDescriptor
    doc=ThisComponent:sh=doc.Sheets.getByName("Инфо")
    If sh.DrawPage.Forms.hasByName("WMSIMP_FORM") Then Exit Sub
    f=doc.createInstance("com.sun.star.form.component.Form")
    sh.DrawPage.Forms.insertByName("WMSIMP_FORM",f)
    button=doc.createInstance("com.sun.star.form.component.CommandButton")
    button.Name="WMSIMP_OPEN":button.Label="ЗАГРУЗИТЬ ЗАКАЗ ЗАКУПЩИКА"
    f.insertByName(button.Name,button)
    r=WMSCore_LastContentRow(sh,9)+2
    sh.getCellByPosition(0,r).String="ИМПОРТ ЗАКАЗОВ ИЗ ТАБЛИЦЫ"
    shape=doc.createInstance("com.sun.star.drawing.ControlShape"):shape.Control=button
    pos=sh.getCellByPosition(0,r+1).Position
    sz=CreateUnoStruct("com.sun.star.awt.Size"):sz.Width=8500:sz.Height=900
    shape.Position=pos:shape.Size=sz:sh.DrawPage.add(shape)
    ev.ListenerType="com.sun.star.awt.XActionListener":ev.EventMethod="actionPerformed"
    ev.ScriptType="Script":ev.ScriptCode="vnd.sun.star.script:Standard.WMS_36_OrderImport.WMSIMP_Import?language=Basic&location=document"
    f.registerScriptEvent 0,ev
End Sub
