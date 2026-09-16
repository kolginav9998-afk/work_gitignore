Option Explicit

' ============================================================================
' WMS_14_UniversalReceiptV2
' v0.1.0 — universal incoming flow for Production / Office / Parts / Ozon
'
' Design:
' - all incoming goods still land in sheet "Заказы";
' - one receipt event gets one immutable _WMS_ReceiptEventID;
' - Ozon allows the same external order + same invoice with multiple UPDs;
' - every later delivery is a NEW receipt event, therefore a NEW lot;
' - article lookup is exact first, then unique partial match (>=3 chars);
' - no destructive changes to existing rows.
' ============================================================================

Global Const WMSRC_VERSION = "0.2.6-PARTS-ARTICLE-ROW-FIX"
Global Const WMSRC_SHEET = "Заказы"

Global Const WMSRC_H_EVENT = "_WMS_ReceiptEventID"
Global Const WMSRC_H_MODE = "_WMS_ReceiptMode"
Global Const WMSRC_H_SOURCE = "_WMS_ReceiptSource"
Global Const WMSRC_H_FROM = "_WMS_ReceivedFrom"
Global Const WMSRC_H_EXTORDER = "_WMS_ExternalOrderNo"
Global Const WMSRC_H_UPD = "_WMS_UPDNo"

Sub WMSRC_Install()
    Dim oDoc As Object,oSh As Object,oCon As Object,sErr As String
    oDoc=ThisComponent
    On Error GoTo EH

    If Not oDoc.Sheets.hasByName(WMSRC_SHEET) Then
        MsgBox "Лист 'Заказы' не найден.",16,"WMS — Универсальный приход"
        Exit Sub
    End If
    oSh=oDoc.Sheets.getByName(WMSRC_SHEET)

    WMSRC_GetOrCreateTechCol oSh,WMSRC_H_EVENT
    WMSRC_GetOrCreateTechCol oSh,WMSRC_H_MODE
    WMSRC_GetOrCreateTechCol oSh,WMSRC_H_SOURCE
    WMSRC_GetOrCreateTechCol oSh,WMSRC_H_FROM
    WMSRC_GetOrCreateTechCol oSh,WMSRC_H_EXTORDER
    WMSRC_GetOrCreateTechCol oSh,WMSRC_H_UPD
    WMSRC_HideTechCols oSh

    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then
        MsgBox "Поля Calc установлены, но Firebird недоступен:" & Chr(10) & sErr,48,"WMS — Универсальный приход"
        Exit Sub
    End If

    If Not WMSRC_EnsureDB(oCon,sErr) Then
        WMSDB_Close
        MsgBox sErr,16,"WMS — Универсальный приход"
        Exit Sub
    End If

    On Error Resume Next
    oCon.commit()
    On Error GoTo EH
    If Not WMSDB_SaveDatabaseDocument(sErr) Then
        WMSDB_Close
        MsgBox "Схема создана, но Base не сохранилась:" & Chr(10) & sErr,16,"WMS — Универсальный приход"
        Exit Sub
    End If
    WMSDB_Close

    MsgBox "Универсальный приход V2 установлен." & Chr(10) & _
           "Производство / Офис / Детали / Ozon готовы." & Chr(10) & _
           "Версия: " & WMSRC_VERSION,64,"WMS — Универсальный приход"
    Exit Sub
EH:
    On Error Resume Next
    WMSDB_Close
    MsgBox "WMSRC_Install: " & CStr(Err) & " " & Error$,16,"WMS — Универсальный приход"
End Sub

Sub WMSRC_NewReceiptMenu()
    Dim s As String
    s=InputBox("Какой приход создать?" & Chr(10) & Chr(10) & _
               "1 — Производство" & Chr(10) & _
               "2 — Детали" & Chr(10) & _
               "3 — Офис" & Chr(10) & _
               "4 — Ozon" & Chr(10) & _
               "5 — Поставщик / другое" & Chr(10) & _
               "6 — Начальный остаток","WMS — Новый приход","1")
    If Trim(s)="" Then Exit Sub

    Select Case Trim(s)
        Case "1": WMSRC_NewProduction
        Case "2": WMSRC_NewParts
        Case "3": WMSRC_NewOffice
        Case "4": WMSRC_NewOzon
        Case "5": WMSRC_NewOther
        Case "6": WMSRC_NewOpeningBalance
        Case Else: MsgBox "Введите число от 1 до 6.",48,"WMS — Новый приход"
    End Select
End Sub

Sub WMSRC_NewProduction()
    Dim who As String,n As Long
    who=InputBox("От кого / какой участок?" & Chr(10) & _
                 "Можно оставить пустым, если неизвестно.","WMS — Производство","Производство")
    n=WMSRC_AskRows("Сколько позиций принесли?",2)
    If n<=0 Then Exit Sub
    WMSRC_CreateBlock "PRODUCTION","Производство",who,"","","",n,True
End Sub

Sub WMSRC_NewParts()
    Dim who As String,n As Long
    who=InputBox("От кого поступили детали?","WMS — Детали","Производство")
    n=WMSRC_AskRows("Сколько позиций деталей?",2)
    If n<=0 Then Exit Sub
    WMSRC_CreateBlock "PARTS","Производство",who,"","","",n,True
    MsgBox "Для деталей заполняй:" & Chr(10) & _
           "B — название" & Chr(10) & _
           "F — артикул детали" & Chr(10) & _
           "G — количество." & Chr(10) & Chr(10) & _
           "Потом нажми 'Распознать по артикулам'.",64,"WMS — Детали"
End Sub

Sub WMSRC_NewOffice()
    Dim who As String,n As Long
    who=InputBox("От кого / из какого офиса приехал товар?","WMS — Офис","Офис")
    n=WMSRC_AskRows("Сколько строк подготовить?" & Chr(10) & _
                    "Для массового офисного прихода можно сразу 30.",30)
    If n<=0 Then Exit Sub
    WMSRC_CreateBlock "OFFICE","Офис",who,"","","",n,True
End Sub

Sub WMSRC_NewOzon()
    Dim oDoc As Object,oSh As Object,extOrder As String,invoice As String,upd As String,who As String,n As Long
    Dim found As Long,orderedQty As Double,receivedQty As Double,sErr As String

    oDoc=ThisComponent
    If Not oDoc.Sheets.hasByName(WMSRC_SHEET) Then Exit Sub
    oSh=oDoc.Sheets.getByName(WMSRC_SHEET)

    extOrder=InputBox("Номер УЖЕ существующего заказа Ozon:" & Chr(10) & _
                      "Приход создаётся как допоставка к этому заказу.","WMS — Ozon","")
    If Trim(extOrder)="" Then Exit Sub

    found=WMSRC_FindExistingOzonOrder(oSh,extOrder,invoice,who,orderedQty)
    If found=0 Then
        MsgBox "Заказ '" & extOrder & "' не найден на листе 'Заказы'." & Chr(10) & _
               "Сначала должен существовать сам заказ. Новый Ozon-заказ из прихода не создаётся.",48,"WMS — Ozon"
        Exit Sub
    End If

    upd=InputBox("Номер УПД этой поставки." & Chr(10) & _
                 "Если УПД ещё нет — оставь пустым. Её можно добавить позже к ReceiptID.","WMS — Ozon","")
    n=WMSRC_AskRows("Сколько товарных строк приехало сейчас?",1)
    If n<=0 Then Exit Sub

    WMSRC_CreateBlock "OZON","Ozon",who,extOrder,invoice,upd,n,False
End Sub

Function WMSRC_FindExistingOzonOrder(oSh As Object,orderNo As String,ByRef invoice As String,ByRef who As String,ByRef orderedQty As Double) As Long
    Dim lastR As Long,r As Long,cExt As Long,v As String
    WMSRC_FindExistingOzonOrder=0
    cExt=WMSRC_GetOrCreateTechCol(oSh,WMSRC_H_EXTORDER)
    lastR=WMSRC_LastBusinessRow(oSh)

    For r=1 To lastR
        v=Trim(oSh.getCellByPosition(cExt,r).String)
        If v="" Then v=Trim(oSh.getCellByPosition(2,r).String)
        If UCase(v)=UCase(Trim(orderNo)) Then
            WMSRC_FindExistingOzonOrder=WMSRC_FindExistingOzonOrder+1
            If invoice="" Then invoice=Trim(oSh.getCellByPosition(3,r).String)
            If who="" Then who=Trim(oSh.getCellByPosition(11,r).String)
            orderedQty=orderedQty+oSh.getCellByPosition(7,r).Value
        End If
    Next r
End Function

Sub WMSRC_AttachUPDToSelectedReceipt()
    Dim oDoc As Object,oSh As Object,oSel As Object
    Dim srcR As Long,dstR As Long,i As Long
    Dim cEvent As Long,cMode As Long,cSource As Long,cFrom As Long,cExt As Long,cUpd As Long
    Dim newEventID As String,upd As String,sQty As String,qty As Double
    Dim srcName As String,srcCode As String,srcArticle As String
    Dim sourceName As String,externalOrder As String
    Dim orderedQty As Double,receivedBefore As Double,draftReceived As Double
    Dim sCheckErr As String,extraHeader As Variant,extraCol As Long

    oDoc=ThisComponent
    If Not oDoc.Sheets.hasByName(WMSRC_SHEET) Then Exit Sub
    oSh=oDoc.Sheets.getByName(WMSRC_SHEET)
    oSel=oDoc.CurrentSelection
    srcR=oSel.CellAddress.Row

    If srcR<1 Then
        MsgBox "Выберите товарную строку существующего заказа.",48,"WMS — Допоставка"
        Exit Sub
    End If

    cEvent=WMSRC_GetOrCreateTechCol(oSh,WMSRC_H_EVENT)
    cMode=WMSRC_GetOrCreateTechCol(oSh,WMSRC_H_MODE)
    cSource=WMSRC_GetOrCreateTechCol(oSh,WMSRC_H_SOURCE)
    cFrom=WMSRC_GetOrCreateTechCol(oSh,WMSRC_H_FROM)
    cExt=WMSRC_GetOrCreateTechCol(oSh,WMSRC_H_EXTORDER)
    cUpd=WMSRC_GetOrCreateTechCol(oSh,WMSRC_H_UPD)

    srcName=Trim(oSh.getCellByPosition(1,srcR).String)
    srcArticle=Trim(oSh.getCellByPosition(5,srcR).String)
    srcCode=Trim(oSh.getCellByPosition(21,srcR).String)

    If srcName="" And srcArticle="" And srcCode="" Then
        MsgBox "В выбранной строке нет товара.",48,"WMS — Допоставка"
        Exit Sub
    End If

    ' Source can come from a pasted manager order and may not yet be normalized.
    sourceName=Trim(oSh.getCellByPosition(23,srcR).String)
    If sourceName="" Then sourceName=Trim(oSh.getCellByPosition(cSource,srcR).String)
    If sourceName="" Then
        sourceName=InputBox("Источник прихода этой допоставки:" & Chr(10) & _
                            "Например: Ozon, Поставщик, Офис, Производство","WMS — Допоставка","")
        If Trim(sourceName)="" Then Exit Sub
    End If

    externalOrder=Trim(oSh.getCellByPosition(cExt,srcR).String)
    If externalOrder="" Then
        externalOrder=Trim(oSh.getCellByPosition(2,srcR).String)
    End If
    If externalOrder="" Then
        externalOrder=InputBox("Номер заказа / связь с исходным заказом:" & Chr(10) & _
                               "Он нужен, чтобы WMS могла контролировать допоставки и перепоставку.", _
                               "WMS — Допоставка","")
        If Trim(externalOrder)="" Then Exit Sub
    End If

    upd=InputBox("Номер УПД НОВОЙ допоставки." & Chr(10) & _
                 "Если УПД ещё нет — оставьте пустым.","WMS — Допоставка","")

    sQty=InputBox("Сколько выбранного товара приехало СЕЙЧАС?","WMS — Допоставка","")
    If Trim(sQty)="" Then Exit Sub

    On Error GoTo BadQty
    qty=CDbl(sQty)
    If qty<=0 Then GoTo BadQty
    On Error GoTo EH

    ' Production safety: preserve the original ordered quantity and count
    ' already conducted + still visible draft deliveries before allowing overreceipt.
    orderedQty=oSh.getCellByPosition(7,srcR).Value
    sCheckErr=""
    receivedBefore=WMSRC_ReceiptCountConsolidated(oDoc,oSh,externalOrder,srcCode,srcArticle,srcName,Trim(oSh.getCellByPosition(8,srcR).String),sCheckErr)
    If sCheckErr<>"" Then
        MsgBox "Не удалось проверить проведённые и ожидающие проведения поступления." & Chr(10) & sCheckErr,16,"WMS — Допоставка"
        Exit Sub
    End If
    If orderedQty>0 Then
        If Not WMSSAFE_ConfirmOverReceipt(externalOrder,orderedQty,receivedBefore,qty) Then
            Exit Sub
        End If
    End If

    dstR=WMSRC_FirstEmptyBusinessRow(oSh)
    newEventID=WMSRC_NewEventID(UCase(Left(sourceName,12)))

    gWMSORD_Busy=True

    ' Copy the selected manager/order row as template.
    ' Old receipt identifiers are replaced below.
    For i=0 To 24
        oSh.getCellByPosition(i,dstR).String=oSh.getCellByPosition(i,srcR).String
        If oSh.getCellByPosition(i,srcR).Type=1 Then
            oSh.getCellByPosition(i,dstR).Value=oSh.getCellByPosition(i,srcR).Value
        End If
    Next i

    ' Copy business extras by name, keeping dates numeric and display format intact.
    For Each extraHeader In Array("Назначение","Ожидаемая дата поступления")
        extraCol=Orders_FindHeader(oSh,CStr(extraHeader))
        If extraCol>=0 Then
            oSh.getCellByPosition(extraCol,dstR).Formula=oSh.getCellByPosition(extraCol,srcR).Formula
            oSh.getCellByPosition(extraCol,dstR).NumberFormat=oSh.getCellByPosition(extraCol,srcR).NumberFormat
        End If
    Next extraHeader

    ' New physical delivery of the same item.
    oSh.getCellByPosition(6,dstR).Value=qty
    ' H keeps the original ordered quantity copied from the parent row.
    ' This lets Firebird retain order semantics while G is this physical delivery.
    oSh.getCellByPosition(12,dstR).Value=CDbl(Date)

    ' C = current UPD/document for this delivery.
    oSh.getCellByPosition(2,dstR).String=upd

    ' Preserve the order relationship, create a NEW receipt event.
    oSh.getCellByPosition(cEvent,dstR).String=newEventID
    oSh.getCellByPosition(cMode,dstR).String="ORDER_DELIVERY"
    oSh.getCellByPosition(cSource,dstR).String=sourceName
    oSh.getCellByPosition(cFrom,dstR).String=oSh.getCellByPosition(cFrom,srcR).String
    oSh.getCellByPosition(cExt,dstR).String=externalOrder
    oSh.getCellByPosition(cUpd,dstR).String=upd
    oSh.getCellByPosition(23,dstR).String=sourceName

    Orders_SetReceiptType oSh,dstR,"DIRECT_RECEIPT"
    gWMSORD_Busy=False

    Orders_ProcessRow oDoc,oSh,dstR,True,True,True,True
    oDoc.CurrentController.select(oSh.getCellByPosition(1,dstR))

    MsgBox "Создана отдельная позиция допоставки." & Chr(10) & _
           "Товар: " & srcName & Chr(10) & _
           "Источник: " & sourceName & Chr(10) & _
           "Количество: " & CStr(qty) & Chr(10) & _
           "УПД: " & IIf(Trim(upd)="","ожидается",upd) & Chr(10) & _
           "Новый ReceiptID: " & newEventID & Chr(10) & Chr(10) & _
           "Исходная строка заказа не изменена.",64,"WMS — Допоставка"
    Exit Sub

BadQty:
    gWMSORD_Busy=False
    MsgBox "Количество должно быть числом больше нуля.",48,"WMS — Допоставка"
    Exit Sub
EH:
    gWMSORD_Busy=False
    MsgBox "Создание допоставки: " & CStr(Err) & " " & Error$,16,"WMS — Допоставка"
End Sub

Sub WMSRC_DeleteDraftReceipt()
    Dim oDoc As Object,oSh As Object,oSel As Object,r As Long,cEvent As Long,eventID As String
    Dim lastR As Long,i As Long,n As Long
    oDoc=ThisComponent:oSh=oDoc.Sheets.getByName(WMSRC_SHEET)
    oSel=oDoc.CurrentSelection:r=oSel.CellAddress.Row
    cEvent=WMSRC_GetOrCreateTechCol(oSh,WMSRC_H_EVENT)
    eventID=Trim(oSh.getCellByPosition(cEvent,r).String)
    If eventID="" Then MsgBox "Выберите строку черновика прихода.",48,"WMS — Удалить черновик":Exit Sub

    If MsgBox("Удалить только НЕПРОВЕДЁННЫЙ черновик " & eventID & "?" & Chr(10) & _
              "Проведённые записи Firebird эта команда не удаляет.",36,"WMS — Удалить черновик")<>6 Then Exit Sub

    lastR=WMSRC_LastBusinessRow(oSh)
    For i=lastR To 1 Step -1
        If Trim(oSh.getCellByPosition(cEvent,i).String)=eventID Then
            oSh.Rows.removeByIndex(i,1):n=n+1
        End If
    Next i
    MsgBox "Удалено строк черновика: " & CStr(n),64,"WMS — Удалить черновик"
End Sub

Sub WMSRC_NewOpeningBalance()
    Dim n As Long
    n=WMSRC_AskRows("Сколько позиций начального остатка внести?" & Chr(10) & _
                    "Используйте только для реального товара, который физически уже есть на складе на дату запуска WMS.",10)
    If n<=0 Then Exit Sub
    WMSRC_CreateBlock "OPENING","Начальный остаток","Инвентаризация на запуск","","","",n,True
    MsgBox "Подготовлены строки начального остатка." & Chr(10) & _
           "Заполните фактическое количество, единицу, код/артикул и место хранения." & Chr(10) & _
           "Проводите только после физической проверки товара.",64,"WMS — Начальный остаток"
End Sub

Sub WMSRC_NewOther()
    Dim src As String,who As String,n As Long
    src=InputBox("Источник прихода:","WMS — Другой приход","Поставщик")
    If Trim(src)="" Then Exit Sub
    who=InputBox("От кого / поставщик / отправитель:","WMS — Другой приход",src)
    n=WMSRC_AskRows("Сколько позиций?",1)
    If n<=0 Then Exit Sub
    WMSRC_CreateBlock "OTHER",src,who,"","","",n,True
End Sub

Function WMSRC_AskRows(prompt As String,defaultN As Long) As Long
    Dim s As String,n As Long
    WMSRC_AskRows=0
    s=InputBox(prompt,"WMS — Новый приход",CStr(defaultN))
    If Trim(s)="" Then Exit Function
    On Error GoTo Bad
    n=CLng(s)
    If n<1 Or n>500 Then GoTo Bad
    WMSRC_AskRows=n
    Exit Function
Bad:
    MsgBox "Количество строк должно быть от 1 до 500.",48,"WMS — Новый приход"
End Function

Sub WMSRC_CreateBlock(mode As String,src As String,who As String,extOrder As String,invoice As String,upd As String,n As Long,mirrorFactToOrdered As Boolean)
    Dim oDoc As Object,oSh As Object,startR As Long,r As Long,i As Long,eventID As String
    Dim cEvent As Long,cMode As Long,cSource As Long,cFrom As Long,cExt As Long,cUpd As Long
    Dim formulaText As String

    oDoc=ThisComponent
    If Not oDoc.Sheets.hasByName(WMSRC_SHEET) Then Exit Sub
    oSh=oDoc.Sheets.getByName(WMSRC_SHEET)
    oDoc.CurrentController.setActiveSheet(oSh)

    cEvent=WMSRC_GetOrCreateTechCol(oSh,WMSRC_H_EVENT)
    cMode=WMSRC_GetOrCreateTechCol(oSh,WMSRC_H_MODE)
    cSource=WMSRC_GetOrCreateTechCol(oSh,WMSRC_H_SOURCE)
    cFrom=WMSRC_GetOrCreateTechCol(oSh,WMSRC_H_FROM)
    cExt=WMSRC_GetOrCreateTechCol(oSh,WMSRC_H_EXTORDER)
    cUpd=WMSRC_GetOrCreateTechCol(oSh,WMSRC_H_UPD)
    WMSRC_HideTechCols oSh

    eventID=WMSRC_NewEventID(mode)
    startR=WMSRC_FirstEmptyBusinessRow(oSh)

    On Error GoTo EH
    gWMSORD_Busy=True
    For i=0 To n-1
        r=startR+i
        oSh.getCellByPosition(0,r).Value=i+1
        oSh.getCellByPosition(11,r).String=who
        oSh.getCellByPosition(23,r).String=src
        oSh.getCellByPosition(12,r).Value=CDbl(Date)

        If Trim(invoice)<>"" Then oSh.getCellByPosition(3,r).String=invoice
        If Trim(upd)<>"" Then oSh.getCellByPosition(2,r).String=upd

        oSh.getCellByPosition(cEvent,r).String=eventID
        oSh.getCellByPosition(cMode,r).String=mode
        oSh.getCellByPosition(cSource,r).String=src
        oSh.getCellByPosition(cFrom,r).String=who
        oSh.getCellByPosition(cExt,r).String=extOrder
        oSh.getCellByPosition(cUpd,r).String=upd

        Orders_SetReceiptType oSh,r,"DIRECT_RECEIPT"

        If mirrorFactToOrdered Then
            formulaText="=IF(G" & CStr(r+1) & "="""";"""";G" & CStr(r+1) & ")"
            oSh.getCellByPosition(7,r).Formula=formulaText
        End If
    Next i
    gWMSORD_Busy=False

    For i=0 To n-1
        Orders_ProcessRow oDoc,oSh,startR+i,True,True,True,True
    Next i

    oDoc.CurrentController.select(oSh.getCellByPosition(1,startR))
    MsgBox "Создан приход " & eventID & Chr(10) & _
           "Источник: " & src & Chr(10) & _
           "Строк: " & CStr(n) & _
           IIf(mode="OZON",Chr(10) & "Заказ Ozon: " & extOrder & Chr(10) & _
               "Счёт: " & invoice & Chr(10) & "УПД: " & upd,""), _
           64,"WMS — Новый приход"
    Exit Sub
EH:
    gWMSORD_Busy=False
    MsgBox "Создание прихода: " & CStr(Err) & " " & Error$,16,"WMS — Новый приход"
End Sub

Sub WMSRC_ResolveArticles()
    Dim oDoc As Object,oSh As Object,oCon As Object,sErr As String
    Dim lastR As Long,r As Long,art As String,code As String,nm As String,unitName As String,loc As String
    Dim currentCode As String,mode As String
    Dim nFound As Long,nNew As Long,nAmb As Long,nMiss As Long,nSkip As Long
    Dim cMode As Long,result As Integer

    On Error GoTo EH
    oDoc=ThisComponent
    If Not oDoc.Sheets.hasByName(WMSRC_SHEET) Then Exit Sub
    oSh=oDoc.Sheets.getByName(WMSRC_SHEET)

    cMode=WMSRC_GetOrCreateTechCol(oSh,WMSRC_H_MODE)
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then MsgBox sErr,16,"WMS — Артикулы":Exit Sub

    lastR=WMSRC_LastBusinessRow(oSh)
    For r=1 To lastR
        art=Trim(oSh.getCellByPosition(5,r).String)
        currentCode=Trim(oSh.getCellByPosition(21,r).String)
        mode=UCase(Trim(oSh.getCellByPosition(cMode,r).String))

        If art<>"" Then
            ' PARTS rows are checked even when column V was already filled automatically.
            ' This fixes first-receipt rows prepared by Orders_ProcessRow.
            If mode="PARTS" Or currentCode="" Then
                code="":nm="":unitName="":loc="":sErr=""
                result=WMSRC_FindByArticleCon(oCon,art,code,nm,unitName,loc,sErr)
                If sErr<>"" Then GoTo DBERR

                Select Case result
                    Case 1
                        oSh.getCellByPosition(21,r).String=code
                        If Trim(oSh.getCellByPosition(1,r).String)="" Then oSh.getCellByPosition(1,r).String=nm
                        If Trim(oSh.getCellByPosition(8,r).String)="" Then oSh.getCellByPosition(8,r).String=unitName
                        If Trim(oSh.getCellByPosition(19,r).String)="" Then oSh.getCellByPosition(19,r).String=loc
                        nFound=nFound+1

                    Case 2
                        nAmb=nAmb+1

                    Case Else
                        If mode="PARTS" Then
                            If currentCode="" Then
                                oSh.getCellByPosition(21,r).String=art
                            End If
                            nNew=nNew+1
                        Else
                            nMiss=nMiss+1
                        End If
                End Select
            Else
                nSkip=nSkip+1
            End If
        End If
    Next r

    WMSDB_Close
    Orders_CheckAll
    MsgBox "Распознавание артикулов завершено." & Chr(10) & _
           "Найдено в БД: " & CStr(nFound) & Chr(10) & _
           "Новые детали подготовлены: " & CStr(nNew) & Chr(10) & _
           "Неоднозначно: " & CStr(nAmb) & Chr(10) & _
           "Не найдено: " & CStr(nMiss) & Chr(10) & _
           "Уже готовые обычные строки: " & CStr(nSkip) & Chr(10) & Chr(10) & _
           IIf(nAmb=0 And nMiss=0,"Можно нажимать «Провести заказ».", _
               "Строки с неоднозначным/не найденным артикулом не проводите."), _
           IIf(nAmb=0 And nMiss=0,64,48),"WMS — Артикулы"
    Exit Sub

DBERR:
    WMSDB_Close
    MsgBox sErr,16,"WMS — Артикулы"
    Exit Sub
EH:
    On Error Resume Next
    WMSDB_Close
    MsgBox "Распознавание артикулов: " & CStr(Err) & " " & Error$,16,"WMS — Артикулы"
End Sub

Function WMSRC_FindByArticleCon(oCon As Object,art As String,ByRef code As String,ByRef nm As String,ByRef unitName As String,ByRef loc As String,ByRef sErr As String) As Integer
    Dim oStmt As Object,oRS As Object,sql As String,n As Long
    WMSRC_FindByArticleCon=0:sErr=""
    On Error GoTo EH

    sql="SELECT PRODUCT_CODE,PRODUCT_NAME,UNIT_NAME,DEFAULT_LOCATION FROM WMS_PRODUCTS " & _
        "WHERE ACTIVE_FLAG=1 AND SUPPLIER_ARTICLE=" & WMSDB_SQLText(Trim(art))
    oStmt=oCon.createStatement():oRS=oStmt.executeQuery(sql)
    Do While oRS.next()
        n=n+1
        If n=1 Then
            code=oRS.getString(1):nm=oRS.getString(2):unitName=oRS.getString(3):loc=oRS.getString(4)
        End If
        If n>1 Then Exit Do
    Loop

    If n=1 Then
        WMSRC_FindByArticleCon=1
    ElseIf n>1 Then
        WMSRC_FindByArticleCon=2
    End If
    Exit Function
EH:
    sErr="Поиск артикула: " & CStr(Err) & " " & Error$
End Function

Function WMSRC_EnsureDB(oCon As Object,ByRef sErr As String) As Boolean
    Dim cols,types,i As Long
    WMSRC_EnsureDB=False:sErr=""
    cols=Array("RECEIPT_EVENT_ID","RECEIPT_MODE","RECEIPT_SOURCE","RECEIVED_FROM","EXTERNAL_ORDER_NO","UPD_NO")
    types=Array("VARCHAR(128)","VARCHAR(48)","VARCHAR(120)","VARCHAR(180)","VARCHAR(120)","VARCHAR(120)")

    If Not WMSRC_TableExists(oCon,"WMS_ORDER_LINES",sErr) Then
        If sErr="" Then sErr="Нет WMS_ORDER_LINES. Сначала установите WMS_06."
        Exit Function
    End If

    For i=0 To UBound(cols)
        If Not WMSRC_ColumnExists(oCon,"WMS_ORDER_LINES",CStr(cols(i)),sErr) Then
            If sErr<>"" Then Exit Function
            If Not WMSRC_ExecDDL(oCon,"ALTER TABLE WMS_ORDER_LINES ADD " & CStr(cols(i)) & " " & CStr(types(i)),sErr) Then Exit Function
        End If
    Next i
    WMSRC_EnsureDB=True
End Function

Function WMSRC_TableExists(oCon As Object,t As String,ByRef sErr As String) As Boolean
    Dim oStmt As Object,oRS As Object,sql As String
    WMSRC_TableExists=False:sErr=""
    On Error GoTo EH
    sql="SELECT COUNT(*) FROM RDB$RELATIONS WHERE TRIM(RDB$RELATION_NAME)=" & WMSDB_SQLText(UCase(t)) & " AND COALESCE(RDB$SYSTEM_FLAG,0)=0"
    oStmt=oCon.createStatement():oRS=oStmt.executeQuery(sql)
    If oRS.next() Then WMSRC_TableExists=(oRS.getInt(1)>0)
    Exit Function
EH:
    sErr="Проверка таблицы " & t & ": " & CStr(Err) & " " & Error$
End Function

Function WMSRC_ColumnExists(oCon As Object,t As String,c As String,ByRef sErr As String) As Boolean
    Dim oStmt As Object,oRS As Object,sql As String
    WMSRC_ColumnExists=False:sErr=""
    On Error GoTo EH
    sql="SELECT COUNT(*) FROM RDB$RELATION_FIELDS WHERE TRIM(RDB$RELATION_NAME)=" & WMSDB_SQLText(UCase(t)) & _
        " AND TRIM(RDB$FIELD_NAME)=" & WMSDB_SQLText(UCase(c))
    oStmt=oCon.createStatement():oRS=oStmt.executeQuery(sql)
    If oRS.next() Then WMSRC_ColumnExists=(oRS.getInt(1)>0)
    Exit Function
EH:
    sErr="Проверка поля " & c & ": " & CStr(Err) & " " & Error$
End Function

Function WMSRC_ExecDDL(oCon As Object,sql As String,ByRef sErr As String) As Boolean
    Dim oStmt As Object
    WMSRC_ExecDDL=False:sErr=""
    On Error GoTo EH
    oStmt=oCon.createStatement():oStmt.execute(sql)
    WMSRC_ExecDDL=True
    Exit Function
EH:
    sErr="DDL: " & CStr(Err) & " " & Error$
End Function

Function WMSRC_GetOrCreateTechCol(oSh As Object,h As String) As Long
    Dim cur As Object,lastCol As Long,c As Long
    cur=oSh.createCursor():cur.gotoEndOfUsedArea(True)
    lastCol=cur.RangeAddress.EndColumn
    If lastCol<25 Then lastCol=25
    For c=26 To lastCol
        If Trim(oSh.getCellByPosition(c,0).String)=h Then
            WMSRC_GetOrCreateTechCol=c
            Exit Function
        End If
    Next c
    c=lastCol+1
    oSh.getCellByPosition(c,0).String=h
    WMSRC_GetOrCreateTechCol=c
End Function

Sub WMSRC_HideTechCols(oSh As Object)
    Dim a,i As Long,c As Long
    a=Array(WMSRC_H_EVENT,WMSRC_H_MODE,WMSRC_H_SOURCE,WMSRC_H_FROM,WMSRC_H_EXTORDER,WMSRC_H_UPD)
    On Error Resume Next
    For i=0 To UBound(a)
        c=WMSRC_GetOrCreateTechCol(oSh,CStr(a(i)))
        oSh.Columns.getByIndex(c).IsVisible=False
    Next i
    On Error GoTo 0
End Sub

Function WMSRC_FirstEmptyBusinessRow(oSh As Object) As Long
    Dim lastR As Long,r As Long
    lastR=WMSRC_LastBusinessRow(oSh)+1
    If lastR<1 Then lastR=1
    For r=1 To lastR+500
        If Trim(oSh.getCellByPosition(1,r).String)="" And _
           Trim(oSh.getCellByPosition(5,r).String)="" And _
           Trim(oSh.getCellByPosition(21,r).String)="" And _
           Abs(oSh.getCellByPosition(6,r).Value)<0.0000001 Then
            WMSRC_FirstEmptyBusinessRow=r
            Exit Function
        End If
    Next r
    WMSRC_FirstEmptyBusinessRow=lastR+1
End Function

Function WMSRC_LastBusinessRow(oSh As Object) As Long
    Dim cur As Object
    cur=oSh.createCursor():cur.gotoEndOfUsedArea(True)
    WMSRC_LastBusinessRow=cur.RangeAddress.EndRow
End Function


Function WMSRC_ReceivedBeforeDB(oDoc As Object,externalOrder As String,productCode As String,articleCode As String,productName As String,ByRef sErr As String) As Double
    Dim oCon As Object,oStmt As Object,oRS As Object
    Dim sql As String,matchSQL As String

    WMSRC_ReceivedBeforeDB=0
    sErr=""
    If Trim(externalOrder)="" Then Exit Function

    On Error GoTo EH
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then Exit Function

    If Trim(productCode)<>"" Then
        matchSQL="PRODUCT_CODE=" & WMSDB_SQLText(Trim(productCode))
    ElseIf Trim(articleCode)<>"" Then
        matchSQL="SUPPLIER_ARTICLE=" & WMSDB_SQLText(Trim(articleCode))
    Else
        matchSQL="PRODUCT_NAME=" & WMSDB_SQLText(Trim(productName))
    End If

    sql="SELECT COALESCE(SUM(FACT_QTY),0) FROM WMS_ORDER_LINES WHERE " & _
        "EXTERNAL_ORDER_NO=" & WMSDB_SQLText(Trim(externalOrder)) & _
        " AND " & matchSQL & _
        " AND COALESCE(RECEIPT_EVENT_ID,'')<>''"

    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery(sql)
    If oRS.next() Then WMSRC_ReceivedBeforeDB=oRS.getDouble(1)
    WMSDB_Close
    Exit Function
EH:
    sErr="Контроль допоставки: " & CStr(Err) & " " & Error$
    On Error Resume Next
    WMSDB_Close
End Function

Function WMSRC_DraftReceived(oSh As Object,cExt As Long,cMode As Long,externalOrder As String,productCode As String,articleCode As String,productName As String) As Double
    Dim lastR As Long,r As Long
    Dim sameProduct As Boolean,modeName As String,extNo As String

    WMSRC_DraftReceived=0
    If Trim(externalOrder)="" Then Exit Function

    lastR=WMSRC_LastBusinessRow(oSh)
    For r=1 To lastR
        extNo=Trim(oSh.getCellByPosition(cExt,r).String)
        modeName=UCase(Trim(oSh.getCellByPosition(cMode,r).String))
        If UCase(extNo)=UCase(Trim(externalOrder)) And modeName="ORDER_DELIVERY" Then
            sameProduct=False
            If Trim(productCode)<>"" Then
                sameProduct=(UCase(Trim(oSh.getCellByPosition(21,r).String))=UCase(Trim(productCode)))
            ElseIf Trim(articleCode)<>"" Then
                sameProduct=(UCase(Trim(oSh.getCellByPosition(5,r).String))=UCase(Trim(articleCode)))
            Else
                sameProduct=(UCase(Trim(oSh.getCellByPosition(1,r).String))=UCase(Trim(productName)))
            End If

            If sameProduct Then
                WMSRC_DraftReceived=WMSRC_DraftReceived+oSh.getCellByPosition(6,r).Value
            End If
        End If
    Next r
End Function

Function WMSRC_NewEventID(mode As String) As String
    Randomize
    WMSRC_NewEventID="RCPT-" & UCase(Left(mode,4)) & "-" & _
        Right("0000" & CStr(Year(Now)),4) & _
        Right("00" & CStr(Month(Now)),2) & _
        Right("00" & CStr(Day(Now)),2) & "-" & _
        Right("00" & CStr(Hour(Now)),2) & _
        Right("00" & CStr(Minute(Now)),2) & _
        Right("00" & CStr(Second(Now)),2) & "-" & _
        Right("000000" & CStr(CLng(Rnd()*999999)),6)
End Function


Function WMSRC_ReceiptCountConsolidated(oDoc As Object,oSh As Object,externalOrder As String, _
    productCode As String,articleCode As String,productName As String,orderUnit As String,ByRef sErr As String) As Double
    Dim oCon As Object,oStmt As Object,oRS As Object,sql As String,matchSQL As String
    Dim postedIDs() As String,draftIDs() As String,nPosted As Long,nDraft As Long
    Dim cExt As Long,cMode As Long,cSID As Long,lastCol As Long,lastRow As Long
    Dim data As Variant,item As Variant,r As Long,i As Long,sid As String,matched As Boolean,posted As Boolean
    Dim qty As Double,totalQty As Double,detail As String,receiptUnit As String
    WMSRC_ReceiptCountConsolidated=0:sErr=""
    On Error GoTo EH
    If Trim(externalOrder)="" Then sErr="Не задана связь с заказом.":Exit Function
    If Trim(orderUnit)="" Then sErr="Не указана единица измерения заказа.":Exit Function
    cExt=Orders_FindHeader(oSh,WMSRC_H_EXTORDER)
    cMode=Orders_FindHeader(oSh,WMSRC_H_MODE)
    cSID=Orders_FindHeader(oSh,"_WMS_SourceID")
    If cExt<0 Or cMode<0 Or cSID<0 Then sErr="Не найдены служебные поля связи поступлений. Обновите структуру заказа.":Exit Function
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then GoTo Finished
    If Trim(productCode)<>"" Then
        matchSQL="UPPER(TRIM(l.PRODUCT_CODE))=" & WMSDB_SQLText(UCase(Trim(productCode)))
    ElseIf Trim(articleCode)<>"" Then
        matchSQL="UPPER(TRIM(l.SUPPLIER_ARTICLE))=" & WMSDB_SQLText(UCase(Trim(articleCode)))
    Else
        matchSQL="UPPER(TRIM(l.PRODUCT_NAME))=" & WMSDB_SQLText(UCase(Trim(productName)))
    End If
    ' Physical receipts are established by posted ledger movements, not by synced drafts.
    sql="SELECT l.SOURCE_ID,COALESCE(m.ENTRY_QTY,m.QTY),CASE WHEN m.ENTRY_QTY IS NULL THEN m.UNIT_NAME ELSE m.ENTRY_UNIT END FROM WMS_ORDER_LINES l " & _
        "JOIN WMS_STOCK_MOVEMENTS m ON m.MOVEMENT_ID='IN-' || l.SOURCE_ID " & _
        "WHERE UPPER(TRIM(l.EXTERNAL_ORDER_NO))=" & WMSDB_SQLText(UCase(Trim(externalOrder))) & _
        " AND " & matchSQL & " AND m.MOVEMENT_TYPE='IN' AND COALESCE(m.STATUS_NAME,'POSTED')='POSTED'"
    oStmt=oCon.createStatement():oRS=oStmt.executeQuery(sql)
    ReDim postedIDs(0 To 0):nPosted=0
    Do While oRS.next()
        sid=Trim(oRS.getString(1)):qty=oRS.getDouble(2)
        If sid="" Or qty<=0 Then sErr="В проведённом поступлении отсутствует идентификатор или положительное количество.":GoTo Finished
        receiptUnit=Trim(oRS.getString(3))
        If UCase(receiptUnit)<>UCase(Trim(orderUnit)) Then
            sErr="Поступление " & sid & ": единица измерения отличается от заказа. Нужен проверенный пересчёт единиц.":GoTo Finished
        End If
        For i=0 To nPosted-1
            If postedIDs(i)=sid Then sErr="Повторяется проведённое поступление " & sid:GoTo Finished
        Next i
        ReDim Preserve postedIDs(0 To nPosted)
        postedIDs(nPosted)=sid:nPosted=nPosted+1:totalQty=totalQty+qty
    Loop
    WMSDBX_CloseRS oRS:WMSDBX_CloseStmt oStmt
    lastCol=Orders_LastHeaderCol(oSh):lastRow=WMSRC_LastBusinessRow(oSh)
    If lastRow<1 Then WMSRC_ReceiptCountConsolidated=totalQty:GoTo Finished
    data=oSh.getCellRangeByPosition(0,1,lastCol,lastRow).getDataArray()
    ReDim draftIDs(0 To UBound(data)):nDraft=0
    For r=0 To UBound(data)
        item=data(r)
        If UCase(Trim(CStr(item(cExt))))=UCase(Trim(externalOrder)) And UCase(Trim(CStr(item(cMode))))="ORDER_DELIVERY" Then
            If Trim(productCode)<>"" Then
                matched=(UCase(Trim(CStr(item(21))))=UCase(Trim(productCode)))
            ElseIf Trim(articleCode)<>"" Then
                matched=(UCase(Trim(CStr(item(5))))=UCase(Trim(articleCode)))
            Else
                matched=(UCase(Trim(CStr(item(1))))=UCase(Trim(productName)))
            End If
            If matched Then
                sid=Trim(CStr(item(cSID))):posted=False
                For i=0 To nPosted-1
                    If postedIDs(i)=sid Then posted=True:Exit For
                Next i
                If Not posted Then
                    If sid<>"" Then
                        For i=0 To nDraft-1
                            If draftIDs(i)=sid Then sErr="Повторяется идентификатор черновика " & sid & ". Допоставка остановлена.":GoTo Finished
                        Next i
                        draftIDs(nDraft)=sid:nDraft=nDraft+1
                    End If
                    If VarType(item(6))=8 Then
                        If Trim(CStr(item(6)))<>"" Then sErr="Строка " & CStr(r+2) & ": количество поступления записано текстом.":GoTo Finished
                        qty=0
                    Else
                        qty=CDbl(item(6))
                    End If
                    If qty<0 Then sErr="Строка " & CStr(r+2) & ": отрицательное количество поступления.":GoTo Finished
                    If qty>0 And UCase(Trim(CStr(item(8))))<>UCase(Trim(orderUnit)) Then
                        sErr="Строка " & CStr(r+2) & ": единица измерения отличается от заказа. Нужен проверенный пересчёт единиц.":GoTo Finished
                    End If
                    totalQty=totalQty+qty
                End If
            End If
        End If
    Next r
    WMSRC_ReceiptCountConsolidated=totalQty
Finished:
    WMSDBX_CloseRS oRS:WMSDBX_CloseStmt oStmt
    WMSDB_Close
    Exit Function
EH:
    detail=CStr(Err) & " " & Error$
    sErr="Контроль допоставки: " & detail
    WMSRC_ReceiptCountConsolidated=0
    Resume Finished
End Function
