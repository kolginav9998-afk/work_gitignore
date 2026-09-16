Option Explicit
' 1.3.4: cast aggregated lot quantity at SDBC read boundary; retain reopen verification.

Global Const WMSWF_VERSION = "3.3.0-ONE-CONNECTION-RESUMABLE-BATCH"
Global gWMSWF_Busy As Boolean
Global Const WMSWF_FORM = "WMS_MANUAL_FORM"

Sub WMSWF_Install()
    Dim oDoc As Object
    oDoc=ThisComponent
    On Error GoTo EH
    oDoc.lockControllers()

    WMSWF_PrepareSheet oDoc,"Приход — Производство"
    WMSWF_PrepareSheet oDoc,"Приход — Офис"
    WMSWF_PrepareSheet oDoc,"Приход — Детали"
    WMSWF_PrepareSheet oDoc,"Расход — Производство"
    WMSWF_PrepareSheet oDoc,"Расход — Офис"
    WMSWF_PrepareSheet oDoc,"Расход — Детали"

    WMSWF_PrepareSheet oDoc,"Приход — Цех"
    WMSWF_PrepareSheet oDoc,"Расход — Цех"
    oDoc.unlockControllers()
    WMSREF_ApplyValidations oDoc
    Exit Sub
EH:
    On Error Resume Next:oDoc.unlockControllers()
    MsgBox "Рабочие листы: " & CStr(Err) & " " & Error$,16,"WMS — Workflow"
End Sub

Sub WMSWF_PrepareSheet(oDoc As Object,sName As String)
    Dim oSh As Object,h As Variant,i As Long,lastCol As Long,bg As Long
    If oDoc.Sheets.hasByName(sName) Then
        oSh=oDoc.Sheets.getByName(sName)
    Else
        oDoc.Sheets.insertNewByName(sName,oDoc.Sheets.getCount())
        oSh=oDoc.Sheets.getByName(sName)
    End If
    oSh.IsVisible=True

    h=WMSWF_Headers(sName):lastCol=UBound(h):bg=WMSWF_Color(sName)

    ' Do not touch working rows below row 5.
    oSh.getCellRangeByPosition(0,0,15,4).clearContents(1023)
    oSh.getCellRangeByPosition(0,0,15,4).CellBackColor=-1

    oSh.getCellByPosition(0,0).String=UCase(sName)
    oSh.getCellByPosition(0,1).String=WMSWF_Subtitle(sName)
    oSh.getCellByPosition(0,2).String="Заполняйте с 6-й строки. Код товара определяется автоматически. Панель действий находится выше заголовков и не перекрывает таблицу."

    If Right(sName,3)="Цех" Then
        oSh.getCellByPosition(0,1).String="Общий учёт продукции и деталей производства."
        oSh.getCellByPosition(0,2).String="Артикул необязателен. Без артикула: для существующего товара укажите внутренний код в L; новый приход без кода создаст новый товар. Расход: партии подбираются автоматически."
    End If
    oSh.getCellRangeByPosition(0,0,lastCol,0).CellBackColor=bg
    oSh.getCellRangeByPosition(0,0,lastCol,0).CharColor=16777215
    oSh.getCellRangeByPosition(0,0,lastCol,0).CharWeight=150
    oSh.getCellRangeByPosition(0,1,lastCol,2).CellBackColor=15790320

    For i=0 To lastCol:oSh.getCellByPosition(i,4).String=CStr(h(i)):Next i
    oSh.getCellRangeByPosition(0,4,lastCol,4).CellBackColor=bg
    oSh.getCellRangeByPosition(0,4,lastCol,4).CharColor=16777215
    oSh.getCellRangeByPosition(0,4,lastCol,4).CharWeight=150

    WMSWF_SetWidths oSh,sName
    WMSWF_InstallButtons oDoc,oSh
End Sub

Function WMSWF_Headers(sName As String) As Variant
    Select Case sName
        Case "Приход — Цех"
            WMSWF_Headers=Array("Дата","Наименование","Артикул (если есть)","Кол-во","Ед. изм.","Кто сдал","Место хранения","Категория","Подкатегория","Документ / чертёж","Комментарий","Внутренний код")
        Case "Расход — Цех"
            WMSWF_Headers=Array("Дата","Наименование","Артикул (если есть)","Кол-во","Ед. изм.","Кому выдано","Назначение / проект","Откуда","Партия (необязательно)","Возвратный","Комментарий","Внутренний код")
        Case "Приход — Производство"
            WMSWF_Headers=Array("Дата","Наименование","Артикул","Кол-во","Ед. изм.","Кто сдал","Место хранения","Категория","Подкатегория","Документ","Комментарий")
        Case "Приход — Офис"
            WMSWF_Headers=Array("Дата","Наименование","Артикул","Кол-во","Ед. изм.","От кого / отдел","Контур","Место хранения","Категория","Подкатегория","Документ","Комментарий")
        Case "Приход — Детали"
            WMSWF_Headers=Array("Дата","Наименование детали","Артикул","Кол-во","Ед. изм.","Кто сдал","Место хранения","Категория детали","№ чертежа / партии","Комментарий")
        Case "Расход — Производство"
            WMSWF_Headers=Array("Дата","Наименование","Артикул","Кол-во","Ед. изм.","Кому выдано","Участок / назначение","Откуда","Партия","Комментарий")
        Case "Расход — Офис"
            WMSWF_Headers=Array("Дата","Наименование","Артикул","Кол-во","Ед. изм.","Кому / отдел","Контур","Откуда","Партия","Комментарий")
        Case "Расход — Детали"
            WMSWF_Headers=Array("Дата","Наименование детали","Артикул","Кол-во","Ед. изм.","Кому выдано","Машина / проект","Откуда","Партия","Возвратный","Комментарий")
    End Select
End Function

Function WMSWF_Color(sName As String) As Long
    Select Case sName
        Case "Приход — Цех","Приход — Производство":WMSWF_Color=33792
        Case "Приход — Офис":WMSWF_Color=26367
        Case "Приход — Детали":WMSWF_Color=10040064
        Case "Расход — Цех","Расход — Производство":WMSWF_Color=10027008
        Case "Расход — Офис":WMSWF_Color=8421504
        Case "Расход — Детали":WMSWF_Color=10824234
    End Select
End Function

Function WMSWF_Subtitle(sName As String) As String
    Select Case sName
        Case "Приход — Производство":WMSWF_Subtitle="Фактическое поступление ТМЦ с производства."
        Case "Приход — Офис":WMSWF_Subtitle="Фактическое поступление ТМЦ из офиса. Контур сохраняется в документе."
        Case "Приход — Детали":WMSWF_Subtitle="Детали: поиск по точному артикулу; новый артикул создаёт новую номенклатуру."
        Case "Расход — Производство":WMSWF_Subtitle="Передача ТМЦ в производство. Одна строка = одна партия."
        Case "Расход — Офис":WMSWF_Subtitle="Передача ТМЦ в офис. Одна строка = одна партия."
        Case "Расход — Детали":WMSWF_Subtitle="Выдача деталей. Возвратные позиции попадают в контроль возвратов."
    End Select
End Function

Sub WMSWF_SetWidths(oSh As Object,sName As String)
    Dim i As Long
    For i=0 To 11:oSh.Columns.getByIndex(i).Width=3600:Next i
    oSh.Columns.getByIndex(0).Width=2500
    oSh.Columns.getByIndex(1).Width=6800
    oSh.Columns.getByIndex(2).Width=3900
    oSh.Columns.getByIndex(3).Width=2400
    oSh.Columns.getByIndex(4).Width=2200
    oSh.Columns.getByIndex(5).Width=4800
    oSh.Columns.getByIndex(6).Width=4800
    oSh.Columns.getByIndex(7).Width=4200
    oSh.Columns.getByIndex(8).Width=4300
    oSh.Columns.getByIndex(9).Width=4800
    oSh.Columns.getByIndex(10).Width=6500
End Sub

Sub WMSWF_InstallButtons(oDoc As Object,oSh As Object)
    Dim forms As Object,form As Object
    WMSWF_RemoveButtons oSh
    forms=oSh.DrawPage.Forms
    On Error Resume Next
    If forms.hasByName(WMSWF_FORM) Then forms.removeByName(WMSWF_FORM)
    On Error GoTo EH
    form=oDoc.createInstance("com.sun.star.form.component.Form"):form.Name=WMSWF_FORM
    forms.insertByName(WMSWF_FORM,form)

    ' Toolbar is intentionally above headers and spaced across columns.
    WMSWF_AddButton oDoc,oSh,form,"WMS_WF_ROW","Провести строку",0,3,3400,750,"WMSWF_ConductSelected"
    WMSWF_AddButton oDoc,oSh,form,"WMS_WF_ALL","Провести заполненные",2,3,4000,750,"WMSWF_ConductAll"
    WMSWF_AddButton oDoc,oSh,form,"WMS_WF_ACT","Создать акт",4,3,3200,750,"WMSWF_CreateAct"
    WMSWF_AddButton oDoc,oSh,form,"WMS_WF_STOCK","Остаток",6,3,2600,750,"WMSWF_OpenStock"
    WMSWF_AddButton oDoc,oSh,form,"WMS_WF_SEARCH","Поиск в базе",8,3,3000,750,"WMSWF_OpenSearch"
    WMSWF_AddButton oDoc,oSh,form,"WMS_WF_FILLART","Заполнить по артикулу",10,3,3800,750,"WMSWF_AutofillSelectedByArticle"
    WMSWF_AddButton oDoc,oSh,form,"WMS_WF_REPEAT","Повторить реквизиты",12,3,3800,750,"WMSWF_RepeatSelectedFields"
    Exit Sub
EH:
    MsgBox "Кнопки " & oSh.Name & ": " & CStr(Err) & " " & Error$,16,"WMS — Workflow"
End Sub

Sub WMSWF_AddButton(oDoc As Object,oSh As Object,oForm As Object,sName As String,sLabel As String,nCol As Long,nRow As Long,nW As Long,nH As Long,sMacro As String)
    Dim model As Object,shape As Object,a As Object,sz As Object,idx As Long
    Dim ev As New com.sun.star.script.ScriptEventDescriptor
    model=oDoc.createInstance("com.sun.star.form.component.CommandButton")
    model.Name=sName:model.Label=sLabel:model.Tabstop=False
    oForm.insertByName(sName,model):idx=oForm.Count-1
    shape=oDoc.createInstance("com.sun.star.drawing.ControlShape"):shape.Control=model
    a=oSh.getCellByPosition(nCol,nRow).Position:shape.Position=a
    sz=CreateUnoStruct("com.sun.star.awt.Size"):sz.Width=nW:sz.Height=nH:shape.Size=sz
    oSh.DrawPage.add(shape)
    ev.ListenerType="com.sun.star.awt.XActionListener":ev.EventMethod="actionPerformed":ev.ScriptType="Script"
    ev.ScriptCode="vnd.sun.star.script:Standard.WMS_20_ManualOperations." & sMacro & "?language=Basic&location=document"
    oForm.registerScriptEvent idx,ev
End Sub

Sub WMSWF_RemoveButtons(oSh As Object)
    Dim i As Long,shape As Object,ctl As Object,nm As String
    On Error Resume Next
    For i=oSh.DrawPage.getCount()-1 To 0 Step -1
        shape=oSh.DrawPage.getByIndex(i):nm="":ctl=shape.Control:nm=ctl.Name
        If Left(nm,7)="WMS_WF_" Or Left(nm,8)="WMS_MAN_" Then oSh.DrawPage.remove(shape)
    Next i
    On Error GoTo 0
End Sub

Sub WMSWF_ConductSelected()
    Dim oDoc As Object,oSh As Object,sel As Object,r As Long,sErr As String,sSaveErr As String
    Dim oCon As Object
    If Not WMSDBX_TryEnter("WORKFLOW") Then
        MsgBox "Другая операция WMS ещё выполняется. Дождитесь завершения.",48,"WMS"
        Exit Sub
    End If
    oDoc=ThisComponent:oSh=oDoc.CurrentController.ActiveSheet
    If Not WMSWF_IsWorkflowSheet(oSh.Name) Then
        WMSDBX_Leave
        Exit Sub
    End If
    On Error GoTo EH
    sel=oDoc.CurrentSelection:r=sel.CellAddress.Row
    If r<5 Then
        WMSDBX_Leave
        MsgBox "Выберите заполненную строку.",48,"WMS — Проведение"
        Exit Sub
    End If

    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then GoTo Fail
    If Not WMSARCH_EnsureSchema(oCon,sErr) Then GoTo Fail
    If Not WMSINT_EnsureSchema(oCon,sErr) Then GoTo Fail
    If Not WMSDBX_Begin(oCon,sErr) Then GoTo Fail

    If Not WMSWF_ConductRowCon(oCon,oSh,r,sErr) Then GoTo Fail
    If Not WMSDBX_Commit(oCon,sErr) Then GoTo FailAfterCommit
    If Not WMSDBX_SaveODB(sSaveErr) Then
        sErr="Firebird подтвердил операцию, но ODB не удалось сохранить. Строка сохранена на листе. Перед повтором сверьте журнал движений: " & sSaveErr
        GoTo FailAfterCommit
    End If
    WMSDB_Close
    If gWMSDBX_UnsafeConnection Then
        sErr=gWMSDBX_UnsafeReason
        GoTo FailAfterCommit
    End If
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then GoTo FailAfterCommit
    If Not WMSWF_VerifyPersistedRowCon(oCon,oSh,r,sErr) Then
        WMSDBX_BlockConnection sErr
        GoTo FailAfterCommit
    End If
    WMSWF_ClearAfterSuccess oSh,r
    WMSDBX_Leave
    On Error Resume Next
    WMSDBST_RefreshStock
    WMSREF_RefreshCache
    On Error GoTo 0
    MsgBox "Проведено, сохранено и проверено.",64,"WMS"
    Exit Sub

Fail:
    WMSDBX_RollbackQuiet oCon
FailAfterCommit:
    On Error Resume Next
    WMSDB_Close
    On Error GoTo 0
    WMSDBX_Leave
    MsgBox sErr,16,"WMS — Не проведено"
    Exit Sub
EH:
    sErr="Проведение: " & CStr(Err) & " " & Error$
    Resume Fail
End Sub

Sub WMSWF_ConductAll()
    Dim oDoc As Object,oSh As Object,oCon As Object
    Dim r As Long,lastR As Long,rowCount As Long,i As Long
    Dim rows() As Long
    Dim sErr As String,sSaveErr As String,rep As String

    If Not WMSDBX_TryEnter("WORKFLOW_BULK") Then
        MsgBox "Другая операция WMS ещё выполняется. Дождитесь завершения.",48,"WMS"
        Exit Sub
    End If
    oDoc=ThisComponent:oSh=oDoc.CurrentController.ActiveSheet
    If Not WMSWF_IsWorkflowSheet(oSh.Name) Then
        WMSDBX_Leave
        Exit Sub
    End If
    On Error GoTo EH
    oDoc.lockControllers()

    lastR=WMSWF_LastRow(oSh)
    For r=5 To lastR
        If WMSWF_RowHasData(oSh,r) Then
            rowCount=rowCount+1
            ReDim Preserve rows(1 To rowCount)
            rows(rowCount)=r
        End If
    Next r
    If rowCount=0 Then
        oDoc.unlockControllers()
        WMSDBX_Leave
        MsgBox "Нет заполненных строк для проведения.",48,"WMS"
        Exit Sub
    End If

    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then GoTo Fail
    ' Схема проверяется один раз на весь пакет, а не для каждой строки.
    If Not WMSARCH_EnsureSchema(oCon,sErr) Then GoTo Fail
    If Not WMSINT_EnsureSchema(oCon,sErr) Then GoTo Fail
    If Not WMSDBX_Begin(oCon,sErr) Then GoTo Fail

    ' Весь пакет — одна транзакция. Ошибка одной строки откатывает весь новый пакет.
    For i=1 To rowCount
        r=rows(i)
        sErr=""
        If Not WMSWF_ConductRowCon(oCon,oSh,r,sErr) Then
            rep="Строка " & CStr(r+1) & ": " & sErr
            GoTo Fail
        End If
    Next i

    If Not WMSDBX_Commit(oCon,sErr) Then GoTo FailAfterCommit
    ' ODB сериализуется один раз на весь пакет.
    If Not WMSDBX_SaveODB(sSaveErr) Then
        sErr="Транзакция Firebird завершена, но контейнер ODB не удалось сохранить. Строки НЕ очищены: " & sSaveErr
        GoTo FailAfterCommit
    End If
    WMSDB_Close
    If gWMSDBX_UnsafeConnection Then
        sErr=gWMSDBX_UnsafeReason
        GoTo FailAfterCommit
    End If
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then GoTo FailAfterCommit
    For i=1 To rowCount
        r=rows(i)
        If Not WMSWF_VerifyPersistedRowCon(oCon,oSh,r,sErr) Then
            WMSDBX_BlockConnection sErr
            GoTo FailAfterCommit
        End If
    Next i

    ' Clear only after reopening the ODB and verifying EVERY submitted row.
    For i=1 To rowCount
        WMSWF_ClearAfterSuccess oSh,rows(i)
    Next i

    On Error Resume Next
    oDoc.unlockControllers()
    On Error GoTo 0
    WMSDBX_Leave
    On Error Resume Next
    WMSDBST_RefreshStock
    WMSREF_RefreshCache
    On Error GoTo 0
    MsgBox "Проведено строк: " & CStr(rowCount) & Chr(10) & _
           "Все строки повторно прочитаны из сохранённой ODB. Проверка 1.4.0.",64,"WMS — Пакетное проведение"
    Exit Sub

Fail:
    WMSDBX_RollbackQuiet oCon
    If rep="" Then rep=sErr
FailAfterCommit:
    On Error Resume Next
    WMSDB_Close
    oDoc.unlockControllers()
    On Error GoTo 0
    WMSDBX_Leave
    MsgBox "Пакет НЕ очищен. Результат не подтверждён полностью. Перед повторным вводом требуется сверка базы." & Chr(10) & Chr(10) & rep & IIf(sErr<>rep,Chr(10) & sErr,""),16,"WMS — Проведение остановлено"
    Exit Sub
EH:
    sErr="Пакетное проведение: " & CStr(Err) & " " & Error$
    Resume Fail
End Sub

Function WMSWF_ConductRow(oDoc As Object,oSh As Object,r As Long,ByRef sErr As String) As Boolean
    Dim oCon As Object,sSaveErr As String
    WMSWF_ConductRow=False:sErr=""
    On Error GoTo EH
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then Exit Function
    If Not WMSARCH_EnsureSchema(oCon,sErr) Then GoTo Fail
    If Not WMSINT_EnsureSchema(oCon,sErr) Then GoTo Fail
    If Not WMSDBX_Begin(oCon,sErr) Then GoTo Fail
    If Not WMSWF_ConductRowCon(oCon,oSh,r,sErr) Then GoTo Fail
    If Not WMSDBX_Commit(oCon,sErr) Then GoTo FailCommitted
    If Not WMSDBX_SaveODB(sSaveErr) Then
        sErr="ODB не сохранён: " & sSaveErr
        GoTo FailCommitted
    End If
    WMSDB_Close
    If gWMSDBX_UnsafeConnection Then
        sErr=gWMSDBX_UnsafeReason
        GoTo FailCommitted
    End If
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then GoTo FailCommitted
    If Not WMSWF_VerifyPersistedRowCon(oCon,oSh,r,sErr) Then
        WMSDBX_BlockConnection sErr
        GoTo FailCommitted
    End If
    WMSWF_ClearAfterSuccess oSh,r
    WMSWF_ConductRow=True
    Exit Function
Fail:
    WMSDBX_RollbackQuiet oCon
FailCommitted:
    On Error Resume Next
    WMSDB_Close
    On Error GoTo 0
    Exit Function
EH:
    sErr="ConductRow: " & CStr(Err) & " " & Error$
    Resume Fail
End Function

Function WMSWF_ConductRowCon(oCon As Object,oSh As Object,r As Long,ByRef sErr As String) As Boolean
    Dim dir As String,channel As String
    WMSWF_Context oSh.Name,dir,channel
    If dir="IN" Then
        WMSWF_ConductRowCon=WMSWF_ReceiptRowCon(oCon,oSh,r,channel,sErr)
    ElseIf dir="OUT" Then
        WMSWF_ConductRowCon=WMSWF_IssueRowCon(oCon,oSh,r,channel,sErr)
    Else
        sErr="Неизвестный рабочий лист."
        WMSWF_ConductRowCon=False
    End If
End Function

Function WMSWF_ReceiptRow(oDoc As Object,oSh As Object,r As Long,channel As String,ByRef sErr As String) As Boolean
    WMSWF_ReceiptRow=WMSWF_ConductRow(oDoc,oSh,r,sErr)
End Function

Function WMSWF_ReceiptRowCon(oCon As Object,oSh As Object,r As Long,channel As String,ByRef sErr As String) As Boolean
    Dim dt As Double,nm As String,art As String,unitName As String,person As String,loc As String,cat As String,subcat As String,docNo As String,note As String
    Dim qty As Double,code As String,lotID As String,docID As String,movID As String,opID As String
    Dim oStmt As Object,nExisting As Long

    WMSWF_ReceiptRowCon=False:sErr=""
    On Error GoTo EH
    WMSWF_ReadReceipt oSh,r,channel,dt,nm,art,qty,unitName,person,loc,cat,subcat,docNo,note
    If nm="" Then sErr="Не заполнено наименование.":Exit Function
    If art="" And channel<>"WORKSHOP" Then sErr="Не заполнен артикул.":Exit Function
    If qty<=0 Then sErr="Количество должно быть больше 0.":Exit Function
    If unitName="" Then sErr="Не заполнена единица измерения.":Exit Function
    If loc="" Then sErr="Не заполнено место хранения.":Exit Function
    If dt<=0 Then dt=Date

    opID=WMSWF_GetOrCreateTokenCon(oCon,oSh,r,"IN",sErr)
    If sErr<>"" Or opID="" Then GoTo Done
    nExisting=WMSDBX_CountByOperation(oCon,opID,sErr)
    If sErr<>"" Then GoTo Done
    If nExisting>0 Then
        sErr="Строка " & CStr(r+1) & ": номер операции уже есть в базе (" & opID & "). Возможно повторное проведение или совпадение старых номеров. Строка не очищена. Сверьте журнал движений перед повторным вводом."
        GoTo Done
    End If

    code=WMSWF_ResolveRowProductCon(oCon,oSh,r,art,nm,unitName,loc,True,sErr)
    If sErr<>"" Or code="" Then GoTo Done

    docID="DOC-" & opID:lotID="LOT-" & opID:movID="MOV-" & opID
    oStmt=oCon.createStatement()
    oStmt.executeUpdate("INSERT INTO WMS_DOCUMENTS (DOC_ID,DOC_TYPE,DOC_NO,DOC_DATE,FLOW_CHANNEL,PARTY_FROM,STATUS_NAME,COMMENT_TEXT) VALUES (" & _
        WMSDB_SQLText(docID) & ",'RECEIPT'," & WMSDB_SQLText(docNo) & "," & WMSWF_SQLDate(dt) & "," & WMSDB_SQLText(channel) & "," & WMSDB_SQLText(person) & ",'POSTED'," & WMSDB_SQLText(note) & ")")
    oStmt.executeUpdate("INSERT INTO WMS_DOCUMENT_LINES (DOC_ID,LINE_NO,PRODUCT_CODE,ARTICLE_CODE,PRODUCT_NAME,QTY,UNIT_NAME,LOCATION_NAME,LOT_ID,NOTE_TEXT) VALUES (" & _
        WMSDB_SQLText(docID) & ",1," & WMSDB_SQLText(code) & "," & WMSDB_SQLText(art) & "," & WMSDB_SQLText(nm) & "," & WMSWF_Num(qty) & "," & WMSDB_SQLText(unitName) & "," & WMSDB_SQLText(loc) & "," & WMSDB_SQLText(lotID) & "," & WMSDB_SQLText(note) & ")")
    oStmt.executeUpdate("INSERT INTO WMS_STOCK_LOTS (LOT_ID,PRODUCT_CODE,SOURCE_ID,SOURCE_TYPE,DOC_QTY,DOC_UNIT,BASE_QTY,BASE_UNIT,LOCATION_NAME,ORIGIN_NAME,DEST_CATEGORY,DEST_SUBCATEGORY,RECEIPT_DATE,NOTE_TEXT,ACTIVE_FLAG) VALUES (" & _
        WMSDB_SQLText(lotID) & "," & WMSDB_SQLText(code) & "," & WMSDB_SQLText(docID) & ",'MANUAL_RECEIPT'," & WMSWF_Num(qty) & "," & WMSDB_SQLText(unitName) & "," & WMSWF_Num(qty) & "," & WMSDB_SQLText(unitName) & "," & WMSDB_SQLText(loc) & "," & WMSDB_SQLText(channel) & "," & WMSDB_SQLText(cat) & "," & WMSDB_SQLText(subcat) & "," & WMSWF_SQLDate(dt) & "," & WMSDB_SQLText(note) & ",1)")
    oStmt.executeUpdate("INSERT INTO WMS_LOT_UNITS (LOT_ID,UNIT_NAME,FACTOR_TO_BASE,UNIT_ROLE) VALUES (" & WMSDB_SQLText(lotID) & "," & WMSDB_SQLText(unitName) & ",1,'BASE')")
    oStmt.executeUpdate("INSERT INTO WMS_STOCK_MOVEMENTS (MOVEMENT_ID,SOURCE_ID,SOURCE_TYPE,MOVEMENT_TYPE,PRODUCT_CODE,PRODUCT_NAME,QTY,UNIT_NAME,LOCATION_NAME,MOVEMENT_DATE,NOTE_TEXT,ORIGIN_NAME,DEST_CATEGORY,DEST_SUBCATEGORY,LOT_ID,ENTRY_QTY,ENTRY_UNIT,BASE_QTY,BASE_UNIT,STATUS_NAME,FLOW_CHANNEL,OPERATION_ID) VALUES (" & _
        WMSDB_SQLText(movID) & "," & WMSDB_SQLText(docID) & ",'MANUAL_RECEIPT','IN'," & WMSDB_SQLText(code) & "," & WMSDB_SQLText(nm) & "," & WMSWF_Num(qty) & "," & WMSDB_SQLText(unitName) & "," & WMSDB_SQLText(loc) & "," & WMSWF_SQLDate(dt) & "," & WMSDB_SQLText(note) & "," & WMSDB_SQLText(channel) & "," & WMSDB_SQLText(cat) & "," & WMSDB_SQLText(subcat) & "," & WMSDB_SQLText(lotID) & "," & WMSWF_Num(qty) & "," & WMSDB_SQLText(unitName) & "," & WMSWF_Num(qty) & "," & WMSDB_SQLText(unitName) & ",'POSTED'," & WMSDB_SQLText(channel) & "," & WMSDB_SQLText(opID) & ")")

    If Not WMSDBX_VerifyMovement(oCon,movID,lotID,qty,sErr) Then GoTo Done
    WMSWF_AddRefCon oCon,"UNIT",unitName,unitName,""
    WMSWF_AddRefCon oCon,"LOCATION",loc,loc,""
    If cat<>"" Then WMSWF_AddRefCon oCon,"CATEGORY",cat,cat,""
    If subcat<>"" Then WMSWF_AddRefCon oCon,"SUBCATEGORY",subcat,subcat,cat
    WMSWF_ReceiptRowCon=True
Done:
    WMSDBX_CloseStmt oStmt
    Exit Function
EH:
    sErr="Приход: " & CStr(Err) & " " & Error$
    Resume Done
End Function

Function WMSWF_IssueRow(oDoc As Object,oSh As Object,r As Long,channel As String,ByRef sErr As String) As Boolean
    WMSWF_IssueRow=WMSWF_ConductRow(oDoc,oSh,r,sErr)
End Function

Function WMSWF_IssueRowCon(oCon As Object,oSh As Object,r As Long,channel As String,ByRef sErr As String) As Boolean
    Dim dt As Double,nm As String,art As String,unitName As String,person As String,dest As String,loc As String,lotID As String,note As String,ret As String
    Dim qty As Double,code As String,available As Double,baseQty As Double,baseUnit As String,origin As String,cat As String,subcat As String
    Dim oStmt As Object,docID As String,movID As String,issueID As String,opID As String,nExisting As Long

    Dim plan As Variant,part As Long,piece As Variant
    WMSWF_IssueRowCon=False:sErr=""
    On Error GoTo EH
    WMSWF_ReadIssue oSh,r,channel,dt,nm,art,qty,unitName,person,dest,loc,lotID,ret,note
    If nm="" Then sErr="Не заполнено наименование.":Exit Function
    If art="" And channel<>"WORKSHOP" Then sErr="Не заполнен артикул.":Exit Function
    If qty<=0 Then sErr="Количество должно быть больше 0.":Exit Function
    If unitName="" Then sErr="Не заполнена единица измерения.":Exit Function
    If person="" Then sErr="Не заполнено кому выдано.":Exit Function
    If loc="" Then sErr="Не заполнено Откуда.":Exit Function
    If dt<=0 Then dt=Date

    opID=WMSWF_GetOrCreateTokenCon(oCon,oSh,r,"OUT",sErr)
    If sErr<>"" Or opID="" Then GoTo Done
    nExisting=WMSDBX_CountByOperation(oCon,opID,sErr)
    If sErr<>"" Then GoTo Done
    If nExisting>0 Then
        sErr="Строка " & CStr(r+1) & ": номер операции уже есть в базе (" & opID & "). Возможно повторное проведение или совпадение старых номеров. Строка не очищена. Сверьте журнал движений перед повторным вводом."
        GoTo Done
    End If

    code=WMSWF_ResolveRowProductCon(oCon,oSh,r,art,nm,unitName,loc,False,sErr)
    If sErr<>"" Or code="" Then GoTo Done
    plan=WMSWF_PlanIssueCon(oCon,code,loc,lotID,qty,unitName,sErr)
    If sErr<>"" Then GoTo Done

    docID="DOC-" & opID
    oStmt=oCon.createStatement()
    oStmt.executeUpdate("INSERT INTO WMS_DOCUMENTS (DOC_ID,DOC_TYPE,DOC_DATE,FLOW_CHANNEL,PARTY_TO,STATUS_NAME,COMMENT_TEXT) VALUES (" & WMSDB_SQLText(docID) & ",'ISSUE'," & WMSWF_SQLDate(dt) & "," & WMSDB_SQLText(channel) & "," & WMSDB_SQLText(person) & ",'POSTED'," & WMSDB_SQLText(note) & ")")
    For part=0 To UBound(plan)
        piece=plan(part)
        lotID=CStr(piece(0)):qty=CDbl(piece(1)):baseQty=CDbl(piece(2))
        baseUnit=CStr(piece(3)):origin=CStr(piece(4)):cat=CStr(piece(5)):subcat=CStr(piece(6))
        movID="MOV-" & opID & "-" & CStr(part+1):issueID="ISS-" & opID & "-" & CStr(part+1)
    oStmt.executeUpdate("INSERT INTO WMS_DOCUMENT_LINES (DOC_ID,LINE_NO,PRODUCT_CODE,ARTICLE_CODE,PRODUCT_NAME,QTY,UNIT_NAME,LOCATION_NAME,LOT_ID,RETURNABLE,NOTE_TEXT) VALUES (" & WMSDB_SQLText(docID) & "," & CStr(part+1) & "," & WMSDB_SQLText(code) & "," & WMSDB_SQLText(art) & "," & WMSDB_SQLText(nm) & "," & WMSWF_Num(qty) & "," & WMSDB_SQLText(unitName) & "," & WMSDB_SQLText(loc) & "," & WMSDB_SQLText(lotID) & "," & WMSDB_SQLText(ret) & "," & WMSDB_SQLText(note) & ")")
    oStmt.executeUpdate("INSERT INTO WMS_ISSUES (SOURCE_ID,PRODUCT_CODE,PRODUCT_NAME,ISSUE_QTY,UNIT_NAME,EMPLOYEE_NAME,ISSUE_DATE,SOURCE_LOCATION,RETURNABLE,RETURNED_QTY,NOTE_TEXT,ROW_MODE,ROW_STATE,RETURN_STATE,ALLOC_CATEGORY,ALLOC_SUBCATEGORY,STOCK_ORIGIN,LOT_ID,ENTRY_QTY,ENTRY_UNIT,BASE_QTY,BASE_UNIT) VALUES (" & WMSDB_SQLText(issueID) & "," & WMSDB_SQLText(code) & "," & WMSDB_SQLText(nm) & "," & WMSWF_Num(qty) & "," & WMSDB_SQLText(unitName) & "," & WMSDB_SQLText(person) & "," & WMSWF_SQLDate(dt) & "," & WMSDB_SQLText(loc) & "," & WMSDB_SQLText(ret) & ",0," & WMSDB_SQLText(note) & "," & WMSDB_SQLText(channel) & ",'POSTED'," & WMSDB_SQLText(IIf(UCase(ret)="ДА","OPEN","N/A")) & "," & WMSDB_SQLText(cat) & "," & WMSDB_SQLText(subcat) & "," & WMSDB_SQLText(origin) & "," & WMSDB_SQLText(lotID) & "," & WMSWF_Num(qty) & "," & WMSDB_SQLText(unitName) & "," & WMSWF_Num(baseQty) & "," & WMSDB_SQLText(baseUnit) & ")")
    oStmt.executeUpdate("INSERT INTO WMS_STOCK_MOVEMENTS (MOVEMENT_ID,SOURCE_ID,SOURCE_TYPE,MOVEMENT_TYPE,PRODUCT_CODE,PRODUCT_NAME,QTY,UNIT_NAME,LOCATION_NAME,MOVEMENT_DATE,NOTE_TEXT,ORIGIN_NAME,DEST_CATEGORY,DEST_SUBCATEGORY,LOT_ID,ENTRY_QTY,ENTRY_UNIT,BASE_QTY,BASE_UNIT,STATUS_NAME,FLOW_CHANNEL,OPERATION_ID) VALUES (" & WMSDB_SQLText(movID) & "," & WMSDB_SQLText(issueID) & ",'MANUAL_ISSUE','OUT'," & WMSDB_SQLText(code) & "," & WMSDB_SQLText(nm) & "," & WMSWF_Num(-baseQty) & "," & WMSDB_SQLText(baseUnit) & "," & WMSDB_SQLText(loc) & "," & WMSWF_SQLDate(dt) & "," & WMSDB_SQLText(note) & "," & WMSDB_SQLText(origin) & "," & WMSDB_SQLText(cat) & "," & WMSDB_SQLText(subcat) & "," & WMSDB_SQLText(lotID) & "," & WMSWF_Num(-qty) & "," & WMSDB_SQLText(unitName) & "," & WMSWF_Num(-baseQty) & "," & WMSDB_SQLText(baseUnit) & ",'POSTED'," & WMSDB_SQLText(channel) & "," & WMSDB_SQLText(opID) & ")")

    If Not WMSDBX_VerifyMovement(oCon,movID,lotID,-baseQty,sErr) Then GoTo Done
    Next part
    WMSWF_AddRefCon oCon,"EMPLOYEE",person,person,""
    WMSWF_IssueRowCon=True
Done:
    WMSDBX_CloseStmt oStmt
    Exit Function
EH:
    sErr="Расход: " & CStr(Err) & " " & Error$
    Resume Done
End Function

Sub WMSWF_ReadReceipt(oSh As Object,r As Long,channel As String,ByRef dt As Double,ByRef nm As String,ByRef art As String,ByRef qty As Double,ByRef unitName As String,ByRef person As String,ByRef loc As String,ByRef cat As String,ByRef subcat As String,ByRef docNo As String,ByRef note As String)
    dt=oSh.getCellByPosition(0,r).Value
    nm=Trim(oSh.getCellByPosition(1,r).String):art=Trim(oSh.getCellByPosition(2,r).String)
    qty=oSh.getCellByPosition(3,r).Value:unitName=Trim(oSh.getCellByPosition(4,r).String)
    person=Trim(oSh.getCellByPosition(5,r).String)
    If channel="OFFICE" Then
        loc=Trim(oSh.getCellByPosition(7,r).String):cat=Trim(oSh.getCellByPosition(8,r).String):subcat=Trim(oSh.getCellByPosition(9,r).String)
        docNo=Trim(oSh.getCellByPosition(10,r).String):note="Контур: " & Trim(oSh.getCellByPosition(6,r).String) & " | " & Trim(oSh.getCellByPosition(11,r).String)
    ElseIf channel="PARTS" Then
        loc=Trim(oSh.getCellByPosition(6,r).String):cat=Trim(oSh.getCellByPosition(7,r).String):subcat=""
        docNo=Trim(oSh.getCellByPosition(8,r).String):note=Trim(oSh.getCellByPosition(9,r).String)
    Else
        loc=Trim(oSh.getCellByPosition(6,r).String):cat=Trim(oSh.getCellByPosition(7,r).String):subcat=Trim(oSh.getCellByPosition(8,r).String)
        docNo=Trim(oSh.getCellByPosition(9,r).String):note=Trim(oSh.getCellByPosition(10,r).String)
    End If
End Sub

Sub WMSWF_ReadIssue(oSh As Object,r As Long,channel As String,ByRef dt As Double,ByRef nm As String,ByRef art As String,ByRef qty As Double,ByRef unitName As String,ByRef person As String,ByRef dest As String,ByRef loc As String,ByRef lotID As String,ByRef ret As String,ByRef note As String)
    dt=oSh.getCellByPosition(0,r).Value:nm=Trim(oSh.getCellByPosition(1,r).String):art=Trim(oSh.getCellByPosition(2,r).String)
    qty=oSh.getCellByPosition(3,r).Value:unitName=Trim(oSh.getCellByPosition(4,r).String):person=Trim(oSh.getCellByPosition(5,r).String)
    If channel="OFFICE" Then
        dest=Trim(oSh.getCellByPosition(6,r).String):loc=Trim(oSh.getCellByPosition(7,r).String):lotID=Trim(oSh.getCellByPosition(8,r).String):ret="Нет":note=Trim(oSh.getCellByPosition(9,r).String)
    ElseIf channel="PARTS" Or channel="WORKSHOP" Then
        dest=Trim(oSh.getCellByPosition(6,r).String):loc=Trim(oSh.getCellByPosition(7,r).String):lotID=Trim(oSh.getCellByPosition(8,r).String):ret=Trim(oSh.getCellByPosition(9,r).String):note=Trim(oSh.getCellByPosition(10,r).String)
        If ret="" Then ret="Нет"
    Else
        dest=Trim(oSh.getCellByPosition(6,r).String):loc=Trim(oSh.getCellByPosition(7,r).String):lotID=Trim(oSh.getCellByPosition(8,r).String):ret="Нет":note=Trim(oSh.getCellByPosition(9,r).String)
    End If
    If dest<>"" Then note=dest & IIf(note="",""," | " & note)
End Sub

Function WMSWF_ResolveProductCon(oCon As Object,art As String,nm As String,unitName As String,loc As String,allowCreate As Boolean,ByRef sErr As String) As String
    Dim oStmt As Object,oRS As Object,sql As String,n As Long,code As String,generated As String
    WMSWF_ResolveProductCon="":sErr=""
    On Error GoTo EH

    oStmt=oCon.createStatement()
    sql="SELECT DISTINCT PRODUCT_CODE FROM WMS_PRODUCT_ALIASES WHERE ACTIVE_FLAG=1 AND UPPER(TRIM(ARTICLE_CODE))=UPPER(TRIM(" & WMSDB_SQLText(art) & "))"
    oRS=oStmt.executeQuery(sql)
    Do While oRS.next()
        n=n+1
        If n=1 Then code=oRS.getString(1)
        If n>1 Then Exit Do
    Loop
    WMSDBX_CloseRS oRS
    If n>1 Then sErr="Артикул неоднозначен: найден у нескольких товаров.":GoTo Done
    If n=1 Then WMSWF_ResolveProductCon=code:GoTo Done

    n=0
    oRS=oStmt.executeQuery("SELECT PRODUCT_CODE FROM WMS_PRODUCTS WHERE ACTIVE_FLAG=1 AND UPPER(TRIM(COALESCE(SUPPLIER_ARTICLE,'')))=UPPER(TRIM(" & WMSDB_SQLText(art) & "))")
    Do While oRS.next()
        n=n+1
        If n=1 Then code=oRS.getString(1)
        If n>1 Then Exit Do
    Loop
    WMSDBX_CloseRS oRS
    If n>1 Then sErr="Артикул неоднозначен в номенклатуре.":GoTo Done
    If n=1 Then
        WMSWF_EnsureAliasCon oCon,code,art,sErr
        If sErr="" Then WMSWF_ResolveProductCon=code
        GoTo Done
    End If

    If Not allowCreate Then sErr="Артикул не найден в номенклатуре. Сначала оформите приход.":GoTo Done
    generated=WMSARCH_NextProductCodeCon(oCon,sErr)
    If sErr<>"" Then GoTo Done
    oStmt.executeUpdate("INSERT INTO WMS_PRODUCTS (PRODUCT_CODE,PRODUCT_NAME,UNIT_NAME,DEFAULT_LOCATION,SUPPLIER_ARTICLE,ACTIVE_FLAG,ITEM_ID) VALUES (" & _
        WMSDB_SQLText(generated) & "," & WMSDB_SQLText(nm) & "," & WMSDB_SQLText(unitName) & "," & WMSDB_SQLText(loc) & "," & WMSDB_SQLText(art) & ",1," & WMSDB_SQLText("ITEM-" & generated) & ")")
    WMSWF_EnsureAliasCon oCon,generated,art,sErr
    If sErr="" Then WMSWF_ResolveProductCon=generated
Done:
    WMSDBX_CloseRS oRS
    WMSDBX_CloseStmt oStmt
    Exit Function
EH:
    sErr="Разрешение товара: " & CStr(Err) & " " & Error$
    Resume Done
End Function


Sub WMSWF_EnsureAliasCon(oCon As Object,code As String,art As String,ByRef sErr As String)
    Dim oStmt As Object,oRS As Object,n As Long,aliasID As String
    sErr=""
    On Error GoTo EH
    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery("SELECT COUNT(*) FROM WMS_PRODUCT_ALIASES WHERE PRODUCT_CODE=" & WMSDB_SQLText(code) & " AND UPPER(TRIM(ARTICLE_CODE))=UPPER(TRIM(" & WMSDB_SQLText(art) & "))")
    If oRS.next() Then n=oRS.getLong(1)
    WMSDBX_CloseRS oRS
    If n=0 Then
        aliasID=WMSWF_NewID("ALIAS")
        oStmt.executeUpdate("INSERT INTO WMS_PRODUCT_ALIASES (ALIAS_ID,PRODUCT_CODE,ARTICLE_CODE,ALIAS_TYPE,FACTOR_TO_BASE,ACTIVE_FLAG) VALUES (" & _
            WMSDB_SQLText(aliasID) & "," & WMSDB_SQLText(code) & "," & WMSDB_SQLText(art) & ",'SUPPLIER_ARTICLE',1,1)")
    End If
Done:
    WMSDBX_CloseRS oRS
    WMSDBX_CloseStmt oStmt
    Exit Sub
EH:
    sErr="Сопоставление артикула: " & CStr(Err) & " " & Error$
    Resume Done
End Sub


Function WMSWF_SelectLotCon(oCon As Object,code As String,loc As String,ByRef lotID As String,entryQty As Double,entryUnit As String,ByRef available As Double,ByRef baseQty As Double,ByRef baseUnit As String,ByRef origin As String,ByRef cat As String,ByRef subcat As String,ByRef sErr As String) As Boolean
    Dim oStmt As Object,oRS As Object,sql As String,n As Long,list As String,q As Double,factor As Double
    Dim oneLot As String,oneUnit As String,oneOrigin As String,oneCat As String,oneSub As String
    WMSWF_SelectLotCon=False:sErr=""
    On Error GoTo EH
    oStmt=oCon.createStatement()

    If Trim(lotID)<>"" Then
        sql="SELECT LOT_ID,CAST(QTY AS DOUBLE PRECISION),UNIT_NAME,ORIGIN_NAME,CATEGORY_NAME,SUBCATEGORY_NAME FROM WMS_V_LOT_BALANCE WHERE LOT_ID=" & WMSDB_SQLText(lotID) & _
            " AND PRODUCT_CODE=" & WMSDB_SQLText(code) & " AND LOCATION_NAME=" & WMSDB_SQLText(loc) & " AND QTY>0"
    Else
        sql="SELECT LOT_ID,CAST(QTY AS DOUBLE PRECISION),UNIT_NAME,ORIGIN_NAME,CATEGORY_NAME,SUBCATEGORY_NAME FROM WMS_V_LOT_BALANCE WHERE PRODUCT_CODE=" & WMSDB_SQLText(code) & _
            " AND LOCATION_NAME=" & WMSDB_SQLText(loc) & " AND QTY>0 ORDER BY RECEIPT_DATE,LOT_ID"
    End If

    oRS=oStmt.executeQuery(sql)
    Do While oRS.next()
        n=n+1:q=oRS.getDouble(2)
        If n=1 Then
            oneLot=oRS.getString(1):available=q:oneUnit=oRS.getString(3):oneOrigin=oRS.getString(4):oneCat=oRS.getString(5):oneSub=oRS.getString(6)
        End If
        If n<=10 Then list=list & oRS.getString(1) & " | доступно " & WMSWF_Num(q) & " " & oRS.getString(3) & Chr(10)
    Loop
    WMSDBX_CloseRS oRS

    If n=0 Then sErr="Нет доступной партии товара в месте хранения '" & loc & "'.":GoTo Done
    If Trim(lotID)="" And n>1 Then
        sErr="Найдено несколько партий. Укажите партию в колонке «Партия»:" & Chr(10) & list
        GoTo Done
    End If

    lotID=oneLot:baseUnit=oneUnit:origin=oneOrigin:cat=oneCat:subcat=oneSub
    If UCase(Trim(entryUnit))=UCase(Trim(baseUnit)) Then
        baseQty=entryQty
    Else
        factor=WMSWF_LotFactorCon(oCon,lotID,entryUnit,sErr)
        If sErr<>"" Then GoTo Done
        If factor<=0 Then sErr="Для партии нет коэффициента единицы '" & entryUnit & "' к базовой единице '" & baseUnit & "'.":GoTo Done
        baseQty=entryQty*factor
    End If
    If available+0.000001<baseQty Then
        sErr="В партии недостаточно. Доступно " & WMSWF_Num(available) & " " & baseUnit & ", требуется " & WMSWF_Num(baseQty) & " " & baseUnit & "."
        GoTo Done
    End If
    WMSWF_SelectLotCon=True
Done:
    WMSDBX_CloseRS oRS
    WMSDBX_CloseStmt oStmt
    Exit Function
EH:
    sErr="Выбор партии: " & CStr(Err) & " " & Error$
    Resume Done
End Function


Function WMSWF_LotFactorCon(oCon As Object,lotID As String,unitName As String,ByRef sErr As String) As Double
    Dim oStmt As Object,oRS As Object
    WMSWF_LotFactorCon=0:sErr=""
    On Error GoTo EH
    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery("SELECT FACTOR_TO_BASE FROM WMS_LOT_UNITS WHERE LOT_ID=" & WMSDB_SQLText(lotID) & " AND UPPER(TRIM(UNIT_NAME))=UPPER(TRIM(" & WMSDB_SQLText(unitName) & "))")
    If oRS.next() Then WMSWF_LotFactorCon=oRS.getDouble(1)
Done:
    WMSDBX_CloseRS oRS
    WMSDBX_CloseStmt oStmt
    Exit Function
EH:
    sErr="Коэффициент партии: " & CStr(Err) & " " & Error$
    Resume Done
End Function


Function WMSWF_VerifyMovementAndLotCon(oCon As Object,movID As String,lotID As String,expectedQty As Double,ByRef sErr As String) As Boolean
    WMSWF_VerifyMovementAndLotCon=WMSDBX_VerifyMovement(oCon,movID,lotID,expectedQty,sErr)
End Function

Sub WMSWF_AddRefCon(oCon As Object,typ As String,code As String,nm As String,parent As String)
    Dim oStmt As Object,oRS As Object,n As Long
    If Trim(code)="" Then Exit Sub
    On Error GoTo Done
    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery("SELECT COUNT(*) FROM WMS_REFERENCE WHERE REF_TYPE=" & WMSDB_SQLText(UCase(typ)) & " AND REF_CODE=" & WMSDB_SQLText(code))
    If oRS.next() Then n=oRS.getLong(1)
    WMSDBX_CloseRS oRS
    If n=0 Then oStmt.executeUpdate("INSERT INTO WMS_REFERENCE (REF_TYPE,REF_CODE,REF_NAME,PARENT_CODE) VALUES (" & WMSDB_SQLText(UCase(typ)) & "," & WMSDB_SQLText(code) & "," & WMSDB_SQLText(nm) & "," & WMSDB_SQLText(parent) & ")")
Done:
    WMSDBX_CloseRS oRS
    WMSDBX_CloseStmt oStmt
End Sub


Function WMSWF_GetOrCreateTokenCon(oCon As Object,oSh As Object,r As Long,directionName As String,ByRef sErr As String) As String
    Dim c As Object,t As String,oStmt As Object,oRS As Object,nExisting As Long
    WMSWF_GetOrCreateTokenCon="":sErr=""
    On Error GoTo EH
    c=oSh.getCellByPosition(50,r)
    t=Trim(c.String)
    ' Retire legacy time tokens only when no movement with that token exists.
    If Left(t,3)="WF-" Then
        nExisting=WMSDBX_CountByOperation(oCon,t,sErr)
        If sErr<>"" Then GoTo Done
        If nExisting=0 Then t=""
    End If
    If t="" Then
        oStmt=oCon.createStatement()
        oRS=oStmt.executeQuery("SELECT UUID_TO_CHAR(GEN_UUID()) FROM RDB$DATABASE")
        If Not oRS.next() Then
            sErr="Не удалось получить идентификатор операции. Строка не проведена."
            GoTo Done
        End If
        t=Trim(oRS.getString(1))
        If Len(t)<>36 Then
            sErr="Некорректный идентификатор операции. Строка не проведена."
            GoTo Done
        End If
        t="WF2-" & directionName & "-" & t
        c.String=t
        oSh.Columns.getByIndex(50).IsVisible=False
    End If
    WMSWF_GetOrCreateTokenCon=t
Done:
    WMSDBX_CloseRS oRS
    WMSDBX_CloseStmt oStmt
    Exit Function
EH:
    sErr="Идентификатор операции: " & CStr(Err) & " " & Error$
    Resume Done
End Function

Sub WMSWF_ClearAfterSuccess(oSh As Object,r As Long)
    On Error Resume Next
    oSh.getCellRangeByPosition(0,r,11,r).clearContents(1023)
    oSh.getCellByPosition(50,r).String=""
    On Error GoTo 0
End Sub

Sub WMSWF_Context(sName As String,ByRef dir As String,ByRef channel As String)
    dir="":channel=""
    Select Case sName
        Case "Приход — Цех":dir="IN":channel="WORKSHOP"
        Case "Расход — Цех":dir="OUT":channel="WORKSHOP"
        Case "Приход — Производство":dir="IN":channel="PRODUCTION"
        Case "Приход — Офис":dir="IN":channel="OFFICE"
        Case "Приход — Детали":dir="IN":channel="PARTS"
        Case "Расход — Производство":dir="OUT":channel="PRODUCTION"
        Case "Расход — Офис":dir="OUT":channel="OFFICE"
        Case "Расход — Детали":dir="OUT":channel="PARTS"
    End Select
End Sub

Function WMSWF_IsWorkflowSheet(sName As String) As Boolean
    Dim d As String,c As String:WMSWF_Context sName,d,c:WMSWF_IsWorkflowSheet=(d<>"")
End Function

Function WMSWF_RowHasData(oSh As Object,r As Long) As Boolean
    WMSWF_RowHasData=(Trim(oSh.getCellByPosition(1,r).String)<>"" Or Trim(oSh.getCellByPosition(2,r).String)<>"" Or oSh.getCellByPosition(3,r).Value<>0)
End Function

Function WMSWF_LastRow(oSh As Object) As Long
    Dim c As Object:c=oSh.createCursor():c.gotoEndOfUsedArea(True):WMSWF_LastRow=c.RangeAddress.EndRow
End Function

Function WMSWF_NewID(prefix As String) As String
    Randomize
    WMSWF_NewID=prefix & "-" & Format(Now,"YYYYMMDDHHMMSS") & "-" & CStr(Int(Rnd()*900000)+100000)
End Function

Function WMSWF_Num(v As Double) As String
    WMSWF_Num=Replace(CStr(v),",",".")
End Function

Function WMSWF_SQLDate(d As Double) As String
    If d<=0 Then d=Date
    WMSWF_SQLDate=WMSDB_SQLText(Format(CDate(d),"YYYY-MM-DD"))
End Function

Sub WMSWF_CreateAct()
    WMSACTWF_CreateFromActiveSheet
End Sub

Sub WMSWF_OpenStock()
    If ThisComponent.Sheets.hasByName("Остаток") Then ThisComponent.CurrentController.setActiveSheet(ThisComponent.Sheets.getByName("Остаток")):WMSDBST_RefreshStock
End Sub

Sub WMSWF_OpenSearch()
    If ThisComponent.Sheets.hasByName("База - Поиск") Then ThisComponent.CurrentController.setActiveSheet(ThisComponent.Sheets.getByName("База - Поиск"))
End Sub


' ==========================================================================
' 3.3 workflow speed helpers. No OnChange listeners: only explicit buttons.
' ==========================================================================
Sub WMSWF_AutofillSelectedByArticle()
    Dim oDoc As Object,oSh As Object,sel As Object,r As Long
    Dim art As String,sErr As String,oCon As Object,ps As Object,rs As Object
    Dim code As String,nm As String,unitName As String,loc As String,cat As String,subcat As String,n As Long
    oDoc=ThisComponent
    oSh=oDoc.CurrentController.ActiveSheet
    If Not WMSWF_IsWorkflowSheet(oSh.Name) Then Exit Sub
    On Error GoTo EH
    sel=oDoc.CurrentSelection
    r=sel.CellAddress.Row
    If r<5 Then MsgBox "Выберите рабочую строку.",48,"WMS":Exit Sub
    art=Trim(oSh.getCellByPosition(2,r).String)
    If art="" Then MsgBox "Сначала введите артикул.",48,"WMS":Exit Sub
    If Not WMSDBX_TryEnter("AUTOFILL") Then MsgBox "Другая операция WMS ещё выполняется.",48,"WMS":Exit Sub
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then GoTo Fail
    ps=oCon.prepareStatement("SELECT p.PRODUCT_CODE,p.PRODUCT_NAME,p.UNIT_NAME,p.DEFAULT_LOCATION FROM WMS_PRODUCT_ALIASES a JOIN WMS_PRODUCTS p ON p.PRODUCT_CODE=a.PRODUCT_CODE WHERE a.ACTIVE_FLAG=1 AND UPPER(TRIM(a.ARTICLE_CODE))=UPPER(TRIM(?))")
    ps.setString(1,art)
    rs=ps.executeQuery()
    Do While rs.next()
        n=n+1
        If n=1 Then
            code=rs.getString(1)
            nm=rs.getString(2)
            unitName=rs.getString(3)
            loc=rs.getString(4)
        End If
        If n>1 Then Exit Do
    Loop
    WMSDBX_CloseRS rs
    WMSDBX_CloseStmt ps
    If n=0 Then
        ps=oCon.prepareStatement("SELECT PRODUCT_CODE,PRODUCT_NAME,UNIT_NAME,DEFAULT_LOCATION FROM WMS_PRODUCTS WHERE ACTIVE_FLAG=1 AND UPPER(TRIM(COALESCE(SUPPLIER_ARTICLE,'')))=UPPER(TRIM(?))")
        ps.setString(1,art)
        rs=ps.executeQuery()
        Do While rs.next()
            n=n+1
            If n=1 Then
                code=rs.getString(1)
                nm=rs.getString(2)
                unitName=rs.getString(3)
                loc=rs.getString(4)
            End If
            If n>1 Then Exit Do
        Loop
    End If
    WMSDBX_CloseRS rs
    WMSDBX_CloseStmt ps
    If n=0 Then GoTo NotFound
    If n>1 Then GoTo Ambig
    ps=oCon.prepareStatement("SELECT FIRST 1 COALESCE(DEST_CATEGORY,''),COALESCE(DEST_SUBCATEGORY,'') FROM WMS_STOCK_LOTS WHERE PRODUCT_CODE=? ORDER BY RECEIPT_DATE DESC,LOT_ID DESC")
    ps.setString(1,code)
    rs=ps.executeQuery()
    If rs.next() Then
        cat=rs.getString(1)
        subcat=rs.getString(2)
    End If
    WMSDBX_CloseRS rs
    WMSDBX_CloseStmt ps
    WMSDB_Close
    If Trim(oSh.getCellByPosition(1,r).String)="" Then oSh.getCellByPosition(1,r).String=nm
    If Trim(oSh.getCellByPosition(4,r).String)="" Then oSh.getCellByPosition(4,r).String=unitName
    WMSWF_AutofillLocationCategory oSh,r,loc,cat,subcat
    WMSDBX_Leave
    Exit Sub
NotFound:
    WMSDB_Close
    WMSDBX_Leave
    MsgBox "Артикул пока не найден. Для нового товара заполните строку вручную и оформите приход.",48,"WMS — Артикул"
    Exit Sub
Ambig:
    WMSDB_Close
    WMSDBX_Leave
    MsgBox "Артикул найден у нескольких товаров. Автозаполнение остановлено.",48,"WMS — Артикул"
    Exit Sub
Fail:
    WMSDBX_CloseRS rs
    WMSDBX_CloseStmt ps
    On Error Resume Next
    WMSDB_Close
    On Error GoTo 0
    WMSDBX_Leave
    MsgBox sErr,16,"WMS — Автозаполнение"
    Exit Sub
EH:
    sErr="Автозаполнение: " & CStr(Err) & " " & Error$
    Resume Fail
End Sub

Sub WMSWF_AutofillLocationCategory(oSh As Object,r As Long,loc As String,cat As String,subcat As String)
    If Left(oSh.Name,6)="Приход" Then
        If oSh.Name="Приход — Офис" Then
            If Trim(oSh.getCellByPosition(7,r).String)="" Then oSh.getCellByPosition(7,r).String=loc
            If Trim(oSh.getCellByPosition(8,r).String)="" Then oSh.getCellByPosition(8,r).String=cat
            If Trim(oSh.getCellByPosition(9,r).String)="" Then oSh.getCellByPosition(9,r).String=subcat
        ElseIf oSh.Name="Приход — Детали" Then
            If Trim(oSh.getCellByPosition(6,r).String)="" Then oSh.getCellByPosition(6,r).String=loc
            If Trim(oSh.getCellByPosition(7,r).String)="" Then oSh.getCellByPosition(7,r).String=cat
        Else
            If Trim(oSh.getCellByPosition(6,r).String)="" Then oSh.getCellByPosition(6,r).String=loc
            If Trim(oSh.getCellByPosition(7,r).String)="" Then oSh.getCellByPosition(7,r).String=cat
            If Trim(oSh.getCellByPosition(8,r).String)="" Then oSh.getCellByPosition(8,r).String=subcat
        End If
    Else
        If Trim(oSh.getCellByPosition(7,r).String)="" Then oSh.getCellByPosition(7,r).String=loc
    End If
End Sub

Sub WMSWF_RepeatSelectedFields()
    Dim oDoc As Object,oSh As Object,sel As Object,r As Long,nr As Long,lastCol As Long,c As Long
    oDoc=ThisComponent
    oSh=oDoc.CurrentController.ActiveSheet
    If Not WMSWF_IsWorkflowSheet(oSh.Name) Then Exit Sub
    On Error GoTo EH
    sel=oDoc.CurrentSelection
    r=sel.CellAddress.Row
    If r<5 Then MsgBox "Выберите рабочую строку.",48,"WMS":Exit Sub
    nr=r+1
    Do While WMSWF_RowHasData(oSh,nr)
        nr=nr+1
        If nr>5000 Then MsgBox "Не найдена свободная строка.",48,"WMS":Exit Sub
    Loop
    lastCol=UBound(WMSWF_Headers(oSh.Name))
    For c=0 To lastCol
        oSh.getCellByPosition(c,nr).String=oSh.getCellByPosition(c,r).String
        If oSh.getCellByPosition(c,r).Type=1 Then oSh.getCellByPosition(c,nr).Value=oSh.getCellByPosition(c,r).Value
    Next c
    ' Product-specific values are intentionally blank; recurring logistics fields stay copied.
    oSh.getCellByPosition(1,nr).String=""
    oSh.getCellByPosition(2,nr).String=""
    oSh.getCellByPosition(3,nr).String=""
    oSh.getCellByPosition(3,nr).Value=0
    oDoc.CurrentController.select(oSh.getCellByPosition(2,nr))
    Exit Sub
EH:
    MsgBox "Повтор реквизитов: " & CStr(Err) & " " & Error$,16,"WMS"
End Sub

Sub WMSWF_SyntaxProbe()
End Sub

Function WMSWF_VerifyPersistedRowCon(oCon As Object,oSh As Object,r As Long,ByRef sErr As String) As Boolean
    Dim dir As String,channel As String,opID As String,nm As String,art As String
    Dim unitName As String,person As String,loc As String,cat As String,subcat As String
    Dim docNo As String,note As String,dest As String,lotID As String,ret As String
    Dim dt As Double,qty As Double,expectedQty As Double,totalQty As Double,n As Long,ps As Object,rs As Object
    WMSWF_VerifyPersistedRowCon=False:sErr=""
    On Error GoTo EH
    WMSWF_Context oSh.Name,dir,channel
    If dir="IN" Then
        WMSWF_ReadReceipt oSh,r,channel,dt,nm,art,qty,unitName,person,loc,cat,subcat,docNo,note
        expectedQty=qty
    ElseIf dir="OUT" Then
        WMSWF_ReadIssue oSh,r,channel,dt,nm,art,qty,unitName,person,dest,loc,lotID,ret,note
        expectedQty=-qty
    Else
        sErr="Неизвестный лист проверки."
        GoTo Done
    End If
    opID=Trim(oSh.getCellByPosition(50,r).String)
    If opID="" Then
        sErr="Отсутствует номер операции."
        GoTo Done
    End If
    ps=oCon.prepareStatement("SELECT PRODUCT_NAME,ENTRY_QTY,ENTRY_UNIT,LOCATION_NAME,FLOW_CHANNEL,MOVEMENT_TYPE,STATUS_NAME FROM WMS_STOCK_MOVEMENTS WHERE OPERATION_ID=?")
    ps.setString(1,opID)
    rs=ps.executeQuery()
    totalQty=0:n=0
    Do While rs.next()
        n=n+1
        If rs.getString(1)<>nm Or rs.getString(3)<>unitName Or rs.getString(4)<>loc Or rs.getString(5)<>channel Or rs.getString(6)<>dir Or rs.getString(7)<>"POSTED" Then
            sErr="После повторного открытия ODB реквизиты движения не совпали."
            GoTo Done
        End If
        totalQty=totalQty+rs.getDouble(2)
    Loop
    If n=0 Or (dir="IN" And n<>1) Or Abs(totalQty-expectedQty)>0.0001 Then
        sErr="После повторного открытия ODB отсутствуют движения или не совпадает общее количество."
        GoTo Done
    End If
    WMSWF_VerifyPersistedRowCon=True
Done:
    WMSDBX_CloseRS rs
    WMSDBX_CloseStmt ps
    If sErr<>"" Then sErr="Строка " & CStr(r+1) & " (" & opID & "): " & sErr & " Ввод сохранён на листе. Не проводите его повторно без сверки."
    Exit Function
EH:
    sErr="Проверка сохранённой строки: " & CStr(Err) & " " & Error$
    Resume Done
End Function

Function WMSWF_PlanIssueCon(oCon As Object,code As String,loc As String,requestedLot As String,qty As Double,unitName As String,ByRef sErr As String) As Variant
    Dim st As Object,rs As Object,sql As String,parts() As Variant,n As Long
    Dim remaining As Double,avail As Double,factor As Double,takeQty As Double,baseQty As Double
    Dim lid As String,bu As String,origin As String,cat As String,subcat As String
    sErr="":remaining=qty:n=0
    On Error GoTo EH
    sql="SELECT LOT_ID,CAST(QTY AS DOUBLE PRECISION),UNIT_NAME,ORIGIN_NAME,CATEGORY_NAME,SUBCATEGORY_NAME FROM WMS_V_LOT_BALANCE WHERE PRODUCT_CODE=" & WMSDB_SQLText(code) & " AND LOCATION_NAME=" & WMSDB_SQLText(loc) & " AND QTY>0"
    If Trim(requestedLot)<>"" Then sql=sql & " AND LOT_ID=" & WMSDB_SQLText(requestedLot)
    sql=sql & " ORDER BY RECEIPT_DATE,LOT_ID"
    st=oCon.createStatement():rs=st.executeQuery(sql)
    Do While rs.next()
        lid=rs.getString(1):avail=rs.getDouble(2):bu=rs.getString(3)
        origin=rs.getString(4):cat=rs.getString(5):subcat=rs.getString(6)
        factor=1
        If UCase(Trim(unitName))<>UCase(Trim(bu)) Then
            factor=WMSWF_LotFactorCon(oCon,lid,unitName,sErr)
            If sErr<>"" Then GoTo Done
            If factor<=0 Then
                sErr="Для партии " & lid & " нет пересчёта из " & unitName & " в " & bu & "."
                GoTo Done
            End If
        End If
        takeQty=remaining
        If avail/factor<takeQty Then takeQty=avail/factor
        baseQty=takeQty*factor
        ReDim Preserve parts(0 To n)
        parts(n)=Array(lid,takeQty,baseQty,bu,origin,cat,subcat)
        n=n+1:remaining=remaining-takeQty
        If remaining<=0.0000001 Then Exit Do
    Loop
    If remaining>0.0000001 Or n=0 Then
        sErr="Недостаточно остатка. Требуется " & WMSWF_Num(qty) & " " & unitName & "; доступно " & WMSWF_Num(qty-remaining) & "."
        GoTo Done
    End If
    WMSWF_PlanIssueCon=parts()
Done:
    WMSDBX_CloseRS rs:WMSDBX_CloseStmt st
    Exit Function
EH:
    sErr="Автоматический подбор партий: " & CStr(Err) & " " & Error$
    Resume Done
End Function

Function WMSWF_ResolveRowProductCon(con As Object,sh As Object,r As Long,art As String,nm As String,unitName As String,loc As String,allowCreate As Boolean,ByRef sErr As String) As String
    Dim code As String,st As Object,rs As Object,sql As String
    sErr="":WMSWF_ResolveRowProductCon=""
    On Error GoTo EH
    If Right(sh.Name,3)<>"Цех" Then
        WMSWF_ResolveRowProductCon=WMSWF_ResolveProductCon(con,art,nm,unitName,loc,allowCreate,sErr)
        Exit Function
    End If
    code=Trim(sh.getCellByPosition(11,r).String)
    If code="" And art<>"" Then
        code=WMSWF_ResolveProductCon(con,art,nm,unitName,loc,allowCreate,sErr)
        If sErr<>"" Then GoTo Done
    ElseIf code="" Then
        If Not allowCreate Then
            sErr="Для выдачи без артикула укажите внутренний код товара в колонке L. Он виден в остатке и поиске."
            GoTo Done
        End If
        code=WMSARCH_NextProductCodeCon(con,sErr)
        If sErr<>"" Then GoTo Done
        st=con.createStatement()
        st.executeUpdate("INSERT INTO WMS_PRODUCTS (PRODUCT_CODE,PRODUCT_NAME,UNIT_NAME,DEFAULT_LOCATION,SUPPLIER_ARTICLE,ACTIVE_FLAG,ITEM_ID) VALUES (" & WMSDB_SQLText(code) & "," & WMSDB_SQLText(nm) & "," & WMSDB_SQLText(unitName) & "," & WMSDB_SQLText(loc) & ",'',1," & WMSDB_SQLText("ITEM-" & code) & ")")
    End If
    st=con.createStatement()
    rs=st.executeQuery("SELECT PRODUCT_NAME,UNIT_NAME FROM WMS_PRODUCTS WHERE ACTIVE_FLAG=1 AND PRODUCT_CODE=" & WMSDB_SQLText(code))
    If Not rs.next() Then sErr="Внутренний код не найден: " & code:GoTo Done
    If UCase(Trim(rs.getString(1)))<>UCase(Trim(nm)) Then sErr="Код или артикул относится к другому наименованию. Проверьте товар.":GoTo Done
    If allowCreate And UCase(Trim(rs.getString(2)))<>UCase(Trim(unitName)) Then sErr="Приход в другой единице требует настроенного пересчёта. Укажите базовую единицу товара.":GoTo Done
    WMSWF_ResolveRowProductCon=code
Done:
    WMSDBX_CloseRS rs:WMSDBX_CloseStmt st
    Exit Function
EH:
    sErr="Определение товара: " & CStr(Err) & " " & Error$
    Resume Done
End Function
