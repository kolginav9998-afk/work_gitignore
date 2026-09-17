Option Explicit

' ============================================================================
' WMS_03_Issues_FINAL.bas
' Production controller for sheet "Выдачи" in the clean WMS architecture.
' LibreOffice Calc / ODS / Linux / LibreOffice Basic + UNO only.
'
' SOURCE RULES PRESERVED:
' - A №, B Код, C Наименование, D Кол-во, E Ед. изм., F Кто получил,
'   G Дата, H Откуда, J Возвращено, K Дата возврата, L Примечание are
'   user-entered fields.
' - F gets a non-blocking employee suggestion drop-down built from real history.
' - I Возвратный is a strict Да/Нет drop-down.
' - Code/name matching never auto-merges fuzzy items. Fuzzy is suggestion only.
' - Returns are cumulative: 5 issued -> 3 returned -> 4 means return delta +1.
' - Returned <= issued; return date >= issue date; posted return history is never
'   silently reduced or overwritten.
' - Sorting/filtering must keep hidden SourceID/return metadata with the row.
' - OnChange is light: normalize/validate/SourceID/DIRTY/queue/suggestions only.
' - EMPTY ROW GUARD: column A number alone is not a business row; B:L activates it.
' - BUTTONS 1.0.4: working actions are available as buttons directly on the sheet.
' - Movement/balance posting remains the responsibility of the common movement
'   engine; this controller exposes a stable integration API.
' ============================================================================

Global Const WMSISS_VERSION = "1.0.13-BULK-CONDUCT-UI"
Global Const WMSISS_SCHEMA = "1.0"
Global Const WMSISS_TARGET_LO = "25.8.7.3"
Global Const WMSISS_SHEET = "Выдачи"
Global Const WMSISS_QUEUE = "SYS_WMS_QUEUE"
Global Const WMSISS_SETTINGS = "Настройки WMS"
Global Const WMSISS_AUDIT = "Журнал WMS"
Global Const WMSISS_ERRORS = "Ошибки WMS"
Global Const WMSISS_NOM = "Номенклатура"
Global Const WMSISS_HEADER_ROW = 0
Global Const WMSISS_USER_LAST_COL = 11
Global Const WMSISS_MIN_STYLE_ROWS = 500
Global Const WMSISS_FILTER_BUFFER = 100
Global Const WMSISS_BULK_THRESHOLD = 50
Global Const WMSISS_MAX_EVENT_ROWS = 5000
Global Const WMSISS_EPS = 0.0000001

Global Const WMSISS_YES = "Да"
Global Const WMSISS_NO = "Нет"
Global Const WMSISS_STATE_DRAFT = "Черновик"
Global Const WMSISS_STATE_DONE = "Проведено"
Global Const WMSISS_STATE_CANCEL = "Отменено"

Global Const WMSISS_T_ID = "_WMS_SourceID"
Global Const WMSISS_T_STATE = "_WMS_State"
Global Const WMSISS_T_HASH = "_WMS_SyncHash"
Global Const WMSISS_T_ERROR = "_WMS_Error"
Global Const WMSISS_T_RETPOSTED = "_WMS_ReturnPostedQty"
Global Const WMSISS_T_RETSEQ = "_WMS_ReturnSeq"
Global Const WMSISS_T_RETHASH = "_WMS_ReturnHash"
Global Const WMSISS_T_RETERR = "_WMS_ReturnError"
Global Const WMSISS_T_DIRTY = "_WMS_Dirty"
Global Const WMSISS_T_MODE = "_WMS_Mode"
Global Const WMSISS_T_LEGACY = "_WMS_LegacyKey"
Global Const WMSISS_T_TOUCH = "_WMS_LastTouch"

Global Const WMSISS_X_ROWVER = "_WMS_IssueRowVersion"
Global Const WMSISS_X_BASEFP = "_WMS_BaseFingerprint"
Global Const WMSISS_X_RETFP = "_WMS_ReturnFingerprint"
Global Const WMSISS_X_FLAGS = "_WMS_ValidationFlags"
Global Const WMSISS_X_SEVERITY = "_WMS_Severity"
Global Const WMSISS_X_VALIDATED = "_WMS_LastValidated"
Global Const WMSISS_X_CODESUG = "_WMS_CodeSuggestion"
Global Const WMSISS_X_SUGNAME = "_WMS_CodeSuggestionName"
Global Const WMSISS_X_SUGSCORE = "_WMS_CodeSuggestionScore"
Global Const WMSISS_X_SUGSOURCE = "_WMS_CodeSuggestionSource"
Global Const WMSISS_X_EMPKEY = "_WMS_EmployeeKey"
Global Const WMSISS_X_RETSTATE = "_WMS_ReturnState"
Global Const WMSISS_X_LASTQMODE = "_WMS_LastQueueMode"
Global Const WMSISS_X_RETHINT = "_WMS_ReturnableHint"
Global Const WMSISS_X_UNITHINT = "_WMS_UnitHint"
Global Const WMSISS_X_PLACEHINT = "_WMS_PlaceHint"

Global Const WMSISS_LIST_EMP = "WMS_ISSUE_EMPLOYEES"
Global Const WMSISS_LIST_UNIT = "WMS_ISSUE_UNITS"
Global Const WMSISS_LIST_PLACE = "WMS_ISSUE_PLACES"
Global Const WMSISS_HELP_EMP_COL = 7
Global Const WMSISS_HELP_UNIT_COL = 8
Global Const WMSISS_HELP_PLACE_COL = 9
Global Const WMSISS_HELP_MAX = 300

Global gWMSISS_Busy As Boolean
Global gWMSISS_EventDepth As Long
Global gWMSISS_LastError As String
Global gWMSISS_LastReport As String
Global gWMSISS_LastEventMs As Double
Global gWMSISS_HeaderCache As Variant
Global gWMSISS_HeaderReady As Boolean
Global gWMSISS_LastHeaderColCached As Long
Global gWMSISS_QueueReady As Boolean
Global gWMSISS_QueueNextRow As Long
Global gWMSISS_QKeys() As String
Global gWMSISS_QRows() As Long
Global gWMSISS_QCap As Long
Global gWMSISS_QCount As Long
Global gWMSISS_NomReady As Boolean
Global gWMSISS_NomCount As Long
Global gWMSISS_NomCode() As String
Global gWMSISS_NomName() As String
Global gWMSISS_NomUnit() As String
Global gWMSISS_NomPlace() As String
Global gWMSISS_NomReturnable() As String
Global gWMSISS_NomActive() As Boolean
Global gWMSISS_BulkMode As Boolean


' ============================================================================
' FAST UPGRADE 1.0.6
' Use this when the previous 1.0.4 controller is already installed.
' It does NOT restyle 500 rows, rebuild all validations, scan/revalidate all
' issues, touch filters, or run the full self-check. It only makes the schema
' safe and replaces the button panel, including "Провести выдачу".
' ============================================================================
Sub InstallIssuesFast()
    Dim oDoc As Object,oSh As Object,sErr As String
    oDoc=ThisComponent
    On Error GoTo EH

    If Not Issues_Preflight(oDoc,sErr) Then
        MsgBox sErr,16,"WMS — Выдачи: быстрое обновление"
        Exit Sub
    End If

    oSh=oDoc.Sheets.getByName(WMSISS_SHEET)
    gWMSISS_Busy=True

    If Not Issues_EnsureExtensionColumns(oSh,sErr) Then GoTo Fail
    Issues_ResetHeaderCache

    If Not Issues_InstallButtonsCore(oDoc,oSh,sErr) Then GoTo Fail

    Issues_SetSetting oDoc,"WMS_IssuesVersion",WMSISS_VERSION,"Версия контроллера Выдачи",True
    Issues_SetSetting oDoc,"WMS_IssuesState","READY","Состояние контроллера Выдачи",True

    gWMSISS_Busy=False
    MsgBox "Быстрое обновление Выдачи завершено." & Chr(10) & _
           "Кнопка 'Провести выдачу' установлена." & Chr(10) & _
           "Версия: " & WMSISS_VERSION,64,"WMS — Выдачи"
    Exit Sub

Fail:
    gWMSISS_Busy=False
    MsgBox "Быстрое обновление остановлено: " & sErr,16,"WMS — Выдачи"
    Exit Sub

EH:
    gWMSISS_Busy=False
    MsgBox "Ошибка быстрого обновления: " & CStr(Err) & " " & Error$,16,"WMS — Выдачи"
End Sub

Sub InstallIssuesConductButtonOnly()
    Dim oDoc As Object,oSh As Object,sErr As String
    oDoc=ThisComponent
    On Error GoTo EH
    If Not oDoc.Sheets.hasByName(WMSISS_SHEET) Then
        MsgBox "Лист 'Выдачи' не найден.",16,"WMS — Выдачи"
        Exit Sub
    End If
    oSh=oDoc.Sheets.getByName(WMSISS_SHEET)
    gWMSISS_Busy=True
    If Not Issues_InstallButtonsCore(oDoc,oSh,sErr) Then
        gWMSISS_Busy=False
        MsgBox "Кнопки не установлены: " & sErr,16,"WMS — Выдачи"
        Exit Sub
    End If
    gWMSISS_Busy=False
    MsgBox "Панель кнопок обновлена. 'Провести выдачу' добавлена.",64,"WMS — Выдачи"
    Exit Sub
EH:
    gWMSISS_Busy=False
    MsgBox "Ошибка панели кнопок: " & CStr(Err) & " " & Error$,16,"WMS — Выдачи"
End Sub

' ============================================================================
' PUBLIC INSTALL / UPGRADE / OPERATIONS / QA
' ============================================================================

Sub InstallIssues()
    Issues_InstallStableInput
End Sub

Sub Issues_InstallStableInput()
    Dim oDoc As Object,oSh As Object,sErr As String
    oDoc=ThisComponent

    On Error GoTo EH

    MsgBox "Установка стабильного режима Выдачи началась.",64,"WMS — Выдачи"

    If Not oDoc.Sheets.hasByName(WMSISS_SHEET) Then
        MsgBox "Лист 'Выдачи' не найден.",16,"WMS — Выдачи"
        Exit Sub
    End If

    oSh=oDoc.Sheets.getByName(WMSISS_SHEET)

    gWMSISS_Busy=True
    gWMSISS_EventDepth=0

    ' 1.0.11: оставляем только лёгкий OnChange для колонки B (Код).
    ' Никаких полных проверок/перерисовок при вводе.
    If Not Issues_InstallEvent(oSh,sErr) Then
        gWMSISS_Busy=False
        MsgBox "Не удалось назначить лёгкое автозаполнение по коду:" & Chr(10) & sErr,48,"WMS — Выдачи"
        Exit Sub
    End If

    ' Убираем живые validation-ограничения из рабочей области.
    ' Они могли откатывать введённое значение при переходе в другую ячейку.
    Issues_RemoveLiveValidations oSh

    ' Возвращаем нормальный контроллер после возможного оборванного макроса.
    On Error Resume Next
    oDoc.unlockControllers
    Err=0
    On Error GoTo EH

    oDoc.CurrentController.setActiveSheet(oSh)

    ' Кнопки оставляем, но тяжёлую переустановку листа не запускаем.
    If Not Issues_InstallButtonsCore(oDoc,oSh,sErr) Then
        gWMSISS_Busy=False
        MsgBox "Режим ввода восстановлен, но панель кнопок не обновилась:" & Chr(10) & sErr,48,"WMS — Выдачи"
        Exit Sub
    End If

    Issues_SetSetting oDoc,"WMS_IssuesVersion",WMSISS_VERSION,"Версия контроллера Выдачи",True
    Issues_SetSetting oDoc,"WMS_IssuesState","STABLE_INPUT","Состояние контроллера Выдачи",True

    gWMSISS_Busy=False

    MsgBox "Готово." & Chr(10) & _
           "Версия: " & WMSISS_VERSION & Chr(10) & _
           "Автозаполнение включено только для колонки B «Код»." & Chr(10) & _
           "По коду из Firebird заполняются Наименование, Ед. изм. и, когда место однозначно, Откуда." & Chr(10) & _
           "Дата выдачи подставляется автоматически, если пустая." & Chr(10) & _
           "Тяжёлые проверки при вводе по-прежнему отключены.", _
           64,"WMS — Выдачи"
    Exit Sub

EH:
    gWMSISS_Busy=False
    MsgBox "Установка стабильного режима остановлена:" & Chr(10) & _
           CStr(Err) & " " & Error$,16,"WMS — Выдачи"
End Sub

Sub Issues_RemoveOnChangeEvent(oSh As Object)
    On Error Resume Next
    oSh.Events.replaceByName "OnChange", Array()
    Err=0
    On Error GoTo 0
End Sub

Sub Issues_RemoveLiveValidations(oSh As Object)
    Dim oRange As Object,oV As Object,endRow As Long
    On Error GoTo Done

    endRow=WMSISS_MIN_STYLE_ROWS
    If endRow<1000 Then endRow=1000
    If endRow>oSh.Rows.getCount()-1 Then endRow=oSh.Rows.getCount()-1

    oRange=oSh.getCellRangeByPosition(0,1,WMSISS_USER_LAST_COL,endRow)
    oV=oRange.Validation
    oV.Type=com.sun.star.sheet.ValidationType.ANY
    oV.ShowErrorMessage=False
    oV.ShowInputMessage=False
    oRange.Validation=oV
Done:
End Sub

Sub Issues_InputDiagnostic()
    Dim oDoc As Object,oSh As Object,ev As Variant,s As String,v As Object
    oDoc=ThisComponent
    On Error GoTo EH

    If Not oDoc.Sheets.hasByName(WMSISS_SHEET) Then
        MsgBox "Лист 'Выдачи' не найден.",16,"WMS — Диагностика ввода"
        Exit Sub
    End If
    oSh=oDoc.Sheets.getByName(WMSISS_SHEET)

    s="Версия: " & WMSISS_VERSION & Chr(10)

    On Error Resume Next
    ev=oSh.Events.getByName("OnChange")
    If Err<>0 Then
        Err=0
        s=s & "OnChange: отсутствует/отключён" & Chr(10)
    Else
        s=s & "OnChange: проверен" & Chr(10)
    End If

    v=oSh.getCellByPosition(1,1).Validation
    s=s & "Live validation B2: " & CStr(v.Type) & " (0 = ANY)" & Chr(10)
    On Error GoTo EH

    s=s & "Режим: стабильный ручной ввод"
    MsgBox s,64,"WMS — Диагностика ввода"
    Exit Sub
EH:
    MsgBox "Диагностика ввода: " & CStr(Err) & " " & Error$,16,"WMS — Диагностика ввода"
End Sub

Sub InstallIssuesQuiet()
    Dim sReport As String
    Call Issues_InstallCore(ThisComponent, False, sReport)
End Sub

Sub UpgradeIssues()
    InstallIssues
End Sub

Sub Issues_RefreshDesign()
    Dim oDoc As Object,oSh As Object,s As String
    oDoc=ThisComponent
    If Not Issues_Preflight(oDoc,s) Then MsgBox s,16,"WMS — Выдачи":Exit Sub
    oSh=oDoc.Sheets.getByName(WMSISS_SHEET)
    gWMSISS_Busy=True
    Issues_ApplyDesign oDoc,oSh
    Issues_RebuildHelperLists oDoc,oSh
    Issues_ApplyValidations oDoc,oSh
    Issues_RefreshFilter oDoc,oSh
    Issues_ApplyFreeze oDoc,oSh
    gWMSISS_Busy=False
    MsgBox "Дизайн, сетка, форматы, подсказочные списки, валидации и фильтр листа 'Выдачи' обновлены.",64,"WMS — Выдачи"
End Sub

Sub Issues_RevalidateAll()
    Dim s As String
    If Issues_RevalidateAllCore(ThisComponent,True,s) Then
        MsgBox s,64,"WMS — Выдачи"
    Else
        MsgBox s,48,"WMS — Выдачи"
    End If
End Sub

Sub Issues_EmptyRowGuardCleanup()
    Dim oDoc As Object,oSh As Object,last As Long,r As Long,nClean As Long,nProtected As Long,sID As String
    Dim cID As Long,cState As Long,cPosted As Long,state As String,posted As Double
    oDoc=ThisComponent
    If Not oDoc.Sheets.hasByName(WMSISS_SHEET) Then MsgBox "Лист 'Выдачи' не найден.",16,"WMS — Выдачи":Exit Sub
    oSh=oDoc.Sheets.getByName(WMSISS_SHEET)
    cID=Issues_FindHeader(oSh,WMSISS_T_ID):cState=Issues_FindHeader(oSh,WMSISS_T_STATE):cPosted=Issues_FindHeader(oSh,WMSISS_T_RETPOSTED)
    last=Issues_LastContentRow(oSh,Issues_LastHeaderCol(oSh))
    If last<1 Then MsgBox "Проверять нечего.",64,"WMS — Выдачи":Exit Sub
    gWMSISS_Busy=True
    On Error GoTo EH
    For r=1 To last
        If Not Issues_RowHasBusinessData(oSh,r) Then
            state="":posted=0:sID=""
            If cState>=0 Then state=Trim(oSh.getCellByPosition(cState,r).String)
            If cPosted>=0 Then posted=Issues_CellNumber(oSh.getCellByPosition(cPosted,r))
            If cID>=0 Then sID=Trim(oSh.getCellByPosition(cID,r).String)
            If Issues_Canon(state)=Issues_Canon(WMSISS_STATE_DONE) Or posted>WMSISS_EPS Then
                nProtected=nProtected+1
            Else
                If sID<>"" Then Issues_CancelPendingQueueBySourceID oDoc,sID,"EMPTY_ROW_GUARD"
                Issues_HandleBlankBusinessRow oSh,r
                nClean=nClean+1
            End If
        End If
    Next r
    Issues_ResetQueueCache
    gWMSISS_Busy=False
    MsgBox "Empty Row Guard завершён." & Chr(10) & _
           "Очищено пустых черновых строк: " & CStr(nClean) & Chr(10) & _
           "Защищено проведённых строк: " & CStr(nProtected) & Chr(10) & _
           "Номер в колонке A сам по себе больше не создаёт выдачу.",64,"WMS — Выдачи"
    Exit Sub
EH:
    gWMSISS_Busy=False
    MsgBox "Очистка остановлена: " & CStr(Err) & " " & Error$,16,"WMS — Выдачи"
End Sub

Sub Issues_EmptyRowGuardCleanupCore(oDoc As Object,oSh As Object)
    Dim last As Long,r As Long,sID As String,cID As Long,cState As Long,cPosted As Long,state As String,posted As Double
    cID=Issues_FindHeader(oSh,WMSISS_T_ID):cState=Issues_FindHeader(oSh,WMSISS_T_STATE):cPosted=Issues_FindHeader(oSh,WMSISS_T_RETPOSTED)
    last=Issues_LastContentRow(oSh,Issues_LastHeaderCol(oSh))
    For r=1 To last
        If Not Issues_RowHasBusinessData(oSh,r) Then
            state="":posted=0:sID=""
            If cState>=0 Then state=Trim(oSh.getCellByPosition(cState,r).String)
            If cPosted>=0 Then posted=Issues_CellNumber(oSh.getCellByPosition(cPosted,r))
            If Issues_Canon(state)<>Issues_Canon(WMSISS_STATE_DONE) And posted<=WMSISS_EPS Then
                If cID>=0 Then sID=Trim(oSh.getCellByPosition(cID,r).String)
                If sID<>"" Then Issues_CancelPendingQueueBySourceID oDoc,sID,"EMPTY_ROW_GUARD"
                Issues_HandleBlankBusinessRow oSh,r
            End If
        End If
    Next r
    Issues_ResetQueueCache
End Sub

Sub Issues_RebuildLists()
    Dim oDoc As Object,oSh As Object,s As String
    oDoc=ThisComponent
    If Not Issues_Preflight(oDoc,s) Then MsgBox s,16,"WMS — Выдачи":Exit Sub
    oSh=oDoc.Sheets.getByName(WMSISS_SHEET)
    gWMSISS_Busy=True
    Issues_RebuildHelperLists oDoc,oSh
    Issues_ApplyValidations oDoc,oSh
    gWMSISS_Busy=False
    MsgBox "Списки сотрудников, единиц и мест обновлены из фактических данных WMS.",64,"WMS — Выдачи"
End Sub

Sub Issues_RefreshSelectedSuggestion()
    Dim oDoc As Object,oSh As Object,oSel As Object,r As Long,s As String
    oDoc=ThisComponent
    If Not Issues_Preflight(oDoc,s) Then MsgBox s,16,"WMS — Выдачи":Exit Sub
    oSh=oDoc.CurrentController.ActiveSheet
    If oSh.Name<>WMSISS_SHEET Then MsgBox "Откройте лист 'Выдачи'.",48,"WMS — Выдачи":Exit Sub
    oSel=oDoc.CurrentSelection
    On Error GoTo Bad
    r=oSel.CellAddress.Row
    If r<1 Then GoTo Bad
    gWMSISS_Busy=True
    Issues_ResetNomCache
    Issues_ValidateRow oDoc,oSh,r,True,True
    Issues_ApplyRowVisual oSh,r
    gWMSISS_Busy=False
    MsgBox Issues_RowDiagnosticText(oSh,r),64,"WMS — Выдачи: строка " & CStr(r+1)
    Exit Sub
Bad:
    gWMSISS_Busy=False
    MsgBox "Выделите одну рабочую ячейку в строке выдачи.",48,"WMS — Выдачи"
End Sub

Sub Issues_ShowSelectedDiagnostics()
    Dim oDoc As Object,oSh As Object,oSel As Object,r As Long
    oDoc=ThisComponent:oSh=oDoc.CurrentController.ActiveSheet
    If oSh.Name<>WMSISS_SHEET Then MsgBox "Откройте лист 'Выдачи'.",48,"WMS — Выдачи":Exit Sub
    oSel=oDoc.CurrentSelection
    On Error GoTo Bad
    r=oSel.CellAddress.Row
    If r<1 Then GoTo Bad
    MsgBox Issues_RowDiagnosticText(oSh,r),64,"WMS — Выдачи: диагностика"
    Exit Sub
Bad:
    MsgBox "Выделите одну ячейку рабочей строки.",48,"WMS — Выдачи"
End Sub

Sub Issues_ButtonFillAllByCodes()
    Dim oDoc As Object,oSh As Object,oCon As Object,sErr As String
    Dim last As Long,r As Long,nCodes As Long,nFilled As Long,nMiss As Long
    Dim code As String,oldName As String

    oDoc=ThisComponent
    If Not oDoc.Sheets.hasByName(WMSISS_SHEET) Then Exit Sub
    oSh=oDoc.Sheets.getByName(WMSISS_SHEET)
    oDoc.CurrentController.setActiveSheet(oSh)

    sErr=""
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then MsgBox sErr,16,"WMS — Выдачи":Exit Sub

    last=Issues_LastContentRow(oSh,WMSISS_USER_LAST_COL)
    If last<1 Then
        MsgBox "Нет строк для массового заполнения.",48,"WMS — Выдачи"
        Exit Sub
    End If

    gWMSISS_Busy=True
    On Error GoTo EH

    For r=1 To last
        code=Trim(oSh.getCellByPosition(1,r).String)
        If code<>"" Then
            nCodes=nCodes+1
            oldName=Trim(oSh.getCellByPosition(2,r).String)
            sErr=""
            If Issues_AutofillRowByCodeCon(oCon,oSh,r,sErr) Then
                nFilled=nFilled+1
            Else
                nMiss=nMiss+1
            End If
        End If
    Next r

    gWMSISS_Busy=False
    MsgBox "Массовое заполнение завершено." & Chr(10) & _
           "Кодов обработано: " & CStr(nCodes) & Chr(10) & _
           "Найдено в БД: " & CStr(nFilled) & Chr(10) & _
           "Не найдено: " & CStr(nMiss), _
           IIf(nMiss=0,64,48),"WMS — Выдачи"
    Exit Sub
EH:
    gWMSISS_Busy=False
    MsgBox "Массовое заполнение остановлено: " & CStr(Err) & " " & Error$,16,"WMS — Выдачи"
End Sub

Sub Issues_ButtonLookupCode()
    Dim oDoc As Object,oSh As Object,oSel As Object,r As Long,oCon As Object,sErr As String
    oDoc=ThisComponent
    If Not oDoc.Sheets.hasByName(WMSISS_SHEET) Then Exit Sub
    oSh=oDoc.Sheets.getByName(WMSISS_SHEET)
    If oDoc.CurrentController.ActiveSheet.Name<>WMSISS_SHEET Then
        oDoc.CurrentController.setActiveSheet(oSh)
    End If
    oSel=oDoc.CurrentSelection
    On Error GoTo Bad
    r=oSel.CellAddress.Row
    If r<1 Then GoTo Bad
    sErr=""
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then MsgBox sErr,16,"WMS — Выдачи":Exit Sub
    If Issues_AutofillRowByCodeCon(oCon,oSh,r,sErr) Then
        MsgBox "Данные товара заполнены по коду.",64,"WMS — Выдачи"
    ElseIf sErr<>"" Then
        MsgBox sErr,16,"WMS — Выдачи"
    Else
        MsgBox "Такой код товара не найден в базе.",48,"WMS — Выдачи"
    End If
    Exit Sub
Bad:
    MsgBox "Выделите рабочую строку выдачи.",48,"WMS — Выдачи"
End Sub

Sub Issues_InstallButtons()
    Dim oDoc As Object,oSh As Object,sErr As String
    oDoc=ThisComponent
    If Not Issues_Preflight(oDoc,sErr) Then MsgBox sErr,16,"WMS — Выдачи":Exit Sub
    oSh=oDoc.Sheets.getByName(WMSISS_SHEET)
    If Issues_InstallButtonsCore(oDoc,oSh,sErr) Then
        MsgBox "Кнопки листа 'Выдачи' установлены/обновлены.",64,"WMS — Выдачи"
    Else
        MsgBox "Не удалось установить кнопки: " & sErr,48,"WMS — Выдачи"
    End If
End Sub

Sub Issues_StabilityRepair()
    Issues_InstallStableInput
End Sub

Sub Issues_ButtonNewIssue()
    Dim oDoc As Object,oSh As Object,r As Long,last As Long
    oDoc=ThisComponent
    If Not oDoc.Sheets.hasByName(WMSISS_SHEET) Then MsgBox "Лист 'Выдачи' не найден.",16,"WMS — Выдачи":Exit Sub
    oSh=oDoc.Sheets.getByName(WMSISS_SHEET)
    oDoc.CurrentController.setActiveSheet(oSh)
    last=Issues_LastContentRow(oSh,WMSISS_USER_LAST_COL)
    If last<1 Then last=1
    r=last+1
    If r<1 Then r=1
    Do While r<50000 And Issues_RowHasBusinessData(oSh,r)
        r=r+1
    Loop
    ' Новая выдача начинается с внутреннего кода товара.
    If oSh.getCellByPosition(6,r).Value=0 And Trim(oSh.getCellByPosition(6,r).String)="" Then
        oSh.getCellByPosition(6,r).Value=CDbl(Date)
        On Error Resume Next
        Issues_WriteDate oDoc,oSh.getCellByPosition(6,r),CDbl(Date)
        Err=0
        On Error GoTo 0
    End If
    oDoc.CurrentController.select(oSh.getCellByPosition(1,r))
End Sub

Sub Issues_ButtonReturn()
    On Error GoTo Missing
    WMSDBR_ReturnDialog
    Exit Sub
Missing:
    MsgBox "Модуль WMS_08_DB_Returns не установлен или недоступен.",48,"WMS — Возврат"
End Sub

Sub Issues_ButtonConductAll()
    On Error GoTo EH
    WMSDBIu_ConductAllIssues
    Exit Sub
EH:
    MsgBox "Массовое проведение недоступно. Проверьте WMS_07 DB Issues 0.3.0." & _
           Chr(10) & CStr(Err) & " " & Error$,16,"WMS — Выдачи"
End Sub

Sub Issues_ButtonConductIssue()
    On Error GoTo Missing
    WMSDBIu_ConductSelectedIssue
    Exit Sub
Missing:
    MsgBox "Модуль WMS_07_DB_Issues не установлен или недоступен.",48,"WMS — Выдачи / БД"
End Sub

Sub Issues_ButtonCheckSelected()
    Issues_RefreshSelectedSuggestion
End Sub

Sub Issues_ButtonShowDiagnostics()
    Issues_ShowSelectedDiagnostics
End Sub

Sub Issues_ButtonValidateAll()
    Issues_RevalidateAll
End Sub

Sub Issues_ButtonRefreshLists()
    Issues_RebuildLists
End Sub

Sub Issues_ButtonSelfCheck()
    Issues_SelfCheck
End Sub

Sub Issues_SelfCheck()
    Dim s As String
    If Issues_RunSelfCheck(ThisComponent,s) Then
        MsgBox s,64,"WMS — Выдачи self-check"
    Else
        MsgBox s,48,"WMS — Выдачи self-check"
    End If
End Sub

Function Issues_SelfCheckText() As String
    Dim s As String
    Call Issues_RunSelfCheck(ThisComponent,s)
    Issues_SelfCheckText=s
End Function

Function Issues_LastReportText() As String
    Issues_LastReportText=gWMSISS_LastReport
End Function

Function Issues_CompileProbe() As String
    Issues_CompileProbe=WMSISS_VERSION & "|" & WMSISS_SCHEMA & "|" & WMSISS_SHEET & "|" & CStr(Issues_ExtensionHeaderCount())
End Function

Function Issues_IsInstalled() As Boolean
    Issues_IsInstalled=(Issues_GetSetting(ThisComponent,"WMS_IssuesVersion","")=WMSISS_VERSION)
End Function

Function Issues_InstallCore(oDoc As Object,bShow As Boolean,ByRef sReport As String) As Boolean
    Dim oSh As Object,sErr As String,sReval As String,bLocked As Boolean
    Issues_InstallCore=False:gWMSISS_LastError="":gWMSISS_LastReport="":sReport=""
    On Error GoTo Fatal
    If Not Issues_Preflight(oDoc,sErr) Then sReport=sErr:GoTo Done
    gWMSISS_Busy=True
    On Error Resume Next
    oDoc.lockControllers
    If Err=0 Then
        bLocked=True
    Else
        Err=0
    End If
    On Error GoTo Fatal
    oSh=oDoc.Sheets.getByName(WMSISS_SHEET)
    If Not Issues_EnsureExtensionColumns(oSh,sErr) Then sReport=sErr:GoTo SafeFail
    Issues_ResetHeaderCache
    Issues_EnsureSettings oDoc
    Issues_ApplyDesign oDoc,oSh
    Issues_RebuildHelperLists oDoc,oSh
    Issues_ApplyValidations oDoc,oSh
    If Not Issues_InstallEvent(oSh,sErr) Then sReport=sErr:GoTo SafeFail
    Issues_RefreshFilter oDoc,oSh
    Issues_ApplyFreeze oDoc,oSh
    Issues_ResetQueueCache
    Issues_ResetNomCache
    Issues_EmptyRowGuardCleanupCore oDoc,oSh
    If Not Issues_InstallButtonsCore(oDoc,oSh,sErr) Then sReport="Модуль установлен, но кнопки не созданы: " & sErr:GoTo SafeFail
    If Not Issues_RevalidateAllCore(oDoc,False,sReval) Then sReport="Модуль установлен частично, но проверка строк не завершилась: " & sReval:GoTo SafeFail
    If Not Issues_RunSelfCheck(oDoc,sErr) Then sReport=sErr:GoTo SafeFail
    Issues_SetSetting oDoc,"WMS_IssuesVersion",WMSISS_VERSION,"Версия production-контроллера листа Выдачи",True
    Issues_SetSetting oDoc,"WMS_IssuesState","READY","Состояние контроллера Выдачи",True
    Issues_SetSetting oDoc,"WMS_IssuesLastCheck",Issues_Stamp(Now),"Последний успешный self-check Выдачи",True
    Issues_LogAudit oDoc,"ISSUES_INSTALL","WMS_03_Issues " & WMSISS_VERSION & " установлен; schema " & WMSISS_SCHEMA
    sReport="WMS_03_Issues " & WMSISS_VERSION & " установлен." & Chr(10) & _
      "Лист 'Выдачи': стабильный режим ввода, кнопки действий, employee/returnable dropdowns, code suggestions, даты, частичные накопительные возвраты, SourceID/DIRTY/queue, фильтр и self-check — OK." & Chr(10) & _
      "Проведение движения и списание остатка будет выполнять общий Movement/Batch модуль; контроллер уже готов к интеграции."
    Issues_InstallCore=True:GoTo Done
SafeFail:
    gWMSISS_LastError=sReport
    Issues_LogError oDoc,"ISSUES_INSTALL",sReport
    Issues_SetSetting oDoc,"WMS_IssuesState","ERROR","Последняя установка/проверка Выдачи завершилась ошибкой",True
    GoTo Done
Fatal:
    sReport="WMS_03_Issues runtime error " & CStr(Err) & ": " & Error$
    Resume SafeFail
Done:
    If bLocked Then On Error Resume Next:oDoc.unlockControllers:On Error GoTo 0
    gWMSISS_Busy=False:gWMSISS_LastReport=sReport
End Function

' ============================================================================
' PREFLIGHT / SCHEMA / SETTINGS
' ============================================================================

Function Issues_Preflight(oDoc As Object,ByRef sErr As String) As Boolean
    Dim oSh As Object,a As Variant,i As Long,actual As String,builder As String
    Issues_Preflight=False:sErr=""
    On Error GoTo EH
    If IsNull(oDoc) Or IsEmpty(oDoc) Then sErr="Нет открытой книги Calc.":Exit Function
    If Not oDoc.supportsService("com.sun.star.sheet.SpreadsheetDocument") Then sErr="Модуль Выдачи работает только в LibreOffice Calc.":Exit Function
    If Not oDoc.Sheets.hasByName(WMSISS_SHEET) Then sErr="Нет листа 'Выдачи'. Сначала выполните WMS_00 CreateWorkbook.":Exit Function
    If Not oDoc.Sheets.hasByName(WMSISS_SETTINGS) Then sErr="Нет листа 'Настройки WMS'. Сначала выполните WMS_00 CreateWorkbook.":Exit Function
    builder=Issues_GetSetting(oDoc,"WMS_BuilderVersion","")
    If builder="" Then sErr="Книга не подтверждена как новая WMS, созданная WMS_00. Установка Выдачи остановлена.":Exit Function
    oSh=oDoc.Sheets.getByName(WMSISS_SHEET):a=Issues_BaseHeaders()
    For i=0 To UBound(a)
        actual=Trim(oSh.getCellByPosition(i,WMSISS_HEADER_ROW).String)
        If Issues_Canon(actual)<>Issues_Canon(CStr(a(i))) Then
            sErr="Лист 'Выдачи': столбец " & Issues_ColName(i) & " должен быть '" & CStr(a(i)) & "', сейчас '" & actual & "'. Автоматическая перестройка остановлена, чтобы не повредить данные."
            Exit Function
        End If
    Next i
    Issues_Preflight=True:Exit Function
EH:
    sErr="Preflight Выдачи: " & CStr(Err) & " " & Error$
End Function

Function Issues_BaseHeaders() As Variant
    Issues_BaseHeaders=Array("№","Код","Наименование","Кол-во","Ед. изм.","Кто получил","Дата","Откуда","Возвратный","Возвращено","Дата возврата","Примечание", _
        WMSISS_T_ID,WMSISS_T_STATE,WMSISS_T_HASH,WMSISS_T_ERROR,WMSISS_T_RETPOSTED,WMSISS_T_RETSEQ,WMSISS_T_RETHASH,WMSISS_T_RETERR,WMSISS_T_DIRTY,WMSISS_T_MODE,WMSISS_T_LEGACY,WMSISS_T_TOUCH)
End Function

Function Issues_ExtensionHeaders() As Variant
    Issues_ExtensionHeaders=Array(WMSISS_X_ROWVER,WMSISS_X_BASEFP,WMSISS_X_RETFP,WMSISS_X_FLAGS,WMSISS_X_SEVERITY,WMSISS_X_VALIDATED,WMSISS_X_CODESUG,WMSISS_X_SUGNAME,WMSISS_X_SUGSCORE,WMSISS_X_SUGSOURCE,WMSISS_X_EMPKEY,WMSISS_X_RETSTATE,WMSISS_X_LASTQMODE,WMSISS_X_RETHINT,WMSISS_X_UNITHINT,WMSISS_X_PLACEHINT)
End Function

Function Issues_ExtensionHeaderCount() As Long
    Dim a As Variant:a=Issues_ExtensionHeaders():Issues_ExtensionHeaderCount=UBound(a)-LBound(a)+1
End Function

Function Issues_EnsureExtensionColumns(oSh As Object,ByRef sErr As String) As Boolean
    Dim a As Variant,i As Long,c As Long,nAt As Long
    Issues_EnsureExtensionColumns=False:sErr=""
    On Error GoTo EH
    Issues_ResetHeaderCache
    a=Issues_ExtensionHeaders()
    For i=LBound(a) To UBound(a)
        c=Issues_FindHeader(oSh,CStr(a(i)))
        If c<0 Then
            nAt=Issues_LastHeaderColRaw(oSh)+1
            oSh.getCellByPosition(nAt,WMSISS_HEADER_ROW).String=CStr(a(i))
            oSh.Columns.getByIndex(nAt).IsVisible=False
            Issues_ResetHeaderCache
        Else
            oSh.Columns.getByIndex(c).IsVisible=False
        End If
    Next i
    Issues_ResetHeaderCache
    ' Base technical columns must also remain hidden.
    For c=12 To 23
        If Issues_Canon(oSh.getCellByPosition(c,0).String)=Issues_Canon(CStr(Issues_BaseHeaders()(c))) Then oSh.Columns.getByIndex(c).IsVisible=False
    Next c
    Issues_EnsureExtensionColumns=True:Exit Function
EH:
    sErr="Не удалось подготовить техполя Выдачи: " & CStr(Err) & " " & Error$
End Function

Sub Issues_EnsureSettings(oDoc As Object)
    Issues_SetSetting oDoc,"WMS_IssuesTargetLO",WMSISS_TARGET_LO,"Целевая версия LibreOffice контроллера Выдачи",False
    Issues_SetSetting oDoc,"WMS_IssuesFuzzySuggestion","ON","Нечёткий поиск используется только как подсказка кода, никогда как автослияние",False
    Issues_SetSetting oDoc,"WMS_IssuesEmployeeDropdown","AUTO_HISTORY","Список сотрудников строится из фактических значений Кто получил",False
    Issues_SetSetting oDoc,"WMS_IssuesBulkThreshold",CStr(WMSISS_BULK_THRESHOLD),"С какого количества строк включается облегчённая bulk-обработка",False
End Sub

Sub Issues_SetSetting(oDoc As Object,sKey As String,sValue As String,sDesc As String,bOverwrite As Boolean)
    Dim oSh As Object,r As Long
    If Not oDoc.Sheets.hasByName(WMSISS_SETTINGS) Then Exit Sub
    oSh=oDoc.Sheets.getByName(WMSISS_SETTINGS):r=Issues_FindSettingRow(oSh,sKey)
    If r<0 Then r=Issues_NextRow(oSh,0,1):oSh.getCellByPosition(0,r).String=sKey:oSh.getCellByPosition(1,r).String=sValue:oSh.getCellByPosition(2,r).String=sDesc
    If r>=0 And bOverwrite Then oSh.getCellByPosition(1,r).String=sValue:oSh.getCellByPosition(2,r).String=sDesc
End Sub

Function Issues_GetSetting(oDoc As Object,sKey As String,sDefault As String) As String
    Dim oSh As Object,r As Long
    Issues_GetSetting=sDefault
    On Error GoTo Done
    If Not oDoc.Sheets.hasByName(WMSISS_SETTINGS) Then Exit Function
    oSh=oDoc.Sheets.getByName(WMSISS_SETTINGS):r=Issues_FindSettingRow(oSh,sKey)
    If r>=0 Then Issues_GetSetting=Trim(oSh.getCellByPosition(1,r).String)
Done:
End Function

Function Issues_FindSettingRow(oSh As Object,sKey As String) As Long
    Dim r As Long,last As Long:Issues_FindSettingRow=-1:last=Issues_LastContentRow(oSh,2)
    For r=1 To last
        If Issues_Canon(oSh.getCellByPosition(0,r).String)=Issues_Canon(sKey) Then Issues_FindSettingRow=r:Exit Function
    Next r
End Function

' ============================================================================
' DESIGN / FORMATS / VALIDATIONS / FILTER / FREEZE
' ============================================================================

Sub Issues_ApplyDesign(oDoc As Object,oSh As Object)
    Dim last As Long,endRow As Long,c As Long,oHead As Object,oBody As Object
    On Error GoTo Done
    last=Issues_LastContentRow(oSh,WMSISS_USER_LAST_COL):endRow=last+200
    If endRow<WMSISS_MIN_STYLE_ROWS Then endRow=WMSISS_MIN_STYLE_ROWS
    If endRow>oSh.Rows.getCount()-1 Then endRow=oSh.Rows.getCount()-1
    oSh.TabColor=RGB(14,116,144)
    oHead=oSh.getCellRangeByPosition(0,0,WMSISS_USER_LAST_COL,0)
    oHead.CellBackColor=RGB(15,71,86):oHead.CharColor=RGB(255,255,255):oHead.CharWeight=com.sun.star.awt.FontWeight.BOLD
    oHead.CharHeight=10.5:oHead.HoriJustify=com.sun.star.table.CellHoriJustify.CENTER:oHead.VertJustify=com.sun.star.table.CellVertJustify.CENTER:oHead.IsTextWrapped=True
    oSh.Rows.getByIndex(0).Height=1050
    oBody=oSh.getCellRangeByPosition(0,1,WMSISS_USER_LAST_COL,endRow)
    oBody.CellBackColor=RGB(255,255,255):oBody.CharColor=RGB(31,41,55):oBody.CharHeight=9.5:oBody.VertJustify=com.sun.star.table.CellVertJustify.CENTER
    ' Operational bands: employee and return fields are visually obvious.
    oSh.getCellRangeByPosition(5,1,5,endRow).CellBackColor=RGB(255,251,235)
    oSh.getCellRangeByPosition(8,1,10,endRow).CellBackColor=RGB(239,246,255)
    oSh.getCellRangeByPosition(11,1,11,endRow).CellBackColor=RGB(248,250,252)
    Issues_ApplyTableGrid oSh,0,0,WMSISS_USER_LAST_COL,endRow
    Issues_SetWidth oSh,0,1250:Issues_SetWidth oSh,1,2400:Issues_SetWidth oSh,2,6200:Issues_SetWidth oSh,3,1800
    Issues_SetWidth oSh,4,2100:Issues_SetWidth oSh,5,4300:Issues_SetWidth oSh,6,2600:Issues_SetWidth oSh,7,3900
    Issues_SetWidth oSh,8,2700:Issues_SetWidth oSh,9,2400:Issues_SetWidth oSh,10,2900:Issues_SetWidth oSh,11,5200
    Issues_SetAlignCenter oSh,Array(0,1,3,4,6,8,9,10),1,endRow
    Issues_SetWrap oSh,Array(2,5,7,11),1,endRow
    Issues_ApplyNumberFormat oDoc,oSh.getCellRangeByPosition(0,1,2,endRow),"TEXT"
    Issues_ApplyNumberFormat oDoc,oSh.getCellRangeByPosition(3,1,3,endRow),"QTY"
    Issues_ApplyNumberFormat oDoc,oSh.getCellRangeByPosition(4,1,5,endRow),"TEXT"
    Issues_ApplyNumberFormat oDoc,oSh.getCellRangeByPosition(6,1,6,endRow),"DATE"
    Issues_ApplyNumberFormat oDoc,oSh.getCellRangeByPosition(7,1,8,endRow),"TEXT"
    Issues_ApplyNumberFormat oDoc,oSh.getCellRangeByPosition(9,1,9,endRow),"QTY"
    Issues_ApplyNumberFormat oDoc,oSh.getCellRangeByPosition(10,1,10,endRow),"DATE"
    Issues_ApplyNumberFormat oDoc,oSh.getCellRangeByPosition(11,1,11,endRow),"TEXT"
    For c=12 To Issues_LastHeaderCol(oSh)
        oSh.Columns.getByIndex(c).IsVisible=False
    Next c
Done:
End Sub

Sub Issues_ApplyTableGrid(oSh As Object,c1 As Long,r1 As Long,c2 As Long,r2 As Long)
    Dim oRange As Object,oTop As Object,oLeft As Object
    Dim ln As New com.sun.star.table.BorderLine2
    Dim outer As New com.sun.star.table.BorderLine2
    On Error GoTo Done
    oRange=oSh.getCellRangeByPosition(c1,r1,c2,r2)
    outer.Color=RGB(100,116,139):outer.LineStyle=com.sun.star.table.BorderLineStyle.SOLID:outer.LineWidth=50
    oTop=oSh.getCellRangeByPosition(c1,r1,c2,r1):oTop.TopBorder2=outer
    oLeft=oSh.getCellRangeByPosition(c1,r1,c1,r2):oLeft.LeftBorder2=outer
    ln.Color=RGB(176,186,198):ln.LineStyle=com.sun.star.table.BorderLineStyle.SOLID:ln.LineWidth=35
    oRange.BottomBorder2=ln:oRange.RightBorder2=ln
Done:
End Sub

Sub Issues_ApplyWorkingGrid(oSh As Object,r1 As Long,r2 As Long)
    If r1<1 Then r1=1
    If r2>=r1 Then Issues_ApplyTableGrid oSh,0,r1,WMSISS_USER_LAST_COL,r2
End Sub

Sub Issues_SetAlignCenter(oSh As Object,aCols As Variant,r1 As Long,r2 As Long)
    Dim i As Long,c As Long
    For i=LBound(aCols) To UBound(aCols)
        c=CLng(aCols(i))
        oSh.getCellRangeByPosition(c,r1,c,r2).HoriJustify=com.sun.star.table.CellHoriJustify.CENTER
    Next i
End Sub

Sub Issues_SetWrap(oSh As Object,aCols As Variant,r1 As Long,r2 As Long)
    Dim i As Long,c As Long
    For i=LBound(aCols) To UBound(aCols)
        c=CLng(aCols(i))
        oSh.getCellRangeByPosition(c,r1,c,r2).IsTextWrapped=True
    Next i
End Sub

Sub Issues_SetWidth(oSh As Object,c As Long,nWidth As Long)
    On Error Resume Next:oSh.Columns.getByIndex(c).Width=nWidth:On Error GoTo 0
End Sub

Sub Issues_ApplyNumberFormat(oDoc As Object,oRange As Object,sKind As String)
    Dim k As Long:k=Issues_NumberFormatKey(oDoc,sKind):If k>=0 Then oRange.NumberFormat=k
End Sub

Function Issues_NumberFormatKey(oDoc As Object,sKind As String) As Long
    Dim oFormats As Object,aLocale As New com.sun.star.lang.Locale,fmt As String,k As Long
    Issues_NumberFormatKey=-1
    On Error GoTo Done
    oFormats=oDoc.NumberFormats:aLocale.Language="ru":aLocale.Country="RU"
    Select Case UCase(sKind)
        Case "QTY":Issues_NumberFormatKey=oFormats.getStandardIndex(aLocale):Exit Function
        Case "DATE":fmt="DD.MM.YYYY"
        Case "TEXT":fmt="@"
        Case Else:Exit Function
    End Select
    k=oFormats.queryKey(fmt,aLocale,False):If k=-1 Then k=oFormats.addNew(fmt,aLocale)
    Issues_NumberFormatKey=k
Done:
End Function

Sub Issues_ApplyValidations(oDoc As Object,oSh As Object)
    Dim last As Long,endRow As Long,oV As Object,oRange As Object
    On Error GoTo Done
    last=Issues_LastContentRow(oSh,WMSISS_USER_LAST_COL):endRow=last+250
    If endRow<WMSISS_MIN_STYLE_ROWS Then endRow=WMSISS_MIN_STYLE_ROWS
    If endRow>oSh.Rows.getCount()-1 Then endRow=oSh.Rows.getCount()-1
    Issues_SetNumberValidation oSh.getCellRangeByPosition(3,1,3,endRow),True,"Количество","Количество выдачи должно быть положительным числом."
    Issues_SetNumberValidation oSh.getCellRangeByPosition(9,1,9,endRow),False,"Возвращено","Введите накопительное возвращённое количество от 0 до количества выдачи."
    oRange=oSh.getCellRangeByPosition(8,1,8,endRow):oV=oRange.Validation
    oV.Type=com.sun.star.sheet.ValidationType.LIST:oV.Operator=com.sun.star.sheet.ConditionOperator.EQUAL:oV.Formula1="""Да"";""Нет"""
    oV.IgnoreBlankCells=True:oV.ShowList=1:oV.ShowErrorMessage=True:oV.ErrorAlertStyle=com.sun.star.sheet.ValidationAlertStyle.STOP
    oV.ErrorTitle="WMS — возвратный":oV.ErrorMessage="Выберите Да или Нет.":oV.ShowInputMessage=True:oV.InputTitle="Возвратный":oV.InputMessage="Да — WMS будет вести накопительные частичные возвраты. Нет — поля Возвращено/Дата возврата должны оставаться пустыми или нулевыми."
    oRange.Validation=oV
    Issues_SetListValidation oSh.getCellRangeByPosition(5,1,5,endRow),WMSISS_LIST_EMP,"Кто получил","Выберите сотрудника из истории или введите нового вручную.",False
    Issues_SetListValidation oSh.getCellRangeByPosition(4,1,4,endRow),WMSISS_LIST_UNIT,"Ед. изм.","Можно выбрать ранее использованную единицу или ввести свою.",False
    Issues_SetListValidation oSh.getCellRangeByPosition(7,1,7,endRow),WMSISS_LIST_PLACE,"Откуда","Можно выбрать известное место хранения или ввести другое вручную.",False
    Issues_SetInputHelp oSh.getCellRangeByPosition(1,1,1,endRow),"Код","Код вводите сами. Если код пуст, WMS может показать только подсказку по похожему названию — автослияния по fuzzy нет."
    Issues_SetInputHelp oSh.getCellRangeByPosition(6,1,6,endRow),"Дата выдачи","Можно вводить 24.08, 24/08 или 24-08 — для текущей операции WMS нормализует дату с текущим годом."
    Issues_SetInputHelp oSh.getCellRangeByPosition(9,1,9,endRow),"Возвращено","Это НАКОПИТЕЛЬНЫЙ итог. Выдали 5: сначала 3, потом исправили на 4 — будущий Movement создаст только возврат +1."
    Issues_SetInputHelp oSh.getCellRangeByPosition(10,1,10,endRow),"Дата возврата","При Возвращено > 0 дата обязательна и не может быть раньше даты выдачи."
Done:
End Sub

Sub Issues_SetNumberValidation(oRange As Object,bPositive As Boolean,sTitle As String,sMessage As String)
    Dim oV As Object
    On Error GoTo Done
    oV=oRange.Validation:oV.Type=com.sun.star.sheet.ValidationType.DECIMAL
    If bPositive Then
        oV.Operator=com.sun.star.sheet.ConditionOperator.GREATER
        oV.Formula1="0"
    Else
        oV.Operator=com.sun.star.sheet.ConditionOperator.GREATER_EQUAL
        oV.Formula1="0"
    End If
    oV.IgnoreBlankCells=True:oV.ShowErrorMessage=True:oV.ErrorAlertStyle=com.sun.star.sheet.ValidationAlertStyle.WARNING:oV.ErrorTitle="WMS — " & sTitle:oV.ErrorMessage=sMessage:oRange.Validation=oV
Done:
End Sub

Sub Issues_SetListValidation(oRange As Object,sFormula As String,sTitle As String,sMessage As String,bStrict As Boolean)
    Dim oV As Object
    On Error GoTo Done
    oV=oRange.Validation:oV.Type=com.sun.star.sheet.ValidationType.LIST:oV.Operator=com.sun.star.sheet.ConditionOperator.EQUAL:oV.Formula1=sFormula
    oV.IgnoreBlankCells=True:oV.ShowList=1:oV.ShowInputMessage=True:oV.InputTitle=sTitle:oV.InputMessage=sMessage
    oV.ShowErrorMessage=bStrict
    If bStrict Then oV.ErrorAlertStyle=com.sun.star.sheet.ValidationAlertStyle.STOP:oV.ErrorTitle="WMS — " & sTitle:oV.ErrorMessage=sMessage
    oRange.Validation=oV
Done:
End Sub

Sub Issues_SetInputHelp(oRange As Object,sTitle As String,sMessage As String)
    Dim oV As Object
    On Error GoTo Done
    oV=oRange.Validation:oV.ShowInputMessage=True:oV.InputTitle=sTitle:oV.InputMessage=sMessage:oRange.Validation=oV
Done:
End Sub

Sub Issues_ApplyFreeze(oDoc As Object,oSh As Object)
    Dim ctl As Object
    On Error GoTo Done
    ctl=oDoc.CurrentController:If IsNull(ctl) Or IsEmpty(ctl) Then Exit Sub
    ctl.setActiveSheet oSh:ctl.freezeAtPosition 2,1
    On Error Resume Next:ctl.ZoomValue=90:On Error GoTo 0
Done:
End Sub

Sub Issues_RefreshFilter(oDoc As Object,oSh As Object)
    Dim dbs As Object,db As Object,addr As Variant,last As Long,endRow As Long,lastCol As Long,dbName As String
    On Error GoTo Done
    last=Issues_LastContentRow(oSh,WMSISS_USER_LAST_COL):endRow=last+WMSISS_FILTER_BUFFER
    If endRow<WMSISS_MIN_STYLE_ROWS Then endRow=WMSISS_MIN_STYLE_ROWS
    If endRow>oSh.Rows.getCount()-1 Then endRow=oSh.Rows.getCount()-1
    lastCol=Issues_LastHeaderCol(oSh):If lastCol<WMSISS_USER_LAST_COL Then lastCol=WMSISS_USER_LAST_COL
    addr=oSh.getCellRangeByPosition(0,0,lastCol,endRow).RangeAddress:dbs=oDoc.DatabaseRanges:dbName="WMS00_DB_03"
    If dbs.hasByName(dbName) Then
        db=dbs.getByName(dbName)
        db.setDataArea addr
    Else
        dbs.addNewByName dbName,addr
        db=dbs.getByName(dbName)
    End If
    db.AutoFilter=True
Done:
End Sub

Function Issues_InstallEvent(oSh As Object,ByRef sErr As String) As Boolean
    Dim ev(1) As New com.sun.star.beans.PropertyValue
    Issues_InstallEvent=False
    On Error GoTo EH
    ev(0).Name="EventType":ev(0).Value="Script"
    ev(1).Name="Script":ev(1).Value="vnd.sun.star.script:Standard.WMS_03_Issues_FINAL.Issues_OnContentChanged?language=Basic&location=document"
    oSh.Events.replaceByName "OnChange",ev():Issues_InstallEvent=True:Exit Function
EH:
    sErr="Не удалось назначить событие Content changed для Выдачи: " & CStr(Err) & " " & Error$
End Function

' ============================================================================
' DYNAMIC HELPER LISTS — HISTORY-BASED, NON-BLOCKING
' ============================================================================

Sub Issues_RebuildHelperLists(oDoc As Object,oSh As Object)
    Dim oSet As Object,aEmp() As String,aUnit() As String,aPlace() As String,nEmp As Long,nUnit As Long,nPlace As Long
    Dim r As Long,last As Long,s As String,oNom As Object,cUnit As Long,cPlace As Long
    On Error GoTo Done
    oSet=oDoc.Sheets.getByName(WMSISS_SETTINGS)
    ReDim aEmp(0 To WMSISS_HELP_MAX-1):ReDim aUnit(0 To WMSISS_HELP_MAX-1):ReDim aPlace(0 To WMSISS_HELP_MAX-1)
    last=Issues_LastContentRow(oSh,WMSISS_USER_LAST_COL)
    For r=1 To last
        s=Trim(oSh.getCellByPosition(5,r).String):Issues_AddUnique aEmp,nEmp,s
        s=Trim(oSh.getCellByPosition(4,r).String):Issues_AddUnique aUnit,nUnit,s
        s=Trim(oSh.getCellByPosition(7,r).String):Issues_AddUnique aPlace,nPlace,s
    Next r
    If oDoc.Sheets.hasByName(WMSISS_NOM) Then
        oNom=oDoc.Sheets.getByName(WMSISS_NOM):cUnit=Issues_FindHeader(oNom,"Базовая ед."):cPlace=Issues_FindHeader(oNom,"Место")
        last=Issues_LastContentRow(oNom,9)
        For r=1 To last
            If cUnit>=0 Then Issues_AddUnique aUnit,nUnit,Trim(oNom.getCellByPosition(cUnit,r).String)
            If cPlace>=0 Then Issues_AddUnique aPlace,nPlace,Trim(oNom.getCellByPosition(cPlace,r).String)
        Next r
    End If
    Issues_WriteHelperList oSet,WMSISS_HELP_EMP_COL,"_WMS_IssueEmployees",aEmp,nEmp
    Issues_WriteHelperList oSet,WMSISS_HELP_UNIT_COL,"_WMS_IssueUnits",aUnit,nUnit
    Issues_WriteHelperList oSet,WMSISS_HELP_PLACE_COL,"_WMS_IssuePlaces",aPlace,nPlace
    Issues_DefineListName oDoc,oSet,WMSISS_LIST_EMP,WMSISS_HELP_EMP_COL
    Issues_DefineListName oDoc,oSet,WMSISS_LIST_UNIT,WMSISS_HELP_UNIT_COL
    Issues_DefineListName oDoc,oSet,WMSISS_LIST_PLACE,WMSISS_HELP_PLACE_COL
    oSet.Columns.getByIndex(WMSISS_HELP_EMP_COL).IsVisible=False:oSet.Columns.getByIndex(WMSISS_HELP_UNIT_COL).IsVisible=False:oSet.Columns.getByIndex(WMSISS_HELP_PLACE_COL).IsVisible=False
Done:
End Sub

Sub Issues_AddUnique(a() As String,ByRef n As Long,s As String)
    Dim i As Long
    s=Trim(s):If s="" Or n>=WMSISS_HELP_MAX Then Exit Sub
    For i=0 To n-1:If Issues_Canon(a(i))=Issues_Canon(s) Then Exit Sub
    Next i
    a(n)=s:n=n+1
End Sub

Sub Issues_WriteHelperList(oSh As Object,c As Long,sHeader As String,a() As String,n As Long)
    Dim r As Long
    oSh.getCellRangeByPosition(c,0,c,WMSISS_HELP_MAX).clearContents(1023)
    oSh.getCellByPosition(c,0).String=sHeader
    For r=0 To n-1
        oSh.getCellByPosition(c,r+1).String=a(r)
    Next r
End Sub

Sub Issues_DefineListName(oDoc As Object,oSh As Object,sName As String,c As Long)
    Dim nrs As Object,aBase As Variant,content As String
    On Error GoTo Done
    nrs=oDoc.NamedRanges:aBase=oSh.getCellByPosition(c,0).CellAddress
    content="$'Настройки WMS'." & "$" & Issues_ColName(c) & "$2:$" & Issues_ColName(c) & "$" & CStr(WMSISS_HELP_MAX+1)
    If nrs.hasByName(sName) Then
        nrs.getByName(sName).setContent content
    Else
        nrs.addNewByName sName,content,aBase,0
    End If
Done:
End Sub

' ============================================================================
' SHEET EVENT — LIGHT CONTROLLER
' ============================================================================

Sub Issues_OnContentChanged(oEvent As Object)
    Dim oSh As Object,r1 As Long,r2 As Long,c1 As Long,c2 As Long
    Dim oCon As Object,sErr As String,r As Long

    If gWMSISS_Busy Then Exit Sub
    If Not Issues_GetChangedRange(oEvent,oSh,r1,r2,c1,c2) Then Exit Sub
    If oSh.Name<>WMSISS_SHEET Then Exit Sub

    ' Только B = Код. На любые другие изменения обработчик ничего не делает.
    If c1>1 Or c2<1 Then Exit Sub
    If r2<1 Then Exit Sub
    If r1<1 Then r1=1

    gWMSISS_Busy=True
    On Error GoTo EH

    sErr=""
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then GoTo SafeExit

    For r=r1 To r2
        If Trim(oSh.getCellByPosition(1,r).String)<>"" Then
            sErr=""
            Call Issues_AutofillRowByCodeCon(oCon,oSh,r,sErr)
        End If
    Next r

SafeExit:
    gWMSISS_Busy=False
    Exit Sub
EH:
    gWMSISS_Busy=False
End Sub

Function Issues_SQLQuote(s As String) As String
    Issues_SQLQuote=Replace(CStr(s),"'","''")
End Function

Function Issues_AutofillRowByCodeCon(oCon As Object,oSh As Object,r As Long,ByRef sErr As String) As Boolean
    Dim code As String,sql As String,oSt As Object,oRS As Object
    Dim nm As String,unitName As String,defLoc As String
    Dim loc As String,stockUnit As String,nLoc As Long
    Issues_AutofillRowByCodeCon=False
    sErr=""
    On Error GoTo EH

    code=Trim(oSh.getCellByPosition(1,r).String)
    If code="" Then Exit Function

    ' Основные данные товара.
    sql="SELECT PRODUCT_NAME,UNIT_NAME,DEFAULT_LOCATION FROM WMS_PRODUCTS " & _
        "WHERE PRODUCT_CODE='" & Issues_SQLQuote(code) & "'"
    oSt=oCon.createStatement()
    oRS=oSt.executeQuery(sql)

    If Not oRS.next() Then
        ' Не стираем пользовательский ввод, если код пока неизвестен.
        GoTo CloseQuery
    End If

    nm=Trim(oRS.getString(1))
    unitName=Trim(oRS.getString(2))
    defLoc=Trim(oRS.getString(3))
    ' OPT-1: release first result before replacing its variables.
    oRS.close()
    oSt.close()

    If nm<>"" Then oSh.getCellByPosition(2,r).String=nm
    If unitName<>"" Then oSh.getCellByPosition(4,r).String=unitName

    ' Текущее место хранения берём из фактического остатка.
    ' Автоподстановка безопасна только если положительный остаток находится
    ' ровно в одном месте. Если мест несколько — ничего не угадываем.
    sql="SELECT LOCATION_NAME,UNIT_NAME,SUM(QTY) FROM WMS_STOCK_MOVEMENTS " & _
        "WHERE PRODUCT_CODE='" & Issues_SQLQuote(code) & "' " & _
        "GROUP BY LOCATION_NAME,UNIT_NAME HAVING SUM(QTY)>0"
    oSt=oCon.createStatement()
    oRS=oSt.executeQuery(sql)
    nLoc=0:loc="":stockUnit=""
    Do While oRS.next()
        nLoc=nLoc+1
        If nLoc=1 Then
            loc=Trim(oRS.getString(1))
            stockUnit=Trim(oRS.getString(2))
        End If
        If nLoc>1 Then Exit Do
    Loop

    If nLoc=1 Then
        If loc<>"" Then oSh.getCellByPosition(7,r).String=loc
        If stockUnit<>"" Then oSh.getCellByPosition(4,r).String=stockUnit
    ElseIf nLoc=0 And defLoc<>"" Then
        ' Для нового товара без движения допускается справочное место.
        oSh.getCellByPosition(7,r).String=defLoc
    End If

    ' Дата выдачи — сегодня, только если пользователь её ещё не указал.
    If oSh.getCellByPosition(6,r).Value=0 And Trim(oSh.getCellByPosition(6,r).String)="" Then
        oSh.getCellByPosition(6,r).Value=CDbl(Date)
        On Error Resume Next
        Issues_WriteDate ThisComponent,oSh.getCellByPosition(6,r),CDbl(Date)
        Err=0
        On Error GoTo EH
    End If

    Issues_AutofillRowByCodeCon=True
    GoTo CloseQuery
EH:
    sErr="Автозаполнение выдачи по коду: " & CStr(Err) & " " & Error$
CloseQuery:
    On Error Resume Next
    oRS.close()
    oSt.close()
    On Error GoTo 0
End Function

Function Issues_GetChangedRange(oEvent As Object,ByRef oSh As Object,ByRef r1 As Long,ByRef r2 As Long,ByRef c1 As Long,ByRef c2 As Long) As Boolean
    Dim a As Variant:Issues_GetChangedRange=False
    On Error GoTo TryCell
    a=oEvent.RangeAddress:r1=a.StartRow:r2=a.EndRow:c1=a.StartColumn:c2=a.EndColumn:oSh=ThisComponent.Sheets.getByIndex(a.Sheet):Issues_GetChangedRange=True:Exit Function
TryCell:
    On Error GoTo Bad
    a=oEvent.CellAddress:r1=a.Row:r2=a.Row:c1=a.Column:c2=a.Column:oSh=ThisComponent.Sheets.getByIndex(a.Sheet):Issues_GetChangedRange=True
Bad:
End Function

Sub Issues_EnsureChangedRangeFormat(oDoc As Object,oSh As Object,r1 As Long,r2 As Long)
    On Error GoTo Done
    Issues_ApplyNumberFormat oDoc,oSh.getCellRangeByPosition(0,r1,2,r2),"TEXT"
    Issues_ApplyNumberFormat oDoc,oSh.getCellRangeByPosition(3,r1,3,r2),"QTY"
    Issues_ApplyNumberFormat oDoc,oSh.getCellRangeByPosition(4,r1,5,r2),"TEXT"
    Issues_ApplyNumberFormat oDoc,oSh.getCellRangeByPosition(6,r1,6,r2),"DATE"
    Issues_ApplyNumberFormat oDoc,oSh.getCellRangeByPosition(7,r1,8,r2),"TEXT"
    Issues_ApplyNumberFormat oDoc,oSh.getCellRangeByPosition(9,r1,9,r2),"QTY"
    Issues_ApplyNumberFormat oDoc,oSh.getCellRangeByPosition(10,r1,10,r2),"DATE"
    Issues_ApplyNumberFormat oDoc,oSh.getCellRangeByPosition(11,r1,11,r2),"TEXT"
    Issues_ApplyWorkingGrid oSh,r1,r2
Done:
End Sub

Sub Issues_NormalizeDates(oDoc As Object,oSh As Object,r As Long,c1 As Long,c2 As Long)
    Dim c As Long,v As Double
    If c1<=6 And c2>=6 Then c=6:v=Issues_CellDate(oSh.getCellByPosition(c,r),True):If v>0 Then Issues_WriteDate oDoc,oSh.getCellByPosition(c,r),v
    If c1<=10 And c2>=10 Then c=10:v=Issues_CellDate(oSh.getCellByPosition(c,r),True):If v>0 Then Issues_WriteDate oDoc,oSh.getCellByPosition(c,r),v
End Sub

Sub Issues_ProcessBulkLight(oDoc As Object,oSh As Object,r1 As Long,r2 As Long,c1 As Long,c2 As Long)
    Dim aUser As Variant,aTech As Variant,outTech() As Variant,outIssue() As Variant,outReturn() As Variant
    Dim i As Long,n As Long,r As Long,t As Variant,u As Variant,stamp As String
    Dim sFlags As String,sSeverity As String,sError As String,sRetError As String,sID As String,sState As String,sMode As String,sLegacy As String
    Dim sNo As String,sCode As String,sName As String,sUnit As String,sEmp As String,sFrom As String,sReturnable As String
    Dim qty As Double,ret As Double,posted As Double,bQty As Boolean,bRet As Boolean,dIssue As Double,dReturn As Double
    Dim idx As Long,sugCode As String,sugName As String,sugSource As String,sRetHint As String,sUnitHint As String,sPlaceHint As String,score As Double
    Dim baseFP As String,oldBase As String,retFP As String,qMode As String,nVer As Double,issueOut As Variant,returnOut As Variant
    Dim badRows() As Long,badCount As Long,techLast As Long
    n=r2-r1+1:If n<=0 Then Exit Sub
    techLast=Issues_FindHeader(oSh,WMSISS_X_PLACEHINT):If techLast<12 Then Exit Sub
    aUser=oSh.getCellRangeByPosition(0,r1,WMSISS_USER_LAST_COL,r2).getDataArray()
    aTech=oSh.getCellRangeByPosition(12,r1,techLast,r2).getDataArray()
    ReDim outTech(0 To n-1):ReDim outIssue(0 To n-1):ReDim outReturn(0 To n-1):ReDim badRows(0 To n-1)
    stamp=Issues_Stamp(Now)
    Call Issues_LoadNomCache(oDoc)
    gWMSISS_BulkMode=True
    For i=0 To n-1
        r=r1+i:u=aUser(i):t=aTech(i)
        issueOut=u(6):returnOut=u(10)
        If Issues_BulkRowHasBusiness(u) Then
            sFlags="":sNo=Trim(CStr(u(0))):sCode=Trim(CStr(u(1))):sName=Trim(CStr(u(2))):sUnit=Trim(CStr(u(4))):sEmp=Trim(CStr(u(5))):sFrom=Trim(CStr(u(7))):sReturnable=Issues_Canon(CStr(u(8)))
            bQty=Issues_BulkTryNumber(u(3),qty):bRet=Issues_BulkTryNumber(u(9),ret):If Not bRet Then ret=0
            posted=Issues_BulkNumber(t(4))
            dIssue=Issues_BulkDateValue(u(6),True):dReturn=Issues_BulkDateValue(u(10),True)
            If dIssue>0 Then issueOut=dIssue
            If dReturn>0 Then returnOut=dReturn
            sID=Trim(CStr(t(0))):If sID="" Then sID=Issues_NewSourceID(r)
            sState=Trim(CStr(t(1))):If sState="" Then sState=WMSISS_STATE_DRAFT
            sMode=Trim(CStr(t(9))):If sMode="" Then sMode="CURRENT"
            sLegacy=Trim(CStr(t(10))):If UCase(sMode)="OPEN_LEGACY" And sLegacy="" Then sLegacy=Issues_NewLegacyKey(r)
            If sNo="" Then Issues_AddFlag sFlags,"WARN:NO_NUMBER:Не заполнен № выдачи."
            If sName="" Then Issues_AddFlag sFlags,"CRIT:NO_NAME:Не заполнено Наименование."
            If Not bQty Or qty<=0 Then Issues_AddFlag sFlags,"CRIT:QTY_BAD:Количество выдачи должно быть больше нуля."
            If sUnit="" Then Issues_AddFlag sFlags,"CRIT:NO_UNIT:Не заполнена Ед. изм."
            If sEmp="" Then Issues_AddFlag sFlags,"CRIT:NO_EMPLOYEE:Не заполнено Кто получил."
            If dIssue<=0 Then
                Issues_AddFlag sFlags,"CRIT:ISSUE_DATE_BAD:Не заполнена или некорректна дата выдачи."
            ElseIf dIssue>CDbl(Date)+WMSISS_EPS Then
                Issues_AddFlag sFlags,"WARN:ISSUE_DATE_FUTURE:Дата выдачи находится в будущем."
            End If
            If sFrom="" Then Issues_AddFlag sFlags,"CRIT:NO_FROM:Не заполнено Откуда."
            If sReturnable<>Issues_Canon(WMSISS_YES) And sReturnable<>Issues_Canon(WMSISS_NO) Then Issues_AddFlag sFlags,"CRIT:RETURNABLE_BAD:Выберите Возвратный = Да или Нет."
            If ret<0 Then Issues_AddFlag sFlags,"CRIT:RETURN_NEGATIVE:Возвращено не может быть отрицательным."
            If bQty And qty>0 And ret>qty+WMSISS_EPS Then Issues_AddFlag sFlags,"CRIT:RETURN_GT_ISSUED:Возвращено больше выданного."
            If posted<0 Then Issues_AddFlag sFlags,"CRIT:POSTED_RETURN_NEGATIVE:Системное проведённое количество возврата отрицательное."
            If bQty And posted>qty+WMSISS_EPS Then Issues_AddFlag sFlags,"CRIT:POSTED_GT_ISSUED:В истории проведено возвратов больше, чем выдано."
            If ret+WMSISS_EPS<posted Then Issues_AddFlag sFlags,"CRIT:RETURN_BELOW_POSTED:Нельзя уменьшить Возвращено ниже уже проведённого возврата " & Issues_NumText(posted) & ". Историю нужно отменять безопасной операцией."
            If sReturnable=Issues_Canon(WMSISS_NO) Then
                If ret>WMSISS_EPS Then Issues_AddFlag sFlags,"CRIT:NONRETURNABLE_HAS_RETURN:Для невозвратной выдачи Возвращено должно быть 0/пусто."
                If dReturn>0 Then Issues_AddFlag sFlags,"WARN:NONRETURNABLE_RETURN_DATE:Для невозвратной выдачи указана дата возврата."
            ElseIf sReturnable=Issues_Canon(WMSISS_YES) Then
                If ret>WMSISS_EPS And dReturn<=0 Then Issues_AddFlag sFlags,"CRIT:RETURN_DATE_REQUIRED:При Возвращено > 0 обязательна Дата возврата."
                If ret<=WMSISS_EPS And dReturn>0 Then Issues_AddFlag sFlags,"WARN:RETURN_DATE_WITH_ZERO:Дата возврата заполнена, но Возвращено = 0."
                If ret>WMSISS_EPS And dReturn>0 And dIssue>0 And dReturn<dIssue-WMSISS_EPS Then Issues_AddFlag sFlags,"CRIT:RETURN_BEFORE_ISSUE:Дата возврата не может быть раньше даты выдачи."
            End If
            sugCode="":sugName="":score=0:sugSource="":sRetHint="":sUnitHint="":sPlaceHint=""
            If gWMSISS_NomCount>0 Then
                If sCode<>"" Then
                    idx=Issues_NomIndexByCode(sCode)
                    If idx>=0 Then
                        sugCode=gWMSISS_NomCode(idx):sugName=gWMSISS_NomName(idx):score=100:sugSource="EXACT_CODE":sRetHint=gWMSISS_NomReturnable(idx):sUnitHint=gWMSISS_NomUnit(idx):sPlaceHint=gWMSISS_NomPlace(idx)
                        If Not gWMSISS_NomActive(idx) Then Issues_AddFlag sFlags,"WARN:CODE_INACTIVE:Код найден в номенклатуре, но позиция не активна."
                        If sName<>"" And sugName<>"" And Issues_NameSimilarity(sName,sugName)<45 Then Issues_AddFlag sFlags,"WARN:CODE_NAME_MISMATCH:Код относится к '" & sugName & "'. Проверьте наименование."
                        If sUnit<>"" And sUnitHint<>"" And Issues_CanonUnit(sUnit)<>Issues_CanonUnit(sUnitHint) Then Issues_AddFlag sFlags,"WARN:UNIT_HINT:В номенклатуре базовая ед. '" & sUnitHint & "'. Проверьте единицу выдачи."
                        If sFrom<>"" And sPlaceHint<>"" And Issues_Canon(sFrom)<>Issues_Canon(sPlaceHint) Then Issues_AddFlag sFlags,"INFO:PLACE_HINT:В номенклатуре основное место '" & sPlaceHint & "'."
                        If sReturnable<>"" And sRetHint<>"" And Issues_Canon(sRetHint)<>sReturnable Then Issues_AddFlag sFlags,"WARN:RETURNABLE_HINT:В номенклатуре Возвратный = '" & sRetHint & "'. Проверьте выбор."
                    Else
                        Issues_AddFlag sFlags,"WARN:CODE_UNKNOWN:Код пока не найден в номенклатуре."
                    End If
                Else
                    Issues_AddFlag sFlags,"WARN:NO_CODE:Код не заполнен. Для старых позиций допустимо, но идентификация будет слабее."
                End If
            ElseIf sCode="" Then
                Issues_AddFlag sFlags,"WARN:NO_CODE:Код не заполнен. Для старых позиций допустимо, но идентификация будет слабее."
            End If
            baseFP=Issues_BulkBaseFingerprint(u,dIssue):retFP=Issues_BulkReturnFingerprint(u,dReturn):oldBase=Trim(CStr(t(13)))
            If Issues_Canon(sState)=Issues_Canon(WMSISS_STATE_DONE) And oldBase<>"" And oldBase<>baseFP Then Issues_AddFlag sFlags,"CRIT:POSTED_CHANGED:Проведённая выдача изменена. Сначала отмените старое движение и перепроведите безопасно."
            sSeverity=Issues_SeverityFromFlags(sFlags):sError=Issues_FirstProblem(sFlags):sRetError=Issues_FirstReturnProblem(sFlags)
            nVer=Issues_BulkNumber(t(12))+1
            qMode="ISSUE":If ret>WMSISS_EPS Or dReturn>0 Then qMode="ISSUE_RETURN"
            If oldBase="" Or Issues_Canon(sState)<>Issues_Canon(WMSISS_STATE_DONE) Then oldBase=baseFP
            outTech(i)=Array(sID,sState,CStr(t(2)),sError,posted,Issues_BulkNumber(t(5)),CStr(t(6)),sRetError,"1",sMode,sLegacy,stamp,nVer,oldBase,retFP,sFlags,sSeverity,stamp,sugCode,sugName,score,sugSource,Issues_Canon(sEmp),Issues_ReturnStateText(sReturnable,qty,ret,sFlags),qMode,sRetHint,sUnitHint,sPlaceHint)
            If sSeverity<>"OK" Then badRows(badCount)=r:badCount=badCount+1
        Else
            sState=Trim(CStr(t(1))):posted=Issues_BulkNumber(t(4))
            If Issues_Canon(sState)=Issues_Canon(WMSISS_STATE_DONE) Or posted>WMSISS_EPS Then
                t(3)="Проведённую выдачу/возврат нельзя очищать как обычную строку. Восстановите данные или используйте безопасную отмену движения."
                t(15)="CRIT:POSTED_ROW_CLEARED:Проведённую выдачу/возврат нельзя очищать как обычную строку.":t(16)="CRITICAL":badRows(badCount)=r:badCount=badCount+1
                outTech(i)=t
            Else
                outTech(i)=Array("","","","",0,0,"","","","","","",0,"","","","","","","",0,"","","","","","","")
            End If
        End If
        outIssue(i)=Array(issueOut):outReturn(i)=Array(returnOut)
    Next i
    oSh.getCellRangeByPosition(12,r1,techLast,r2).setDataArray(outTech())
    If c1<=6 And c2>=6 Then oSh.getCellRangeByPosition(6,r1,6,r2).setDataArray(outIssue())
    If c1<=10 And c2>=10 Then oSh.getCellRangeByPosition(10,r1,10,r2).setDataArray(outReturn())
    gWMSISS_BulkMode=False
    Issues_QueueRowsBatch oDoc,oSh,r1,r2
    ' Base visual is applied by range; expensive cell-by-cell highlights only for problem rows.
    oSh.getCellRangeByPosition(0,r1,WMSISS_USER_LAST_COL,r2).CellBackColor=RGB(255,255,255)
    oSh.getCellRangeByPosition(0,r1,WMSISS_USER_LAST_COL,r2).CharColor=RGB(31,41,55)
    oSh.getCellRangeByPosition(0,r1,WMSISS_USER_LAST_COL,r2).CharWeight=com.sun.star.awt.FontWeight.NORMAL
    oSh.getCellRangeByPosition(5,r1,5,r2).CellBackColor=RGB(255,251,235)
    oSh.getCellRangeByPosition(8,r1,10,r2).CellBackColor=RGB(239,246,255)
    oSh.getCellRangeByPosition(11,r1,11,r2).CellBackColor=RGB(248,250,252)
    For i=0 To badCount-1:Issues_ApplyRowVisual oSh,badRows(i):Next i
End Sub

Function Issues_BulkRowHasBusiness(aRow As Variant) As Boolean
    Dim c As Long,s As String
    ' Empty Row Guard 1.0.3:
    ' column A (№) alone is only a visual/prepared row number and MUST NOT create
    ' SourceID / DIRTY / validation / queue work.  A row becomes a business row
    ' only when at least one real field B:L contains data.
    For c=1 To WMSISS_USER_LAST_COL
        s=Trim(CStr(aRow(c)))
        If s<>"" Then Issues_BulkRowHasBusiness=True:Exit Function
    Next c
End Function

Function Issues_BulkTryNumber(v As Variant,ByRef d As Double) As Boolean
    Dim s As String:d=0
    On Error GoTo Bad
    If IsNumeric(v) Then d=CDbl(v):Issues_BulkTryNumber=True:Exit Function
    s=Trim(CStr(v)):If s="" Then Exit Function
    s=Replace(s,Chr(160),""):s=Replace(s," ",""):s=Replace(s,",",".")
    If IsNumeric(s) Then d=CDbl(s):Issues_BulkTryNumber=True
    Exit Function
Bad:
End Function

Function Issues_BulkNumber(v As Variant) As Double
    Dim d As Double:If Issues_BulkTryNumber(v,d) Then Issues_BulkNumber=d
End Function

Function Issues_BulkDateValue(v As Variant,bInferYear As Boolean) As Double
    Dim s As String,p As Variant,d As Integer,m As Integer,y As Integer,dt As Date,x As Double
    On Error GoTo Bad
    If IsNumeric(v) Then x=CDbl(v):If x>0 Then Issues_BulkDateValue=x:Exit Function
    s=Trim(CStr(v)):If s="" Then Exit Function
    s=Replace(s,"/","."):s=Replace(s,"-","."):p=Split(s,".")
    If UBound(p)=1 And bInferYear Then
        d=CInt(p(0)):m=CInt(p(1)):y=Year(Date)
    ElseIf UBound(p)=2 Then
        If Len(p(0))=4 Then y=CInt(p(0)):m=CInt(p(1)):d=CInt(p(2)) Else d=CInt(p(0)):m=CInt(p(1)):y=CInt(p(2)):If y<100 Then y=2000+y
    Else
        GoTo Bad
    End If
    dt=DateSerial(y,m,d):If Year(dt)=y And Month(dt)=m And Day(dt)=d Then Issues_BulkDateValue=CDbl(dt)
Bad:
End Function

Function Issues_BulkBaseFingerprint(aRow As Variant,dIssue As Double) As String
    Dim c As Long,s As String,v As String
    For c=0 To WMSISS_USER_LAST_COL
        If c<>9 And c<>10 Then
            If c=6 And dIssue>0 Then v=CStr(dIssue) Else v=CStr(aRow(c))
            s=s & "|" & CStr(c) & "=" & v
        End If
    Next c
    Issues_BulkBaseFingerprint=Issues_SimpleHash(s)
End Function

Function Issues_BulkReturnFingerprint(aRow As Variant,dReturn As Double) As String
    Dim v As String:v=CStr(aRow(10)):If dReturn>0 Then v=CStr(dReturn)
    Issues_BulkReturnFingerprint=Issues_SimpleHash(CStr(aRow(9)) & "|" & v)
End Function

Sub Issues_HandleBlankBusinessRow(oSh As Object,r As Long)
    Dim cState As Long,cPosted As Long,cErr As Long,cFlags As Long,cSev As Long,state As String,posted As Double,a As Variant,i As Long,c As Long
    cState=Issues_FindHeader(oSh,WMSISS_T_STATE)
    cPosted=Issues_FindHeader(oSh,WMSISS_T_RETPOSTED)
    cErr=Issues_FindHeader(oSh,WMSISS_T_ERROR)
    cFlags=Issues_FindHeader(oSh,WMSISS_X_FLAGS)
    cSev=Issues_FindHeader(oSh,WMSISS_X_SEVERITY)
    If cState>=0 Then state=Trim(oSh.getCellByPosition(cState,r).String)
    If cPosted>=0 Then posted=Issues_CellNumber(oSh.getCellByPosition(cPosted,r))
    If Issues_Canon(state)=Issues_Canon(WMSISS_STATE_DONE) Or posted>WMSISS_EPS Then
        If cErr>=0 Then oSh.getCellByPosition(cErr,r).String="Проведённую выдачу/возврат нельзя очищать как обычную строку. Восстановите данные или используйте безопасную отмену движения."
        If cFlags>=0 Then oSh.getCellByPosition(cFlags,r).String="CRIT:POSTED_ROW_CLEARED:Проведённую выдачу/возврат нельзя очищать как обычную строку."
        If cSev>=0 Then oSh.getCellByPosition(cSev,r).String="CRITICAL"
        oSh.getCellRangeByPosition(0,r,WMSISS_USER_LAST_COL,r).CellBackColor=RGB(254,202,202)
        Issues_SetWMSAnnotation oSh.getCellByPosition(0,r),"WMS: Проведённую выдачу/возврат нельзя очищать. Восстановите строку или выполните безопасную отмену."
    Else
        ' A never-posted blank draft has no business identity and can be cleaned completely.
        For c=12 To Issues_LastHeaderCol(oSh)
            oSh.getCellByPosition(c,r).String=""
        Next c
        Issues_ClearWMSAnnotation oSh.getCellByPosition(0,r)
        Issues_ClearWMSAnnotation oSh.getCellByPosition(1,r)
        oSh.getCellRangeByPosition(0,r,WMSISS_USER_LAST_COL,r).CellBackColor=RGB(255,255,255)
        oSh.getCellByPosition(5,r).CellBackColor=RGB(255,251,235)
        oSh.getCellRangeByPosition(8,r,10,r).CellBackColor=RGB(239,246,255)
        oSh.getCellByPosition(11,r).CellBackColor=RGB(248,250,252)
    End If
End Sub

' ============================================================================
' SOURCE ID / SORT-COPY SAFETY
' ============================================================================

Sub Issues_EnsureSIDUniquenessForEvent(oSh As Object,r1 As Long,r2 As Long)
    Dim cID As Long,last As Long,aAll As Variant,aBiz As Variant,aEvt As Variant,outIDs() As Variant
    Dim cap As Long,keys() As String,counts() As Long,i As Long,s As String,slot As Long,found As Boolean,newID As String,n As Long
    cID=Issues_FindHeader(oSh,WMSISS_T_ID):If cID<0 Then Exit Sub
    last=Issues_LastContentRow(oSh,WMSISS_USER_LAST_COL):If last<1 Then Exit Sub
    aAll=oSh.getCellRangeByPosition(cID,1,cID,last).getDataArray()
    aBiz=oSh.getCellRangeByPosition(0,r1,WMSISS_USER_LAST_COL,r2).getDataArray()
    aEvt=oSh.getCellRangeByPosition(cID,r1,cID,r2).getDataArray()
    n=r2-r1+1:ReDim outIDs(0 To n-1)
    cap=16:Do While cap<(last+1)*2:cap=cap*2:Loop
    ReDim keys(0 To cap-1):ReDim counts(0 To cap-1)
    For i=0 To UBound(aAll)
        s=Trim(CStr(aAll(i)(0)))
        If s<>"" Then slot=Issues_LocalHashSlot(keys,cap,s,found):If found Then counts(slot)=counts(slot)+1 Else keys(slot)=s:counts(slot)=1
    Next i
    For i=0 To n-1
        s=Trim(CStr(aEvt(i)(0)))
        If Issues_BulkRowHasBusiness(aBiz(i)) Then
            If s="" Then
                s=Issues_NewSourceID(r1+i)
            Else
                slot=Issues_LocalHashSlot(keys,cap,s,found)
                If found And counts(slot)>1 Then s=Issues_NewSourceID(r1+i):counts(slot)=counts(slot)-1
            End If
        End If
        outIDs(i)=Array(s)
    Next i
    oSh.getCellRangeByPosition(cID,r1,cID,r2).setDataArray(outIDs())
End Sub

Function Issues_LocalHashSlot(keys() As String,cap As Long,sKey As String,ByRef bFound As Boolean) As Long
    Dim h As Long,i As Long,ch As Long,slot As Long
    h=0
    For i=1 To Len(sKey)
        ch=Asc(Mid(sKey,i,1))
        If ch<0 Then ch=ch+256
        h=(h*31+ch) Mod 2147483
    Next i
    slot=h Mod cap
    Do
        If keys(slot)="" Then
            keys(slot)=sKey
            bFound=False
            Issues_LocalHashSlot=slot
            Exit Function
        End If
        If keys(slot)=sKey Then
            bFound=True
            Issues_LocalHashSlot=slot
            Exit Function
        End If
        slot=slot+1
        If slot>=cap Then slot=0
    Loop
End Function

Function Issues_NewSourceID(r As Long) As String
    Randomize:Issues_NewSourceID="ISS-" & Issues_TimestampCompact(Now) & "-" & Right("00000" & CStr(r+1),5) & "-" & Right("000000" & CStr(CLng(Rnd()*999999)),6)
End Function

Function Issues_NewLegacyKey(r As Long) As String
    Randomize:Issues_NewLegacyKey="LEG-ISS-" & Issues_TimestampCompact(Now) & "-" & Right("00000" & CStr(r+1),5) & "-" & Right("0000" & CStr(CLng(Rnd()*9999)),4)
End Function

' ============================================================================
' ROW VALIDATION / SUGGESTIONS / RETURN LOGIC
' ============================================================================

Sub Issues_ValidateRow(oDoc As Object,oSh As Object,r As Long,bMarkDirty As Boolean,bAllowFuzzy As Boolean)
    Dim cID As Long,cState As Long,cErr As Long,cRetErr As Long,cDirty As Long,cMode As Long,cLegacy As Long,cTouch As Long
    Dim cVer As Long,cBase As Long,cRetFp As Long,cFlags As Long,cSev As Long,cValidated As Long,cCodeSug As Long,cSugName As Long,cSugScore As Long,cSugSource As Long,cEmp As Long,cRetState As Long,cLastQ As Long,cRetHint As Long,cUnitHint As Long,cPlaceHint As Long
    Dim sFlags As String,sSeverity As String,sError As String,sRetError As String,sID As String,sState As String,sMode As String,baseFP As String,oldBase As String,retFP As String
    Dim sNo As String,sCode As String,sName As String,sUnit As String,sEmp As String,sFrom As String,sReturnable As String,sNote As String
    Dim qty As Double,ret As Double,posted As Double,bQty As Boolean,bRet As Boolean,dIssue As Double,dReturn As Double
    Dim idx As Long,sugCode As String,sugName As String,sugSource As String,sRetHint As String,sUnitHint As String,sPlaceHint As String,score As Double
    Dim qMode As String,nVer As Double
    If r<1 Or Not Issues_RowHasBusinessData(oSh,r) Then Exit Sub
    cID=Issues_FindHeader(oSh,WMSISS_T_ID):cState=Issues_FindHeader(oSh,WMSISS_T_STATE):cErr=Issues_FindHeader(oSh,WMSISS_T_ERROR):cRetErr=Issues_FindHeader(oSh,WMSISS_T_RETERR)
    cDirty=Issues_FindHeader(oSh,WMSISS_T_DIRTY):cMode=Issues_FindHeader(oSh,WMSISS_T_MODE):cLegacy=Issues_FindHeader(oSh,WMSISS_T_LEGACY):cTouch=Issues_FindHeader(oSh,WMSISS_T_TOUCH)
    cVer=Issues_FindHeader(oSh,WMSISS_X_ROWVER)
    cBase=Issues_FindHeader(oSh,WMSISS_X_BASEFP)
    cRetFp=Issues_FindHeader(oSh,WMSISS_X_RETFP)
    cFlags=Issues_FindHeader(oSh,WMSISS_X_FLAGS)
    cSev=Issues_FindHeader(oSh,WMSISS_X_SEVERITY)
    cValidated=Issues_FindHeader(oSh,WMSISS_X_VALIDATED)
    cCodeSug=Issues_FindHeader(oSh,WMSISS_X_CODESUG)
    cSugName=Issues_FindHeader(oSh,WMSISS_X_SUGNAME)
    cSugScore=Issues_FindHeader(oSh,WMSISS_X_SUGSCORE)
    cSugSource=Issues_FindHeader(oSh,WMSISS_X_SUGSOURCE)
    cEmp=Issues_FindHeader(oSh,WMSISS_X_EMPKEY)
    cRetState=Issues_FindHeader(oSh,WMSISS_X_RETSTATE)
    cLastQ=Issues_FindHeader(oSh,WMSISS_X_LASTQMODE)
    cRetHint=Issues_FindHeader(oSh,WMSISS_X_RETHINT)
    cUnitHint=Issues_FindHeader(oSh,WMSISS_X_UNITHINT)
    cPlaceHint=Issues_FindHeader(oSh,WMSISS_X_PLACEHINT)
    sNo=Trim(oSh.getCellByPosition(0,r).String)
    sCode=Trim(oSh.getCellByPosition(1,r).String)
    sName=Trim(oSh.getCellByPosition(2,r).String)
    sUnit=Trim(oSh.getCellByPosition(4,r).String)
    sEmp=Trim(oSh.getCellByPosition(5,r).String)
    sFrom=Trim(oSh.getCellByPosition(7,r).String)
    sReturnable=Issues_Canon(oSh.getCellByPosition(8,r).String)
    sNote=Trim(oSh.getCellByPosition(11,r).String)
    bQty=Issues_TryNumber(oSh.getCellByPosition(3,r),qty)
    bRet=Issues_TryNumber(oSh.getCellByPosition(9,r),ret)
    posted=Issues_CellNumber(oSh.getCellByPosition(Issues_FindHeader(oSh,WMSISS_T_RETPOSTED),r))
    dIssue=Issues_CellDate(oSh.getCellByPosition(6,r),True):dReturn=Issues_CellDate(oSh.getCellByPosition(10,r),True)
    sID=Trim(oSh.getCellByPosition(cID,r).String):If sID="" Then sID=Issues_NewSourceID(r):oSh.getCellByPosition(cID,r).String=sID
    sState=Trim(oSh.getCellByPosition(cState,r).String):If sState="" Then sState=WMSISS_STATE_DRAFT:oSh.getCellByPosition(cState,r).String=sState
    sMode=Trim(oSh.getCellByPosition(cMode,r).String):If sMode="" Then sMode="CURRENT":oSh.getCellByPosition(cMode,r).String=sMode
    If UCase(sMode)="OPEN_LEGACY" And Trim(oSh.getCellByPosition(cLegacy,r).String)="" Then oSh.getCellByPosition(cLegacy,r).String=Issues_NewLegacyKey(r)
    If sNo="" Then Issues_AddFlag sFlags,"WARN:NO_NUMBER:Не заполнен № выдачи."
    If sName="" Then Issues_AddFlag sFlags,"CRIT:NO_NAME:Не заполнено Наименование."
    If Not bQty Or qty<=0 Then Issues_AddFlag sFlags,"CRIT:QTY_BAD:Количество выдачи должно быть больше нуля."
    If sUnit="" Then Issues_AddFlag sFlags,"CRIT:NO_UNIT:Не заполнена Ед. изм."
    If sEmp="" Then Issues_AddFlag sFlags,"CRIT:NO_EMPLOYEE:Не заполнено Кто получил."
    If dIssue<=0 Then
        Issues_AddFlag sFlags,"CRIT:ISSUE_DATE_BAD:Не заполнена или некорректна дата выдачи."
    ElseIf dIssue>CDbl(Date)+WMSISS_EPS Then
        Issues_AddFlag sFlags,"WARN:ISSUE_DATE_FUTURE:Дата выдачи находится в будущем."
    End If
    If sFrom="" Then Issues_AddFlag sFlags,"CRIT:NO_FROM:Не заполнено Откуда."
    If sReturnable<>Issues_Canon(WMSISS_YES) And sReturnable<>Issues_Canon(WMSISS_NO) Then Issues_AddFlag sFlags,"CRIT:RETURNABLE_BAD:Выберите Возвратный = Да или Нет."
    If Not bRet Then ret=0
    If ret<0 Then Issues_AddFlag sFlags,"CRIT:RETURN_NEGATIVE:Возвращено не может быть отрицательным."
    If bQty And qty>0 And ret>qty+WMSISS_EPS Then Issues_AddFlag sFlags,"CRIT:RETURN_GT_ISSUED:Возвращено больше выданного."
    If posted<0 Then Issues_AddFlag sFlags,"CRIT:POSTED_RETURN_NEGATIVE:Системное проведённое количество возврата отрицательное."
    If bQty And posted>qty+WMSISS_EPS Then Issues_AddFlag sFlags,"CRIT:POSTED_GT_ISSUED:В истории проведено возвратов больше, чем выдано."
    If ret+WMSISS_EPS<posted Then Issues_AddFlag sFlags,"CRIT:RETURN_BELOW_POSTED:Нельзя уменьшить Возвращено ниже уже проведённого возврата " & Issues_NumText(posted) & ". Историю нужно отменять безопасной операцией."
    If sReturnable=Issues_Canon(WMSISS_NO) Then
        If ret>WMSISS_EPS Then Issues_AddFlag sFlags,"CRIT:NONRETURNABLE_HAS_RETURN:Для невозвратной выдачи Возвращено должно быть 0/пусто."
        If dReturn>0 Then Issues_AddFlag sFlags,"WARN:NONRETURNABLE_RETURN_DATE:Для невозвратной выдачи указана дата возврата."
    ElseIf sReturnable=Issues_Canon(WMSISS_YES) Then
        If ret>WMSISS_EPS And dReturn<=0 Then Issues_AddFlag sFlags,"CRIT:RETURN_DATE_REQUIRED:При Возвращено > 0 обязательна Дата возврата."
        If ret<=WMSISS_EPS And dReturn>0 Then Issues_AddFlag sFlags,"WARN:RETURN_DATE_WITH_ZERO:Дата возврата заполнена, но Возвращено = 0."
        If ret>WMSISS_EPS And dReturn>0 And dIssue>0 And dReturn<dIssue-WMSISS_EPS Then Issues_AddFlag sFlags,"CRIT:RETURN_BEFORE_ISSUE:Дата возврата не может быть раньше даты выдачи."
    End If
    ' Exact code resolution and non-destructive hints.
    sugCode="":sugName="":score=0:sugSource="":sRetHint="":sUnitHint="":sPlaceHint=""
    If Issues_LoadNomCache(oDoc) Then
        If sCode<>"" Then
            idx=Issues_NomIndexByCode(sCode)
            If idx>=0 Then
                sugCode=gWMSISS_NomCode(idx):sugName=gWMSISS_NomName(idx):score=100:sugSource="EXACT_CODE":sRetHint=gWMSISS_NomReturnable(idx):sUnitHint=gWMSISS_NomUnit(idx):sPlaceHint=gWMSISS_NomPlace(idx)
                If Not gWMSISS_NomActive(idx) Then Issues_AddFlag sFlags,"WARN:CODE_INACTIVE:Код найден в номенклатуре, но позиция не активна."
                If sName<>"" And sugName<>"" And Issues_NameSimilarity(sName,sugName)<45 Then Issues_AddFlag sFlags,"WARN:CODE_NAME_MISMATCH:Код относится к '" & sugName & "'. Проверьте наименование."
                If sUnit<>"" And sUnitHint<>"" And Issues_CanonUnit(sUnit)<>Issues_CanonUnit(sUnitHint) Then Issues_AddFlag sFlags,"WARN:UNIT_HINT:В номенклатуре базовая ед. '" & sUnitHint & "'. Проверьте единицу выдачи."
                If sFrom<>"" And sPlaceHint<>"" And Issues_Canon(sFrom)<>Issues_Canon(sPlaceHint) Then Issues_AddFlag sFlags,"INFO:PLACE_HINT:В номенклатуре основное место '" & sPlaceHint & "'."
                If sReturnable<>"" And sRetHint<>"" And Issues_Canon(sRetHint)<>sReturnable Then Issues_AddFlag sFlags,"WARN:RETURNABLE_HINT:В номенклатуре Возвратный = '" & sRetHint & "'. Проверьте выбор."
            Else
                Issues_AddFlag sFlags,"WARN:CODE_UNKNOWN:Код пока не найден в номенклатуре."
            End If
        ElseIf sName<>"" And bAllowFuzzy Then
            Issues_BestNameSuggestion sName,sugCode,sugName,score,sugSource,sRetHint,sUnitHint,sPlaceHint
            If sugCode<>"" And score>=55 Then
                Issues_AddFlag sFlags,"INFO:CODE_SUGGESTION:Похожая позиция: " & sugCode & " — " & sugName & ". Это только подсказка."
            Else
                Issues_AddFlag sFlags,"WARN:NO_CODE:Код не заполнен. Для старых позиций допустимо, но идентификация будет слабее."
            End If
        ElseIf sCode="" Then
            Issues_AddFlag sFlags,"WARN:NO_CODE:Код не заполнен. Для старых позиций допустимо, но идентификация будет слабее."
        End If
    End If
    baseFP=Issues_BaseFingerprint(oSh,r):retFP=Issues_ReturnFingerprint(oSh,r):oldBase=Trim(oSh.getCellByPosition(cBase,r).String)
    If Issues_Canon(sState)=Issues_Canon(WMSISS_STATE_DONE) And oldBase<>"" And oldBase<>baseFP Then Issues_AddFlag sFlags,"CRIT:POSTED_CHANGED:Проведённая выдача изменена. Сначала отмените старое движение и перепроведите безопасно."
    sSeverity=Issues_SeverityFromFlags(sFlags):sError=Issues_FirstProblem(sFlags):sRetError=Issues_FirstReturnProblem(sFlags)
    oSh.getCellByPosition(cErr,r).String=sError:oSh.getCellByPosition(cRetErr,r).String=sRetError:oSh.getCellByPosition(cFlags,r).String=sFlags:oSh.getCellByPosition(cSev,r).String=sSeverity:oSh.getCellByPosition(cValidated,r).String=Issues_Stamp(Now)
    oSh.getCellByPosition(cCodeSug,r).String=sugCode:oSh.getCellByPosition(cSugName,r).String=sugName:oSh.getCellByPosition(cSugScore,r).Value=score:oSh.getCellByPosition(cSugSource,r).String=sugSource:oSh.getCellByPosition(cEmp,r).String=Issues_Canon(sEmp)
    oSh.getCellByPosition(cRetHint,r).String=sRetHint:oSh.getCellByPosition(cUnitHint,r).String=sUnitHint:oSh.getCellByPosition(cPlaceHint,r).String=sPlaceHint
    oSh.getCellByPosition(cRetState,r).String=Issues_ReturnStateText(sReturnable,qty,ret,sFlags)
    If oldBase="" Or Issues_Canon(sState)<>Issues_Canon(WMSISS_STATE_DONE) Then oSh.getCellByPosition(cBase,r).String=baseFP
    oSh.getCellByPosition(cRetFp,r).String=retFP
    nVer=Issues_CellNumber(oSh.getCellByPosition(cVer,r)):If bMarkDirty Then nVer=nVer+1:oSh.getCellByPosition(cVer,r).Value=nVer
    If bMarkDirty Then oSh.getCellByPosition(cDirty,r).String="1":oSh.getCellByPosition(cTouch,r).String=Issues_Stamp(Now)
    qMode="ISSUE"
    If ret>WMSISS_EPS Or dReturn>0 Then qMode="ISSUE_RETURN"
    oSh.getCellByPosition(cLastQ,r).String=qMode
    If bMarkDirty And Not gWMSISS_BulkMode Then Issues_QueueRow oDoc,oSh,r,qMode,sError
    If Not gWMSISS_BulkMode Then Issues_WriteRowAnnotations oSh,r,sError,sugCode,sugName,score
End Sub

Function Issues_ReturnStateText(sReturnable As String,qty As Double,ret As Double,sFlags As String) As String
    If InStr(1,sFlags,"CRIT:",1)>0 Then Issues_ReturnStateText="Ошибка":Exit Function
    If sReturnable=Issues_Canon(WMSISS_NO) Then Issues_ReturnStateText="Невозвратный":Exit Function
    If sReturnable<>Issues_Canon(WMSISS_YES) Then Issues_ReturnStateText="Не определено":Exit Function
    If ret<=WMSISS_EPS Then Issues_ReturnStateText="Открыт":Exit Function
    If qty>0 And ret>=qty-WMSISS_EPS Then
        Issues_ReturnStateText="Возвращено полностью"
    Else
        Issues_ReturnStateText="Частичный возврат"
    End If
End Function

Sub Issues_AddFlag(ByRef sFlags As String,sFlag As String)
    If sFlags<>"" Then sFlags=sFlags & "|"
    sFlags=sFlags & sFlag
End Sub

Function Issues_SeverityFromFlags(sFlags As String) As String
    If InStr(1,sFlags,"CRIT:",1)>0 Then Issues_SeverityFromFlags="CRITICAL":Exit Function
    If InStr(1,sFlags,"WARN:",1)>0 Then Issues_SeverityFromFlags="WARNING":Exit Function
    If InStr(1,sFlags,"INFO:",1)>0 Then Issues_SeverityFromFlags="INFO":Exit Function
    Issues_SeverityFromFlags="OK"
End Function

Function Issues_FirstProblem(sFlags As String) As String
    Dim a As Variant,i As Long,p As Long,s As String,pass As Long,prefix As String
    If Trim(sFlags)="" Then Exit Function
    a=Split(sFlags,"|")
    For pass=1 To 2
        If pass=1 Then
            prefix="CRIT:"
        Else
            prefix="WARN:"
        End If
        For i=LBound(a) To UBound(a)
            s=CStr(a(i))
            If Left(s,5)=prefix Then
                p=InStr(6,s,":")
                If p>0 Then
                    Issues_FirstProblem=Mid(s,p+1)
                Else
                    Issues_FirstProblem=s
                End If
                Exit Function
            End If
        Next i
    Next pass
End Function

Function Issues_FirstReturnProblem(sFlags As String) As String
    Dim a As Variant,i As Long,s As String,p As Long,pass As Long,prefix As String
    If Trim(sFlags)="" Then Exit Function
    a=Split(sFlags,"|")
    For pass=1 To 3
        If pass=1 Then
            prefix="CRIT:"
        ElseIf pass=2 Then
            prefix="WARN:"
        Else
            prefix="INFO:"
        End If
        For i=LBound(a) To UBound(a)
            s=CStr(a(i))
            If Left(s,5)=prefix And (InStr(1,s,"RETURN",1)>0 Or InStr(1,s,"POSTED",1)>0) Then
                p=InStr(6,s,":")
                If p>0 Then
                    Issues_FirstReturnProblem=Mid(s,p+1)
                Else
                    Issues_FirstReturnProblem=s
                End If
                Exit Function
            End If
        Next i
    Next pass
End Function

' ============================================================================
' NOMENCLATURE CACHE / FUZZY SUGGESTION ONLY
' ============================================================================

Sub Issues_ResetNomCache()
    gWMSISS_NomReady=False:gWMSISS_NomCount=0
End Sub

Function Issues_LoadNomCache(oDoc As Object) As Boolean
    Dim oSh As Object,last As Long,r As Long,cCode As Long,cName As Long,cUnit As Long,cPlace As Long,cRet As Long,cAct As Long,n As Long,s As String
    If gWMSISS_NomReady Then Issues_LoadNomCache=(gWMSISS_NomCount>0):Exit Function
    gWMSISS_NomReady=True:gWMSISS_NomCount=0
    If Not oDoc.Sheets.hasByName(WMSISS_NOM) Then Exit Function
    oSh=oDoc.Sheets.getByName(WMSISS_NOM):cCode=Issues_FindHeader(oSh,"Код"):cName=Issues_FindHeader(oSh,"Наименование"):cUnit=Issues_FindHeader(oSh,"Базовая ед."):cPlace=Issues_FindHeader(oSh,"Место"):cRet=Issues_FindHeader(oSh,"Возвратный"):cAct=Issues_FindHeader(oSh,"Активен")
    If cCode<0 Or cName<0 Then Exit Function
    last=Issues_LastContentRow(oSh,9):If last<1 Then Exit Function
    ReDim gWMSISS_NomCode(0 To last-1):ReDim gWMSISS_NomName(0 To last-1):ReDim gWMSISS_NomUnit(0 To last-1):ReDim gWMSISS_NomPlace(0 To last-1):ReDim gWMSISS_NomReturnable(0 To last-1):ReDim gWMSISS_NomActive(0 To last-1)
    For r=1 To last
        s=Trim(oSh.getCellByPosition(cCode,r).String)
        If s<>"" Or Trim(oSh.getCellByPosition(cName,r).String)<>"" Then
            gWMSISS_NomCode(n)=s:gWMSISS_NomName(n)=Trim(oSh.getCellByPosition(cName,r).String)
            If cUnit>=0 Then gWMSISS_NomUnit(n)=Trim(oSh.getCellByPosition(cUnit,r).String)
            If cPlace>=0 Then gWMSISS_NomPlace(n)=Trim(oSh.getCellByPosition(cPlace,r).String)
            If cRet>=0 Then gWMSISS_NomReturnable(n)=Trim(oSh.getCellByPosition(cRet,r).String)
            gWMSISS_NomActive(n)=True
            If cAct>=0 Then If Issues_Canon(oSh.getCellByPosition(cAct,r).String)="нет" Or Issues_Canon(oSh.getCellByPosition(cAct,r).String)="false" Or Issues_Canon(oSh.getCellByPosition(cAct,r).String)="0" Then gWMSISS_NomActive(n)=False
            n=n+1
        End If
    Next r
    gWMSISS_NomCount=n:Issues_LoadNomCache=(n>0)
End Function

Function Issues_NomIndexByCode(ByVal sCode As String) As Long
    Dim i As Long,c As String:Issues_NomIndexByCode=-1:c=Issues_Canon(sCode)
    For i=0 To gWMSISS_NomCount-1:If Issues_Canon(gWMSISS_NomCode(i))=c Then Issues_NomIndexByCode=i:Exit Function
    Next i
End Function

Sub Issues_BestNameSuggestion(sName As String,ByRef sCode As String,ByRef sBestName As String,ByRef score As Double,ByRef sSource As String,ByRef sRet As String,ByRef sUnit As String,ByRef sPlace As String)
    Dim i As Long,d As Double
    sCode="":sBestName="":score=0:sSource="":sRet="":sUnit="":sPlace=""
    For i=0 To gWMSISS_NomCount-1
        If gWMSISS_NomActive(i) Then
            d=Issues_NameSimilarity(sName,gWMSISS_NomName(i))
            If d>score Then score=d:sCode=gWMSISS_NomCode(i):sBestName=gWMSISS_NomName(i):sRet=gWMSISS_NomReturnable(i):sUnit=gWMSISS_NomUnit(i):sPlace=gWMSISS_NomPlace(i)
        End If
    Next i
    If score>=99 Then
        sSource="EXACT_NAME"
    ElseIf score>=55 Then
        sSource="FUZZY_NAME"
    Else
        sCode="":sBestName="":sRet="":sUnit="":sPlace="":sSource=""
    End If
End Sub

Function Issues_NameSimilarity(ByVal a As String,ByVal b As String) As Double
    Dim x As String,y As String,ax As Variant,ay As Variant,i As Long,j As Long,hit As Long,nx As Long,ny As Long,t As String,prefix As Double
    x=Issues_CanonName(a):y=Issues_CanonName(b)
    If x="" Or y="" Then Exit Function
    If x=y Then Issues_NameSimilarity=100:Exit Function
    If InStr(1,x,y,1)>0 Or InStr(1,y,x,1)>0 Then Issues_NameSimilarity=85:Exit Function
    ax=Split(x," "):ay=Split(y," ")
    For i=LBound(ax) To UBound(ax)
        t=Trim(CStr(ax(i))):If Len(t)>=3 Then nx=nx+1:For j=LBound(ay) To UBound(ay):If Issues_Canon(t)=Issues_Canon(CStr(ay(j))) Then hit=hit+1:Exit For
        Next j
    Next i
    For j=LBound(ay) To UBound(ay):If Len(Trim(CStr(ay(j))))>=3 Then ny=ny+1
    Next j
    If Left(x,4)=Left(y,4) Then prefix=10
    If nx>0 And ny>0 Then Issues_NameSimilarity=(hit/((nx+ny)/2))*75+prefix
    If Issues_NameSimilarity>100 Then Issues_NameSimilarity=100
End Function

Function Issues_CanonName(ByVal s As String) As String
    s=LCase(Trim(s)):s=Replace(s,"ё","е"):s=Replace(s,Chr(160)," ")
    s=Replace(s,"/"," "):s=Replace(s,"-"," "):s=Replace(s,"_"," "):s=Replace(s,","," "):s=Replace(s,"."," "):s=Replace(s,"("," "):s=Replace(s,")"," ")
    Do While InStr(s,"  ")>0:s=Replace(s,"  "," "):Loop
    Issues_CanonName=Trim(s)
End Function

Function Issues_CanonUnit(ByVal s As String) As String
    s=Issues_Canon(s):s=Replace(s,".","")
    Select Case s
        Case "штук","штука","штуки","шт":Issues_CanonUnit="шт"
        Case "метр","метры","метров","м":Issues_CanonUnit="м"
        Case "сантиметр","сантиметры","см":Issues_CanonUnit="см"
        Case "миллиметр","миллиметры","мм":Issues_CanonUnit="мм"
        Case "килограмм","килограммы","кг":Issues_CanonUnit="кг"
        Case "грамм","граммы","г":Issues_CanonUnit="г"
        Case "литр","литры","л":Issues_CanonUnit="л"
        Case "миллилитр","миллилитры","мл":Issues_CanonUnit="мл"
        Case Else:Issues_CanonUnit=s
    End Select
End Function

' ============================================================================
' ANNOTATIONS / VISUAL STATE
' ============================================================================

Sub Issues_WriteRowAnnotations(oSh As Object,r As Long,sProblem As String,sCode As String,sName As String,score As Double)
    Dim msg As String
    Issues_SetWMSAnnotation oSh.getCellByPosition(0,r),IIf(sProblem<>"","WMS: " & sProblem,"")
    If Trim(oSh.getCellByPosition(1,r).String)="" And sCode<>"" And score>=55 Then
        msg="WMS: Подсказка, не автослияние: " & sCode & " — " & sName & " (" & CStr(CLng(score)) & "%). Код решаете вы."
        Issues_SetWMSAnnotation oSh.getCellByPosition(1,r),msg
    Else
        Issues_ClearWMSAnnotation oSh.getCellByPosition(1,r)
    End If
End Sub

Sub Issues_SetWMSAnnotation(oCell As Object,sText As String)
    Dim oSh As Object,oNotes As Object,cur As String,pos As Variant
    On Error GoTo Done
    If sText="" Then Issues_ClearWMSAnnotation oCell:Exit Sub
    cur=Trim(oCell.Annotation.String)
    ' Preserve a note written manually by the user. WMS owns only notes prefixed WMS:.
    If cur<>"" And Left(cur,4)<>"WMS:" Then Exit Sub
    oSh=oCell.Spreadsheet
    If cur<>"" Then Issues_RemoveAnnotationAt oSh,oCell.CellAddress
    oNotes=oSh.Annotations
    pos=oCell.CellAddress
    oNotes.insertNew(pos,sText)
Done:
End Sub

Sub Issues_ClearWMSAnnotation(oCell As Object)
    Dim cur As String,oSh As Object
    On Error GoTo Done
    cur=Trim(oCell.Annotation.String)
    If Left(cur,4)<>"WMS:" Then Exit Sub
    oSh=oCell.Spreadsheet
    Issues_RemoveAnnotationAt oSh,oCell.CellAddress
Done:
End Sub

Sub Issues_RemoveAnnotationAt(oSh As Object,addr As Variant)
    Dim oNotes As Object,a As Object,i As Long,p As Variant
    On Error GoTo Done
    oNotes=oSh.Annotations
    For i=oNotes.Count-1 To 0 Step -1
        a=oNotes.getByIndex(i)
        p=a.Position
        If p.Column=addr.Column And p.Row=addr.Row Then
            oNotes.removeByIndex(i)
            Exit For
        End If
    Next i
Done:
End Sub

Sub Issues_ApplyRowVisual(oSh As Object,r As Long)
    Dim flags As String,sev As String,retState As String
    flags=Issues_CellText(oSh,r,WMSISS_X_FLAGS):sev=Issues_CellText(oSh,r,WMSISS_X_SEVERITY):retState=Issues_CellText(oSh,r,WMSISS_X_RETSTATE)
    oSh.getCellRangeByPosition(0,r,WMSISS_USER_LAST_COL,r).CellBackColor=RGB(255,255,255)
    oSh.getCellRangeByPosition(0,r,WMSISS_USER_LAST_COL,r).CharColor=RGB(31,41,55)
    oSh.getCellRangeByPosition(0,r,WMSISS_USER_LAST_COL,r).CharWeight=com.sun.star.awt.FontWeight.NORMAL
    oSh.getCellByPosition(5,r).CellBackColor=RGB(255,251,235)
    oSh.getCellRangeByPosition(8,r,10,r).CellBackColor=RGB(239,246,255)
    oSh.getCellByPosition(11,r).CellBackColor=RGB(248,250,252)
    Select Case retState
        Case "Возвращено полностью":oSh.getCellRangeByPosition(8,r,10,r).CellBackColor=RGB(187,247,208)
        Case "Частичный возврат":oSh.getCellRangeByPosition(8,r,10,r).CellBackColor=RGB(254,240,138)
        Case "Невозвратный":oSh.getCellRangeByPosition(8,r,10,r).CellBackColor=RGB(226,232,240)
    End Select
    If InStr(flags,"NO_NUMBER")>0 Then Issues_WarnCell oSh.getCellByPosition(0,r)
    If InStr(flags,"NO_CODE")>0 Or InStr(flags,"CODE_UNKNOWN")>0 Then Issues_WarnCell oSh.getCellByPosition(1,r)
    If InStr(flags,"CODE_SUGGESTION")>0 Then oSh.getCellByPosition(1,r).CellBackColor=RGB(191,219,254)
    If InStr(flags,"CODE_NAME_MISMATCH")>0 Then Issues_WarnCell oSh.getCellByPosition(1,r):Issues_WarnCell oSh.getCellByPosition(2,r)
    If InStr(flags,"NO_NAME")>0 Then Issues_ErrorCell oSh.getCellByPosition(2,r)
    If InStr(flags,"QTY_BAD")>0 Then Issues_ErrorCell oSh.getCellByPosition(3,r)
    If InStr(flags,"NO_UNIT")>0 Then Issues_ErrorCell oSh.getCellByPosition(4,r)
    If InStr(flags,"UNIT_HINT")>0 Then Issues_WarnCell oSh.getCellByPosition(4,r)
    If InStr(flags,"NO_EMPLOYEE")>0 Then Issues_ErrorCell oSh.getCellByPosition(5,r)
    If InStr(flags,"ISSUE_DATE_BAD")>0 Then
        Issues_ErrorCell oSh.getCellByPosition(6,r)
    ElseIf InStr(flags,"ISSUE_DATE_FUTURE")>0 Then
        Issues_WarnCell oSh.getCellByPosition(6,r)
    End If
    If InStr(flags,"NO_FROM")>0 Then Issues_ErrorCell oSh.getCellByPosition(7,r)
    If InStr(flags,"PLACE_HINT")>0 Then Issues_InfoCell oSh.getCellByPosition(7,r)
    If InStr(flags,"RETURNABLE_BAD")>0 Then
        Issues_ErrorCell oSh.getCellByPosition(8,r)
    ElseIf InStr(flags,"RETURNABLE_HINT")>0 Then
        Issues_WarnCell oSh.getCellByPosition(8,r)
    End If
    If InStr(flags,"RETURN_NEGATIVE")>0 Or InStr(flags,"RETURN_GT_ISSUED")>0 Or InStr(flags,"RETURN_BELOW_POSTED")>0 Or InStr(flags,"NONRETURNABLE_HAS_RETURN")>0 Then Issues_ErrorCell oSh.getCellByPosition(9,r)
    If InStr(flags,"RETURN_DATE_REQUIRED")>0 Or InStr(flags,"RETURN_BEFORE_ISSUE")>0 Then Issues_ErrorCell oSh.getCellByPosition(10,r)
    If InStr(flags,"RETURN_DATE_WITH_ZERO")>0 Or InStr(flags,"NONRETURNABLE_RETURN_DATE")>0 Then Issues_WarnCell oSh.getCellByPosition(10,r)
    If InStr(flags,"POSTED_CHANGED")>0 Then oSh.getCellRangeByPosition(0,r,8,r).CellBackColor=RGB(254,202,202)
End Sub

Sub Issues_ErrorCell(oCell As Object)
    oCell.CellBackColor=RGB(254,202,202):oCell.CharColor=RGB(127,29,29):oCell.CharWeight=com.sun.star.awt.FontWeight.BOLD
End Sub
Sub Issues_WarnCell(oCell As Object)
    oCell.CellBackColor=RGB(254,240,138):oCell.CharColor=RGB(113,63,18)
End Sub
Sub Issues_InfoCell(oCell As Object)
    oCell.CellBackColor=RGB(191,219,254):oCell.CharColor=RGB(30,64,175)
End Sub

Function Issues_RowDiagnosticText(oSh As Object,r As Long) As String
    Dim s As String
    s="SourceID: " & Issues_CellText(oSh,r,WMSISS_T_ID) & Chr(10) & _
      "Состояние: " & Issues_CellText(oSh,r,WMSISS_T_STATE) & Chr(10) & _
      "Режим: " & Issues_CellText(oSh,r,WMSISS_T_MODE) & Chr(10) & _
      "Контроль: " & Issues_CellText(oSh,r,WMSISS_X_SEVERITY) & Chr(10) & _
      "Возврат: " & Issues_CellText(oSh,r,WMSISS_X_RETSTATE) & Chr(10) & _
      "Проведено возврата: " & Issues_CellText(oSh,r,WMSISS_T_RETPOSTED) & Chr(10) & _
      "Ошибка: " & Issues_CellText(oSh,r,WMSISS_T_ERROR)
    If Issues_CellText(oSh,r,WMSISS_X_CODESUG)<>"" Then s=s & Chr(10) & "Подсказка кода: " & Issues_CellText(oSh,r,WMSISS_X_CODESUG) & " — " & Issues_CellText(oSh,r,WMSISS_X_SUGNAME) & " (только предложение)"
    Issues_RowDiagnosticText=s
End Function

' ============================================================================
' QUEUE — DEDUP BY SOURCEID
' ============================================================================

Sub Issues_ResetQueueCache()
    gWMSISS_QueueReady=False:gWMSISS_QueueNextRow=1:gWMSISS_QCap=0:gWMSISS_QCount=0
End Sub

Sub Issues_QueueRow(oDoc As Object,oSh As Object,r As Long,sMode As String,sMessage As String)
    Dim oQ As Object,sID As String,qrow As Long,stamp As String
    On Error GoTo EH
    If Not oDoc.Sheets.hasByName(WMSISS_QUEUE) Then Exit Sub
    sID=Issues_CellText(oSh,r,WMSISS_T_ID):If sID="" Then Exit Sub
    oQ=oDoc.Sheets.getByName(WMSISS_QUEUE):Issues_EnsureQueueHeaders oQ
    If Not Issues_EnsureQueueCache(oDoc) Then Exit Sub
    qrow=Issues_QueueHashGet(sID)
    stamp=Issues_Stamp(Now)
    If qrow<1 Then qrow=gWMSISS_QueueNextRow:gWMSISS_QueueNextRow=qrow+1:Issues_QueueHashPut sID,qrow
    oQ.getCellByPosition(0,qrow).String=stamp:oQ.getCellByPosition(1,qrow).String=WMSISS_SHEET:oQ.getCellByPosition(2,qrow).String=sID:oQ.getCellByPosition(3,qrow).Value=r+1:oQ.getCellByPosition(4,qrow).String=sMode:oQ.getCellByPosition(5,qrow).String="PENDING":oQ.getCellByPosition(6,qrow).String="":oQ.getCellByPosition(7,qrow).String=sMessage
    oQ.IsVisible=False
    Exit Sub
EH:
    Issues_ResetQueueCache:Issues_LogError oDoc,"QueueRow","Не удалось поставить выдачу в очередь: " & CStr(Err) & " " & Error$
End Sub

Sub Issues_CancelPendingQueueBySourceID(oDoc As Object,sSourceID As String,sReason As String)
    Dim oQ As Object,last As Long,r As Long,sID As String,status As String
    If Trim(sSourceID)="" Then Exit Sub
    If Not oDoc.Sheets.hasByName(WMSISS_QUEUE) Then Exit Sub
    oQ=oDoc.Sheets.getByName(WMSISS_QUEUE)
    Issues_EnsureQueueHeaders oQ
    last=Issues_LastContentRow(oQ,7)
    For r=1 To last
        sID=Trim(oQ.getCellByPosition(2,r).String)
        If sID=sSourceID Then
            status=UCase(Trim(oQ.getCellByPosition(5,r).String))
            If status="PENDING" Or status="ERROR" Or status="RETRY" Then
                oQ.getCellByPosition(5,r).String="CANCELLED"
                oQ.getCellByPosition(7,r).String=sReason
            End If
        End If
    Next r
End Sub

Sub Issues_EnsureQueueHeaders(oQ As Object)
    Dim a As Variant,i As Long:a=Array("Время","Лист","SourceID","RowHint","Mode","Status","RunID","Message")
    For i=0 To UBound(a):If Trim(oQ.getCellByPosition(i,0).String)="" Then oQ.getCellByPosition(i,0).String=CStr(a(i))
    Next i
End Sub

Function Issues_EnsureQueueCache(oDoc As Object) As Boolean
    Dim oQ As Object,last As Long,r As Long,sID As String,status As String,cap As Long
    If gWMSISS_QueueReady Then Issues_EnsureQueueCache=True:Exit Function
    If Not oDoc.Sheets.hasByName(WMSISS_QUEUE) Then Exit Function
    oQ=oDoc.Sheets.getByName(WMSISS_QUEUE):last=Issues_LastContentRow(oQ,7)
    cap=8192
    Do While cap<(last+WMSISS_MAX_EVENT_ROWS+100)*2
        cap=cap*2
    Loop
    gWMSISS_QCap=cap:ReDim gWMSISS_QKeys(0 To cap-1):ReDim gWMSISS_QRows(0 To cap-1):gWMSISS_QCount=0
    For r=1 To last
        sID=Trim(oQ.getCellByPosition(2,r).String):status=UCase(Trim(oQ.getCellByPosition(5,r).String))
        If sID<>"" And (status="PENDING" Or status="ERROR" Or status="RETRY") Then Issues_QueueHashPut sID,r
    Next r
    gWMSISS_QueueNextRow=last+1:If gWMSISS_QueueNextRow<1 Then gWMSISS_QueueNextRow=1
    gWMSISS_QueueReady=True:Issues_EnsureQueueCache=True
End Function

Function Issues_QHash(sKey As String) As Long
    Dim i As Long,h As Long,ch As Long:h=0
    For i=1 To Len(sKey)
        ch=Asc(Mid(sKey,i,1))
        If ch<0 Then ch=ch+256
        h=(h*31+ch) Mod 2147483
    Next i
    Issues_QHash=h
End Function

Function Issues_QueueHashSlot(sKey As String,ByRef found As Boolean) As Long
    Dim slot As Long
    If gWMSISS_QCap<=0 Then Issues_QueueHashSlot=-1:Exit Function
    slot=Issues_QHash(sKey) Mod gWMSISS_QCap
    Do
        If gWMSISS_QKeys(slot)="" Then found=False:Issues_QueueHashSlot=slot:Exit Function
        If gWMSISS_QKeys(slot)=sKey Then found=True:Issues_QueueHashSlot=slot:Exit Function
        slot=slot+1:If slot>=gWMSISS_QCap Then slot=0
    Loop
End Function

Sub Issues_QueueHashPut(sKey As String,r As Long)
    Dim found As Boolean,slot As Long:If sKey="" Or gWMSISS_QCap<=0 Then Exit Sub
    slot=Issues_QueueHashSlot(sKey,found):If slot<0 Then Exit Sub
    If Not found Then gWMSISS_QKeys(slot)=sKey:gWMSISS_QCount=gWMSISS_QCount+1
    gWMSISS_QRows(slot)=r
End Sub

Function Issues_QueueHashGet(sKey As String) As Long
    Dim found As Boolean,slot As Long:slot=Issues_QueueHashSlot(sKey,found):If slot>=0 And found Then Issues_QueueHashGet=gWMSISS_QRows(slot)
End Function


Sub Issues_QueueRowsBatch(oDoc As Object,oSh As Object,r1 As Long,r2 As Long)
    Dim oQ As Object,lastQ As Long,nChanged As Long,maxRows As Long,total As Long,i As Long,r As Long,idx As Long
    Dim qData As Variant,srcData As Variant,modeData As Variant,errData As Variant
    Dim out() As Variant,keys() As String,rows() As Long,cap As Long,slot As Long,found As Boolean
    Dim sID As String,sMode As String,sMsg As String,status As String,stamp As String
    On Error GoTo EH
    If Not oDoc.Sheets.hasByName(WMSISS_QUEUE) Then Exit Sub
    If r2<r1 Then Exit Sub
    oQ=oDoc.Sheets.getByName(WMSISS_QUEUE):Issues_EnsureQueueHeaders oQ
    lastQ=Issues_LastContentRow(oQ,7)
    nChanged=r2-r1+1
    maxRows=lastQ+nChanged+4
    If maxRows<4 Then maxRows=4
    ReDim out(0 To maxRows-1)
    cap=16
    Do While cap<(maxRows+16)*2:cap=cap*2:Loop
    ReDim keys(0 To cap-1):ReDim rows(0 To cap-1)
    total=0
    If lastQ>=1 Then
        qData=oQ.getCellRangeByPosition(0,1,7,lastQ).getDataArray()
        For i=0 To UBound(qData)
            out(total)=qData(i)
            sID=Trim(CStr(qData(i)(2))):status=UCase(Trim(CStr(qData(i)(5))))
            If sID<>"" And (status="PENDING" Or status="ERROR" Or status="RETRY") Then
                slot=Issues_BatchHashSlot(keys,cap,sID,found)
                If Not found Then keys(slot)=sID
                rows(slot)=total
            End If
            total=total+1
        Next i
    End If
    srcData=oSh.getCellRangeByPosition(Issues_FindHeader(oSh,WMSISS_T_ID),r1,Issues_FindHeader(oSh,WMSISS_T_ID),r2).getDataArray()
    modeData=oSh.getCellRangeByPosition(Issues_FindHeader(oSh,WMSISS_X_LASTQMODE),r1,Issues_FindHeader(oSh,WMSISS_X_LASTQMODE),r2).getDataArray()
    errData=oSh.getCellRangeByPosition(Issues_FindHeader(oSh,WMSISS_T_ERROR),r1,Issues_FindHeader(oSh,WMSISS_T_ERROR),r2).getDataArray()
    stamp=Issues_Stamp(Now)
    For i=0 To nChanged-1
        sID=Trim(CStr(srcData(i)(0)))
        If sID<>"" Then
            sMode=Trim(CStr(modeData(i)(0))):If sMode="" Then sMode="ISSUE"
            sMsg=Trim(CStr(errData(i)(0)))
            slot=Issues_BatchHashSlot(keys,cap,sID,found)
            If found Then
                idx=rows(slot)
            Else
                idx=total:total=total+1:keys(slot)=sID:rows(slot)=idx
            End If
            out(idx)=Array(stamp,WMSISS_SHEET,sID,CDbl(r1+i+1),sMode,"PENDING","",sMsg)
        End If
    Next i
    If total>0 Then
        ReDim Preserve out(0 To total-1)
        oQ.getCellRangeByPosition(0,1,7,total).setDataArray(out())
    End If
    oQ.IsVisible=False
    Issues_ResetQueueCache
    Exit Sub
EH:
    gWMSISS_BulkMode=False
    Issues_ResetQueueCache
    Issues_LogError oDoc,"QueueRowsBatch","Не удалось пакетно поставить выдачи в очередь: " & CStr(Err) & " " & Error$
End Sub

Function Issues_BatchHashSlot(ByRef keys() As String,cap As Long,sKey As String,ByRef bFound As Boolean) As Long
    Dim h As Long,i As Long,ch As Long,slot As Long
    h=0
    For i=1 To Len(sKey)
        ch=Asc(Mid(sKey,i,1)):If ch<0 Then ch=ch+256
        h=(h*31+ch) Mod 2147483
    Next i
    slot=h Mod cap
    Do
        If keys(slot)="" Then bFound=False:Issues_BatchHashSlot=slot:Exit Function
        If keys(slot)=sKey Then bFound=True:Issues_BatchHashSlot=slot:Exit Function
        slot=slot+1:If slot>=cap Then slot=0
    Loop
End Function

' ============================================================================
' REVALIDATE / API FOR FUTURE MOVEMENT + NONRETURNS
' ============================================================================

Function Issues_RevalidateAllCore(oDoc As Object,bMarkDirty As Boolean,ByRef sReport As String) As Boolean
    Dim oSh As Object,last As Long,r As Long,n As Long,bLocked As Boolean
    Issues_RevalidateAllCore=False:sReport=""
    On Error GoTo EH
    If Not Issues_Preflight(oDoc,sReport) Then Exit Function
    oSh=oDoc.Sheets.getByName(WMSISS_SHEET):last=Issues_LastContentRow(oSh,WMSISS_USER_LAST_COL)
    gWMSISS_Busy=True
    On Error Resume Next
    oDoc.lockControllers
    If Err=0 Then
        bLocked=True
    Else
        Err=0
    End If
    On Error GoTo EH
    Issues_ResetNomCache
    For r=1 To last
        If Issues_RowHasBusinessData(oSh,r) Then Issues_EnsureSIDUniquenessForEvent oSh,r,r:Issues_ValidateRow oDoc,oSh,r,bMarkDirty,True:Issues_ApplyRowVisual oSh,r:n=n+1
    Next r
    Issues_RebuildHelperLists oDoc,oSh:Issues_ApplyValidations oDoc,oSh:Issues_RefreshFilter oDoc,oSh
    sReport="Проверено строк Выдачи: " & CStr(n) & ". Ошибки/возвраты/подсказки/очередь обновлены.":Issues_RevalidateAllCore=True
Done:
    If bLocked Then On Error Resume Next:oDoc.unlockControllers:On Error GoTo 0
    gWMSISS_Busy=False:Exit Function
EH:
    sReport="Revalidate Выдачи: " & CStr(Err) & " " & Error$:Resume Done
End Function

Function Issues_API_FindRowBySourceID(sSourceID As String) As Long
    Dim oSh As Object,c As Long,r As Long,last As Long:Issues_API_FindRowBySourceID=-1
    If Not ThisComponent.Sheets.hasByName(WMSISS_SHEET) Then Exit Function
    oSh=ThisComponent.Sheets.getByName(WMSISS_SHEET):c=Issues_FindHeader(oSh,WMSISS_T_ID):If c<0 Then Exit Function:last=Issues_LastContentRow(oSh,WMSISS_USER_LAST_COL)
    For r=1 To last:If Trim(oSh.getCellByPosition(c,r).String)=Trim(sSourceID) Then Issues_API_FindRowBySourceID=r:Exit Function
    Next r
End Function

Function Issues_API_ValidateForPosting(nRow As Long,ByRef sMessage As String) As Boolean
    Dim oSh As Object,sev As String
    Issues_API_ValidateForPosting=False:sMessage=""
    If Not ThisComponent.Sheets.hasByName(WMSISS_SHEET) Then sMessage="Нет листа Выдачи.":Exit Function
    oSh=ThisComponent.Sheets.getByName(WMSISS_SHEET)
    If nRow<1 Then sMessage="Неверная строка.":Exit Function
    gWMSISS_Busy=True:Issues_ValidateRow ThisComponent,oSh,nRow,False,False:gWMSISS_Busy=False
    sev=Issues_CellText(oSh,nRow,WMSISS_X_SEVERITY)
    If sev="CRITICAL" Then sMessage=Issues_CellText(oSh,nRow,WMSISS_T_ERROR):Exit Function
    If Issues_CellNumber(oSh.getCellByPosition(3,nRow))<=0 Then sMessage="Количество выдачи должно быть положительным.":Exit Function
    If Trim(oSh.getCellByPosition(5,nRow).String)="" Then sMessage="Не заполнено Кто получил.":Exit Function
    sMessage="OK":Issues_API_ValidateForPosting=True
End Function

Function Issues_API_GetPayload(nRow As Long) As Variant
    Dim oSh As Object,a(0 To 15) As Variant
    If Not ThisComponent.Sheets.hasByName(WMSISS_SHEET) Then Issues_API_GetPayload=a():Exit Function
    oSh=ThisComponent.Sheets.getByName(WMSISS_SHEET)
    a(0)=Issues_CellText(oSh,nRow,WMSISS_T_ID):a(1)=oSh.getCellByPosition(1,nRow).String:a(2)=oSh.getCellByPosition(2,nRow).String:a(3)=Issues_CellNumber(oSh.getCellByPosition(3,nRow)):a(4)=oSh.getCellByPosition(4,nRow).String:a(5)=oSh.getCellByPosition(5,nRow).String
    a(6)=Issues_CellDate(oSh.getCellByPosition(6,nRow),True):a(7)=oSh.getCellByPosition(7,nRow).String:a(8)=oSh.getCellByPosition(8,nRow).String:a(9)=Issues_CellNumber(oSh.getCellByPosition(9,nRow)):a(10)=Issues_CellDate(oSh.getCellByPosition(10,nRow),True):a(11)=oSh.getCellByPosition(11,nRow).String
    a(12)=Issues_CellText(oSh,nRow,WMSISS_T_MODE):a(13)=Issues_CellText(oSh,nRow,WMSISS_T_STATE):a(14)=Issues_CellText(oSh,nRow,WMSISS_T_ERROR):a(15)=Issues_CellText(oSh,nRow,WMSISS_X_RETSTATE)
    Issues_API_GetPayload=a()
End Function

Function Issues_API_GetReturnDelta(nRow As Long) As Double
    Dim oSh As Object,ret As Double,posted As Double,qty As Double
    If Not ThisComponent.Sheets.hasByName(WMSISS_SHEET) Then Exit Function
    oSh=ThisComponent.Sheets.getByName(WMSISS_SHEET):qty=Issues_CellNumber(oSh.getCellByPosition(3,nRow)):ret=Issues_CellNumber(oSh.getCellByPosition(9,nRow)):posted=Issues_CellNumber(oSh.getCellByPosition(Issues_FindHeader(oSh,WMSISS_T_RETPOSTED),nRow))
    If ret<0 Or posted<0 Or ret>qty+WMSISS_EPS Or ret<posted-WMSISS_EPS Then Exit Function
    If ret-posted>WMSISS_EPS Then Issues_API_GetReturnDelta=ret-posted
End Function

Sub Issues_API_MarkPosted(nRow As Long,sSyncHash As String)
    Dim oSh As Object
    If Not ThisComponent.Sheets.hasByName(WMSISS_SHEET) Then Exit Sub
    oSh=ThisComponent.Sheets.getByName(WMSISS_SHEET):gWMSISS_Busy=True
    oSh.getCellByPosition(Issues_FindHeader(oSh,WMSISS_T_STATE),nRow).String=WMSISS_STATE_DONE:oSh.getCellByPosition(Issues_FindHeader(oSh,WMSISS_T_HASH),nRow).String=sSyncHash:oSh.getCellByPosition(Issues_FindHeader(oSh,WMSISS_T_DIRTY),nRow).String="0":oSh.getCellByPosition(Issues_FindHeader(oSh,WMSISS_X_BASEFP),nRow).String=Issues_BaseFingerprint(oSh,nRow)
    gWMSISS_Busy=False
End Sub

Sub Issues_API_MarkReturnPosted(nRow As Long,dTotalPosted As Double,nSeq As Long,sReturnHash As String)
    Dim oSh As Object
    If Not ThisComponent.Sheets.hasByName(WMSISS_SHEET) Then Exit Sub
    oSh=ThisComponent.Sheets.getByName(WMSISS_SHEET):gWMSISS_Busy=True
    oSh.getCellByPosition(Issues_FindHeader(oSh,WMSISS_T_RETPOSTED),nRow).Value=dTotalPosted:oSh.getCellByPosition(Issues_FindHeader(oSh,WMSISS_T_RETSEQ),nRow).Value=nSeq:oSh.getCellByPosition(Issues_FindHeader(oSh,WMSISS_T_RETHASH),nRow).String=sReturnHash:oSh.getCellByPosition(Issues_FindHeader(oSh,WMSISS_T_RETERR),nRow).String="":oSh.getCellByPosition(Issues_FindHeader(oSh,WMSISS_X_RETFP),nRow).String=Issues_ReturnFingerprint(oSh,nRow)
    Issues_ValidateRow ThisComponent,oSh,nRow,False,False:Issues_ApplyRowVisual oSh,nRow:gWMSISS_Busy=False
End Sub

' ============================================================================
' SELF CHECK
' ============================================================================

Function Issues_RunSelfCheck(oDoc As Object,ByRef sReport As String) As Boolean
    Dim oSh As Object,a As Variant,i As Long,c As Long,bad As Long,s As String,ev As Variant,v As Object,db As Object,addr As Variant,tb As Variant,k As Long,nfmt As Object,preview As String
    sReport="":Issues_RunSelfCheck=False
    On Error GoTo EH
    If Not Issues_Preflight(oDoc,s) Then sReport=s:Exit Function
    oSh=oDoc.Sheets.getByName(WMSISS_SHEET):a=Issues_ExtensionHeaders()
    For i=LBound(a) To UBound(a)
        c=Issues_FindHeader(oSh,CStr(a(i)))
        If c<0 Then
            sReport=sReport & "Нет техполя " & CStr(a(i)) & "." & Chr(10)
            bad=bad+1
        ElseIf oSh.Columns.getByIndex(c).IsVisible Then
            sReport=sReport & "Техполе видно: " & CStr(a(i)) & "." & Chr(10)
            bad=bad+1
        End If
    Next i
    ev=oSh.Events.getByName("OnChange")
    If IsNull(ev) Or IsEmpty(ev) Then
        sReport=sReport & "Не назначено OnChange." & Chr(10)
        bad=bad+1
    ElseIf InStr(Issues_EventScript(ev),"WMS_03_Issues_FINAL.Issues_OnContentChanged")=0 Then
        sReport=sReport & "OnChange назначен не на Issues_OnContentChanged." & Chr(10)
        bad=bad+1
    End If
    v=oSh.getCellByPosition(8,1).Validation:If v.Type<>com.sun.star.sheet.ValidationType.LIST Or InStr(v.Formula1,"Да")=0 Or InStr(v.Formula1,"Нет")=0 Then sReport=sReport & "В I Возвратный нет корректного Да/Нет validation." & Chr(10):bad=bad+1
    v=oSh.getCellByPosition(5,1).Validation:If v.Type<>com.sun.star.sheet.ValidationType.LIST Then sReport=sReport & "В F Кто получил нет списка-подсказки." & Chr(10):bad=bad+1
    If Not oDoc.NamedRanges.hasByName(WMSISS_LIST_EMP) Then sReport=sReport & "Нет named range сотрудников." & Chr(10):bad=bad+1
    If Not oDoc.DatabaseRanges.hasByName("WMS00_DB_03") Then
        sReport=sReport & "Нет AutoFilter range WMS00_DB_03." & Chr(10)
        bad=bad+1
    Else
        db=oDoc.DatabaseRanges.getByName("WMS00_DB_03")
        addr=db.getDataArea()
        If addr.StartColumn<>0 Or addr.EndColumn<Issues_LastHeaderCol(oSh) Then
            sReport=sReport & "Фильтр Выдачи не охватывает скрытые техполя." & Chr(10)
            bad=bad+1
        End If
    End If
    If oSh.getCellByPosition(5,2).BottomBorder2.LineWidth<=0 Then sReport=sReport & "Нет горизонтальных разделителей ячеек." & Chr(10):bad=bad+1
    If oSh.getCellByPosition(5,2).RightBorder2.LineWidth<=0 Then sReport=sReport & "Нет вертикальных разделителей ячеек." & Chr(10):bad=bad+1
    k=Issues_NumberFormatKey(oDoc,"QTY"):nfmt=CreateUnoService("com.sun.star.util.NumberFormatter"):nfmt.attachNumberFormatsSupplier(oDoc):preview=nfmt.convertNumberToString(k,1):If preview<>"1" Then sReport=sReport & "Формат количества показывает 1 как '" & preview & "'." & Chr(10):bad=bad+1
    If bad=0 Then
        sReport="WMS_03_Issues " & WMSISS_VERSION & " self-check: OK. Схема, сетка, формат количества, скрытые поля, dropdowns, named ranges, событие и фильтр корректны."
        Issues_RunSelfCheck=True
    Else
        sReport="WMS_03_Issues self-check: найдено проблем " & CStr(bad) & "." & Chr(10) & sReport
    End If
    Exit Function
EH:
    sReport="Issues self-check runtime error " & CStr(Err) & ": " & Error$
End Function

Function Issues_EventScript(ev As Variant) As String
    Dim i As Long
    On Error GoTo Done
    For i=LBound(ev) To UBound(ev):If ev(i).Name="Script" Then Issues_EventScript=CStr(ev(i).Value):Exit Function
    Next i
Done:
End Function

' ============================================================================
' FINGERPRINTS — BASE ISSUE EXCLUDES RETURN FIELDS J/K
' ============================================================================

Function Issues_BaseFingerprint(oSh As Object,r As Long) As String
    Dim c As Long,s As String
    For c=0 To WMSISS_USER_LAST_COL
        If c<>9 And c<>10 Then s=s & "|" & CStr(c) & "=" & oSh.getCellByPosition(c,r).Formula
    Next c
    Issues_BaseFingerprint=Issues_SimpleHash(s)
End Function

Function Issues_ReturnFingerprint(oSh As Object,r As Long) As String
    Issues_ReturnFingerprint=Issues_SimpleHash(oSh.getCellByPosition(9,r).Formula & "|" & oSh.getCellByPosition(10,r).Formula)
End Function

Function Issues_SimpleHash(s As String) As String
    Dim i As Long,h1 As Double,h2 As Double,ch As Long
    h1=5381:h2=7919
    For i=1 To Len(s)
        ch=Asc(Mid(s,i,1))
        If ch<0 Then ch=ch+256
        h1=(h1*33+ch) Mod 1000003
        h2=(h2*131+ch+i) Mod 1000033
    Next i
    Issues_SimpleHash=Hex(CLng(h1)) & "-" & Hex(CLng(h2)) & "-" & CStr(Len(s))
End Function

' ============================================================================
' GENERIC CELL / DATE / HEADER HELPERS
' ============================================================================

Sub Issues_ResetHeaderCache()
    gWMSISS_HeaderReady=False:gWMSISS_LastHeaderColCached=-1
End Sub

Function Issues_EnsureHeaderCache(oSh As Object) As Boolean
    Dim last As Long,c As Long,a() As String
    If gWMSISS_HeaderReady Then Issues_EnsureHeaderCache=True:Exit Function
    last=Issues_LastHeaderColRaw(oSh):If last<0 Then Exit Function
    ReDim a(0 To last)
    For c=0 To last
        a(c)=Issues_Canon(oSh.getCellByPosition(c,0).String)
    Next c
    gWMSISS_HeaderCache=a():gWMSISS_LastHeaderColCached=last:gWMSISS_HeaderReady=True:Issues_EnsureHeaderCache=True
End Function

Function Issues_FindHeader(oSh As Object,sHeader As String) As Long
    Dim c As Long,want As String,last As Long
    Issues_FindHeader=-1
    want=Issues_Canon(sHeader)
    ' The global cache belongs ONLY to the controlled Issues sheet. Other sheets
    ' have different schemas and must be scanned independently.
    If oSh.Name=WMSISS_SHEET Then
        If Issues_EnsureHeaderCache(oSh) Then
            For c=0 To gWMSISS_LastHeaderColCached
                If CStr(gWMSISS_HeaderCache(c))=want Then Issues_FindHeader=c:Exit Function
            Next c
        End If
    End If
    last=Issues_LastHeaderColRaw(oSh)
    For c=0 To last
        If Issues_Canon(oSh.getCellByPosition(c,0).String)=want Then Issues_FindHeader=c:Exit Function
    Next c
End Function

Function Issues_LastHeaderCol(oSh As Object) As Long
    If Issues_EnsureHeaderCache(oSh) Then
        Issues_LastHeaderCol=gWMSISS_LastHeaderColCached
    Else
        Issues_LastHeaderCol=Issues_LastHeaderColRaw(oSh)
    End If
End Function

Function Issues_LastHeaderColRaw(oSh As Object) As Long
    Dim c As Long,last As Long:last=-1
    For c=0 To 127:If Trim(oSh.getCellByPosition(c,0).String)<>"" Then last=c
    Next c
    Issues_LastHeaderColRaw=last
End Function

Function Issues_RowHasBusinessData(oSh As Object,r As Long) As Boolean
    Dim c As Long
    ' Empty Row Guard 1.0.3:
    ' A (№) may be prefilled for paper-style numbering. It is not sufficient to
    ' create a WMS issue. Real business content starts in B:L.
    For c=1 To WMSISS_USER_LAST_COL
        If Trim(oSh.getCellByPosition(c,r).Formula)<>"" Then Issues_RowHasBusinessData=True:Exit Function
    Next c
End Function

Function Issues_CellText(oSh As Object,r As Long,sHeader As String) As String
    Dim c As Long:c=Issues_FindHeader(oSh,sHeader):If c>=0 Then Issues_CellText=Trim(oSh.getCellByPosition(c,r).String)
End Function

Function Issues_LastContentRow(oSh As Object,nLastCol As Long) As Long
    ' OPT-1: existing shared batch reader replaces per-cell UNO calls.
    Issues_LastContentRow=0
    If nLastCol<0 Then Exit Function
    Issues_LastContentRow=WMSCore_LastContentRow(oSh,nLastCol)
End Function

Function Issues_NextRow(oSh As Object,nKeyCol As Long,nStart As Long) As Long
    Dim r As Long
    r=Issues_LastContentRow(oSh,nKeyCol)+1
    If r<nStart Then r=nStart
    Do While Trim(oSh.getCellByPosition(nKeyCol,r).Formula)<>"":r=r+1:Loop
    Issues_NextRow=r
End Function

Function Issues_Canon(ByVal s As String) As String
    s=LCase(Trim(s)):s=Replace(s,"ё","е"):s=Replace(s,Chr(160)," "):Do While InStr(s,"  ")>0:s=Replace(s,"  "," "):Loop:Issues_Canon=s
End Function

Function Issues_TryNumber(oCell As Object,ByRef d As Double) As Boolean
    Dim s As String:d=0
    If oCell.Type=1 Then d=oCell.Value:Issues_TryNumber=True:Exit Function
    s=Trim(oCell.String):If s="" Then Exit Function
    s=Replace(s,Chr(160),""):s=Replace(s," ",""):s=Replace(s,",",".")
    If IsNumeric(s) Then d=CDbl(s):Issues_TryNumber=True
End Function

Function Issues_CellNumber(oCell As Object) As Double
    Dim d As Double:If Issues_TryNumber(oCell,d) Then Issues_CellNumber=d
End Function

Function Issues_CellDate(oCell As Object,bInferYear As Boolean) As Double
    Dim s As String,p As Variant,d As Integer,m As Integer,y As Integer,dt As Date
    On Error GoTo Bad
    If oCell.Type=1 And oCell.Value>0 Then
        Issues_CellDate=oCell.Value
        Exit Function
    End If
    s=Trim(oCell.String)
    If s="" Then Exit Function
    s=Replace(s,"/",".")
    s=Replace(s,"-",".")
    p=Split(s,".")
    If UBound(p)=1 And bInferYear Then
        d=CInt(p(0))
        m=CInt(p(1))
        y=Year(Date)
    ElseIf UBound(p)=2 Then
        If Len(p(0))=4 Then
            y=CInt(p(0))
            m=CInt(p(1))
            d=CInt(p(2))
        Else
            d=CInt(p(0))
            m=CInt(p(1))
            y=CInt(p(2))
            If y<100 Then y=2000+y
        End If
    Else
        GoTo Bad
    End If
    dt=DateSerial(y,m,d)
    If Year(dt)=y And Month(dt)=m And Day(dt)=d Then Issues_CellDate=CDbl(dt)
Bad:
End Function

Sub Issues_WriteDate(oDoc As Object,oCell As Object,v As Double)
    oCell.Value=v:oCell.NumberFormat=Issues_NumberFormatKey(oDoc,"DATE")
End Sub

Function Issues_NumText(d As Double) As String
    If Abs(d-Int(d))<WMSISS_EPS Then
        Issues_NumText=CStr(CLng(d))
    Else
        Issues_NumText=CStr(d)
    End If
End Function

Function Issues_Stamp(v As Variant) As String
    Issues_Stamp=Format(CDate(v),"DD.MM.YYYY HH:MM:SS")
End Function

Function Issues_TimestampCompact(v As Variant) As String
    Issues_TimestampCompact=Format(CDate(v),"YYYYMMDD-HHMMSS")
End Function

Function Issues_ColName(n As Long) As String
    Dim s As String,x As Long:x=n+1
    Do While x>0:x=x-1:s=Chr(65+(x Mod 26)) & s:x=Int(x/26):Loop
    Issues_ColName=s
End Function

' ============================================================================
' AUDIT / ERRORS
' ============================================================================

Sub Issues_LogAudit(oDoc As Object,sEvent As String,sDesc As String)
    Dim oSh As Object,r As Long
    On Error GoTo Done
    If Not oDoc.Sheets.hasByName(WMSISS_AUDIT) Then Exit Sub
    oSh=oDoc.Sheets.getByName(WMSISS_AUDIT):r=Issues_NextRow(oSh,0,1):oSh.getCellByPosition(0,r).String=Issues_Stamp(Now):oSh.getCellByPosition(1,r).String="":oSh.getCellByPosition(2,r).String=WMSISS_SHEET:oSh.getCellByPosition(3,r).String=sEvent:oSh.getCellByPosition(4,r).String=sDesc
Done:
End Sub

Sub Issues_LogError(oDoc As Object,sWhere As String,sErr As String)
    Dim oSh As Object,r As Long
    On Error GoTo Done
    If Not oDoc.Sheets.hasByName(WMSISS_ERRORS) Then Exit Sub
    oSh=oDoc.Sheets.getByName(WMSISS_ERRORS):r=Issues_NextRow(oSh,0,1):oSh.getCellByPosition(0,r).String=Issues_Stamp(Now):oSh.getCellByPosition(1,r).String="":oSh.getCellByPosition(2,r).String="Выдачи/" & sWhere:oSh.getCellByPosition(3,r).String=sErr:oSh.getCellByPosition(4,r).String="OPEN"
Done:
End Sub


' ============================================================================
' BUTTON PANEL 1.0.4
' ============================================================================

Function Issues_InstallButtonsCore(oDoc As Object,oSh As Object,ByRef sErr As String) As Boolean
    Dim oDP As Object,oForms As Object,oForm As Object
    Dim p As New com.sun.star.awt.Point, z As New com.sun.star.awt.Size
    Dim baseX As Long,baseY As Long,w As Long,h As Long,gap As Long
    Issues_InstallButtonsCore=False:sErr=""
    On Error GoTo EH
    oDP=oSh.DrawPage
    oForms=oDP.Forms
    Issues_RemoveButtonPanel oSh
    oForm=oDoc.createInstance("com.sun.star.form.component.Form")
    oForm.Name="WMS_ISSUES_BUTTON_PANEL"
    oForms.insertByName("WMS_ISSUES_BUTTON_PANEL",oForm)

    baseX=oSh.getCellByPosition(WMSISS_USER_LAST_COL,0).Position.X + oSh.getCellByPosition(WMSISS_USER_LAST_COL,0).Size.Width + 350
    baseY=oSh.getCellByPosition(0,0).Position.Y + 120
    w=4200:h=650:gap=120

    Issues_AddButton oDoc,oSh,oForm,"WMS_ISS_BTN_NEW","Новая выдача",baseX,baseY,w,h,"Issues_ButtonNewIssue"
    Issues_AddButton oDoc,oSh,oForm,"WMS_ISS_BTN_CONDUCT","Провести выдачу",baseX,baseY+(h+gap),w,h,"Issues_ButtonConductIssue"
    Issues_AddButton oDoc,oSh,oForm,"WMS_ISS_BTN_CONDUCTALL","Провести все",baseX,baseY+2*(h+gap),w,h,"Issues_ButtonConductAll"
    Issues_AddButton oDoc,oSh,oForm,"WMS_ISS_BTN_RETURN","Возврат",baseX,baseY+3*(h+gap),w,h,"Issues_ButtonReturn"
    Issues_AddButton oDoc,oSh,oForm,"WMS_ISS_BTN_CHECK","Проверить строку",baseX,baseY+4*(h+gap),w,h,"Issues_ButtonCheckSelected"
    Issues_AddButton oDoc,oSh,oForm,"WMS_ISS_BTN_DIAG","Диагностика",baseX,baseY+5*(h+gap),w,h,"Issues_ButtonShowDiagnostics"
    Issues_AddButton oDoc,oSh,oForm,"WMS_ISS_BTN_ALL","Проверить все",baseX,baseY+6*(h+gap),w,h,"Issues_ButtonValidateAll"
    Issues_AddButton oDoc,oSh,oForm,"WMS_ISS_BTN_LISTS","Обновить списки",baseX,baseY+7*(h+gap),w,h,"Issues_ButtonRefreshLists"
    Issues_AddButton oDoc,oSh,oForm,"WMS_ISS_BTN_LOOKUP","Найти по коду",baseX,baseY+8*(h+gap),w,h,"Issues_ButtonLookupCode"
    Issues_AddButton oDoc,oSh,oForm,"WMS_ISS_BTN_BULKFILL","Заполнить все по кодам",baseX,baseY+9*(h+gap),w,h,"Issues_ButtonFillAllByCodes"
    Issues_AddButton oDoc,oSh,oForm,"WMS_ISS_BTN_SELF","SelfCheck",baseX,baseY+10*(h+gap),w,h,"Issues_ButtonSelfCheck"

    Issues_InstallButtonsCore=True
    Exit Function
EH:
    sErr="Ошибка панели кнопок " & CStr(Err) & ": " & Error$
End Function

Sub Issues_RemoveButtonPanel(oSh As Object)
    Dim oDP As Object,oForms As Object,i As Long,oShape As Object,oCtl As Object,nm As String
    On Error Resume Next
    oDP=oSh.DrawPage
    For i=oDP.Count-1 To 0 Step -1
        oShape=oDP.getByIndex(i)
        nm=""
        oCtl=oShape.Control
        nm=oCtl.Name
        If Left(nm,12)="WMS_ISS_BTN_" Then oDP.remove(oShape)
    Next i
    oForms=oDP.Forms
    If oForms.hasByName("WMS_ISSUES_BUTTON_PANEL") Then oForms.removeByName("WMS_ISSUES_BUTTON_PANEL")
    On Error GoTo 0
End Sub

Sub Issues_AddButton(oDoc As Object,oSh As Object,oForm As Object,ctlName As String,labelText As String,x As Long,y As Long,w As Long,h As Long,macroName As String)
    Dim oCtl As Object,oShape As Object,ev As New com.sun.star.script.ScriptEventDescriptor
    Dim p As New com.sun.star.awt.Point,z As New com.sun.star.awt.Size,idx As Long
    oCtl=oDoc.createInstance("com.sun.star.form.component.CommandButton")
    oCtl.Name=ctlName
    oCtl.Label=labelText
    oCtl.Tabstop=False
    oForm.insertByName(ctlName,oCtl)
    idx=oForm.Count-1

    oShape=oDoc.createInstance("com.sun.star.drawing.ControlShape")
    p.X=x:p.Y=y:z.Width=w:z.Height=h
    oShape.Position=p:oShape.Size=z:oShape.Control=oCtl
    oSh.DrawPage.add(oShape)

    ev.ListenerType="com.sun.star.awt.XActionListener"
    ev.EventMethod="actionPerformed"
    ev.AddListenerParam=""
    ev.ScriptType="Script"
    ev.ScriptCode="vnd.sun.star.script:Standard.WMS_03_Issues_FINAL." & macroName & "?language=Basic&location=document"
    oForm.registerScriptEvent(idx,ev)
End Sub
