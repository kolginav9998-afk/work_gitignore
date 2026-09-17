Option Explicit

' ============================================================================
' WMS_02_Orders_FINAL.bas
' Production controller for sheet "Заказы" in the clean WMS architecture.
' LibreOffice Calc / ODS / Linux / LibreOffice Basic + UNO only.
'
' VERIFIED SOURCE RULES PRESERVED:
' - A,B,D,E,F,H,I,J,K,L,O,P are user-entered order-stage fields.
' - C,G,N,T,W are filled by the user when goods/documents arrive.
' - M receipt date is automatic on first positive fact, but remains editable.
' - Q Status is automatic, but user can override from a drop-down.
' - R Category is exact-match automation when safe; user may override.
' - S Control is fully automatic.
' - U Comment remains fully user controlled.
' - V Product code is exact-match automation when safe; a user code wins.
' - W Seller is user entered.
' - Automatic status never sets "Оприходовано".
' - Order can enter movement flow only with positive fact + document + manual
'   "Оприходовано". This controller prepares/queues; common movement engine
'   performs historical-safe posting later.
' - Fact > ordered is CRITICAL. Fact < ordered is a warning only.
' - No fuzzy auto-merge. Exact mappings/code/history only.
' - Column A is the user-entered position number inside an order: every 1 starts
'   a new order; 2,3,4... continue that order. It is a hard grouping boundary.
' - Invoice is one per OrderGroupID; document number/date may differ by row.
'   subsequent numbered positions inherit them logically via hidden metadata.
' - Sorting is allowed and must not detach SourceID from the logical row.
' ============================================================================

Global Const WMSORD_VERSION = "1.0.26-GENERIC-ORDER-DELIVERY-UI"
Global Const WMSORD_SCHEMA = "1.5"
Global Const WMSORD_TARGET_LO = "25.8.7.3"
Global Const WMSORD_SHEET = "Заказы"
Global Const WMSORD_QUEUE = "SYS_WMS_QUEUE"
Global Const WMSORD_SETTINGS = "Настройки WMS"
Global Const WMSORD_AUDIT = "Журнал WMS"
Global Const WMSORD_ERRORS = "Ошибки WMS"
Global Const WMSORD_MAP = "Сопоставления"
Global Const WMSORD_NOM = "Номенклатура"
Global Const WMSORD_HEADER_ROW = 0
Global Const WMSORD_USER_LAST_COL = 25
Global Const WMSORD_MIN_STYLE_ROWS = 500
Global Const WMSORD_FILTER_BUFFER = 100
Global Const WMSORD_DEFAULT_LEGACY_DAYS = 30
Global Const WMSORD_MAX_EVENT_ROWS = 2500
Global Const WMSORD_BULK_THRESHOLD = 50
Global Const WMSORD_EPS = 0.0000001

Global Const WMSORD_STATUS_TRANSIT = "В пути"
Global Const WMSORD_STATUS_NO_DOCS = "Получено без документов"
Global Const WMSORD_STATUS_RECEIVED = "Получено"
Global Const WMSORD_STATUS_POST = "Оприходовано"
Global Const WMSORD_STATUS_CANCEL = "Отменено"

Global Const WMSORD_T_ID = "_WMS_SourceID"
Global Const WMSORD_T_STATE = "_WMS_State"
Global Const WMSORD_T_HASH = "_WMS_SyncHash"
Global Const WMSORD_T_ERROR = "_WMS_Error"
Global Const WMSORD_T_DIRTY = "_WMS_Dirty"
Global Const WMSORD_T_MODE = "_WMS_Mode"
Global Const WMSORD_T_LEGACY = "_WMS_LegacyKey"
Global Const WMSORD_T_TOUCH = "_WMS_LastTouch"

Global Const WMSORD_X_ROWVER = "_WMS_OrderRowVersion"
Global Const WMSORD_X_STATUSMODE = "_WMS_StatusMode"
Global Const WMSORD_X_CODESUG = "_WMS_CodeSuggestion"
Global Const WMSORD_X_CODEORIGIN = "_WMS_CodeOrigin"
Global Const WMSORD_X_CATSUG = "_WMS_CategorySuggestion"
Global Const WMSORD_X_CATORIGIN = "_WMS_CategoryOrigin"
Global Const WMSORD_X_LOCSUG = "_WMS_LocationSuggestion"
Global Const WMSORD_X_FLAGS = "_WMS_ValidationFlags"
Global Const WMSORD_X_SEVERITY = "_WMS_Severity"
Global Const WMSORD_X_VALIDATED = "_WMS_LastValidated"
Global Const WMSORD_X_DUPKEY = "_WMS_DuplicateKey"
Global Const WMSORD_X_DATEORIGIN = "_WMS_ReceiptDateOrigin"
Global Const WMSORD_X_FINGERPRINT = "_WMS_AutoFingerprint"
Global Const WMSORD_X_SUGSOURCE = "_WMS_SuggestionSource"
Global Const WMSORD_X_GROUPID = "_WMS_OrderGroupID"
Global Const WMSORD_X_EFFINVOICE = "_WMS_EffectiveInvoice"
Global Const WMSORD_X_EFFDOC = "_WMS_EffectiveDocNo"
Global Const WMSORD_X_EFFDOCDATE = "_WMS_EffectiveDocDate"
Global Const WMSORD_X_EFFSUPPLIER = "_WMS_EffectiveSupplier"
Global Const WMSORD_X_GROUPFLAGS = "_WMS_OrderGroupFlags"

Global gWMSORD_Busy As Boolean
Global gWMSORD_EventDepth As Long
Global gWMSORD_LastError As String
Global gWMSORD_LastReport As String
Global gWMSORD_LastTouched As Long
Global gWMSORD_LastEventMs As Double
Global gWMSORD_DupCache As Variant
Global gWMSORD_DupCacheReady As Boolean
Global gWMSORD_QueueCacheReady As Boolean
Global gWMSORD_QueueNextRow As Long
Global gWMSORD_QueueHashKeys() As String
Global gWMSORD_QueueHashRows() As Long
Global gWMSORD_QueueHashCap As Long
Global gWMSORD_QueueHashCount As Long
Global gWMSORD_HeaderCache As Variant
Global gWMSORD_HeaderCacheReady As Boolean
Global gWMSORD_LastHeaderColCached As Long
Global gWMSORD_HistoryCache As Variant
Global gWMSORD_HistoryCacheReady As Boolean
Global gWMSORD_GroupSeq As Long

' ============================================================================
' PUBLIC INSTALL / UPGRADE / QA
' ============================================================================

Sub InstallOrders()
    Dim sReport As String
    If Orders_InstallCore(ThisComponent, True, sReport) Then
        MsgBox sReport, 64, "WMS — Заказы готовы"
    Else
        MsgBox sReport, 16, "WMS — Заказы: установка остановлена"
    End If
End Sub

Sub InstallOrdersQuiet()
    Dim sReport As String
    Call Orders_InstallCore(ThisComponent, False, sReport)
End Sub

Sub UpgradeOrders()
    InstallOrders
End Sub

Sub Orders_RefreshDesign()
    Dim oDoc As Object, oSh As Object
    oDoc = ThisComponent
    If Not Orders_Preflight(oDoc, gWMSORD_LastError) Then
        MsgBox gWMSORD_LastError, 16, "WMS — Заказы"
        Exit Sub
    End If
    oSh = oDoc.Sheets.getByName(WMSORD_SHEET)
    gWMSORD_Busy = True
    Orders_ApplyDesign oDoc, oSh
    Orders_ApplyValidations oDoc, oSh
    Orders_RefreshFilter oDoc, oSh
    Orders_ApplyFreeze oDoc, oSh
    gWMSORD_Busy = False
    MsgBox "Дизайн, форматы, валидации и фильтр листа 'Заказы' обновлены без изменения пользовательских данных.", 64, "WMS — Заказы"
End Sub

Sub Orders_RevalidateAll()
    Dim sReport As String
    If Orders_RevalidateAllCore(ThisComponent, True, sReport) Then
        MsgBox sReport, 64, "WMS — Заказы"
    Else
        MsgBox sReport, 48, "WMS — Заказы"
    End If
End Sub

Sub Orders_SelfCheck()
    Dim s As String
    If Orders_RunSelfCheck(ThisComponent, s) Then
        MsgBox s, 64, "WMS — Заказы self-check"
    Else
        MsgBox s, 48, "WMS — Заказы self-check"
    End If
End Sub

Function Orders_SelfCheckText() As String
    Dim s As String
    Call Orders_RunSelfCheck(ThisComponent, s)
    Orders_SelfCheckText = s
End Function

Function Orders_LastReportText() As String
    Orders_LastReportText = gWMSORD_LastReport
End Function

Function Orders_CompileProbe() As String
    Orders_CompileProbe = WMSORD_VERSION & "|" & WMSORD_SCHEMA & "|" & WMSORD_SHEET & "|" & CStr(Orders_ExtensionHeaderCount())
End Function

Function Orders_IsInstalled() As Boolean
    Dim oDoc As Object
    oDoc = ThisComponent
    Orders_IsInstalled = (Orders_GetSetting(oDoc, "WMS_OrdersVersion", "") = WMSORD_VERSION)
End Function


Function Orders_UpgradeUniversalReceiptColumns(oDoc As Object,ByRef sErr As String) As Boolean
    Dim oSh As Object,h23 As String,h24 As String,h25 As String
    Orders_UpgradeUniversalReceiptColumns=False:sErr=""
    On Error GoTo EH

    If Not oDoc.Sheets.hasByName(WMSORD_SHEET) Then
        sErr="Нет листа 'Заказы'."
        Exit Function
    End If
    oSh=oDoc.Sheets.getByName(WMSORD_SHEET)

    h23=Trim(oSh.getCellByPosition(23,0).String)
    h24=Trim(oSh.getCellByPosition(24,0).String)
    h25=Trim(oSh.getCellByPosition(25,0).String)

    ' 1.0.19 target:
    ' R = Категория (existing product category)
    ' X = Источник прихода
    ' Y = Подкатегория
    ' Z = hidden legacy destination field kept only to avoid destructive migration.
    If h23="Источник прихода" And h24="Подкатегория" Then
        If h25="" Then oSh.getCellByPosition(25,0).String="_WMS_LegacyDestination"
        oSh.Columns.getByIndex(25).IsVisible=False
        Orders_UpgradeUniversalReceiptColumns=True
        Exit Function
    End If

    ' Upgrade from 1.0.18. Preserve old destination values in hidden Z.
    ' We do NOT reinterpret them as product categories/subcategories.
    If h23="Источник прихода" And h24="Категория назначения" And h25="Подкатегория назначения" Then
        ' Preserve both old fields in one hidden legacy text when possible.
        Dim r As Long,last As Long,a As String,b As String
        last=Orders_LastContentRow(oSh,25)
        For r=1 To last
            a=Trim(oSh.getCellByPosition(24,r).String)
            b=Trim(oSh.getCellByPosition(25,r).String)
            If a<>"" Or b<>"" Then
                oSh.getCellByPosition(25,r).String="Категория назначения=" & a & "; Подкатегория назначения=" & b
            End If
            ' New product subcategory starts empty; no false semantic conversion.
            oSh.getCellByPosition(24,r).String=""
        Next r
        oSh.getCellByPosition(24,0).String="Подкатегория"
        oSh.getCellByPosition(25,0).String="_WMS_LegacyDestination"
        oSh.Columns.getByIndex(25).IsVisible=False
        Orders_UpgradeUniversalReceiptColumns=True
        Exit Function
    End If

    ' Upgrade from 1.0.17 and older.
    If h23="_WMS_SourceID" Or h23="_WMS_ReceiptType" Then
        oSh.Columns.insertByIndex(23,3)
        oSh.getCellByPosition(23,0).String="Источник прихода"
        oSh.getCellByPosition(24,0).String="Подкатегория"
        oSh.getCellByPosition(25,0).String="_WMS_LegacyDestination"
        oSh.Columns.getByIndex(25).IsVisible=False
        Orders_UpgradeUniversalReceiptColumns=True
        Exit Function
    End If

    sErr="Не удалось безопасно обновить лист 'Заказы': X='" & h23 & _
         "', Y='" & h24 & "', Z='" & h25 & "'. Данные не изменены."
    Exit Function
EH:
    sErr="Миграция Заказы 1.0.19: " & CStr(Err) & " " & Error$
End Function

Sub Orders_RepairAccidentalReceiptTypeHeader(oDoc As Object)
    Dim oSh As Object,srcCol As Long,c As Long,cur As Object,lastCol As Long
    On Error Resume Next
    If Not oDoc.Sheets.hasByName(WMSORD_SHEET) Then Exit Sub
    oSh=oDoc.Sheets.getByName(WMSORD_SHEET)

    srcCol=WMSORD_USER_LAST_COL+1
    If Trim(oSh.getCellByPosition(srcCol,0).String)="_WMS_ReceiptType" Then
        oSh.getCellByPosition(srcCol,0).String="_WMS_SourceID"
        cur=oSh.createCursor()
        cur.gotoEndOfUsedArea(True)
        lastCol=cur.RangeAddress.EndColumn
        For c=srcCol+1 To lastCol
            If Trim(oSh.getCellByPosition(c,0).String)="_WMS_ReceiptType" Then
                oSh.Columns.getByIndex(c).IsVisible=False
                Exit Sub
            End If
        Next c
        c=lastCol+1
        oSh.getCellByPosition(c,0).String="_WMS_ReceiptType"
        oSh.Columns.getByIndex(c).IsVisible=False
    End If
    On Error GoTo 0
End Sub

Function Orders_InstallCore(oDoc As Object, bShow As Boolean, ByRef sReport As String) As Boolean
    Dim oSh As Object, sErr As String, sReval As String, bLocked As Boolean
    Orders_InstallCore = False
    gWMSORD_LastError = ""
    gWMSORD_LastReport = ""
    sReport = ""

    On Error GoTo Fatal

    ' 1.0.18 migration: insert three visible universal-receipt fields
    ' before all existing technical columns. Existing data/tech columns shift
    ' to the right intact; nothing is overwritten.
    If Not Orders_UpgradeUniversalReceiptColumns(oDoc, sErr) Then sReport=sErr:GoTo Done

    ' Keep compatibility with the old 1.0.14 repair path after migration.
    Orders_RepairAccidentalReceiptTypeHeader oDoc

    ' 1.0.24: accept the old real-world L header and migrate only the header text.
    ' User data in column L is never changed.
    If Not Orders_MigrateSupplierHeader(oDoc,sErr) Then sReport=sErr:GoTo Done

    If Not Orders_Preflight(oDoc, sErr) Then sReport = sErr: GoTo Done

    gWMSORD_Busy = True
    On Error Resume Next
    oDoc.lockControllers
    If Err = 0 Then bLocked = True Else Err = 0
    On Error GoTo Fatal

    oSh = oDoc.Sheets.getByName(WMSORD_SHEET)
    If Not Orders_EnsureExtensionColumns(oSh, sErr) Then sReport = sErr: GoTo SafeFail
    Orders_ResetHeaderCache
    Orders_EnsureSettings oDoc
    Orders_ApplyDesign oDoc, oSh
    Orders_ApplyValidations oDoc, oSh
    If Not Orders_InstallEvent(oSh, sErr) Then sReport = sErr: GoTo SafeFail
    Orders_RefreshFilter oDoc, oSh
    Orders_ApplyFreeze oDoc, oSh

    ' 1.0.12: install/update the working button panel on sheet "Заказы".
    If Not Orders_InstallButtonsCore(oDoc, oSh, sErr) Then
        sReport = "Модуль установлен частично, но панель кнопок не создана: " & sErr
        GoTo SafeFail
    End If

    Orders_ResetQueueCache

    ' 1.0.11 migration: remove technical ghosts created by the previous
    ' auto-status feedback loop on genuinely empty rows. Conducted/synced rows
    ' are never touched by this migration.
    Call Orders_EmptyRowGuardCleanupCore(oDoc, oSh)
    Orders_ResetQueueCache

    If Not Orders_RevalidateAllCore(oDoc, False, sReval) Then
        sReport = "Модуль установлен частично, но первичная проверка строк не завершилась: " & sReval
        GoTo SafeFail
    End If

    If Not Orders_RunSelfCheck(oDoc, sErr) Then sReport = sErr: GoTo SafeFail

    Orders_SetSetting oDoc, "WMS_OrdersVersion", WMSORD_VERSION, "Версия отдельного production-контроллера листа Заказы", True
    Orders_SetSetting oDoc, "WMS_OrdersState", "READY", "Состояние контроллера Заказы", True
    Orders_SetSetting oDoc, "WMS_OrdersLastCheck", Orders_Stamp(Now), "Последний успешный self-check Заказы", True
    Orders_LogAudit oDoc, "ORDERS_INSTALL", "WMS_02_Orders " & WMSORD_VERSION & " установлен; schema " & WMSORD_SCHEMA

    sReport = "WMS_02_Orders " & WMSORD_VERSION & " установлен." & Chr(10) & _
              "Лист 'Заказы': production-дизайн, автоматические статусы, дата поступления, точные подсказки кода/категории, контроль строки, SourceID/DIRTY/queue, выпадающий статус, фильтр, безопасная сортировка и self-check — OK." & Chr(10) & _
              "Empty Row Guard: пустые строки больше не активируются авто-статусом/контролем." & Chr(10) & _
              "Оприходовано по-прежнему ставится только вручную."
    Orders_InstallCore = True
    GoTo Done

SafeFail:
    gWMSORD_LastError = sReport
    Orders_LogError oDoc, "ORDERS_INSTALL", sReport
    Orders_SetSetting oDoc, "WMS_OrdersState", "ERROR", "Последняя установка/проверка Заказы завершилась ошибкой", True
    GoTo Done

Fatal:
    sReport = "WMS_02_Orders runtime error " & CStr(Err) & ": " & Error$
    Resume SafeFail

Done:
    If bLocked Then
        On Error Resume Next
        oDoc.unlockControllers
        On Error GoTo 0
    End If
    gWMSORD_Busy = False
    gWMSORD_LastReport = sReport
End Function

Function Orders_MigrateSupplierHeader(oDoc As Object,ByRef sErr As String) As Boolean
    Dim oSh As Object,actual As String
    Orders_MigrateSupplierHeader=False
    sErr=""
    On Error GoTo EH

    If IsNull(oDoc) Or IsEmpty(oDoc) Then
        sErr="Нет открытой книги Calc."
        Exit Function
    End If
    If Not oDoc.Sheets.hasByName(WMSORD_SHEET) Then
        sErr="Нет листа 'Заказы'."
        Exit Function
    End If

    oSh=oDoc.Sheets.getByName(WMSORD_SHEET)
    actual=Trim(oSh.getCellByPosition(11,WMSORD_HEADER_ROW).String)

    ' Current header: nothing to do.
    If Orders_Canon(actual)=Orders_Canon("От кого / площадка") Then
        Orders_MigrateSupplierHeader=True
        Exit Function
    End If

    ' Legacy header used in the user's working workbook.
    ' Only the header cell is renamed. All row values stay exactly where they are.
    If Orders_Canon(actual)=Orders_Canon("Поставщик / площадка") Then
        oSh.getCellByPosition(11,WMSORD_HEADER_ROW).String="От кого / площадка"
        Orders_ResetHeaderCache
        Orders_MigrateSupplierHeader=True
        Exit Function
    End If

    sErr="Несовместимая схема 'Заказы': колонка L имеет неизвестный заголовок '" & actual & _
         "'. Ожидалось 'Поставщик / площадка' или 'От кого / площадка'. Данные не изменены."
    Exit Function

EH:
    sErr="Миграция заголовка L: " & CStr(Err) & " " & Error$
End Function

' ============================================================================
' PREFLIGHT / SCHEMA
' ============================================================================

Function Orders_Preflight(oDoc As Object, ByRef sErr As String) As Boolean
    Dim oSh As Object, a As Variant, i As Long, actual As String, builder As String
    Orders_Preflight = False
    sErr = ""
    On Error GoTo EH

    If IsNull(oDoc) Or IsEmpty(oDoc) Then sErr = "Нет открытой книги Calc.": Exit Function
    If Not oDoc.supportsService("com.sun.star.sheet.SpreadsheetDocument") Then sErr = "Модуль Заказы работает только в LibreOffice Calc.": Exit Function
    If Not oDoc.Sheets.hasByName(WMSORD_SHEET) Then sErr = "Нет листа 'Заказы'. Сначала выполните WMS_00 CreateWorkbook.": Exit Function
    If Not oDoc.Sheets.hasByName(WMSORD_SETTINGS) Then sErr = "Нет листа 'Настройки WMS'. Сначала выполните WMS_00 CreateWorkbook.": Exit Function

    builder = Orders_GetSetting(oDoc, "WMS_BuilderVersion", "")
    If builder = "" Then sErr = "Книга не подтверждена как новая WMS, созданная WMS_00. Установка Заказы остановлена.": Exit Function

    oSh = oDoc.Sheets.getByName(WMSORD_SHEET)
    a = Orders_BaseHeaders()
    For i = 0 To UBound(a)
        actual = Trim(oSh.getCellByPosition(i, WMSORD_HEADER_ROW).String)

        ' Column L is compatible with both legacy and new universal-receipt wording.
        If i=11 Then
            If Orders_Canon(actual)<>Orders_Canon("От кого / площадка") And _
               Orders_Canon(actual)<>Orders_Canon("Поставщик / площадка") Then
                sErr="Несовместимая схема 'Заказы': колонка L должна называться 'Поставщик / площадка' " & _
                     "или 'От кого / площадка', сейчас '" & actual & "'. Данные не изменены."
                Exit Function
            End If
        ElseIf Orders_Canon(actual) <> Orders_Canon(CStr(a(i))) Then
            sErr = "Несовместимая схема 'Заказы': колонка " & Orders_ColName(i) & " должна называться '" & CStr(a(i)) & "', сейчас '" & actual & "'. Данные не изменены."
            Exit Function
        End If
    Next i

    Orders_Preflight = True
    Exit Function
EH:
    sErr = "Preflight Заказы: " & CStr(Err) & " " & Error$
End Function

Function Orders_BaseHeaders() As Variant
    Orders_BaseHeaders = Array("Номер", "Полное наименование товара", "№ документа", "Номер счета", "Код поставщика", "Артикул поставщика", "Факт. количество", "Количество", "Ед. изм.", "Цена", "Сумма", "От кого / площадка", "Дата поступления", "Дата документа", "Дата заказа", "Покупатель", "Статус", "Категория", "Контроль", "Место хранения", "Комментарий", "Код товара", "Продавец", "Источник прихода", "Подкатегория", "_WMS_LegacyDestination", _
        WMSORD_T_ID, WMSORD_T_STATE, WMSORD_T_HASH, WMSORD_T_ERROR, WMSORD_T_DIRTY, WMSORD_T_MODE, WMSORD_T_LEGACY, WMSORD_T_TOUCH)
End Function

Function Orders_ExtensionHeaders() As Variant
    Orders_ExtensionHeaders = Array(WMSORD_X_ROWVER, WMSORD_X_STATUSMODE, WMSORD_X_CODESUG, WMSORD_X_CODEORIGIN, WMSORD_X_CATSUG, WMSORD_X_CATORIGIN, WMSORD_X_LOCSUG, WMSORD_X_FLAGS, WMSORD_X_SEVERITY, WMSORD_X_VALIDATED, WMSORD_X_DUPKEY, WMSORD_X_DATEORIGIN, WMSORD_X_FINGERPRINT, WMSORD_X_SUGSOURCE, WMSORD_X_GROUPID, WMSORD_X_EFFINVOICE, WMSORD_X_EFFDOC, WMSORD_X_EFFDOCDATE, WMSORD_X_EFFSUPPLIER, WMSORD_X_GROUPFLAGS)
End Function

Function Orders_ExtensionHeaderCount() As Long
    Dim a As Variant: a = Orders_ExtensionHeaders(): Orders_ExtensionHeaderCount = UBound(a) + 1
End Function

Function Orders_EnsureExtensionColumns(oSh As Object, ByRef sErr As String) As Boolean
    Dim a As Variant, i As Long, c As Long, nextCol As Long, hdr As String
    Orders_EnsureExtensionColumns = False
    On Error GoTo EH
    a = Orders_ExtensionHeaders()
    nextCol = Orders_LastHeaderCol(oSh) + 1
    If nextCol < 31 Then nextCol = 31
    For i = 0 To UBound(a)
        hdr = CStr(a(i))
        c = Orders_FindHeader(oSh, hdr)
        If c < 0 Then
            oSh.getCellByPosition(nextCol, 0).String = hdr
            c = nextCol
            nextCol = nextCol + 1
        End If
        oSh.Columns.getByIndex(c).IsVisible = False
    Next i
    ' Keep every technical column hidden, never U/V.
    For c = 23 To Orders_LastHeaderCol(oSh)
        If Left(Trim(oSh.getCellByPosition(c,0).String),5) = "_WMS_" Then oSh.Columns.getByIndex(c).IsVisible = False
    Next c
    oSh.Columns.getByIndex(20).IsVisible = True
    oSh.Columns.getByIndex(21).IsVisible = True
    oSh.Columns.getByIndex(22).IsVisible = True
    Orders_EnsureExtensionColumns = True
    Exit Function
EH:
    sErr = "Не удалось обновить техническую схему Заказы: " & CStr(Err) & " " & Error$
End Function

' ============================================================================
' SETTINGS / LOGGING
' ============================================================================

Sub Orders_EnsureSettings(oDoc As Object)
    Orders_SetSetting oDoc, "WMS_OrdersVersion", WMSORD_VERSION, "Версия контроллера листа Заказы", True
    Orders_SetSetting oDoc, "ORDERS_AutoStatus", "Да", "Автоматически: В пути / Получено без документов / Получено. Оприходовано не ставится автоматически.", False
    Orders_SetSetting oDoc, "ORDERS_AutoReceiptDate", "Да", "При первом положительном факте заполнить дату поступления, если она пуста; дата остаётся редактируемой.", False
    Orders_SetSetting oDoc, "ORDERS_AutoCodeExact", "Да", "Автокод только по однозначному точному совпадению; ручной код не перезаписывать.", False
    Orders_SetSetting oDoc, "ORDERS_AutoCategoryExact", "Да", "Автокатегория только по однозначному точному совпадению; ручную категорию не перезаписывать.", False
    Orders_SetSetting oDoc, "ORDERS_LegacyAgeDays", CStr(WMSORD_DEFAULT_LEGACY_DAYS), "Возраст для первичной классификации старой строки HISTORY", False
    Orders_SetSetting oDoc, "ORDERS_Marketplaces", "Ozon|Wildberries|Яндекс Маркет", "Подтверждённые предыдущей WMS площадки, для которых полезен продавец", False
    Orders_SetSetting oDoc, "ORDERS_Statuses", WMSORD_STATUS_TRANSIT & "|" & WMSORD_STATUS_NO_DOCS & "|" & WMSORD_STATUS_RECEIVED & "|" & WMSORD_STATUS_POST & "|" & WMSORD_STATUS_CANCEL, "Допустимые статусы Заказы", False
    Orders_SetSetting oDoc, "ORDERS_GroupDocuments", "Да", "Группировка по колонке A Номер: каждое значение 1 начинает новый заказ; 2,3,4... продолжают его. Номер счета един на группу; № документа и дата документа относятся к конкретной строке и могут различаться.", False
    Orders_SetSetting oDoc, "ORDERS_ModuleSchema", WMSORD_SCHEMA, "Схема контроллера Заказы", True
End Sub

Sub Orders_SetSetting(oDoc As Object, sKey As String, sValue As String, sDesc As String, bOverwrite As Boolean)
    Dim oSh As Object, r As Long
    On Error GoTo Done
    If Not oDoc.Sheets.hasByName(WMSORD_SETTINGS) Then Exit Sub
    oSh = oDoc.Sheets.getByName(WMSORD_SETTINGS)
    r = Orders_FindSettingRow(oSh, sKey)
    If r < 1 Then r = Orders_NextRow(oSh, 0, 1): oSh.getCellByPosition(0,r).String = sKey
    If bOverwrite Or Trim(oSh.getCellByPosition(1,r).String) = "" Then oSh.getCellByPosition(1,r).String = sValue
    If Trim(oSh.getCellByPosition(2,r).String) = "" Or bOverwrite Then oSh.getCellByPosition(2,r).String = sDesc
Done:
End Sub

Function Orders_GetSetting(oDoc As Object, sKey As String, sDefault As String) As String
    Dim oSh As Object, r As Long
    Orders_GetSetting = sDefault
    On Error GoTo Done
    If Not oDoc.Sheets.hasByName(WMSORD_SETTINGS) Then Exit Function
    oSh = oDoc.Sheets.getByName(WMSORD_SETTINGS)
    r = Orders_FindSettingRow(oSh, sKey)
    If r >= 1 Then If Trim(oSh.getCellByPosition(1,r).String) <> "" Then Orders_GetSetting = Trim(oSh.getCellByPosition(1,r).String)
Done:
End Function

Function Orders_FindSettingRow(oSh As Object, sKey As String) As Long
    Dim last As Long, r As Long
    Orders_FindSettingRow = -1
    last = Orders_LastContentRow(oSh, 0)
    For r = 1 To last
        If Orders_Canon(oSh.getCellByPosition(0,r).String) = Orders_Canon(sKey) Then Orders_FindSettingRow = r: Exit Function
    Next r
End Function

Sub Orders_LogAudit(oDoc As Object, sEvent As String, sText As String)
    Dim oSh As Object, r As Long
    On Error GoTo Done
    If Not oDoc.Sheets.hasByName(WMSORD_AUDIT) Then Exit Sub
    oSh = oDoc.Sheets.getByName(WMSORD_AUDIT)
    r = Orders_NextRow(oSh, 0, 1)
    oSh.getCellByPosition(0,r).String = Orders_Stamp(Now)
    oSh.getCellByPosition(1,r).String = "ORD-" & Orders_TimestampCompact(Now)
    oSh.getCellByPosition(2,r).String = WMSORD_SHEET
    oSh.getCellByPosition(3,r).String = sEvent
    oSh.getCellByPosition(4,r).String = sText
Done:
End Sub

Sub Orders_LogError(oDoc As Object, sWhere As String, sText As String)
    Dim oSh As Object, r As Long
    On Error GoTo Done
    If Not oDoc.Sheets.hasByName(WMSORD_ERRORS) Then Exit Sub
    oSh = oDoc.Sheets.getByName(WMSORD_ERRORS)
    r = Orders_NextRow(oSh, 0, 1)
    oSh.getCellByPosition(0,r).String = Orders_Stamp(Now)
    oSh.getCellByPosition(1,r).String = "ORD-" & Orders_TimestampCompact(Now)
    oSh.getCellByPosition(2,r).String = WMSORD_SHEET & ":" & sWhere
    oSh.getCellByPosition(3,r).String = sText
    oSh.getCellByPosition(4,r).String = "OPEN"
Done:
End Sub


' ============================================================================
' ORDER GROUP VISUALIZATION
' ============================================================================

Function Orders_GroupColorIndex(sGroup As String) As Long
    Dim a As Variant, token As String, seq As Long, h As Double, i As Long, ch As Long
    Orders_GroupColorIndex = 0
    sGroup = Trim(sGroup)
    If sGroup = "" Then Exit Function
    ' Normal generated groups end with a monotonic group sequence. Using that
    ' sequence gives adjacent order blocks deterministic, alternating colors and
    ' keeps the color attached to the group after sorting.
    a = Split(sGroup, "|")
    If UBound(a) >= 1 Then
        token = Trim(CStr(a(UBound(a))))
        If IsNumeric(token) Then
            On Error Resume Next
            seq = CLng(token)
            If Err = 0 Then
                Orders_GroupColorIndex = Abs(seq) Mod 8
                On Error GoTo 0
                Exit Function
            End If
            Err = 0
            On Error GoTo 0
        End If
    End If
    ' Legacy/nonstandard group IDs use a stable hash fallback.
    h = 2166131
    For i = 1 To Len(sGroup)
        ch = Asc(Mid(sGroup, i, 1))
        h = h + ch * (i + 17)
        If h > 2147000000 Then h = h - 2147000000
    Next i
    Orders_GroupColorIndex = CLng(h Mod 8)
End Function

Sub Orders_RecolorAll()
    Dim oSh As Object
    If Not ThisComponent.Sheets.hasByName(WMSORD_SHEET) Then
        MsgBox "Нет листа Заказы.",48,"WMS — Заказы"
        Exit Sub
    End If
    oSh=ThisComponent.Sheets.getByName(WMSORD_SHEET)
    gWMSORD_Busy=True
    Orders_RecolorAllOrderGroups oSh
    gWMSORD_Busy=False
    MsgBox "Цветовые блоки заказов и проблемные ячейки обновлены. Формат данных не изменён.",64,"WMS — Заказы"
End Sub

Sub Orders_GetGroupPalette(nIdx As Long, ByRef nBack As Long)
    Select Case (nIdx Mod 8)
        Case 0: nBack = RGB(252,252,252)
        Case 1: nBack = RGB(242,247,252)
        Case 2: nBack = RGB(247,244,252)
        Case 3: nBack = RGB(244,250,246)
        Case 4: nBack = RGB(252,247,240)
        Case 5: nBack = RGB(244,249,251)
        Case 6: nBack = RGB(249,247,242)
        Case Else: nBack = RGB(247,247,249)
    End Select
End Sub

Sub Orders_ApplyOrderBlockBackground(oSh As Object, r As Long)
    Dim sGroup As String, nBack As Long, nIdx As Long, oRow As Object
    If r < 1 Then Exit Sub
    sGroup = Orders_CellText(oSh, r, WMSORD_X_GROUPID)
    If sGroup = "" Then
        nBack = RGB(255,255,255)
    Else
        nIdx = Orders_GroupColorIndex(sGroup)
        Orders_GetGroupPalette nIdx, nBack
    End If
    oRow = oSh.getCellRangeByPosition(0,r,WMSORD_USER_LAST_COL,r)
    oRow.CellBackColor = nBack
    ' Stronger separator when the group changes from the previous populated row.
    If r > 1 Then
        If sGroup <> Orders_CellText(oSh,r-1,WMSORD_X_GROUPID) And sGroup <> "" Then
            Dim aLine As New com.sun.star.table.BorderLine2
            aLine.Color = RGB(71,85,105)
            aLine.LineStyle = com.sun.star.table.BorderLineStyle.SOLID
            aLine.LineWidth = 70
            oRow.TopBorder2 = aLine
        End If
    End If
End Sub

Sub Orders_ClearProblemHighlight(oSh As Object, r As Long)
    Dim sGroup As String, nBack As Long, nIdx As Long
    sGroup = Orders_CellText(oSh,r,WMSORD_X_GROUPID)
    If sGroup = "" Then nBack = RGB(255,255,255) Else nIdx=Orders_GroupColorIndex(sGroup):Orders_GetGroupPalette nIdx,nBack
    Dim c As Long
    For c=0 To WMSORD_USER_LAST_COL
        oSh.getCellByPosition(c,r).CellBackColor = nBack
        oSh.getCellByPosition(c,r).CharColor = RGB(30,41,59)
        oSh.getCellByPosition(c,r).CharWeight = 100
    Next c
End Sub

Sub Orders_HighlightProblemCells(oSh As Object, r As Long)
    Dim flags As String, sev As String, nBack As Long, nFont As Long, nWeight As Double
    flags = Orders_CellText(oSh,r,WMSORD_X_FLAGS)
    sev = UCase(Orders_CellText(oSh,r,WMSORD_X_SEVERITY))
    If Trim(flags) = "" Or sev = "OK" Then Exit Sub
    If sev = "CRITICAL" Then
        nBack = RGB(244,199,195): nFont = RGB(128,20,20): nWeight = 150
    ElseIf sev = "WARN" Then
        nBack = RGB(252,232,178): nFont = RGB(120,72,0): nWeight = 150
    Else
        nBack = RGB(226,238,255): nFont = RGB(30,70,140): nWeight = 110
    End If
    Dim c As Long
    ' Map the first/all known issue keywords to the most useful input cells.
    If InStr(1,flags,"номер счета",1)>0 Or InStr(1,flags,"номера счета",1)>0 Then Orders_SetCellVisual oSh.getCellByPosition(3,r),nBack,nFont,nWeight
    If InStr(1,flags,"№ документа",1)>0 Or InStr(1,flags,"номер документа",1)>0 Then Orders_SetCellVisual oSh.getCellByPosition(2,r),nBack,nFont,nWeight
    If InStr(1,flags,"документа",1)>0 Then Orders_SetCellVisual oSh.getCellByPosition(2,r),nBack,nFont,nWeight
    If InStr(1,flags,"поставщик",1)>0 Or InStr(1,flags,"площадки",1)>0 Then Orders_SetCellVisual oSh.getCellByPosition(11,r),nBack,nFont,nWeight
    If InStr(1,flags,"наименование",1)>0 Then Orders_SetCellVisual oSh.getCellByPosition(1,r),nBack,nFont,nWeight
    If InStr(1,flags,"количество",1)>0 Or InStr(1,flags,"факт больше",1)>0 Then
        Orders_SetCellVisual oSh.getCellByPosition(6,r),nBack,nFont,nWeight
        Orders_SetCellVisual oSh.getCellByPosition(7,r),nBack,nFont,nWeight
    End If
    If InStr(1,flags,"дата",1)>0 Then
        Orders_SetCellVisual oSh.getCellByPosition(12,r),nBack,nFont,nWeight
        Orders_SetCellVisual oSh.getCellByPosition(13,r),nBack,nFont,nWeight
        Orders_SetCellVisual oSh.getCellByPosition(14,r),nBack,nFont,nWeight
    End If
    If InStr(1,flags,"единиц",1)>0 Then Orders_SetCellVisual oSh.getCellByPosition(8,r),nBack,nFont,nWeight
    If InStr(1,flags,"категори",1)>0 Then Orders_SetCellVisual oSh.getCellByPosition(17,r),nBack,nFont,nWeight
    If InStr(1,flags,"место хранения",1)>0 Then Orders_SetCellVisual oSh.getCellByPosition(19,r),nBack,nFont,nWeight
    If InStr(1,flags,"код товара",1)>0 Or InStr(1,flags,"кода товара",1)>0 Then Orders_SetCellVisual oSh.getCellByPosition(21,r),nBack,nFont,nWeight
    If InStr(1,flags,"продавц",1)>0 Then Orders_SetCellVisual oSh.getCellByPosition(22,r),nBack,nFont,nWeight
    If InStr(1,flags,"Номер должна",1)>0 Or InStr(1,flags,"нумерац",1)>0 Or InStr(1,flags,"позиции заказа",1)>0 Then Orders_SetCellVisual oSh.getCellByPosition(0,r),nBack,nFont,nWeight
    ' The Control cell is always the most prominent state indicator.
    Orders_SetCellVisual oSh.getCellByPosition(18,r),nBack,nFont,nWeight
End Sub

Sub Orders_RecolorAllOrderGroups(oSh As Object)
    Dim last As Long, r As Long
    last = Orders_LastContentRow(oSh,WMSORD_USER_LAST_COL)
    If last < 1 Then Exit Sub
    For r=1 To last
        Orders_ApplyRowVisual oSh,r
    Next r
End Sub

' ============================================================================
' DESIGN / UX
' ============================================================================

Sub Orders_ApplyDesign(oDoc As Object, oSh As Object)
    Dim last As Long, endRow As Long, c As Long
    Dim oRng As Object
    On Error GoTo Done

    oSh.TabColor = RGB(29,78,216)
    oSh.Rows.getByIndex(0).Height = 1000

    ' Header: one restrained professional navy bar.
    oRng = oSh.getCellRangeByPosition(0,0,WMSORD_USER_LAST_COL,0)
    oRng.CellBackColor = RGB(31,58,95)
    oRng.CharColor = RGB(255,255,255)
    oRng.CharWeight = 150
    oRng.CharHeight = 9.5
    oRng.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER
    oRng.VertJustify = com.sun.star.table.CellVertJustify.CENTER
    oRng.IsTextWrapped = True

    last = Orders_LastContentRow(oSh, WMSORD_USER_LAST_COL)
    endRow = last + 80
    If endRow < WMSORD_MIN_STYLE_ROWS Then endRow = WMSORD_MIN_STYLE_ROWS
    If endRow > oSh.Rows.getCount()-1 Then endRow = oSh.Rows.getCount()-1

    ' Base body.
    oRng = oSh.getCellRangeByPosition(0,1,WMSORD_USER_LAST_COL,endRow)
    oRng.CharColor = RGB(30,41,59)
    oRng.CharHeight = 9.5
    oRng.VertJustify = com.sun.star.table.CellVertJustify.CENTER

    ' Base body is intentionally neutral; each populated row receives a controlled order-group tint below.
    oRng.CellBackColor = RGB(255,255,255)

    ' Clear visual grid: every working cell must be visibly separated.
    Orders_ApplyTableGrid oSh, 0, 0, WMSORD_USER_LAST_COL, endRow

    ' Alignment and wrapping.
    Orders_SetAlignCenter oSh, Array(0,2,3,4,5,6,7,8,9,10,12,13,14,16,17,19,21), 1, endRow
    Orders_SetWrap oSh, Array(1,11,18,20,22,23,24,25), 1, endRow

    ' Widths are tuned for daily warehouse use, not presentation slides.
    Orders_SetWidth oSh, 0, 1800
    Orders_SetWidth oSh, 1, 6500
    Orders_SetWidth oSh, 2, 2700
    Orders_SetWidth oSh, 3, 2600
    Orders_SetWidth oSh, 4, 2400
    Orders_SetWidth oSh, 5, 3000
    Orders_SetWidth oSh, 6, 1900
    Orders_SetWidth oSh, 7, 1900
    Orders_SetWidth oSh, 8, 1500
    Orders_SetWidth oSh, 9, 1900
    Orders_SetWidth oSh,10, 2200
    Orders_SetWidth oSh,11, 3900
    Orders_SetWidth oSh,12, 2300
    Orders_SetWidth oSh,13, 2300
    Orders_SetWidth oSh,14, 2300
    Orders_SetWidth oSh,15, 3000
    Orders_SetWidth oSh,16, 3300
    Orders_SetWidth oSh,17, 2900
    Orders_SetWidth oSh,18, 5200
    Orders_SetWidth oSh,19, 3000
    Orders_SetWidth oSh,20, 4700
    Orders_SetWidth oSh,21, 2500
    Orders_SetWidth oSh,22, 3600
    Orders_SetWidth oSh,23, 3900
    Orders_SetWidth oSh,24, 3600
    Orders_SetWidth oSh,25, 3900

    ' Formats.
    Orders_ApplyNumberFormat oDoc, oSh.getCellRangeByPosition(6,1,7,endRow), "QTY"
    Orders_ApplyNumberFormat oDoc, oSh.getCellRangeByPosition(9,1,10,endRow), "MONEY"
    Orders_ApplyNumberFormat oDoc, oSh.getCellRangeByPosition(12,1,14,endRow), "DATE"
    Orders_ApplyNumberFormat oDoc, oSh.getCellRangeByPosition(0,1,5,endRow), "TEXT"
    Orders_ApplyNumberFormat oDoc, oSh.getCellRangeByPosition(21,1,25,endRow), "TEXT"

    ' Hybrid/automatic fields visually stronger.
    oSh.getCellRangeByPosition(16,1,16,endRow).CharWeight = 110
    oSh.getCellRangeByPosition(18,1,18,endRow).CharWeight = 110
    oSh.getCellRangeByPosition(21,1,21,endRow).CharWeight = 110

    ' Rebuild order block visuals from the real hidden group identity, so pasted foreign formatting is always corrected.
    If last >= 1 Then Orders_RecolorAllOrderGroups oSh

    ' Hide every technical field after user column W.
    For c = 23 To Orders_LastHeaderCol(oSh)
        If Left(Trim(oSh.getCellByPosition(c,0).String),5) = "_WMS_" Then oSh.Columns.getByIndex(c).IsVisible = False
    Next c
    oSh.Columns.getByIndex(20).IsVisible = True
    oSh.Columns.getByIndex(21).IsVisible = True
    oSh.Columns.getByIndex(22).IsVisible = True
    oSh.Columns.getByIndex(23).IsVisible = True
    oSh.Columns.getByIndex(24).IsVisible = True
    oSh.Columns.getByIndex(25).IsVisible = False
Done:
End Sub

Sub Orders_SetColumnBand(oSh As Object, aCols As Variant, r1 As Long, r2 As Long, nColor As Long)
    Dim i As Long, c As Long
    For i = LBound(aCols) To UBound(aCols)
        c = CLng(aCols(i))
        oSh.getCellRangeByPosition(c,r1,c,r2).CellBackColor = nColor
    Next i
End Sub

Sub Orders_ApplyTableGrid(oSh As Object, c1 As Long, r1 As Long, c2 As Long, r2 As Long)
    Dim oRange As Object, oTop As Object, oLeft As Object
    Dim aLine As New com.sun.star.table.BorderLine2
    Dim aOuter As New com.sun.star.table.BorderLine2
    On Error GoTo Done

    oRange = oSh.getCellRangeByPosition(c1,r1,c2,r2)

    ' IMPORTANT: Calc does not reliably materialize TableBorder2.HorizontalLine /
    ' VerticalLine on every cell of a large range. Apply BottomBorder2 and
    ' RightBorder2 directly to the range: Calc then writes the separator to every
    ' individual cell, so the table remains visibly separated on screen and print.
    aOuter.Color = RGB(100,116,139)
    aOuter.LineStyle = com.sun.star.table.BorderLineStyle.SOLID
    aOuter.LineWidth = 50
    ' Apply outer top/left FIRST. Calc range-border setters may reset other sides
    ' of the touched cells; the per-cell grid is therefore applied LAST.
    oTop = oSh.getCellRangeByPosition(c1,r1,c2,r1): oTop.TopBorder2 = aOuter
    oLeft = oSh.getCellRangeByPosition(c1,r1,c1,r2): oLeft.LeftBorder2 = aOuter

    aLine.Color = RGB(176,186,198)
    aLine.LineStyle = com.sun.star.table.BorderLineStyle.SOLID
    aLine.LineWidth = 35
    oRange.BottomBorder2 = aLine
    oRange.RightBorder2 = aLine
Done:
End Sub

Sub Orders_SetAlignCenter(oSh As Object, aCols As Variant, r1 As Long, r2 As Long)
    Dim i As Long, c As Long
    For i = LBound(aCols) To UBound(aCols)
        c = CLng(aCols(i))
        oSh.getCellRangeByPosition(c,r1,c,r2).HoriJustify = com.sun.star.table.CellHoriJustify.CENTER
    Next i
End Sub

Sub Orders_SetWrap(oSh As Object, aCols As Variant, r1 As Long, r2 As Long)
    Dim i As Long, c As Long
    For i = LBound(aCols) To UBound(aCols)
        c = CLng(aCols(i))
        oSh.getCellRangeByPosition(c,r1,c,r2).IsTextWrapped = True
    Next i
End Sub

Sub Orders_SetWidth(oSh As Object, c As Long, nWidth As Long)
    On Error Resume Next
    oSh.Columns.getByIndex(c).Width = nWidth
    On Error GoTo 0
End Sub

Sub Orders_ApplyNumberFormat(oDoc As Object, oRange As Object, sKind As String)
    Dim k As Long
    k = Orders_NumberFormatKey(oDoc, sKind)
    If k >= 0 Then oRange.NumberFormat = k
End Sub

Function Orders_NumberFormatKey(oDoc As Object, sKind As String) As Long
    Dim oFormats As Object
    Dim aLocale As New com.sun.star.lang.Locale
    Dim fmt As String, k As Long, baseKey As Long
    Orders_NumberFormatKey = -1
    On Error GoTo Done

    oFormats = oDoc.NumberFormats
    aLocale.Language = "ru": aLocale.Country = "RU"

    Select Case UCase(sKind)
        Case "QTY"
            ' Locale-safe General format: 1 -> 1, 1.5 -> 1,5.
            ' Do not hardcode dot/comma separators: they differ by locale.
            k = oFormats.getStandardIndex(aLocale)
            Orders_NumberFormatKey = k
            Exit Function

        Case "MONEY"
            ' Generate the format from LibreOffice locale data instead of parsing
            ' an English pattern under ru-RU (which caused 456 -> 456.000.00).
            baseKey = oFormats.getStandardFormat(com.sun.star.util.NumberFormat.NUMBER, aLocale)
            fmt = oFormats.generateFormat(baseKey, aLocale, False, False, 2, 1)

        Case "DATE"
            fmt = "DD.MM.YYYY"

        Case "TEXT"
            fmt = "@"

        Case Else
            Exit Function
    End Select

    k = oFormats.queryKey(fmt, aLocale, False)
    If k = -1 Then k = oFormats.addNew(fmt, aLocale)
    Orders_NumberFormatKey = k
Done:
End Function

Sub Orders_ApplyValidations(oDoc As Object, oSh As Object)
    Dim last As Long, endRow As Long, oV As Object, oRange As Object
    On Error GoTo Done
    last = Orders_LastContentRow(oSh, WMSORD_USER_LAST_COL)
    endRow = last + 200
    If endRow < WMSORD_MIN_STYLE_ROWS Then endRow = WMSORD_MIN_STYLE_ROWS
    If endRow > oSh.Rows.getCount()-1 Then endRow = oSh.Rows.getCount()-1

    ' Status: strict drop-down. Manual override is allowed, but only valid statuses.
    oRange = oSh.getCellRangeByPosition(16,1,16,endRow)
    oV = oRange.Validation
    oV.Type = com.sun.star.sheet.ValidationType.LIST
    oV.Operator = com.sun.star.sheet.ConditionOperator.EQUAL
    oV.Formula1 = """В пути"";""Получено без документов"";""Получено"";""Оприходовано"";""Отменено"""
    oV.IgnoreBlankCells = True
    oV.ShowList = 1
    oV.ShowErrorMessage = True
    oV.ErrorAlertStyle = com.sun.star.sheet.ValidationAlertStyle.STOP
    oV.ErrorTitle = "WMS — статус заказа"
    oV.ErrorMessage = "Выберите статус из списка. 'Оприходовано' — только осознанно вручную."
    oV.ShowInputMessage = True
    oV.InputTitle = "Статус"
    oV.InputMessage = "Автоматика ставит В пути / Получено без документов / Получено. Ручной выбор фиксируется. Чтобы вернуть авто-режим — очистите статус."
    oRange.Validation = oV

    Orders_SetNumberValidation oSh.getCellRangeByPosition(0,1,0,endRow), True, "Номер позиции заказа", "Введите целое число: 1 для первой позиции нового заказа, затем 2, 3, 4... Для следующего заказа снова начните с 1."
    Orders_SetInputHelp oSh.getCellRangeByPosition(0,1,0,endRow), "Номер позиции заказа", "Каждый новый заказ начинается с 1. Последующие позиции этого же заказа: 2, 3, 4... Новый заказ — снова 1. WMS использует эту нумерацию как жёсткую границу группы."
    Orders_SetNumberValidation oSh.getCellRangeByPosition(6,1,6,endRow), False, "Факт. количество", "Факт должен быть числом не меньше нуля."
    Orders_SetNumberValidation oSh.getCellRangeByPosition(7,1,7,endRow), True, "Количество", "Заказанное количество должно быть положительным числом."
    Orders_SetNumberValidation oSh.getCellRangeByPosition(9,1,10,endRow), False, "Цена / сумма", "Значение должно быть числом не меньше нуля."

    Orders_SetInputHelp oSh.getCellRangeByPosition(12,1,12,endRow), "Дата поступления", "При первом положительном факте WMS заполнит сегодняшнюю дату, если поле пусто. Можно исправить вручную. 24.08, 24/08 и 24-08 нормализуются с текущим годом."
    Orders_SetInputHelp oSh.getCellRangeByPosition(17,1,17,endRow), "Категория", "WMS подставляет категорию только при однозначном точном совпадении. Любая ваша ручная категория имеет приоритет. Очистите ячейку, чтобы снова разрешить авто-подстановку."
    Orders_SetInputHelp oSh.getCellRangeByPosition(18,1,18,endRow), "Контроль", "Полностью автоматическое поле. Здесь показывается одно главное действие/проблема. Полная диагностика хранится в скрытых техполях."
    Orders_SetInputHelp oSh.getCellRangeByPosition(21,1,21,endRow), "Код товара", "Введите внутренний код. Если код уже есть в Firebird, полное наименование в колонке B заполнится автоматически."

    ' Универсальный приход.
    oRange=oSh.getCellRangeByPosition(23,1,23,endRow)
    oV=oRange.Validation
    oV.Type=com.sun.star.sheet.ValidationType.LIST
    oV.Operator=com.sun.star.sheet.ConditionOperator.EQUAL
    oV.Formula1="""Поставщик / заказ"";""Поставщик / прямой приход"";""Производство"";""Офис"";""Другое"""
    oV.IgnoreBlankCells=True
    oV.ShowList=1
    oV.ShowErrorMessage=True
    oV.ErrorAlertStyle=com.sun.star.sheet.ValidationAlertStyle.STOP
    oV.ErrorTitle="WMS — источник прихода"
    oV.ErrorMessage="Выберите источник из списка."
    oRange.Validation=oV

    Orders_SetInputHelp oSh.getCellRangeByPosition(23,1,23,endRow), "Источник прихода", "Откуда эта конкретная партия поступила на склад."
    Orders_SetInputHelp oSh.getCellRangeByPosition(24,1,24,endRow), "Подкатегория", "Подкатегория относится к обычной категории товара в колонке R. Например: Категория Крепёж → Подкатегория Болты."
    ' Z с 1.0.19 — скрытое legacy-поле; пользователь его не заполняет.
Done:
End Sub

Sub Orders_SetNumberValidation(oRange As Object, bStrictPositive As Boolean, sTitle As String, sMessage As String)
    Dim oV As Object
    On Error GoTo Done
    oV = oRange.Validation
    oV.Type = com.sun.star.sheet.ValidationType.DECIMAL
    If bStrictPositive Then
        oV.Operator = com.sun.star.sheet.ConditionOperator.GREATER
        oV.Formula1 = "0"
    Else
        oV.Operator = com.sun.star.sheet.ConditionOperator.GREATER_EQUAL
        oV.Formula1 = "0"
    End If
    oV.IgnoreBlankCells = True
    oV.ShowErrorMessage = True
    oV.ErrorAlertStyle = com.sun.star.sheet.ValidationAlertStyle.WARNING
    oV.ErrorTitle = "WMS — " & sTitle
    oV.ErrorMessage = sMessage
    oRange.Validation = oV
Done:
End Sub

Sub Orders_SetInputHelp(oRange As Object, sTitle As String, sMessage As String)
    Dim oV As Object
    On Error GoTo Done
    oV = oRange.Validation
    If oV.Type = com.sun.star.sheet.ValidationType.ANY Then oV.Type = com.sun.star.sheet.ValidationType.ANY
    oV.ShowInputMessage = True
    oV.InputTitle = sTitle
    oV.InputMessage = sMessage
    oRange.Validation = oV
Done:
End Sub

Sub Orders_ApplyFreeze(oDoc As Object, oSh As Object)
    Dim oCtl As Object
    On Error GoTo Done
    oCtl = oDoc.CurrentController
    If IsNull(oCtl) Or IsEmpty(oCtl) Then Exit Sub
    oCtl.setActiveSheet oSh
    oCtl.freezeAtPosition 2, 1
    On Error Resume Next
    oCtl.ZoomValue = 85
    On Error GoTo 0
Done:
End Sub

Sub Orders_RefreshFilter(oDoc As Object, oSh As Object)
    Dim dbs As Object, db As Object, addr As Variant, last As Long, endRow As Long, dbName As String, lastFilterCol As Long
    On Error GoTo Done
    last = Orders_LastContentRow(oSh, WMSORD_USER_LAST_COL)
    endRow = last + WMSORD_FILTER_BUFFER
    If endRow < WMSORD_MIN_STYLE_ROWS Then endRow = WMSORD_MIN_STYLE_ROWS
    If endRow > oSh.Rows.getCount()-1 Then endRow = oSh.Rows.getCount()-1
    ' IMPORTANT: the database range includes hidden technical columns too.
    ' This makes native AutoFilter sorting move SourceID/fingerprint/audit metadata
    ' together with the user's business row instead of detaching identities.
    lastFilterCol = Orders_LastHeaderCol(oSh)
    If lastFilterCol < WMSORD_USER_LAST_COL Then lastFilterCol = WMSORD_USER_LAST_COL
    addr = oSh.getCellRangeByPosition(0,0,lastFilterCol,endRow).RangeAddress
    dbs = oDoc.DatabaseRanges
    dbName = "WMS00_DB_02"
    If dbs.hasByName(dbName) Then
        db = dbs.getByName(dbName)
        db.setDataArea addr
    Else
        dbs.addNewByName dbName, addr
        db = dbs.getByName(dbName)
    End If
    db.AutoFilter = True
Done:
End Sub

' ============================================================================
' SHEET EVENT — LIGHT ROW CONTROLLER
' ============================================================================

Sub Orders_EnsureChangedRangeFormat(oDoc As Object,oSh As Object,r1 As Long,r2 As Long)
    On Error GoTo Done
    If r1<1 Then r1=1
    If r2<r1 Then Exit Sub

    Orders_ApplyNumberFormat oDoc,oSh.getCellRangeByPosition(6,r1,7,r2),"QTY"
    Orders_ApplyNumberFormat oDoc,oSh.getCellRangeByPosition(9,r1,10,r2),"MONEY"
    Orders_ApplyNumberFormat oDoc,oSh.getCellRangeByPosition(12,r1,14,r2),"DATE"
    Orders_ApplyNumberFormat oDoc,oSh.getCellRangeByPosition(0,r1,5,r2),"TEXT"
    Orders_ApplyNumberFormat oDoc,oSh.getCellRangeByPosition(21,r1,22,r2),"TEXT"
    Orders_ApplyWorkingRowGrid oSh,0,r1,WMSORD_USER_LAST_COL,r2
Done:
End Sub

Sub Orders_ApplyWorkingRowGrid(oSh As Object,c1 As Long,r1 As Long,c2 As Long,r2 As Long)
    Dim oRange As Object
    Dim aLine As New com.sun.star.table.BorderLine2
    On Error GoTo Done
    oRange=oSh.getCellRangeByPosition(c1,r1,c2,r2)
    aLine.Color=RGB(176,186,198)
    aLine.LineStyle=com.sun.star.table.BorderLineStyle.SOLID
    aLine.LineWidth=35
    oRange.BottomBorder2=aLine
    oRange.RightBorder2=aLine
Done:
End Sub

Function Orders_InstallEvent(oSh As Object, ByRef sErr As String) As Boolean
    Dim ev(1) As New com.sun.star.beans.PropertyValue
    Orders_InstallEvent = False
    On Error GoTo EH
    ev(0).Name = "EventType": ev(0).Value = "Script"
    ev(1).Name = "Script": ev(1).Value = "vnd.sun.star.script:Standard.WMS_02_Orders_FINAL.Orders_OnContentChanged?language=Basic&location=document"
    oSh.Events.replaceByName "OnChange", ev()
    Orders_InstallEvent = True
    Exit Function
EH:
    sErr = "Не удалось назначить событие Content changed для Заказы: " & CStr(Err) & " " & Error$
End Function


' ============================================================================
' ORDER GROUP CONTEXT — one order can contain many item rows.
' Column A numbering is the hard group boundary. Invoice/document/document date
' and supplier/platform may be entered only on the first row of the order; all
' following 2,3,4... rows inherit hidden effective values without filling visible
' cells. The hidden group survives sorting/filtering.
' ============================================================================

Function Orders_GroupRelevantChange(c1 As Long,c2 As Long) As Boolean
    ' A number, C document, D invoice, L supplier/platform, N document date,
    ' O order date and P buyer may change the logical group context.
    Orders_GroupRelevantChange = ((c1<=0 And c2>=0) Or (c1<=2 And c2>=2) Or (c1<=3 And c2>=3) Or (c1<=11 And c2>=11) Or (c1<=13 And c2>=13) Or (c1<=14 And c2>=14) Or (c1<=15 And c2>=15))
End Function

Function Orders_OrderPositionState(oSh As Object,r As Long,ByRef nPos As Long) As Integer
    Dim d As Double,s As String
    Orders_OrderPositionState=0:nPos=0
    If r<1 Then Exit Function
    s=Trim(oSh.getCellByPosition(0,r).String)
    If s="" And Abs(oSh.getCellByPosition(0,r).Value)<=WMSORD_EPS Then Exit Function
    If Not Orders_TryNumber(oSh.getCellByPosition(0,r),d) Then Orders_OrderPositionState=-1:Exit Function
    If d<1 Or Abs(d-Int(d))>WMSORD_EPS Then Orders_OrderPositionState=-1:Exit Function
    nPos=CLng(d):Orders_OrderPositionState=1
End Function

Function Orders_BulkOrderPositionState(v As Variant,ByRef nPos As Long) As Integer
    Dim d As Double,s As String
    Orders_BulkOrderPositionState=0:nPos=0
    s=Orders_BulkText(v):If s="" Then Exit Function
    If Not Orders_BulkTryNumber(v,d) Then Orders_BulkOrderPositionState=-1:Exit Function
    If d<1 Or Abs(d-Int(d))>WMSORD_EPS Then Orders_BulkOrderPositionState=-1:Exit Function
    nPos=CLng(d):Orders_BulkOrderPositionState=1
End Function

Function Orders_NewOrderGroupID(r As Long) As String
    Dim tick As Long
    gWMSORD_GroupSeq=gWMSORD_GroupSeq+1
    If gWMSORD_GroupSeq>999999 Then gWMSORD_GroupSeq=1
    tick=CLng(Timer*100)
    Orders_NewOrderGroupID="ORDSEQ|" & Orders_TimestampCompact(Now) & "|" & Right("000000" & CStr(r+1),6) & "|" & Right("00000000" & CStr(tick),8) & "|" & Right("000000" & CStr(gWMSORD_GroupSeq),6)
End Function

Function Orders_MakeGroupKey(oSh As Object,r As Long) As String
    ' Compatibility fallback for rows where user numbering in A is absent.
    ' Numbered rows use ORDSEQ groups instead.
    Dim invoice As String,platform As String,buyer As String,dOrder As Double
    invoice=Trim(oSh.getCellByPosition(3,r).String)
    If invoice="" Then Exit Function
    platform=Trim(oSh.getCellByPosition(11,r).String)
    buyer=Trim(oSh.getCellByPosition(15,r).String)
    dOrder=Orders_CellDate(oSh.getCellByPosition(14,r),True)
    Orders_MakeGroupKey="ORD|" & Orders_Canon(platform) & "|" & Orders_Canon(invoice) & "|" & CStr(CLng(dOrder)) & "|" & Orders_Canon(buyer)
End Function

Function Orders_SameSupplierContext(oSh As Object,r1 As Long,r2 As Long) As Boolean
    Dim a As String,b As String
    a=Orders_Canon(Trim(oSh.getCellByPosition(11,r1).String))
    b=Orders_Canon(Trim(oSh.getCellByPosition(11,r2).String))
    If a="" Or b="" Then Orders_SameSupplierContext=True Else Orders_SameSupplierContext=(a=b)
End Function

Sub Orders_ClearGroupContextRow(oSh As Object,r As Long)
    Dim c As Long
    c=Orders_FindHeader(oSh,WMSORD_X_GROUPID):If c>=0 Then oSh.getCellByPosition(c,r).String=""
    c=Orders_FindHeader(oSh,WMSORD_X_EFFINVOICE):If c>=0 Then oSh.getCellByPosition(c,r).String=""
    c=Orders_FindHeader(oSh,WMSORD_X_EFFDOC):If c>=0 Then oSh.getCellByPosition(c,r).String=""
    c=Orders_FindHeader(oSh,WMSORD_X_EFFDOCDATE):If c>=0 Then oSh.getCellByPosition(c,r).Value=0
    c=Orders_FindHeader(oSh,WMSORD_X_EFFSUPPLIER):If c>=0 Then oSh.getCellByPosition(c,r).String=""
    c=Orders_FindHeader(oSh,WMSORD_X_GROUPFLAGS):If c>=0 Then oSh.getCellByPosition(c,r).String=""
End Sub

Sub Orders_PrepareGroupEdits(oSh As Object,r1 As Long,r2 As Long,c1 As Long,c2 As Long)
    Dim last As Long,stopRow As Long,i As Long,nPos As Long,posState As Integer
    Dim aNums As Variant,rowv As Variant,c As Long
    Dim aHdr As Variant,h As Variant
    ' Group identity is driven by column A. For an A edit/paste, invalidate the
    ' touched sequence and let Orders_RebuildGroupContexts assign it again.
    ' IMPORTANT: do this in vector/range operations, not row-by-row UNO calls,
    ' so a 1000-row paste stays linear and fast.
    If Not (c1<=0 And c2>=0) Then Exit Sub
    last=Orders_LastContentRow(oSh,WMSORD_USER_LAST_COL)
    If last<1 Then Exit Sub
    If r1<1 Then r1=1
    If r2<r1 Then r2=r1
    stopRow=last

    ' Find the next untouched row that explicitly starts a new order (A=1).
    If r2<last Then
        aNums=oSh.getCellRangeByPosition(0,r2+1,0,last).DataArray
        For i=0 To UBound(aNums)
            rowv=aNums(i)
            posState=Orders_BulkOrderPositionState(rowv(0),nPos)
            If posState=1 And nPos=1 Then stopRow=r2+i:Exit For
        Next i
    End If
    If stopRow<r1 Then Exit Sub

    aHdr=Array(WMSORD_X_GROUPID,WMSORD_X_EFFINVOICE,WMSORD_X_EFFDOC,WMSORD_X_EFFDOCDATE,WMSORD_X_EFFSUPPLIER,WMSORD_X_GROUPFLAGS)
    On Error Resume Next
    For Each h In aHdr
        c=Orders_FindHeader(oSh,CStr(h))
        If c>=0 Then oSh.getCellRangeByPosition(c,r1,c,stopRow).clearContents(1023)
    Next h
    On Error GoTo 0
End Sub

Function Orders_BulkGroupKey(aRow As Variant) As String
    ' Compatibility fallback only for unnumbered legacy rows.
    Dim invoice As String,platform As String,buyer As String,dOrder As Double
    invoice=Orders_BulkText(aRow(3))
    If invoice="" Then Exit Function
    platform=Orders_BulkText(aRow(11)):buyer=Orders_BulkText(aRow(15))
    dOrder=Orders_BulkDateValue(aRow(14),True)
    Orders_BulkGroupKey="ORD|" & Orders_Canon(platform) & "|" & Orders_Canon(invoice) & "|" & CStr(CLng(dOrder)) & "|" & Orders_Canon(buyer)
End Function

Function Orders_GroupHashSlot(aKeys() As String,nCap As Long,sKey As String,ByRef bFound As Boolean) As Long
    Dim idx As Long,probe As Long,k As String
    Orders_GroupHashSlot=-1:bFound=False
    If nCap<=0 Or sKey="" Then Exit Function
    idx=Orders_QueueHashBase(sKey,nCap)
    For probe=0 To nCap-1
        k=aKeys(idx)
        If k="" Then Orders_GroupHashSlot=idx:Exit Function
        If k=sKey Then bFound=True:Orders_GroupHashSlot=idx:Exit Function
        idx=idx+1:If idx>=nCap Then idx=0
    Next probe
End Function

Sub Orders_GroupHashTextPut(aKeys() As String,aVals() As String,aFlags() As String,nCap As Long,sGroup As String,sValue As String,sConflict As String)
    Dim slot As Long,bFound As Boolean,old As String
    If sGroup="" Or Trim(sValue)="" Then Exit Sub
    slot=Orders_GroupHashSlot(aKeys(),nCap,sGroup,bFound):If slot<0 Then Exit Sub
    If Not bFound Then aKeys(slot)=sGroup
    old=aVals(slot)
    If old="" Then
        aVals(slot)=sValue
    ElseIf Orders_Canon(old)<>Orders_Canon(sValue) Then
        Orders_GroupHashFlagPut aFlags(),slot,sConflict
    End If
End Sub

Sub Orders_GroupHashDatePut(aKeys() As String,aVals() As Double,aFlags() As String,nCap As Long,sGroup As String,dValue As Double,sConflict As String)
    Dim slot As Long,bFound As Boolean,old As Double
    If sGroup="" Or dValue<=0 Then Exit Sub
    slot=Orders_GroupHashSlot(aKeys(),nCap,sGroup,bFound):If slot<0 Then Exit Sub
    If Not bFound Then aKeys(slot)=sGroup
    old=aVals(slot)
    If old<=0 Then
        aVals(slot)=dValue
    ElseIf Abs(old-dValue)>WMSORD_EPS Then
        Orders_GroupHashFlagPut aFlags(),slot,sConflict
    End If
End Sub

Sub Orders_GroupHashFlagPut(aFlags() As String,nSlot As Long,sText As String)
    If nSlot<0 Then Exit Sub
    If aFlags(nSlot)="" Then
        aFlags(nSlot)=sText
    ElseIf InStr(1,aFlags(nSlot),sText,1)=0 Then
        aFlags(nSlot)=aFlags(nSlot) & " | " & sText
    End If
End Sub

Sub Orders_RebuildGroupContexts(oSh As Object)
    Dim cGroupAbs As Long,cEffInvAbs As Long,cEffDocAbs As Long,cEffDateAbs As Long,cEffSupplierAbs As Long,cFlagsAbs As Long
    Dim techStart As Long,techEnd As Long,cGroup As Long,cEffInv As Long,cEffDoc As Long,cEffDate As Long,cEffSupplier As Long,cFlags As Long
    Dim last As Long,i As Long,groupID As String,prevGroup As String,invoice As String,docNo As String,supplier As String,prevSupplier As String
    Dim dDoc As Double,desired As String,aData As Variant,aTech As Variant,drow As Variant,trow As Variant
    Dim cap As Long,slot As Long,bFound As Boolean,startSlot As Long,startFound As Boolean
    Dim keys() As String,invVals() As String,docVals() As String,supplierVals() As String,dateVals() As Double,flagVals() As String,startKeys() As String,seqFlags() As String
    Dim docMixed() As Boolean,dateMixed() As Boolean
    Dim pos As Long,prevPos As Long,posState As Integer,prevPosState As Integer,sSeq As String
    On Error GoTo EH
    cGroupAbs=Orders_FindHeader(oSh,WMSORD_X_GROUPID):cEffInvAbs=Orders_FindHeader(oSh,WMSORD_X_EFFINVOICE)
    cEffDocAbs=Orders_FindHeader(oSh,WMSORD_X_EFFDOC):cEffDateAbs=Orders_FindHeader(oSh,WMSORD_X_EFFDOCDATE)
    cEffSupplierAbs=Orders_FindHeader(oSh,WMSORD_X_EFFSUPPLIER):cFlagsAbs=Orders_FindHeader(oSh,WMSORD_X_GROUPFLAGS)
    If cGroupAbs<0 Or cEffInvAbs<0 Or cEffDocAbs<0 Or cEffDateAbs<0 Or cEffSupplierAbs<0 Or cFlagsAbs<0 Then Exit Sub
    last=Orders_LastContentRow(oSh,WMSORD_USER_LAST_COL)
    If last<1 Then Exit Sub
    techStart=23:techEnd=Orders_LastHeaderCol(oSh)
    cGroup=cGroupAbs-techStart:cEffInv=cEffInvAbs-techStart:cEffDoc=cEffDocAbs-techStart:cEffDate=cEffDateAbs-techStart:cEffSupplier=cEffSupplierAbs-techStart:cFlags=cFlagsAbs-techStart
    aData=oSh.getCellRangeByPosition(0,1,WMSORD_USER_LAST_COL,last).DataArray
    aTech=oSh.getCellRangeByPosition(techStart,1,techEnd,last).DataArray
    cap=64:Do While cap<(last+8)*2:cap=cap*2:Loop
    ReDim keys(0 To cap-1):ReDim invVals(0 To cap-1):ReDim docVals(0 To cap-1):ReDim supplierVals(0 To cap-1):ReDim dateVals(0 To cap-1):ReDim flagVals(0 To cap-1)
    ReDim docMixed(0 To cap-1):ReDim dateMixed(0 To cap-1)
    ReDim startKeys(0 To cap-1):ReDim seqFlags(0 To UBound(aData))

    ' Pass 1. Column A is the hard order boundary:
    '   1 = first position of a NEW order
    '   2,3,4... = continuation of that order
    ' Existing hidden GroupID is trusted after sorting. Blank/new contexts are
    ' assigned from the sequence. A duplicated GroupID on a second "1" is healed.
    prevGroup="":prevSupplier="":prevPos=0:prevPosState=0
    For i=0 To UBound(aData)
        drow=aData(i):trow=aTech(i):seqFlags(i)=""
        If Not Orders_BulkRowHasData(drow) Then
            prevGroup="":prevSupplier="":prevPos=0:prevPosState=0
            GoTo Next1
        End If
        supplier=Orders_Canon(Orders_BulkText(drow(11)))
        groupID=Trim(CStr(trow(cGroup)))
        invoice=Orders_BulkText(drow(3))
        posState=Orders_BulkOrderPositionState(drow(0),pos)

        If posState=1 Then
            If pos=1 Then
                If groupID="" Then
                    groupID=Orders_NewOrderGroupID(i+1)
                Else
                    startSlot=Orders_GroupHashSlot(startKeys(),cap,groupID,startFound)
                    If startFound Or (prevGroup<>"" And groupID=prevGroup) Then groupID=Orders_NewOrderGroupID(i+1)
                End If
                startSlot=Orders_GroupHashSlot(startKeys(),cap,groupID,startFound)
                If startSlot>=0 And Not startFound Then startKeys(startSlot)=groupID
            ElseIf groupID="" Then
                If prevGroup<>"" And prevPosState=1 And prevPos+1=pos Then
                    groupID=prevGroup
                Else
                    groupID=Orders_NewOrderGroupID(i+1)
                    seqFlags(i)="Нарушена нумерация заказа: позиция " & CStr(pos) & " должна идти после " & CStr(pos-1) & " или новый заказ должен начинаться с 1"
                End If
            End If
        ElseIf posState<0 Then
            If groupID="" Then groupID=Orders_NewOrderGroupID(i+1)
            seqFlags(i)="Колонка Номер должна содержать целое положительное число: 1, 2, 3..."
        Else
            ' Compatibility fallback for old/legacy rows without numbering.
            If groupID="" Then
                If invoice<>"" Then
                    desired=Orders_BulkGroupKey(drow):groupID=desired
                ElseIf prevGroup<>"" Then
                    If supplier="" Or prevSupplier="" Or supplier=prevSupplier Then groupID=prevGroup
                End If
            End If
        End If

        trow(cGroup)=groupID
        If groupID<>"" Then
            ' Ensure a hash slot exists even when this row has no invoice/supplier.
            slot=Orders_GroupHashSlot(keys(),cap,groupID,bFound)
            If slot>=0 And Not bFound Then keys(slot)=groupID
            If invoice<>"" Then Orders_GroupHashTextPut keys(),invVals(),flagVals(),cap,groupID,invoice,"В одном заказе обнаружены разные номера счета"
            If Orders_BulkText(drow(11))<>"" Then Orders_GroupHashTextPut keys(),supplierVals(),flagVals(),cap,groupID,Orders_BulkText(drow(11)),"В одном заказе обнаружены разные поставщики / площадки"
            ' UПД/№ документа и дата могут быть row-level. However, when an
            ' order has exactly ONE non-empty document (or exactly ONE non-empty
            ' document date), a blank cell is understood as continuation of
            ' that single document/date. If several different documents/dates
            ' exist in the same order, blanks are NOT guessed.
            docNo=Orders_BulkText(drow(2))
            dDoc=Orders_BulkDateValue(drow(13),True)
            slot=Orders_GroupHashSlot(keys(),cap,groupID,bFound)
            If slot>=0 Then
                If docNo<>"" Then
                    If docVals(slot)="" Then
                        docVals(slot)=docNo
                    ElseIf Orders_Canon(docVals(slot))<>Orders_Canon(docNo) Then
                        docMixed(slot)=True
                    End If
                End If
                If dDoc>0 Then
                    If dateVals(slot)<=0 Then
                        dateVals(slot)=dDoc
                    ElseIf Abs(dateVals(slot)-dDoc)>WMSORD_EPS Then
                        dateMixed(slot)=True
                    End If
                End If
            End If
        End If
        prevGroup=groupID:prevSupplier=supplier:prevPos=pos:prevPosState=posState
        aTech(i)=trow
Next1:
    Next i

    ' Pass 2: invoice is group-level. Document/date use a conservative
    ' continuation rule: inherit only when the group contains exactly one
    ' distinct non-empty value. With multiple distinct values, blanks stay blank.
    For i=0 To UBound(aTech)
        trow=aTech(i):groupID=Trim(CStr(trow(cGroup))):sSeq=seqFlags(i)
        If groupID="" Then
            trow(cEffInv)="":trow(cEffDoc)="":trow(cEffDate)=0:trow(cEffSupplier)="":trow(cFlags)=sSeq
        Else
            slot=Orders_GroupHashSlot(keys(),cap,groupID,bFound)
            If slot>=0 And bFound Then
                trow(cEffInv)=invVals(slot)
                If Not docMixed(slot) Then
                    trow(cEffDoc)=docVals(slot)
                Else
                    trow(cEffDoc)=Orders_BulkText(aData(i)(2))
                End If
                If Not dateMixed(slot) Then
                    trow(cEffDate)=dateVals(slot)
                Else
                    trow(cEffDate)=Orders_BulkDateValue(aData(i)(13),True)
                End If
                trow(cEffSupplier)=supplierVals(slot)
                trow(cFlags)=flagVals(slot)
            Else
                trow(cEffInv)="":trow(cEffDoc)=Orders_BulkText(aData(i)(2)):trow(cEffDate)=Orders_BulkDateValue(aData(i)(13),True):trow(cEffSupplier)="":trow(cFlags)=""
            End If
            If sSeq<>"" Then
                If Trim(CStr(trow(cFlags)))="" Then
                    trow(cFlags)=sSeq
                ElseIf InStr(1,CStr(trow(cFlags)),sSeq,1)=0 Then
                    trow(cFlags)=CStr(trow(cFlags)) & " | " & sSeq
                End If
            End If
        End If
        aTech(i)=trow
    Next i
    oSh.getCellRangeByPosition(techStart,1,techEnd,last).DataArray=aTech
    Exit Sub
EH:
    Orders_LogError ThisComponent,"OrderGroups","Не удалось перестроить группы заказов: " & CStr(Err) & " " & Error$
End Sub

Sub Orders_GroupMapPut(mVal As Variant,mFlags As Variant,sGroup As String,sValue As String,sConflict As String)
    Dim old As String
    If sGroup="" Or Trim(sValue)="" Then Exit Sub
    If mVal.Exists(sGroup) Then
        old=CStr(mVal.Item(sGroup))
        If Orders_Canon(old)<>Orders_Canon(sValue) Then Orders_GroupFlagPut mFlags,sGroup,sConflict
    Else
        mVal.Add sGroup,sValue
    End If
End Sub

Sub Orders_GroupDateMapPut(mVal As Variant,mFlags As Variant,sGroup As String,dValue As Double,sConflict As String)
    Dim old As Double
    If sGroup="" Or dValue<=0 Then Exit Sub
    If mVal.Exists(sGroup) Then
        old=CDbl(mVal.Item(sGroup))
        If Abs(old-dValue)>WMSORD_EPS Then Orders_GroupFlagPut mFlags,sGroup,sConflict
    Else
        mVal.Add sGroup,dValue
    End If
End Sub

Sub Orders_GroupFlagPut(mFlags As Variant,sGroup As String,sText As String)
    Dim old As String
    If mFlags.Exists(sGroup) Then
        old=CStr(mFlags.Item(sGroup))
        If InStr(1,old,sText,1)=0 Then mFlags.ReplaceItem sGroup,old & " | " & sText
    Else
        mFlags.Add sGroup,sText
    End If
End Sub

Sub Orders_EnsureRowGroupContext(oSh As Object,r As Long)
    Dim cGroup As Long,cEffInv As Long,cEffDoc As Long,cEffDate As Long,cEffSupplier As Long,cFlags As Long
    Dim groupID As String,invoice As String,prevGroup As String,nPos As Long,nPrev As Long,posState As Integer,prevState As Integer,sIssue As String
    If r<1 Then Exit Sub
    cGroup=Orders_FindHeader(oSh,WMSORD_X_GROUPID):cEffInv=Orders_FindHeader(oSh,WMSORD_X_EFFINVOICE)
    cEffDoc=Orders_FindHeader(oSh,WMSORD_X_EFFDOC):cEffDate=Orders_FindHeader(oSh,WMSORD_X_EFFDOCDATE)
    cEffSupplier=Orders_FindHeader(oSh,WMSORD_X_EFFSUPPLIER):cFlags=Orders_FindHeader(oSh,WMSORD_X_GROUPFLAGS)
    If cGroup<0 Or cEffInv<0 Or cEffDoc<0 Or cEffDate<0 Or cEffSupplier<0 Or cFlags<0 Then Exit Sub
    groupID=Trim(oSh.getCellByPosition(cGroup,r).String)
    invoice=Trim(oSh.getCellByPosition(3,r).String)
    posState=Orders_OrderPositionState(oSh,r,nPos)

    If groupID="" Then
        If posState=1 Then
            If nPos=1 Then
                groupID=Orders_NewOrderGroupID(r)
            ElseIf r>1 Then
                prevState=Orders_OrderPositionState(oSh,r-1,nPrev)
                prevGroup=Trim(oSh.getCellByPosition(cGroup,r-1).String)
                If prevState=1 And nPrev+1=nPos And prevGroup<>"" Then
                    groupID=prevGroup
                Else
                    groupID=Orders_NewOrderGroupID(r)
                    sIssue="Нарушена нумерация заказа: позиция " & CStr(nPos) & " должна идти после " & CStr(nPos-1) & " или новый заказ должен начинаться с 1"
                End If
            Else
                groupID=Orders_NewOrderGroupID(r)
                sIssue="Новый заказ должен начинаться с номера 1"
            End If
        ElseIf posState<0 Then
            groupID=Orders_NewOrderGroupID(r)
            sIssue="Колонка Номер должна содержать целое положительное число: 1, 2, 3..."
        ElseIf invoice<>"" Then
            groupID=Orders_MakeGroupKey(oSh,r)
        ElseIf r>1 Then
            prevGroup=Trim(oSh.getCellByPosition(cGroup,r-1).String)
            If prevGroup<>"" And Orders_RowHasBusinessData(oSh,r-1) And Orders_SameSupplierContext(oSh,r-1,r) Then groupID=prevGroup
        End If
        oSh.getCellByPosition(cGroup,r).String=groupID
    End If

    If sIssue<>"" Then oSh.getCellByPosition(cFlags,r).String=sIssue
    If groupID<>"" Then
        If invoice<>"" Then oSh.getCellByPosition(cEffInv,r).String=invoice
        If Trim(oSh.getCellByPosition(11,r).String)<>"" Then oSh.getCellByPosition(cEffSupplier,r).String=Trim(oSh.getCellByPosition(11,r).String)
        ' If this row inherits the group and effective values are still empty,
        ' copy them only from a real predecessor of the same numbered sequence.
        If r>1 And posState=1 And nPos>1 Then
            prevState=Orders_OrderPositionState(oSh,r-1,nPrev)
            prevGroup=Trim(oSh.getCellByPosition(cGroup,r-1).String)
            If prevState=1 And nPrev+1=nPos And prevGroup=groupID Then
                If Trim(oSh.getCellByPosition(cEffInv,r).String)="" Then oSh.getCellByPosition(cEffInv,r).String=oSh.getCellByPosition(cEffInv,r-1).String
                If Trim(oSh.getCellByPosition(cEffSupplier,r).String)="" Then oSh.getCellByPosition(cEffSupplier,r).String=oSh.getCellByPosition(cEffSupplier,r-1).String
            End If
        End If
    End If
End Sub

Function Orders_EffectiveInvoice(oSh As Object,r As Long) As String
    Dim s As String,c As Long
    s=Trim(oSh.getCellByPosition(3,r).String)
    If s="" Then c=Orders_FindHeader(oSh,WMSORD_X_EFFINVOICE):If c>=0 Then s=Trim(oSh.getCellByPosition(c,r).String)
    Orders_EffectiveInvoice=s
End Function

Function Orders_EffectiveDocNo(oSh As Object,r As Long) As String
    ' Effective UПД/№ документа: if this order has exactly one distinct
    ' non-empty document, blank rows inherit that document logically.
    ' If several distinct documents exist, blanks are not guessed.
    Dim s As String,c As Long
    s=Trim(oSh.getCellByPosition(2,r).String)
    If s<>"" Then Orders_EffectiveDocNo=s:Exit Function
    Orders_EffectiveDocNo=Orders_GetSingleGroupDocNo(oSh,r)
End Function

Function Orders_EffectiveDocDate(oSh As Object,r As Long) As Double
    ' Effective document date follows the same conservative rule as UПД:
    ' inherit only when the order has exactly one distinct non-empty date.
    Dim d As Double
    d=Orders_CellDate(oSh.getCellByPosition(13,r),True)
    If d>0 Then Orders_EffectiveDocDate=d:Exit Function
    Orders_EffectiveDocDate=Orders_GetSingleGroupDocDate(oSh,r)
End Function

Function Orders_GetNumberedGroupBounds(oSh As Object,r As Long,ByRef firstRow As Long,ByRef lastRow As Long) As Boolean
    Dim last As Long,i As Long,n As Long,txt As String
    Orders_GetNumberedGroupBounds=False
    firstRow=-1
    lastRow=-1
    If r<1 Then Exit Function
    last=Orders_LastContentRow(oSh,WMSORD_USER_LAST_COL)
    If r>last Then Exit Function

    firstRow=r
    For i=r To 1 Step -1
        txt=Trim(oSh.getCellByPosition(0,i).String)
        If Orders_TryIntegerText(txt,n) Then
            If n=1 Then
                firstRow=i
                Exit For
            End If
        End If
    Next i

    lastRow=last
    For i=firstRow+1 To last
        txt=Trim(oSh.getCellByPosition(0,i).String)
        If Orders_TryIntegerText(txt,n) Then
            If n=1 Then
                lastRow=i-1
                Exit For
            End If
        End If
    Next i
    Orders_GetNumberedGroupBounds=True
End Function

Function Orders_TryIntegerText(s As String,ByRef n As Long) As Boolean
    Dim x As Double,t As String
    Orders_TryIntegerText=False
    n=0
    t=Trim(s)
    If t="" Then Exit Function
    If Not IsNumeric(t) Then Exit Function
    x=CDbl(t)
    If x<>Fix(x) Then Exit Function
    n=CLng(x)
    Orders_TryIntegerText=True
End Function

Function Orders_GetSingleGroupDocNo(oSh As Object,r As Long) As String
    Dim cGroup As Long,c As Long,last As Long,rg As String,s As String,v As String
    Dim firstRow As Long,lastRow As Long,i As Long
    Orders_GetSingleGroupDocNo=""

    cGroup=Orders_FindHeader(oSh,WMSORD_X_GROUPID)
    If cGroup>=0 Then
        rg=Trim(oSh.getCellByPosition(cGroup,r).String)
    Else
        rg=""
    End If

    c=2
    If rg<>"" Then
        last=Orders_LastContentRow(oSh,WMSORD_USER_LAST_COL)
        For i=1 To last
            If Trim(oSh.getCellByPosition(cGroup,i).String)=rg Then
                v=Trim(oSh.getCellByPosition(c,i).String)
                If v<>"" Then
                    If s="" Then
                        s=v
                    Else
                        If Orders_Canon(s)<>Orders_Canon(v) Then
                            Orders_GetSingleGroupDocNo=""
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next i
    Else
        If Orders_GetNumberedGroupBounds(oSh,r,firstRow,lastRow) Then
            For i=firstRow To lastRow
                v=Trim(oSh.getCellByPosition(c,i).String)
                If v<>"" Then
                    If s="" Then
                        s=v
                    Else
                        If Orders_Canon(s)<>Orders_Canon(v) Then
                            Orders_GetSingleGroupDocNo=""
                            Exit Function
                        End If
                    End If
                End If
            Next i
        End If
    End If
    Orders_GetSingleGroupDocNo=s
End Function

Function Orders_GetSingleGroupDocDate(oSh As Object,r As Long) As Double
    Dim cGroup As Long,last As Long,rg As String,d As Double,s As Double
    Dim firstRow As Long,lastRow As Long,i As Long
    Orders_GetSingleGroupDocDate=0

    cGroup=Orders_FindHeader(oSh,WMSORD_X_GROUPID)
    If cGroup>=0 Then
        rg=Trim(oSh.getCellByPosition(cGroup,r).String)
    Else
        rg=""
    End If

    If rg<>"" Then
        last=Orders_LastContentRow(oSh,WMSORD_USER_LAST_COL)
        For i=1 To last
            If Trim(oSh.getCellByPosition(cGroup,i).String)=rg Then
                d=Orders_CellDate(oSh.getCellByPosition(13,i),True)
                If d>0 Then
                    If s<=0 Then
                        s=d
                    Else
                        If Abs(s-d)>WMSORD_EPS Then
                            Orders_GetSingleGroupDocDate=0
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next i
    Else
        If Orders_GetNumberedGroupBounds(oSh,r,firstRow,lastRow) Then
            For i=firstRow To lastRow
                d=Orders_CellDate(oSh.getCellByPosition(13,i),True)
                If d>0 Then
                    If s<=0 Then
                        s=d
                    Else
                        If Abs(s-d)>WMSORD_EPS Then
                            Orders_GetSingleGroupDocDate=0
                            Exit Function
                        End If
                    End If
                End If
            Next i
        End If
    End If
    Orders_GetSingleGroupDocDate=s
End Function

Function Orders_EffectiveSupplier(oSh As Object,r As Long) As String
    Dim s As String,c As Long
    s=Trim(oSh.getCellByPosition(11,r).String)
    If s="" Then c=Orders_FindHeader(oSh,WMSORD_X_EFFSUPPLIER):If c>=0 Then s=Trim(oSh.getCellByPosition(c,r).String)
    Orders_EffectiveSupplier=s
End Function

Function Orders_GroupIDsForRange(oSh As Object,r1 As Long,r2 As Long) As String
    Dim c As Long,r As Long,g As String,s As String,sep As String
    sep=Chr(30)
    c=Orders_FindHeader(oSh,WMSORD_X_GROUPID):If c<0 Then Exit Function
    For r=r1 To r2
        g=Trim(oSh.getCellByPosition(c,r).String)
        If g<>"" Then If InStr(1,sep & s & sep,sep & g & sep,0)=0 Then If s="" Then s=g Else s=s & sep & g
    Next r
    Orders_GroupIDsForRange=s
End Function

Function Orders_GroupListHas(sGroups As String,sGroup As String) As Boolean
    Dim a As Variant,i As Long
    If Trim(sGroups)="" Or Trim(sGroup)="" Then Exit Function
    a=Split(sGroups,Chr(30))
    For i=LBound(a) To UBound(a)
        If CStr(a(i))=sGroup Then Orders_GroupListHas=True:Exit Function
    Next i
End Function

Sub Orders_ProcessGroupSiblings(oDoc As Object,oSh As Object,sGroups As String,r1 As Long,r2 As Long)
    Dim cGroup As Long,cFp As Long,last As Long,r As Long,g As String
    If Trim(sGroups)="" Then Exit Sub
    cGroup=Orders_FindHeader(oSh,WMSORD_X_GROUPID):cFp=Orders_FindHeader(oSh,WMSORD_X_FINGERPRINT)
    If cGroup<0 Then Exit Sub
    last=Orders_LastContentRow(oSh,WMSORD_USER_LAST_COL)
    Orders_ResetDupCache
    For r=1 To last
        If r<r1 Or r>r2 Then
            g=Trim(oSh.getCellByPosition(cGroup,r).String)
            If Orders_GroupListHas(sGroups,g) Then
                Orders_ProcessRow oDoc,oSh,r,True,True,False,False
                If cFp>=0 Then oSh.getCellByPosition(cFp,r).String=Orders_BusinessFingerprint(oSh,r)
            End If
        End If
    Next r
    Orders_ResetDupCache
End Sub

Sub Orders_RebuildOrderGroups()
    Dim oSh As Object,s As String,bOK As Boolean
    If Not ThisComponent.Sheets.hasByName(WMSORD_SHEET) Then
        MsgBox "Нет листа Заказы.",48,"WMS — Заказы"
        Exit Sub
    End If
    oSh=ThisComponent.Sheets.getByName(WMSORD_SHEET)
    gWMSORD_Busy=True
    Orders_RebuildGroupContexts oSh
    gWMSORD_Busy=False
    bOK=Orders_RevalidateAllCore(ThisComponent,True,s)
    If bOK Then
        MsgBox "Группы заказов перестроены." & Chr(10) & s,64,"WMS — Заказы"
    Else
        MsgBox s,48,"WMS — Заказы"
    End If
End Sub


Function Orders_DBQuote(s As String) As String
    Orders_DBQuote = Replace(CStr(s), "'", "''")
End Function

Function Orders_FillProductNameFromDB(oSh As Object, r As Long, ByRef sErr As String) As Boolean
    Dim code As String, oCon As Object, oStmt As Object, oRS As Object, sql As String
    Dim productName As String
    Orders_FillProductNameFromDB = False
    sErr = ""
    On Error GoTo EH

    code = Trim(oSh.getCellByPosition(21,r).String) ' V = Код товара
    If code = "" Then Exit Function

    oCon = WMSDB_GetConnectionEx(ThisComponent, sErr)
    If sErr <> "" Then Exit Function

    If Not WMSDB_TableExistsSimple(oCon, "WMS_PRODUCTS", sErr) Then
        If sErr = "" Then sErr = "В Firebird нет таблицы WMS_PRODUCTS."
        Exit Function
    End If

    sql = "SELECT PRODUCT_NAME FROM WMS_PRODUCTS WHERE PRODUCT_CODE='" & Orders_DBQuote(code) & "'"
    oStmt = oCon.createStatement()
    oRS = oStmt.executeQuery(sql)

    If oRS.next() Then
        productName = Trim(oRS.getString(1))
        If productName <> "" Then
            ' B = Полное наименование товара
            oSh.getCellByPosition(1,r).String = productName
            Orders_FillProductNameFromDB = True
        End If
    End If
    GoTo CloseQuery
EH:
    sErr = "Автозаполнение по коду: " & CStr(Err) & " " & Error$
CloseQuery:
    On Error Resume Next
    oRS.close()
    oStmt.close()
    On Error GoTo 0
End Function

Sub Orders_OnContentChanged(oEvent As Object)
    Dim oSh As Object, r1 As Long, r2 As Long, c1 As Long, c2 As Long, r As Long
    Dim t0 As Double, oldFp As String, newFp As String, finalFp As String, hadChange As Boolean
    Dim cFp As Long, sourceID As String, bDupCacheWasReady As Boolean, bCopyCacheRebuilt As Boolean
    Dim bCheckIdentity As Boolean, bGroupChanged As Boolean, sGroupTargets As String
    On Error GoTo EH
    If gWMSORD_Busy Then Exit Sub
    If gWMSORD_EventDepth > 0 Then Exit Sub
    If Not Orders_GetChangedRange(oEvent, oSh, r1, r2, c1, c2) Then Exit Sub
    If oSh.Name <> WMSORD_SHEET Then Exit Sub
    If r2 < 1 Then Exit Sub
    If r1 < 1 Then r1 = 1
    If r2 - r1 + 1 > WMSORD_MAX_EVENT_ROWS Then r2 = r1 + WMSORD_MAX_EVENT_ROWS - 1

    gWMSORD_Busy = True
    gWMSORD_EventDepth = gWMSORD_EventDepth + 1
    gWMSORD_LastTouched = 0
    t0 = Timer

    ' 1.0.20: лёгкий точечный поиск только при изменении V = Код товара.
    ' Никакого полного перечитывания листа и никаких тяжёлых обработчиков.
    If c1 <= 21 And c2 >= 21 Then
        Dim rrLookup As Long, sLookupErr As String
        For rrLookup = r1 To r2
            sLookupErr = ""
            Call Orders_FillProductNameFromDB(oSh, rrLookup, sLookupErr)
            If sLookupErr <> "" Then gWMSORD_LastError = sLookupErr
        Next rrLookup
    End If

    ' Reassert locale-safe formats and visible cell separators on every changed
    ' row. This also repairs formatting pasted from external procurement files.
    Orders_EnsureChangedRangeFormat ThisComponent, oSh, r1, r2

    ' A native sort includes hidden technical columns. In that case the group
    ' context already travels with the row and must not be rebuilt from visual
    ' adjacency. User/paste edits of A:W still rebuild group context normally.
    bGroupChanged=Orders_GroupRelevantChange(c1,c2)
    If c2>WMSORD_USER_LAST_COL Then bGroupChanged=False
    If bGroupChanged Then
        Orders_PrepareGroupEdits oSh,r1,r2,c1,c2
        Orders_RebuildGroupContexts oSh
        If (r2-r1+1)<=WMSORD_BULK_THRESHOLD Then sGroupTargets=Orders_GroupIDsForRange(oSh,r1,r2)
    End If

    cFp = Orders_FindHeader(oSh, WMSORD_X_FINGERPRINT)
    If (r2-r1+1) > WMSORD_BULK_THRESHOLD Then
        Orders_ProcessBulkLight ThisComponent,oSh,r1,r2,c1,c2
        If bGroupChanged And sGroupTargets<>"" Then Orders_ProcessGroupSiblings ThisComponent,oSh,sGroupTargets,r1,r2
        Orders_RefreshFilter ThisComponent,oSh
        Orders_ResetDupCache
        Orders_ResetHistoryCache
        gWMSORD_LastEventMs = (Timer - t0) * 1000
        If gWMSORD_LastEventMs < 0 Then gWMSORD_LastEventMs = gWMSORD_LastEventMs + 86400000
        GoTo Done
    End If
    ' Duplicate detection uses an in-memory exact-key index. Rebuild once after
    ' reopen/install, then update only changed rows. This avoids O(n^2) scans
    ' when procurement data is pasted in bulk.
    bDupCacheWasReady = gWMSORD_DupCacheReady
    If Not bDupCacheWasReady Then Call Orders_EnsureDupCache(oSh)

    For r = r1 To r2
        oldFp = ""
        If cFp >= 0 Then oldFp = Trim(oSh.getCellByPosition(cFp,r).String)
        newFp = Orders_BusinessFingerprint(oSh, r)
        hadChange = (oldFp = "" Or oldFp <> newFp)

        ' Sorting carries hidden fingerprint with the row. If the business content did
        ' not change, do not mark manual overrides/dirty just because row number moved.
        If hadChange Then
            Orders_DetectManualOverrides oSh, r, c1, c2
            bCheckIdentity = (c2 > WMSORD_USER_LAST_COL)
            Orders_ProcessRow ThisComponent, oSh, r, True, True, bCheckIdentity, bDupCacheWasReady
            finalFp = Orders_BusinessFingerprint(oSh, r)
            If cFp >= 0 Then oSh.getCellByPosition(cFp,r).String = finalFp
            gWMSORD_LastTouched = gWMSORD_LastTouched + 1
        Else
            ' A copied whole row can carry both SourceID and fingerprint. Keep the
            ' expensive uniqueness scan behind nested conditions: LibreOffice Basic
            ' does not guarantee short-circuit evaluation for And expressions.
            sourceID = Orders_CellText(oSh, r, WMSORD_T_ID)
            If (r2-r1+1) <= 50 Then
                If sourceID <> "" Then
                    If Orders_SourceIDCount(oSh,sourceID) > 1 Then
                        ' The business duplicate cache must include the newly copied
                        ' physical row. Rebuild it once for the whole copy event.
                        If Not bCopyCacheRebuilt Then
                            Orders_ResetDupCache
                            Call Orders_EnsureDupCache(oSh)
                            bCopyCacheRebuilt = True
                        End If
                        Orders_ProcessRow ThisComponent, oSh, r, True, True, True, False
                        finalFp = Orders_BusinessFingerprint(oSh, r)
                        If cFp >= 0 Then oSh.getCellByPosition(cFp,r).String = finalFp
                        gWMSORD_LastTouched = gWMSORD_LastTouched + 1
                    Else
                        Orders_ApplyRowVisual oSh, r
                    End If
                Else
                    Orders_ApplyRowVisual oSh, r
                End If
            Else
                Orders_ApplyRowVisual oSh, r
            End If
        End If
    Next r

    If bGroupChanged Then Orders_ProcessGroupSiblings ThisComponent,oSh,sGroupTargets,r1,r2
    Orders_RefreshFilter ThisComponent, oSh
    If (c1<=5 And c2>=5) Or (c1<=11 And c2>=11) Or (c1<=17 And c2>=17) Or (c1<=19 And c2>=19) Or (c1<=21 And c2>=21) Or (c1<=22 And c2>=22) Then Orders_ResetHistoryCache
    gWMSORD_LastEventMs = (Timer - t0) * 1000
    If gWMSORD_LastEventMs < 0 Then gWMSORD_LastEventMs = gWMSORD_LastEventMs + 86400000

Done:
    gWMSORD_EventDepth = gWMSORD_EventDepth - 1
    If gWMSORD_EventDepth < 0 Then gWMSORD_EventDepth = 0
    gWMSORD_Busy = False
    Exit Sub
EH:
    gWMSORD_LastError = "Orders_OnContentChanged: " & CStr(Err) & " " & Error$
    Orders_LogError ThisComponent, "OnContentChanged", gWMSORD_LastError
    Resume Done
End Sub

Function Orders_GetChangedRange(oEvent As Object, ByRef oSh As Object, ByRef r1 As Long, ByRef r2 As Long, ByRef c1 As Long, ByRef c2 As Long) As Boolean
    Dim a As Variant
    Orders_GetChangedRange = False
    On Error GoTo TryCell
    a = oEvent.RangeAddress
    r1 = a.StartRow: r2 = a.EndRow: c1 = a.StartColumn: c2 = a.EndColumn
    oSh = ThisComponent.Sheets.getByIndex(a.Sheet)
    Orders_GetChangedRange = True
    Exit Function
TryCell:
    On Error GoTo Bad
    a = oEvent.CellAddress
    r1 = a.Row: r2 = a.Row: c1 = a.Column: c2 = a.Column
    oSh = ThisComponent.Sheets.getByIndex(a.Sheet)
    Orders_GetChangedRange = True
    Exit Function
Bad:
End Function

Sub Orders_DetectManualOverrides(oSh As Object, r As Long, c1 As Long, c2 As Long)
    Dim c As Long, s As String
    ' Status Q: any direct nonblank user edit becomes manual. Clearing returns AUTO.
    If c1 <= 16 And c2 >= 16 Then
        c = Orders_FindHeader(oSh, WMSORD_X_STATUSMODE)
        If c >= 0 Then
            s = Trim(oSh.getCellByPosition(16,r).String)
            If s = "" Then oSh.getCellByPosition(c,r).String = "AUTO" Else oSh.getCellByPosition(c,r).String = "MANUAL"
        End If
    End If
    ' Category R.
    If c1 <= 17 And c2 >= 17 Then
        c = Orders_FindHeader(oSh, WMSORD_X_CATORIGIN)
        If c >= 0 Then
            s = Trim(oSh.getCellByPosition(17,r).String)
            If s = "" Then oSh.getCellByPosition(c,r).String = "AUTO" Else oSh.getCellByPosition(c,r).String = "MANUAL"
        End If
    End If
    ' Receipt date M.
    If c1 <= 12 And c2 >= 12 Then
        c = Orders_FindHeader(oSh, WMSORD_X_DATEORIGIN)
        If c >= 0 Then
            If Orders_CellDate(oSh.getCellByPosition(12,r), True) > 0 Then oSh.getCellByPosition(c,r).String = "MANUAL" Else oSh.getCellByPosition(c,r).String = "AUTO"
        End If
    End If
    ' Code V.
    If c1 <= 21 And c2 >= 21 Then
        c = Orders_FindHeader(oSh, WMSORD_X_CODEORIGIN)
        If c >= 0 Then
            s = Trim(oSh.getCellByPosition(21,r).String)
            If s = "" Then oSh.getCellByPosition(c,r).String = "AUTO" Else oSh.getCellByPosition(c,r).String = "MANUAL"
        End If
    End If
End Sub

Sub Orders_ProcessBulkLight(oDoc As Object,oSh As Object,r1 As Long,r2 As Long,c1 As Long,c2 As Long)
    Dim n As Long,i As Long,r As Long,techStart As Long,techEnd As Long
    Dim cID As Long,cState As Long,cErr As Long,cDirty As Long,cMode As Long,cLegacy As Long,cTouch As Long
    Dim cRowVer As Long,cStatusMode As Long,cCodeOrigin As Long,cCatOrigin As Long,cFlags As Long,cSeverity As Long,cValidated As Long,cDup As Long,cDateOrigin As Long,cFp As Long,cSugSource As Long,cCodeSug As Long,cCatSug As Long,cLocSug As Long,cGroup As Long,cEffInv As Long,cEffDoc As Long,cEffDate As Long,cEffSupplier As Long,cGroupFlags As Long
    Dim aData As Variant,aFormula As Variant,aTech As Variant,drow As Variant,trow As Variant,frow As Variant
    Dim outM() As Variant,outN() As Variant,outO() As Variant,outQ() As Variant,outS() As Variant
    Dim aChanged() As Boolean,aSID() As String,aMode() As String,aUse() As Boolean,aNewSID() As Boolean
    Dim oldFp As String,newFp As String,sid As String,state As String,statusMode As String,codeOrigin As String,catOrigin As String,dateOrigin As String
    Dim status As String,docNo As String,sFlags As String,sSeverity As String,sControl As String,stamp As String
    Dim fact As Double,ordered As Double,okFact As Boolean,okOrdered As Boolean,mDate As Double,nDate As Double,oDate As Double,refDate As Double
    Dim legacyDays As Long,age As Double,mode As String,nVer As Double,hasData As Boolean,bAnyChanged As Boolean
    Dim sidCounts As Variant,sidData As Variant,sidRow As Variant,last As Long,j As Long,cnt As Long,bTechTouched As Boolean

    On Error GoTo EH
    n=r2-r1+1
    If n<=0 Then Exit Sub
    techStart=23:techEnd=Orders_LastHeaderCol(oSh)
    If techEnd<techStart Then Exit Sub

    cID=Orders_FindHeader(oSh,WMSORD_T_ID)-techStart
    cState=Orders_FindHeader(oSh,WMSORD_T_STATE)-techStart
    cErr=Orders_FindHeader(oSh,WMSORD_T_ERROR)-techStart
    cDirty=Orders_FindHeader(oSh,WMSORD_T_DIRTY)-techStart
    cMode=Orders_FindHeader(oSh,WMSORD_T_MODE)-techStart
    cLegacy=Orders_FindHeader(oSh,WMSORD_T_LEGACY)-techStart
    cTouch=Orders_FindHeader(oSh,WMSORD_T_TOUCH)-techStart
    cRowVer=Orders_FindHeader(oSh,WMSORD_X_ROWVER)-techStart
    cStatusMode=Orders_FindHeader(oSh,WMSORD_X_STATUSMODE)-techStart
    cCodeSug=Orders_FindHeader(oSh,WMSORD_X_CODESUG)-techStart
    cCodeOrigin=Orders_FindHeader(oSh,WMSORD_X_CODEORIGIN)-techStart
    cCatSug=Orders_FindHeader(oSh,WMSORD_X_CATSUG)-techStart
    cCatOrigin=Orders_FindHeader(oSh,WMSORD_X_CATORIGIN)-techStart
    cLocSug=Orders_FindHeader(oSh,WMSORD_X_LOCSUG)-techStart
    cFlags=Orders_FindHeader(oSh,WMSORD_X_FLAGS)-techStart
    cSeverity=Orders_FindHeader(oSh,WMSORD_X_SEVERITY)-techStart
    cValidated=Orders_FindHeader(oSh,WMSORD_X_VALIDATED)-techStart
    cDup=Orders_FindHeader(oSh,WMSORD_X_DUPKEY)-techStart
    cDateOrigin=Orders_FindHeader(oSh,WMSORD_X_DATEORIGIN)-techStart
    cFp=Orders_FindHeader(oSh,WMSORD_X_FINGERPRINT)-techStart
    cSugSource=Orders_FindHeader(oSh,WMSORD_X_SUGSOURCE)-techStart
    cGroup=Orders_FindHeader(oSh,WMSORD_X_GROUPID)-techStart
    cEffInv=Orders_FindHeader(oSh,WMSORD_X_EFFINVOICE)-techStart
    cEffDoc=Orders_FindHeader(oSh,WMSORD_X_EFFDOC)-techStart
    cEffDate=Orders_FindHeader(oSh,WMSORD_X_EFFDOCDATE)-techStart
    cEffSupplier=Orders_FindHeader(oSh,WMSORD_X_EFFSUPPLIER)-techStart
    cGroupFlags=Orders_FindHeader(oSh,WMSORD_X_GROUPFLAGS)-techStart

    aData=oSh.getCellRangeByPosition(0,r1,WMSORD_USER_LAST_COL,r2).DataArray
    aFormula=oSh.getCellRangeByPosition(0,r1,WMSORD_USER_LAST_COL,r2).FormulaArray
    aTech=oSh.getCellRangeByPosition(techStart,r1,techEnd,r2).DataArray
    ReDim outM(0 To n-1):ReDim outN(0 To n-1):ReDim outO(0 To n-1):ReDim outQ(0 To n-1):ReDim outS(0 To n-1)
    ReDim aChanged(0 To n-1):ReDim aSID(0 To n-1):ReDim aMode(0 To n-1):ReDim aUse(0 To n-1):ReDim aNewSID(0 To n-1)
    stamp=Orders_Stamp(Now)
    legacyDays=Orders_LongSetting(oDoc,"ORDERS_LegacyAgeDays",WMSORD_DEFAULT_LEGACY_DAYS)
    bTechTouched=(c2>WMSORD_USER_LAST_COL)

    ' Large native sorts also include hidden columns. Build SourceID counts once so
    ' copied whole rows can be healed without turning a sort into O(n^2).
    If bTechTouched Then
        GlobalScope.BasicLibraries.LoadLibrary("ScriptForge")
        sidCounts=CreateScriptService("Dictionary",False)
        last=Orders_LastContentRow(oSh,WMSORD_USER_LAST_COL)
        If last>=1 Then
            sidData=oSh.getCellRangeByPosition(Orders_FindHeader(oSh,WMSORD_T_ID),1,Orders_FindHeader(oSh,WMSORD_T_ID),last).DataArray
            For j=0 To UBound(sidData)
                sidRow=sidData(j):sid=Trim(CStr(sidRow(0)))
                If sid<>"" Then
                    If sidCounts.Exists(sid) Then
                        cnt=CLng(sidCounts.Item(sid)):sidCounts.ReplaceItem sid,cnt+1
                    Else
                        sidCounts.Add sid,1
                    End If
                End If
            Next j
        End If
    End If

    ' First pass: detect real business changes using the carried fingerprint.
    For i=0 To n-1
        trow=aTech(i):frow=aFormula(i)
        oldFp="":If cFp>=0 Then oldFp=Trim(CStr(trow(cFp)))
        newFp=Orders_FingerprintFromFormulaRow(frow)
        aChanged(i)=(oldFp="" Or oldFp<>newFp)
        If bTechTouched And cID>=0 Then
            sid=Trim(CStr(trow(cID)))
            If sid<>"" Then
                If sidCounts.Exists(sid) Then
                    If CLng(sidCounts.Item(sid))>1 Then aChanged(i)=True
                End If
            End If
        End If
        If aChanged(i) Then bAnyChanged=True
    Next i
    If Not bAnyChanged Then
        For i=0 To n-1
            r=r1+i: Orders_ApplyRowVisual oSh,r
        Next i
        GoTo BulkDone
    End If

    For i=0 To n-1
        drow=aData(i):trow=aTech(i):r=r1+i
        outM(i)=Array(drow(12)):outN(i)=Array(drow(13)):outO(i)=Array(drow(14)):outQ(i)=Array(drow(16)):outS(i)=Array(drow(18))
        ' Use hidden effective group context only for group-level fields.
        ' Document number/date stay row-level and are never inherited.
        If cEffInv>=0 Then If Orders_BulkText(trow(cEffInv))<>"" Then drow(3)=trow(cEffInv)
        If cEffDoc>=0 Then If Orders_BulkText(trow(cEffDoc))<>"" Then drow(2)=trow(cEffDoc)
        If cEffDate>=0 Then If IsNumeric(trow(cEffDate)) Then If CDbl(trow(cEffDate))>0 Then drow(13)=trow(cEffDate)
        If cEffSupplier>=0 Then If Orders_BulkText(trow(cEffSupplier))<>"" Then drow(11)=trow(cEffSupplier)
        If Not aChanged(i) Then GoTo NextBulkRow

        hasData=Orders_BulkRowHasData(drow)
        If Not hasData Then
            If cID>=0 Then sid=Trim(CStr(trow(cID))) Else sid=""
            If sid="" Then GoTo NextBulkRow
        End If

        sid="":If cID>=0 Then sid=Trim(CStr(trow(cID)))
        state="":If cState>=0 Then state=Trim(CStr(trow(cState)))
        If sid="" Then
            sid=Orders_NewSourceID(r)
            aNewSID(i)=True
        ElseIf bTechTouched Then
            If sidCounts.Exists(sid) Then
                If CLng(sidCounts.Item(sid))>1 Then
                    If Orders_Canon(state)<>Orders_Canon("Проведено") Then
                        sid=Orders_NewSourceID(r)
                        aNewSID(i)=True
                    End If
                End If
            End If
        End If
        If cID>=0 Then trow(cID)=sid
        If bTechTouched And aNewSID(i) And cGroup>=0 Then
            Dim healPos As Long,healState As Integer
            healState=Orders_BulkOrderPositionState(drow(0),healPos)
            If healState=1 And healPos=1 Then
                ' A copied whole row can carry the old hidden GroupID. Duplicate
                ' SourceID proves this is a copy, not a sort; A=1 is a hard new-order
                ' boundary, so heal group identity together with SourceID.
                trow(cGroup)=Orders_NewOrderGroupID(r)
                If cEffInv>=0 Then trow(cEffInv)=Orders_BulkText(drow(3))
                If cEffDoc>=0 Then trow(cEffDoc)=""
                If cEffDate>=0 Then trow(cEffDate)=0
                If cEffSupplier>=0 Then trow(cEffSupplier)=Orders_BulkText(drow(11))
                If cGroupFlags>=0 Then trow(cGroupFlags)=""
            End If
        End If
        If state="" Then state="Черновик":If cState>=0 Then trow(cState)=state

        ' Normalize current-operation dates in memory.
        mDate=Orders_BulkDateValue(drow(12),True):nDate=Orders_BulkDateValue(drow(13),True):oDate=Orders_BulkDateValue(drow(14),True)
        If mDate>0 Then drow(12)=mDate
        If nDate>0 Then drow(13)=nDate
        If oDate>0 Then drow(14)=oDate

        dateOrigin="":If cDateOrigin>=0 Then dateOrigin=UCase(Trim(CStr(trow(cDateOrigin))))
        If c1<=12 And c2>=12 Then
            If mDate>0 Then dateOrigin="MANUAL" Else dateOrigin="AUTO"
        End If
        okFact=Orders_BulkTryNumber(drow(6),fact)
        If okFact And fact>0 And mDate<=0 And Orders_BulkText(drow(12))="" And dateOrigin<>"MANUAL" Then
            mDate=CDbl(Date):drow(12)=mDate:dateOrigin="AUTO_FIRST_FACT"
        End If
        If cDateOrigin>=0 Then trow(cDateOrigin)=dateOrigin

        mode="":If cMode>=0 Then mode=UCase(Trim(CStr(trow(cMode))))
        If mode="" Then
            refDate=mDate:If refDate<=0 Then refDate=oDate:If refDate<=0 Then refDate=nDate
            mode="CURRENT"
            If refDate>0 Then
                age=CDbl(Date)-refDate
                If age>legacyDays Then mode="HISTORY"
            End If
            If cMode>=0 Then trow(cMode)=mode
            If mode<>"CURRENT" And cLegacy>=0 Then
                If Trim(CStr(trow(cLegacy)))="" Then trow(cLegacy)=Orders_NewLegacyKey(r)
            End If
        End If

        statusMode="":If cStatusMode>=0 Then statusMode=UCase(Trim(CStr(trow(cStatusMode))))
        status=Orders_BulkText(drow(16))
        If c1<=16 And c2>=16 Then
            If status="" Then statusMode="AUTO" Else statusMode="MANUAL"
        End If
        If statusMode="" Then statusMode="AUTO"
        If statusMode<>"MANUAL" Then
            docNo=Orders_BulkText(drow(2))
            If docNo="" And cEffDoc>=0 Then docNo=Orders_BulkText(trow(cEffDoc))
            If Not okFact Or fact<=0 Then
                status=WMSORD_STATUS_TRANSIT
            ElseIf docNo="" Then
                status=WMSORD_STATUS_NO_DOCS
            Else
                status=WMSORD_STATUS_RECEIVED
            End If
            drow(16)=status
        End If
        If cStatusMode>=0 Then trow(cStatusMode)=statusMode

        codeOrigin="":If cCodeOrigin>=0 Then codeOrigin=UCase(Trim(CStr(trow(cCodeOrigin))))
        If c1<=21 And c2>=21 Then
            If Orders_BulkText(drow(21))="" Then codeOrigin="AUTO" Else codeOrigin="MANUAL"
        End If
        If cCodeOrigin>=0 Then trow(cCodeOrigin)=codeOrigin
        catOrigin="":If cCatOrigin>=0 Then catOrigin=UCase(Trim(CStr(trow(cCatOrigin))))
        If c1<=17 And c2>=17 Then
            If Orders_BulkText(drow(17))="" Then catOrigin="AUTO" Else catOrigin="MANUAL"
        End If
        If cCatOrigin>=0 Then trow(cCatOrigin)=catOrigin

        If (c1<=5 And c2>=5) Or (c1<=11 And c2>=11) Or (c1<=22 And c2>=22) Then
            If cCodeSug>=0 Then trow(cCodeSug)=""
            If cCatSug>=0 Then trow(cCatSug)=""
            If cLocSug>=0 Then trow(cLocSug)=""
            If cSugSource>=0 Then trow(cSugSource)="DEFERRED_BULK"
        End If

        sFlags="":sSeverity="":sControl=""
        Orders_BulkValidate drow,state,sSeverity,sFlags,sControl,legacyDays
        If cGroupFlags>=0 Then
            If Orders_BulkText(trow(cGroupFlags))<>"" Then
                Orders_AppendFlag sFlags,"CRITICAL: " & Orders_BulkText(trow(cGroupFlags))
                sSeverity="CRITICAL":sControl=Orders_BulkText(trow(cGroupFlags))
            End If
        End If
        drow(18)=sControl
        If cFlags>=0 Then trow(cFlags)=sFlags
        If cSeverity>=0 Then trow(cSeverity)=sSeverity
        If cValidated>=0 Then trow(cValidated)=stamp
        If cDup>=0 Then trow(cDup)=Orders_BulkDuplicateKey(drow)
        If cErr>=0 Then
            If sSeverity="CRITICAL" Then trow(cErr)=sControl Else trow(cErr)=""
        End If
        If cDirty>=0 Then trow(cDirty)="DIRTY"
        If cTouch>=0 Then trow(cTouch)=stamp
        If cRowVer>=0 Then
            nVer=0:If IsNumeric(trow(cRowVer)) Then nVer=CDbl(trow(cRowVer))
            trow(cRowVer)=nVer+1
        End If
        aSID(i)=sid:aMode(i)=mode:aUse(i)=True
        aTech(i)=trow
        outM(i)=Array(drow(12)):outO(i)=Array(drow(14)):outQ(i)=Array(drow(16)):outS(i)=Array(drow(18))
NextBulkRow:
    Next i

    ' Five vector writes instead of thousands of cell-by-cell writes.
    oSh.getCellRangeByPosition(12,r1,12,r2).DataArray=outM
    oSh.getCellRangeByPosition(13,r1,13,r2).DataArray=outN
    oSh.getCellRangeByPosition(14,r1,14,r2).DataArray=outO
    oSh.getCellRangeByPosition(16,r1,16,r2).DataArray=outQ
    oSh.getCellRangeByPosition(18,r1,18,r2).DataArray=outS

    ' Re-read final formulas once so sort fingerprint exactly matches Calc's own
    ' representation of dates/numbers/text after our vector writes.
    aFormula=oSh.getCellRangeByPosition(0,r1,WMSORD_USER_LAST_COL,r2).FormulaArray
    For i=0 To n-1
        If aChanged(i) Then
            trow=aTech(i):frow=aFormula(i)
            If cFp>=0 Then trow(cFp)=Orders_FingerprintFromFormulaRow(frow)
            aTech(i)=trow
        End If
    Next i
    oSh.getCellRangeByPosition(techStart,r1,techEnd,r2).DataArray=aTech
    Orders_BulkQueue oDoc,oSh,r1,aSID(),aMode(),aUse(),aNewSID()
    gWMSORD_LastTouched=n
BulkDone:
    On Error Resume Next
    If bTechTouched Then sidCounts.Dispose
    On Error GoTo 0
    Exit Sub
EH:
    gWMSORD_LastError="Orders_ProcessBulkLight: " & CStr(Err) & " " & Error$
    Orders_LogError oDoc,"BulkLight",gWMSORD_LastError
End Sub

Function Orders_BulkText(v As Variant) As String
    On Error GoTo Done
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    Orders_BulkText=Trim(CStr(v))
Done:
End Function

Function Orders_BulkTryNumber(v As Variant,ByRef d As Double) As Boolean
    Dim s As String
    Orders_BulkTryNumber=False:d=0
    On Error GoTo Done
    If IsEmpty(v) Or IsNull(v) Then Exit Function
    If IsNumeric(v) Then d=CDbl(v):Orders_BulkTryNumber=True:Exit Function
    s=Trim(CStr(v)):If s="" Then Exit Function
    s=Replace(s,Chr(160),""):s=Replace(s," ",""):s=Replace(s,",",".")
    If IsNumeric(s) Then d=CDbl(s):Orders_BulkTryNumber=True
Done:
End Function

Function Orders_BulkDateValue(v As Variant,bInferYear As Boolean) As Double
    Dim s As String,a As Variant,dd As Integer,mm As Integer,yy As Integer,dt As Date,d As Double
    Orders_BulkDateValue=0
    On Error GoTo Done
    If IsNumeric(v) Then d=CDbl(v):If d>0 Then Orders_BulkDateValue=d:Exit Function
    s=Trim(CStr(v)):If s="" Then Exit Function
    s=Replace(s,"/","."):s=Replace(s,"-",".")
    Do While InStr(s,"..")>0:s=Replace(s,"..","."):Loop
    a=Split(s,".")
    If UBound(a)=1 Then
        If Not bInferYear Then Exit Function
        dd=CInt(a(0)):mm=CInt(a(1)):yy=Year(Date)
    ElseIf UBound(a)=2 Then
        dd=CInt(a(0)):mm=CInt(a(1)):yy=CInt(a(2)):If yy<100 Then yy=2000+yy
    Else
        Exit Function
    End If
    dt=DateSerial(yy,mm,dd)
    If Day(dt)=dd And Month(dt)=mm And Year(dt)=yy Then Orders_BulkDateValue=CDbl(dt)
Done:
End Function

Function Orders_BulkRowHasData(aRow As Variant) As Boolean
    Dim c As Long
    ' 1.0.13 Revalidate Guard:
    ' A (Номер) may be prefilled as a template and by itself is NOT an order row.
    ' Q (Статус) and S (Контроль) are generated by WMS and are also ignored.
    For c=0 To WMSORD_USER_LAST_COL
        If c<>0 And c<>16 And c<>18 Then
            If Orders_BulkText(aRow(c))<>"" Then Orders_BulkRowHasData=True:Exit Function
        End If
    Next c
End Function

Function Orders_FingerprintFromFormulaRow(aRow As Variant) As String
    Dim c As Long,s As String,p As String,h1 As Double,h2 As Double,i As Long,ch As Long
    For c=0 To WMSORD_USER_LAST_COL
        If c<>18 Then p=CStr(aRow(c)):s=s & "|" & CStr(c) & "=" & p
    Next c
    h1=5381:h2=7919
    For i=1 To Len(s)
        ch=Asc(Mid(s,i,1)):If ch<0 Then ch=ch+256
        h1=(h1*33+ch) Mod 1000003:h2=(h2*131+ch+i) Mod 1000033
    Next i
    Orders_FingerprintFromFormulaRow=Hex(CLng(h1)) & "-" & Hex(CLng(h2)) & "-" & CStr(Len(s))
End Function

Function Orders_BulkDuplicateKey(aRow As Variant) As String
    Dim no As String,invoice As String,sup As String,art As String,buyer As String,d As Double,q As Double,ok As Boolean
    no=Orders_BulkText(aRow(0)):invoice=Orders_BulkText(aRow(3)):sup=Orders_BulkText(aRow(4)):art=Orders_BulkText(aRow(5)):buyer=Orders_BulkText(aRow(15))
    d=Orders_BulkDateValue(aRow(14),True):ok=Orders_BulkTryNumber(aRow(7),q)
    If no="" Or invoice="" Or art="" Or d<=0 Or Not ok Then Exit Function
    Orders_BulkDuplicateKey=Orders_Canon(no) & "|" & Orders_Canon(invoice) & "|" & Orders_Canon(sup) & "|" & Orders_Canon(art) & "|" & Orders_NumKey(q) & "|" & CStr(CLng(d)) & "|" & Orders_Canon(buyer)
End Function

Sub Orders_BulkValidate(aRow As Variant,sState As String,ByRef sSeverity As String,ByRef sFlags As String,ByRef sControl As String,legacyDays As Long)
    Dim name As String,docNo As String,invoice As String,unit As String,platform As String,seller As String,status As String,cat As String,place As String,code As String
    Dim ordered As Double,fact As Double,price As Double,total As Double,okOrdered As Boolean,okFact As Boolean,okPrice As Boolean,okTotal As Boolean
    Dim dArr As Double,dDoc As Double,dOrder As Double,expected As Double,age As Long,firstCritical As String,firstWarn As String,firstInfo As String
    Dim orderPos As Long,posState As Integer
    name=Orders_BulkText(aRow(1)):docNo=Orders_BulkText(aRow(2)):invoice=Orders_BulkText(aRow(3)):unit=Orders_BulkText(aRow(8))
    platform=Orders_BulkText(aRow(11)):seller=Orders_BulkText(aRow(22)):status=Orders_BulkText(aRow(16)):cat=Orders_BulkText(aRow(17)):place=Orders_BulkText(aRow(19)):code=Orders_BulkText(aRow(21))
    okOrdered=Orders_BulkTryNumber(aRow(7),ordered):okFact=Orders_BulkTryNumber(aRow(6),fact):okPrice=Orders_BulkTryNumber(aRow(9),price):okTotal=Orders_BulkTryNumber(aRow(10),total)
    dArr=Orders_BulkDateValue(aRow(12),True):dDoc=Orders_BulkDateValue(aRow(13),True):dOrder=Orders_BulkDateValue(aRow(14),True)
    sSeverity="OK":sControl="":sFlags=""
    posState=Orders_BulkOrderPositionState(aRow(0),orderPos)
    If posState=0 Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Не заполнен номер позиции заказа в колонке Номер"
    If posState<0 Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"CRITICAL","Колонка Номер должна содержать целое положительное число: 1, 2, 3..."
    If name="" Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"CRITICAL","Заполните наименование товара"
    If Not okOrdered Or ordered<=0 Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"CRITICAL","Проверьте заказанное количество"
    If unit="" Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Не заполнена единица измерения"
    If invoice="" Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Не заполнен номер счета"
    If platform="" Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Не заполнен поставщик / площадка"
    If dOrder<=0 Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Не заполнена дата заказа"
    If Orders_BulkText(aRow(6))<>"" Then If Not okFact Or fact<0 Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"CRITICAL","Факт. количество должно быть числом не меньше нуля"
    If okFact And okOrdered Then
        If fact>ordered+WMSORD_EPS Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"CRITICAL","Факт больше заказанного — проверьте количество"
        If fact>WMSORD_EPS And fact<ordered-WMSORD_EPS Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Факт меньше заказа; остаток может приехать другим документом"
    End If
    If okFact And fact>WMSORD_EPS Then
        If dArr<=0 Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Укажите дату поступления"
        If place="" Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Укажите место хранения"
        If cat="" Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Укажите категорию"
        If code="" Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"INFO","У позиции пока нет кода товара"
        If docNo="" Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Товар получен без документов"
    ElseIf dArr>0 Then
        Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Дата поступления есть, а положительного факта нет"
    End If
    If docNo<>"" And dDoc<=0 Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Есть № документа, но нет даты документа"
    If docNo="" And dDoc>0 Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Есть дата документа, но нет № документа"
    If dOrder>CDbl(Date)+WMSORD_EPS Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Дата заказа находится в будущем"
    If dArr>CDbl(Date)+WMSORD_EPS Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Дата поступления находится в будущем"
    If dDoc>CDbl(Date)+WMSORD_EPS Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Дата документа находится в будущем"
    If dOrder>0 And dArr>0 And dArr+WMSORD_EPS<dOrder Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Дата поступления раньше даты заказа"
    If Orders_IsMarketplace(platform) And seller="" Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Для маркетплейса укажите продавца"
    If okOrdered And okPrice And okTotal And ordered>=0 And price>=0 Then
        expected=ordered*price:If Abs(total-expected)>0.02 Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Сумма не равна Количество × Цена"
    End If
    If Orders_Canon(status)=Orders_Canon(WMSORD_STATUS_POST) Then
        If Not okFact Or fact<=0 Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"CRITICAL","Оприходовано требует положительный факт"
        If docNo="" Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"CRITICAL","Оприходование без № документа запрещено"
    ElseIf Orders_Canon(status)=Orders_Canon(WMSORD_STATUS_CANCEL) Then
        Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"INFO","Заказ отменён вручную"
    End If
    If dOrder>0 And (Not okFact Or fact<=0) Then age=CLng(CDbl(Date)-dOrder):If age>legacyDays Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Заказ остаётся без факта " & CStr(age) & " дн."
    If Orders_Canon(sState)=Orders_Canon("Проведено") Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"CRITICAL","Проведённая строка изменена — требуется безопасное отмена/перепроведение"
    If firstCritical<>"" Then sSeverity="CRITICAL":sControl=firstCritical:Exit Sub
    If firstWarn<>"" Then sSeverity="WARN":sControl=firstWarn:Exit Sub
    If firstInfo<>"" Then sSeverity="INFO":sControl=firstInfo:Exit Sub
    If Orders_Canon(status)=Orders_Canon(WMSORD_STATUS_CANCEL) Then
        sSeverity="INFO":sControl="Отменено"
    ElseIf Orders_Canon(status)=Orders_Canon(WMSORD_STATUS_POST) Then
        sSeverity="OK":sControl="Готово к проведению"
    ElseIf okFact And fact>0 And docNo="" Then
        sSeverity="WARN":sControl="Ждём документы"
    ElseIf okFact And fact>0 And docNo<>"" Then
        sSeverity="OK":sControl="Получено — проверьте и оприходуйте"
    Else
        sSeverity="INFO":sControl="В пути"
    End If
End Sub

Sub Orders_BulkQueue(oDoc As Object,oSh As Object,r1 As Long,aSID() As String,aMode() As String,aUse() As Boolean,aNewSID() As Boolean)
    Dim oQ As Object,i As Long,found As Long,status As String,appendCount As Long,startRow As Long,idx As Long,target As Long,stamp As String
    Dim qLast As Long,qData As Variant,qRow As Variant,aRows() As Variant,aAppend() As Boolean,bModified As Boolean
    If Not oDoc.Sheets.hasByName(WMSORD_QUEUE) Then Exit Sub
    On Error GoTo EH
    oQ=oDoc.Sheets.getByName(WMSORD_QUEUE)
    Orders_EnsureQueueHeaders oQ
    Call Orders_EnsureQueueCache(oDoc)
    qLast=Orders_QueueLastRowFast(oQ)
    If qLast>=1 Then qData=oQ.getCellRangeByPosition(0,1,7,qLast).DataArray
    ReDim aAppend(LBound(aSID) To UBound(aSID))
    stamp=Orders_Stamp(Now)

    ' Resolve existing PENDING records entirely in memory. Newly generated
    ' SourceIDs are known not to exist in the queue, so they bypass any lookup.
    For i=LBound(aSID) To UBound(aSID)
        If aUse(i) And Trim(aSID(i))<>"" Then
            found=-1
            If Not aNewSID(i) Then
                If gWMSORD_QueueCacheReady Then found=Orders_QueueHashGet(aSID(i))
                If found<1 And qLast>=1 And Not gWMSORD_QueueCacheReady Then found=Orders_QueueArrayFindPending(qData,aSID(i))
            End If
            If found>=1 And found<=qLast Then
                qRow=qData(found-1)
                status=UCase(Trim(CStr(qRow(5))))
                If CStr(qRow(1))=WMSORD_SHEET And Trim(CStr(qRow(2)))=aSID(i) And (status="PENDING" Or status="") Then
                    qRow(0)=stamp:qRow(3)=r1+i:qRow(4)=aMode(i):qRow(5)="PENDING":qRow(6)="":qRow(7)="ORDER_CHANGED"
                    qData(found-1)=qRow
                    bModified=True
                Else
                    found=-1
                End If
            End If
            If found<1 Then aAppend(i)=True:appendCount=appendCount+1
        End If
    Next i

    If bModified And qLast>=1 Then oQ.getCellRangeByPosition(0,1,7,qLast).DataArray=qData
    If appendCount<=0 Then GoTo Done

    startRow=qLast+1
    If gWMSORD_QueueNextRow>startRow Then startRow=gWMSORD_QueueNextRow
    If startRow<1 Then startRow=1
    ReDim aRows(0 To appendCount-1)
    idx=0
    For i=LBound(aSID) To UBound(aSID)
        If aAppend(i) Then
            target=startRow+idx
            aRows(idx)=Array(stamp,WMSORD_SHEET,aSID(i),r1+i,aMode(i),"PENDING","","ORDER_CHANGED")
            idx=idx+1
        End If
    Next i
    If idx>0 Then oQ.getCellRangeByPosition(0,startRow,7,startRow+idx-1).DataArray=aRows

    ' Update the native in-memory hash only after the vector write. This avoids
    ' 1000+ cross-service Dictionary calls during procurement bulk paste.
    idx=0
    For i=LBound(aSID) To UBound(aSID)
        If aAppend(i) Then
            target=startRow+idx
            If gWMSORD_QueueCacheReady Then Orders_QueueHashPut aSID(i),target
            idx=idx+1
        End If
    Next i
    gWMSORD_QueueNextRow=startRow+idx
Done:
    On Error Resume Next:oQ.IsVisible=False:On Error GoTo 0
    Exit Sub
EH:
    Orders_ResetQueueCache
    Orders_LogError oDoc,"BulkQueue","Не удалось обновить очередь Заказы: " & CStr(Err) & " " & Error$
End Sub

Function Orders_QueueArrayFindPending(aData As Variant,sSourceID As String) As Long
    Dim i As Long,aRow As Variant,status As String
    Orders_QueueArrayFindPending=-1
    On Error GoTo Done
    For i=LBound(aData) To UBound(aData)
        aRow=aData(i)
        If CStr(aRow(1))=WMSORD_SHEET And Trim(CStr(aRow(2)))=sSourceID Then
            status=UCase(Trim(CStr(aRow(5))))
            If status="PENDING" Or status="" Then Orders_QueueArrayFindPending=i+1:Exit Function
        End If
    Next i
Done:
End Function

' ============================================================================
' ROW PROCESSING PIPELINE
' ============================================================================

Sub Orders_ProcessRow(oDoc As Object, oSh As Object, r As Long, bMarkDirty As Boolean, bQueue As Boolean, bCheckIdentity As Boolean, bUpdateDupCache As Boolean)
    Dim hasData As Boolean, sourceID As String, sSeverity As String, sFlags As String, sControl As String
    Dim c As Long, oldState As String, oldDupKey As String, newDupKey As String, nDupPeers As Long
    Dim bCacheWasReady As Boolean,bIdentityHealed As Boolean,healPos As Long,healState As Integer
    If r <= WMSORD_HEADER_ROW Then Exit Sub

    hasData = Orders_RowHasBusinessData(oSh, r)
    sourceID = Orders_CellText(oSh, r, WMSORD_T_ID)
    oldState = Orders_CellText(oSh, r, WMSORD_T_STATE)

    ' A previously identified row that is now cleared still needs identity so that
    ' a future movement engine can safely cancel/reconcile it instead of losing history.
    If Not hasData And sourceID = "" Then
        Orders_ClearAutomaticVisibleFields oSh, r
        Orders_SetTech oSh, r, WMSORD_X_FLAGS, ""
        Orders_SetTech oSh, r, WMSORD_X_SEVERITY, ""
        Orders_SetTech oSh, r, WMSORD_X_VALIDATED, Orders_Stamp(Now)
        Orders_ApplyRowVisual oSh, r
        Exit Sub
    End If

    If sourceID = "" Then
        sourceID = Orders_NewSourceID(r)
        Orders_SetTech oSh, r, WMSORD_T_ID, sourceID
    ElseIf bCheckIdentity Then
        If Orders_SourceIDCount(oSh, sourceID) > 1 Then
            If Orders_Canon(oldState) <> Orders_Canon("Проведено") Then
                sourceID = Orders_NewSourceID(r)
                Orders_SetTech oSh, r, WMSORD_T_ID, sourceID
                bIdentityHealed=True
            Else
                Orders_AppendFlag sFlags, "CRITICAL: дублирующий SourceID у проведённой строки"
            End If
        End If
    End If

    If bIdentityHealed Then
        healState=Orders_OrderPositionState(oSh,r,healPos)
        If healState=1 And healPos=1 Then Orders_ClearGroupContextRow oSh,r
    End If
    If oldState = "" Then Orders_SetTech oSh, r, WMSORD_T_STATE, "Черновик"
    Orders_NormalizeRowDates oDoc, oSh, r
    Orders_EnsureMode oDoc, oSh, r
    Orders_EnsureRowGroupContext oSh, r
    Orders_ApplyReceiptDateAuto oDoc, oSh, r
    Orders_ApplyExactSuggestions oDoc, oSh, r, sFlags
    Orders_ApplyAutomaticStatus oDoc, oSh, r

    oldDupKey = Orders_CellText(oSh, r, WMSORD_X_DUPKEY)
    newDupKey = Orders_DuplicateKey(oSh, r)
    bCacheWasReady = gWMSORD_DupCacheReady
    If Not bCacheWasReady Then Call Orders_EnsureDupCache(oSh)
    If bUpdateDupCache And bCacheWasReady Then Orders_DupCacheMove oldDupKey, newDupKey
    nDupPeers = Orders_DupCachePeerCount(oSh, newDupKey)

    Orders_ValidateRow oDoc, oSh, r, sSeverity, sFlags, sControl, nDupPeers
    Orders_SetCellText oSh, r, "Контроль", sControl
    Orders_SetTech oSh, r, WMSORD_X_FLAGS, sFlags
    Orders_SetTech oSh, r, WMSORD_X_SEVERITY, sSeverity
    Orders_SetTech oSh, r, WMSORD_X_VALIDATED, Orders_Stamp(Now)
    Orders_SetTech oSh, r, WMSORD_X_DUPKEY, newDupKey
    Orders_IncrementRowVersion oSh, r

    If sSeverity = "CRITICAL" Then
        Orders_SetTech oSh, r, WMSORD_T_ERROR, Orders_FirstCritical(sFlags)
    Else
        Orders_SetTech oSh, r, WMSORD_T_ERROR, ""
    End If

    If bMarkDirty Then
        Orders_SetTech oSh, r, WMSORD_T_DIRTY, "DIRTY"
        Orders_SetTech oSh, r, WMSORD_T_TOUCH, Orders_Stamp(Now)
        If bQueue Then Orders_QueueRow oDoc, oSh, r, sourceID
    End If

    Orders_ApplyRowVisual oSh, r
End Sub

Sub Orders_ClearAutomaticVisibleFields(oSh As Object, r As Long)
    Dim mode As String
    mode = Orders_CellText(oSh,r,WMSORD_X_STATUSMODE)
    If mode = "AUTO" Then oSh.getCellByPosition(16,r).String = ""
    oSh.getCellByPosition(18,r).String = ""
End Sub

Sub Orders_NormalizeRowDates(oDoc As Object, oSh As Object, r As Long)
    Orders_NormalizeDateCell oDoc, oSh.getCellByPosition(12,r), True
    Orders_NormalizeDateCell oDoc, oSh.getCellByPosition(13,r), True
    Orders_NormalizeDateCell oDoc, oSh.getCellByPosition(14,r), True
End Sub

Sub Orders_NormalizeDateCell(oDoc As Object, oCell As Object, bInferYear As Boolean)
    Dim d As Double, k As Long
    On Error GoTo Done
    d = Orders_CellDate(oCell, bInferYear)
    If d > 0 Then
        oCell.Value = d
        k = Orders_NumberFormatKey(oDoc, "DATE")
        If k >= 0 Then oCell.NumberFormat = k
    End If
Done:
End Sub

Sub Orders_EnsureMode(oDoc As Object, oSh As Object, r As Long)
    Dim c As Long, mode As String, d As Double, age As Double, lim As Long
    c = Orders_FindHeader(oSh, WMSORD_T_MODE): If c < 0 Then Exit Sub
    mode = UCase(Trim(oSh.getCellByPosition(c,r).String))
    If mode <> "" Then Exit Sub
    d = Orders_RowReferenceDate(oSh,r)
    lim = Orders_LongSetting(oDoc,"ORDERS_LegacyAgeDays",WMSORD_DEFAULT_LEGACY_DAYS)
    If d > 0 Then
        age = CDbl(Date) - d
        If age > lim Then mode = "HISTORY" Else mode = "CURRENT"
    Else
        mode = "CURRENT"
    End If
    oSh.getCellByPosition(c,r).String = mode
    If mode <> "CURRENT" And Orders_CellText(oSh,r,WMSORD_T_LEGACY) = "" Then Orders_SetTech oSh,r,WMSORD_T_LEGACY,Orders_NewLegacyKey(r)
End Sub

Sub Orders_ApplyReceiptDateAuto(oDoc As Object, oSh As Object, r As Long)
    Dim fact As Double, okFact As Boolean, cOrigin As Long, origin As String, d As Double
    okFact = Orders_TryNumber(oSh.getCellByPosition(6,r), fact)
    If Not okFact Or fact <= 0 Then Exit Sub
    d = Orders_CellDate(oSh.getCellByPosition(12,r), True)
    If d > 0 Then Exit Sub
    cOrigin = Orders_FindHeader(oSh, WMSORD_X_DATEORIGIN)
    If cOrigin >= 0 Then origin = UCase(Trim(oSh.getCellByPosition(cOrigin,r).String))
    If origin = "MANUAL" Then Exit Sub
    oSh.getCellByPosition(12,r).Value = CDbl(Date)
    Orders_ApplyNumberFormat oDoc, oSh.getCellByPosition(12,r), "DATE"
    If cOrigin >= 0 Then oSh.getCellByPosition(cOrigin,r).String = "AUTO_FIRST_FACT"
End Sub

Sub Orders_ApplyAutomaticStatus(oDoc As Object, oSh As Object, r As Long)
    Dim cMode As Long, mode As String, status As String, fact As Double, okFact As Boolean, docNo As String, autoStatus As String
    cMode = Orders_FindHeader(oSh, WMSORD_X_STATUSMODE)
    If cMode >= 0 Then mode = UCase(Trim(oSh.getCellByPosition(cMode,r).String))
    status = Trim(oSh.getCellByPosition(16,r).String)

    If status = "" And mode = "MANUAL" Then mode = "AUTO": If cMode >= 0 Then oSh.getCellByPosition(cMode,r).String = "AUTO"
    If mode = "" Then mode = "AUTO": If cMode >= 0 Then oSh.getCellByPosition(cMode,r).String = "AUTO"
    If mode = "MANUAL" Then Exit Sub

    okFact = Orders_TryNumber(oSh.getCellByPosition(6,r), fact)
    docNo = Orders_EffectiveDocNo(oSh,r)
    If Not okFact Or fact <= 0 Then
        autoStatus = WMSORD_STATUS_TRANSIT
    ElseIf docNo = "" Then
        autoStatus = WMSORD_STATUS_NO_DOCS
    Else
        autoStatus = WMSORD_STATUS_RECEIVED
    End If
    ' CRITICAL invariant: automation never sets Оприходовано.
    oSh.getCellByPosition(16,r).String = autoStatus
End Sub

' ============================================================================
' EXACT SUGGESTIONS — CODE / CATEGORY / LOCATION
' ============================================================================

Sub Orders_ApplyExactSuggestions(oDoc As Object, oSh As Object, r As Long, ByRef sFlags As String)
    Dim code As String, cat As String, place As String, itemID As String, source As String
    Dim mapState As String, hCode As String, hCat As String, hPlace As String
    Dim c As Long, origin As String, curr As String

    curr = Trim(oSh.getCellByPosition(21,r).String)
    ' Highest confidence: user's/manual physical code can resolve Nomenclature.
    If curr <> "" Then
        If Orders_ResolveByCode(oDoc,curr,itemID,cat,place) Then source = "PHYSICAL_CODE"
    End If

    If source = "" Then
        mapState = Orders_ResolveSupplierMapping(oDoc,oSh,r,itemID,code,cat,place)
        If mapState = "AMBIGUOUS" Then Orders_AppendFlag sFlags,"WARN: точное сопоставление поставщика неоднозначно"
        If mapState = "FOUND" Then source = "EXACT_MAPPING"
    End If

    ' Fallback learning only from exact same platform+seller+article in existing orders.
    If source = "" Then
        Orders_HistoryExactSuggestions oSh,r,hCode,hCat,hPlace,mapState
        If mapState = "FOUND" Then
            code = hCode: cat = hCat: place = hPlace: source = "EXACT_ORDER_HISTORY"
        ElseIf mapState = "AMBIGUOUS" Then
            Orders_AppendFlag sFlags,"INFO: история точного товара содержит разные код/категорию; авто-подстановка отключена"
        End If
    End If

    Orders_SetTech oSh,r,WMSORD_X_CODESUG,code
    Orders_SetTech oSh,r,WMSORD_X_CATSUG,cat
    Orders_SetTech oSh,r,WMSORD_X_LOCSUG,place
    Orders_SetTech oSh,r,WMSORD_X_SUGSOURCE,source

    ' Product code V: user value always wins.
    c = Orders_FindHeader(oSh,WMSORD_X_CODEORIGIN): origin = "": If c>=0 Then origin=UCase(Trim(oSh.getCellByPosition(c,r).String))
    curr = Trim(oSh.getCellByPosition(21,r).String)
    If curr = "" And origin <> "MANUAL" And code <> "" Then
        oSh.getCellByPosition(21,r).String = code
        If c>=0 Then oSh.getCellByPosition(c,r).String = "AUTO_EXACT"
    ElseIf curr <> "" And code <> "" And Orders_Canon(curr) <> Orders_Canon(code) Then
        Orders_AppendFlag sFlags,"WARN: ручной код отличается от точного предложения '" & code & "'; ручное значение сохранено"
    End If

    ' Category R: user value always wins.
    c = Orders_FindHeader(oSh,WMSORD_X_CATORIGIN): origin = "": If c>=0 Then origin=UCase(Trim(oSh.getCellByPosition(c,r).String))
    curr = Trim(oSh.getCellByPosition(17,r).String)
    If curr = "" And origin <> "MANUAL" And cat <> "" Then
        oSh.getCellByPosition(17,r).String = cat
        If c>=0 Then oSh.getCellByPosition(c,r).String = "AUTO_EXACT"
    ElseIf curr <> "" And cat <> "" And Orders_Canon(curr) <> Orders_Canon(cat) Then
        Orders_AppendFlag sFlags,"WARN: ручная категория отличается от точного предложения '" & cat & "'; ручное значение сохранено"
    End If

    ' Location T remains user controlled by explicit requirement. Only compare/suggest.
    curr = Trim(oSh.getCellByPosition(19,r).String)
    If curr <> "" And place <> "" And Orders_Canon(curr) <> Orders_Canon(place) Then Orders_AppendFlag sFlags,"WARN: место хранения отличается от основного точного места '" & place & "'"
End Sub

Function Orders_ResolveByCode(oDoc As Object, sCode As String, ByRef sItemID As String, ByRef sCat As String, ByRef sPlace As String) As Boolean
    Dim oNom As Object, cID As Long,cCode As Long,cCat As Long,cPlace As Long,r As Long,last As Long,n As Long, id As String
    Orders_ResolveByCode = False: sItemID="":sCat="":sPlace=""
    If Trim(sCode)="" Then Exit Function
    If Not oDoc.Sheets.hasByName(WMSORD_NOM) Then Exit Function
    oNom=oDoc.Sheets.getByName(WMSORD_NOM)
    cID=Orders_FindHeader(oNom,"ID"):cCode=Orders_FindHeader(oNom,"Код"):cCat=Orders_FindHeader(oNom,"Категория"):cPlace=Orders_FindHeader(oNom,"Место")
    If cID<0 Or cCode<0 Then Exit Function
    last=Orders_LastContentRow(oNom,9)
    For r=1 To last
        If Trim(oNom.getCellByPosition(cCode,r).String)<>"" And Orders_Canon(oNom.getCellByPosition(cCode,r).String)=Orders_Canon(sCode) Then
            id=Trim(oNom.getCellByPosition(cID,r).String)
            If id<>"" Then n=n+1:sItemID=id:If cCat>=0 Then sCat=Trim(oNom.getCellByPosition(cCat,r).String):If cPlace>=0 Then sPlace=Trim(oNom.getCellByPosition(cPlace,r).String)
        End If
    Next r
    Orders_ResolveByCode=(n=1)
End Function

Function Orders_ResolveSupplierMapping(oDoc As Object, oOrders As Object, rOrder As Long, ByRef sItemID As String, ByRef sCode As String, ByRef sCat As String, ByRef sPlace As String) As String
    Dim oMap As Object,oNom As Object,p As String,seller As String,art As String
    Dim cID As Long,cCode As Long,cP As Long,cS As Long,cA As Long,cActive As Long,r As Long,last As Long,n As Long,id As String,firstID As String,mapCode As String
    Dim nNom As Long,cnID As Long,cnCode As Long,cnCat As Long,cnPlace As Long,rr As Long,lastNom As Long
    Orders_ResolveSupplierMapping="NONE":sItemID="":sCode="":sCat="":sPlace=""
    p=Trim(oOrders.getCellByPosition(11,rOrder).String):seller=Trim(oOrders.getCellByPosition(22,rOrder).String):art=Trim(oOrders.getCellByPosition(5,rOrder).String)
    If art="" Or p="" Then Exit Function
    If Not oDoc.Sheets.hasByName(WMSORD_MAP) Then Exit Function
    oMap=oDoc.Sheets.getByName(WMSORD_MAP)
    cID=Orders_FindHeader(oMap,"ItemID"):cCode=Orders_FindHeader(oMap,"Код товара"):cP=Orders_FindHeader(oMap,"Поставщик / площадка"):cS=Orders_FindHeader(oMap,"Продавец"):cA=Orders_FindHeader(oMap,"Артикул поставщика"):cActive=Orders_FindHeader(oMap,"Активно")
    If cID<0 Or cP<0 Or cS<0 Or cA<0 Then Exit Function
    last=Orders_LastContentRow(oMap,9)
    For r=1 To last
        If Orders_IsActiveMapping(oMap,cActive,r) Then
            If Orders_Canon(oMap.getCellByPosition(cP,r).String)=Orders_Canon(p) And Orders_Canon(oMap.getCellByPosition(cS,r).String)=Orders_Canon(seller) And Orders_Canon(oMap.getCellByPosition(cA,r).String)=Orders_Canon(art) Then
                id=Trim(oMap.getCellByPosition(cID,r).String)
                If id<>"" Then
                    If firstID="" Then firstID=id
                    If Orders_Canon(firstID)<>Orders_Canon(id) Then Orders_ResolveSupplierMapping="AMBIGUOUS":Exit Function
                    n=n+1
                    If cCode>=0 And Trim(oMap.getCellByPosition(cCode,r).String)<>"" Then
                        If mapCode="" Then
                            mapCode=Trim(oMap.getCellByPosition(cCode,r).String)
                        ElseIf Orders_Canon(mapCode)<>Orders_Canon(oMap.getCellByPosition(cCode,r).String) Then
                            Orders_ResolveSupplierMapping="AMBIGUOUS"
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next r
    If n=0 Then Exit Function
    sItemID=firstID:sCode=mapCode

    If oDoc.Sheets.hasByName(WMSORD_NOM) Then
        oNom=oDoc.Sheets.getByName(WMSORD_NOM)
        cnID=Orders_FindHeader(oNom,"ID"):cnCode=Orders_FindHeader(oNom,"Код"):cnCat=Orders_FindHeader(oNom,"Категория"):cnPlace=Orders_FindHeader(oNom,"Место")
        If cnID>=0 Then
            lastNom=Orders_LastContentRow(oNom,9)
            For rr=1 To lastNom
                If Orders_Canon(oNom.getCellByPosition(cnID,rr).String)=Orders_Canon(firstID) And Trim(oNom.getCellByPosition(cnID,rr).String)<>"" Then
                    nNom=nNom+1
                    If sCode="" And cnCode>=0 Then sCode=Trim(oNom.getCellByPosition(cnCode,rr).String)
                    If cnCat>=0 Then sCat=Trim(oNom.getCellByPosition(cnCat,rr).String)
                    If cnPlace>=0 Then sPlace=Trim(oNom.getCellByPosition(cnPlace,rr).String)
                End If
            Next rr
            If nNom>1 Then Orders_ResolveSupplierMapping="AMBIGUOUS":Exit Function
        End If
    End If
    Orders_ResolveSupplierMapping="FOUND"
End Function

Function Orders_IsActiveMapping(oMap As Object, cActive As Long, r As Long) As Boolean
    Dim s As String
    If cActive<0 Then Orders_IsActiveMapping=True:Exit Function
    s=Orders_Canon(oMap.getCellByPosition(cActive,r).String)
    Orders_IsActiveMapping=(s<>"НЕТ" And s<>"FALSE" And s<>"0" And s<>"НЕАКТИВНО")
End Function

Sub Orders_ResetHistoryCache()
    On Error Resume Next
    If gWMSORD_HistoryCacheReady Then gWMSORD_HistoryCache.Dispose
    gWMSORD_HistoryCache=Empty
    gWMSORD_HistoryCacheReady=False
    On Error GoTo 0
End Sub

Function Orders_HistoryKey(oSh As Object,r As Long) As String
    Dim p As String,seller As String,art As String
    p=Orders_Canon(oSh.getCellByPosition(11,r).String)
    seller=Orders_Canon(oSh.getCellByPosition(22,r).String)
    art=Orders_Canon(oSh.getCellByPosition(5,r).String)
    If p="" Or art="" Then Exit Function
    Orders_HistoryKey=p & Chr(31) & seller & Chr(31) & art
End Function

Function Orders_HistoryPack(sCode As String,sCat As String,sPlace As String,bAmb As Boolean) As String
    Orders_HistoryPack=sCode & Chr(30) & sCat & Chr(30) & sPlace & Chr(30) & IIf(bAmb,"1","0")
End Function

Sub Orders_HistoryUnpack(sPacked As String,ByRef sCode As String,ByRef sCat As String,ByRef sPlace As String,ByRef bAmb As Boolean)
    Dim a As Variant
    sCode="":sCat="":sPlace="":bAmb=False
    a=Split(sPacked,Chr(30))
    If UBound(a)>=0 Then sCode=CStr(a(0))
    If UBound(a)>=1 Then sCat=CStr(a(1))
    If UBound(a)>=2 Then sPlace=CStr(a(2))
    If UBound(a)>=3 Then bAmb=(CStr(a(3))="1")
End Sub

Function Orders_EnsureHistoryCache(oSh As Object) As Boolean
    Dim last As Long,r As Long,k As String,packed As String
    Dim codeSeen As String,catSeen As String,placeSeen As String,v As String,bAmb As Boolean
    Orders_EnsureHistoryCache=False
    On Error GoTo EH
    If gWMSORD_HistoryCacheReady Then Orders_EnsureHistoryCache=True:Exit Function
    GlobalScope.BasicLibraries.LoadLibrary("ScriptForge")
    gWMSORD_HistoryCache=CreateScriptService("Dictionary",False)
    last=Orders_LastContentRow(oSh,WMSORD_USER_LAST_COL)
    For r=1 To last
        k=Orders_HistoryKey(oSh,r)
        If k<>"" Then
            codeSeen="":catSeen="":placeSeen="":bAmb=False
            If gWMSORD_HistoryCache.Exists(k) Then
                packed=CStr(gWMSORD_HistoryCache.Item(k))
                Orders_HistoryUnpack packed,codeSeen,catSeen,placeSeen,bAmb
            End If
            v=Trim(oSh.getCellByPosition(21,r).String)
            If v<>"" Then
                If codeSeen="" Then
                    codeSeen=v
                ElseIf Orders_Canon(codeSeen)<>Orders_Canon(v) Then
                    bAmb=True
                End If
            End If
            v=Trim(oSh.getCellByPosition(17,r).String)
            If v<>"" Then
                If catSeen="" Then
                    catSeen=v
                ElseIf Orders_Canon(catSeen)<>Orders_Canon(v) Then
                    bAmb=True
                End If
            End If
            v=Trim(oSh.getCellByPosition(19,r).String)
            If v<>"" Then
                If placeSeen="" Then
                    placeSeen=v
                ElseIf Orders_Canon(placeSeen)<>Orders_Canon(v) Then
                    bAmb=True
                End If
            End If
            packed=Orders_HistoryPack(codeSeen,catSeen,placeSeen,bAmb)
            If gWMSORD_HistoryCache.Exists(k) Then
                gWMSORD_HistoryCache.ReplaceItem k,packed
            Else
                gWMSORD_HistoryCache.Add k,packed
            End If
        End If
    Next r
    gWMSORD_HistoryCacheReady=True
    Orders_EnsureHistoryCache=True
    Exit Function
EH:
    gWMSORD_HistoryCacheReady=False
    Orders_LogError ThisComponent,"HistoryCache","Не удалось построить индекс точной истории заказов: " & CStr(Err) & " " & Error$
End Function

Sub Orders_HistoryExactSuggestions(oSh As Object, rOrder As Long, ByRef sCode As String, ByRef sCat As String, ByRef sPlace As String, ByRef sState As String)
    Dim k As String,packed As String,bAmb As Boolean
    sCode="":sCat="":sPlace="":sState="NONE"
    k=Orders_HistoryKey(oSh,rOrder)
    If k="" Then Exit Sub
    If Not Orders_EnsureHistoryCache(oSh) Then Exit Sub
    On Error GoTo Done
    If Not gWMSORD_HistoryCache.Exists(k) Then Exit Sub
    packed=CStr(gWMSORD_HistoryCache.Item(k))
    Orders_HistoryUnpack packed,sCode,sCat,sPlace,bAmb
    If bAmb Then
        sState="AMBIGUOUS"
    ElseIf sCode<>"" Or sCat<>"" Or sPlace<>"" Then
        sState="FOUND"
    End If
Done:
End Sub

' ============================================================================
' BUSINESS VALIDATION / CONTROL
' ============================================================================

Sub Orders_ValidateRow(oDoc As Object,oSh As Object,r As Long,ByRef sSeverity As String,ByRef sFlags As String,ByRef sControl As String,nDuplicatePeers As Long)
    Dim name As String,docNo As String,invoice As String,supCode As String,art As String,unit As String,platform As String,seller As String
    Dim buyer As String,status As String,cat As String,place As String,code As String
    Dim ordered As Double,fact As Double,price As Double,total As Double,okOrdered As Boolean,okFact As Boolean,okPrice As Boolean,okTotal As Boolean
    Dim dArr As Double,dDoc As Double,dOrder As Double,expected As Double,age As Long
    Dim firstCritical As String,firstWarn As String,firstInfo As String,manualMode As String
    Dim orderPos As Long,posState As Integer

    sSeverity="OK":sControl="": If sFlags="" Then sFlags=""
    name=Trim(oSh.getCellByPosition(1,r).String):docNo=Orders_EffectiveDocNo(oSh,r):invoice=Orders_EffectiveInvoice(oSh,r)
    supCode=Trim(oSh.getCellByPosition(4,r).String):art=Trim(oSh.getCellByPosition(5,r).String):unit=Trim(oSh.getCellByPosition(8,r).String)
    platform=Orders_EffectiveSupplier(oSh,r):seller=Trim(oSh.getCellByPosition(22,r).String):buyer=Trim(oSh.getCellByPosition(15,r).String)
    status=Trim(oSh.getCellByPosition(16,r).String):cat=Trim(oSh.getCellByPosition(17,r).String):place=Trim(oSh.getCellByPosition(19,r).String):code=Trim(oSh.getCellByPosition(21,r).String)
    okOrdered=Orders_TryNumber(oSh.getCellByPosition(7,r),ordered):okFact=Orders_TryNumber(oSh.getCellByPosition(6,r),fact):okPrice=Orders_TryNumber(oSh.getCellByPosition(9,r),price):okTotal=Orders_TryNumber(oSh.getCellByPosition(10,r),total)
    dArr=Orders_CellDate(oSh.getCellByPosition(12,r),True):dDoc=Orders_EffectiveDocDate(oSh,r):dOrder=Orders_CellDate(oSh.getCellByPosition(14,r),True)
    Dim groupIssue As String: groupIssue=Orders_CellText(oSh,r,WMSORD_X_GROUPFLAGS)
    posState=Orders_OrderPositionState(oSh,r,orderPos)

    If posState=0 Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Не заполнен номер позиции заказа в колонке Номер"
    If posState<0 Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"CRITICAL","Колонка Номер должна содержать целое положительное число: 1, 2, 3..."
    If groupIssue<>"" Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"CRITICAL",groupIssue
    If name="" Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"CRITICAL","Заполните наименование товара"
    If Not okOrdered Or ordered<=0 Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"CRITICAL","Проверьте заказанное количество"
    If unit="" Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Не заполнена единица измерения"
    If invoice="" Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Не заполнен номер счета"
    If platform="" Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Не заполнен поставщик / площадка"
    If dOrder<=0 Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Не заполнена дата заказа"

    If Trim(oSh.getCellByPosition(6,r).String)<>"" Or Abs(oSh.getCellByPosition(6,r).Value)>WMSORD_EPS Then
        If Not okFact Or fact<0 Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"CRITICAL","Факт. количество должно быть числом не меньше нуля"
    End If
    If okFact And okOrdered Then
        If fact>ordered+WMSORD_EPS Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"CRITICAL","Факт больше заказанного — проверьте количество"
        If fact>WMSORD_EPS And fact<ordered-WMSORD_EPS Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Факт меньше заказа; остаток может приехать другим документом"
    End If

    If okFact And fact>WMSORD_EPS Then
        If dArr<=0 Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Укажите дату поступления"
        If place="" Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Укажите место хранения"
        If cat="" Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Укажите категорию"
        If code="" Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"INFO","У позиции пока нет кода товара"
        If docNo="" Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Товар получен без документов"
    ElseIf dArr>0 Then
        Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Дата поступления есть, а положительного факта нет"
    End If

    If docNo<>"" And dDoc<=0 Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Есть № документа, но нет даты документа"
    If docNo="" And dDoc>0 Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Есть дата документа, но нет № документа"
    If dOrder>CDbl(Date)+WMSORD_EPS Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Дата заказа находится в будущем"
    If dArr>CDbl(Date)+WMSORD_EPS Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Дата поступления находится в будущем"
    If dDoc>CDbl(Date)+WMSORD_EPS Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Дата документа находится в будущем"
    If dOrder>0 And dArr>0 And dArr+WMSORD_EPS<dOrder Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Дата поступления раньше даты заказа"

    If Orders_IsMarketplace(platform) And seller="" Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Для маркетплейса укажите продавца"

    If okOrdered And okPrice And okTotal And ordered>=0 And price>=0 Then
        expected=ordered*price
        If Abs(total-expected)>0.02 Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Сумма не равна Количество × Цена"
    ElseIf okOrdered And okPrice And (Trim(oSh.getCellByPosition(10,r).String)="" And Abs(oSh.getCellByPosition(10,r).Value)<WMSORD_EPS) Then
        Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"INFO","Сумма не заполнена"
    End If

    ' Status consistency. Automation never auto-posts.
    If Orders_Canon(status)=Orders_Canon(WMSORD_STATUS_POST) Then
        If Not okFact Or fact<=0 Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"CRITICAL","Оприходовано требует положительный факт"
        If docNo="" Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"CRITICAL","Оприходование без № документа запрещено"
    ElseIf Orders_Canon(status)=Orders_Canon(WMSORD_STATUS_CANCEL) Then
        Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"INFO","Заказ отменён вручную"
    ElseIf Orders_Canon(status)=Orders_Canon(WMSORD_STATUS_TRANSIT) And okFact And fact>0 Then
        Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Статус В пути не соответствует положительному факту"
    ElseIf Orders_Canon(status)=Orders_Canon(WMSORD_STATUS_NO_DOCS) And docNo<>"" Then
        Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Документ уже указан, а статус оставлен Получено без документов"
    ElseIf Orders_Canon(status)=Orders_Canon(WMSORD_STATUS_RECEIVED) Then
        If Not okFact Or fact<=0 Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Статус Получено без положительного факта"
        If docNo="" Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Статус Получено без № документа"
    End If

    If nDuplicatePeers>0 Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Найдена другая строка с тем же точным ключом заказа — проверьте дубль"

    If dOrder>0 And (Not okFact Or fact<=0) Then
        age=CLng(CDbl(Date)-dOrder)
        If age>Orders_LongSetting(oDoc,"ORDERS_LegacyAgeDays",WMSORD_DEFAULT_LEGACY_DAYS) Then Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"WARN","Заказ остаётся без факта " & CStr(age) & " дн."
    End If

    If Orders_Canon(Orders_CellText(oSh,r,WMSORD_T_STATE))=Orders_Canon("Проведено") And Orders_Canon(Orders_CellText(oSh,r,WMSORD_T_DIRTY))=Orders_Canon("DIRTY") Then
        Orders_AddIssue sFlags,firstCritical,firstWarn,firstInfo,"CRITICAL","Проведённая строка изменена — требуется безопасное отмена/перепроведение"
    End If

    If firstCritical<>"" Then sSeverity="CRITICAL":sControl=firstCritical:Exit Sub
    If firstWarn<>"" Then sSeverity="WARN":sControl=firstWarn:Exit Sub
    If firstInfo<>"" Then sSeverity="INFO":sControl=firstInfo:Exit Sub

    ' No issues: control reflects the workflow state, not generic "OK".
    If Orders_Canon(status)=Orders_Canon(WMSORD_STATUS_CANCEL) Then
        sSeverity="INFO":sControl="Отменено"
    ElseIf Orders_Canon(status)=Orders_Canon(WMSORD_STATUS_POST) Then
        sSeverity="OK":sControl="Готово к проведению"
    ElseIf okFact And fact>0 And docNo="" Then
        sSeverity="WARN":sControl="Ждём документы"
    ElseIf okFact And fact>0 And docNo<>"" Then
        sSeverity="OK":sControl="Получено — проверьте и оприходуйте"
    Else
        sSeverity="INFO":sControl="В пути"
    End If
End Sub

Sub Orders_AddIssue(ByRef flags As String,ByRef firstCritical As String,ByRef firstWarn As String,ByRef firstInfo As String,sLevel As String,sMessage As String)
    Orders_AppendFlag flags,sLevel & ": " & sMessage
    Select Case UCase(sLevel)
        Case "CRITICAL": If firstCritical="" Then firstCritical=sMessage
        Case "WARN": If firstWarn="" Then firstWarn=sMessage
        Case Else: If firstInfo="" Then firstInfo=sMessage
    End Select
End Sub

Sub Orders_AppendFlag(ByRef sFlags As String,sText As String)
    If Trim(sText)="" Then Exit Sub
    If sFlags="" Then sFlags=sText Else sFlags=sFlags & " | " & sText
End Sub

Function Orders_FirstCritical(sFlags As String) As String
    Dim a As Variant,i As Long,s As String
    a=Split(sFlags," | ")
    For i=LBound(a) To UBound(a)
        s=Trim(CStr(a(i)))
        If Left(UCase(s),9)="CRITICAL:" Then Orders_FirstCritical=Mid(s,10):Exit Function
    Next i
End Function

Function Orders_IsMarketplace(sPlatform As String) As Boolean
    Dim c As String
    c=Orders_Canon(sPlatform)
    Orders_IsMarketplace=(InStr(c,"OZON")>0 Or InStr(c,"WILDBERRIES")>0 Or InStr(c,"ВАЙЛДБЕРРИЗ")>0 Or InStr(c,"ЯНДЕКС МАРКЕТ")>0 Or InStr(c,"ЯНДЕКСМАРКЕТ")>0)
End Function

Function Orders_DuplicateKey(oSh As Object,r As Long) As String
    Dim no As String,invoice As String,sup As String,art As String,buyer As String,d As Double,q As Double,ok As Boolean
    no=Trim(oSh.getCellByPosition(0,r).String):invoice=Orders_EffectiveInvoice(oSh,r):sup=Trim(oSh.getCellByPosition(4,r).String):art=Trim(oSh.getCellByPosition(5,r).String):buyer=Trim(oSh.getCellByPosition(15,r).String)
    d=Orders_CellDate(oSh.getCellByPosition(14,r),True):ok=Orders_TryNumber(oSh.getCellByPosition(7,r),q)
    ' Conservative exact key: only when enough identifiers exist. It is a warning,
    ' never automatic deletion/merge.
    If no="" Or invoice="" Or art="" Or d<=0 Or Not ok Then Exit Function
    Orders_DuplicateKey=Orders_Canon(no) & "|" & Orders_Canon(invoice) & "|" & Orders_Canon(sup) & "|" & Orders_Canon(art) & "|" & Orders_NumKey(q) & "|" & CStr(CLng(d)) & "|" & Orders_Canon(buyer)
End Function

Sub Orders_ResetDupCache()
    On Error Resume Next
    If gWMSORD_DupCacheReady Then gWMSORD_DupCache.Dispose
    gWMSORD_DupCache = Empty
    gWMSORD_DupCacheReady = False
    On Error GoTo 0
End Sub

Function Orders_EnsureDupCache(oSh As Object) As Boolean
    Dim last As Long,r As Long,k As String,n As Long
    Orders_EnsureDupCache=False
    On Error GoTo EH
    If gWMSORD_DupCacheReady Then Orders_EnsureDupCache=True:Exit Function
    GlobalScope.BasicLibraries.LoadLibrary("ScriptForge")
    gWMSORD_DupCache = CreateScriptService("Dictionary", False)
    last=Orders_LastContentRow(oSh,WMSORD_USER_LAST_COL)
    For r=1 To last
        If Orders_RowHasBusinessData(oSh,r) Then
            k=Orders_DuplicateKey(oSh,r)
            If k<>"" Then
                If gWMSORD_DupCache.Exists(k) Then
                    n=CLng(gWMSORD_DupCache.Item(k))
                    gWMSORD_DupCache.ReplaceItem k,n+1
                Else
                    gWMSORD_DupCache.Add k,1
                End If
            End If
        End If
    Next r
    gWMSORD_DupCacheReady=True
    Orders_EnsureDupCache=True
    Exit Function
EH:
    gWMSORD_DupCacheReady=False
    Orders_LogError ThisComponent,"DupCache","Не удалось построить индекс точных дублей: " & CStr(Err) & " " & Error$
End Function

Sub Orders_DupCacheMove(sOldKey As String,sNewKey As String)
    Dim n As Long
    If Not gWMSORD_DupCacheReady Then Exit Sub
    On Error GoTo EH
    If sOldKey=sNewKey Then Exit Sub
    If sOldKey<>"" Then
        If gWMSORD_DupCache.Exists(sOldKey) Then
            n=CLng(gWMSORD_DupCache.Item(sOldKey))
            If n<=1 Then
                gWMSORD_DupCache.Remove sOldKey
            Else
                gWMSORD_DupCache.ReplaceItem sOldKey,n-1
            End If
        End If
    End If
    If sNewKey<>"" Then
        If gWMSORD_DupCache.Exists(sNewKey) Then
            n=CLng(gWMSORD_DupCache.Item(sNewKey))
            gWMSORD_DupCache.ReplaceItem sNewKey,n+1
        Else
            gWMSORD_DupCache.Add sNewKey,1
        End If
    End If
    Exit Sub
EH:
    Orders_ResetDupCache
End Sub

Function Orders_DupCachePeerCount(oSh As Object,sKey As String) As Long
    Dim n As Long
    Orders_DupCachePeerCount=0
    If sKey="" Then Exit Function
    If Not gWMSORD_DupCacheReady Then
        If Not Orders_EnsureDupCache(oSh) Then Exit Function
    End If
    On Error GoTo EH
    If gWMSORD_DupCache.Exists(sKey) Then
        n=CLng(gWMSORD_DupCache.Item(sKey))
        If n>1 Then Orders_DupCachePeerCount=n-1
    End If
    Exit Function
EH:
    Orders_ResetDupCache
End Function

Function Orders_DuplicateKeyCount(oSh As Object,sKey As String,rSelf As Long) As Long
    Dim r As Long,last As Long,n As Long
    If sKey="" Then Exit Function
    last=Orders_LastContentRow(oSh,WMSORD_USER_LAST_COL)
    For r=1 To last
        If r<>rSelf Then If Orders_DuplicateKey(oSh,r)=sKey Then n=n+1
    Next r
    Orders_DuplicateKeyCount=n
End Function

' ============================================================================
' VISUAL STATE OF ROW
' ============================================================================

Sub Orders_ApplyRowVisual(oSh As Object,r As Long)
    Dim status As String,severity As String,codeOrigin As String,catOrigin As String,dateOrigin As String
    Dim c As Long
    If r<1 Then Exit Sub
    status=Trim(oSh.getCellByPosition(16,r).String)
    severity=UCase(Orders_CellText(oSh,r,WMSORD_X_SEVERITY))
    codeOrigin=UCase(Orders_CellText(oSh,r,WMSORD_X_CODEORIGIN))
    catOrigin=UCase(Orders_CellText(oSh,r,WMSORD_X_CATORIGIN))
    dateOrigin=UCase(Orders_CellText(oSh,r,WMSORD_X_DATEORIGIN))

    ' Order block background is the primary visual layer. It overwrites pasted/imported fills.
    Orders_ApplyOrderBlockBackground oSh,r

    ' Status Q.
    Select Case Orders_Canon(status)
        Case "В ПУТИ"
            Orders_SetCellVisual oSh.getCellByPosition(16,r),RGB(37,99,235),RGB(255,255,255),150
        Case "ПОЛУЧЕНО БЕЗ ДОКУМЕНТОВ"
            Orders_SetCellVisual oSh.getCellByPosition(16,r),RGB(245,158,11),RGB(69,26,3),150
        Case "ПОЛУЧЕНО"
            Orders_SetCellVisual oSh.getCellByPosition(16,r),RGB(22,163,74),RGB(255,255,255),150
        Case "ОПРИХОДОВАНО"
            Orders_SetCellVisual oSh.getCellByPosition(16,r),RGB(21,128,61),RGB(255,255,255),150
        Case "ОТМЕНЕНО"
            Orders_SetCellVisual oSh.getCellByPosition(16,r),RGB(100,116,139),RGB(255,255,255),150
        Case Else
            Orders_SetCellVisual oSh.getCellByPosition(16,r),RGB(246,250,255),RGB(30,41,59),110
    End Select

    ' Automatic Control S.
    Select Case severity
        Case "CRITICAL": Orders_SetCellVisual oSh.getCellByPosition(18,r),RGB(220,38,38),RGB(255,255,255),150
        Case "WARN": Orders_SetCellVisual oSh.getCellByPosition(18,r),RGB(245,158,11),RGB(69,26,3),150
        Case "OK": Orders_SetCellVisual oSh.getCellByPosition(18,r),RGB(22,163,74),RGB(255,255,255),150
        Case "INFO": Orders_SetCellVisual oSh.getCellByPosition(18,r),RGB(37,99,235),RGB(255,255,255),150
        Case Else: Orders_SetCellVisual oSh.getCellByPosition(18,r),RGB(245,247,250),RGB(71,85,105),110
    End Select

    ' Hybrid fields tell user whether current value was automatic or manual.
    If Left(codeOrigin,4)="AUTO" Then
        Orders_SetCellVisual oSh.getCellByPosition(21,r),RGB(235,247,239),RGB(21,94,59),110
    Else
        Orders_SetCellVisual oSh.getCellByPosition(21,r),RGB(246,250,255),RGB(30,41,59),110
    End If
    If Left(catOrigin,4)="AUTO" Then
        Orders_SetCellVisual oSh.getCellByPosition(17,r),RGB(235,247,239),RGB(21,94,59),110
    Else
        Orders_SetCellVisual oSh.getCellByPosition(17,r),RGB(246,250,255),RGB(30,41,59),110
    End If
    If Left(dateOrigin,4)="AUTO" Then
        Orders_SetCellVisual oSh.getCellByPosition(12,r),RGB(235,247,239),RGB(21,94,59),110
    Else
        Orders_SetCellVisual oSh.getCellByPosition(12,r),RGB(242,248,255),RGB(30,41,59),110
    End If

    ' Finally overlay problem cells so a warning/critical state is visible above the order tint.
    Orders_HighlightProblemCells oSh,r
End Sub

Sub Orders_SetCellVisual(oCell As Object,nBack As Long,nFont As Long,nWeight As Double)
    On Error Resume Next
    oCell.CellBackColor=nBack:oCell.CharColor=nFont:oCell.CharWeight=nWeight
    On Error GoTo 0
End Sub

' ============================================================================
' QUEUE / DIRTY / SOURCE ID
' ============================================================================

Sub Orders_ResetQueueCache()
    On Error Resume Next
    ReDim gWMSORD_QueueHashKeys(0 To 0)
    ReDim gWMSORD_QueueHashRows(0 To 0)
    gWMSORD_QueueHashCap=0
    gWMSORD_QueueHashCount=0
    gWMSORD_QueueCacheReady=False
    gWMSORD_QueueNextRow=1
    On Error GoTo 0
End Sub

Sub Orders_QueueHashInit(nExpected As Long)
    Dim cap As Long
    cap=64
    If nExpected<1 Then nExpected=1
    Do While cap<nExpected*2
        cap=cap*2
        If cap>1048576 Then Exit Do
    Loop
    ReDim gWMSORD_QueueHashKeys(0 To cap-1)
    ReDim gWMSORD_QueueHashRows(0 To cap-1)
    gWMSORD_QueueHashCap=cap
    gWMSORD_QueueHashCount=0
End Sub

Function Orders_QueueHashBase(sKey As String,nCap As Long) As Long
    Dim h As Double,i As Long
    h=5381
    For i=1 To Len(sKey)
        h=h*33+Asc(Mid(sKey,i,1))
        If h>1000000000 Then h=h-Int(h/1000000007)*1000000007
    Next i
    If h<0 Then h=-h
    Orders_QueueHashBase=CLng(h) Mod nCap
End Function

Function Orders_QueueHashSlot(sKey As String,ByRef bFound As Boolean) As Long
    Dim idx As Long,probe As Long,k As String
    Orders_QueueHashSlot=-1:bFound=False
    If gWMSORD_QueueHashCap<=0 Or sKey="" Then Exit Function
    idx=Orders_QueueHashBase(sKey,gWMSORD_QueueHashCap)
    For probe=0 To gWMSORD_QueueHashCap-1
        k=gWMSORD_QueueHashKeys(idx)
        If k="" Then Orders_QueueHashSlot=idx:Exit Function
        If k=sKey Then bFound=True:Orders_QueueHashSlot=idx:Exit Function
        idx=idx+1:If idx>=gWMSORD_QueueHashCap Then idx=0
    Next probe
End Function

Sub Orders_QueueHashGrow(nNeeded As Long)
    Dim oldKeys() As String,oldRows() As Long,oldCap As Long,i As Long,k As String,r As Long,newCap As Long,bFound As Boolean,slot As Long
    If gWMSORD_QueueHashCap>0 Then
        If nNeeded*10<=gWMSORD_QueueHashCap*6 Then Exit Sub
    End If
    oldCap=gWMSORD_QueueHashCap
    If oldCap>0 Then
        ReDim oldKeys(0 To oldCap-1):ReDim oldRows(0 To oldCap-1)
        For i=0 To oldCap-1:oldKeys(i)=gWMSORD_QueueHashKeys(i):oldRows(i)=gWMSORD_QueueHashRows(i):Next i
    End If
    newCap=64:Do While newCap<nNeeded*2:newCap=newCap*2:Loop
    ReDim gWMSORD_QueueHashKeys(0 To newCap-1):ReDim gWMSORD_QueueHashRows(0 To newCap-1)
    gWMSORD_QueueHashCap=newCap:gWMSORD_QueueHashCount=0
    If oldCap>0 Then
        For i=0 To oldCap-1
            k=oldKeys(i):r=oldRows(i)
            If k<>"" Then
                slot=Orders_QueueHashSlot(k,bFound)
                If slot>=0 Then gWMSORD_QueueHashKeys(slot)=k:gWMSORD_QueueHashRows(slot)=r:gWMSORD_QueueHashCount=gWMSORD_QueueHashCount+1
            End If
        Next i
    End If
End Sub

Function Orders_QueueHashGet(sKey As String) As Long
    Dim bFound As Boolean,slot As Long
    Orders_QueueHashGet=-1
    If Not gWMSORD_QueueCacheReady Then Exit Function
    slot=Orders_QueueHashSlot(sKey,bFound)
    If bFound And slot>=0 Then Orders_QueueHashGet=gWMSORD_QueueHashRows(slot)
End Function

Sub Orders_QueueHashPut(sKey As String,nRow As Long)
    Dim bFound As Boolean,slot As Long
    If sKey="" Then Exit Sub
    If gWMSORD_QueueHashCap<=0 Then Orders_QueueHashInit 64
    Orders_QueueHashGrow gWMSORD_QueueHashCount+1
    slot=Orders_QueueHashSlot(sKey,bFound)
    If slot<0 Then Exit Sub
    If Not bFound Then gWMSORD_QueueHashKeys(slot)=sKey:gWMSORD_QueueHashCount=gWMSORD_QueueHashCount+1
    gWMSORD_QueueHashRows(slot)=nRow
End Sub

Function Orders_QueueLastRowFast(oQ As Object) As Long
    Dim cur As Object,a As Variant,nEnd As Long,data As Variant,rowData As Variant,r As Long,c As Long
    Orders_QueueLastRowFast=0
    On Error GoTo Done
    cur=oQ.createCursor():cur.gotoEndOfUsedArea(True):a=cur.RangeAddress:nEnd=a.EndRow
    If nEnd<1 Then Exit Function
    data=oQ.getCellRangeByPosition(0,0,7,nEnd).DataArray
    For r=UBound(data) To 1 Step -1
        rowData=data(r)
        For c=0 To 7
            If Trim(CStr(rowData(c)))<>"" Then Orders_QueueLastRowFast=r:Exit Function
        Next c
    Next r
Done:
End Function

Function Orders_EnsureQueueCache(oDoc As Object) As Boolean
    Dim oQ As Object,last As Long,rr As Long,sID As String,sSheet As String,status As String,data As Variant,rowData As Variant
    Orders_EnsureQueueCache=False
    On Error GoTo EH
    If gWMSORD_QueueCacheReady Then Orders_EnsureQueueCache=True:Exit Function
    If Not oDoc.Sheets.hasByName(WMSORD_QUEUE) Then Orders_QueueHashInit 64:gWMSORD_QueueCacheReady=True:Orders_EnsureQueueCache=True:Exit Function
    oQ=oDoc.Sheets.getByName(WMSORD_QUEUE)
    Orders_EnsureQueueHeaders oQ
    last=Orders_QueueLastRowFast(oQ)
    gWMSORD_QueueNextRow=last+1:If gWMSORD_QueueNextRow<1 Then gWMSORD_QueueNextRow=1
    Orders_QueueHashInit last+64
    If last>=1 Then
        data=oQ.getCellRangeByPosition(0,1,7,last).DataArray
        For rr=0 To UBound(data)
            rowData=data(rr)
            sSheet=Trim(CStr(rowData(1))):sID=Trim(CStr(rowData(2))):status=UCase(Trim(CStr(rowData(5))))
            If sSheet=WMSORD_SHEET And sID<>"" And (status="PENDING" Or status="") Then Orders_QueueHashPut sID,rr+1
        Next rr
    End If
    gWMSORD_QueueCacheReady=True
    Orders_EnsureQueueCache=True
    Exit Function
EH:
    Orders_ResetQueueCache
    Orders_LogError oDoc,"QueueCache","Не удалось построить индекс очереди: " & CStr(Err) & " " & Error$
End Function

Sub Orders_QueueRow(oDoc As Object,oSh As Object,r As Long,sSourceID As String)
    Dim oQ As Object,found As Long,mode As String,status As String
    If Trim(sSourceID)="" Then Exit Sub
    If Not oDoc.Sheets.hasByName(WMSORD_QUEUE) Then Exit Sub
    On Error GoTo EH
    oQ=oDoc.Sheets.getByName(WMSORD_QUEUE)
    Orders_EnsureQueueHeaders oQ
    Call Orders_EnsureQueueCache(oDoc)
    found=Orders_QueueHashGet(sSourceID)
    If found>=1 Then
        status=UCase(Trim(oQ.getCellByPosition(5,found).String))
        If Trim(oQ.getCellByPosition(2,found).String)<>sSourceID Or Trim(oQ.getCellByPosition(1,found).String)<>WMSORD_SHEET Or (status<>"PENDING" And status<>"") Then found=-1
    End If
    mode=Orders_CellText(oSh,r,WMSORD_T_MODE):If mode="" Then mode="CURRENT"
    If found<1 Then
        found=gWMSORD_QueueNextRow:If found<1 Then found=1
        Do While Trim(oQ.getCellByPosition(0,found).String)<>"" Or Trim(oQ.getCellByPosition(2,found).String)<>"":found=found+1:Loop
        gWMSORD_QueueNextRow=found+1
        If gWMSORD_QueueCacheReady Then Orders_QueueHashPut sSourceID,found
    End If
    oQ.getCellRangeByPosition(0,found,7,found).DataArray=Array(Array(Orders_Stamp(Now),WMSORD_SHEET,sSourceID,r,mode,"PENDING","","ORDER_CHANGED"))
    On Error Resume Next:oQ.IsVisible=False:On Error GoTo 0
    Exit Sub
EH:
    Orders_ResetQueueCache
    Orders_LogError oDoc,"QueueRow","Не удалось поставить заказ в очередь: " & CStr(Err) & " " & Error$
End Sub

Sub Orders_EnsureQueueHeaders(oQ As Object)
    Dim a As Variant,i As Long
    a=Array("Время","Лист","SourceID","RowHint","Mode","Status","RunID","Message")
    For i=0 To UBound(a)
        If Trim(oQ.getCellByPosition(i,0).String)="" Then oQ.getCellByPosition(i,0).String=CStr(a(i))
    Next i
End Sub

Function Orders_NewSourceID(r As Long) As String
    Randomize
    Orders_NewSourceID="ORD-" & Orders_TimestampCompact(Now) & "-" & Right("00000" & CStr(r+1),5) & "-" & Right("000000" & CStr(CLng(Rnd()*999999)),6)
End Function

Function Orders_NewLegacyKey(r As Long) As String
    Randomize
    Orders_NewLegacyKey="LEG-ORD-" & Orders_TimestampCompact(Now) & "-" & Right("00000" & CStr(r+1),5) & "-" & Right("0000" & CStr(CLng(Rnd()*9999)),4)
End Function

Function Orders_SourceIDCount(oSh As Object,sID As String) As Long
    Dim c As Long,r As Long,last As Long,n As Long
    If Trim(sID)="" Then Exit Function
    c=Orders_FindHeader(oSh,WMSORD_T_ID):If c<0 Then Exit Function
    last=Orders_LastContentRow(oSh,WMSORD_USER_LAST_COL)
    For r=1 To last
        If Trim(oSh.getCellByPosition(c,r).String)=sID Then n=n+1
    Next r
    Orders_SourceIDCount=n
End Function

Sub Orders_IncrementRowVersion(oSh As Object,r As Long)
    Dim c As Long,n As Double
    c=Orders_FindHeader(oSh,WMSORD_X_ROWVER):If c<0 Then Exit Sub
    n=oSh.getCellByPosition(c,r).Value
    If n<0 Then n=0
    oSh.getCellByPosition(c,r).Value=n+1
End Sub

' ============================================================================
' FULL REVALIDATION / BULK QA
' ============================================================================

Function Orders_RevalidateAllCore(oDoc As Object,bMarkDirty As Boolean,ByRef sReport As String) As Boolean
    Dim oSh As Object,last As Long,r As Long,nData As Long,nCrit As Long,nWarn As Long,nInfo As Long,cFp As Long,nClean As Long
    Dim sev As String,fp As String,bLocked As Boolean,sErr As String
    Orders_RevalidateAllCore=False:sReport=""
    On Error GoTo EH
    If Not Orders_Preflight(oDoc,sErr) Then sReport=sErr:Exit Function
    oSh=oDoc.Sheets.getByName(WMSORD_SHEET)
    Orders_EnsureReceiptTypeColumn oSh

    ' First purge technical remnants from rows that contain only a prefilled number
    ' or WMS-generated fields. This prevents old ghost SourceID/state from being
    ' counted by "Проверить все".
    nClean=Orders_EmptyRowGuardCleanupCore(oDoc,oSh)

    Orders_RebuildGroupContexts oSh
    gWMSORD_Busy=True
    On Error Resume Next:oDoc.lockControllers:If Err=0 Then bLocked=True Else Err=0:On Error GoTo EH
    last=Orders_LastContentRow(oSh,WMSORD_USER_LAST_COL)
    cFp=Orders_FindHeader(oSh,WMSORD_X_FINGERPRINT)
    Orders_ResetDupCache
    Call Orders_EnsureDupCache(oSh)
    Orders_ResetHistoryCache
    Call Orders_EnsureHistoryCache(oSh)
    For r=1 To last
        If Orders_RowHasBusinessData(oSh,r) Or Orders_CellText(oSh,r,WMSORD_T_ID)<>"" Then
            nData=nData+1
            Orders_ProcessRow oDoc,oSh,r,bMarkDirty,bMarkDirty,True,False
            fp=Orders_BusinessFingerprint(oSh,r):If cFp>=0 Then oSh.getCellByPosition(cFp,r).String=fp
            sev=UCase(Orders_CellText(oSh,r,WMSORD_X_SEVERITY))
            If sev="CRITICAL" Then
                nCrit=nCrit+1
            ElseIf sev="WARN" Then
                nWarn=nWarn+1
            ElseIf sev="INFO" Then
                nInfo=nInfo+1
            End If
        End If
    Next r
    Orders_ResetHistoryCache
    Call Orders_EnsureHistoryCache(oSh)
    Orders_ApplyDesign oDoc,oSh
    Orders_ApplyValidations oDoc,oSh
    Orders_RefreshFilter oDoc,oSh
    sReport="Заказы проверены: строк " & CStr(nData) & ", critical " & CStr(nCrit) & ", warning " & CStr(nWarn) & ", info " & CStr(nInfo) & "."
    If nClean>0 Then sReport=sReport & Chr(10) & "Очищено пустых шаблонных строк: " & CStr(nClean) & "."
    Orders_RevalidateAllCore=True
Done:
    If bLocked Then On Error Resume Next:oDoc.unlockControllers:On Error GoTo 0
    gWMSORD_Busy=False
    Exit Function
EH:
    sReport="Revalidate Заказы: " & CStr(Err) & " " & Error$
    Orders_LogError oDoc,"RevalidateAll",sReport
    Resume Done
End Function

' ============================================================================
' INTEGRATION API FOR FUTURE BATCH / MOVEMENT ENGINE
' ============================================================================

Function Orders_API_FindRowBySourceID(sSourceID As String) As Long
    Dim oSh As Object,c As Long,last As Long,r As Long
    Orders_API_FindRowBySourceID=-1
    If Not ThisComponent.Sheets.hasByName(WMSORD_SHEET) Then Exit Function
    oSh=ThisComponent.Sheets.getByName(WMSORD_SHEET):c=Orders_FindHeader(oSh,WMSORD_T_ID):If c<0 Then Exit Function
    last=Orders_LastContentRow(oSh,WMSORD_USER_LAST_COL)
    For r=1 To last
        If Trim(oSh.getCellByPosition(c,r).String)=Trim(sSourceID) Then Orders_API_FindRowBySourceID=r:Exit Function
    Next r
End Function

Function Orders_API_ValidateForPosting(nRow As Long,ByRef sMessage As String) As Boolean
    Dim oSh As Object,fact As Double,ok As Boolean,status As String,docNo As String,mode As String,sev As String,flags As String,control As String
    Dim dupKey As String,nDupPeers As Long
    Orders_API_ValidateForPosting=False:sMessage=""
    If Not ThisComponent.Sheets.hasByName(WMSORD_SHEET) Then sMessage="Нет листа Заказы.":Exit Function
    oSh=ThisComponent.Sheets.getByName(WMSORD_SHEET)
    If nRow<1 Then sMessage="Выбрана строка заголовков.":Exit Function
    mode=UCase(Orders_CellText(oSh,nRow,WMSORD_T_MODE))
    If mode<>"" And mode<>"CURRENT" Then sMessage="Режим " & mode & " не проводится как новый текущий приход.":Exit Function
    ok=Orders_TryNumber(oSh.getCellByPosition(6,nRow),fact):If Not ok Or fact<=0 Then sMessage="Для оприходования нужен положительный факт прихода.":Exit Function
    status=Trim(oSh.getCellByPosition(16,nRow).String):If Orders_Canon(status)<>Orders_Canon(WMSORD_STATUS_POST) Then sMessage="Приход создаётся только после осознанного ручного статуса 'Оприходовано'.":Exit Function
    docNo=Orders_EffectiveDocNo(oSh,nRow):If docNo="" Then sMessage="Оприходование без № документа запрещено.":Exit Function
    dupKey=Orders_DuplicateKey(oSh,nRow)
    nDupPeers=Orders_DupCachePeerCount(oSh,dupKey)
    flags="":Orders_ValidateRow ThisComponent,oSh,nRow,sev,flags,control,nDupPeers
    If sev="CRITICAL" Then sMessage=control:Exit Function
    Orders_API_ValidateForPosting=True:sMessage=flags
End Function

Function Orders_API_GetPayload(nRow As Long) As Variant
    Dim oSh As Object,a(15) As Variant
    oSh=ThisComponent.Sheets.getByName(WMSORD_SHEET)
    a(0)=Orders_CellText(oSh,nRow,WMSORD_T_ID)
    a(1)=Orders_CellDate(oSh.getCellByPosition(12,nRow),True)
    a(2)=Trim(oSh.getCellByPosition(1,nRow).String)
    a(3)=Orders_EffectiveDocNo(oSh,nRow)
    a(4)=Orders_EffectiveInvoice(oSh,nRow)
    a(5)=Trim(oSh.getCellByPosition(5,nRow).String)
    a(6)=oSh.getCellByPosition(6,nRow).Value
    a(7)=Trim(oSh.getCellByPosition(8,nRow).String)
    a(8)=Trim(oSh.getCellByPosition(11,nRow).String)
    a(9)=Trim(oSh.getCellByPosition(22,nRow).String)
    a(10)=Trim(oSh.getCellByPosition(19,nRow).String)
    a(11)=Trim(oSh.getCellByPosition(20,nRow).String)
    a(12)=Trim(oSh.getCellByPosition(21,nRow).String)
    a(13)=Trim(oSh.getCellByPosition(17,nRow).String)
    a(14)=Orders_CellText(oSh,nRow,WMSORD_T_MODE)
    a(15)=Orders_BusinessFingerprint(oSh,nRow)
    Orders_API_GetPayload=a()
End Function

Sub Orders_API_MarkConducted(sSourceID As String,sSyncHash As String)
    Dim r As Long,oSh As Object
    r=Orders_API_FindRowBySourceID(sSourceID):If r<1 Then Exit Sub
    oSh=ThisComponent.Sheets.getByName(WMSORD_SHEET)
    gWMSORD_Busy=True
    Orders_SetTech oSh,r,WMSORD_T_STATE,"Проведено"
    Orders_SetTech oSh,r,WMSORD_T_HASH,sSyncHash
    Orders_SetTech oSh,r,WMSORD_T_DIRTY,""
    Orders_SetTech oSh,r,WMSORD_T_ERROR,""
    Orders_ApplyRowVisual oSh,r
    gWMSORD_Busy=False
End Sub

Sub Orders_API_MarkCancelled(sSourceID As String,sReason As String)
    Dim r As Long,oSh As Object
    r=Orders_API_FindRowBySourceID(sSourceID):If r<1 Then Exit Sub
    oSh=ThisComponent.Sheets.getByName(WMSORD_SHEET)
    gWMSORD_Busy=True
    Orders_SetTech oSh,r,WMSORD_T_STATE,"Отменено"
    Orders_SetTech oSh,r,WMSORD_T_DIRTY,""
    If Trim(sReason)<>"" Then Orders_SetTech oSh,r,WMSORD_T_ERROR,sReason
    Orders_ApplyRowVisual oSh,r
    gWMSORD_Busy=False
End Sub

' ============================================================================
' SELF CHECK
' ============================================================================

Function Orders_RunSelfCheck(oDoc As Object,ByRef sReport As String) As Boolean
    Dim oSh As Object,a As Variant,i As Long,c As Long,bad As Long,s As String,ev As Variant,v As Object,db As Object,addr As Variant
    Dim tb As Variant, gridCell As Object, qtyKey As Long, moneyKey As Long, nfmt As Object
    Dim qtyPreview As String, moneyPreview As String
    sReport="":Orders_RunSelfCheck=False
    On Error GoTo EH
    If Not Orders_Preflight(oDoc,s) Then sReport=s:Exit Function
    oSh=oDoc.Sheets.getByName(WMSORD_SHEET)
    a=Orders_ExtensionHeaders()
    For i=0 To UBound(a)
        c=Orders_FindHeader(oSh,CStr(a(i)))
        If c<0 Then
            sReport=sReport & "Нет техполя " & CStr(a(i)) & "." & Chr(10)
            bad=bad+1
        ElseIf oSh.Columns.getByIndex(c).IsVisible Then
            sReport=sReport & "Техполе видно: " & CStr(a(i)) & "." & Chr(10)
            bad=bad+1
        End If
    Next i
    If Not oSh.Columns.getByIndex(20).IsVisible Then sReport=sReport & "U Комментарий ошибочно скрыт." & Chr(10):bad=bad+1
    If Not oSh.Columns.getByIndex(21).IsVisible Then sReport=sReport & "V Код товара ошибочно скрыт." & Chr(10):bad=bad+1
    If Not oSh.Columns.getByIndex(22).IsVisible Then sReport=sReport & "W Продавец ошибочно скрыт." & Chr(10):bad=bad+1
    If Not oSh.Columns.getByIndex(23).IsVisible Then sReport=sReport & "X Источник прихода ошибочно скрыт." & Chr(10):bad=bad+1
    If Not oSh.Columns.getByIndex(24).IsVisible Then sReport=sReport & "Y Подкатегория ошибочно скрыта." & Chr(10):bad=bad+1
    If oSh.Columns.getByIndex(25).IsVisible Then sReport=sReport & "Z legacy-поле должно быть скрыто." & Chr(10):bad=bad+1

    ev=oSh.Events.getByName("OnChange")
    If IsNull(ev) Or IsEmpty(ev) Then
        sReport=sReport & "Не назначено событие OnChange." & Chr(10):bad=bad+1
    Else
        If InStr(Orders_EventScript(ev),"WMS_02_Orders_FINAL.Orders_OnContentChanged")=0 Then sReport=sReport & "OnChange назначен не на Orders_OnContentChanged." & Chr(10):bad=bad+1
    End If

    v=oSh.getCellByPosition(16,1).Validation
    If v.Type<>com.sun.star.sheet.ValidationType.LIST Then sReport=sReport & "В Q Статус нет LIST validation." & Chr(10):bad=bad+1
    If InStr(v.Formula1,"Оприходовано")=0 Then sReport=sReport & "В списке статусов нет Оприходовано." & Chr(10):bad=bad+1

    If Not oDoc.DatabaseRanges.hasByName("WMS00_DB_02") Then
        sReport=sReport & "Нет AutoFilter database range WMS00_DB_02." & Chr(10):bad=bad+1
    Else
        db=oDoc.DatabaseRanges.getByName("WMS00_DB_02"):addr=db.getDataArea()
        If addr.StartColumn<>0 Or addr.EndColumn<Orders_LastHeaderCol(oSh) Then sReport=sReport & "Фильтр Заказы не охватывает всю строку вместе со скрытыми техполями." & Chr(10):bad=bad+1
    End If

    If Orders_GetSetting(oDoc,"WMS_OrdersVersion","")<>WMSORD_VERSION Then sReport=sReport & "Настройка WMS_OrdersVersion не соответствует модулю." & Chr(10):bad=bad+1

    ' Visual grid is applied as per-cell BottomBorder2/RightBorder2 because Calc
    ' does not reliably report TableBorder2 inner-line validity on large ranges.
    gridCell=oSh.getCellByPosition(5,2)
    If gridCell.BottomBorder2.LineWidth<=0 Then sReport=sReport & "Нет горизонтальных разделителей ячеек." & Chr(10):bad=bad+1
    If gridCell.RightBorder2.LineWidth<=0 Then sReport=sReport & "Нет вертикальных разделителей ячеек." & Chr(10):bad=bad+1

    ' Number formats are checked through the LibreOffice formatter itself because
    ' user paste may temporarily alter an individual cell format before OnChange
    ' repairs it. The generated locale-safe keys must render the canonical values.
    qtyKey=Orders_NumberFormatKey(oDoc,"QTY")
    moneyKey=Orders_NumberFormatKey(oDoc,"MONEY")
    nfmt=CreateUnoService("com.sun.star.util.NumberFormatter")
    nfmt.attachNumberFormatsSupplier(oDoc)
    qtyPreview=nfmt.convertNumberToString(qtyKey,1)
    moneyPreview=nfmt.convertNumberToString(moneyKey,456)
    If qtyPreview<>"1" Then sReport=sReport & "Формат количества отображает 1 как '" & qtyPreview & "'." & Chr(10):bad=bad+1
    If Orders_GroupColorIndex("GROUP-A-1") = Orders_GroupColorIndex("GROUP-B-2") Then sReport=sReport & "Палитра заказов не различает тестовые группы." & Chr(10):bad=bad+1
    If moneyPreview<>"456,00" Then sReport=sReport & "Формат цены отображает 456 как '" & moneyPreview & "'." & Chr(10):bad=bad+1

    If bad=0 Then
        sReport="WMS_02_Orders " & WMSORD_VERSION & " self-check: OK. Схема, сетка, форматы, скрытые поля, U/V/W, событие, статус-validation, фильтр, нумерация групп и группировка и правила реквизитов заказа корректны."
        Orders_RunSelfCheck=True
    Else
        sReport="WMS_02_Orders self-check: найдено проблем " & CStr(bad) & "." & Chr(10) & sReport
    End If
    Exit Function
EH:
    sReport="Orders self-check runtime error " & CStr(Err) & ": " & Error$
End Function

Function Orders_EventScript(ev As Variant) As String
    Dim i As Long
    On Error GoTo Done
    For i=LBound(ev) To UBound(ev)
        If ev(i).Name="Script" Then Orders_EventScript=CStr(ev(i).Value):Exit Function
    Next i
Done:
End Function

' ============================================================================
' FINGERPRINT / SORT SAFETY
' ============================================================================

Function Orders_BusinessFingerprint(oSh As Object,r As Long) As String
    Dim c As Long,s As String,p As String,h1 As Double,h2 As Double,i As Long,ch As Long
    ' Control S (18) is automatic and intentionally excluded.
    For c=0 To WMSORD_USER_LAST_COL
        If c<>18 Then
            p=oSh.getCellByPosition(c,r).Formula
            s=s & "|" & CStr(c) & "=" & p
        End If
    Next c
    h1=5381:h2=7919
    For i=1 To Len(s)
        ch=Asc(Mid(s,i,1))
        If ch<0 Then ch=ch+256
        ' Keep arithmetic safely inside LibreOffice Basic 32-bit Mod range.
        h1=(h1*33+ch) Mod 1000003
        h2=(h2*131+ch+i) Mod 1000033
    Next i
    Orders_BusinessFingerprint=Hex(CLng(h1)) & "-" & Hex(CLng(h2)) & "-" & CStr(Len(s))
End Function

Sub Orders_ResetHeaderCache()
    On Error Resume Next
    If gWMSORD_HeaderCacheReady Then gWMSORD_HeaderCache.Dispose
    gWMSORD_HeaderCache=Empty
    gWMSORD_HeaderCacheReady=False
    gWMSORD_LastHeaderColCached=0
    On Error GoTo 0
End Sub

Function Orders_EnsureHeaderCache(oSh As Object) As Boolean
    Dim cur As Object,a As Variant,c As Long,h As String
    Orders_EnsureHeaderCache=False
    If oSh.Name<>WMSORD_SHEET Then Exit Function
    On Error GoTo EH
    If gWMSORD_HeaderCacheReady Then Orders_EnsureHeaderCache=True:Exit Function
    GlobalScope.BasicLibraries.LoadLibrary("ScriptForge")
    gWMSORD_HeaderCache=CreateScriptService("Dictionary",False)
    cur=oSh.createCursor():cur.gotoEndOfUsedArea(True):a=cur.RangeAddress
    gWMSORD_LastHeaderColCached=0
    For c=0 To a.EndColumn
        h=Orders_Canon(oSh.getCellByPosition(c,0).String)
        If h<>"" Then
            If Not gWMSORD_HeaderCache.Exists(h) Then gWMSORD_HeaderCache.Add h,c
            gWMSORD_LastHeaderColCached=c
        End If
    Next c
    gWMSORD_HeaderCacheReady=True
    Orders_EnsureHeaderCache=True
    Exit Function
EH:
    gWMSORD_HeaderCacheReady=False
End Function

Sub Orders_EmptyRowGuardCleanup()
    Dim oDoc As Object,oSh As Object,n As Long
    On Error GoTo EH
    oDoc=ThisComponent
    If Not oDoc.Sheets.hasByName(WMSORD_SHEET) Then MsgBox "Лист 'Заказы' не найден.",48,"WMS — Заказы":Exit Sub
    oSh=oDoc.Sheets.getByName(WMSORD_SHEET)
    gWMSORD_Busy=True
    On Error Resume Next
    oDoc.lockControllers
    On Error GoTo EH
    n=Orders_EmptyRowGuardCleanupCore(oDoc,oSh)
    Orders_ResetQueueCache
    On Error Resume Next
    oDoc.unlockControllers
    On Error GoTo 0
    gWMSORD_Busy=False
    MsgBox "Empty Row Guard завершён." & Chr(10) & _
           "Очищено ложных пустых строк: " & CStr(n) & Chr(10) & _
           "Реальные, проведённые и уже синхронизированные строки не изменялись.",64,"WMS — Заказы"
    Exit Sub
EH:
    On Error Resume Next
    oDoc.unlockControllers
    On Error GoTo 0
    gWMSORD_Busy=False
    MsgBox "Empty Row Guard: ошибка " & CStr(Err) & ": " & Error$,16,"WMS — Заказы"
End Sub

Function Orders_EmptyRowGuardCleanupCore(oDoc As Object,oSh As Object) As Long
    Dim lastRow As Long,r As Long,sID As String,sState As String,sHash As String
    Dim aBase As Variant,aExt As Variant,i As Long,c As Long,n As Long
    Orders_EmptyRowGuardCleanupCore=0
    On Error GoTo Done

    lastRow=Orders_LastContentRow(oSh,Orders_LastHeaderCol(oSh))
    If lastRow<1 Then Exit Function
    aBase=Orders_BaseHeaders()
    aExt=Orders_ExtensionHeaders()

    For r=1 To lastRow
        If Not Orders_RowHasBusinessData(oSh,r) Then
            sID=Orders_CellText(oSh,r,WMSORD_T_ID)
            sState=Orders_Canon(Orders_CellText(oSh,r,WMSORD_T_STATE))
            sHash=Trim(Orders_CellText(oSh,r,WMSORD_T_HASH))

            ' Conservative migration rule: never purge anything that has already
            ' been conducted or has a synchronization hash/history marker.
            If sState<>Orders_Canon("Проведено") And sHash="" Then
                ' Clear only WMS-generated visible fields. User columns stay intact.
                oSh.getCellByPosition(16,r).String=""
                oSh.getCellByPosition(18,r).String=""

                ' Clear all technical metadata columns belonging to Orders.
                For i=23 To UBound(aBase)
                    c=Orders_FindHeader(oSh,CStr(aBase(i)))
                    If c>=0 Then oSh.getCellByPosition(c,r).String=""
                Next i
                For i=0 To UBound(aExt)
                    c=Orders_FindHeader(oSh,CStr(aExt(i)))
                    If c>=0 Then
                        oSh.getCellByPosition(c,r).String=""
                        If CStr(aExt(i))=WMSORD_X_EFFDOCDATE Then oSh.getCellByPosition(c,r).Value=0
                    End If
                Next i

                If sID<>"" Then Orders_QueueCancelSourceID oDoc,sID,"EMPTY_ROW_GUARD"
                Orders_ApplyRowVisual oSh,r
                n=n+1
            End If
        End If
    Next r
    Orders_EmptyRowGuardCleanupCore=n
Done:
End Function

Sub Orders_QueueCancelSourceID(oDoc As Object,sSourceID As String,sReason As String)
    Dim oQ As Object,last As Long,r As Long
    If Trim(sSourceID)="" Then Exit Sub
    If Not oDoc.Sheets.hasByName(WMSORD_QUEUE) Then Exit Sub
    On Error GoTo Done
    oQ=oDoc.Sheets.getByName(WMSORD_QUEUE)
    last=Orders_QueueLastRowFast(oQ)
    For r=1 To last
        If Trim(oQ.getCellByPosition(1,r).String)=WMSORD_SHEET And _
           Trim(oQ.getCellByPosition(2,r).String)=Trim(sSourceID) Then
            If UCase(Trim(oQ.getCellByPosition(5,r).String))="PENDING" Or _
               Trim(oQ.getCellByPosition(5,r).String)="" Then
                oQ.getCellByPosition(5,r).String="CANCELLED"
                oQ.getCellByPosition(7,r).String=sReason
            End If
        End If
    Next r
Done:
End Sub


' ============================================================================
' BUTTON PANEL 1.0.12
' ============================================================================

Sub Orders_ButtonFillAllByCodes(Optional oEvent)
    Dim oDoc As Object,oSh As Object,oCon As Object,sErr As String
    Dim last As Long,r As Long,nCodes As Long,nFilled As Long,nMiss As Long
    Dim code As String

    oDoc=ThisComponent
    If Not oDoc.Sheets.hasByName(WMSORD_SHEET) Then Exit Sub
    oSh=oDoc.Sheets.getByName(WMSORD_SHEET)
    oDoc.CurrentController.setActiveSheet(oSh)

    sErr=""
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then MsgBox sErr,16,"WMS — Заказы":Exit Sub

    last=Orders_LastContentRow(oSh,21)
    If last<1 Then
        MsgBox "Нет строк для массового заполнения.",48,"WMS — Заказы"
        Exit Sub
    End If

    On Error GoTo EH
    For r=1 To last
        code=Trim(oSh.getCellByPosition(21,r).String)
        If code<>"" Then
            nCodes=nCodes+1
            sErr=""
            If Orders_FillProductNameFromDB(oSh,r,sErr) Then
                nFilled=nFilled+1
            Else
                nMiss=nMiss+1
            End If
        End If
    Next r

    MsgBox "Массовое заполнение Заказов завершено." & Chr(10) & _
           "Кодов обработано: " & CStr(nCodes) & Chr(10) & _
           "Найдено в БД: " & CStr(nFilled) & Chr(10) & _
           "Не найдено: " & CStr(nMiss), _
           IIf(nMiss=0,64,48),"WMS — Заказы"
    Exit Sub
EH:
    MsgBox "Массовое заполнение остановлено: " & CStr(Err) & " " & Error$,16,"WMS — Заказы"
End Sub

Sub Orders_ButtonLotUnits(Optional oEvent)
    On Error GoTo EH
    WMSDBSR_ConfigureSelected
    Exit Sub
EH:
    MsgBox "Единицы партии недоступны. Проверьте WMS_12 SmartReceipt: " & _
           CStr(Err) & " " & Error$,16,"WMS — Заказы"
End Sub

Sub Orders_InstallButtons()
    Dim oDoc As Object,oSh As Object,sErr As String
    oDoc=ThisComponent
    If Not oDoc.Sheets.hasByName(WMSORD_SHEET) Then
        MsgBox "Лист 'Заказы' не найден.",16,"WMS — Заказы"
        Exit Sub
    End If
    oSh=oDoc.Sheets.getByName(WMSORD_SHEET)
    If Orders_InstallButtonsCore(oDoc,oSh,sErr) Then
        MsgBox "Панель кнопок листа 'Заказы' создана/обновлена.",64,"WMS — Заказы"
    Else
        MsgBox sErr,16,"WMS — Заказы"
    End If
End Sub


Sub Orders_ButtonCheckSelected()
    Dim oDoc As Object,oSh As Object,oSel As Object,r As Long
    oDoc=ThisComponent
    If Not oDoc.Sheets.hasByName(WMSORD_SHEET) Then Exit Sub
    oSh=oDoc.CurrentController.ActiveSheet
    If oSh.Name<>WMSORD_SHEET Then
        MsgBox "Откройте лист 'Заказы'.",48,"WMS — Заказы"
        Exit Sub
    End If
    On Error GoTo Bad
    oSel=oDoc.CurrentSelection
    r=oSel.CellAddress.Row
    If r<1 Then GoTo Bad
    gWMSORD_Busy=True
    Orders_ProcessRow oDoc,oSh,r,False,False,True,True
    gWMSORD_Busy=False
    Orders_ShowSelectedRowDiagnostics
    Exit Sub
Bad:
    gWMSORD_Busy=False
    MsgBox "Выделите одну ячейку рабочей строки заказа.",48,"WMS — Заказы"
End Sub

Sub Orders_ButtonValidateAll()
    Orders_RevalidateAll
End Sub

Sub Orders_ButtonRebuildGroups()
    Orders_RebuildOrderGroups
End Sub

Sub Orders_ButtonAutoStatus()
    Orders_ResetSelectedStatusToAuto
End Sub

Sub Orders_ButtonAutoCode()
    Orders_ResetSelectedCodeToAuto
End Sub

Sub Orders_ButtonAutoCategory()
    Orders_ResetSelectedCategoryToAuto
End Sub

Sub Orders_ButtonSyncPositionDB()
    On Error GoTo Missing
    WMSDBO_SyncSelectedPosition
    Exit Sub
Missing:
    MsgBox "Модуль WMS_06_DB_Orders не установлен или недоступен.",48,"WMS — Заказы / БД"
End Sub

Sub Orders_ButtonSyncOrderDB()
    On Error GoTo Missing
    WMSDBO_SyncCurrentOrder
    Exit Sub
Missing:
    MsgBox "Модуль WMS_06_DB_Orders не установлен или недоступен.",48,"WMS — Заказы / БД"
End Sub

Sub Orders_ButtonSyncAllDB()
    On Error GoTo Missing
    WMSDBO_SyncAllReadyPositions
    Exit Sub
Missing:
    MsgBox "Модуль WMS_06_DB_Orders не установлен или недоступен.",48,"WMS — Заказы / БД"
End Sub

Sub Orders_ButtonViewDB()
    On Error GoTo Missing
    WMSDBO_ViewSelectedPositionDB
    Exit Sub
Missing:
    MsgBox "Модуль WMS_06_DB_Orders не установлен или недоступен.",48,"WMS — Заказы / БД"
End Sub

Sub Orders_ButtonDBStatus()
    On Error GoTo Missing
    WMSDBO_SelectedPositionStatus
    Exit Sub
Missing:
    MsgBox "Модуль WMS_06_DB_Orders не установлен или недоступен.",48,"WMS — Заказы / БД"
End Sub

Sub Orders_ButtonRefreshDesign()
    Orders_RefreshDesign
End Sub

Sub Orders_ButtonSelfCheck()
    Orders_SelfCheck
End Sub

Function Orders_InstallButtonsCore(oDoc As Object,oSh As Object,ByRef sErr As String) As Boolean
    Dim oDP As Object,oForms As Object,oForm As Object
    Dim baseX As Long,baseY As Long,w As Long,h As Long,gap As Long,colGap As Long
    Orders_InstallButtonsCore=False:sErr=""
    On Error GoTo EH

    oDP=oSh.DrawPage
    oForms=oDP.Forms
    Orders_RemoveButtonPanel oSh

    oForm=oDoc.createInstance("com.sun.star.form.component.Form")
    oForm.Name="WMS_ORDERS_BUTTON_PANEL"
    oForms.insertByName("WMS_ORDERS_BUTTON_PANEL",oForm)

    ' The panel starts immediately after the last user column W.
    ' Two columns keep the panel compact and avoid covering order data.
    baseX=oSh.getCellByPosition(WMSORD_USER_LAST_COL,0).Position.X + _
          oSh.getCellByPosition(WMSORD_USER_LAST_COL,0).Size.Width + 350
    baseY=oSh.getCellByPosition(0,0).Position.Y + 120
    w=4100:h=650:gap=120:colGap=220

    ' Left column — work with Calc orders.
    Orders_AddButton oDoc,oSh,oForm,"WMS_ORD_BTN_NEW","Новый заказ",baseX,baseY,w,h,"Orders_ButtonNewOrder"
    Orders_AddButton oDoc,oSh,oForm,"WMS_ORD_BTN_DIRECT","Приход вне заказа",baseX,baseY+(h+gap),w,h,"Orders_ButtonDirectReceipt"
    Orders_AddButton oDoc,oSh,oForm,"WMS_ORD_BTN_UNIVERSAL","Новый приход",baseX,baseY+2*(h+gap),w,h,"Orders_ButtonUniversalReceipt"
    Orders_AddButton oDoc,oSh,oForm,"WMS_ORD_BTN_CHECK","Проверить строку",baseX,baseY+3*(h+gap),w,h,"Orders_ButtonCheckSelected"
    Orders_AddButton oDoc,oSh,oForm,"WMS_ORD_BTN_ALL","Проверить все",baseX,baseY+4*(h+gap),w,h,"Orders_ButtonValidateAll"
    Orders_AddButton oDoc,oSh,oForm,"WMS_ORD_BTN_GROUPS","Пересобрать группы",baseX,baseY+5*(h+gap),w,h,"Orders_ButtonRebuildGroups"
    Orders_AddButton oDoc,oSh,oForm,"WMS_ORD_BTN_STATUS","Статус: авто",baseX,baseY+6*(h+gap),w,h,"Orders_ButtonAutoStatus"
    Orders_AddButton oDoc,oSh,oForm,"WMS_ORD_BTN_CODE","Код: авто",baseX,baseY+7*(h+gap),w,h,"Orders_ButtonAutoCode"
    Orders_AddButton oDoc,oSh,oForm,"WMS_ORD_BTN_CAT","Категория: авто",baseX,baseY+8*(h+gap),w,h,"Orders_ButtonAutoCategory"

    ' Right column — normal user actions. Test/service sync procedures remain
    ' available as macros, but are not placed on the everyday panel.
    baseX=baseX+w+colGap
    Orders_AddButton oDoc,oSh,oForm,"WMS_ORD_BTN_CONDUCT_POS","Провести позицию",baseX,baseY,w,h,"Orders_ButtonConductPositionDB"
    Orders_AddButton oDoc,oSh,oForm,"WMS_ORD_BTN_CONDUCT","Провести заказ",baseX,baseY+(h+gap),w,h,"Orders_ButtonConductToDB"
    Orders_AddButton oDoc,oSh,oForm,"WMS_ORD_BTN_CONDUCTALL","Провести все",baseX,baseY+2*(h+gap),w,h,"Orders_ButtonConductAllDB"
    Orders_AddButton oDoc,oSh,oForm,"WMS_ORD_BTN_DBVIEW","БД: показать",baseX,baseY+3*(h+gap),w,h,"Orders_ButtonViewDB"
    Orders_AddButton oDoc,oSh,oForm,"WMS_ORD_BTN_DBSTATE","БД: статус",baseX,baseY+4*(h+gap),w,h,"Orders_ButtonDBStatus"
    Orders_AddButton oDoc,oSh,oForm,"WMS_ORD_BTN_DESIGN","Обновить оформление",baseX,baseY+5*(h+gap),w,h,"Orders_ButtonRefreshDesign"
    Orders_AddButton oDoc,oSh,oForm,"WMS_ORD_BTN_SELF","SelfCheck",baseX,baseY+6*(h+gap),w,h,"Orders_ButtonSelfCheck"
    Orders_AddButton oDoc,oSh,oForm,"WMS_ORD_BTN_LOTUNITS","Единицы партии",baseX,baseY+7*(h+gap),w,h,"Orders_ButtonLotUnits"
    Orders_AddButton oDoc,oSh,oForm,"WMS_ORD_BTN_BULKFILL","Заполнить все по кодам",baseX,baseY+8*(h+gap),w,h,"Orders_ButtonFillAllByCodes"
    Orders_AddButton oDoc,oSh,oForm,"WMS_ORD_BTN_ARTICLE","Распознать по артикулам",baseX,baseY+9*(h+gap),w,h,"Orders_ButtonResolveArticles"
    Orders_AddButton oDoc,oSh,oForm,"WMS_ORD_BTN_UPD","Допоставка к заказу",baseX,baseY+10*(h+gap),w,h,"Orders_ButtonAddDelivery"
    Orders_AddButton oDoc,oSh,oForm,"WMS_ORD_BTN_DELDRAFT","Удалить черновик",baseX,baseY+11*(h+gap),w,h,"Orders_ButtonDeleteReceiptDraft"
    Orders_AddButton oDoc,oSh,oForm,"WMS_ORD_BTN_RECOVER","Незавершённые операции",baseX,baseY+12*(h+gap),w,h,"Orders_ButtonSafetyIncomplete"
    Orders_AddButton oDoc,oSh,oForm,"WMS_ORD_BTN_AUDIT","Журнал действий",baseX,baseY+13*(h+gap),w,h,"Orders_ButtonSafetyAudit"

    Orders_InstallButtonsCore=True
    Exit Function
EH:
    sErr="Ошибка панели кнопок: " & CStr(Err) & " " & Error$
End Function

Sub Orders_RemoveButtonPanel(oSh As Object)
    Dim oDP As Object,oForms As Object,i As Long,oShape As Object,oCtl As Object,nm As String
    On Error Resume Next
    oDP=oSh.DrawPage
    For i=oDP.Count-1 To 0 Step -1
        oShape=oDP.getByIndex(i)
        nm=""
        oCtl=oShape.Control
        nm=oCtl.Name
        If Left(nm,12)="WMS_ORD_BTN_" Then oDP.remove(oShape)
    Next i
    oForms=oDP.Forms
    If oForms.hasByName("WMS_ORDERS_BUTTON_PANEL") Then oForms.removeByName("WMS_ORDERS_BUTTON_PANEL")
    On Error GoTo 0
End Sub

Sub Orders_AddButton(oDoc As Object,oSh As Object,oForm As Object,ctlName As String,labelText As String,x As Long,y As Long,w As Long,h As Long,macroName As String)
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
    ev.ScriptCode="vnd.sun.star.script:Standard.WMS_02_Orders_FINAL." & macroName & "?language=Basic&location=document"
    oForm.registerScriptEvent(idx,ev)
End Sub



' ============================================================================
' DIRECT RECEIPT + SAFE CONDUCT 1.0.14
' ============================================================================

Sub Orders_ButtonDirectReceipt()
    Orders_CreateNewEntryWithSource "DIRECT_RECEIPT","Поставщик / прямой приход"
End Sub

Sub Orders_ButtonNewOrder()
    Orders_CreateNewEntryWithSource "ORDER","Поставщик / заказ"
End Sub

Sub Orders_ButtonUniversalReceipt()
    On Error GoTo Legacy
    WMSRC_NewReceiptMenu
    Exit Sub
Legacy:
    MsgBox "Модуль WMS_14 UniversalReceiptV2 не установлен." & Chr(10) & _
           "Установите WMS_14 и запустите WMSRC_Install.",48,"WMS — Новый приход"
End Sub

Sub Orders_CreateNewEntry(entryType As String)
    Dim src As String
    If UCase(Trim(entryType))="ORDER" Then
        src="Поставщик / заказ"
    Else
        src="Поставщик / прямой приход"
    End If
    Orders_CreateNewEntryWithSource entryType,src
End Sub

Sub Orders_CreateNewEntryWithSource(entryType As String,src As String)
    Dim oDoc As Object,oSh As Object,r As Long,maxRow As Long
    oDoc=ThisComponent
    If Not oDoc.Sheets.hasByName(WMSORD_SHEET) Then Exit Sub
    oSh=oDoc.Sheets.getByName(WMSORD_SHEET)
    oDoc.CurrentController.setActiveSheet(oSh)

    maxRow=Orders_LastContentRow(oSh,WMSORD_USER_LAST_COL)+1
    If maxRow<1 Then maxRow=1
    r=1
    Do While r<=maxRow+500
        If Not Orders_RowHasBusinessData(oSh,r) Then Exit Do
        r=r+1
    Loop

    gWMSORD_Busy=True
    oSh.getCellByPosition(0,r).Value=1
    oSh.getCellByPosition(23,r).String=src
    Orders_SetReceiptType oSh,r,entryType
    gWMSORD_Busy=False

    Orders_ProcessRow oDoc,oSh,r,True,True,True,True
    oDoc.CurrentController.select(oSh.getCellByPosition(1,r))
End Sub

Sub Orders_ButtonDeleteReceiptDraft()
    WMSRC_DeleteDraftReceipt
End Sub

Sub Orders_ButtonAddDelivery()
    WMSRC_AttachUPDToSelectedReceipt
End Sub

Sub Orders_ButtonAttachUPD()
    WMSRC_AttachUPDToSelectedReceipt
End Sub

Sub Orders_ButtonSafetyIncomplete()
    WMSSAFE_CheckIncomplete
End Sub

Sub Orders_ButtonSafetyAudit()
    WMSSAFE_ShowAudit
End Sub

Sub Orders_ButtonResolveArticles()
    On Error GoTo EH
    WMSRC_ResolveArticles
    Exit Sub
EH:
    MsgBox "Поиск по артикулам недоступен. Проверьте WMS_14 UniversalReceiptV2.",48,"WMS — Заказы"
End Sub

Sub Orders_ButtonConductAllDB()
    On Error GoTo EH
    WMSDBO_ConductAllReady
    Exit Sub
EH:
    MsgBox "Массовое проведение недоступно. Проверьте WMS_06 DB Orders 0.6.0." & Chr(10) & _
           CStr(Err) & " " & Error$,16,"WMS — Заказы"
End Sub

Sub Orders_ButtonConductPositionDB()
    On Error GoTo Missing
    WMSDBO_ConductSelectedPosition
    Exit Sub
Missing:
    MsgBox "Модуль WMS_06_DB_Orders с проведением позиции не установлен.",48,"WMS — Заказы / БД"
End Sub

Sub Orders_ButtonConductToDB()
    On Error GoTo Missing
    WMSDBO_ConductCurrentOrder
    Exit Sub
Missing:
    MsgBox "Модуль WMS_06_DB_Orders с процедурой проведения не установлен.",48,"WMS — Заказы / БД"
End Sub

Function Orders_GetReceiptType(oSh As Object,r As Long) As String
    Dim s As String,c As Long
    c=Orders_GetOrCreateTechColumn(oSh,"_WMS_ReceiptType")
    On Error Resume Next
    s=UCase(Trim(oSh.getCellByPosition(c,r).String))
    On Error GoTo 0
    If s="" Then s="ORDER"
    Orders_GetReceiptType=s
End Function

Sub Orders_SetReceiptType(oSh As Object,r As Long,entryType As String)
    Dim c As Long
    c=Orders_GetOrCreateTechColumn(oSh,"_WMS_ReceiptType")
    oSh.getCellByPosition(c,r).String=entryType
End Sub

Sub Orders_EnsureReceiptTypeColumn(oSh As Object)
    Dim c As Long
    c=Orders_GetOrCreateTechColumn(oSh,"_WMS_ReceiptType")
    On Error Resume Next
    oSh.Columns.getByIndex(c).IsVisible=False
    On Error GoTo 0
End Sub

Function Orders_GetOrCreateTechColumn(oSh As Object,headerName As String) As Long
    Dim cur As Object,lastCol As Long,c As Long,s As String
    cur=oSh.createCursor()
    cur.gotoEndOfUsedArea(True)
    lastCol=cur.RangeAddress.EndColumn
    If lastCol<WMSORD_USER_LAST_COL Then lastCol=WMSORD_USER_LAST_COL

    For c=WMSORD_USER_LAST_COL+1 To lastCol
        s=Trim(oSh.getCellByPosition(c,0).String)
        If s=headerName Then
            Orders_GetOrCreateTechColumn=c
            Exit Function
        End If
    Next c

    ' Append only after all existing technical columns. Never overwrite X/Y/...
    c=lastCol+1
    oSh.getCellByPosition(c,0).String=headerName
    Orders_GetOrCreateTechColumn=c
End Function

' ============================================================================
' GENERIC HELPERS
' ============================================================================

Function Orders_RowHasBusinessData(oSh As Object,r As Long) As Boolean
    Dim c As Long,cell As Object
    ' 1.0.13 Revalidate Guard:
    ' A (Номер) alone is only a row/order template marker and must not activate a row.
    ' Q (Статус) and S (Контроль) are generated by WMS itself.
    For c=0 To WMSORD_USER_LAST_COL
        If c<>0 And c<>16 And c<>18 Then
            cell=oSh.getCellByPosition(c,r)
            If Trim(cell.Formula)<>"" Then Orders_RowHasBusinessData=True:Exit Function
        End If
    Next c
End Function

Function Orders_CellText(oSh As Object,r As Long,sHeader As String) As String
    Dim c As Long:c=Orders_FindHeader(oSh,sHeader):If c>=0 Then Orders_CellText=Trim(oSh.getCellByPosition(c,r).String)
End Function

Sub Orders_SetCellText(oSh As Object,r As Long,sHeader As String,sValue As String)
    Dim c As Long:c=Orders_FindHeader(oSh,sHeader):If c>=0 Then oSh.getCellByPosition(c,r).String=sValue
End Sub

Sub Orders_SetTech(oSh As Object,r As Long,sHeader As String,sValue As String)
    Dim c As Long:c=Orders_FindHeader(oSh,sHeader):If c>=0 Then oSh.getCellByPosition(c,r).String=sValue
End Sub

Function Orders_FindHeader(oSh As Object,sHeader As String) As Long
    Dim c As Long,last As Long,k As String
    Orders_FindHeader=-1
    k=Orders_Canon(sHeader)
    If oSh.Name=WMSORD_SHEET Then
        If Orders_EnsureHeaderCache(oSh) Then
            On Error Resume Next
            If gWMSORD_HeaderCache.Exists(k) Then Orders_FindHeader=CLng(gWMSORD_HeaderCache.Item(k))
            On Error GoTo 0
            Exit Function
        End If
    End If
    last=Orders_LastHeaderCol(oSh)
    For c=0 To last
        If Orders_Canon(oSh.getCellByPosition(c,0).String)=k Then Orders_FindHeader=c:Exit Function
    Next c
End Function

Function Orders_LastHeaderCol(oSh As Object) As Long
    Dim cur As Object,a As Variant,c As Long
    Orders_LastHeaderCol=0
    If oSh.Name=WMSORD_SHEET Then
        If Orders_EnsureHeaderCache(oSh) Then Orders_LastHeaderCol=gWMSORD_LastHeaderColCached:Exit Function
    End If
    On Error GoTo Done
    cur=oSh.createCursor():cur.gotoEndOfUsedArea(True):a=cur.RangeAddress
    For c=a.EndColumn To 0 Step -1
        If Trim(oSh.getCellByPosition(c,0).String)<>"" Then Orders_LastHeaderCol=c:Exit Function
    Next c
Done:
End Function

Function Orders_LastContentRow(oSh As Object,nLastCol As Long) As Long
    ' OPT-1: existing shared batch reader replaces per-cell UNO calls.
    Orders_LastContentRow=0
    If nLastCol<0 Then Exit Function
    Orders_LastContentRow=WMSCore_LastContentRow(oSh,nLastCol)
End Function

Function Orders_NextRow(oSh As Object,nKeyCol As Long,nStart As Long) As Long
    Dim last As Long:last=Orders_LastContentRow(oSh,nKeyCol):If last<nStart Then Orders_NextRow=nStart Else Orders_NextRow=last+1
End Function

Function Orders_Canon(s As String) As String
    s=Trim(s):s=Replace(s,Chr(160)," "):s=Replace(s,Chr(9)," ")
    Do While InStr(s,"  ")>0:s=Replace(s,"  "," "):Loop
    Orders_Canon=UCase(s)
End Function

Function Orders_TryNumber(oCell As Object,ByRef d As Double) As Boolean
    Dim s As String
    Orders_TryNumber=False:d=0
    On Error GoTo Bad
    If oCell.Type=com.sun.star.table.CellContentType.VALUE Then d=oCell.Value:Orders_TryNumber=True:Exit Function
    s=Trim(oCell.String):If s="" Then Exit Function
    s=Replace(s,Chr(160),""):s=Replace(s," ",""):s=Replace(s,",",".")
    If Not IsNumeric(s) Then Exit Function
    d=CDbl(s):Orders_TryNumber=True
Bad:
End Function

Function Orders_CellDate(oCell As Object,bInferYear As Boolean) As Double
    Dim s As String,a As Variant,dd As Integer,mm As Integer,yy As Integer,dt As Date
    Orders_CellDate=0
    On Error GoTo Bad
    If oCell.Type=com.sun.star.table.CellContentType.VALUE And oCell.Value>0 Then Orders_CellDate=oCell.Value:Exit Function
    s=Trim(oCell.String):If s="" Then Exit Function
    s=Replace(s,"/","."):s=Replace(s,"-",".")
    Do While InStr(s,"..")>0:s=Replace(s,"..","."):Loop
    a=Split(s,".")
    If UBound(a)=1 Then
        If Not bInferYear Then Exit Function
        dd=CInt(a(0)):mm=CInt(a(1)):yy=Year(Date)
    ElseIf UBound(a)=2 Then
        dd=CInt(a(0)):mm=CInt(a(1)):yy=CInt(a(2)):If yy<100 Then yy=2000+yy
    Else
        Exit Function
    End If
    dt=DateSerial(yy,mm,dd)
    If Day(dt)<>dd Or Month(dt)<>mm Or Year(dt)<>yy Then Exit Function
    Orders_CellDate=CDbl(dt)
Bad:
End Function

Function Orders_RowReferenceDate(oSh As Object,r As Long) As Double
    Dim d As Double
    d=Orders_CellDate(oSh.getCellByPosition(12,r),True):If d>0 Then Orders_RowReferenceDate=d:Exit Function
    d=Orders_CellDate(oSh.getCellByPosition(14,r),True):If d>0 Then Orders_RowReferenceDate=d:Exit Function
    d=Orders_CellDate(oSh.getCellByPosition(13,r),True):If d>0 Then Orders_RowReferenceDate=d
End Function

Function Orders_LongSetting(oDoc As Object,sKey As String,nDefault As Long) As Long
    Dim s As String
    s=Orders_GetSetting(oDoc,sKey,CStr(nDefault))
    If IsNumeric(s) Then Orders_LongSetting=CLng(s) Else Orders_LongSetting=nDefault
End Function

Function Orders_NumKey(d As Double) As String
    Orders_NumKey=Replace(CStr(d),",",".")
End Function

Function Orders_Stamp(v As Variant) As String
    Orders_Stamp=Right("0" & CStr(Day(v)),2) & "." & Right("0" & CStr(Month(v)),2) & "." & CStr(Year(v)) & " " & Right("0" & CStr(Hour(v)),2) & ":" & Right("0" & CStr(Minute(v)),2) & ":" & Right("0" & CStr(Second(v)),2)
End Function

Function Orders_TimestampCompact(v As Variant) As String
    Orders_TimestampCompact=CStr(Year(v)) & Right("0" & CStr(Month(v)),2) & Right("0" & CStr(Day(v)),2) & Right("0" & CStr(Hour(v)),2) & Right("0" & CStr(Minute(v)),2) & Right("0" & CStr(Second(v)),2)
End Function

Function Orders_ColName(n As Long) As String
    Dim s As String,x As Long
    x=n+1
    Do While x>0
        x=x-1:s=Chr(65+(x Mod 26)) & s:x=x\26
    Loop
    Orders_ColName=s
End Function

' ============================================================================
' OPERATOR UTILITIES (optional; normal daily work remains automatic)
' ============================================================================

Sub Orders_ShowSelectedRowDiagnostics()
    Dim oCtl As Object,oSel As Object,a As Variant,oSh As Object,r As Long,s As String
    On Error GoTo EH
    oCtl=ThisComponent.CurrentController:oSel=oCtl.Selection
    a=oSel.RangeAddress:oSh=ThisComponent.Sheets.getByIndex(a.Sheet)
    If oSh.Name<>WMSORD_SHEET Then MsgBox "Откройте лист Заказы.",48,"WMS — Заказы":Exit Sub
    r=a.StartRow:If r<1 Then Exit Sub
    s="Строка " & CStr(r+1) & Chr(10) & "SourceID: " & Orders_CellText(oSh,r,WMSORD_T_ID) & Chr(10) & "Mode: " & Orders_CellText(oSh,r,WMSORD_T_MODE) & Chr(10) & "Severity: " & Orders_CellText(oSh,r,WMSORD_X_SEVERITY) & Chr(10) & Chr(10) & Orders_CellText(oSh,r,WMSORD_X_FLAGS)
    MsgBox s,64,"WMS — диагностика заказа"
    Exit Sub
EH:
    MsgBox "Не удалось прочитать выбранную строку: " & CStr(Err) & " " & Error$,48,"WMS — Заказы"
End Sub

Sub Orders_ResetSelectedStatusToAuto()
    Dim oCtl As Object,oSel As Object,a As Variant,oSh As Object,r1 As Long,r2 As Long,r As Long,c As Long
    On Error GoTo EH
    oCtl=ThisComponent.CurrentController:oSel=oCtl.Selection:a=oSel.RangeAddress:oSh=ThisComponent.Sheets.getByIndex(a.Sheet)
    If oSh.Name<>WMSORD_SHEET Then MsgBox "Откройте лист Заказы.",48,"WMS — Заказы":Exit Sub
    r1=a.StartRow:r2=a.EndRow:If r1<1 Then r1=1
    gWMSORD_Busy=True:c=Orders_FindHeader(oSh,WMSORD_X_STATUSMODE)
    For r=r1 To r2
        If c>=0 Then oSh.getCellByPosition(c,r).String="AUTO"
        oSh.getCellByPosition(16,r).String=""
        Orders_ProcessRow ThisComponent,oSh,r,True,True,False,True
        Orders_SetTech oSh,r,WMSORD_X_FINGERPRINT,Orders_BusinessFingerprint(oSh,r)
    Next r
    gWMSORD_Busy=False
    MsgBox "Для выбранных строк статус снова управляется автоматически.",64,"WMS — Заказы"
    Exit Sub
EH:
    gWMSORD_Busy=False:MsgBox "Ошибка: " & CStr(Err) & " " & Error$,48,"WMS — Заказы"
End Sub

Sub Orders_ResetSelectedCodeToAuto()
    Dim oCtl As Object,oSel As Object,a As Variant,oSh As Object,r1 As Long,r2 As Long,r As Long,c As Long
    On Error GoTo EH
    oCtl=ThisComponent.CurrentController:oSel=oCtl.Selection:a=oSel.RangeAddress:oSh=ThisComponent.Sheets.getByIndex(a.Sheet)
    If oSh.Name<>WMSORD_SHEET Then MsgBox "Откройте лист Заказы.",48,"WMS — Заказы":Exit Sub
    r1=a.StartRow:r2=a.EndRow:If r1<1 Then r1=1
    gWMSORD_Busy=True:c=Orders_FindHeader(oSh,WMSORD_X_CODEORIGIN)
    For r=r1 To r2
        If c>=0 Then oSh.getCellByPosition(c,r).String="AUTO"
        oSh.getCellByPosition(21,r).String=""
        Orders_ProcessRow ThisComponent,oSh,r,True,True,False,True
        Orders_SetTech oSh,r,WMSORD_X_FINGERPRINT,Orders_BusinessFingerprint(oSh,r)
    Next r
    gWMSORD_Busy=False
    MsgBox "Для выбранных строк код снова разрешено подбирать автоматически по точным совпадениям.",64,"WMS — Заказы"
    Exit Sub
EH:
    gWMSORD_Busy=False:MsgBox "Ошибка: " & CStr(Err) & " " & Error$,48,"WMS — Заказы"
End Sub

Sub Orders_ResetSelectedCategoryToAuto()
    Dim oCtl As Object,oSel As Object,a As Variant,oSh As Object,r1 As Long,r2 As Long,r As Long,c As Long
    On Error GoTo EH
    oCtl=ThisComponent.CurrentController:oSel=oCtl.Selection:a=oSel.RangeAddress:oSh=ThisComponent.Sheets.getByIndex(a.Sheet)
    If oSh.Name<>WMSORD_SHEET Then MsgBox "Откройте лист Заказы.",48,"WMS — Заказы":Exit Sub
    r1=a.StartRow:r2=a.EndRow:If r1<1 Then r1=1
    gWMSORD_Busy=True:c=Orders_FindHeader(oSh,WMSORD_X_CATORIGIN)
    For r=r1 To r2
        If c>=0 Then oSh.getCellByPosition(c,r).String="AUTO"
        oSh.getCellByPosition(17,r).String=""
        Orders_ProcessRow ThisComponent,oSh,r,True,True,False,True
        Orders_SetTech oSh,r,WMSORD_X_FINGERPRINT,Orders_BusinessFingerprint(oSh,r)
    Next r
    gWMSORD_Busy=False
    MsgBox "Для выбранных строк категория снова разрешена к автоматическому точному подбору.",64,"WMS — Заказы"
    Exit Sub
EH:
    gWMSORD_Busy=False:MsgBox "Ошибка: " & CStr(Err) & " " & Error$,48,"WMS — Заказы"
End Sub

