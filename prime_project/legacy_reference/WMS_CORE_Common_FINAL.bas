Option Explicit

' ============================================================================
' WMS_CORE_Common_FINAL.bas
' Shared infrastructure for the clean WMS-lite architecture.
' LibreOffice Calc / ODS / Linux / LibreOffice Basic + UNO only.
'
' PURPOSE
' - one reusable implementation of schema helpers, header lookup/cache,
'   common palette, number formats, batch locks, settings, queue, index,
'   IDs, audit/error logging and diagnostics;
' - leaf modules keep their business rules, but must not reimplement these
'   infrastructure concerns as the WMS grows.
'
' IMPORTANT
' - no business posting/balance rules live here;
' - no fuzzy matching lives here;
' - no whole-workbook scan is performed on cell OnChange;
' - bulk helpers read/write ranges through DataArray/FormulaArray;
' - Install/Upgrade are idempotent and do not delete user rows.
' ============================================================================

Global Const WMSC_VERSION = "1.0.0-FINAL"
Global Const WMSC_SCHEMA = "1.0"
Global Const WMSC_TARGET_LO = "25.8.7.3"
Global Const WMSC_TARGET_OS = "Linux"
Global Const WMSC_MODULE = "CommonCore"

Global Const WMSC_SETTINGS = "Настройки WMS"
Global Const WMSC_AUDIT = "Журнал WMS"
Global Const WMSC_ERRORS = "Ошибки WMS"
Global Const WMSC_QUEUE = "SYS_WMS_QUEUE"
Global Const WMSC_INDEX = "SYS_WMS_INDEX"

Global Const WMSC_ERR_VALIDATION = "VALIDATION"
Global Const WMSC_ERR_INTEGRITY = "INTEGRITY"
Global Const WMSC_ERR_IO = "IO"
Global Const WMSC_ERR_PERFORMANCE = "PERFORMANCE_LIMIT"
Global Const WMSC_ERR_SCHEMA = "SCHEMA"
Global Const WMSC_ERR_RUNTIME = "RUNTIME"

Global Const WMSC_ROLE_EDITABLE = "editable"
Global Const WMSC_ROLE_SYSTEM = "system"
Global Const WMSC_ROLE_OK = "ok"
Global Const WMSC_ROLE_WARN = "warn"
Global Const WMSC_ROLE_CRITICAL = "critical"

Global Const WMSC_FONT_NAME = "Liberation Sans"
Global Const WMSC_FONT_SIZE = 10
Global Const WMSC_HEADER_HEIGHT = 800
Global Const WMSC_ZOOM = 100
Global Const WMSC_DEFAULT_BATCH_CHUNK = 300
Global Const WMSC_TECH_PREFIX = "_WMS_"

Global gWMSC_LastReport As String
Global gWMSC_LastError As String
Global gWMSC_LastRunID As String
Global gWMSC_RepairedHeaders As String
Global gWMSC_IDCounter As Long
Global gWMSC_IDSession As String

' Nested batch state.
Global gWMSC_BatchDepth As Long
Global gWMSC_BatchAutoCalc As Boolean
Global gWMSC_BatchControllersLocked As Boolean
Global gWMSC_BatchActionLocked As Boolean

' Header cache: hash key = canonical sheet name | canonical header.
Global gWMSC_HReady As Boolean
Global gWMSC_HCap As Long
Global gWMSC_HKeys() As String
Global gWMSC_HCols() As Long
Global gWMSC_HBuiltCap As Long
Global gWMSC_HBuiltKeys() As String

' Queue cache: hash key = canonical sheet name | SourceID.
Global gWMSC_QReady As Boolean
Global gWMSC_QCap As Long
Global gWMSC_QKeys() As String
Global gWMSC_QRows() As Long
Global gWMSC_QNextRow As Long

' Optional status indicator used by long explicit batch operations.
Global gWMSC_Status As Object
Global gWMSC_StatusActive As Boolean

' ============================================================================
' PUBLIC ENTRY POINTS
' ============================================================================

Sub WMSCore_Install()
    Dim sReport As String
    If WMSCore_InstallCore(ThisComponent, True, sReport) Then
        MsgBox sReport, 64, "WMS — Common Core"
    Else
        MsgBox sReport, 16, "WMS — Common Core"
    End If
End Sub

Sub WMSCore_InstallQuiet()
    Dim sReport As String
    Call WMSCore_InstallCore(ThisComponent, False, sReport)
End Sub

Sub WMSCore_Upgrade()
    WMSCore_Install
End Sub

Sub WMSCore_SelfCheck()
    Dim sReport As String
    If WMSCore_RunSelfCheck(ThisComponent, sReport) Then
        MsgBox sReport, 64, "WMS — Common Core self-check"
    Else
        MsgBox sReport, 48, "WMS — Common Core self-check"
    End If
End Sub

Function WMSCore_SelfCheckText() As String
    Dim sReport As String
    If WMSCore_RunSelfCheck(ThisComponent, sReport) Then
        WMSCore_SelfCheckText = "OK|" & sReport
    Else
        WMSCore_SelfCheckText = "FAIL|" & sReport
    End If
End Function

Function WMSCore_CompileProbe() As String
    WMSCore_CompileProbe = WMSC_VERSION & "|" & WMSC_SCHEMA & "|" & WMSC_TARGET_LO
End Function

Function WMSCore_IsInstalled() As Boolean
    Dim s As String
    WMSCore_IsInstalled = False
    On Error GoTo Done
    If Not WMSCore_IsCalcDocument(ThisComponent) Then Exit Function
    If Not ThisComponent.Sheets.hasByName(WMSC_SETTINGS) Then Exit Function
    s = WMSCore_GetSetting(ThisComponent, "WMS.CommonCore.Version", "")
    WMSCore_IsInstalled = (Trim(s) <> "")
Done:
End Function

Function WMSCore_LastReportText() As String
    WMSCore_LastReportText = gWMSC_LastReport
End Function

Function WMSCore_LastErrorText() As String
    WMSCore_LastErrorText = gWMSC_LastError
End Function

Sub WMSCore_InvalidateCaches()
    WMSCore_InvalidateHeaderCache
    WMSCore_ResetQueueCache
End Sub

' ============================================================================
' INSTALL / PREFLIGHT / SYSTEM SCHEMA
' ============================================================================

Function WMSCore_InstallCore(oDoc As Object, bShow As Boolean, ByRef sReport As String) As Boolean
    Dim sErr As String
    Dim bBatch As Boolean
    WMSCore_InstallCore = False
    gWMSC_LastError = ""
    gWMSC_LastRunID = WMSCore_NewRunID()
    On Error GoTo EH

    If Not WMSCore_Preflight(oDoc, sErr) Then
        sReport = sErr
        gWMSC_LastError = sErr
        Exit Function
    End If

    bBatch = WMSCore_BeginBatch(oDoc)
    WMSCore_EnsureSystemExtensions oDoc
    WMSCore_EnsureCommonSettings oDoc
    WMSCore_EnsureQueueSchema oDoc
    WMSCore_EnsureIndexSchema oDoc
    WMSCore_InvalidateCaches

    WMSCore_LogAudit oDoc, WMSC_MODULE, "INSTALL", "", "Common Core " & WMSC_VERSION & " schema " & WMSC_SCHEMA & " установлен/обновлён"
    If gWMSC_RepairedHeaders <> "" Then WMSCore_LogAudit oDoc, WMSC_MODULE, "SYSTEM_HEADER_REPAIR", "", "Без удаления данных восстановлены заголовки: " & gWMSC_RepairedHeaders

    If bBatch Then WMSCore_EndBatch oDoc, False
    sReport = "WMS Common Core " & WMSC_VERSION & " установлен. Общие schema/queue/index/design/performance helpers готовы." & IIf(gWMSC_RepairedHeaders <> "", " Восстановлены системные заголовки: " & gWMSC_RepairedHeaders & ".", "")
    gWMSC_LastReport = sReport
    WMSCore_InstallCore = True
    Exit Function
EH:
    On Error Resume Next
    If bBatch Then WMSCore_EndBatch oDoc, False
    sErr = "[" & WMSC_ERR_RUNTIME & "] Install: " & CStr(Err) & " " & Error$
    gWMSC_LastError = sErr
    sReport = sErr
    WMSCore_LogError oDoc, WMSC_ERR_RUNTIME, WMSC_MODULE, "Install", "", "CORE_INSTALL", sErr
End Function

Function WMSCore_Preflight(oDoc As Object, ByRef sErr As String) As Boolean
    Dim a As Variant, i As Long
    WMSCore_Preflight = False
    If Not WMSCore_IsCalcDocument(oDoc) Then
        sErr = "Открытый документ не является LibreOffice Calc."
        Exit Function
    End If

    a = Array(WMSC_SETTINGS, WMSC_AUDIT, WMSC_ERRORS, WMSC_QUEUE, WMSC_INDEX)
    For i = 0 To UBound(a)
        If Not oDoc.Sheets.hasByName(CStr(a(i))) Then
            sErr = "Не найден системный лист '" & CStr(a(i)) & "'. Сначала установите WMS_00_CreateWorkbook_FINAL.bas в новую книгу."
            Exit Function
        End If
    Next i

    ' Repair a known legacy-controller defect safely: older sheet controllers
    ' could write their first setting/audit record into row 1 (index 0), replacing
    ' the system header. Insert a new header row; existing records are shifted,
    ' never deleted. This repair runs before any header-cache lookup.
    gWMSC_RepairedHeaders = ""
    WMSCore_RepairSystemHeaderRows oDoc
    WMSCore_InvalidateHeaderCache

    If WMSCore_FindHeader(oDoc.Sheets.getByName(WMSC_SETTINGS), "Параметр") < 0 Then
        sErr = "Лист 'Настройки WMS' имеет несовместимую схему. Ожидается заголовок 'Параметр'."
        Exit Function
    End If
    WMSCore_Preflight = True
End Function

Sub WMSCore_RepairSystemHeaderRows(oDoc As Object)
    WMSCore_RepairHeaderRow oDoc.Sheets.getByName(WMSC_SETTINGS), Array("Параметр", "Значение", "Описание")
    WMSCore_RepairHeaderRow oDoc.Sheets.getByName(WMSC_AUDIT), Array("Время", "RunID", "Источник", "Событие", "Описание")
    WMSCore_RepairHeaderRow oDoc.Sheets.getByName(WMSC_ERRORS), Array("Время", "RunID", "Где", "Ошибка", "Статус")
    WMSCore_RepairHeaderRow oDoc.Sheets.getByName(WMSC_QUEUE), Array("Время", "Лист", "SourceID", "RowHint", "Mode", "Status", "RunID", "Message")
    WMSCore_RepairHeaderRow oDoc.Sheets.getByName(WMSC_INDEX), Array("Тип ключа", "Ключ", "ItemID", "Код", "Артикул", "Наименование", "Ед.", "Место", "Источник", "Строка", "Обновлено")
End Sub

Sub WMSCore_RepairHeaderRow(oSh As Object, aExpected As Variant)
    Dim i As Long, ok As Boolean, hasData As Boolean, last As Long, a As Variant, rowData As Variant
    On Error GoTo Done
    ok = True
    For i = LBound(aExpected) To UBound(aExpected)
        If WMSCore_CanonHeader(oSh.getCellByPosition(i,0).String) <> WMSCore_CanonHeader(CStr(aExpected(i))) Then ok = False: Exit For
    Next i
    If ok Then Exit Sub

    last = WMSCore_LastHeaderColUncached(oSh)
    If last < UBound(aExpected) Then last = UBound(aExpected)
    a = oSh.getCellRangeByPosition(0,0,last,0).getFormulaArray()
    rowData = a(0)
    For i = 0 To UBound(rowData)
        If Trim(CStr(rowData(i))) <> "" Then hasData = True: Exit For
    Next i
    If hasData Then oSh.Rows.insertByIndex(0,1)
    For i = LBound(aExpected) To UBound(aExpected)
        oSh.getCellByPosition(i,0).String = CStr(aExpected(i))
    Next i
    If gWMSC_RepairedHeaders <> "" Then gWMSC_RepairedHeaders = gWMSC_RepairedHeaders & ", "
    gWMSC_RepairedHeaders = gWMSC_RepairedHeaders & oSh.Name
Done:
End Sub

Sub WMSCore_EnsureSystemExtensions(oDoc As Object)
    Dim oErr As Object, oLog As Object
    On Error GoTo EH
    oErr = oDoc.Sheets.getByName(WMSC_ERRORS)
    Call WMSCore_EnsureColumn(oErr, "_WMS_Category", 5, True)
    Call WMSCore_EnsureColumn(oErr, "_WMS_Module", 6, True)
    Call WMSCore_EnsureColumn(oErr, "_WMS_SourceID", 7, True)
    Call WMSCore_EnsureColumn(oErr, "_WMS_ErrorCode", 8, True)
    Call WMSCore_EnsureColumn(oErr, "_WMS_LastTouch", 9, True)

    oLog = oDoc.Sheets.getByName(WMSC_AUDIT)
    Call WMSCore_EnsureColumn(oLog, "_WMS_Category", 5, True)
    Call WMSCore_EnsureColumn(oLog, "_WMS_SourceID", 6, True)
    Call WMSCore_EnsureColumn(oLog, "_WMS_Module", 7, True)
    Call WMSCore_EnsureColumn(oLog, "_WMS_LastTouch", 8, True)
    Exit Sub
EH:
    WMSCore_LogError oDoc, WMSC_ERR_SCHEMA, WMSC_MODULE, "EnsureSystemExtensions", "", "CORE_SYS_EXT", CStr(Err) & " " & Error$
End Sub

Sub WMSCore_EnsureCommonSettings(oDoc As Object)
    WMSCore_SetSetting oDoc, "WMS.CommonCore.Version", WMSC_VERSION, "Версия общего инфраструктурного Core", True
    WMSCore_SetSetting oDoc, "WMS.CommonCore.Schema", WMSC_SCHEMA, "Версия схемы общего инфраструктурного Core", True
    WMSCore_SetSetting oDoc, "WMS.CommonCore.TargetLO", WMSC_TARGET_LO, "Целевая версия LibreOffice", True
    WMSCore_SetSetting oDoc, "WMS.UI.FontName", WMSC_FONT_NAME, "Единый шрифт пользовательских листов", False
    WMSCore_SetSetting oDoc, "WMS.UI.FontSize", CStr(WMSC_FONT_SIZE), "Размер текста данных", False
    WMSCore_SetSetting oDoc, "WMS.UI.Zoom", CStr(WMSC_ZOOM), "Масштаб рабочих листов, %", False
    WMSCore_SetSetting oDoc, "WMS.Batch.ChunkSize", CStr(WMSC_DEFAULT_BATCH_CHUNK), "Технический размер порции batch-обработки", False
End Sub

Sub WMSCore_EnsureQueueSchema(oDoc As Object)
    Dim oSh As Object, a As Variant, i As Long
    oSh = oDoc.Sheets.getByName(WMSC_QUEUE)
    a = Array("Время", "Лист", "SourceID", "RowHint", "Mode", "Status", "RunID", "Message")
    For i = 0 To UBound(a)
        If Trim(oSh.getCellByPosition(i,0).String) = "" Then oSh.getCellByPosition(i,0).String = CStr(a(i))
    Next i
    On Error Resume Next
    oSh.IsVisible = False
    On Error GoTo 0
End Sub

Sub WMSCore_EnsureIndexSchema(oDoc As Object)
    Dim oSh As Object, a As Variant, i As Long
    oSh = oDoc.Sheets.getByName(WMSC_INDEX)
    a = Array("Тип ключа", "Ключ", "ItemID", "Код", "Артикул", "Наименование", "Ед.", "Место", "Источник", "Строка", "Обновлено")
    For i = 0 To UBound(a)
        If Trim(oSh.getCellByPosition(i,0).String) = "" Then oSh.getCellByPosition(i,0).String = CStr(a(i))
    Next i
    On Error Resume Next
    oSh.IsVisible = False
    On Error GoTo 0
End Sub

' ============================================================================
' IDEMPOTENT COLUMN / HEADER MAP / HEADER CACHE
' ============================================================================

Function WMSCore_EnsureColumn(oSheet As Object, sHeader As String, nPos As Long, bHidden As Boolean) As Long
    Dim c As Long, last As Long, target As Long, sExisting As String
    WMSCore_EnsureColumn = -1
    On Error GoTo EH
    If Trim(sHeader) = "" Then Exit Function

    c = WMSCore_FindHeader(oSheet, sHeader)
    If c >= 0 Then
        On Error Resume Next
        oSheet.Columns.getByIndex(c).IsVisible = Not bHidden
        On Error GoTo EH
        WMSCore_EnsureColumn = c
        Exit Function
    End If

    last = WMSCore_LastHeaderCol(oSheet)
    If last < 0 Then last = -1
    target = nPos
    If target < 0 Or target > last + 1 Then target = last + 1
    If target < 0 Then target = 0

    If target <= last Then
        sExisting = Trim(oSheet.getCellByPosition(target,0).String)
        If sExisting <> "" Then
            oSheet.Columns.insertByIndex(target, 1)
        End If
    End If

    oSheet.getCellByPosition(target,0).String = sHeader
    oSheet.Columns.getByIndex(target).IsVisible = Not bHidden
    WMSCore_InvalidateHeaderCache
    WMSCore_EnsureColumn = target
    Exit Function
EH:
    gWMSC_LastError = "EnsureColumn(" & sHeader & "): " & CStr(Err) & " " & Error$
End Function

Function WMSCore_GetHeaderMap(oSheet As Object) As Object
    Dim cMap As New Collection
    Dim aData As Variant, aRow As Variant, last As Long, c As Long
    Dim k As String
    On Error GoTo Done
    last = WMSCore_LastHeaderCol(oSheet)
    If last < 0 Then WMSCore_GetHeaderMap = cMap: Exit Function
    aData = oSheet.getCellRangeByPosition(0,0,last,0).getDataArray()
    aRow = aData(0)
    For c = 0 To last
        k = WMSCore_CanonHeader(CStr(aRow(c)))
        If k <> "" Then
            On Error Resume Next
            cMap.Add c, k
            On Error GoTo Done
        End If
    Next c
Done:
    WMSCore_GetHeaderMap = cMap
End Function

Function WMSCore_HeaderIndex(oHeaderMap As Object, sHeader As String) As Long
    Dim k As String
    WMSCore_HeaderIndex = -1
    On Error GoTo Done
    k = WMSCore_CanonHeader(sHeader)
    If k = "" Then Exit Function
    WMSCore_HeaderIndex = CLng(oHeaderMap.Item(k))
Done:
End Function

Function WMSCore_FindHeader(oSheet As Object, sHeader As String) As Long
    Dim key As String, slot As Long, found As Boolean
    WMSCore_FindHeader = -1
    If Trim(sHeader) = "" Then Exit Function
    If Not WMSCore_EnsureHeaderSheetCache(oSheet) Then Exit Function
    key = WMSCore_Canon(oSheet.Name) & "|" & WMSCore_CanonHeader(sHeader)
    slot = WMSCore_HSlot(key, found)
    If slot >= 0 And found Then WMSCore_FindHeader = gWMSC_HCols(slot)
End Function

Function WMSCore_EnsureHeaderSheetCache(oSheet As Object) As Boolean
    Dim sheetKey As String, slotBuilt As Long, built As Boolean
    Dim last As Long, aData As Variant, aRow As Variant, c As Long
    Dim key As String, slot As Long, found As Boolean
    WMSCore_EnsureHeaderSheetCache = False
    On Error GoTo EH
    WMSCore_InitHeaderCache
    sheetKey = WMSCore_Canon(oSheet.Name)
    slotBuilt = WMSCore_HBuiltSlot(sheetKey, built)
    If slotBuilt >= 0 And built Then WMSCore_EnsureHeaderSheetCache = True: Exit Function

    last = WMSCore_LastHeaderColUncached(oSheet)
    If last < 0 Then
        If slotBuilt >= 0 Then gWMSC_HBuiltKeys(slotBuilt) = sheetKey
        WMSCore_EnsureHeaderSheetCache = True
        Exit Function
    End If

    aData = oSheet.getCellRangeByPosition(0,0,last,0).getDataArray()
    aRow = aData(0)
    For c = 0 To last
        If WMSCore_CanonHeader(CStr(aRow(c))) <> "" Then
            key = sheetKey & "|" & WMSCore_CanonHeader(CStr(aRow(c)))
            slot = WMSCore_HSlot(key, found)
            If slot >= 0 Then
                If Not found Then gWMSC_HKeys(slot) = key
                gWMSC_HCols(slot) = c
            End If
        End If
    Next c
    If slotBuilt >= 0 Then gWMSC_HBuiltKeys(slotBuilt) = sheetKey
    WMSCore_EnsureHeaderSheetCache = True
    Exit Function
EH:
    gWMSC_LastError = "HeaderCache(" & oSheet.Name & "): " & CStr(Err) & " " & Error$
End Function

Sub WMSCore_InitHeaderCache()
    If gWMSC_HReady Then Exit Sub
    gWMSC_HCap = 16384
    ReDim gWMSC_HKeys(0 To gWMSC_HCap-1)
    ReDim gWMSC_HCols(0 To gWMSC_HCap-1)
    gWMSC_HBuiltCap = 256
    ReDim gWMSC_HBuiltKeys(0 To gWMSC_HBuiltCap-1)
    gWMSC_HReady = True
End Sub

Sub WMSCore_InvalidateHeaderCache()
    gWMSC_HReady = False
    gWMSC_HCap = 0
    gWMSC_HBuiltCap = 0
End Sub

Function WMSCore_HSlot(sKey As String, ByRef bFound As Boolean) As Long
    Dim slot As Long, start As Long
    WMSCore_HSlot = -1: bFound = False
    If Not gWMSC_HReady Or gWMSC_HCap <= 0 Then Exit Function
    slot = WMSCore_HashSlotStart(sKey, gWMSC_HCap): start = slot
    Do
        If gWMSC_HKeys(slot) = "" Then WMSCore_HSlot = slot: Exit Function
        If gWMSC_HKeys(slot) = sKey Then bFound = True: WMSCore_HSlot = slot: Exit Function
        slot = slot + 1: If slot >= gWMSC_HCap Then slot = 0
    Loop While slot <> start
End Function

Function WMSCore_HBuiltSlot(sKey As String, ByRef bFound As Boolean) As Long
    Dim slot As Long, start As Long
    WMSCore_HBuiltSlot = -1: bFound = False
    If gWMSC_HBuiltCap <= 0 Then Exit Function
    slot = WMSCore_HashSlotStart(sKey, gWMSC_HBuiltCap): start = slot
    Do
        If gWMSC_HBuiltKeys(slot) = "" Then WMSCore_HBuiltSlot = slot: Exit Function
        If gWMSC_HBuiltKeys(slot) = sKey Then bFound = True: WMSCore_HBuiltSlot = slot: Exit Function
        slot = slot + 1: If slot >= gWMSC_HBuiltCap Then slot = 0
    Loop While slot <> start
End Function

Function WMSCore_LastHeaderCol(oSheet As Object) As Long
    WMSCore_LastHeaderCol = WMSCore_LastHeaderColUncached(oSheet)
End Function

Function WMSCore_LastHeaderColUncached(oSheet As Object) As Long
    Dim cur As Object, addr As Variant, endCol As Long, aData As Variant, aRow As Variant, c As Long
    WMSCore_LastHeaderColUncached = -1
    On Error GoTo Done
    cur = oSheet.createCursor(): cur.gotoEndOfUsedArea(True): addr = cur.RangeAddress
    endCol = addr.EndColumn
    If endCol < 0 Then Exit Function
    aData = oSheet.getCellRangeByPosition(0,0,endCol,0).getDataArray()
    aRow = aData(0)
    For c = endCol To 0 Step -1
        If Trim(CStr(aRow(c))) <> "" Then WMSCore_LastHeaderColUncached = c: Exit Function
    Next c
Done:
End Function

Function WMSCore_LastContentRow(oSheet As Object, nLastCol As Long) As Long
    Dim cur As Object, addr As Variant, maxR As Long, topR As Long
    Dim a As Variant, rr As Long, cc As Long, rowData As Variant
    WMSCore_LastContentRow = 0
    On Error GoTo Done
    cur = oSheet.createCursor(): cur.gotoEndOfUsedArea(True): addr = cur.RangeAddress
    maxR = addr.EndRow
    If nLastCol < 0 Then nLastCol = addr.EndColumn
    If nLastCol > addr.EndColumn Then nLastCol = addr.EndColumn
    If maxR < 1 Or nLastCol < 0 Then Exit Function
    maxR = WMSCore_ContentEndBound(oSheet, nLastCol, maxR)

    Do While maxR >= 1
        topR = maxR - 255: If topR < 1 Then topR = 1
        a = oSheet.getCellRangeByPosition(0,topR,nLastCol,maxR).getFormulaArray()
        For rr = UBound(a) To 0 Step -1
            rowData = a(rr)
            For cc = 0 To UBound(rowData)
                If Trim(CStr(rowData(cc))) <> "" Then WMSCore_LastContentRow = topR + rr: Exit Function
            Next cc
        Next rr
        maxR = topR - 1
    Loop
Done:
End Function

Function WMSCore_NextRow(oSheet As Object, nKeyCol As Long, nMinRow As Long) As Long
    Dim last As Long
    last = WMSCore_LastContentRow(oSheet, nKeyCol)
    If last < nMinRow Then WMSCore_NextRow = nMinRow Else WMSCore_NextRow = last + 1
End Function

' ============================================================================
' COMMON DESIGN / PALETTE / FORMAT HELPERS
' ============================================================================

Sub WMSCore_ApplyPalette(oRange As Object, sRole As String)
    Dim role As String
    On Error GoTo Done
    role = LCase(Trim(sRole))
    oRange.CharFontName = WMSC_FONT_NAME
    oRange.CharHeight = WMSC_FONT_SIZE
    oRange.CharColor = RGB(31,41,55)
    Select Case role
        Case WMSC_ROLE_EDITABLE
            oRange.CellBackColor = RGB(255,255,255)
        Case WMSC_ROLE_SYSTEM
            oRange.CellBackColor = RGB(242,242,242)
        Case WMSC_ROLE_OK
            oRange.CellBackColor = RGB(217,234,211)
        Case WMSC_ROLE_WARN
            oRange.CellBackColor = RGB(252,232,178)
        Case WMSC_ROLE_CRITICAL
            oRange.CellBackColor = RGB(244,199,195)
        Case Else
            oRange.CellBackColor = RGB(255,255,255)
    End Select
Done:
End Sub

Sub WMSCore_ApplyStandardHeader(oSheet As Object, nFirstCol As Long, nLastCol As Long)
    Dim oHdr As Object
    On Error GoTo Done
    If nLastCol < nFirstCol Then Exit Sub
    oHdr = oSheet.getCellRangeByPosition(nFirstCol,0,nLastCol,0)
    oHdr.CellBackColor = RGB(31,58,95)
    oHdr.CharColor = RGB(255,255,255)
    oHdr.CharFontName = WMSC_FONT_NAME
    oHdr.CharHeight = WMSC_FONT_SIZE
    oHdr.CharWeight = 150
    oHdr.HoriJustify = com.sun.star.table.CellHoriJustify.CENTER
    oHdr.VertJustify = com.sun.star.table.CellVertJustify.CENTER
    oHdr.IsTextWrapped = True
    oSheet.Rows.getByIndex(0).Height = WMSC_HEADER_HEIGHT
Done:
End Sub

Sub WMSCore_ApplyTableGrid(oSheet As Object, c1 As Long, r1 As Long, c2 As Long, r2 As Long)
    Dim oRange As Object, oTop As Object, oLeft As Object
    Dim aLine As New com.sun.star.table.BorderLine2
    Dim aOuter As New com.sun.star.table.BorderLine2
    On Error GoTo Done
    If c2 < c1 Or r2 < r1 Then Exit Sub
    oRange = oSheet.getCellRangeByPosition(c1,r1,c2,r2)
    aOuter.Color = RGB(100,116,139)
    aOuter.LineStyle = com.sun.star.table.BorderLineStyle.SOLID
    aOuter.LineWidth = 50
    oTop = oSheet.getCellRangeByPosition(c1,r1,c2,r1): oTop.TopBorder2 = aOuter
    oLeft = oSheet.getCellRangeByPosition(c1,r1,c1,r2): oLeft.LeftBorder2 = aOuter
    aLine.Color = RGB(176,186,198)
    aLine.LineStyle = com.sun.star.table.BorderLineStyle.SOLID
    aLine.LineWidth = 35
    oRange.BottomBorder2 = aLine
    oRange.RightBorder2 = aLine
Done:
End Sub

Sub WMSCore_ApplyWorkingView(oDoc As Object, oSheet As Object)
    Dim oCtrl As Object, oView As Object
    On Error GoTo Done
    oCtrl = oDoc.CurrentController
    If IsNull(oCtrl) Then Exit Sub
    oCtrl.setActiveSheet(oSheet)
    oView = oCtrl.ViewSettings
    On Error Resume Next
    oView.ShowGrid = False
    oView.ZoomType = com.sun.star.view.DocumentZoomType.BY_VALUE
    oView.ZoomValue = CLng(WMSCore_GetSetting(oDoc, "WMS.UI.Zoom", CStr(WMSC_ZOOM)))
    On Error GoTo 0
Done:
End Sub

Sub WMSCore_ApplyNumberFormat(oDoc As Object, oRange As Object, sKind As String)
    Dim k As Long
    k = WMSCore_NumberFormatKey(oDoc, sKind)
    If k >= 0 Then oRange.NumberFormat = k
End Sub

Function WMSCore_NumberFormatKey(oDoc As Object, sKind As String) As Long
    Dim oFormats As Object, aLocale As New com.sun.star.lang.Locale
    Dim fmt As String, k As Long, baseKey As Long
    WMSCore_NumberFormatKey = -1
    On Error GoTo Done
    oFormats = oDoc.NumberFormats
    aLocale.Language = "ru": aLocale.Country = "RU"
    Select Case UCase(Trim(sKind))
        Case "QTY", "NUMBER"
            WMSCore_NumberFormatKey = oFormats.getStandardIndex(aLocale): Exit Function
        Case "MONEY"
            baseKey = oFormats.getStandardFormat(com.sun.star.util.NumberFormat.NUMBER, aLocale)
            fmt = oFormats.generateFormat(baseKey, aLocale, False, False, 2, 1)
        Case "DATE"
            fmt = "DD.MM.YYYY"
        Case "DATETIME"
            fmt = "DD.MM.YYYY HH:MM:SS"
        Case "TEXT"
            fmt = "@"
        Case Else
            Exit Function
    End Select
    k = oFormats.queryKey(fmt, aLocale, False)
    If k = -1 Then k = oFormats.addNew(fmt, aLocale)
    WMSCore_NumberFormatKey = k
Done:
End Function

' ============================================================================
' BATCH LOCKS / PROGRESS / PERFORMANCE
' ============================================================================

Function WMSCore_BeginBatch(oDoc As Object) As Boolean
    WMSCore_BeginBatch = False
    On Error GoTo EH
    gWMSC_BatchDepth = gWMSC_BatchDepth + 1
    If gWMSC_BatchDepth > 1 Then WMSCore_BeginBatch = True: Exit Function

    gWMSC_BatchControllersLocked = False
    gWMSC_BatchActionLocked = False
    gWMSC_BatchAutoCalc = True

    On Error Resume Next
    gWMSC_BatchAutoCalc = oDoc.isAutomaticCalculationEnabled()
    oDoc.lockControllers(): If Err = 0 Then gWMSC_BatchControllersLocked = True
    Err = 0
    oDoc.addActionLock(): If Err = 0 Then gWMSC_BatchActionLocked = True
    Err = 0
    oDoc.enableAutomaticCalculation(False)
    On Error GoTo EH
    WMSCore_BeginBatch = True
    Exit Function
EH:
    gWMSC_LastError = "BeginBatch: " & CStr(Err) & " " & Error$
    WMSCore_BeginBatch = True
End Function

Sub WMSCore_EndBatch(oDoc As Object, bCalculate As Boolean)
    On Error Resume Next
    If gWMSC_BatchDepth <= 0 Then Exit Sub
    gWMSC_BatchDepth = gWMSC_BatchDepth - 1
    If gWMSC_BatchDepth > 0 Then Exit Sub

    oDoc.enableAutomaticCalculation(gWMSC_BatchAutoCalc)
    If bCalculate And gWMSC_BatchAutoCalc Then oDoc.calculateAll()
    If gWMSC_BatchActionLocked Then oDoc.removeActionLock()
    If gWMSC_BatchControllersLocked Then oDoc.unlockControllers()
    gWMSC_BatchControllersLocked = False
    gWMSC_BatchActionLocked = False
    On Error GoTo 0
End Sub

Sub WMSCore_StatusStart(oDoc As Object, sText As String, nMax As Long)
    gWMSC_StatusActive = False
    On Error GoTo Done
    gWMSC_Status = oDoc.CurrentController.Frame.createStatusIndicator()
    If nMax < 1 Then nMax = 1
    gWMSC_Status.start(sText, nMax)
    gWMSC_StatusActive = True
Done:
End Sub

Sub WMSCore_StatusUpdate(nValue As Long, Optional sText As String)
    On Error GoTo Done
    If Not gWMSC_StatusActive Then Exit Sub
    If Trim(sText) <> "" Then gWMSC_Status.setText(sText)
    gWMSC_Status.setValue(nValue)
Done:
End Sub

Sub WMSCore_StatusEnd()
    On Error Resume Next
    If gWMSC_StatusActive Then gWMSC_Status.end()
    gWMSC_StatusActive = False
    On Error GoTo 0
End Sub

Sub WMSCore_RecordPerformance(oDoc As Object, sMetric As String, nRows As Long, nSeconds As Double)
    Dim key As String
    key = "PERF." & WMSCore_CanonKey(sMetric)
    WMSCore_SetSetting oDoc, key & ".Rows", CStr(nRows), "Последний фактический объём замера", True
    WMSCore_SetSetting oDoc, key & ".Seconds", WMSCore_NumText(nSeconds), "Последнее фактическое время, сек", True
    WMSCore_SetSetting oDoc, key & ".Updated", WMSCore_Stamp(Now), "Дата последнего замера", True
End Sub

' ============================================================================
' SETTINGS
' ============================================================================

Sub WMSCore_SetSetting(oDoc As Object, sKey As String, sValue As String, sDescription As String, bForce As Boolean)
    Dim oSh As Object, r As Long, cKey As Long, cVal As Long, cDesc As Long
    On Error GoTo Done
    If Not oDoc.Sheets.hasByName(WMSC_SETTINGS) Then Exit Sub
    oSh = oDoc.Sheets.getByName(WMSC_SETTINGS)
    cKey = WMSCore_FindHeader(oSh, "Параметр")
    cVal = WMSCore_FindHeader(oSh, "Значение")
    cDesc = WMSCore_FindHeader(oSh, "Описание")
    If cKey < 0 Or cVal < 0 Then Exit Sub
    r = WMSCore_FindSettingRow(oSh, cKey, sKey)
    If r < 0 Then
        r = WMSCore_NextRow(oSh, cKey, 1)
        oSh.getCellByPosition(cKey,r).String = sKey
        oSh.getCellByPosition(cVal,r).String = sValue
        If cDesc >= 0 Then oSh.getCellByPosition(cDesc,r).String = sDescription
    ElseIf bForce Then
        oSh.getCellByPosition(cVal,r).String = sValue
        If cDesc >= 0 And Trim(sDescription) <> "" Then oSh.getCellByPosition(cDesc,r).String = sDescription
    Else
        If Trim(oSh.getCellByPosition(cVal,r).String) = "" Then oSh.getCellByPosition(cVal,r).String = sValue
        If cDesc >= 0 And Trim(oSh.getCellByPosition(cDesc,r).String) = "" Then oSh.getCellByPosition(cDesc,r).String = sDescription
    End If
Done:
End Sub

Function WMSCore_GetSetting(oDoc As Object, sKey As String, sDefault As String) As String
    Dim oSh As Object, r As Long, cKey As Long, cVal As Long
    WMSCore_GetSetting = sDefault
    On Error GoTo Done
    If Not oDoc.Sheets.hasByName(WMSC_SETTINGS) Then Exit Function
    oSh = oDoc.Sheets.getByName(WMSC_SETTINGS)
    cKey = WMSCore_FindHeader(oSh, "Параметр"): cVal = WMSCore_FindHeader(oSh, "Значение")
    If cKey < 0 Or cVal < 0 Then Exit Function
    r = WMSCore_FindSettingRow(oSh, cKey, sKey)
    If r >= 0 Then
        If Trim(oSh.getCellByPosition(cVal,r).String) <> "" Then WMSCore_GetSetting = oSh.getCellByPosition(cVal,r).String
    End If
Done:
End Function

Function WMSCore_FindSettingRow(oSh As Object, cKey As Long, sKey As String) As Long
    Dim last As Long, a As Variant, i As Long
    WMSCore_FindSettingRow = -1
    last = WMSCore_LastContentRow(oSh, cKey)
    If last < 1 Then Exit Function
    On Error GoTo Done
    a = oSh.getCellRangeByPosition(cKey,1,cKey,last).getDataArray()
    For i = 0 To UBound(a)
        If WMSCore_Canon(CStr(a(i)(0))) = WMSCore_Canon(sKey) Then WMSCore_FindSettingRow = i + 1: Exit Function
    Next i
Done:
End Function

' ============================================================================
' QUEUE — SINGLE AND BULK DEDUPLICATED ENQUEUE
' ============================================================================

Sub WMSCore_EnqueueDirty(sSheet As String, sSourceID As String, sMode As String)
    WMSCore_EnqueueDirtyEx ThisComponent, sSheet, sSourceID, 0, sMode, "CHANGED"
End Sub

Sub WMSCore_EnqueueDirtyEx(oDoc As Object, sSheet As String, sSourceID As String, nRowHint As Long, sMode As String, sMessage As String)
    Dim oQ As Object, key As String, found As Boolean, slot As Long, qrow As Long
    Dim aRow As Variant
    On Error GoTo EH
    If Trim(sSheet) = "" Or Trim(sSourceID) = "" Then Exit Sub
    If Not oDoc.Sheets.hasByName(WMSC_QUEUE) Then Exit Sub
    WMSCore_EnsureQueueSchema oDoc
    oQ = oDoc.Sheets.getByName(WMSC_QUEUE)
    If Not WMSCore_EnsureQueueCache(oDoc) Then Exit Sub
    key = WMSCore_QueueKey(sSheet, sSourceID)
    slot = WMSCore_QSlot(key, found)
    If slot < 0 Then
        WMSCore_ResetQueueCache
        If Not WMSCore_EnsureQueueCache(oDoc) Then Exit Sub
        slot = WMSCore_QSlot(key, found)
        If slot < 0 Then Exit Sub
    End If
    If found Then
        qrow = gWMSC_QRows(slot)
    Else
        qrow = gWMSC_QNextRow: If qrow < 1 Then qrow = 1
        gWMSC_QNextRow = qrow + 1
        gWMSC_QKeys(slot) = key: gWMSC_QRows(slot) = qrow
    End If
    If Trim(sMode) = "" Then sMode = "CHANGED"
    aRow = Array(Array(WMSCore_Stamp(Now), sSheet, sSourceID, CDbl(nRowHint), sMode, "PENDING", "", sMessage))
    oQ.getCellRangeByPosition(0,qrow,7,qrow).setDataArray(aRow)
    oQ.IsVisible = False
    Exit Sub
EH:
    WMSCore_ResetQueueCache
    WMSCore_LogError oDoc, WMSC_ERR_RUNTIME, WMSC_MODULE, "EnqueueDirty", sSourceID, "QUEUE_ENQUEUE", CStr(Err) & " " & Error$
End Sub

Function WMSCore_EnqueueBatch(oDoc As Object, sSheet As String, aSourceIDs As Variant, aRowHints As Variant, aModes As Variant, sMessage As String) As Long
    Dim oQ As Object, i As Long, n As Long, sID As String, mode As String, hint As Long
    Dim key As String, found As Boolean, slot As Long, qrow As Long, maxRow As Long
    Dim aData As Variant, aRow As Variant, idx As Long
    Dim targets() As Long, ids() As String, modes() As String, hints() As Long
    WMSCore_EnqueueBatch = 0
    On Error GoTo EH
    If Trim(sSheet) = "" Then Exit Function
    If Not oDoc.Sheets.hasByName(WMSC_QUEUE) Then Exit Function
    n = UBound(aSourceIDs) - LBound(aSourceIDs) + 1
    If n <= 0 Then Exit Function
    ReDim targets(0 To n-1): ReDim ids(0 To n-1): ReDim modes(0 To n-1): ReDim hints(0 To n-1)

    WMSCore_EnsureQueueSchema oDoc
    oQ = oDoc.Sheets.getByName(WMSC_QUEUE)
    If Not WMSCore_EnsureQueueCache(oDoc) Then Exit Function
    maxRow = gWMSC_QNextRow - 1

    For i = 0 To n-1
        sID = Trim(CStr(aSourceIDs(LBound(aSourceIDs)+i)))
        If sID <> "" Then
            key = WMSCore_QueueKey(sSheet, sID)
            slot = WMSCore_QSlot(key, found)
            If slot < 0 Then
                WMSCore_ResetQueueCache
                Call WMSCore_EnsureQueueCache(oDoc)
                slot = WMSCore_QSlot(key, found)
            End If
            If slot >= 0 Then
                If found Then
                    qrow = gWMSC_QRows(slot)
                Else
                    qrow = gWMSC_QNextRow: gWMSC_QNextRow = qrow + 1
                    gWMSC_QKeys(slot) = key: gWMSC_QRows(slot) = qrow
                End If
                mode = "CHANGED"
                On Error Resume Next
                mode = Trim(CStr(aModes(LBound(aModes)+i))): If mode = "" Then mode = "CHANGED"
                hint = CLng(aRowHints(LBound(aRowHints)+i))
                On Error GoTo EH
                targets(i) = qrow: ids(i) = sID: modes(i) = mode: hints(i) = hint
                If qrow > maxRow Then maxRow = qrow
            End If
        End If
    Next i

    If maxRow < 1 Then Exit Function
    aData = oQ.getCellRangeByPosition(0,1,7,maxRow).getDataArray()
    For i = 0 To n-1
        If targets(i) >= 1 And ids(i) <> "" Then
            idx = targets(i) - 1
            aRow = Array(WMSCore_Stamp(Now), sSheet, ids(i), CDbl(hints(i)), modes(i), "PENDING", "", sMessage)
            aData(idx) = aRow
            WMSCore_EnqueueBatch = WMSCore_EnqueueBatch + 1
        End If
    Next i
    oQ.getCellRangeByPosition(0,1,7,maxRow).setDataArray(aData)
    oQ.IsVisible = False
    Exit Function
EH:
    WMSCore_ResetQueueCache
    WMSCore_LogError oDoc, WMSC_ERR_RUNTIME, WMSC_MODULE, "EnqueueBatch", "", "QUEUE_BATCH", CStr(Err) & " " & Error$
End Function

Function WMSCore_EnsureQueueCache(oDoc As Object) As Boolean
    Dim oQ As Object, last As Long, a As Variant, i As Long, rowData As Variant
    Dim key As String, status As String, slot As Long, found As Boolean
    WMSCore_EnsureQueueCache = False
    On Error GoTo EH
    If gWMSC_QReady Then WMSCore_EnsureQueueCache = True: Exit Function
    oQ = oDoc.Sheets.getByName(WMSC_QUEUE)
    last = WMSCore_LastContentRow(oQ, 7)
    gWMSC_QCap = WMSCore_NextPow2((last + 2048) * 4)
    If gWMSC_QCap < 8192 Then gWMSC_QCap = 8192
    ReDim gWMSC_QKeys(0 To gWMSC_QCap-1)
    ReDim gWMSC_QRows(0 To gWMSC_QCap-1)
    If last >= 1 Then
        a = oQ.getCellRangeByPosition(0,1,7,last).getDataArray()
        For i = 0 To UBound(a)
            rowData = a(i)
            status = UCase(Trim(CStr(rowData(5))))
            If Trim(CStr(rowData(1))) <> "" And Trim(CStr(rowData(2))) <> "" And (status = "PENDING" Or status = "ERROR" Or status = "RETRY" Or status = "") Then
                key = WMSCore_QueueKey(CStr(rowData(1)), CStr(rowData(2)))
                slot = WMSCore_QSlot(key, found)
                If slot >= 0 Then gWMSC_QKeys(slot) = key: gWMSC_QRows(slot) = i + 1
            End If
        Next i
    End If
    gWMSC_QNextRow = last + 1: If gWMSC_QNextRow < 1 Then gWMSC_QNextRow = 1
    gWMSC_QReady = True
    WMSCore_EnsureQueueCache = True
    Exit Function
EH:
    WMSCore_ResetQueueCache
End Function

Sub WMSCore_ResetQueueCache()
    gWMSC_QReady = False: gWMSC_QCap = 0: gWMSC_QNextRow = 1
End Sub

Function WMSCore_QSlot(sKey As String, ByRef bFound As Boolean) As Long
    Dim slot As Long, start As Long
    WMSCore_QSlot = -1: bFound = False
    If Not gWMSC_QReady And gWMSC_QCap <= 0 Then Exit Function
    If gWMSC_QCap <= 0 Then Exit Function
    slot = WMSCore_HashSlotStart(sKey, gWMSC_QCap): start = slot
    Do
        If gWMSC_QKeys(slot) = "" Then WMSCore_QSlot = slot: Exit Function
        If gWMSC_QKeys(slot) = sKey Then bFound = True: WMSCore_QSlot = slot: Exit Function
        slot = slot + 1: If slot >= gWMSC_QCap Then slot = 0
    Loop While slot <> start
End Function

Function WMSCore_QueueKey(sSheet As String, sSourceID As String) As String
    WMSCore_QueueKey = WMSCore_Canon(sSheet) & "|" & Trim(sSourceID)
End Function

' ============================================================================
' INDEX — EXACT, BATCH-BUILT, NO FUZZY MERGE
' ============================================================================

Function WMSCore_BuildIndex(sIndexType As String, sSourceSheet As String, sKeyHeader As String, sItemIDHeader As String, sCodeHeader As String, sArticleHeader As String, sNameHeader As String, sUnitHeader As String, sPlaceHeader As String) As Long
    Dim oDoc As Object
    oDoc = ThisComponent
    WMSCore_BuildIndex = WMSCore_BuildIndexCore(oDoc, sIndexType, sSourceSheet, sKeyHeader, sItemIDHeader, sCodeHeader, sArticleHeader, sNameHeader, sUnitHeader, sPlaceHeader)
End Function

Function WMSCore_BuildIndexCore(oDoc As Object, sIndexType As String, sSourceSheet As String, sKeyHeader As String, sItemIDHeader As String, sCodeHeader As String, sArticleHeader As String, sNameHeader As String, sUnitHeader As String, sPlaceHeader As String) As Long
    Dim oSrc As Object, oIdx As Object, last As Long, lastCol As Long, aSrc As Variant
    Dim cKey As Long, cItem As Long, cCode As Long, cArt As Long, cName As Long, cUnit As Long, cPlace As Long
    Dim aOld As Variant, oldLast As Long, aOut() As Variant, outN As Long, keepN As Long, newN As Long, totalN As Long
    Dim i As Long, rowData As Variant, key As String, typ As String, src As String, stamp As String
    Dim bBatch As Boolean
    WMSCore_BuildIndexCore = 0
    On Error GoTo EH
    If Not oDoc.Sheets.hasByName(sSourceSheet) Or Not oDoc.Sheets.hasByName(WMSC_INDEX) Then Exit Function
    If Trim(sIndexType) = "" Or Trim(sKeyHeader) = "" Then Exit Function
    oSrc = oDoc.Sheets.getByName(sSourceSheet): oIdx = oDoc.Sheets.getByName(WMSC_INDEX)

    ' Index rebuild is an explicit batch operation; refresh schema cache once so
    ' a legitimate installer/upgrade performed earlier in the session cannot
    ' leave stale header positions behind.
    WMSCore_InvalidateHeaderCache
    cKey = WMSCore_FindHeader(oSrc, sKeyHeader): If cKey < 0 Then Exit Function
    cItem = WMSCore_FindHeader(oSrc, sItemIDHeader)
    cCode = WMSCore_FindHeader(oSrc, sCodeHeader)
    cArt = WMSCore_FindHeader(oSrc, sArticleHeader)
    cName = WMSCore_FindHeader(oSrc, sNameHeader)
    cUnit = WMSCore_FindHeader(oSrc, sUnitHeader)
    cPlace = WMSCore_FindHeader(oSrc, sPlaceHeader)
    lastCol = WMSCore_LastHeaderCol(oSrc)
    last = WMSCore_LastContentRow(oSrc, lastCol)
    stamp = WMSCore_Stamp(Now)

    oldLast = WMSCore_LastContentRow(oIdx, 10)
    keepN = 0
    If oldLast >= 1 Then
        aOld = oIdx.getCellRangeByPosition(0,1,10,oldLast).getDataArray()
        For i = 0 To UBound(aOld)
            rowData = aOld(i)
            typ = WMSCore_Canon(CStr(rowData(0))): src = WMSCore_Canon(CStr(rowData(8)))
            If WMSCore_CleanText(rowData(0)) <> "" Then
                If Not (typ = WMSCore_Canon(sIndexType) And src = WMSCore_Canon(sSourceSheet)) Then keepN = keepN + 1
            End If
        Next i
    End If

    newN = 0
    If last >= 1 Then
        aSrc = oSrc.getCellRangeByPosition(0,1,lastCol,last).getDataArray()
        For i = 0 To UBound(aSrc)
            rowData = aSrc(i)
            key = WMSCore_CleanText(WMSCore_ArrayValue(rowData, cKey))
            If key <> "" Then newN = newN + 1
        Next i
    End If

    totalN = keepN + newN
    If totalN > 0 Then ReDim aOut(0 To totalN-1)
    outN = 0

    If keepN > 0 Then
        For i = 0 To UBound(aOld)
            rowData = aOld(i)
            typ = WMSCore_Canon(CStr(rowData(0))): src = WMSCore_Canon(CStr(rowData(8)))
            If WMSCore_CleanText(rowData(0)) <> "" Then
                If Not (typ = WMSCore_Canon(sIndexType) And src = WMSCore_Canon(sSourceSheet)) Then
                    aOut(outN) = rowData: outN = outN + 1
                End If
            End If
        Next i
    End If

    If newN > 0 Then
        For i = 0 To UBound(aSrc)
            rowData = aSrc(i)
            key = WMSCore_CleanText(WMSCore_ArrayValue(rowData, cKey))
            If key <> "" Then
                aOut(outN) = Array(sIndexType, key, WMSCore_ArrayValue(rowData,cItem), WMSCore_ArrayValue(rowData,cCode), WMSCore_ArrayValue(rowData,cArt), WMSCore_ArrayValue(rowData,cName), WMSCore_ArrayValue(rowData,cUnit), WMSCore_ArrayValue(rowData,cPlace), sSourceSheet, CDbl(i+2), stamp)
                outN = outN + 1
                WMSCore_BuildIndexCore = WMSCore_BuildIndexCore + 1
            End If
        Next i
    End If

    bBatch = WMSCore_BeginBatch(oDoc)
    If oldLast >= 1 Then oIdx.getCellRangeByPosition(0,1,10,oldLast).clearContents(23)
    If totalN > 0 Then oIdx.getCellRangeByPosition(0,1,10,totalN).setDataArray(aOut())
    oIdx.IsVisible = False
    If bBatch Then WMSCore_EndBatch oDoc, False
    Exit Function
EH:
    On Error Resume Next
    If bBatch Then WMSCore_EndBatch oDoc, False
    WMSCore_LogError oDoc, WMSC_ERR_RUNTIME, WMSC_MODULE, "BuildIndex", "", "INDEX_BUILD", CStr(Err) & " " & Error$
End Function

Function WMSCore_ArrayValue(aRow As Variant, c As Long) As Variant
    WMSCore_ArrayValue = ""
    On Error GoTo Done
    If c >= 0 And c <= UBound(aRow) Then WMSCore_ArrayValue = aRow(c)
Done:
End Function

' ============================================================================
' AUDIT / ERROR TAXONOMY
' ============================================================================

Sub WMSCore_LogAudit(oDoc As Object, sModule As String, sEvent As String, sSourceID As String, sDescription As String)
    Dim oSh As Object, r As Long, c As Long
    On Error GoTo Done
    If Not oDoc.Sheets.hasByName(WMSC_AUDIT) Then Exit Sub
    oSh = oDoc.Sheets.getByName(WMSC_AUDIT)
    r = WMSCore_NextRow(oSh, 0, 1)
    oSh.getCellRangeByPosition(0,r,4,r).setDataArray(Array(Array(WMSCore_Stamp(Now), gWMSC_LastRunID, sModule, sEvent, sDescription)))
    c = WMSCore_FindHeader(oSh, "_WMS_Category"): If c >= 0 Then oSh.getCellByPosition(c,r).String = "AUDIT"
    c = WMSCore_FindHeader(oSh, "_WMS_SourceID"): If c >= 0 Then oSh.getCellByPosition(c,r).String = sSourceID
    c = WMSCore_FindHeader(oSh, "_WMS_Module"): If c >= 0 Then oSh.getCellByPosition(c,r).String = sModule
    c = WMSCore_FindHeader(oSh, "_WMS_LastTouch"): If c >= 0 Then oSh.getCellByPosition(c,r).String = WMSCore_Stamp(Now)
Done:
End Sub

Sub WMSCore_LogError(oDoc As Object, sCategory As String, sModule As String, sWhere As String, sSourceID As String, sCode As String, sText As String)
    Dim oSh As Object, r As Long, c As Long, cat As String, shown As String
    On Error GoTo Done
    If Not oDoc.Sheets.hasByName(WMSC_ERRORS) Then Exit Sub
    cat = UCase(Trim(sCategory)): If cat = "" Then cat = WMSC_ERR_RUNTIME
    shown = "[" & cat
    If Trim(sCode) <> "" Then shown = shown & "/" & Trim(sCode)
    shown = shown & "] " & sText
    oSh = oDoc.Sheets.getByName(WMSC_ERRORS)
    r = WMSCore_NextRow(oSh, 0, 1)
    oSh.getCellRangeByPosition(0,r,4,r).setDataArray(Array(Array(WMSCore_Stamp(Now), gWMSC_LastRunID, sModule & "/" & sWhere, shown, "OPEN")))
    c = WMSCore_FindHeader(oSh, "_WMS_Category"): If c >= 0 Then oSh.getCellByPosition(c,r).String = cat
    c = WMSCore_FindHeader(oSh, "_WMS_Module"): If c >= 0 Then oSh.getCellByPosition(c,r).String = sModule
    c = WMSCore_FindHeader(oSh, "_WMS_SourceID"): If c >= 0 Then oSh.getCellByPosition(c,r).String = sSourceID
    c = WMSCore_FindHeader(oSh, "_WMS_ErrorCode"): If c >= 0 Then oSh.getCellByPosition(c,r).String = sCode
    c = WMSCore_FindHeader(oSh, "_WMS_LastTouch"): If c >= 0 Then oSh.getCellByPosition(c,r).String = WMSCore_Stamp(Now)
Done:
End Sub

' ============================================================================
' IDS / TEXT / HASH / DATE / NUMBER UTILITIES
' ============================================================================

Function WMSCore_NewSourceID() As String
    WMSCore_NewSourceID = WMSCore_NewID("SRC")
End Function

Function WMSCore_NewMovementID() As String
    WMSCore_NewMovementID = WMSCore_NewID("MOV")
End Function

Function WMSCore_NewRunID() As String
    WMSCore_NewRunID = WMSCore_NewID("RUN")
End Function

Function WMSCore_NewID(sPrefix As String) As String
    If gWMSC_IDSession = "" Then
        Randomize
        gWMSC_IDSession = Right("000000" & CStr(Int(Rnd()*1000000)),6)
    End If
    gWMSC_IDCounter = gWMSC_IDCounter + 1
    If gWMSC_IDCounter > 999999 Then gWMSC_IDCounter = 1
    WMSCore_NewID = UCase(Trim(sPrefix)) & "-" & WMSCore_TimestampCompact(Now) & "-" & gWMSC_IDSession & "-" & Right("000000" & CStr(gWMSC_IDCounter),6)
End Function

Function WMSCore_Stamp(v As Variant) As String
    On Error GoTo Fallback
    WMSCore_Stamp = Right("0" & CStr(Day(v)),2) & "." & Right("0" & CStr(Month(v)),2) & "." & CStr(Year(v)) & " " & Right("0" & CStr(Hour(v)),2) & ":" & Right("0" & CStr(Minute(v)),2) & ":" & Right("0" & CStr(Second(v)),2)
    Exit Function
Fallback:
    WMSCore_Stamp = CStr(v)
End Function

Function WMSCore_TimestampCompact(v As Variant) As String
    WMSCore_TimestampCompact = CStr(Year(v)) & Right("0" & CStr(Month(v)),2) & Right("0" & CStr(Day(v)),2) & Right("0" & CStr(Hour(v)),2) & Right("0" & CStr(Minute(v)),2) & Right("0" & CStr(Second(v)),2)
End Function

Function WMSCore_CleanText(v As Variant) As String
    Dim s As String
    s = CStr(v)
    s = Replace(s, Chr(160), " ")
    s = Replace(s, Chr(9), " ")
    s = Replace(s, Chr(13), " ")
    s = Replace(s, Chr(10), " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    WMSCore_CleanText = Trim(s)
End Function

Function WMSCore_Canon(v As Variant) As String
    WMSCore_Canon = UCase(WMSCore_CleanText(v))
End Function

Function WMSCore_CanonHeader(v As Variant) As String
    Dim s As String
    s = WMSCore_Canon(v)
    s = Replace(s, "Ё", "Е")
    WMSCore_CanonHeader = s
End Function

Function WMSCore_CanonKey(v As Variant) As String
    Dim s As String
    s = WMSCore_Canon(v)
    s = Replace(s, " ", "_")
    s = Replace(s, "/", "_")
    s = Replace(s, "\\", "_")
    s = Replace(s, ".", "_")
    WMSCore_CanonKey = s
End Function

Function WMSCore_TryNumber(v As Variant, ByRef nOut As Double) As Boolean
    Dim s As String
    WMSCore_TryNumber = False: nOut = 0
    On Error GoTo Done
    If IsNumeric(v) Then nOut = CDbl(v): WMSCore_TryNumber = True: Exit Function
    s = WMSCore_CleanText(v)
    If s = "" Then Exit Function
    s = Replace(s, Chr(160), "")
    s = Replace(s, " ", "")
    If InStr(s, ",") > 0 And InStr(s, ".") = 0 Then s = Replace(s, ",", ".")
    If IsNumeric(s) Then nOut = CDbl(s): WMSCore_TryNumber = True
Done:
End Function

Function WMSCore_ParseWorkingDateText(sIn As String, ByRef dOut As Date) As Boolean
    Dim s As String, a As Variant, dd As Long, mm As Long, yy As Long
    WMSCore_ParseWorkingDateText = False
    s = WMSCore_CleanText(sIn): If s = "" Then Exit Function
    s = Replace(s, "/", "."): s = Replace(s, "-", ".")
    a = Split(s, ".")
    On Error GoTo Done
    If UBound(a) = 1 Then
        dd = CLng(a(0)): mm = CLng(a(1)): yy = Year(Date)
    ElseIf UBound(a) = 2 Then
        dd = CLng(a(0)): mm = CLng(a(1)): yy = CLng(a(2))
        If yy < 100 Then yy = 2000 + yy
    Else
        Exit Function
    End If
    dOut = DateSerial(yy,mm,dd)
    If Day(dOut) <> dd Or Month(dOut) <> mm Or Year(dOut) <> yy Then Exit Function
    WMSCore_ParseWorkingDateText = True
Done:
End Function

Function WMSCore_NumText(n As Double) As String
    Dim s As String
    s = CStr(n)
    WMSCore_NumText = Replace(s, ",", ".")
End Function

Function WMSCore_SimpleHash(s As String) As Long
    Dim i As Long, h As Double, code As Long
    h = 2166136261#
    For i = 1 To Len(s)
        code = Asc(Mid(s,i,1)): If code < 0 Then code = code + 65536
        h = h + code * 16777619#
        h = h - Int(h / 2147483629#) * 2147483629#
    Next i
    If h < 0 Then h = -h
    WMSCore_SimpleHash = CLng(h Mod 2147483000#)
End Function

Function WMSCore_HashSlotStart(sKey As String, nCap As Long) As Long
    If nCap <= 0 Then WMSCore_HashSlotStart = 0 Else WMSCore_HashSlotStart = WMSCore_SimpleHash(sKey) Mod nCap
End Function

Function WMSCore_NextPow2(n As Long) As Long
    Dim p As Long: p = 1
    If n < 1 Then WMSCore_NextPow2 = 1: Exit Function
    Do While p < n And p < 1073741824
        p = p * 2
    Loop
    WMSCore_NextPow2 = p
End Function

' ============================================================================
' SELF-CHECK / DIAGNOSTICS
' ============================================================================

Function WMSCore_RunSelfCheck(oDoc As Object, ByRef sReport As String) As Boolean
    Dim issues As String, oSh As Object, c As Long, k As Long, s As String
    Dim id1 As String, id2 As String, oMap As Object
    WMSCore_RunSelfCheck = False
    issues = ""
    On Error GoTo EH

    If Not WMSCore_IsCalcDocument(oDoc) Then WMSCore_AppendIssue issues, "не Calc-документ": GoTo Done
    If Not oDoc.Sheets.hasByName(WMSC_SETTINGS) Then WMSCore_AppendIssue issues, "нет Настройки WMS"
    If Not oDoc.Sheets.hasByName(WMSC_AUDIT) Then WMSCore_AppendIssue issues, "нет Журнал WMS"
    If Not oDoc.Sheets.hasByName(WMSC_ERRORS) Then WMSCore_AppendIssue issues, "нет Ошибки WMS"
    If Not oDoc.Sheets.hasByName(WMSC_QUEUE) Then WMSCore_AppendIssue issues, "нет SYS_WMS_QUEUE"
    If Not oDoc.Sheets.hasByName(WMSC_INDEX) Then WMSCore_AppendIssue issues, "нет SYS_WMS_INDEX"
    If issues <> "" Then GoTo Done

    oSh = oDoc.Sheets.getByName(WMSC_SETTINGS)
    oMap = WMSCore_GetHeaderMap(oSh)
    If WMSCore_HeaderIndex(oMap, "Параметр") < 0 Then WMSCore_AppendIssue issues, "header-map не видит Параметр"
    If WMSCore_FindHeader(oSh, "Значение") < 0 Then WMSCore_AppendIssue issues, "header-cache не видит Значение"

    s = WMSCore_GetSetting(oDoc, "WMS.CommonCore.Version", "")
    If s <> WMSC_VERSION Then WMSCore_AppendIssue issues, "версия Common Core в настройках не совпадает"
    If WMSCore_GetSetting(oDoc, "WMS.Batch.ChunkSize", "") = "" Then WMSCore_AppendIssue issues, "нет WMS.Batch.ChunkSize"

    oSh = oDoc.Sheets.getByName(WMSC_ERRORS)
    If WMSCore_FindHeader(oSh, "_WMS_Category") < 0 Then WMSCore_AppendIssue issues, "нет taxonomy поля ошибок"
    If WMSCore_FindHeader(oSh, "_WMS_ErrorCode") < 0 Then WMSCore_AppendIssue issues, "нет кода ошибки"

    oSh = oDoc.Sheets.getByName(WMSC_QUEUE)
    For c = 0 To 7
        If Trim(oSh.getCellByPosition(c,0).String) = "" Then WMSCore_AppendIssue issues, "неполная схема queue": Exit For
    Next c
    WMSCore_ResetQueueCache
    If Not WMSCore_EnsureQueueCache(oDoc) Then WMSCore_AppendIssue issues, "queue cache не строится"

    k = WMSCore_NumberFormatKey(oDoc, "DATE"): If k < 0 Then WMSCore_AppendIssue issues, "не создаётся формат даты"
    k = WMSCore_NumberFormatKey(oDoc, "MONEY"): If k < 0 Then WMSCore_AppendIssue issues, "не создаётся денежный формат"

    id1 = WMSCore_NewSourceID(): id2 = WMSCore_NewSourceID()
    If id1 = "" Or id2 = "" Or id1 = id2 Then WMSCore_AppendIssue issues, "генератор SourceID не уникален"

Done:
    If issues = "" Then
        sReport = "WMS Common Core " & WMSC_VERSION & " self-check: OK"
        WMSCore_RunSelfCheck = True
    Else
        sReport = "WMS Common Core " & WMSC_VERSION & " self-check: FAIL — " & issues
    End If
    Exit Function
EH:
    WMSCore_AppendIssue issues, "runtime " & CStr(Err) & " " & Error$
    Resume Done
End Function

Sub WMSCore_AppendIssue(ByRef s As String, sIssue As String)
    If Trim(sIssue) = "" Then Exit Sub
    If s <> "" Then s = s & "; "
    s = s & sIssue
End Sub

Function WMSCore_IsCalcDocument(oDoc As Object) As Boolean
    WMSCore_IsCalcDocument = False
    On Error GoTo Done
    WMSCore_IsCalcDocument = oDoc.supportsService("com.sun.star.sheet.SpreadsheetDocument")
Done:
End Function


' PO KATAK OPT-1: exclude formatting-only tail without changing cell semantics.
' queryContentCells includes values, dates, text and formulas (1+2+4+16).
' On an unsupported query API retain the original bounded batch scan.
Function WMSCore_ContentEndBound(oSheet As Object, nLastCol As Long, originalEnd As Long) As Long
    Dim oMatches As Object, aRanges As Variant, rangeItem As Variant
    Dim i As Long, contentEnd As Long
    WMSCore_ContentEndBound = originalEnd
    On Error GoTo KeepOriginal
    If originalEnd < 1 Or nLastCol < 0 Then Exit Function
    oMatches = oSheet.getCellRangeByPosition(0,1,nLastCol,originalEnd).queryContentCells(23)
    aRanges = oMatches.getRangeAddresses()
    contentEnd = 0
    For i = LBound(aRanges) To UBound(aRanges)
        rangeItem = aRanges(i)
        If rangeItem.EndRow > contentEnd Then contentEnd = rangeItem.EndRow
    Next i
    WMSCore_ContentEndBound = contentEnd
KeepOriginal:
End Function
