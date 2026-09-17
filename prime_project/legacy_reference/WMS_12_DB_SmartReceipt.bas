Option Explicit

' ============================================================================
' WMS_12_DB_SmartReceipt
' SMART RECEIPT / LOTS INTEGRATION
' LibreOffice Calc + embedded Firebird
'
' Version 0.1.0
'
' Purpose:
' - connect Orders receipts to WMS_11 Units/Lots;
' - preserve quantity/unit exactly as entered from the document;
' - post physical stock in the BASE quantity/unit;
' - create a deterministic receipt lot per SourceID;
' - remember the conversion for that exact lot forever;
' - add a light "Единицы партии" button to Orders.
'
' Examples:
'   1 упак in document -> 1000 шт in physical stock
'   10 шт metal in document -> 100 кг in physical stock
'
' IMPORTANT:
' - no heavy OnChange handler is installed;
' - ordinary 1:1 receipts require no extra dialog;
' - if product already has another BASE unit, conduct asks to confirm
'   the physical quantity;
' - for the FIRST unusual receipt, select the row and press
'   "Единицы партии" before conducting.
' ============================================================================

Global Const WMSDBSR_VERSION = "0.2.1-PERSIST-SUPPLIER-ARTICLE"
Global Const WMSDBSR_SHEET = "Заказы"
Global Const WMSDBSR_FORM = "WMS_SMART_RECEIPT_PANEL"

Global Const WMSDBSR_T_LOT = "_WMS_LotID"
Global Const WMSDBSR_T_DOCQTY = "_WMS_DocQty"
Global Const WMSDBSR_T_DOCUNIT = "_WMS_DocUnit"
Global Const WMSDBSR_T_BASEQTY = "_WMS_BaseQty"
Global Const WMSDBSR_T_BASEUNIT = "_WMS_BaseUnit"
Global Const WMSDBSR_T_FACTOR = "_WMS_UnitFactor"
Global Const WMSDBSR_T_CONFIG = "_WMS_UnitsConfigured"
Global Const WMSDBSR_T_ACTQTY = "_WMS_ActualQty"
Global Const WMSDBSR_T_ACTUNIT = "_WMS_ActualUnit"
Global Const WMSDBSR_T_ACTFACTOR = "_WMS_ActualFactorToBase"

Sub WMSDBSR_Install()
    Dim oDoc As Object,oSh As Object,sErr As String
    oDoc=ThisComponent
    On Error GoTo EH

    If Not oDoc.Sheets.hasByName(WMSDBSR_SHEET) Then
        MsgBox "Лист 'Заказы' не найден.",16,"WMS — Умный приход"
        Exit Sub
    End If
    oSh=oDoc.Sheets.getByName(WMSDBSR_SHEET)

    ' First verify that the database foundation exists.
    If Not WMSDBSR_CheckFoundation(sErr) Then
        MsgBox "Сначала должен быть установлен WMS_11 Units & Lots." & Chr(10) & Chr(10) & sErr, _
               16,"WMS — Умный приход"
        Exit Sub
    End If

    WMSDBSR_EnsureTechColumns oSh
    WMSDBSR_InstallButton oDoc,oSh

    MsgBox "Умный приход установлен." & Chr(10) & _
           "Версия: " & WMSDBSR_VERSION & Chr(10) & Chr(10) & _
           "Обычные приходы 1:1 работают без дополнительных окон." & Chr(10) & _
           "Для необычной первой партии используйте кнопку 'Единицы партии'.", _
           64,"WMS — Умный приход"
    Exit Sub
EH:
    MsgBox "Установка Smart Receipt: " & CStr(Err) & " " & Error$,16,"WMS — Умный приход"
End Sub

Sub WMSDBSR_SelfCheck()
    Dim oDoc As Object,oSh As Object,oCon As Object,sErr As String,s As String
    Dim ok As Boolean
    oDoc=ThisComponent
    ok=True
    s=""

    On Error GoTo EH

    If Not oDoc.Sheets.hasByName(WMSDBSR_SHEET) Then
        s=s & "Лист Заказы: НЕТ" & Chr(10)
        ok=False
    Else
        oSh=oDoc.Sheets.getByName(WMSDBSR_SHEET)
        If WMSDBSR_FindHeader(oSh,WMSDBSR_T_LOT)<0 Then ok=False:s=s & WMSDBSR_T_LOT & ": НЕТ" & Chr(10)
        If WMSDBSR_FindHeader(oSh,WMSDBSR_T_DOCQTY)<0 Then ok=False:s=s & WMSDBSR_T_DOCQTY & ": НЕТ" & Chr(10)
        If WMSDBSR_FindHeader(oSh,WMSDBSR_T_DOCUNIT)<0 Then ok=False:s=s & WMSDBSR_T_DOCUNIT & ": НЕТ" & Chr(10)
        If WMSDBSR_FindHeader(oSh,WMSDBSR_T_BASEQTY)<0 Then ok=False:s=s & WMSDBSR_T_BASEQTY & ": НЕТ" & Chr(10)
        If WMSDBSR_FindHeader(oSh,WMSDBSR_T_BASEUNIT)<0 Then ok=False:s=s & WMSDBSR_T_BASEUNIT & ": НЕТ" & Chr(10)
        If WMSDBSR_FindHeader(oSh,WMSDBSR_T_FACTOR)<0 Then ok=False:s=s & WMSDBSR_T_FACTOR & ": НЕТ" & Chr(10)
        If WMSDBSR_FindHeader(oSh,WMSDBSR_T_CONFIG)<0 Then ok=False:s=s & WMSDBSR_T_CONFIG & ": НЕТ" & Chr(10)
        If WMSDBSR_FindHeader(oSh,WMSDBSR_T_ACTQTY)<0 Then ok=False:s=s & WMSDBSR_T_ACTQTY & ": НЕТ" & Chr(10)
        If WMSDBSR_FindHeader(oSh,WMSDBSR_T_ACTUNIT)<0 Then ok=False:s=s & WMSDBSR_T_ACTUNIT & ": НЕТ" & Chr(10)
        If WMSDBSR_FindHeader(oSh,WMSDBSR_T_ACTFACTOR)<0 Then ok=False:s=s & WMSDBSR_T_ACTFACTOR & ": НЕТ" & Chr(10)
    End If

    If Not WMSDBSR_CheckFoundation(sErr) Then
        ok=False
        s=s & "Units/Lots foundation: ОШИБКА — " & sErr & Chr(10)
    Else
        s=s & "Units/Lots foundation: OK" & Chr(10)
    End If

    If ok Then
        MsgBox "Smart Receipt SelfCheck: OK" & Chr(10) & _
               "Версия: " & WMSDBSR_VERSION & Chr(10) & _
               "Техполя Заказов: OK" & Chr(10) & _
               "Тяжёлый OnChange: НЕ используется",64,"WMS — Умный приход"
    Else
        MsgBox "Smart Receipt SelfCheck: ОШИБКА" & Chr(10) & Chr(10) & s,16,"WMS — Умный приход"
    End If
    Exit Sub
EH:
    MsgBox "SelfCheck Smart Receipt: " & CStr(Err) & " " & Error$,16,"WMS — Умный приход"
End Sub

Function WMSDBSR_CheckFoundation(ByRef sErr As String) As Boolean
    Dim oCon As Object,rep As String
    WMSDBSR_CheckFoundation=False
    sErr=""
    On Error GoTo EH

    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then Exit Function

    If Not WMSDBUL_SelfCheckCore(oCon,rep) Then
        sErr=rep
        Exit Function
    End If

    WMSDBSR_CheckFoundation=True
    Exit Function
EH:
    sErr="Проверка WMS_11: " & CStr(Err) & " " & Error$
End Function

' ============================================================================
' USER ACTION: configure physical/base quantity for selected receipt row.
' ============================================================================

Sub WMSDBSR_ConfigureSelected()
    Dim oDoc As Object,oSh As Object,oSel As Object,r As Long
    Dim code As String,nm As String,docUnit As String,actualUnit As String,baseUnit As String
    Dim sBaseQty As String,sErr As String,sCurrentBase As String
    Dim dDocQty As Double,dActualQty As Double,dBaseQty As Double,dDocFactor As Double,dActualFactor As Double
    Dim answer As Integer

    oDoc=ThisComponent
    On Error GoTo EH

    If Not oDoc.Sheets.hasByName(WMSDBSR_SHEET) Then
        MsgBox "Лист 'Заказы' не найден.",16,"WMS — Единицы партии"
        Exit Sub
    End If
    oSh=oDoc.CurrentController.ActiveSheet
    If oSh.Name<>WMSDBSR_SHEET Then
        MsgBox "Откройте лист 'Заказы' и выберите нужную позицию.",48,"WMS — Единицы партии"
        Exit Sub
    End If

    oSel=oDoc.CurrentSelection:r=oSel.CellAddress.Row
    If r<1 Then MsgBox "Выберите строку товара.",48,"WMS — Единицы партии":Exit Sub

    code=WMSDBSR_CellText(oSh,r,"Код товара")
    nm=WMSDBSR_CellText(oSh,r,"Полное наименование товара")

    ' H/I = документ. G = фактически принятое количество.
    dDocQty=oSh.getCellByPosition(7,r).Value
    dActualQty=oSh.getCellByPosition(6,r).Value
    docUnit=Trim(oSh.getCellByPosition(8,r).String)

    If dDocQty<=0 Then dDocQty=dActualQty
    If code="" Then MsgBox "Сначала заполните Код товара.",48,"WMS — Единицы партии":Exit Sub
    If dActualQty<=0 Then MsgBox "Сначала заполните Факт. количество.",48,"WMS — Единицы партии":Exit Sub
    If dDocQty<=0 Then MsgBox "Не найдено документальное количество.",48,"WMS — Единицы партии":Exit Sub
    If docUnit="" Then MsgBox "Заполните документальную единицу в 'Ед. изм.'.",48,"WMS — Единицы партии":Exit Sub

    actualUnit=InputBox( _
        "Товар: " & nm & Chr(10) & _
        "Код: " & code & Chr(10) & Chr(10) & _
        "По документу: " & WMSDBSR_NumText(dDocQty) & " " & docUnit & Chr(10) & _
        "Фактически принято: " & WMSDBSR_NumText(dActualQty) & Chr(10) & Chr(10) & _
        "В какой единице вы считаете ФАКТ?" & Chr(10) & _
        "Например: шт, рул, упак." & Chr(10) & _
        "Документальную единицу в I менять не нужно.", _
        "WMS — Фактическая единица",docUnit)
    actualUnit=Trim(actualUnit)
    If actualUnit="" Then Exit Sub

    sCurrentBase=WMSDBSR_GetProductBaseUnit(code,sErr)
    If sErr<>"" Then MsgBox sErr,16,"WMS — Единицы партии":Exit Sub
    If sCurrentBase="" Then sCurrentBase=docUnit

    baseUnit=InputBox( _
        "В какой единице товар должен отображаться в ОСТАТКЕ?" & Chr(10) & Chr(10) & _
        "Документ: " & WMSDBSR_NumText(dDocQty) & " " & docUnit & Chr(10) & _
        "Факт: " & WMSDBSR_NumText(dActualQty) & " " & actualUnit, _
        "WMS — Базовая единица остатка",sCurrentBase)
    baseUnit=Trim(baseUnit)
    If baseUnit="" Then Exit Sub

    sBaseQty=""
    If LCase(baseUnit)=LCase(docUnit) Then sBaseQty=WMSDBSR_NumText(dDocQty)

    sBaseQty=InputBox( _
        "Сколько '" & baseUnit & "' фактически должно попасть в остаток?" & Chr(10) & Chr(10) & _
        "Пример: документ 100 кг, фактически 10 шт, остаток 100 кг.", _
        "WMS — Количество в остатке",sBaseQty)
    If Trim(sBaseQty)="" Then Exit Sub
    If Not WMSDBSR_ParseNumber(sBaseQty,dBaseQty) Or dBaseQty<=0 Then
        MsgBox "Некорректное базовое количество.",48,"WMS — Единицы партии":Exit Sub
    End If

    dDocFactor=dBaseQty/dDocQty
    dActualFactor=dBaseQty/dActualQty

    answer=MsgBox( _
        "Проверьте:" & Chr(10) & Chr(10) & _
        "Документ: " & WMSDBSR_NumText(dDocQty) & " " & docUnit & Chr(10) & _
        "Факт: " & WMSDBSR_NumText(dActualQty) & " " & actualUnit & Chr(10) & _
        "Остаток: " & WMSDBSR_NumText(dBaseQty) & " " & baseUnit & Chr(10) & Chr(10) & _
        "1 " & actualUnit & " = " & WMSDBSR_NumText(dActualFactor) & " " & baseUnit, _
        36,"WMS — Единицы партии")
    If answer<>6 Then Exit Sub

    WMSDBSR_EnsureTechColumns oSh
    WMSDBSR_SaveRowConfig oSh,r,dDocQty,docUnit,dBaseQty,baseUnit,dDocFactor
    WMSDBSR_SetTechNumber oSh,r,WMSDBSR_T_ACTQTY,dActualQty
    WMSDBSR_SetTechText oSh,r,WMSDBSR_T_ACTUNIT,actualUnit
    WMSDBSR_SetTechNumber oSh,r,WMSDBSR_T_ACTFACTOR,dActualFactor

    MsgBox "Партия настроена." & Chr(10) & _
           "Документ: " & WMSDBSR_NumText(dDocQty) & " " & docUnit & Chr(10) & _
           "Факт: " & WMSDBSR_NumText(dActualQty) & " " & actualUnit & Chr(10) & _
           "Остаток: " & WMSDBSR_NumText(dBaseQty) & " " & baseUnit,64,"WMS — Единицы партии"
    Exit Sub
EH:
    MsgBox "Настройка единиц: " & CStr(Err) & " " & Error$,16,"WMS — Единицы партии"
End Sub

Sub WMSDBSR_ResetSelected()
    Dim oDoc As Object,oSh As Object,r As Long,oSel As Object
    oDoc=ThisComponent
    On Error GoTo EH
    oSh=oDoc.CurrentController.ActiveSheet
    If oSh.Name<>WMSDBSR_SHEET Then Exit Sub
    oSel=oDoc.CurrentSelection:r=oSel.CellAddress.Row
    If r<1 Then Exit Sub
    WMSDBSR_ClearRowConfig oSh,r
    MsgBox "Настройка единиц выбранной позиции сброшена." & Chr(10) & _
           "При следующем проведении WMS определит её заново.",64,"WMS — Единицы партии"
    Exit Sub
EH:
    MsgBox "Сброс единиц: " & CStr(Err) & " " & Error$,16,"WMS — Единицы партии"
End Sub

' ============================================================================
' API CALLED BY WMS_06 0.5.0
' ============================================================================

Function WMSDBSR_PostReceiptCon(oCon As Object,oSh As Object,r As Long, _
    sMovementID As String,sSourceID As String,sSourceType As String, _
    sProductCode As String,sProductName As String,dDocQty As Double,sDocUnit As String, _
    sLocation As String,sOrigin As String,sDestCategory As String,sDestSubcategory As String, _
    dMovementDate As Double,sNote As String,ByRef sErr As String) As Boolean

    Dim dBaseQty As Double,dFactor As Double
    Dim sBaseUnit As String,sLotID As String
    Dim oStmt As Object,sql As String

    WMSDBSR_PostReceiptCon=False
    sErr=""
    On Error GoTo EH

    If dDocQty<=0 Then sErr="Документальное количество должно быть > 0.":Exit Function
    If Trim(sDocUnit)="" Then sErr="Документальная единица не заполнена.":Exit Function

    If Not WMSDBSR_ResolveRowUnits(oCon,oSh,r,sProductCode,dDocQty,sDocUnit, _
        dBaseQty,sBaseUnit,dFactor,sErr) Then Exit Function

    Dim dStoredDocQty As Double,sStoredDocUnit As String,dActualQty As Double,sActualUnit As String,dActualFactor As Double
    dStoredDocQty=WMSDBSR_TechNumber(oSh,r,WMSDBSR_T_DOCQTY)
    sStoredDocUnit=WMSDBSR_TechText(oSh,r,WMSDBSR_T_DOCUNIT)
    dActualQty=WMSDBSR_TechNumber(oSh,r,WMSDBSR_T_ACTQTY)
    sActualUnit=WMSDBSR_TechText(oSh,r,WMSDBSR_T_ACTUNIT)
    dActualFactor=WMSDBSR_TechNumber(oSh,r,WMSDBSR_T_ACTFACTOR)
    If dStoredDocQty<=0 Then dStoredDocQty=dDocQty
    If sStoredDocUnit="" Then sStoredDocUnit=sDocUnit
    If dActualQty<=0 Then dActualQty=dDocQty
    If sActualUnit="" Then sActualUnit=sDocUnit

    sLotID=WMSDBSR_MakeLotID(sSourceID)
    WMSDBSR_SetTechText oSh,r,WMSDBSR_T_LOT,sLotID

    ' Stock ledger remains backward compatible:
    ' QTY + UNIT_NAME are the BASE physical movement.
    If Not WMSDBST_PostMovementClassifiedCon(oCon,sMovementID,sSourceID,sSourceType,"IN", _
        sProductCode,sProductName,dBaseQty,sBaseUnit,sLocation,sOrigin, _
        sDestCategory,sDestSubcategory,dMovementDate,sNote,sErr) Then
        If sErr="" Then sErr="WMS_10 не подтвердил базовое движение."
        Exit Function
    End If

    ' Preserve supplier article from column F in the product master.
    ' This makes the next receipt resolvable by article.
    Dim sSupplierArticle As String
    sSupplierArticle=Trim(oSh.getCellByPosition(5,r).String)
    If sSupplierArticle<>"" Then
        If Not WMSDBST_EnsureProductCon(oCon,sProductCode,sProductName,sBaseUnit,sLocation,sSupplierArticle,sErr) Then
            If sErr="" Then sErr="Не удалось сохранить артикул в номенклатуре."
            Exit Function
        End If
    End If

    ' If a previous partial attempt already created the lot, history is immutable:
    ' verify it instead of silently changing its conversion.
    If WMSDBUL_LotExists(oCon,sLotID,sErr) Then
        If sErr<>"" Then Exit Function
        If Not WMSDBSR_LotMatchesCon(oCon,sLotID,sProductCode,dStoredDocQty,sStoredDocUnit, _
            dBaseQty,sBaseUnit,sErr) Then Exit Function
    Else
        If sErr<>"" Then Exit Function
        If Not WMSDBUL_CreateLotCon(oCon,sLotID,sProductCode,sSourceID,sSourceType, _
            dStoredDocQty,sStoredDocUnit,dBaseQty,sBaseUnit,sLocation,sOrigin,sDestCategory, _
            sDestSubcategory,dMovementDate,sNote,sErr) Then Exit Function

        ' Base unit already gets factor 1 inside WMS_11.
        If LCase(Trim(sStoredDocUnit))<>LCase(Trim(sBaseUnit)) Then
            If Not WMSDBUL_SetLotUnitCon(oCon,sLotID,sStoredDocUnit,dBaseQty/dStoredDocQty,"DOC",sErr) Then Exit Function
        End If
        If Trim(sActualUnit)<>"" And LCase(Trim(sActualUnit))<>LCase(Trim(sBaseUnit)) Then
            If dActualFactor<=0 Then dActualFactor=dBaseQty/dActualQty
            If Not WMSDBUL_SetLotUnitCon(oCon,sLotID,sActualUnit,dActualFactor,"ACTUAL",sErr) Then Exit Function
        End If
    End If

    ' Attach the immutable movement to the immutable lot and preserve
    ' what the user/document said.
    sql="UPDATE WMS_STOCK_MOVEMENTS SET " & _
        "LOT_ID=" & WMSDB_SQLText(sLotID) & "," & _
        "ENTRY_QTY=" & WMSDBSR_SQLNumber(dStoredDocQty) & "," & _
        "ENTRY_UNIT=" & WMSDB_SQLText(sStoredDocUnit) & "," & _
        "BASE_QTY=" & WMSDBSR_SQLNumber(dBaseQty) & "," & _
        "BASE_UNIT=" & WMSDB_SQLText(sBaseUnit) & _
        " WHERE MOVEMENT_ID=" & WMSDB_SQLText(sMovementID)

    oStmt=oCon.createStatement()
    oStmt.executeUpdate(sql)
    oCon.commit()

    If Not WMSDBSR_VerifyReceiptCon(oCon,sMovementID,sSourceID,sErr) Then Exit Function

    WMSDBSR_PostReceiptCon=True
    Exit Function
EH:
    sErr="Smart Receipt: " & CStr(Err) & " " & Error$
    On Error Resume Next
    oCon.rollback()
End Function

Function WMSDBSR_VerifyReceiptCon(oCon As Object,sMovementID As String,sSourceID As String, _
    ByRef sErr As String) As Boolean

    Dim oStmt As Object,oRS As Object,sql As String,sLot As String,n As Long
    WMSDBSR_VerifyReceiptCon=False
    sErr=""
    On Error GoTo EH

    sql="SELECT LOT_ID,ENTRY_QTY,ENTRY_UNIT,BASE_QTY,BASE_UNIT " & _
        "FROM WMS_STOCK_MOVEMENTS WHERE MOVEMENT_ID=" & WMSDB_SQLText(sMovementID)
    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery(sql)

    If Not oRS.next() Then
        sErr="Движение Smart Receipt не найдено: " & sMovementID
        Exit Function
    End If

    sLot=Trim(oRS.getString(1))
    If sLot="" Then
        sErr="У движения не записан LOT_ID: " & sMovementID
        Exit Function
    End If
    If Trim(oRS.getString(3))="" Or Trim(oRS.getString(5))="" Then
        sErr="У движения не сохранены единицы ENTRY/BASE."
        Exit Function
    End If

    oRS=oStmt.executeQuery("SELECT COUNT(*) FROM WMS_STOCK_LOTS WHERE LOT_ID=" & WMSDB_SQLText(sLot))
    If Not oRS.next() Then
        sErr="Не удалось проверить партию " & sLot
        Exit Function
    End If
    n=oRS.getInt(1)
    If n<>1 Then
        sErr="Партия " & sLot & " не найдена или не уникальна."
        Exit Function
    End If

    WMSDBSR_VerifyReceiptCon=True
    Exit Function
EH:
    sErr="Проверка Smart Receipt: " & CStr(Err) & " " & Error$
End Function

Function WMSDBSR_ResolveRowUnits(oCon As Object,oSh As Object,r As Long, _
    sProductCode As String,dFactQty As Double,sVisibleUnit As String, _
    ByRef dBaseQty As Double,ByRef sBaseUnit As String,ByRef dFactor As Double, _
    ByRef sErr As String) As Boolean

    Dim configured As String,oldDocUnit As String,oldDocQty As Double
    Dim dDocQty As Double,docUnit As String,currentBase As String,lastFactor As Double,sInput As String

    WMSDBSR_ResolveRowUnits=False:sErr=""
    WMSDBSR_EnsureTechColumns oSh

    dDocQty=oSh.getCellByPosition(7,r).Value
    If dDocQty<=0 Then dDocQty=dFactQty
    docUnit=Trim(oSh.getCellByPosition(8,r).String)
    If docUnit="" Then docUnit=sVisibleUnit

    configured=WMSDBSR_TechText(oSh,r,WMSDBSR_T_CONFIG)
    oldDocUnit=WMSDBSR_TechText(oSh,r,WMSDBSR_T_DOCUNIT)
    oldDocQty=WMSDBSR_TechNumber(oSh,r,WMSDBSR_T_DOCQTY)

    If configured="1" Then
        If LCase(Trim(oldDocUnit))=LCase(Trim(docUnit)) And Abs(oldDocQty-dDocQty)<0.000001 Then
            dBaseQty=WMSDBSR_TechNumber(oSh,r,WMSDBSR_T_BASEQTY)
            sBaseUnit=WMSDBSR_TechText(oSh,r,WMSDBSR_T_BASEUNIT)
            dFactor=WMSDBSR_TechNumber(oSh,r,WMSDBSR_T_FACTOR)
            If dBaseQty>0 And sBaseUnit<>"" And dFactor>0 Then WMSDBSR_ResolveRowUnits=True:Exit Function
        Else
            WMSDBSR_ClearRowConfig oSh,r
        End If
    End If

    currentBase=WMSDBSR_GetProductBaseUnitCon(oCon,sProductCode,sErr)
    If sErr<>"" Then Exit Function
    If currentBase="" Then currentBase=docUnit

    If LCase(Trim(currentBase))=LCase(Trim(docUnit)) Then
        sBaseUnit=currentBase
        dBaseQty=dFactQty
        If dBaseQty<=0 Then dBaseQty=dDocQty
        dFactor=dBaseQty/dDocQty
        WMSDBSR_SaveRowConfig oSh,r,dDocQty,docUnit,dBaseQty,sBaseUnit,dFactor
        WMSDBSR_SetTechNumber oSh,r,WMSDBSR_T_ACTQTY,dFactQty
        WMSDBSR_SetTechText oSh,r,WMSDBSR_T_ACTUNIT,docUnit
        If dFactQty>0 Then
            WMSDBSR_SetTechNumber oSh,r,WMSDBSR_T_ACTFACTOR,dBaseQty/dFactQty
        Else
            WMSDBSR_SetTechNumber oSh,r,WMSDBSR_T_ACTFACTOR,1
        End If
        WMSDBSR_ResolveRowUnits=True
        Exit Function
    End If

    sBaseUnit=currentBase
    lastFactor=WMSDBSR_GetLastFactorCon(oCon,sProductCode,docUnit,sBaseUnit,sErr)
    If sErr<>"" Then Exit Function
    sInput=InputBox( _
        "Товар учитывается в остатке в '" & sBaseUnit & "'." & Chr(10) & _
        "Документ: " & WMSDBSR_NumText(dDocQty) & " " & docUnit & Chr(10) & _
        "Факт. количество: " & WMSDBSR_NumText(dFactQty) & Chr(10) & Chr(10) & _
        "Сколько '" & sBaseUnit & "' должно попасть в остаток?", _
        "WMS — Пересчёт партии",IIf(lastFactor>0,WMSDBSR_NumText(dDocQty*lastFactor),""))
    If Trim(sInput)="" Then sErr="Проведение отменено: пересчёт не подтверждён.":Exit Function
    If Not WMSDBSR_ParseNumber(sInput,dBaseQty) Or dBaseQty<=0 Then sErr="Некорректное базовое количество.":Exit Function

    dFactor=dBaseQty/dDocQty
    WMSDBSR_SaveRowConfig oSh,r,dDocQty,docUnit,dBaseQty,sBaseUnit,dFactor
    WMSDBSR_SetTechNumber oSh,r,WMSDBSR_T_ACTQTY,dFactQty
    WMSDBSR_SetTechText oSh,r,WMSDBSR_T_ACTUNIT,docUnit
    If dFactQty>0 Then
        WMSDBSR_SetTechNumber oSh,r,WMSDBSR_T_ACTFACTOR,dBaseQty/dFactQty
    Else
        WMSDBSR_SetTechNumber oSh,r,WMSDBSR_T_ACTFACTOR,dFactor
    End If
    WMSDBSR_ResolveRowUnits=True
End Function

' ============================================================================
' DB HELPERS
' ============================================================================

Function WMSDBSR_GetProductBaseUnit(sProductCode As String,ByRef sErr As String) As String
    Dim oCon As Object
    WMSDBSR_GetProductBaseUnit=""
    sErr=""
    On Error GoTo EH
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then Exit Function
    WMSDBSR_GetProductBaseUnit=WMSDBSR_GetProductBaseUnitCon(oCon,sProductCode,sErr)
    Exit Function
EH:
    sErr="Чтение базовой единицы: " & CStr(Err) & " " & Error$
End Function

Function WMSDBSR_GetProductBaseUnitCon(oCon As Object,sProductCode As String,ByRef sErr As String) As String
    Dim oStmt As Object,oRS As Object
    WMSDBSR_GetProductBaseUnitCon=""
    sErr=""
    On Error GoTo EH
    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery("SELECT UNIT_NAME FROM WMS_PRODUCTS WHERE PRODUCT_CODE=" & WMSDB_SQLText(sProductCode))
    If oRS.next() Then WMSDBSR_GetProductBaseUnitCon=Trim(oRS.getString(1))
    Exit Function
EH:
    sErr="Чтение WMS_PRODUCTS.UNIT_NAME: " & CStr(Err) & " " & Error$
End Function

Function WMSDBSR_GetLastFactor(sProductCode As String,sDocUnit As String,sBaseUnit As String, _
    ByRef sErr As String) As Double
    Dim oCon As Object
    WMSDBSR_GetLastFactor=0
    sErr=""
    On Error GoTo EH
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then Exit Function
    WMSDBSR_GetLastFactor=WMSDBSR_GetLastFactorCon(oCon,sProductCode,sDocUnit,sBaseUnit,sErr)
    Exit Function
EH:
    sErr="Поиск прошлого коэффициента: " & CStr(Err) & " " & Error$
End Function

Function WMSDBSR_GetLastFactorCon(oCon As Object,sProductCode As String,sDocUnit As String, _
    sBaseUnit As String,ByRef sErr As String) As Double

    Dim oStmt As Object,oRS As Object,sql As String
    WMSDBSR_GetLastFactorCon=0
    sErr=""
    On Error GoTo EH

    sql="SELECT u.FACTOR_TO_BASE FROM WMS_STOCK_LOTS l " & _
        "JOIN WMS_LOT_UNITS u ON u.LOT_ID=l.LOT_ID " & _
        "WHERE l.PRODUCT_CODE=" & WMSDB_SQLText(sProductCode) & _
        " AND l.BASE_UNIT=" & WMSDB_SQLText(sBaseUnit) & _
        " AND u.UNIT_NAME=" & WMSDB_SQLText(sDocUnit) & _
        " ORDER BY l.CREATED_AT DESC ROWS 1"

    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery(sql)
    If oRS.next() Then WMSDBSR_GetLastFactorCon=oRS.getDouble(1)
    Exit Function
EH:
    sErr="Поиск коэффициента партии: " & CStr(Err) & " " & Error$
End Function

Function WMSDBSR_LotMatchesCon(oCon As Object,sLotID As String,sProductCode As String, _
    dDocQty As Double,sDocUnit As String,dBaseQty As Double,sBaseUnit As String, _
    ByRef sErr As String) As Boolean

    Dim oStmt As Object,oRS As Object
    Dim q1 As Double,q2 As Double,u1 As String,u2 As String,p As String
    WMSDBSR_LotMatchesCon=False
    sErr=""
    On Error GoTo EH

    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery( _
        "SELECT PRODUCT_CODE,DOC_QTY,DOC_UNIT,BASE_QTY,BASE_UNIT FROM WMS_STOCK_LOTS WHERE LOT_ID=" & _
        WMSDB_SQLText(sLotID))

    If Not oRS.next() Then
        sErr="Не удалось прочитать существующую партию " & sLotID
        Exit Function
    End If

    p=Trim(oRS.getString(1))
    q1=oRS.getDouble(2)
    u1=Trim(oRS.getString(3))
    q2=oRS.getDouble(4)
    u2=Trim(oRS.getString(5))

    If p<>sProductCode Or Abs(q1-dDocQty)>0.0001 Or _
       LCase(u1)<>LCase(Trim(sDocUnit)) Or Abs(q2-dBaseQty)>0.0001 Or _
       LCase(u2)<>LCase(Trim(sBaseUnit)) Then
        sErr="Партия " & sLotID & " уже была записана с другим пересчётом." & Chr(10) & _
             "Исторический коэффициент изменять нельзя."
        Exit Function
    End If

    WMSDBSR_LotMatchesCon=True
    Exit Function
EH:
    sErr="Проверка неизменяемости партии: " & CStr(Err) & " " & Error$
End Function

' ============================================================================
' SHEET / TECH / BUTTON HELPERS
' ============================================================================

Sub WMSDBSR_EnsureTechColumns(oSh As Object)
    WMSDBSR_GetOrCreateTechColumn oSh,WMSDBSR_T_LOT
    WMSDBSR_GetOrCreateTechColumn oSh,WMSDBSR_T_DOCQTY
    WMSDBSR_GetOrCreateTechColumn oSh,WMSDBSR_T_DOCUNIT
    WMSDBSR_GetOrCreateTechColumn oSh,WMSDBSR_T_BASEQTY
    WMSDBSR_GetOrCreateTechColumn oSh,WMSDBSR_T_BASEUNIT
    WMSDBSR_GetOrCreateTechColumn oSh,WMSDBSR_T_FACTOR
    WMSDBSR_GetOrCreateTechColumn oSh,WMSDBSR_T_CONFIG
    WMSDBSR_GetOrCreateTechColumn oSh,WMSDBSR_T_ACTQTY
    WMSDBSR_GetOrCreateTechColumn oSh,WMSDBSR_T_ACTUNIT
    WMSDBSR_GetOrCreateTechColumn oSh,WMSDBSR_T_ACTFACTOR
End Sub

Function WMSDBSR_GetOrCreateTechColumn(oSh As Object,sHeader As String) As Long
    Dim c As Long,last As Long,cur As Object
    c=WMSDBSR_FindHeader(oSh,sHeader)
    If c>=0 Then
        WMSDBSR_GetOrCreateTechColumn=c
        On Error Resume Next
        oSh.Columns.getByIndex(c).IsVisible=False
        On Error GoTo 0
        Exit Function
    End If

    cur=oSh.createCursor()
    cur.gotoEndOfUsedArea(True)
    last=cur.RangeAddress.EndColumn
    c=last+1
    oSh.getCellByPosition(c,0).String=sHeader
    On Error Resume Next
    oSh.Columns.getByIndex(c).IsVisible=False
    On Error GoTo 0
    WMSDBSR_GetOrCreateTechColumn=c
End Function

Sub WMSDBSR_SaveRowConfig(oSh As Object,r As Long,dDocQty As Double,sDocUnit As String, _
    dBaseQty As Double,sBaseUnit As String,dFactor As Double)

    WMSDBSR_SetTechNumber oSh,r,WMSDBSR_T_DOCQTY,dDocQty
    WMSDBSR_SetTechText oSh,r,WMSDBSR_T_DOCUNIT,sDocUnit
    WMSDBSR_SetTechNumber oSh,r,WMSDBSR_T_BASEQTY,dBaseQty
    WMSDBSR_SetTechText oSh,r,WMSDBSR_T_BASEUNIT,sBaseUnit
    WMSDBSR_SetTechNumber oSh,r,WMSDBSR_T_FACTOR,dFactor
    WMSDBSR_SetTechText oSh,r,WMSDBSR_T_CONFIG,"1"
End Sub

Sub WMSDBSR_ClearRowConfig(oSh As Object,r As Long)
    WMSDBSR_SetTechText oSh,r,WMSDBSR_T_LOT,""
    WMSDBSR_SetTechNumber oSh,r,WMSDBSR_T_DOCQTY,0
    WMSDBSR_SetTechText oSh,r,WMSDBSR_T_DOCUNIT,""
    WMSDBSR_SetTechNumber oSh,r,WMSDBSR_T_BASEQTY,0
    WMSDBSR_SetTechText oSh,r,WMSDBSR_T_BASEUNIT,""
    WMSDBSR_SetTechNumber oSh,r,WMSDBSR_T_FACTOR,0
    WMSDBSR_SetTechText oSh,r,WMSDBSR_T_CONFIG,""
    WMSDBSR_SetTechNumber oSh,r,WMSDBSR_T_ACTQTY,0
    WMSDBSR_SetTechText oSh,r,WMSDBSR_T_ACTUNIT,""
    WMSDBSR_SetTechNumber oSh,r,WMSDBSR_T_ACTFACTOR,0
End Sub

Function WMSDBSR_FindHeader(oSh As Object,sHeader As String) As Long
    Dim c As Long,last As Long,cur As Object
    WMSDBSR_FindHeader=-1
    On Error GoTo Done
    cur=oSh.createCursorByRange(oSh.getCellByPosition(0,0))
    cur.gotoEndOfUsedArea(True)
    last=cur.RangeAddress.EndColumn
    For c=0 To last
        If Trim(oSh.getCellByPosition(c,0).String)=Trim(sHeader) Then
            WMSDBSR_FindHeader=c
            Exit Function
        End If
    Next c
Done:
End Function

Function WMSDBSR_CellText(oSh As Object,r As Long,sHeader As String) As String
    Dim c As Long
    WMSDBSR_CellText=""
    c=WMSDBSR_FindHeader(oSh,sHeader)
    If c>=0 Then WMSDBSR_CellText=Trim(oSh.getCellByPosition(c,r).String)
End Function

Function WMSDBSR_CellNumber(oSh As Object,r As Long,sHeader As String) As Double
    Dim c As Long
    WMSDBSR_CellNumber=0
    c=WMSDBSR_FindHeader(oSh,sHeader)
    If c>=0 Then WMSDBSR_CellNumber=oSh.getCellByPosition(c,r).Value
End Function

Function WMSDBSR_TechText(oSh As Object,r As Long,sHeader As String) As String
    WMSDBSR_TechText=WMSDBSR_CellText(oSh,r,sHeader)
End Function

Function WMSDBSR_TechNumber(oSh As Object,r As Long,sHeader As String) As Double
    WMSDBSR_TechNumber=WMSDBSR_CellNumber(oSh,r,sHeader)
End Function

Sub WMSDBSR_SetTechText(oSh As Object,r As Long,sHeader As String,sValue As String)
    Dim c As Long
    c=WMSDBSR_FindHeader(oSh,sHeader)
    If c>=0 Then oSh.getCellByPosition(c,r).String=sValue
End Sub

Sub WMSDBSR_SetTechNumber(oSh As Object,r As Long,sHeader As String,dValue As Double)
    Dim c As Long
    c=WMSDBSR_FindHeader(oSh,sHeader)
    If c>=0 Then oSh.getCellByPosition(c,r).Value=dValue
End Sub

Function WMSDBSR_MakeLotID(sSourceID As String) As String
    WMSDBSR_MakeLotID=Left("LOT-" & Trim(sSourceID),100)
End Function

Sub WMSDBSR_InstallButton(oDoc As Object,oSh As Object)
    Dim oForms As Object,oForm As Object,oModel As Object,oShape As Object
    Dim ev As New com.sun.star.script.ScriptEventDescriptor
    Dim p As New com.sun.star.awt.Point,z As New com.sun.star.awt.Size
    Dim idx As Long,baseX As Long,baseY As Long,w As Long,h As Long,gap As Long,colGap As Long

    On Error GoTo EH

    WMSDBSR_RemoveButton oSh
    oForms=oSh.DrawPage.Forms
    If oForms.hasByName(WMSDBSR_FORM) Then oForms.removeByName(WMSDBSR_FORM)

    oForm=oDoc.createInstance("com.sun.star.form.component.Form")
    oForm.Name=WMSDBSR_FORM
    oForms.insertByName(WMSDBSR_FORM,oForm)

    ' Same geometry as Orders right-hand panel; place below its SelfCheck.
    baseX=oSh.getCellByPosition(25,0).Position.X + _
          oSh.getCellByPosition(25,0).Size.Width + 350
    baseY=oSh.getCellByPosition(0,0).Position.Y + 120
    w=4100:h=650:gap=120:colGap=220
    baseX=baseX+w+colGap
    baseY=baseY+6*(h+gap)

    oModel=oDoc.createInstance("com.sun.star.form.component.CommandButton")
    oModel.Name="WMS_SR_BTN_UNITS"
    oModel.Label="Единицы партии"
    oModel.Tabstop=False
    oForm.insertByName(oModel.Name,oModel)
    idx=oForm.Count-1

    oShape=oDoc.createInstance("com.sun.star.drawing.ControlShape")
    p.X=baseX:p.Y=baseY:z.Width=w:z.Height=h
    oShape.Position=p:oShape.Size=z:oShape.Control=oModel
    oSh.DrawPage.add(oShape)

    ev.ListenerType="com.sun.star.awt.XActionListener"
    ev.EventMethod="actionPerformed"
    ev.AddListenerParam=""
    ev.ScriptType="Script"
    ev.ScriptCode="vnd.sun.star.script:Standard.WMS_12_DB_SmartReceipt.WMSDBSR_ConfigureSelected?language=Basic&location=document"
    oForm.registerScriptEvent(idx,ev)
    Exit Sub
EH:
    MsgBox "Кнопка Smart Receipt: " & CStr(Err) & " " & Error$,16,"WMS — Умный приход"
End Sub

Sub WMSDBSR_RemoveButton(oSh As Object)
    Dim i As Long,oShape As Object,nm As String,oForms As Object
    On Error Resume Next
    For i=oSh.DrawPage.getCount()-1 To 0 Step -1
        oShape=oSh.DrawPage.getByIndex(i)
        nm=""
        nm=oShape.Control.Name
        If Left(nm,7)="WMS_SR_" Then oSh.DrawPage.remove(oShape)
    Next i
    oForms=oSh.DrawPage.Forms
    If oForms.hasByName(WMSDBSR_FORM) Then oForms.removeByName(WMSDBSR_FORM)
    On Error GoTo 0
End Sub

' ============================================================================
' NUMBER HELPERS
' ============================================================================

Function WMSDBSR_ParseNumber(s As String,ByRef d As Double) As Boolean
    Dim t As String
    WMSDBSR_ParseNumber=False
    d=0
    On Error GoTo EH
    t=Trim(s)
    If t="" Then Exit Function
    t=Replace(t," ","")
    t=Replace(t,",",".")
    ' Val() is locale-independent enough for normalized dot decimal text.
    d=Val(t)
    If d<=0 Then Exit Function
    WMSDBSR_ParseNumber=True
    Exit Function
EH:
    WMSDBSR_ParseNumber=False
End Function

Function WMSDBSR_NumText(d As Double) As String
    Dim s As String
    s=Format(d,"0.######")
    WMSDBSR_NumText=s
End Function

Function WMSDBSR_SQLNumber(d As Double) As String
    Dim s As String
    s=CStr(d)
    s=Replace(s,",",".")
    WMSDBSR_SQLNumber=s
End Function

