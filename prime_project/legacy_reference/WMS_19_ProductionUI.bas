Option Explicit

Global Const WMSPROD_VERSION = "3.3.1-VYDACHA-TOVARY-FIX"

Sub WMSPROD_ApplyUI()
    Dim oDoc As Object,oCon As Object,sErr As String
    oDoc=ThisComponent
    On Error GoTo EH

    ' Architecture is migrated before any interface is rebuilt.
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then GoTo Fail
    If Not WMSARCH_EnsureSchema(oCon,sErr) Then GoTo Fail
    If Not WMSINT_EnsureSchema(oCon,sErr) Then GoTo Fail
    oCon.commit():WMSDB_Close

    WMSPROD_SetSheetVisibility oDoc
    WMS26_OptimizeWorkbook

    ' Build complete working interfaces.
    WMSREF_Install
    WMSWF_Install
    WMSDBST_InstallStockSheet oDoc
    WMSOPS_Install
    WMSSEARCH2_Install
    WMSUI_ApplyAll

    On Error Resume Next
    WMSC_Install
    WMSRESET_Install
    WMSX_InstallUI
    On Error GoTo EH

    WMSOEX_Install
    WMSIMP_InstallUI
    WMSX_OrderStockInstall
    WMSRPT_Install
    WMSPRIME_LayoutAll
    WMSPROD_SetSheetVisibility oDoc
    oDoc.store()
    Exit Sub
Fail:
    WMSDB_Close
    MsgBox "Production UI не применён: " & sErr,16,"WMS"
    Exit Sub
EH:
    On Error Resume Next:WMSDB_Close
    MsgBox "Production UI: " & CStr(Err) & " " & Error$,16,"WMS"
End Sub

Sub WMSPROD_SetSheetVisibility(oDoc As Object)
    Dim visible As Variant,hidden As Variant,i As Long,nm As Variant,oSh As Object
    visible=Array("Инфо","Заказы","Выдачи","Приход — Цех","Приход — Офис", _
        "Расход — Цех","Расход — Офис","Остаток","Возвраты","Инвентаризация","База - Поиск","Справочники")
    hidden=Array("Номенклатура","Сопоставления","Журнал WMS","Ошибки WMS","SYS_WMS_QUEUE","SYS_WMS_INDEX","Настройки WMS","SYS_WMS_REF_CACHE")

    On Error Resume Next
    For i=LBound(visible) To UBound(visible)
        nm=CStr(visible(i))
        If oDoc.Sheets.hasByName(nm) Then
            oSh=oDoc.Sheets.getByName(nm):oSh.IsVisible=True
        End If
    Next i
    For i=LBound(hidden) To UBound(hidden)
        nm=CStr(hidden(i))
        If oDoc.Sheets.hasByName(nm) Then
            oSh=oDoc.Sheets.getByName(nm):oSh.IsVisible=False
        End If
    Next i
    On Error GoTo 0
    For Each nm In Array("Приход — Производство","Приход — Детали","Расход — Производство","Расход — Детали")
        If oDoc.Sheets.hasByName(nm) Then
            oSh=oDoc.Sheets.getByName(nm)
            oSh.IsVisible=(WMSCore_LastContentRow(oSh,11)>=5)
        End If
    Next nm
End Sub

Sub WMSPROD_RemoveObsoleteSheets(oDoc As Object)
    ' Production 2.0 is update-only: no operational sheet is deleted automatically.
    ' Old service sheets are merely hidden by WMSPROD_SetSheetVisibility.
End Sub

Sub WMSPROD_PruneButtons(oDoc As Object)
    ' No destructive pruning in 2.0. New panels manage only their own buttons.
End Sub

Sub WMSPROD_SelfCheck()
    Dim oDoc As Object,oCon As Object,sErr As String,s As String,nProblems As Long,errCount As Long
    oDoc=ThisComponent
    s="WMS Production " & WMSPROD_VERSION & Chr(10) & Chr(10)

    WMSPROD_CheckSheet oDoc,"Инфо",True,s,errCount
    WMSPROD_CheckSheet oDoc,"Заказы",True,s,errCount
    WMSPROD_CheckSheet oDoc,"Выдачи",True,s,errCount
    WMSPROD_CheckSheet oDoc,"Приход — Цех",True,s,errCount
    WMSPROD_CheckSheet oDoc,"Приход — Офис",True,s,errCount
    WMSPROD_CheckSheet oDoc,"Приход — Детали",False,s,errCount
    WMSPROD_CheckSheet oDoc,"Расход — Цех",True,s,errCount
    WMSPROD_CheckSheet oDoc,"Расход — Офис",True,s,errCount
    WMSPROD_CheckSheet oDoc,"Расход — Детали",False,s,errCount
    WMSPROD_CheckSheet oDoc,"Остаток",True,s,errCount
    WMSPROD_CheckSheet oDoc,"Возвраты",True,s,errCount
    WMSPROD_CheckSheet oDoc,"Инвентаризация",True,s,errCount
    WMSPROD_CheckSheet oDoc,"База - Поиск",True,s,errCount
    WMSPROD_CheckSheet oDoc,"Справочники",True,s,errCount

    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then
        s=s & "Firebird: ОШИБКА — " & sErr & Chr(10):errCount=errCount+1
    Else
        s=s & WMSARCH_DeepCheckText(oCon,nProblems,sErr)
        WMSDB_Close
        If sErr<>"" Then
            s=s & "Integrity: ОШИБКА — " & sErr & Chr(10):errCount=errCount+1
        ElseIf nProblems>0 Then
            errCount=errCount+nProblems
        End If
    End If

    s=s & Chr(10)
    If errCount=0 Then
        s=s & "РЕЗУЛЬТАТ: система готова к работе."
        MsgBox s,64,"WMS — Полная проверка"
    Else
        s=s & "РЕЗУЛЬТАТ: обнаружено проблемных групп: " & CStr(errCount)
        MsgBox s,48,"WMS — Полная проверка"
    End If
End Sub

Sub WMSPROD_CheckSheet(oDoc As Object,nm As String,mustVisible As Boolean,ByRef s As String,ByRef errCount As Long)
    Dim oSh As Object
    If Not oDoc.Sheets.hasByName(nm) Then
        s=s & "ОШИБКА: нет листа «" & nm & "»." & Chr(10):errCount=errCount+1:Exit Sub
    End If
    oSh=oDoc.Sheets.getByName(nm)
    If mustVisible And Not oSh.IsVisible Then
        s=s & "ОШИБКА: лист «" & nm & "» скрыт." & Chr(10):errCount=errCount+1
    Else
        s=s & nm & ": OK" & Chr(10)
    End If
End Sub

Function WMSPROD_IsManagedButton(nm As String,panel As String) As Boolean
    WMSPROD_IsManagedButton=False
End Function

Function WMSPROD_IsAllowedButton(nm As String,panel As String) As Boolean
    WMSPROD_IsAllowedButton=True
End Function
