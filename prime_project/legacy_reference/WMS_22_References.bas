Option Explicit

Global Const WMSREF_VERSION = "2.0.0-REFERENCES"
Global Const WMSREF_SHEET = "Справочники"
Global Const WMSREF_CACHE = "SYS_WMS_REF_CACHE"
Global Const WMSREF_FORM = "WMS_REF_FORM"

Sub WMSREF_Install()
    Dim oDoc As Object
    oDoc=ThisComponent
    On Error GoTo EH
    WMSREF_PrepareSheet oDoc
    WMSREF_Refresh
    WMSREF_ApplyValidations oDoc
    Exit Sub
EH:
    MsgBox "Справочники: " & CStr(Err) & " " & Error$,16,"WMS — Справочники"
End Sub

Sub WMSREF_PrepareSheet(oDoc As Object)
    Dim oSh As Object,oCache As Object,headers As Variant,i As Long
    If oDoc.Sheets.hasByName(WMSREF_SHEET) Then
        oSh=oDoc.Sheets.getByName(WMSREF_SHEET)
    Else
        oDoc.Sheets.insertNewByName(WMSREF_SHEET,oDoc.Sheets.getCount())
        oSh=oDoc.Sheets.getByName(WMSREF_SHEET)
    End If
    oSh.IsVisible=True

    If oDoc.Sheets.hasByName(WMSREF_CACHE) Then
        oCache=oDoc.Sheets.getByName(WMSREF_CACHE)
    Else
        oDoc.Sheets.insertNewByName(WMSREF_CACHE,oDoc.Sheets.getCount())
        oCache=oDoc.Sheets.getByName(WMSREF_CACHE)
    End If
    oCache.IsVisible=False

    oSh.getCellRangeByPosition(0,0,8,5).clearContents(1023)
    oSh.getCellByPosition(0,0).String="СПРАВОЧНИКИ WMS"
    oSh.getCellByPosition(0,1).String="Единые значения для всех рабочих листов. Добавляйте строки ниже и нажимайте «Сохранить»."
    oSh.getCellRangeByPosition(0,0,6,0).CellBackColor=2500134
    oSh.getCellRangeByPosition(0,0,6,0).CharColor=16777215
    oSh.getCellRangeByPosition(0,0,6,0).CharWeight=150
    oSh.getCellRangeByPosition(0,1,6,1).CellBackColor=15132390

    headers=Array("Тип","Код","Наименование","Родитель","Активен","Порядок","Комментарий")
    For i=0 To UBound(headers):oSh.getCellByPosition(i,4).String=CStr(headers(i)):Next i
    oSh.getCellRangeByPosition(0,4,6,4).CellBackColor=4473924
    oSh.getCellRangeByPosition(0,4,6,4).CharColor=16777215
    oSh.getCellRangeByPosition(0,4,6,4).CharWeight=150

    oSh.Columns.getByIndex(0).Width=3500:oSh.Columns.getByIndex(1).Width=4200
    oSh.Columns.getByIndex(2).Width=6200:oSh.Columns.getByIndex(3).Width=4200
    oSh.Columns.getByIndex(4).Width=2500:oSh.Columns.getByIndex(5).Width=2500
    oSh.Columns.getByIndex(6).Width=6500

    WMSREF_InstallButtons oDoc,oSh
End Sub

Sub WMSREF_Refresh()
    Dim oDoc As Object,oSh As Object,oCon As Object,oStmt As Object,oRS As Object
    Dim sErr As String,r As Long,lastR As Long,c As Object
    oDoc=ThisComponent
    If Not oDoc.Sheets.hasByName(WMSREF_SHEET) Then WMSREF_PrepareSheet oDoc
    oSh=oDoc.Sheets.getByName(WMSREF_SHEET)

    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then MsgBox sErr,16,"WMS — Справочники":Exit Sub

    c=oSh.createCursor():c.gotoEndOfUsedArea(True):lastR=c.RangeAddress.EndRow
    If lastR<5 Then lastR=5
    oSh.getCellRangeByPosition(0,5,6,lastR).clearContents(1023)

    On Error GoTo EH
    oStmt=oCon.createStatement()
    oRS=oStmt.executeQuery("SELECT REF_TYPE,REF_CODE,REF_NAME,COALESCE(PARENT_CODE,''),ACTIVE_FLAG,SORT_ORDER,COALESCE(NOTE_TEXT,'') " & _
        "FROM WMS_REFERENCE ORDER BY REF_TYPE,SORT_ORDER,REF_NAME")
    r=5
    Do While oRS.next()
        oSh.getCellByPosition(0,r).String=oRS.getString(1)
        oSh.getCellByPosition(1,r).String=oRS.getString(2)
        oSh.getCellByPosition(2,r).String=oRS.getString(3)
        oSh.getCellByPosition(3,r).String=oRS.getString(4)
        oSh.getCellByPosition(4,r).String=IIf(oRS.getInt(5)=1,"Да","Нет")
        oSh.getCellByPosition(5,r).Value=oRS.getInt(6)
        oSh.getCellByPosition(6,r).String=oRS.getString(7)
        r=r+1
    Loop
    WMSDB_Close
    WMSREF_RefreshCache
    Exit Sub
EH:
    WMSDB_Close
    MsgBox "Чтение справочников: " & CStr(Err) & " " & Error$,16,"WMS — Справочники"
End Sub

Sub WMSREF_Save()
    Dim oDoc As Object,oSh As Object,oCon As Object,oStmt As Object
    Dim sErr As String,r As Long,lastR As Long,c As Object
    Dim typ As String,code As String,nm As String,parent As String,active As Integer,sortN As Long,note As String,sql As String
    oDoc=ThisComponent:oSh=oDoc.Sheets.getByName(WMSREF_SHEET)
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then MsgBox sErr,16,"WMS — Справочники":Exit Sub
    On Error GoTo EH
    oStmt=oCon.createStatement()
    c=oSh.createCursor():c.gotoEndOfUsedArea(True):lastR=c.RangeAddress.EndRow

    For r=5 To lastR
        typ=UCase(Trim(oSh.getCellByPosition(0,r).String))
        code=Trim(oSh.getCellByPosition(1,r).String)
        nm=Trim(oSh.getCellByPosition(2,r).String)
        parent=Trim(oSh.getCellByPosition(3,r).String)
        active=IIf(UCase(Trim(oSh.getCellByPosition(4,r).String))="НЕТ",0,1)
        sortN=CLng(oSh.getCellByPosition(5,r).Value)
        note=Trim(oSh.getCellByPosition(6,r).String)

        If typ<>"" Or code<>"" Or nm<>"" Then
            If typ="" Or nm="" Then
                sErr="Строка " & CStr(r+1) & ": обязательны Тип и Наименование.":GoTo Fail
            End If
            If code="" Then
                If typ="SUBCATEGORY" And parent<>"" Then
                    code=parent & ">" & nm
                Else
                    code=nm
                End If
            End If
            sql="UPDATE WMS_REFERENCE SET REF_NAME=" & WMSDB_SQLText(nm) & ",PARENT_CODE=" & WMSDB_SQLText(parent) & _
                ",ACTIVE_FLAG=" & CStr(active) & ",SORT_ORDER=" & CStr(sortN) & ",NOTE_TEXT=" & WMSDB_SQLText(note) & _
                ",UPDATED_AT=CURRENT_TIMESTAMP WHERE REF_TYPE=" & WMSDB_SQLText(typ) & " AND REF_CODE=" & WMSDB_SQLText(code)
            If oStmt.executeUpdate(sql)=0 Then
                oStmt.executeUpdate("INSERT INTO WMS_REFERENCE (REF_TYPE,REF_CODE,REF_NAME,PARENT_CODE,ACTIVE_FLAG,SORT_ORDER,NOTE_TEXT) VALUES (" & _
                    WMSDB_SQLText(typ) & "," & WMSDB_SQLText(code) & "," & WMSDB_SQLText(nm) & "," & WMSDB_SQLText(parent) & "," & _
                    CStr(active) & "," & CStr(sortN) & "," & WMSDB_SQLText(note) & ")")
            End If
        End If
    Next r
    oCon.commit()
    WMSDB_Close
    WMSREF_Refresh
    WMSREF_ApplyValidations oDoc
    MsgBox "Справочники сохранены.",64,"WMS — Справочники"
    Exit Sub
Fail:
    On Error Resume Next:oCon.rollback():WMSDB_Close
    MsgBox sErr,16,"WMS — Справочники":Exit Sub
EH:
    sErr="Сохранение справочников: " & CStr(Err) & " " & Error$:Resume Fail
End Sub

Sub WMSREF_RefreshCache()
    Dim oDoc As Object,oSh As Object,oCon As Object,oStmt As Object,oRS As Object
    Dim types As Variant,i As Long,r As Long,sErr As String
    oDoc=ThisComponent
    If Not oDoc.Sheets.hasByName(WMSREF_CACHE) Then
        oDoc.Sheets.insertNewByName(WMSREF_CACHE,oDoc.Sheets.getCount())
    End If
    oSh=oDoc.Sheets.getByName(WMSREF_CACHE):oSh.IsVisible=False
    oSh.getCellRangeByPosition(0,0,12,5000).clearContents(1023)
    types=Array("UNIT","LOCATION","CATEGORY","SUBCATEGORY","SUPPLIER","SELLER","EMPLOYEE","CONTOUR")
    oCon=WMSDB_GetConnectionEx(oDoc,sErr)
    If sErr<>"" Then Exit Sub
    On Error GoTo EH
    For i=0 To UBound(types)
        oSh.getCellByPosition(i,0).String=CStr(types(i))
        oStmt=oCon.createStatement()
        oRS=oStmt.executeQuery("SELECT REF_NAME FROM WMS_REFERENCE WHERE REF_TYPE=" & WMSDB_SQLText(CStr(types(i))) & _
            " AND ACTIVE_FLAG=1 ORDER BY SORT_ORDER,REF_NAME")
        r=1
        Do While oRS.next()
            oSh.getCellByPosition(i,r).String=oRS.getString(1):r=r+1
        Loop
    Next i
    WMSDB_Close
    Exit Sub
EH:
    WMSDB_Close
End Sub

Sub WMSREF_ApplyValidations(oDoc As Object)
    On Error Resume Next
    If oDoc.Sheets.hasByName("Приход — Производство") Then
        WMSREF_SetList oDoc.Sheets.getByName("Приход — Производство"),4,"UNIT"
        WMSREF_SetList oDoc.Sheets.getByName("Приход — Производство"),6,"LOCATION"
        WMSREF_SetList oDoc.Sheets.getByName("Приход — Производство"),7,"CATEGORY"
        WMSREF_SetList oDoc.Sheets.getByName("Приход — Производство"),8,"SUBCATEGORY"
    End If
    If oDoc.Sheets.hasByName("Приход — Офис") Then
        WMSREF_SetList oDoc.Sheets.getByName("Приход — Офис"),4,"UNIT"
        WMSREF_SetList oDoc.Sheets.getByName("Приход — Офис"),6,"CONTOUR"
        WMSREF_SetList oDoc.Sheets.getByName("Приход — Офис"),7,"LOCATION"
        WMSREF_SetList oDoc.Sheets.getByName("Приход — Офис"),8,"CATEGORY"
        WMSREF_SetList oDoc.Sheets.getByName("Приход — Офис"),9,"SUBCATEGORY"
    End If
    If oDoc.Sheets.hasByName("Приход — Детали") Then
        WMSREF_SetList oDoc.Sheets.getByName("Приход — Детали"),4,"UNIT"
        WMSREF_SetList oDoc.Sheets.getByName("Приход — Детали"),6,"LOCATION"
        WMSREF_SetList oDoc.Sheets.getByName("Приход — Детали"),7,"CATEGORY"
    End If
    If oDoc.Sheets.hasByName("Расход — Производство") Then
        WMSREF_SetList oDoc.Sheets.getByName("Расход — Производство"),4,"UNIT"
        WMSREF_SetList oDoc.Sheets.getByName("Расход — Производство"),5,"EMPLOYEE"
        WMSREF_SetList oDoc.Sheets.getByName("Расход — Производство"),7,"LOCATION"
    End If
    If oDoc.Sheets.hasByName("Расход — Офис") Then
        WMSREF_SetList oDoc.Sheets.getByName("Расход — Офис"),4,"UNIT"
        WMSREF_SetList oDoc.Sheets.getByName("Расход — Офис"),5,"EMPLOYEE"
        WMSREF_SetList oDoc.Sheets.getByName("Расход — Офис"),6,"CONTOUR"
        WMSREF_SetList oDoc.Sheets.getByName("Расход — Офис"),7,"LOCATION"
    End If
    If oDoc.Sheets.hasByName("Расход — Детали") Then
        WMSREF_SetList oDoc.Sheets.getByName("Расход — Детали"),4,"UNIT"
        WMSREF_SetList oDoc.Sheets.getByName("Расход — Детали"),5,"EMPLOYEE"
        WMSREF_SetList oDoc.Sheets.getByName("Расход — Детали"),7,"LOCATION"
    End If
    On Error GoTo 0
End Sub

Sub WMSREF_SetList(oSh As Object,nCol As Long,sType As String)
    Dim oDoc As Object,oCache As Object,i As Long,nCacheCol As Long,lastRow As Long,v As Object,formula As String
    oDoc=ThisComponent:oCache=oDoc.Sheets.getByName(WMSREF_CACHE)
    nCacheCol=WMSREF_CacheCol(sType)
    If nCacheCol<0 Then Exit Sub
    lastRow=1
    For i=1 To 5000
        If Trim(oCache.getCellByPosition(nCacheCol,i).String)="" Then Exit For
        lastRow=i
    Next i
    If lastRow<1 Then Exit Sub

    v=oSh.getCellRangeByPosition(nCol,5,nCol,5000).Validation
    v.Type=6
    formula="$'" & WMSREF_CACHE & "'.$" & Chr(65+nCacheCol) & "$2:$" & Chr(65+nCacheCol) & "$" & CStr(lastRow+1)
    v.Formula1=formula
    v.ShowList=1
    v.ShowErrorMessage=False
    oSh.getCellRangeByPosition(nCol,5,nCol,5000).Validation=v
End Sub

Function WMSREF_CacheCol(sType As String) As Long
    Select Case UCase(sType)
        Case "UNIT":WMSREF_CacheCol=0
        Case "LOCATION":WMSREF_CacheCol=1
        Case "CATEGORY":WMSREF_CacheCol=2
        Case "SUBCATEGORY":WMSREF_CacheCol=3
        Case "SUPPLIER":WMSREF_CacheCol=4
        Case "SELLER":WMSREF_CacheCol=5
        Case "EMPLOYEE":WMSREF_CacheCol=6
        Case "CONTOUR":WMSREF_CacheCol=7
        Case Else:WMSREF_CacheCol=-1
    End Select
End Function

Sub WMSREF_InstallButtons(oDoc As Object,oSh As Object)
    Dim oForms As Object,oForm As Object
    WMSREF_RemoveButtons oSh
    oForms=oSh.DrawPage.Forms
    On Error Resume Next
    If oForms.hasByName(WMSREF_FORM) Then oForms.removeByName(WMSREF_FORM)
    On Error GoTo EH
    oForm=oDoc.createInstance("com.sun.star.form.component.Form"):oForm.Name=WMSREF_FORM
    oForms.insertByName(WMSREF_FORM,oForm)
    WMSREF_AddButton oDoc,oSh,oForm,"WMS_REF_SAVE","Сохранить",0,2,3000,800,"WMSREF_Save"
    WMSREF_AddButton oDoc,oSh,oForm,"WMS_REF_REFRESH","Обновить",1,2,3000,800,"WMSREF_Refresh"
    Exit Sub
EH:
    MsgBox "Кнопки справочников: " & CStr(Err) & " " & Error$,16,"WMS — Справочники"
End Sub

Sub WMSREF_AddButton(oDoc As Object,oSh As Object,oForm As Object,sName As String,sLabel As String,nCol As Long,nRow As Long,nW As Long,nH As Long,sMacro As String)
    Dim oModel As Object,oShape As Object,a As Object,sz As Object,idx As Long
    Dim ev As New com.sun.star.script.ScriptEventDescriptor
    oModel=oDoc.createInstance("com.sun.star.form.component.CommandButton")
    oModel.Name=sName:oModel.Label=sLabel:oModel.Tabstop=False
    oForm.insertByName(sName,oModel):idx=oForm.Count-1
    oShape=oDoc.createInstance("com.sun.star.drawing.ControlShape"):oShape.Control=oModel
    a=oSh.getCellByPosition(nCol,nRow).Position:oShape.Position=a
    sz=CreateUnoStruct("com.sun.star.awt.Size"):sz.Width=nW:sz.Height=nH:oShape.Size=sz
    oSh.DrawPage.add(oShape)
    ev.ListenerType="com.sun.star.awt.XActionListener":ev.EventMethod="actionPerformed":ev.ScriptType="Script"
    ev.ScriptCode="vnd.sun.star.script:Standard.WMS_22_References." & sMacro & "?language=Basic&location=document"
    oForm.registerScriptEvent idx,ev
End Sub

Sub WMSREF_RemoveButtons(oSh As Object)
    Dim i As Long,oShape As Object,oCtl As Object,nm As String
    On Error Resume Next
    For i=oSh.DrawPage.getCount()-1 To 0 Step -1
        oShape=oSh.DrawPage.getByIndex(i):nm="":oCtl=oShape.Control:nm=oCtl.Name
        If Left(nm,8)="WMS_REF_" Then oSh.DrawPage.remove(oShape)
    Next i
    On Error GoTo 0
End Sub
