Option Explicit

Global Const WMSINST_VERSION = "1.4.1-STAGED-UPDATE"
Global Const WMSINST_FOLDER = "WMS_UPDATE_PACKAGE"

Sub WMS_INSTALL_FULL()
    MsgBox "Эта WMS уже является рабочей производственной базой." & Chr(10) & Chr(10) & _
           "Полная пересборка ЗАПРЕЩЕНА, чтобы не потерять рабочие данные." & Chr(10) & _
           "Для новых версий используйте только WMS_UPDATE.",48,"WMS — Production"
End Sub

Sub WMS_UPDATE()
    Dim pkg As String,modulesURL As String,sErr As String,stage As String
    Dim nNew As Long,nUpd As Long,nSame As Long,changes As Long
    Dim con As Object,st As Object,exists As Boolean
    On Error GoTo EH
    stage="Проверка папки":WMSINST_LogStage stage
    If Trim(ThisComponent.URL)="" Then MsgBox "Сначала сохраните книгу.",48,"WMS":Exit Sub
    pkg=WMSINST_FindPackageFolder()
    If pkg="" Then MsgBox "Рядом с книгой нет WMS_UPDATE_PACKAGE.",48,"WMS":Exit Sub
    modulesURL=ConvertToURL(pkg & WMSINST_PathSep() & "Modules" & WMSINST_PathSep())
    stage="Сравнение модулей":WMSINST_LogStage stage
    changes=WMSINST_CountChanges(modulesURL,sErr)
    If sErr<>"" Then GoTo Fail
    If changes>0 Then
        stage="Резервная копия перед заменой модулей":WMSINST_LogStage stage
        If Not WMSINST_Backup(sErr) Then GoTo Fail
        stage="Импорт модулей":WMSINST_LogStage stage
        If Not WMSINST_ImportModules(modulesURL,nNew,nUpd,nSame,sErr) Then GoTo Fail
        ' Do not run changed code or rebuild the UI in this Basic session.
        ThisComponent.store()
        WMSINST_LogStage "Модули сохранены; требуется перезапуск LibreOffice"
        MsgBox "Модули сохранены. Полностью закройте LibreOffice, откройте книгу и нажмите «Обновить WMS» ещё раз для проверки базы. Интерфейс не пересобирался.",64,"WMS — Перезапуск"
        Exit Sub
    End If
    stage="Открытие базы":WMSINST_LogStage stage
    con=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then GoTo Fail
    stage="Проверка таблицы снимков заказов":WMSINST_LogStage stage
    exists=WMSDB_TableExistsSimple(con,"WMS_ORDER_SNAPSHOT",sErr)
    If sErr<>"" Then GoTo Fail
    If Not exists Then
        stage="Создание таблицы снимков заказов":WMSINST_LogStage stage
        st=con.createStatement()
        st.executeUpdate("CREATE TABLE WMS_ORDER_SNAPSHOT (SOURCE_ID VARCHAR(96) NOT NULL,COL_INDEX INTEGER NOT NULL,HEADER_NAME VARCHAR(255),VALUE_TEXT VARCHAR(8000),PRIMARY KEY(SOURCE_ID,COL_INDEX))")
        st.close():con.commit()
        stage="Сохранение базы":WMSINST_LogStage stage
        If Not WMSDB_SaveDatabaseDocument(sErr) Then GoTo Fail
    End If
    ' Leave the read connection open; ordinary operations reuse it.
    WMSINST_LogStage "Готово. Модули совпадают, таблица снимков существует."
    MsgBox "Версия 1.4.1: модули проверены, таблица снимков заказов готова. Пересборка листов и полный пересчёт не запускались.",64,"WMS — Готово"
    Exit Sub
Fail:
    WMSINST_LogStage "Ошибка на этапе: " & stage & " — " & sErr
    MsgBox "Обновление остановлено. Этап: " & stage & Chr(10) & sErr,16,"WMS"
    Exit Sub
EH:
    sErr=CStr(Err) & " " & Error$
    Resume Fail
End Sub

Function WMSINST_CountChanges(folderURL As String,ByRef sErr As String) As Long
    Dim sfa As Object,files As Variant,i As Long,nm As String,src As String,libObj As Object
    WMSINST_CountChanges=0:sErr=""
    On Error GoTo EH
    sfa=CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    files=sfa.getFolderContents(folderURL,False):libObj=WMSINST_StandardLibrary()
    For i=0 To UBound(files)
        If LCase(Right(CStr(files(i)),4))=".bas" Then
            nm=WMSINST_FileBaseName(CStr(files(i)))
            If UCase(nm)<>"WMS_99_INSTALLER" Then
                src=WMSINST_ReadText(CStr(files(i)))
                If Trim(src)="" Then sErr="Пустой модуль: " & nm:Exit Function
                If libObj.hasByName(nm) Then
                    If WMSINST_Normalize(CStr(libObj.getByName(nm)))<>WMSINST_Normalize(src) Then WMSINST_CountChanges=WMSINST_CountChanges+1
                Else
                    WMSINST_CountChanges=WMSINST_CountChanges+1
                End If
            End If
        End If
    Next i
    Exit Function
EH:
    sErr="Сравнение модулей: " & CStr(Err) & " " & Error$
End Function

Sub WMSINST_LogStage(stage As String)
    Dim sfa As Object,outStream As Object,txt As Object,folder As String
    On Error GoTo Done
    folder=WMSINST_BaseFolder() & "Diagnostics" & WMSINST_PathSep()
    WMSINST_EnsureFolder folder
    sfa=CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If sfa.exists(ConvertToURL(folder & "update_last.txt")) Then sfa.kill(ConvertToURL(folder & "update_last.txt"))
    outStream=sfa.openFileWrite(ConvertToURL(folder & "update_last.txt"))
    txt=CreateUnoService("com.sun.star.io.TextOutputStream")
    txt.setOutputStream(outStream):txt.setEncoding("UTF-8")
    txt.writeString(Format(Now,"YYYY-MM-DD HH:MM:SS") & " | " & stage & Chr(10))
    txt.closeOutput()
Done:
End Sub

Sub WMS_REPAIR()
    Dim sReport As String
    If WMSINST_RunSafeMigrations(sReport) Then
        WMSPROD_ApplyUI
        WMSC_Refresh
        ThisComponent.store()
        MsgBox "Безопасное восстановление завершено." & Chr(10) & _
               "Рабочие данные и база не пересобирались." & Chr(10) & Chr(10) & sReport,64,"WMS — Восстановление"
    Else
        MsgBox "Восстановление остановлено:" & Chr(10) & Chr(10) & sReport,16,"WMS — Восстановление"
    End If
End Sub

Sub WMS_INSTALLER_STATUS()
    Dim oLib As Object
    Dim sReport As String
    Dim moduleNames As Variant
    Dim i As Long
    Dim moduleSource As String

    On Error GoTo EH
    oLib=WMSINST_StandardLibrary()
    moduleNames=oLib.getElementNames()
    sReport="WMS Update Manager " & WMSINST_VERSION & Chr(10) & Chr(10)

    For i=0 To UBound(moduleNames)
        If Left(CStr(moduleNames(i)),4)="WMS_" Then
            moduleSource=CStr(oLib.getByName(CStr(moduleNames(i))))
            sReport=sReport & CStr(moduleNames(i)) & " | " & WMSINST_SourceVersion(moduleSource) & Chr(10)
        End If
    Next i

    MsgBox sReport,64,"WMS — Установленные модули"
    Exit Sub
EH:
    MsgBox CStr(Err) & " " & Error$,16,"WMS — Статус"
End Sub

Function WMSINST_RunSafeMigrations(ByRef sReport As String) As Boolean
    Dim oCon As Object
    Dim sErr As String

    WMSINST_RunSafeMigrations=False
    sReport=""
    On Error GoTo EH

    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then GoTo Fail

    If Not WMSDBI_EnsureSchema(oCon,sErr) Then GoTo Fail
    If Not WMSDBO_EnsureTable(oCon,sErr) Then GoTo Fail
    If Not WMSDBO_EnsureOrdersTable(oCon,sErr) Then GoTo Fail
    If Not WMSDBIu_EnsureTables(oCon,sErr) Then GoTo Fail
    If Not WMSDBR_EnsureLotColumns(oCon,sErr) Then GoTo Fail
    If Not WMSDBST_EnsureSchema(oCon,sErr) Then GoTo Fail
    If Not WMSDBUL_EnsureSchema(oCon,sErr) Then GoTo Fail
    If Not WMSRC_EnsureDB(oCon,sErr) Then GoTo Fail
    If Not WMSSAFE_EnsureSchema(oCon,sErr) Then GoTo Fail
    If Not WMSACTREG_EnsureSchema(oCon,sErr) Then GoTo Fail

    On Error Resume Next
    oCon.commit()
    On Error GoTo EH

    If Not WMSDB_SaveDatabaseDocument(sErr) Then GoTo Fail
    WMSDB_Close

    WMSINST_EnsureCalcTechColumns

    sReport="Firebird schema: OK" & Chr(10) & _
            "Safety/Audit: OK" & Chr(10) & _
            "Акты: OK" & Chr(10) & _
            "Рабочие листы: без пересборки"
    WMSINST_RunSafeMigrations=True
    Exit Function

Fail:
    On Error Resume Next
    oCon.rollback()
    WMSDB_Close
    sReport="Ошибка миграции: " & sErr
    Exit Function
EH:
    On Error Resume Next
    oCon.rollback()
    WMSDB_Close
    sReport="Миграция: " & CStr(Err) & " " & Error$
End Function

Sub WMSINST_EnsureCalcTechColumns()
    Dim oDoc As Object
    Dim oSh As Object

    On Error Resume Next
    oDoc=ThisComponent
    If oDoc.Sheets.hasByName("Заказы") Then
        oSh=oDoc.Sheets.getByName("Заказы")
        WMSRC_GetOrCreateTechCol oSh,WMSRC_H_EVENT
        WMSRC_GetOrCreateTechCol oSh,WMSRC_H_MODE
        WMSRC_GetOrCreateTechCol oSh,WMSRC_H_SOURCE
        WMSRC_GetOrCreateTechCol oSh,WMSRC_H_FROM
        WMSRC_GetOrCreateTechCol oSh,WMSRC_H_EXTORDER
        WMSRC_GetOrCreateTechCol oSh,WMSRC_H_UPD
        WMSRC_HideTechCols oSh
        WMSDBSR_EnsureTechColumns oSh
    End If
    On Error GoTo 0
End Sub

Function WMSINST_ImportModules(folderURL As String,ByRef nNew As Long,ByRef nUpd As Long,ByRef nSame As Long,ByRef sErr As String) As Boolean
    Dim oSFA As Object
    Dim fileList As Variant
    Dim i As Long
    Dim fileURL As String
    Dim moduleName As String
    Dim moduleSource As String
    Dim oldSource As String
    Dim oLib As Object

    WMSINST_ImportModules=False
    sErr=""
    nNew=0
    nUpd=0
    nSame=0
    On Error GoTo EH

    oSFA=CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If Not oSFA.exists(folderURL) Then
        sErr="Папка Modules не найдена."
        Exit Function
    End If

    oLib=WMSINST_StandardLibrary()
    fileList=oSFA.getFolderContents(folderURL,False)

    For i=0 To UBound(fileList)
        fileURL=CStr(fileList(i))
        If LCase(Right(fileURL,4))=".bas" Then
            moduleName=WMSINST_FileBaseName(fileURL)
            If UCase(moduleName)<>"WMS_99_INSTALLER" Then
                moduleSource=WMSINST_ReadText(fileURL)
                If Trim(moduleSource)="" Then
                    sErr="Пустой модуль: " & moduleName
                    Exit Function
                End If

                If oLib.hasByName(moduleName) Then
                    oldSource=CStr(oLib.getByName(moduleName))
                    If WMSINST_Normalize(oldSource)=WMSINST_Normalize(moduleSource) Then
                        nSame=nSame+1
                    Else
                        oLib.replaceByName(moduleName,moduleSource)
                        nUpd=nUpd+1
                    End If
                Else
                    oLib.insertByName(moduleName,moduleSource)
                    nNew=nNew+1
                End If
            End If
        End If
    Next i

    WMSINST_ImportModules=True
    Exit Function
EH:
    sErr="Импорт " & moduleName & ": " & CStr(Err) & " " & Error$
End Function

Function WMSINST_Backup(ByRef sErr As String) As Boolean
    Dim oSFA As Object
    Dim baseFolder As String
    Dim backupRoot As String
    Dim backupDir As String
    Dim dstURL As String
    Dim odbURL As String

    WMSINST_Backup=False
    sErr=""
    On Error GoTo EH

    ThisComponent.store()
    On Error Resume Next
    WMSDB_Close
    On Error GoTo EH

    baseFolder=WMSINST_BaseFolder()
    backupRoot=baseFolder & "WMS_Backup" & WMSINST_PathSep()
    WMSINST_EnsureFolder backupRoot

    backupDir=backupRoot & WMSINST_TimeStamp() & WMSINST_PathSep()
    WMSINST_EnsureFolder backupDir

    oSFA=CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    dstURL=ConvertToURL(backupDir & WMSINST_FileNameFromURL(ThisComponent.URL))
    oSFA.copy(ThisComponent.URL,dstURL)

    odbURL=ConvertToURL(baseFolder & "WMS_DATA_PORTABLE.odb")
    If oSFA.exists(odbURL) Then
        oSFA.copy(odbURL,ConvertToURL(backupDir & "WMS_DATA_PORTABLE.odb"))
    End If

    WMSINST_Backup=True
    Exit Function
EH:
    sErr="Backup: " & CStr(Err) & " " & Error$
End Function

Function WMSINST_StandardLibrary() As Object
    Dim oLibraries As Object
    oLibraries=ThisComponent.BasicLibraries
    If Not oLibraries.isLibraryLoaded("Standard") Then
        oLibraries.loadLibrary("Standard")
    End If
    WMSINST_StandardLibrary=oLibraries.getByName("Standard")
End Function

Function WMSINST_ReadText(fileURL As String) As String
    Dim oSFA As Object
    Dim oInput As Object
    Dim oText As Object
    Dim sText As String

    oSFA=CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    oInput=oSFA.openFileRead(fileURL)
    oText=CreateUnoService("com.sun.star.io.TextInputStream")
    oText.setInputStream(oInput)
    oText.setEncoding("UTF-8")

    Do While Not oText.isEOF()
        sText=sText & oText.readLine() & Chr(10)
    Loop

    oText.closeInput()
    WMSINST_ReadText=sText
End Function

Function WMSINST_FindPackageFolder() As String
    Dim baseFolder As String
    Dim candidate As String
    Dim oSFA As Object

    baseFolder=WMSINST_BaseFolder()
    oSFA=CreateUnoService("com.sun.star.ucb.SimpleFileAccess")

    candidate=baseFolder & WMSINST_FOLDER
    If oSFA.exists(ConvertToURL(candidate)) Then
        WMSINST_FindPackageFolder=candidate
        Exit Function
    End If

    WMSINST_FindPackageFolder=""
End Function

Function WMSINST_TimeStamp() As String
    Dim d As Date
    d=Now
    WMSINST_TimeStamp=Right("0000" & CStr(Year(d)),4) & _
        Right("00" & CStr(Month(d)),2) & Right("00" & CStr(Day(d)),2) & "_" & _
        Right("00" & CStr(Hour(d)),2) & Right("00" & CStr(Minute(d)),2) & _
        Right("00" & CStr(Second(d)),2)
End Function

Function WMSINST_LastSepPos(s As String,sep As String) As Long
    Dim i As Long
    WMSINST_LastSepPos=0
    If sep="" Then Exit Function
    For i=Len(s) To 1 Step -1
        If Mid(s,i,Len(sep))=sep Then
            WMSINST_LastSepPos=i
            Exit Function
        End If
    Next i
End Function

Function WMSINST_BaseFolder() As String
    Dim localPath As String
    Dim sep As String
    Dim pos As Long

    localPath=ConvertFromURL(ThisComponent.URL)
    sep=WMSINST_PathSep()
    pos=WMSINST_LastSepPos(localPath,sep)

    If pos>0 Then
        WMSINST_BaseFolder=Left(localPath,pos)
    Else
        WMSINST_BaseFolder=""
    End If
End Function

Function WMSINST_PathSep() As String
    Dim localPath As String
    localPath=ConvertFromURL(ThisComponent.URL)
    If InStr(localPath,"\")>0 Then
        WMSINST_PathSep="\"
    Else
        WMSINST_PathSep="/"
    End If
End Function

Sub WMSINST_EnsureFolder(folderPath As String)
    Dim oSFA As Object
    oSFA=CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If Not oSFA.exists(ConvertToURL(folderPath)) Then
        oSFA.createFolder(ConvertToURL(folderPath))
    End If
End Sub

Function WMSINST_FileBaseName(fileURL As String) As String
    Dim s As String
    Dim p As Long
    s=ConvertFromURL(fileURL)
    p=WMSINST_LastSepPos(s,WMSINST_PathSep())
    If p>0 Then s=Mid(s,p+1)
    If LCase(Right(s,4))=".bas" Then s=Left(s,Len(s)-4)
    WMSINST_FileBaseName=s
End Function

Function WMSINST_FileNameFromURL(fileURL As String) As String
    Dim s As String
    Dim p As Long
    s=ConvertFromURL(fileURL)
    p=WMSINST_LastSepPos(s,WMSINST_PathSep())
    If p>0 Then
        WMSINST_FileNameFromURL=Mid(s,p+1)
    Else
        WMSINST_FileNameFromURL=s
    End If
End Function

Function WMSINST_Normalize(s As String) As String
    s=Replace(s,Chr(13),"")
    WMSINST_Normalize=Trim(s)
End Function

Function WMSINST_SourceVersion(src As String) As String
    Dim p As Long
    Dim q As Long
    Dim s As String

    WMSINST_SourceVersion="?"
    p=InStr(src,"VERSION")
    If p=0 Then Exit Function
    s=Mid(src,p,300)
    p=InStr(s,"""")
    If p=0 Then Exit Function
    q=InStr(p+1,s,"""")
    If q=0 Then Exit Function
    WMSINST_SourceVersion=Mid(s,p+1,q-p-1)
End Function
