Option Explicit

Global Const WMSACTREG_VERSION = "3.2.0-ACT-DB-ISOLATION"

Sub WMSACTREG_Install()
    Dim oCon As Object
    Dim sErr As String

    oCon = WMSDB_GetConnectionEx(ThisComponent, sErr)
    If sErr <> "" Then
        MsgBox sErr, 16, "WMS — Реестр актов"
        Exit Sub
    End If

    If Not WMSACTREG_EnsureSchema(oCon, sErr) Then
        WMSDB_Close
        MsgBox sErr, 16, "WMS — Реестр актов"
        Exit Sub
    End If

    On Error Resume Next
    oCon.commit()
    On Error GoTo 0

    If Not WMSDB_SaveDatabaseDocument(sErr) Then
        WMSDB_Close
        MsgBox sErr, 16, "WMS — Реестр актов"
        Exit Sub
    End If

    WMSDB_Close
    MsgBox "Реестр актов установлен." & Chr(10) & _
           "Нумерация, история, позиции и восстановление PDF/ODT готовы.", 64, "WMS — Реестр актов"
End Sub

Function WMSACTREG_EnsureSchema(oCon As Object, ByRef sErr As String) As Boolean
    Dim oStmt As Object

    WMSACTREG_EnsureSchema = False
    sErr = ""
    On Error GoTo EH

    oStmt = oCon.createStatement()

    If Not WMSACTREG_TableExists(oCon, "WMS_ACTS") Then
        oStmt.execute("CREATE TABLE WMS_ACTS (" & _
            "ACT_ID VARCHAR(128) NOT NULL PRIMARY KEY," & _
            "ACT_NO VARCHAR(80) NOT NULL," & _
            "ACT_TYPE VARCHAR(20) NOT NULL," & _
            "ACT_DATE DATE NOT NULL," & _
            "PARTY_FROM VARCHAR(200),PARTY_TO VARCHAR(200)," & _
            "SOURCE_TYPE VARCHAR(40),SOURCE_ID VARCHAR(128),DOC_TITLE VARCHAR(200)," & _
            "FILE_PDF VARCHAR(1000),FILE_ODT VARCHAR(1000)," & _
            "STATUS_NAME VARCHAR(40) NOT NULL,CANCEL_REASON VARCHAR(500)," & _
            "CREATED_AT TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL,UPDATED_AT TIMESTAMP)")
        oStmt.execute("CREATE UNIQUE INDEX UX_WMS_ACTS_NO ON WMS_ACTS (ACT_NO)")
    End If

    If Not WMSACTREG_TableExists(oCon, "WMS_ACT_LINES") Then
        oStmt.execute("CREATE TABLE WMS_ACT_LINES (" & _
            "ACT_ID VARCHAR(128) NOT NULL,LINE_NO INTEGER NOT NULL," & _
            "PRODUCT_NAME VARCHAR(300),ARTICLE_CODE VARCHAR(160),QTY DECIMAL(18,6),UNIT_NAME VARCHAR(64)," & _
            "PRODUCT_CODE VARCHAR(120),LOT_ID VARCHAR(128),LOCATION_NAME VARCHAR(160)," & _
            "PRIMARY KEY (ACT_ID,LINE_NO),FOREIGN KEY (ACT_ID) REFERENCES WMS_ACTS(ACT_ID))")
    End If

    If Not WMSACTREG_TableExists(oCon, "WMS_ACT_COUNTERS") Then
        oStmt.execute("CREATE TABLE WMS_ACT_COUNTERS (COUNTER_KEY VARCHAR(30) NOT NULL PRIMARY KEY,LAST_NO INTEGER NOT NULL)")
    End If

    WMSACTREG_EnsureSchema = True
    Exit Function

EH:
    sErr = "Acts schema: " & CStr(Err) & " " & Error$
End Function

Function WMSACTREG_TableExists(oCon As Object, t As String) As Boolean
    Dim oStmt As Object
    Dim oRS As Object

    WMSACTREG_TableExists = False
    On Error GoTo Done

    oStmt = oCon.createStatement()
    oRS = oStmt.executeQuery("SELECT COUNT(*) FROM RDB$RELATIONS WHERE TRIM(RDB$RELATION_NAME)=" & _
        WMSDB_SQLText(UCase(t)) & " AND COALESCE(RDB$SYSTEM_FLAG,0)=0")

    If oRS.next() Then
        WMSACTREG_TableExists = (oRS.getInt(1) > 0)
    End If

Done:
End Function

Function WMSACTREG_NextNoCon(oCon As Object, directionName As String, ByRef sErr As String) As String
    Dim oStmt As Object
    Dim oRS As Object
    Dim k As String
    Dim prefix As String
    Dim n As Long

    WMSACTREG_NextNoCon = ""
    sErr = ""
    On Error GoTo EH

    If UCase(Trim(directionName)) = "OUT" Then
        prefix = "АКТ-ВЫД-"
    Else
        prefix = "АКТ-ПР-"
    End If

    k = UCase(Trim(directionName)) & "-" & Format(Date, "YYYY")
    oStmt = oCon.createStatement()
    oRS = oStmt.executeQuery("SELECT LAST_NO FROM WMS_ACT_COUNTERS WHERE COUNTER_KEY=" & WMSDB_SQLText(k))

    If oRS.next() Then
        n = oRS.getInt(1) + 1
        oStmt.executeUpdate("UPDATE WMS_ACT_COUNTERS SET LAST_NO=" & CStr(n) & _
            " WHERE COUNTER_KEY=" & WMSDB_SQLText(k))
    Else
        n = 1
        oStmt.executeUpdate("INSERT INTO WMS_ACT_COUNTERS (COUNTER_KEY,LAST_NO) VALUES (" & _
            WMSDB_SQLText(k) & ",1)")
    End If

    WMSACTREG_NextNoCon = prefix & Format(Date, "YYYY") & "-" & Right("000000" & CStr(n), 6)
    Exit Function

EH:
    sErr = "Нумерация актов: " & CStr(Err) & " " & Error$
End Function

Function WMSACTREG_CreateAct(directionName As String, partyFrom As String, partyTo As String, docTitle As String, rowsText As String, sourceID As String, ByRef actNo As String, ByRef pdfPath As String, ByRef odtPath As String, ByRef sErr As String, Optional ByVal senderName As Variant, Optional ByVal receiverName As Variant) As Boolean
    Dim oCon As Object,oStmt As Object
    Dim actID As String,sourceType As String
    Dim lines As Variant,parts As Variant
    Dim i As Long,lineNo As Long,q As Double
    Dim sSaveErr As String

    If IsMissing(senderName) Then senderName=""
    If IsMissing(receiverName) Then receiverName=""
    WMSACTREG_CreateAct=False
    actNo="":pdfPath="":odtPath="":sErr=""
    On Error GoTo EH

    ' ФАЗА 1. Только короткая операция БД: резервируем номер и сразу освобождаем Base/SDBC.
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then Exit Function
    If Not WMSACTREG_EnsureSchema(oCon,sErr) Then GoTo FailDB
    actNo=WMSACTREG_NextNoCon(oCon,directionName,sErr)
    If sErr<>"" Then GoTo FailDB
    oCon.commit()
    If Not WMSDB_SaveDatabaseDocument(sSaveErr) Then
        sErr="Не удалось сохранить счётчик акта: " & sSaveErr
        GoTo FailDB
    End If
    WMSDB_Close

    ' ФАЗА 2. Writer/PDF работают без открытого Firebird connection.
    If Not WMSACT_CreateFiles(actNo,partyFrom,partyTo,docTitle,rowsText,pdfPath,odtPath,sErr,senderName,receiverName) Then Exit Function

    ' ФАЗА 3. Коротко открываем БД и регистрируем уже созданные файлы.
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then Exit Function
    actID=WMSACTREG_NewID("ACT")
    If UCase(Trim(directionName))="OUT" Then
        sourceType="ISSUE"
    Else
        sourceType="RECEIPT"
    End If
    oStmt=oCon.createStatement()
    oStmt.executeUpdate("INSERT INTO WMS_ACTS (ACT_ID,ACT_NO,ACT_TYPE,ACT_DATE,PARTY_FROM,PARTY_TO,SOURCE_TYPE,SOURCE_ID,DOC_TITLE,FILE_PDF,FILE_ODT,STATUS_NAME) VALUES (" & _
        WMSDB_SQLText(actID) & "," & WMSDB_SQLText(actNo) & "," & WMSDB_SQLText(directionName) & ",CURRENT_DATE," & _
        WMSDB_SQLText(partyFrom) & "," & WMSDB_SQLText(partyTo) & "," & WMSDB_SQLText(sourceType) & "," & _
        WMSDB_SQLText(sourceID) & "," & WMSDB_SQLText(docTitle) & "," & WMSDB_SQLText(pdfPath) & "," & WMSDB_SQLText(odtPath) & ",'ACTIVE')")

    lines=Split(rowsText,Chr(10)):lineNo=0
    For i=0 To UBound(lines)
        If Trim(CStr(lines(i)))<>"" Then
            lineNo=lineNo+1:parts=Split(CStr(lines(i)),Chr(9)):q=0
            On Error Resume Next
            If UBound(parts)>=2 Then q=CDbl(parts(2))
            On Error GoTo EH
            oStmt.executeUpdate("INSERT INTO WMS_ACT_LINES (ACT_ID,LINE_NO,PRODUCT_NAME,ARTICLE_CODE,QTY,UNIT_NAME) VALUES (" & _
                WMSDB_SQLText(actID) & "," & CStr(lineNo) & "," & WMSDB_SQLText(WMSACTREG_Part(parts,0)) & "," & _
                WMSDB_SQLText(WMSACTREG_Part(parts,1)) & "," & WMSACTREG_Num(q) & "," & WMSDB_SQLText(WMSACTREG_Part(parts,3)) & ")")
        End If
    Next i
    On Error Resume Next
    WMSSAFE_AuditCon oCon,"ACT_CREATED","ACT",actID,"DONE",actNo & " | " & sourceType & " | " & sourceID
    On Error GoTo EH
    oCon.commit()
    WMSDBX_CloseStmt oStmt
    If Not WMSDB_SaveDatabaseDocument(sSaveErr) Then
        sErr="Акт записан в Firebird, но ODB не удалось сохранить: " & sSaveErr
        WMSDB_Close
        Exit Function
    End If
    WMSDB_Close
    WMSACTREG_CreateAct=True
    Exit Function

FailDB:
    WMSDBX_RollbackQuiet oCon
    WMSDBX_CloseStmt oStmt
    On Error Resume Next
    WMSDB_Close
    On Error GoTo 0
    Exit Function
EH:
    sErr="Создание акта: " & CStr(Err) & " " & Error$
    Resume FailDB
End Function


Function WMSACTREG_Part(parts As Variant, idx As Long) As String
    WMSACTREG_Part = ""
    On Error GoTo Done

    If idx <= UBound(parts) Then
        WMSACTREG_Part = CStr(parts(idx))
    End If

Done:
End Function

Function WMSACTREG_Num(v As Double) As String
    WMSACTREG_Num = Replace(CStr(v), ",", ".")
End Function

Sub WMSACTREG_ShowRecent()
    Dim oCon As Object
    Dim oStmt As Object
    Dim oRS As Object
    Dim sErr As String
    Dim s As String
    Dim n As Long

    oCon = WMSDB_GetConnectionEx(ThisComponent, sErr)
    If sErr <> "" Then
        MsgBox sErr, 16, "WMS — Акты"
        Exit Sub
    End If

    If Not WMSACTREG_EnsureSchema(oCon, sErr) Then
        WMSDB_Close
        MsgBox sErr, 16, "WMS — Акты"
        Exit Sub
    End If

    oStmt = oCon.createStatement()
    oRS = oStmt.executeQuery("SELECT FIRST 30 ACT_NO,ACT_DATE,PARTY_FROM,PARTY_TO,STATUS_NAME,SOURCE_ID FROM WMS_ACTS ORDER BY CREATED_AT DESC")

    Do While oRS.next()
        n = n + 1
        s = s & oRS.getString(1) & " | " & oRS.getString(2) & " | " & oRS.getString(5) & Chr(10) & _
            oRS.getString(3) & " -> " & oRS.getString(4) & Chr(10) & _
            "Операция: " & oRS.getString(6) & Chr(10) & Chr(10)
    Loop

    WMSDB_Close

    If n = 0 Then
        s = "Актов пока нет."
    End If

    MsgBox s, 64, "WMS — Последние акты"
End Sub

Sub WMSACTREG_OpenByNumber()
    Dim actNo As String
    Dim oCon As Object
    Dim oStmt As Object
    Dim oRS As Object
    Dim sErr As String
    Dim p As String

    actNo = InputBox("Номер акта:", "WMS — Открыть акт", "")
    If Trim(actNo) = "" Then
        Exit Sub
    End If

    oCon = WMSDB_GetConnectionEx(ThisComponent, sErr)
    If sErr <> "" Then
        MsgBox sErr, 16, "WMS — Открыть акт"
        Exit Sub
    End If

    oStmt = oCon.createStatement()
    oRS = oStmt.executeQuery("SELECT FILE_PDF,FILE_ODT FROM WMS_ACTS WHERE ACT_NO=" & WMSDB_SQLText(Trim(actNo)))

    If Not oRS.next() Then
        WMSDB_Close
        MsgBox "Акт не найден.", 48, "WMS — Открыть акт"
        Exit Sub
    End If

    p = Trim(oRS.getString(2))
    If p = "" Then
        p = Trim(oRS.getString(1))
    End If

    WMSDB_Close

    If p = "" Then
        MsgBox "Файл акта не сохранён. Используйте пересоздание.", 48, "WMS — Открыть акт"
        Exit Sub
    End If

    WMSACTREG_OpenPath p
End Sub

Sub WMSACTREG_RebuildByNumber()
    Dim actNo As String
    Dim oCon As Object
    Dim oStmt As Object
    Dim oRS As Object
    Dim oLines As Object
    Dim sErr As String
    Dim partyFrom As String
    Dim partyTo As String
    Dim title As String
    Dim rowsText As String
    Dim pdfPath As String
    Dim odtPath As String
    Dim actID As String

    actNo = InputBox("Номер акта для пересоздания:", "WMS — Пересоздать акт", "")
    If Trim(actNo) = "" Then
        Exit Sub
    End If

    oCon = WMSDB_GetConnectionEx(ThisComponent, sErr)
    If sErr <> "" Then
        MsgBox sErr, 16, "WMS — Пересоздать акт"
        Exit Sub
    End If

    oStmt = oCon.createStatement()
    oRS = oStmt.executeQuery("SELECT ACT_ID,PARTY_FROM,PARTY_TO,DOC_TITLE FROM WMS_ACTS WHERE ACT_NO=" & WMSDB_SQLText(Trim(actNo)))

    If Not oRS.next() Then
        WMSDB_Close
        MsgBox "Акт не найден.", 48, "WMS — Пересоздать акт"
        Exit Sub
    End If

    actID = oRS.getString(1)
    partyFrom = oRS.getString(2)
    partyTo = oRS.getString(3)
    title = oRS.getString(4)

    oLines = oStmt.executeQuery("SELECT PRODUCT_NAME,ARTICLE_CODE,QTY,UNIT_NAME FROM WMS_ACT_LINES WHERE ACT_ID=" & _
        WMSDB_SQLText(actID) & " ORDER BY LINE_NO")

    Do While oLines.next()
        rowsText = rowsText & oLines.getString(1) & Chr(9) & oLines.getString(2) & Chr(9) & _
            CStr(oLines.getDouble(3)) & Chr(9) & oLines.getString(4) & Chr(10)
    Loop

    If WMSACT_CreateFiles(actNo, partyFrom, partyTo, title, rowsText, pdfPath, odtPath, sErr) Then
        oStmt.executeUpdate("UPDATE WMS_ACTS SET FILE_PDF=" & WMSDB_SQLText(pdfPath) & _
            ",FILE_ODT=" & WMSDB_SQLText(odtPath) & _
            ",STATUS_NAME='ACTIVE',UPDATED_AT=CURRENT_TIMESTAMP WHERE ACT_ID=" & WMSDB_SQLText(actID))
        oCon.commit()
        WMSDB_Close
        MsgBox "Акт пересоздан: " & actNo, 64, "WMS — Пересоздать акт"
    Else
        WMSDB_Close
        MsgBox sErr, 16, "WMS — Пересоздать акт"
    End If
End Sub

Sub WMSACTREG_CancelByNumber()
    Dim actNo As String
    Dim reason As String
    Dim oCon As Object
    Dim oStmt As Object
    Dim oRS As Object
    Dim sErr As String
    Dim actID As String

    actNo = InputBox("Номер акта:", "WMS — Аннулировать акт", "")
    If Trim(actNo) = "" Then
        Exit Sub
    End If

    reason = InputBox("Причина аннулирования:", "WMS — Аннулировать акт", "")
    If Trim(reason) = "" Then
        Exit Sub
    End If

    If MsgBox("Аннулировать " & actNo & "?" & Chr(10) & "История и файлы останутся.", 36, "WMS — Аннулировать акт") <> 6 Then
        Exit Sub
    End If

    oCon = WMSDB_GetConnectionEx(ThisComponent, sErr)
    If sErr <> "" Then
        MsgBox sErr, 16, "WMS — Аннулировать акт"
        Exit Sub
    End If

    oStmt = oCon.createStatement()
    oRS = oStmt.executeQuery("SELECT ACT_ID,STATUS_NAME FROM WMS_ACTS WHERE ACT_NO=" & WMSDB_SQLText(Trim(actNo)))

    If Not oRS.next() Then
        WMSDB_Close
        MsgBox "Акт не найден.", 48, "WMS — Аннулировать акт"
        Exit Sub
    End If

    actID = oRS.getString(1)

    If UCase(Trim(oRS.getString(2))) = "CANCELLED" Then
        WMSDB_Close
        MsgBox "Этот акт уже аннулирован.", 48, "WMS — Аннулировать акт"
        Exit Sub
    End If

    oStmt.executeUpdate("UPDATE WMS_ACTS SET STATUS_NAME='CANCELLED',CANCEL_REASON=" & _
        WMSDB_SQLText(reason) & ",UPDATED_AT=CURRENT_TIMESTAMP WHERE ACT_ID=" & WMSDB_SQLText(actID))

    On Error Resume Next
    WMSSAFE_AuditCon oCon, "ACT_CANCELLED", "ACT", actID, "DONE", actNo & " | " & reason
    On Error GoTo 0

    oCon.commit()
    WMSDB_Close
    MsgBox "Акт аннулирован. История сохранена.", 64, "WMS — Аннулировать акт"
End Sub

Sub WMSACTREG_OpenPath(p As String)
    Dim oDesktop As Object

    On Error GoTo EH
    oDesktop = CreateUnoService("com.sun.star.frame.Desktop")
    oDesktop.loadComponentFromURL(ConvertToURL(p), "_blank", 0, Array())
    Exit Sub

EH:
    MsgBox "Не удалось открыть файл:" & Chr(10) & p & Chr(10) & CStr(Err) & " " & Error$, 48, "WMS — Акт"
End Sub

Function WMSACTREG_NewID(prefix As String) As String
    Randomize
    WMSACTREG_NewID = prefix & "-" & Format(Now, "YYYYMMDD-HHMMSS") & "-" & _
        Right("000000" & CStr(CLng(Rnd() * 999999)), 6)
End Function
