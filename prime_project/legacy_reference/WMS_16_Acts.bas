Option Explicit

Global Const WMSACT_VERSION = "3.2.0-INTERNAL-COMPACT-ACTS"

Function WMSACT_AskAndCreate(directionName As String, partyFrom As String, partyTo As String, docTitle As String, rowsText As String, operationID As String, Optional ByVal senderName As Variant, Optional ByVal receiverName As Variant) As Boolean
    Dim ans As Integer,actNo As String,pdfPath As String,odtPath As String,sErr As String,msgText As String
    On Error GoTo EH
    If IsMissing(senderName) Then senderName=""
    If IsMissing(receiverName) Then receiverName=""
    WMSACT_AskAndCreate=True
    If Not WMSDBX_TryEnter("ACT_CREATE") Then
        MsgBox "Другая операция WMS ещё выполняется. Дождитесь завершения.",48,"WMS — Акт"
        WMSACT_AskAndCreate=False
        Exit Function
    End If
    ans=MsgBox("Создать внутренний акт передачи ТМЦ?" & Chr(10) & Chr(10) & _
               "Акт будет сохранён в формате ODT. После создания можно открыть его в LibreOffice Writer.",36,"WMS — Акт")
    If ans<>6 Then
        WMSDBX_Leave
        Exit Function
    End If
    senderName=InputBox("ФИО передавшего для строки подписи. Пусто — заполнить от руки.","Акт — Передал",senderName)
    receiverName=InputBox("ФИО принявшего для строки подписи. Пусто — заполнить от руки.","Акт — Принял",receiverName)
    If Not WMSACTREG_CreateAct(directionName,partyFrom,partyTo,docTitle,rowsText,operationID,actNo,pdfPath,odtPath,sErr,senderName,receiverName) Then
        msgText="Акт создать полностью не удалось:" & Chr(10) & sErr
        If actNo<>"" Then msgText=msgText & Chr(10) & "Номер: " & actNo
        WMSDBX_Leave
        MsgBox msgText,48,"WMS — Акт"
        WMSACT_AskAndCreate=False
        Exit Function
    End If
    WMSDBX_Leave
    ans=MsgBox("Акт сохранён: " & actNo & Chr(10) & odtPath & Chr(10) & Chr(10) & _
        "Открыть его в LibreOffice Writer?",36,"WMS — Акт")
    If ans=6 Then WMSACTREG_OpenPath odtPath
    Exit Function
EH:
    msgText=CStr(Err) & " " & Error$
    WMSDBX_Leave
    WMSACT_AskAndCreate=False
    MsgBox "Создание/открытие акта не завершено: " & msgText,48,"WMS — Акт"
End Function


Function WMSACT_CreateFiles(actNo As String, partyFrom As String, partyTo As String, docTitle As String, rowsText As String, ByRef pdfPath As String, ByRef odtPath As String, ByRef sErr As String, Optional ByVal senderName As Variant, Optional ByVal receiverName As Variant) As Boolean
    Dim oDesktop As Object
    Dim oTextDoc As Object
    Dim oText As Object
    Dim oCursor As Object
    Dim oTable As Object
    Dim lines As Variant
    Dim parts As Variant
    Dim i As Long
    Dim nRows As Long
    Dim rr As Long
    Dim baseDir As String
    Dim docsDir As String
    Dim actsDir As String
    Dim yearDir As String
    Dim fileBase As String
    Dim argsODT(1) As New com.sun.star.beans.PropertyValue
    Dim loadArgs(0) As New com.sun.star.beans.PropertyValue

    If IsMissing(senderName) Then senderName=""
    If IsMissing(receiverName) Then receiverName=""
    WMSACT_CreateFiles = False
    pdfPath = ""
    odtPath = ""
    sErr = ""
    On Error GoTo EH

    baseDir = WMSACT_BaseFolder()
    If baseDir = "" Then
        sErr = "Сначала сохраните WMS как .ods."
        Exit Function
    End If

    docsDir = baseDir & "Documents" & WMSACT_PathSep()
    actsDir = docsDir & "Acts" & WMSACT_PathSep()
    yearDir = actsDir & Format(Date, "YYYY") & WMSACT_PathSep()

    WMSACT_EnsureFolder docsDir
    WMSACT_EnsureFolder actsDir
    WMSACT_EnsureFolder yearDir

    fileBase = WMSACT_ActFileBase(actNo)
    pdfPath = ""
    odtPath = yearDir & fileBase & ".odt"

    oDesktop = CreateUnoService("com.sun.star.frame.Desktop")
    loadArgs(0).Name="Hidden":loadArgs(0).Value=True
    oTextDoc = oDesktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, loadArgs())
    oText = oTextDoc.Text
    oCursor = oText.createTextCursor()
    oCursor.CharWeight = 150
    oCursor.CharHeight = 16

    oText.insertString(oCursor, docTitle & Chr(13), False)
    oCursor.CharWeight = 100
    oCursor.CharHeight = 11
    oText.insertString(oCursor, "№ " & actNo & " от " & Format(Date, "DD.MM.YYYY") & Chr(13), False)
    oText.insertString(oCursor, "Передающая сторона: " & partyFrom & Chr(13), False)
    oText.insertString(oCursor, "Принимающая сторона: " & partyTo & Chr(13) & Chr(13), False)
    oText.insertString(oCursor, "Настоящим подтверждается фактическая передача и прием следующих товарно-материальных ценностей:" & Chr(13) & Chr(13), False)

    lines = Split(rowsText, Chr(10))
    nRows = 1
    For i = 0 To UBound(lines)
        If Trim(CStr(lines(i))) <> "" Then
            nRows = nRows + 1
        End If
    Next i

    oTable = oTextDoc.createInstance("com.sun.star.text.TextTable")
    oTable.initialize(nRows, 5)
    oText.insertTextContent(oCursor, oTable, False)

    WMSACT_Cell oTable, "A1", "№"
    WMSACT_Cell oTable, "B1", "Наименование"
    WMSACT_Cell oTable, "C1", "Артикул / код"
    WMSACT_Cell oTable, "D1", "Количество"
    WMSACT_Cell oTable, "E1", "Ед. изм."

    rr = 2
    For i = 0 To UBound(lines)
        If Trim(CStr(lines(i))) <> "" Then
            parts = Split(CStr(lines(i)), Chr(9))
            WMSACT_Cell oTable, "A" & CStr(rr), CStr(rr - 1)
            If UBound(parts) >= 0 Then
                WMSACT_Cell oTable, "B" & CStr(rr), CStr(parts(0))
            End If
            If UBound(parts) >= 1 Then
                WMSACT_Cell oTable, "C" & CStr(rr), CStr(parts(1))
            End If
            If UBound(parts) >= 2 Then
                WMSACT_Cell oTable, "D" & CStr(rr), CStr(parts(2))
            End If
            If UBound(parts) >= 3 Then
                WMSACT_Cell oTable, "E" & CStr(rr), CStr(parts(3))
            End If
            rr = rr + 1
        End If
    Next i

    oCursor = oText.createTextCursorByRange(oText.End)
    oText.insertString(oCursor, Chr(13) & Chr(13) & _
        "Состояние / замечания: ______________________________________________" & Chr(13) & Chr(13) & _
        "Передал: " & IIf(Trim(senderName)="","______________________________",senderName) & "    Подпись: __________________" & Chr(13) & Chr(13) & _
        "Принял: " & IIf(Trim(receiverName)="","______________________________",receiverName) & "    Подпись: __________________" & Chr(13), False)

    argsODT(0).Name = "FilterName"
    argsODT(0).Value = "writer8"
    argsODT(1).Name = "Overwrite"
    argsODT(1).Value = True
    oTextDoc.storeAsURL(ConvertToURL(odtPath), argsODT())

    oTextDoc.close(True)
    WMSACT_CreateFiles = True
    Exit Function

EH:
    sErr = "Создание файлов акта: " & CStr(Err) & " " & Error$
    On Error Resume Next
    If Not IsEmpty(oTextDoc) Then
        oTextDoc.close(True)
    End If
End Function

Sub WMSACT_CreateWriter(partyFrom As String, partyTo As String, docTitle As String, rowsText As String, operationID As String, exportPDF As Boolean)
    Dim pdfPath As String
    Dim odtPath As String
    Dim sErr As String

    If Not WMSACT_CreateFiles(operationID, partyFrom, partyTo, docTitle, rowsText, pdfPath, odtPath, sErr) Then
        MsgBox sErr, 48, "WMS — Акт"
        Exit Sub
    End If

    WMSACTREG_OpenPath odtPath
End Sub

Function WMSACT_PathSep() As String
    Dim p As String

    p = ConvertFromURL(ThisComponent.URL)
    If InStr(p, "\") > 0 Then
        WMSACT_PathSep = "\"
    Else
        WMSACT_PathSep = "/"
    End If
End Function

Function WMSACT_LastSepPos(s As String,sep As String) As Long
    Dim i As Long
    WMSACT_LastSepPos=0
    If sep="" Then Exit Function
    For i=Len(s) To 1 Step -1
        If Mid(s,i,Len(sep))=sep Then
            WMSACT_LastSepPos=i
            Exit Function
        End If
    Next i
End Function

Function WMSACT_BaseFolder() As String
    Dim p As String
    Dim sep As String
    Dim n As Long

    p = ConvertFromURL(ThisComponent.URL)
    sep = WMSACT_PathSep()
    n = WMSACT_LastSepPos(p, sep)

    If n > 0 Then
        WMSACT_BaseFolder = Left(p, n)
    Else
        WMSACT_BaseFolder = ""
    End If
End Function

Sub WMSACT_EnsureFolder(p As String)
    Dim sfa As Object

    On Error Resume Next
    sfa = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If Not sfa.exists(ConvertToURL(p)) Then
        sfa.createFolder(ConvertToURL(p))
    End If
    On Error GoTo 0
End Sub

Sub WMSACT_Cell(oTable As Object, cellName As String, valueText As String)
    On Error Resume Next
    oTable.getCellByName(cellName).String = valueText
End Sub

Function WMSACT_ActFileBase(actNo As String) As String
    Dim t As String,i As Long,ch As String,out As String
    t=Replace(actNo,"АКТ-ПР-","ACT-IN-")
    t=Replace(t,"АКТ-ВЫД-","ACT-OUT-")
    For i=1 To Len(t)
        ch=Mid(t,i,1)
        If (ch>="A" And ch<="Z") Or (ch>="a" And ch<="z") Or _
           (ch>="0" And ch<="9") Or ch="-" Or ch="_" Then
            out=out & ch
        End If
    Next i
    If Trim(out)="" Then
        out="ACT-" & Right("000000" & CStr(CLng(Timer)),6)
    End If
    WMSACT_ActFileBase=out
End Function

Function WMSACT_SafeFileName(s As String) As String
    Dim bad As Variant
    Dim i As Long

    bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = 0 To UBound(bad)
        s = Replace(s, CStr(bad(i)), "_")
    Next i
    WMSACT_SafeFileName = s
End Function

Function WMSACT_OrderRowsText(oSh As Object, receiptEventID As String) As String
    Dim cEvent As Long
    Dim lastR As Long
    Dim r As Long
    Dim s As String
    Dim articleText As String

    cEvent = WMSRC_GetOrCreateTechCol(oSh, "_WMS_ReceiptEventID")
    lastR = WMSRC_LastBusinessRow(oSh)

    For r = 1 To lastR
        If Trim(oSh.getCellByPosition(cEvent, r).String) = Trim(receiptEventID) Then
            articleText = Trim(oSh.getCellByPosition(5, r).String)
            If articleText = "" Then
                articleText = oSh.getCellByPosition(21, r).String
            End If

            s = s & oSh.getCellByPosition(1, r).String & Chr(9) & _
                articleText & Chr(9) & _
                CStr(oSh.getCellByPosition(6, r).Value) & Chr(9) & _
                oSh.getCellByPosition(8, r).String & Chr(10)
        End If
    Next r

    WMSACT_OrderRowsText = s
End Function


Function WMSACT_OrderRangeRowsText(oSh As Object, firstRow As Long, lastRow As Long) As String
    Dim r As Long
    Dim s As String
    Dim articleText As String

    For r=firstRow To lastRow
        If Trim(oSh.getCellByPosition(1,r).String)<>"" Or Trim(oSh.getCellByPosition(21,r).String)<>"" Then
            articleText=Trim(oSh.getCellByPosition(5,r).String)
            If articleText="" Then
                articleText=Trim(oSh.getCellByPosition(21,r).String)
            End If

            s=s & oSh.getCellByPosition(1,r).String & Chr(9) & _
                articleText & Chr(9) & _
                CStr(oSh.getCellByPosition(6,r).Value) & Chr(9) & _
                oSh.getCellByPosition(8,r).String & Chr(10)
        End If
    Next r

    WMSACT_OrderRangeRowsText=s
End Function

Function WMSACT_ShouldOfferReceipt(oSh As Object, r As Long) As Boolean
    Dim modeName As String
    Dim sourceName As String

    WMSACT_ShouldOfferReceipt=False
    On Error GoTo Done

    modeName=UCase(Trim(WMSDBO_TechText(oSh,r,"_WMS_ReceiptMode")))
    sourceName=LCase(Trim(oSh.getCellByPosition(23,r).String))
    If sourceName="" Then
        sourceName=LCase(Trim(WMSDBO_TechText(oSh,r,"_WMS_ReceiptSource")))
    End If

    If modeName="PRODUCTION" Or modeName="OFFICE" Or modeName="PARTS" Then
        WMSACT_ShouldOfferReceipt=True
        Exit Function
    End If

    If InStr(sourceName,"производ")>0 Or InStr(sourceName,"офис")>0 Or InStr(sourceName,"детал")>0 Then
        WMSACT_ShouldOfferReceipt=True
    End If
Done:
End Function

Function WMSACT_ReceiptPartyFrom(oSh As Object, r As Long) As String
    Dim s As String

    On Error GoTo Done
    s=Trim(WMSDBO_TechText(oSh,r,"_WMS_ReceivedFrom"))
    If s="" Then s=Trim(oSh.getCellByPosition(11,r).String)
    If s="" Then s=Trim(oSh.getCellByPosition(23,r).String)
    If s="" Then s="Передающая сторона"
Done:
    WMSACT_ReceiptPartyFrom=s
End Function

Function WMSACT_ReceiptOperationID(oSh As Object, r As Long) As String
    Dim s As String

    On Error GoTo Done
    s=Trim(WMSDBO_TechText(oSh,r,"_WMS_ReceiptEventID"))
    If s="" Then s=Trim(WMSDBO_TechText(oSh,r,"_WMS_SourceID"))
    If s="" Then s="RECEIPT-" & Format(Now,"YYYYMMDD-HHMMSS")
Done:
    WMSACT_ReceiptOperationID=s
End Function

Function WMSACT_IssueRowText(oSh As Object, r As Long) As String
    WMSACT_IssueRowText = oSh.getCellByPosition(2, r).String & Chr(9) & _
        oSh.getCellByPosition(1, r).String & Chr(9) & _
        CStr(oSh.getCellByPosition(3, r).Value) & Chr(9) & _
        oSh.getCellByPosition(4, r).String
End Function
