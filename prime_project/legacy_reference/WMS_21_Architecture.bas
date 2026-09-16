Option Explicit

Global Const WMSARCH_VERSION = "2.0.0-CORE-SCHEMA"

Function WMSARCH_EnsureSchema(oCon As Object, ByRef sErr As String) As Boolean
    Dim oStmt As Object
    WMSARCH_EnsureSchema=False
    sErr=""
    On Error GoTo EH

    oStmt=oCon.createStatement()

    ' ---------- master data ----------
    If Not WMSDB_TableExistsSimple(oCon,"WMS_REFERENCE",sErr) Then
        If sErr<>"" Then Exit Function
        oStmt.executeUpdate("CREATE TABLE WMS_REFERENCE (" & _
            "REF_TYPE VARCHAR(40) NOT NULL," & _
            "REF_CODE VARCHAR(80) NOT NULL," & _
            "REF_NAME VARCHAR(200) NOT NULL," & _
            "PARENT_CODE VARCHAR(80)," & _
            "ACTIVE_FLAG SMALLINT DEFAULT 1 NOT NULL," & _
            "SORT_ORDER INTEGER DEFAULT 0 NOT NULL," & _
            "NOTE_TEXT VARCHAR(500)," & _
            "UPDATED_AT TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL," & _
            "CONSTRAINT PK_WMS_REFERENCE PRIMARY KEY (REF_TYPE,REF_CODE))")
    End If

    If Not WMSDB_TableExistsSimple(oCon,"WMS_PRODUCT_ALIASES",sErr) Then
        If sErr<>"" Then Exit Function
        oStmt.executeUpdate("CREATE TABLE WMS_PRODUCT_ALIASES (" & _
            "ALIAS_ID VARCHAR(128) NOT NULL PRIMARY KEY," & _
            "PRODUCT_CODE VARCHAR(80) NOT NULL," & _
            "SUPPLIER_NAME VARCHAR(180)," & _
            "SELLER_NAME VARCHAR(180)," & _
            "ARTICLE_CODE VARCHAR(120) NOT NULL," & _
            "ALIAS_TYPE VARCHAR(40) DEFAULT 'SUPPLIER_ARTICLE' NOT NULL," & _
            "UNIT_NAME VARCHAR(40)," & _
            "FACTOR_TO_BASE DECIMAL(18,6) DEFAULT 1 NOT NULL," & _
            "ACTIVE_FLAG SMALLINT DEFAULT 1 NOT NULL," & _
            "CREATED_AT TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL," & _
            "CONSTRAINT FK_WMS_ALIAS_PRODUCT FOREIGN KEY (PRODUCT_CODE) REFERENCES WMS_PRODUCTS(PRODUCT_CODE))")
    End If

    If Not WMSDB_TableExistsSimple(oCon,"WMS_DOCUMENTS",sErr) Then
        If sErr<>"" Then Exit Function
        oStmt.executeUpdate("CREATE TABLE WMS_DOCUMENTS (" & _
            "DOC_ID VARCHAR(128) NOT NULL PRIMARY KEY," & _
            "DOC_TYPE VARCHAR(40) NOT NULL," & _
            "DOC_NO VARCHAR(120)," & _
            "DOC_DATE DATE NOT NULL," & _
            "FLOW_CHANNEL VARCHAR(40)," & _
            "PARTY_FROM VARCHAR(200)," & _
            "PARTY_TO VARCHAR(200)," & _
            "STATUS_NAME VARCHAR(40) DEFAULT 'POSTED' NOT NULL," & _
            "COMMENT_TEXT VARCHAR(1000)," & _
            "CREATED_AT TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL," & _
            "UPDATED_AT TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL)")
    End If

    If Not WMSDB_TableExistsSimple(oCon,"WMS_DOCUMENT_LINES",sErr) Then
        If sErr<>"" Then Exit Function
        oStmt.executeUpdate("CREATE TABLE WMS_DOCUMENT_LINES (" & _
            "DOC_ID VARCHAR(128) NOT NULL," & _
            "LINE_NO INTEGER NOT NULL," & _
            "PRODUCT_CODE VARCHAR(80)," & _
            "ARTICLE_CODE VARCHAR(120)," & _
            "PRODUCT_NAME VARCHAR(255) NOT NULL," & _
            "QTY DECIMAL(18,6) NOT NULL," & _
            "UNIT_NAME VARCHAR(40) NOT NULL," & _
            "LOCATION_NAME VARCHAR(120)," & _
            "LOT_ID VARCHAR(128)," & _
            "RETURNABLE VARCHAR(8)," & _
            "NOTE_TEXT VARCHAR(1000)," & _
            "CONSTRAINT PK_WMS_DOCUMENT_LINES PRIMARY KEY (DOC_ID,LINE_NO)," & _
            "CONSTRAINT FK_WMS_DOCL_DOC FOREIGN KEY (DOC_ID) REFERENCES WMS_DOCUMENTS(DOC_ID))")
    End If

    If Not WMSDB_TableExistsSimple(oCon,"WMS_INVENTORY",sErr) Then
        If sErr<>"" Then Exit Function
        oStmt.executeUpdate("CREATE TABLE WMS_INVENTORY (" & _
            "INVENTORY_ID VARCHAR(128) NOT NULL PRIMARY KEY," & _
            "INVENTORY_DATE DATE NOT NULL," & _
            "STATUS_NAME VARCHAR(40) DEFAULT 'DRAFT' NOT NULL," & _
            "COMMENT_TEXT VARCHAR(1000)," & _
            "CREATED_AT TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL," & _
            "UPDATED_AT TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL)")
    End If

    If Not WMSDB_TableExistsSimple(oCon,"WMS_INVENTORY_LINES",sErr) Then
        If sErr<>"" Then Exit Function
        oStmt.executeUpdate("CREATE TABLE WMS_INVENTORY_LINES (" & _
            "INVENTORY_ID VARCHAR(128) NOT NULL," & _
            "LINE_NO INTEGER NOT NULL," & _
            "LOT_ID VARCHAR(128)," & _
            "PRODUCT_CODE VARCHAR(80) NOT NULL," & _
            "LOCATION_NAME VARCHAR(120) NOT NULL," & _
            "SYSTEM_QTY DECIMAL(18,6) NOT NULL," & _
            "FACT_QTY DECIMAL(18,6)," & _
            "DIFF_QTY DECIMAL(18,6)," & _
            "UNIT_NAME VARCHAR(40) NOT NULL," & _
            "STATUS_NAME VARCHAR(40) DEFAULT 'DRAFT' NOT NULL," & _
            "CONSTRAINT PK_WMS_INV_LINES PRIMARY KEY (INVENTORY_ID,LINE_NO)," & _
            "CONSTRAINT FK_WMS_INVL_INV FOREIGN KEY (INVENTORY_ID) REFERENCES WMS_INVENTORY(INVENTORY_ID))")
    End If

    ' ---------- extend existing tables without changing existing PK/API ----------
    If Not WMSARCH_ColumnExists(oCon,"WMS_PRODUCTS","ITEM_ID",sErr) Then
        If sErr<>"" Then Exit Function
        oStmt.executeUpdate("ALTER TABLE WMS_PRODUCTS ADD ITEM_ID VARCHAR(100)")
    End If

    If Not WMSARCH_ColumnExists(oCon,"WMS_STOCK_MOVEMENTS","STATUS_NAME",sErr) Then
        If sErr<>"" Then Exit Function
        oStmt.executeUpdate("ALTER TABLE WMS_STOCK_MOVEMENTS ADD STATUS_NAME VARCHAR(40)")
    End If
    If Not WMSARCH_ColumnExists(oCon,"WMS_STOCK_MOVEMENTS","FLOW_CHANNEL",sErr) Then
        If sErr<>"" Then Exit Function
        oStmt.executeUpdate("ALTER TABLE WMS_STOCK_MOVEMENTS ADD FLOW_CHANNEL VARCHAR(40)")
    End If
    If Not WMSARCH_ColumnExists(oCon,"WMS_STOCK_MOVEMENTS","OPERATION_ID",sErr) Then
        If sErr<>"" Then Exit Function
        oStmt.executeUpdate("ALTER TABLE WMS_STOCK_MOVEMENTS ADD OPERATION_ID VARCHAR(128)")
    End If
    If Not WMSARCH_ColumnExists(oCon,"WMS_STOCK_MOVEMENTS","REVERSAL_OF",sErr) Then
        If sErr<>"" Then Exit Function
        oStmt.executeUpdate("ALTER TABLE WMS_STOCK_MOVEMENTS ADD REVERSAL_OF VARCHAR(100)")
    End If

    ' ---------- safe backfill ----------
    oStmt.executeUpdate("UPDATE WMS_PRODUCTS SET ITEM_ID='ITEM-' || PRODUCT_CODE WHERE ITEM_ID IS NULL OR TRIM(ITEM_ID)=''")
    oStmt.executeUpdate("UPDATE WMS_STOCK_MOVEMENTS SET STATUS_NAME='POSTED' WHERE STATUS_NAME IS NULL OR TRIM(STATUS_NAME)=''")
    oStmt.executeUpdate("UPDATE WMS_STOCK_MOVEMENTS SET FLOW_CHANNEL=ORIGIN_NAME " & _
        "WHERE (FLOW_CHANNEL IS NULL OR TRIM(FLOW_CHANNEL)='') AND MOVEMENT_TYPE='IN' AND ORIGIN_NAME IS NOT NULL")

    ' ---------- sequence for internal product codes ----------
    If Not WMSARCH_SequenceExists(oCon,"GEN_WMS_ITEM_ID",sErr) Then
        If sErr<>"" Then Exit Function
        oStmt.executeUpdate("CREATE SEQUENCE GEN_WMS_ITEM_ID")
    End If

    ' ---------- indexes ----------
    WMSARCH_CreateIndexIfMissing oCon,"IX_WMS_REF_TYPE","WMS_REFERENCE(REF_TYPE,ACTIVE_FLAG)",False,sErr
    If sErr<>"" Then Exit Function
    WMSARCH_CreateIndexIfMissing oCon,"IX_WMS_ALIAS_ART","WMS_PRODUCT_ALIASES(ARTICLE_CODE,ACTIVE_FLAG)",False,sErr
    If sErr<>"" Then Exit Function
    WMSARCH_CreateIndexIfMissing oCon,"IX_WMS_ALIAS_PRODUCT","WMS_PRODUCT_ALIASES(PRODUCT_CODE)",False,sErr
    If sErr<>"" Then Exit Function
    WMSARCH_CreateIndexIfMissing oCon,"UX_WMS_PRODUCTS_ITEMID","WMS_PRODUCTS(ITEM_ID)",True,sErr
    If sErr<>"" Then Exit Function
    WMSARCH_CreateIndexIfMissing oCon,"IX_WMS_MOV_STATUS","WMS_STOCK_MOVEMENTS(STATUS_NAME,PRODUCT_CODE,LOCATION_NAME)",False,sErr
    If sErr<>"" Then Exit Function
    WMSARCH_CreateIndexIfMissing oCon,"IX_WMS_MOV_FLOW","WMS_STOCK_MOVEMENTS(FLOW_CHANNEL,MOVEMENT_DATE)",False,sErr
    If sErr<>"" Then Exit Function
    WMSARCH_CreateIndexIfMissing oCon,"IX_WMS_MOV_LOT","WMS_STOCK_MOVEMENTS(LOT_ID)",False,sErr
    If sErr<>"" Then Exit Function

    ' ---------- views are the single read model for stock ----------
    oStmt.executeUpdate("CREATE OR ALTER VIEW WMS_V_STOCK_BALANCE " & _
        "(PRODUCT_CODE,PRODUCT_NAME,UNIT_NAME,LOCATION_NAME,ORIGIN_NAME,CATEGORY_NAME,SUBCATEGORY_NAME,QTY) AS " & _
        "SELECT m.PRODUCT_CODE,MAX(m.PRODUCT_NAME),m.UNIT_NAME,m.LOCATION_NAME," & _
        "COALESCE(m.ORIGIN_NAME,''),COALESCE(m.DEST_CATEGORY,''),COALESCE(m.DEST_SUBCATEGORY,''),SUM(m.QTY) " & _
        "FROM WMS_STOCK_MOVEMENTS m WHERE COALESCE(m.STATUS_NAME,'POSTED')='POSTED' " & _
        "GROUP BY m.PRODUCT_CODE,m.UNIT_NAME,m.LOCATION_NAME,m.ORIGIN_NAME,m.DEST_CATEGORY,m.DEST_SUBCATEGORY " & _
        "HAVING SUM(m.QTY)<>0")

    oStmt.executeUpdate("CREATE OR ALTER VIEW WMS_V_LOT_BALANCE " & _
        "(LOT_ID,PRODUCT_CODE,PRODUCT_NAME,UNIT_NAME,LOCATION_NAME,ORIGIN_NAME,CATEGORY_NAME,SUBCATEGORY_NAME,QTY,RECEIPT_DATE) AS " & _
        "SELECT m.LOT_ID,m.PRODUCT_CODE,MAX(m.PRODUCT_NAME),m.UNIT_NAME,m.LOCATION_NAME," & _
        "COALESCE(m.ORIGIN_NAME,''),COALESCE(m.DEST_CATEGORY,''),COALESCE(m.DEST_SUBCATEGORY,''),SUM(m.QTY),MAX(l.RECEIPT_DATE) " & _
        "FROM WMS_STOCK_MOVEMENTS m LEFT JOIN WMS_STOCK_LOTS l ON l.LOT_ID=m.LOT_ID " & _
        "WHERE COALESCE(m.STATUS_NAME,'POSTED')='POSTED' AND m.LOT_ID IS NOT NULL " & _
        "GROUP BY m.LOT_ID,m.PRODUCT_CODE,m.UNIT_NAME,m.LOCATION_NAME,m.ORIGIN_NAME,m.DEST_CATEGORY,m.DEST_SUBCATEGORY " & _
        "HAVING SUM(m.QTY)<>0")

    ' ---------- migrate legacy supplier articles into alias layer ----------
    oStmt.executeUpdate("INSERT INTO WMS_PRODUCT_ALIASES " & _
        "(ALIAS_ID,PRODUCT_CODE,ARTICLE_CODE,ALIAS_TYPE,FACTOR_TO_BASE,ACTIVE_FLAG) " & _
        "SELECT 'LEGACY-' || PRODUCT_CODE,PRODUCT_CODE,SUPPLIER_ARTICLE,'SUPPLIER_ARTICLE',1,1 " & _
        "FROM WMS_PRODUCTS p WHERE p.SUPPLIER_ARTICLE IS NOT NULL AND TRIM(p.SUPPLIER_ARTICLE)<>'' " & _
        "AND NOT EXISTS (SELECT 1 FROM WMS_PRODUCT_ALIASES a WHERE a.PRODUCT_CODE=p.PRODUCT_CODE AND a.ARTICLE_CODE=p.SUPPLIER_ARTICLE)")

    ' ---------- seed reference values from real existing data ----------
    WMSARCH_SeedReferences oCon,sErr
    If sErr<>"" Then Exit Function

    ' ---------- release metadata ----------
    WMSARCH_SetMeta oCon,"RELEASE_MODE","PRODUCTION",sErr
    If sErr<>"" Then Exit Function
    WMSARCH_SetMeta oCon,"RELEASE_VERSION","2.0.0",sErr
    If sErr<>"" Then Exit Function
    WMSARCH_SetMeta oCon,"ARCHITECTURE_VERSION","2.0.0",sErr
    If sErr<>"" Then Exit Function
    WMSARCH_SetMeta oCon,"STOCK_READ_MODEL","WMS_V_STOCK_BALANCE",sErr
    If sErr<>"" Then Exit Function
    WMSARCH_SetMeta oCon,"MASTER_DATA_MODEL","WMS_REFERENCE+WMS_PRODUCT_ALIASES",sErr
    If sErr<>"" Then Exit Function

    oCon.commit()
    WMSARCH_EnsureSchema=True
    Exit Function
EH:
    sErr="WMS Architecture migration: " & CStr(Err) & " " & Error$
End Function

Sub WMSARCH_SeedReferences(oCon As Object, ByRef sErr As String)
    Dim oStmt As Object
    sErr=""
    On Error GoTo EH
    oStmt=oCon.createStatement()

    oStmt.executeUpdate("INSERT INTO WMS_REFERENCE (REF_TYPE,REF_CODE,REF_NAME) " & _
        "SELECT DISTINCT 'UNIT',p.UNIT_NAME,p.UNIT_NAME FROM WMS_PRODUCTS p WHERE p.UNIT_NAME IS NOT NULL AND TRIM(p.UNIT_NAME)<>'' " & _
        "AND NOT EXISTS (SELECT 1 FROM WMS_REFERENCE r WHERE r.REF_TYPE='UNIT' AND r.REF_CODE=p.UNIT_NAME)")
    oStmt.executeUpdate("INSERT INTO WMS_REFERENCE (REF_TYPE,REF_CODE,REF_NAME) " & _
        "SELECT DISTINCT 'LOCATION',m.LOCATION_NAME,m.LOCATION_NAME FROM WMS_STOCK_MOVEMENTS m WHERE m.LOCATION_NAME IS NOT NULL AND TRIM(m.LOCATION_NAME)<>'' " & _
        "AND NOT EXISTS (SELECT 1 FROM WMS_REFERENCE r WHERE r.REF_TYPE='LOCATION' AND r.REF_CODE=m.LOCATION_NAME)")
    oStmt.executeUpdate("INSERT INTO WMS_REFERENCE (REF_TYPE,REF_CODE,REF_NAME) " & _
        "SELECT DISTINCT 'CATEGORY',m.DEST_CATEGORY,m.DEST_CATEGORY FROM WMS_STOCK_MOVEMENTS m WHERE m.DEST_CATEGORY IS NOT NULL AND TRIM(m.DEST_CATEGORY)<>'' " & _
        "AND NOT EXISTS (SELECT 1 FROM WMS_REFERENCE r WHERE r.REF_TYPE='CATEGORY' AND r.REF_CODE=m.DEST_CATEGORY)")
    oStmt.executeUpdate("INSERT INTO WMS_REFERENCE (REF_TYPE,REF_CODE,REF_NAME,PARENT_CODE) " & _
        "SELECT DISTINCT 'SUBCATEGORY',COALESCE(m.DEST_CATEGORY,'') || '>' || m.DEST_SUBCATEGORY,m.DEST_SUBCATEGORY,m.DEST_CATEGORY FROM WMS_STOCK_MOVEMENTS m " & _
        "WHERE m.DEST_SUBCATEGORY IS NOT NULL AND TRIM(m.DEST_SUBCATEGORY)<>'' " & _
        "AND NOT EXISTS (SELECT 1 FROM WMS_REFERENCE r WHERE r.REF_TYPE='SUBCATEGORY' AND r.REF_CODE=COALESCE(m.DEST_CATEGORY,'') || '>' || m.DEST_SUBCATEGORY)")
    oStmt.executeUpdate("INSERT INTO WMS_REFERENCE (REF_TYPE,REF_CODE,REF_NAME) " & _
        "SELECT DISTINCT 'SUPPLIER',o.SUPPLIER_NAME,o.SUPPLIER_NAME FROM WMS_ORDERS o WHERE o.SUPPLIER_NAME IS NOT NULL AND TRIM(o.SUPPLIER_NAME)<>'' " & _
        "AND NOT EXISTS (SELECT 1 FROM WMS_REFERENCE r WHERE r.REF_TYPE='SUPPLIER' AND r.REF_CODE=o.SUPPLIER_NAME)")
    oStmt.executeUpdate("INSERT INTO WMS_REFERENCE (REF_TYPE,REF_CODE,REF_NAME) " & _
        "SELECT DISTINCT 'SELLER',o.SELLER_NAME,o.SELLER_NAME FROM WMS_ORDER_LINES o WHERE o.SELLER_NAME IS NOT NULL AND TRIM(o.SELLER_NAME)<>'' " & _
        "AND NOT EXISTS (SELECT 1 FROM WMS_REFERENCE r WHERE r.REF_TYPE='SELLER' AND r.REF_CODE=o.SELLER_NAME)")
    oStmt.executeUpdate("INSERT INTO WMS_REFERENCE (REF_TYPE,REF_CODE,REF_NAME) " & _
        "SELECT DISTINCT 'EMPLOYEE',i.EMPLOYEE_NAME,i.EMPLOYEE_NAME FROM WMS_ISSUES i WHERE i.EMPLOYEE_NAME IS NOT NULL AND TRIM(i.EMPLOYEE_NAME)<>'' " & _
        "AND NOT EXISTS (SELECT 1 FROM WMS_REFERENCE r WHERE r.REF_TYPE='EMPLOYEE' AND r.REF_CODE=i.EMPLOYEE_NAME)")
    Exit Sub
EH:
    sErr="Заполнение справочников: " & CStr(Err) & " " & Error$
End Sub

Function WMSARCH_ColumnExists(oCon As Object,sTable As String,sColumn As String,ByRef sErr As String) As Boolean
    Dim oStmt As Object,oRS As Object,sql As String
    WMSARCH_ColumnExists=False:sErr=""
    On Error GoTo EH
    sql="SELECT COUNT(*) FROM RDB$RELATION_FIELDS WHERE TRIM(RDB$RELATION_NAME)=" & _
        WMSDB_SQLText(UCase(Trim(sTable))) & " AND TRIM(RDB$FIELD_NAME)=" & WMSDB_SQLText(UCase(Trim(sColumn)))
    oStmt=oCon.createStatement():oRS=oStmt.executeQuery(sql)
    If oRS.next() Then WMSARCH_ColumnExists=(oRS.getLong(1)>0)
    Exit Function
EH:
    sErr="Проверка поля " & sTable & "." & sColumn & ": " & CStr(Err) & " " & Error$
End Function

Function WMSARCH_IndexExists(oCon As Object,sIndex As String,ByRef sErr As String) As Boolean
    Dim oStmt As Object,oRS As Object,sql As String
    WMSARCH_IndexExists=False:sErr=""
    On Error GoTo EH
    sql="SELECT COUNT(*) FROM RDB$INDICES WHERE TRIM(RDB$INDEX_NAME)=" & WMSDB_SQLText(UCase(Trim(sIndex)))
    oStmt=oCon.createStatement():oRS=oStmt.executeQuery(sql)
    If oRS.next() Then WMSARCH_IndexExists=(oRS.getLong(1)>0)
    Exit Function
EH:
    sErr="Проверка индекса " & sIndex & ": " & CStr(Err) & " " & Error$
End Function

Function WMSARCH_SequenceExists(oCon As Object,sName As String,ByRef sErr As String) As Boolean
    Dim oStmt As Object,oRS As Object,sql As String
    WMSARCH_SequenceExists=False:sErr=""
    On Error GoTo EH
    sql="SELECT COUNT(*) FROM RDB$GENERATORS WHERE TRIM(RDB$GENERATOR_NAME)=" & WMSDB_SQLText(UCase(Trim(sName))) & _
        " AND COALESCE(RDB$SYSTEM_FLAG,0)=0"
    oStmt=oCon.createStatement():oRS=oStmt.executeQuery(sql)
    If oRS.next() Then WMSARCH_SequenceExists=(oRS.getLong(1)>0)
    Exit Function
EH:
    sErr="Проверка sequence " & sName & ": " & CStr(Err) & " " & Error$
End Function

Sub WMSARCH_CreateIndexIfMissing(oCon As Object,sIndex As String,sTarget As String,bUnique As Boolean,ByRef sErr As String)
    Dim oStmt As Object,sql As String
    If sErr<>"" Then Exit Sub
    If WMSARCH_IndexExists(oCon,sIndex,sErr) Then Exit Sub
    If sErr<>"" Then Exit Sub
    sql="CREATE " & IIf(bUnique,"UNIQUE ","") & "INDEX " & sIndex & " ON " & sTarget
    On Error GoTo EH
    oStmt=oCon.createStatement():oStmt.executeUpdate(sql)
    Exit Sub
EH:
    sErr="Создание индекса " & sIndex & ": " & CStr(Err) & " " & Error$
End Sub

Sub WMSARCH_SetMeta(oCon As Object,k As String,v As String,ByRef sErr As String)
    Dim oStmt As Object,sql As String
    sErr=""
    On Error GoTo EH
    oStmt=oCon.createStatement()
    sql="UPDATE WMS_META SET META_VALUE=" & WMSDB_SQLText(v) & ",UPDATED_AT=CURRENT_TIMESTAMP WHERE META_KEY=" & WMSDB_SQLText(k)
    If oStmt.executeUpdate(sql)=0 Then
        oStmt.executeUpdate("INSERT INTO WMS_META (META_KEY,META_VALUE) VALUES (" & WMSDB_SQLText(k) & "," & WMSDB_SQLText(v) & ")")
    End If
    Exit Sub
EH:
    sErr="WMS_META " & k & ": " & CStr(Err) & " " & Error$
End Sub

Function WMSARCH_NextProductCodeCon(oCon As Object,ByRef sErr As String) As String
    Dim oStmt As Object,oRS As Object,n As Long
    WMSARCH_NextProductCodeCon="":sErr=""
    On Error GoTo EH
    oStmt=oCon.createStatement():oRS=oStmt.executeQuery("SELECT NEXT VALUE FOR GEN_WMS_ITEM_ID FROM RDB$DATABASE")
    If oRS.next() Then n=oRS.getLong(1)
    WMSARCH_NextProductCodeCon="ITM-" & Right("00000000" & CStr(n),8)
    Exit Function
EH:
    sErr="Генерация кода товара: " & CStr(Err) & " " & Error$
End Function

Function WMSARCH_CountCon(oCon As Object,sql As String,ByRef sErr As String) As Long
    Dim oStmt As Object,oRS As Object
    WMSARCH_CountCon=0:sErr=""
    On Error GoTo EH
    oStmt=oCon.createStatement():oRS=oStmt.executeQuery(sql)
    If oRS.next() Then WMSARCH_CountCon=oRS.getLong(1)
    Exit Function
EH:
    sErr="Integrity SQL: " & CStr(Err) & " " & Error$
End Function

Function WMSARCH_DeepCheckText(oCon As Object,ByRef nProblems As Long,ByRef sErr As String) As String
    Dim s As String,n As Long
    nProblems=0:sErr=""
    s="WMS Architecture " & WMSARCH_VERSION & Chr(10) & Chr(10)

    n=WMSARCH_CountCon(oCon,"SELECT COUNT(*) FROM WMS_STOCK_MOVEMENTS m LEFT JOIN WMS_PRODUCTS p ON p.PRODUCT_CODE=m.PRODUCT_CODE WHERE p.PRODUCT_CODE IS NULL",sErr)
    If sErr<>"" Then Exit Function
    s=s & "Движения без товара: " & CStr(n) & Chr(10):If n>0 Then nProblems=nProblems+1

    n=WMSARCH_CountCon(oCon,"SELECT COUNT(*) FROM WMS_STOCK_LOTS l LEFT JOIN WMS_PRODUCTS p ON p.PRODUCT_CODE=l.PRODUCT_CODE WHERE p.PRODUCT_CODE IS NULL",sErr)
    If sErr<>"" Then Exit Function
    s=s & "Партии без товара: " & CStr(n) & Chr(10):If n>0 Then nProblems=nProblems+1

    n=WMSARCH_CountCon(oCon,"SELECT COUNT(*) FROM WMS_STOCK_MOVEMENTS m LEFT JOIN WMS_STOCK_LOTS l ON l.LOT_ID=m.LOT_ID WHERE m.LOT_ID IS NOT NULL AND l.LOT_ID IS NULL",sErr)
    If sErr<>"" Then Exit Function
    s=s & "Движения с отсутствующей партией: " & CStr(n) & Chr(10):If n>0 Then nProblems=nProblems+1

    n=WMSARCH_CountCon(oCon,"SELECT COUNT(*) FROM WMS_V_STOCK_BALANCE WHERE QTY<0",sErr)
    If sErr<>"" Then Exit Function
    s=s & "Отрицательные агрегированные остатки: " & CStr(n) & Chr(10):If n>0 Then nProblems=nProblems+1

    n=WMSARCH_CountCon(oCon,"SELECT COUNT(*) FROM WMS_V_LOT_BALANCE WHERE QTY<0",sErr)
    If sErr<>"" Then Exit Function
    s=s & "Отрицательные остатки партий: " & CStr(n) & Chr(10):If n>0 Then nProblems=nProblems+1

    n=WMSARCH_CountCon(oCon,"SELECT COUNT(*) FROM (SELECT ARTICLE_CODE FROM WMS_PRODUCT_ALIASES WHERE ACTIVE_FLAG=1 GROUP BY ARTICLE_CODE HAVING COUNT(DISTINCT PRODUCT_CODE)>1)",sErr)
    If sErr<>"" Then Exit Function
    s=s & "Неоднозначные активные артикулы: " & CStr(n) & Chr(10):If n>0 Then nProblems=nProblems+1

    n=WMSARCH_CountCon(oCon,"SELECT COUNT(*) FROM WMS_OPERATION_GUARD WHERE COALESCE(STATE_NAME,'') NOT IN ('DONE','FAILED','CANCELLED')",sErr)
    If sErr<>"" Then Exit Function
    s=s & "Незавершённые операции: " & CStr(n) & Chr(10):If n>0 Then nProblems=nProblems+1

    n=WMSARCH_CountCon(oCon,"SELECT COUNT(*) FROM WMS_STOCK_MOVEMENTS",sErr)
    If sErr<>"" Then Exit Function
    s=s & "Всего складских движений: " & CStr(n) & Chr(10)

    n=WMSARCH_CountCon(oCon,"SELECT COUNT(*) FROM WMS_STOCK_LOTS",sErr)
    If sErr<>"" Then Exit Function
    s=s & "Всего партий: " & CStr(n) & Chr(10)

    n=WMSARCH_CountCon(oCon,"SELECT COUNT(*) FROM WMS_PRODUCTS",sErr)
    If sErr<>"" Then Exit Function
    s=s & "Номенклатура: " & CStr(n) & Chr(10)

    WMSARCH_DeepCheckText=s
End Function

Sub WMSARCH_DeepCheck()
    Dim oCon As Object,sErr As String,s As String,n As Long
    oCon=WMSDB_GetConnectionEx(ThisComponent,sErr)
    If sErr<>"" Then MsgBox sErr,16,"WMS — Integrity":Exit Sub
    s=WMSARCH_DeepCheckText(oCon,n,sErr)
    WMSDB_Close
    If sErr<>"" Then MsgBox sErr,16,"WMS — Integrity":Exit Sub
    If n=0 Then
        MsgBox s & Chr(10) & "РЕЗУЛЬТАТ: целостность OK.",64,"WMS — Полная проверка"
    Else
        MsgBox s & Chr(10) & "РЕЗУЛЬТАТ: проблемных групп: " & CStr(n),48,"WMS — Полная проверка"
    End If
End Sub
