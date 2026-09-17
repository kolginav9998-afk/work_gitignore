Option Explicit
Global Const WMSPRIME_UI_VERSION = "0.1.0"

Sub WMSPRIME_LayoutInfo()
    Dim doc As Object,sh As Object,shape As Object,ctl As Object
    Dim i As Long,workIndex As Long,adminIndex As Long,slot As Long,yBase As Long
    Dim p As New com.sun.star.awt.Point
    Dim sz As New com.sun.star.awt.Size
    Dim nm As String,isAdmin As Boolean
    doc=ThisComponent
    On Error GoTo EH
    sh=doc.Sheets.getByName("Инфо")
    sh.getCellRangeByPosition(0,0,7,39).CellBackColor=16185336
    sh.getCellRangeByPosition(0,0,7,39).CharFontName="Liberation Sans"
    sh.getCellRangeByPosition(0,0,7,39).CharHeight=11
    sh.getCellRangeByPosition(0,0,7,39).CharColor=2372941
    sh.getCellRangeByPosition(0,0,7,39).IsTextWrapped=True
    sh.Rows.getByIndex(0).Height=1100
    For i=1 To 39:sh.Rows.getByIndex(i).Height=650:Next i
    sh.Columns.getByIndex(0).Width=4500
    sh.Columns.getByIndex(1).Width=6200
    sh.Columns.getByIndex(2).Width=700
    For i=3 To 7:sh.Columns.getByIndex(i).Width=3000:Next i
    ' Clear only the old generated panel text, retaining diagnostics A4:B16.
    sh.getCellRangeByPosition(3,0,4,4).clearContents(23)
    sh.getCellRangeByPosition(0,18,1,24).clearContents(23)
    sh.getCellByPosition(0,0).String="ПОКАТАК"
    sh.getCellByPosition(1,0).String="УПРАВЛЕНИЕ СКЛАДОМ"
    sh.getCellByPosition(0,1).String="Панель управления"
    sh.getCellByPosition(3,1).String="ДЕЙСТВИЯ"
    sh.getCellByPosition(3,18).String="СЕРВИС"
    sh.getCellByPosition(0,18).String="Локальная работа"
    sh.getCellByPosition(0,19).String="База рядом с книгой. Интернет не требуется."
    sh.getCellByPosition(0,22).String="Обновление"
    sh.getCellByPosition(0,23).String="Перед изменениями — резервная копия."
    sh.getCellRangeByPosition(0,0,7,0).CellBackColor=2240061
    sh.getCellRangeByPosition(0,0,7,0).CharColor=16777215
    sh.getCellRangeByPosition(0,0,7,0).CharHeight=14
    sh.getCellRangeByPosition(0,0,7,0).CharWeight=150
    sh.getCellRangeByPosition(0,3,1,15).CellBackColor=16777215
    workIndex=0:adminIndex=0
    For i=0 To sh.DrawPage.Count-1
        shape=sh.DrawPage.getByIndex(i)
        If shape.supportsService("com.sun.star.drawing.ControlShape") Then
            ctl=shape.Control
            If ctl.supportsService("com.sun.star.form.component.CommandButton") Then
                nm=ctl.Name:isAdmin=(Left(nm,10)="WMS_CLEAN_")
                If isAdmin Then
                    slot=adminIndex:adminIndex=adminIndex+1:yBase=13700
                Else
                    slot=workIndex:workIndex=workIndex+1:yBase=2300
                End If
                p.X=11400+(slot Mod 2)*6500
                p.Y=yBase+(slot\2)*1200
                sz.Width=6100:sz.Height=1000
                shape.Anchor=sh
                shape.Position=p:shape.Size=sz
                ctl.FontName="Liberation Sans":ctl.FontHeight=10
                ctl.MultiLine=True
                If isAdmin Then
                    ctl.BackgroundColor=16443880:ctl.TextColor=9913874
                Else
                    ctl.BackgroundColor=2240061:ctl.TextColor=16777215
                End If
            End If
        End If
    Next i
    Exit Sub
EH:
    MsgBox "Оформление панели: " & CStr(Err) & " " & Error$,16,"ПОКАТАК"
End Sub

Sub WMSPRIME_LayoutAll()
    Dim sh As Object,sheetIndex As Long
    WMSPRIME_LayoutInfo
    For sheetIndex=0 To ThisComponent.Sheets.Count-1
        sh=ThisComponent.Sheets.getByIndex(sheetIndex)
        If sh.Name="Заказы" Or sh.Name="Выдачи" Or Left(sh.Name,6)="Приход" Or Left(sh.Name,6)="Расход" Then
            WMSPRIME_LayoutOperations sh
        End If
    Next sheetIndex
End Sub

Sub WMSPRIME_LayoutOperations(sh As Object)
    Dim i As Long,n As Long,leftEdge As Long,topEdge As Long,isWorkflow As Boolean
    Dim shape As Object,ctl As Object
    Dim p As New com.sun.star.awt.Point
    Dim sz As New com.sun.star.awt.Size
    On Error GoTo EH
    isWorkflow=(Left(sh.Name,6)="Приход" Or Left(sh.Name,6)="Расход")
    If isWorkflow Then
        sh.Rows.getByIndex(3).Height=2400
        leftEdge=100:topEdge=sh.getCellByPosition(0,3).Position.Y+100
    Else
        leftEdge=2147483647:topEdge=100
        For i=0 To sh.DrawPage.Count-1
            shape=sh.DrawPage.getByIndex(i)
            If shape.supportsService("com.sun.star.drawing.ControlShape") Then
                If shape.Control.supportsService("com.sun.star.form.component.CommandButton") Then
                    If shape.Position.X<leftEdge Then leftEdge=shape.Position.X
                End If
            End If
        Next i
        If leftEdge=2147483647 Then Exit Sub
    End If
    n=0
    For i=0 To sh.DrawPage.Count-1
        shape=sh.DrawPage.getByIndex(i)
        If shape.supportsService("com.sun.star.drawing.ControlShape") Then
            ctl=shape.Control
            If ctl.supportsService("com.sun.star.form.component.CommandButton") Then
                If isWorkflow Then
                    p.X=leftEdge+(n Mod 4)*5600:p.Y=topEdge+(n\4)*1150
                    sz.Width=5300
                Else
                    p.X=leftEdge:p.Y=topEdge+n*1150
                    sz.Width=6100
                End If
                sz.Height=1000
                shape.Anchor=sh:shape.Position=p:shape.Size=sz
                ctl.FontName="Liberation Sans":ctl.FontHeight=10:ctl.MultiLine=True
                ctl.BackgroundColor=2240061:ctl.TextColor=16777215
                n=n+1
            End If
        End If
    Next i
    Exit Sub
EH:
    MsgBox "Панель «" & sh.Name & "»: " & CStr(Err) & " " & Error$,16,"ПОКАТАК"
End Sub
