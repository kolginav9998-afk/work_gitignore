Option Explicit

' PRIME_01_Runtime
' Единый рантайм-лок и вложенный event guard. Заменяет два несовместимых механизма 1.4.1
' (gWMSORD_Busy у Заказов и WMSDBX_TryEnter у остальных форм, которые не исключали друг друга -
' см. ARCHITECTURE §0/§2). Любое проведение документа, из какого бы листа оно ни пришло,
' проходит через один и тот же лок.

' PRIME 2.0.1: оба лока переписаны как простые булевы мьютексы (не вложенные счётчики).
' 2.0.0 багфикс: PRIME_TryEnter только считал глубину и ВСЕГДА возвращал True - реального
' взаимного исключения не было, двойной клик/повторный вход не блокировался вовсе. Ни
' PRIME_PostDocument, ни какой-либо другой код в кодовой базе не полагается на вложенный
' (реентерабельный) вызов PRIME_TryEnter изнутри уже захваченной операции - единственный
' вызывающий каждого лока сам ставит guard один раз на верхнем уровне - поэтому простой
' boolean corректен и не меняет легитимные сценарии использования.
Private gPrimeOpLocked As Boolean
Private gPrimeEventLocked As Boolean     ' было: вложенный счётчик; см. фикс ниже
Private gPrimeControllersLocked As Boolean

' Пытается войти в критическую секцию проведения. Возвращает False, если операция уже
' выполняется (double_click_idempotent: двойной клик/повторный вход не должен запускать
' второе проведение). Вызывающий (PRIME_PostDocument) ОБЯЗАН проверить результат и не
' продолжать при False - см. required_behavior в мастер-задании 2.0.1.
Public Function PRIME_TryEnter() As Boolean
    If gPrimeOpLocked Then
        PRIME_TryEnter = False
        Exit Function
    End If
    gPrimeOpLocked = True
    PRIME_TryEnter = True
End Function

' Идемпотентно: снимать лок можно, даже если он уже снят (например, cleanup вызван дважды
' на разных путях выхода) - без риска "занизить" чужой лок, т.к. лок один на процесс/сеанс.
Public Sub PRIME_Leave()
    gPrimeOpLocked = False
    If gPrimeControllersLocked Then
        PRIME_UnlockControllers()
    End If
End Sub

Public Function PRIME_IsOpLocked() As Boolean
    PRIME_IsOpLocked = gPrimeOpLocked
End Function

' Event guard для PRIME_OnContentChanged: программный ввод внутри posting/lookup не должен
' повторно триггерить обработчик события (аналог того, что в 1.4.1 частично делал gWMSORD_Busy,
' но здесь - единая реализация для всех листов).
' 2.0.0 багфикс: при повторном (вложенном) входе старый код увеличивал depth И возвращал
' False. Каждый вызывающий использует паттерн "If Not PRIME_EventEnter() Then Exit Sub" -
' то есть при False он НИКОГДА не вызывает PRIME_EventLeave (выходит раньше). Из-за этого
' depth, увеличенный на вложенном вызове, никогда не уменьшался обратно - guard "залипал"
' навсегда после первого же вложенного события (например, программное автозаполнение поля
' внутри обработчика, которое само генерирует ContentChanged). Теперь при блокировке depth
' НЕ меняется вовсе - симметрия Enter/Leave сохраняется только для успешных (True) входов.
Public Function PRIME_EventEnter() As Boolean
    If gPrimeEventLocked Then
        PRIME_EventEnter = False ' уже обрабатывается - вызывающий обязан выйти, НЕ вызывая EventLeave
        Exit Function
    End If
    gPrimeEventLocked = True
    PRIME_EventEnter = True
End Function

Public Sub PRIME_EventLeave()
    gPrimeEventLocked = False
End Sub

' Сброс на случай, если предыдущий запуск завершился аварийно и лок "залип"
' (кнопка "Проверить PRIME" / диагностика может вызвать это явно, но НЕ автоматически при каждом
' открытии - иначе можно замаскировать реальную зависшую операцию).
Public Sub PRIME_ResetRuntimeLock()
    gPrimeOpLocked = False
    gPrimeEventLocked = False
    If gPrimeControllersLocked Then
        PRIME_UnlockControllers()
    End If
End Sub

' Блокировка контроллеров UI разрешена только для коротких пакетных операций (runtime.lock_controllers)
' и всегда снимается в cleanup, даже при ошибке.
Public Sub PRIME_LockControllersForBatch()
    On Error Resume Next
    ThisComponent.lockControllers()
    gPrimeControllersLocked = True
End Sub

Public Sub PRIME_UnlockControllers()
    On Error Resume Next
    If gPrimeControllersLocked Then
        ThisComponent.unlockControllers()
    End If
    gPrimeControllersLocked = False
End Sub

' Универсальный cleanup, вызывается из каждого posting/event пути в Finally-подобном блоке
' (см. PRIME_04_Posting.PRIME_PostDocument) - гарантирует снятие лока и разблокировку UI
' независимо от того, где именно произошла ошибка.
Public Sub PRIME_RuntimeCleanup(ByVal wasEventScope As Boolean)
    If wasEventScope Then
        PRIME_EventLeave()
    Else
        PRIME_Leave()
    End If
End Sub
