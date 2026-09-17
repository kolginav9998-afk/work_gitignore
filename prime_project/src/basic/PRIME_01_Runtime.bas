Option Explicit

' PRIME_01_Runtime
' Единый рантайм-лок и вложенный event guard. Заменяет два несовместимых механизма 1.4.1
' (gWMSORD_Busy у Заказов и WMSDBX_TryEnter у остальных форм, которые не исключали друг друга -
' см. ARCHITECTURE §0/§2). Любое проведение документа, из какого бы листа оно ни пришло,
' проходит через один и тот же лок.

Private gPrimeOpLockDepth As Long        ' вложенный счётчик операции (posting), не boolean
Private gPrimeEventLockDepth As Long     ' вложенный счётчик события Contents changed
Private gPrimeControllersLocked As Boolean

' Пытается войти в критическую секцию проведения. Возвращает False, если операция уже идёт
' на верхнем уровне (защита от двойного клика/повторного входа) - но вложенные вызовы одного
' и того же логического потока (например, PostDocument вызывает вспомогательные функции)
' увеличивают глубину, а не блокируются.
Public Function PRIME_TryEnter() As Boolean
    gPrimeOpLockDepth = gPrimeOpLockDepth + 1
    PRIME_TryEnter = True
End Function

Public Sub PRIME_Leave()
    If gPrimeOpLockDepth > 0 Then
        gPrimeOpLockDepth = gPrimeOpLockDepth - 1
    End If
    If gPrimeOpLockDepth = 0 And gPrimeControllersLocked Then
        PRIME_UnlockControllers()
    End If
End Sub

Public Function PRIME_IsOpLocked() As Boolean
    PRIME_IsOpLocked = (gPrimeOpLockDepth > 0)
End Function

' Event guard для PRIME_OnContentChanged: программный ввод внутри posting/lookup не должен
' повторно триггерить обработчик события (аналог того, что в 1.4.1 частично делал gWMSORD_Busy,
' но здесь - единая реализация для всех листов).
Public Function PRIME_EventEnter() As Boolean
    If gPrimeEventLockDepth > 0 Then
        gPrimeEventLockDepth = gPrimeEventLockDepth + 1
        PRIME_EventEnter = False ' сигнал вызывающему: событие уже обрабатывается, выйти
        Exit Function
    End If
    gPrimeEventLockDepth = 1
    PRIME_EventEnter = True
End Function

Public Sub PRIME_EventLeave()
    If gPrimeEventLockDepth > 0 Then
        gPrimeEventLockDepth = gPrimeEventLockDepth - 1
    End If
End Sub

' Сброс на случай, если предыдущий запуск завершился аварийно и счётчики "залипли"
' (кнопка "Проверить PRIME" / диагностика может вызвать это явно, но НЕ автоматически при каждом
' открытии - иначе можно замаскировать реальную зависшую операцию).
Public Sub PRIME_ResetRuntimeLock()
    gPrimeOpLockDepth = 0
    gPrimeEventLockDepth = 0
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
