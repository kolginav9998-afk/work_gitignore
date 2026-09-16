Option Explicit

Global Const WMSRT_VERSION = "3.1.0-RUNTIME-COMPAT"

Function WMSRT_TryEnter(opName As String) As Boolean
    WMSRT_TryEnter=WMSDBX_TryEnter(opName)
End Function

Sub WMSRT_Leave()
    WMSDBX_Leave
End Sub

Sub WMSRT_ResetBusy()
    WMSDBX_ResetBusy
End Sub

Sub WMSRT_HealthSnapshot()
    WMS26_RunDiagnostics
End Sub
