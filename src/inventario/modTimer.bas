Option Explicit

' Temporizador de Windows. Sirve para cerrar un escaneo por inactividad cuando
' la pistola no envía ninguna tecla al final del código: si pasan MS_INACTIVIDAD
' milisegundos sin recibir un carácter nuevo, el serial se registra solo.

#If VBA7 Then
    Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
    Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
    Private idTimer As LongPtr
#Else
    Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
    Private idTimer As Long
#End If

Public Const MS_INACTIVIDAD As Long = 250

' La pone en True la ventana de escaneo mientras está abierta: evita que el
' temporizador cree una copia fantasma de la ventana si dispara fuera de tiempo
Public VentanaAbierta As Boolean

Public Sub IniciarTemporizador()
    DetenerTemporizador
    idTimer = SetTimer(0&, 0&, MS_INACTIVIDAD, AddressOf TemporizadorVencido)
End Sub

Public Sub DetenerTemporizador()
    If idTimer <> 0 Then
        KillTimer 0&, idTimer
        idTimer = 0
    End If
End Sub

#If VBA7 Then
Public Sub TemporizadorVencido(ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal idEvento As LongPtr, ByVal tiempo As Long)
#Else
Public Sub TemporizadorVencido(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvento As Long, ByVal tiempo As Long)
#End If
    On Error Resume Next
    DetenerTemporizador
    If Not VentanaAbierta Then Exit Sub
    frmEscaneo.CerrarPorInactividad
End Sub
