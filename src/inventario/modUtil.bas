Option Explicit

Public Function Hj(nombre As String) As Worksheet
    Set Hj = ThisWorkbook.Worksheets(nombre)
End Function

Public Function NuevoDic(Optional texto As Boolean = False) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    If texto Then d.CompareMode = vbTextCompare
    Set NuevoDic = d
End Function

Public Function UltimaFila(ws As Worksheet, Optional col As Long = 1) As Long
    UltimaFila = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

' Serial en mayúsculas y sin espacios ni caracteres de control
Public Function NormalizarSerial(ByVal s As String) As String
    Dim i As Long, c As String, r As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If AscW(c) > 32 Then r = r & c
    Next
    NormalizarSerial = UCase$(r)
End Function

Public Sub GuardarSilencioso()
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Save
    Application.DisplayAlerts = True
End Sub

' ThisWorkbook.Path devuelve una URL (https://...) cuando el archivo está en
' OneDrive. Power Query y el sistema de archivos necesitan la ruta local.
Public Function RutaLocal() As String
    Dim p As String: p = ThisWorkbook.Path
    If LCase(Left(p, 4)) <> "http" Then
        RutaLocal = p
        Exit Function
    End If

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim partes() As String: partes = Split(DecodificarURL(p), "/")
    Dim raices As Variant, r As Variant, i As Long, j As Long, resto As String
    raices = Array(Environ("OneDriveCommercial"), Environ("OneDriveConsumer"), Environ("OneDrive"))
    For Each r In raices
        If Len(r) > 0 Then
            ' Del sufijo más largo al más corto: el primero que exista es la carpeta
            For i = 3 To UBound(partes)
                resto = ""
                For j = i To UBound(partes)
                    resto = resto & "\" & partes(j)
                Next
                If fso.FolderExists(r & resto) Then
                    RutaLocal = r & resto
                    Exit Function
                End If
            Next
        End If
    Next
    RutaLocal = ""
End Function

Private Function DecodificarURL(s As String) As String
    Dim i As Long, c As String, res As String, b1 As Long, b2 As Long, b3 As Long
    On Error GoTo Fin
    i = 1
    Do While i <= Len(s)
        c = Mid(s, i, 1)
        If c = "%" And i + 2 <= Len(s) Then
            b1 = CLng("&H" & Mid(s, i + 1, 2))
            If b1 < &H80 Then
                res = res & Chr(b1): i = i + 3
            ElseIf b1 >= &HE0 And i + 8 <= Len(s) Then
                b2 = CLng("&H" & Mid(s, i + 4, 2)): b3 = CLng("&H" & Mid(s, i + 7, 2))
                res = res & ChrW((b1 And &HF) * 4096 + (b2 And &H3F) * 64 + (b3 And &H3F)): i = i + 9
            ElseIf b1 >= &HC0 And i + 5 <= Len(s) Then
                b2 = CLng("&H" & Mid(s, i + 4, 2))
                res = res & ChrW((b1 And &H1F) * 64 + (b2 And &H3F)): i = i + 6
            Else
                res = res & c: i = i + 1
            End If
        Else
            res = res & c: i = i + 1
        End If
    Loop
    DecodificarURL = res
    Exit Function
Fin:
    DecodificarURL = s
End Function

' Reemplaza la ruta dentro de File.Contents("...") en la fórmula M
Public Function CambiarRuta(formula As String, ruta As String) As String
    Dim marca As String: marca = "File.Contents("""
    Dim ini As Long, fin As Long
    ini = InStr(formula, marca) + Len(marca)
    fin = InStr(ini, formula, """)")
    CambiarRuta = Left(formula, ini - 1) & Replace(ruta, """", """""") & Mid(formula, fin)
End Function
