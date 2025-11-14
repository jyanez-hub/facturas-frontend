Attribute VB_Name = "modHyperlinksComprobantes"
'================== modHyperlinksComprobantes ==================
Option Explicit

' Usa la ruta guardada. Si no hay ruta, NO pregunta: simplemente sale.
Public Sub AutoHyperlinks_AfterImport(Optional ByVal askIfMissing As Boolean = False)
    Dim baseCarpeta As String
    baseCarpeta = GetRutaPDF()

    If Len(baseCarpeta) = 0 Then
        If askIfMissing Then
            baseCarpeta = SeleccionarCarpetaPDFs("Selecciona la carpeta donde están los PDF renombrados")
            If Len(baseCarpeta) = 0 Then Exit Sub
            SetRutaPDF baseCarpeta
        Else
            Exit Sub
        End If
    End If

    If Right$(baseCarpeta, 1) <> "\" Then baseCarpeta = baseCarpeta & "\"
    CrearHipervinculosParaHoja "facturas", baseCarpeta
    CrearHipervinculosParaHoja "retenciones", baseCarpeta
End Sub

' ====== Punto de entrada (si ya tienes la ruta) ======
Public Sub AutoHyperlinks_AfterImport_ConRuta(ByVal baseCarpeta As String)
    If Len(baseCarpeta) = 0 Then Exit Sub
    If Right$(baseCarpeta, 1) <> "\" Then baseCarpeta = baseCarpeta & "\"
    CrearHipervinculosParaHoja "facturas", baseCarpeta
    CrearHipervinculosParaHoja "retenciones", baseCarpeta
End Sub

' ====== Núcleo por hoja (SOLO DataBodyRange; encabezado intacto) ======
Private Sub CrearHipervinculosParaHoja(ByVal nombreHoja As String, ByVal baseCarpeta As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = WBH.Worksheets(nombreHoja)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    If ws.ListObjects.Count = 0 Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim lo As ListObject
    Set lo = ws.ListObjects(1)

    ' 1) Ubicar columna nro_comprobante dentro de la TABLA
    Dim idxNro As Long
    idxNro = ColIndexInTableByHeader(lo, "nro_comprobante")
    If idxNro = 0 Then
        ' fallback con tu ColByHeader en fila 1
        idxNro = ColByHeader(ws, "nro_comprobante", 1)
        If idxNro = 0 Then idxNro = 12 ' L por defecto
    End If

    ' 2) Columnas auxiliares por si hay que reconstruir el nombre
    Dim cTipo As Long, cEst As Long, cPto As Long, cSec As Long
    cTipo = ColIndexInTableByHeader(lo, "tipo"): If cTipo = 0 Then cTipo = ColByHeader(ws, "tipo", 1)
    cEst = ColIndexInTableByHeader(lo, "estab"): If cEst = 0 Then cEst = ColByHeader(ws, "estab", 1)
    cPto = ColIndexInTableByHeader(lo, "ptoEmi"): If cPto = 0 Then cPto = ColByHeader(ws, "ptoEmi", 1)
    cSec = ColIndexInTableByHeader(lo, "secuencial"): If cSec = 0 Then cSec = ColByHeader(ws, "secuencial", 1)

    ' 3) Columna de estado (no tocamos formato de encabezado)
    Dim colRes As Long: colRes = EnsureEstadoCol(ws)
    ws.Columns(colRes).NumberFormat = "@"

    ' 4) Rangos: SOLO cuerpo de la tabla en la columna objetivo
    If lo.DataBodyRange Is Nothing Then GoTo Salir
    Dim rngDatos As Range
    If idxNro > 0 Then
        ' Si idxNro es índice relativo a la tabla
        If idxNro <= lo.ListColumns.Count Then
            Set rngDatos = lo.DataBodyRange.Columns(idxNro)
        Else
            ' Si vino como índice de hoja, tradúcelo al cuerpo
            Set rngDatos = Intersect(lo.DataBodyRange, ws.Columns(idxNro))
        End If
    End If
    If rngDatos Is Nothing Then GoTo Salir

    ' 5) LIMPIEZA: borrar hipervínculos previos SOLO en el cuerpo
    LimpiarHipervinculos_Range rngDatos

    ' Asegurar que el ENCABEZADO NO conserve hipervínculos ni subrayado
    On Error Resume Next
    If Not lo.HeaderRowRange Is Nothing Then
        If lo.HeaderRowRange.Hyperlinks.Count > 0 Then lo.HeaderRowRange.Hyperlinks.Delete
        lo.HeaderRowRange.Font.Underline = xlUnderlineStyleNone
        ' NO tocar .Font.Color para no perder el blanco del estilo corporativo
    End If
    On Error GoTo 0

    ' 6) Crear hipervínculos fila por fila
    Dim c As Range, txt As String, baseName As String, rutaPDF As String
    Dim fila As Long
    For Each c In rngDatos.Cells
        txt = Trim$(CStr(c.Value))
        fila = c.Row

        If txt <> "" Then
            baseName = ResolverBaseName(ws, fila, txt, cTipo, cEst, cPto, cSec)
            rutaPDF = ""

            If baseName <> "" Then
                rutaPDF = BuscarPDF(baseCarpeta, baseName, ExtraerSecuencial(baseName))
            ElseIf cSec > 0 Then
                rutaPDF = BuscarPDF(baseCarpeta, "", PadLeft(Trim$(CStr(ws.Cells(fila, cSec).Value)), 9, "0"))
            End If

            If rutaPDF <> "" Then
                ws.Hyperlinks.Add Anchor:=c, Address:=rutaPDF, TextToDisplay:=txt
                ' Si deseas el link azul en el cuerpo, descomenta:
                ' c.Font.Underline = xlUnderlineStyleSingle
                ' c.Font.Color = vbBlue
                ws.Cells(fila, colRes).Value = "OK"
            Else
                ws.Cells(fila, colRes).Value = "PDF no encontrado"
            End If
        Else
            ws.Cells(fila, colRes).Value = ""
        End If
    Next c

Salir:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

' ====== Helpers ======

' Borra hipervínculos SOLO en el rango dado (no toca encabezado/estilo)
Private Sub LimpiarHipervinculos_Range(ByVal rng As Range)
    On Error Resume Next
    If rng.Hyperlinks.Count > 0 Then rng.Hyperlinks.Delete
    rng.Font.Underline = xlUnderlineStyleNone
    ' No forzamos color a "Automático" para no heredar negro en encabezado
    On Error GoTo 0
End Sub

' Índice de columna por encabezado dentro de un ListObject (case-insensitive)
Private Function ColIndexInTableByHeader(ByVal lo As ListObject, ByVal headerName As String) As Long
    Dim i As Long, t As String, h As String
    h = LCase$(headerName)
    For i = 1 To lo.ListColumns.Count
        t = LCase$(CStr(lo.ListColumns(i).name))
        If t = h Then ColIndexInTableByHeader = i: Exit Function
    Next i
End Function

Private Function ResolverBaseName(ws As Worksheet, ByVal r As Long, ByVal txt As String, _
                                  ByVal cTipo As Long, ByVal cEst As Long, ByVal cPto As Long, ByVal cSec As Long) As String
    Dim re As Object: Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True: re.Global = False

    ' ¿ya viene con el patrón completo?
    re.Pattern = "^[A-Z]{2}-\d{3}-\d{3}-\d{9}$"
    If re.Test(txt) Then ResolverBaseName = txt: Exit Function

    ' Intentar armarlo con columnas auxiliares
    Dim tipo As String, estab As String, pto As String, sec As String, pref As String
    If cTipo > 0 Then tipo = Trim$(CStr(ws.Cells(r, cTipo).Value))
    If cEst > 0 Then estab = Trim$(CStr(ws.Cells(r, cEst).Value))
    If cPto > 0 Then pto = Trim$(CStr(ws.Cells(r, cPto).Value))
    If cSec > 0 Then sec = Trim$(CStr(ws.Cells(r, cSec).Value))

    If sec = "" Then sec = txt ' por si L trae solo el secuencial

    pref = NormalizarPrefijo(tipo, ws.name)

    estab = PadLeft(estab, 3, "0")
    pto = PadLeft(pto, 3, "0")
    sec = PadLeft(SoloDigitos(sec), 9, "0")

    If pref <> "" And estab <> "" And pto <> "" And sec <> "" Then
        ResolverBaseName = pref & "-" & estab & "-" & pto & "-" & sec
    End If
End Function

Private Function NormalizarPrefijo(ByVal tipo As String, ByVal nombreHoja As String) As String
    Dim t As String: t = UCase$(Trim$(tipo))
    ' Si ya viene FC/NC/ND/CR
    If t = "FC" Or t = "NC" Or t = "ND" Or t = "CR" Then
        NormalizarPrefijo = t: Exit Function
    End If
    ' Si viene codDoc (01/04/05/07)
    Select Case t
        Case "01": NormalizarPrefijo = "FC": Exit Function
        Case "04": NormalizarPrefijo = "NC": Exit Function
        Case "05": NormalizarPrefijo = "ND": Exit Function
        Case "07": NormalizarPrefijo = "CR": Exit Function
    End Select
    ' Si no hay tipo, infiere por la hoja
    If LCase$(nombreHoja) = "facturas" Then NormalizarPrefijo = "FC"
    If LCase$(nombreHoja) = "retenciones" Then NormalizarPrefijo = "CR"
End Function

Private Function BuscarPDF(ByVal carpeta As String, ByVal baseName As String, ByVal secuencial As String) As String
    Dim f As String
    If baseName <> "" Then
        f = carpeta & baseName & ".pdf"
        If dir$(f, vbNormal) <> "" Then BuscarPDF = f: Exit Function
    End If
    If Len(secuencial) >= 4 Then
        f = dir$(carpeta & "*" & secuencial & "*.pdf", vbNormal)
        If f <> "" Then BuscarPDF = carpeta & f
    End If
End Function

Private Function ColByHeader(ByVal ws As Worksheet, ByVal headerName As String, ByVal filaHeader As Long) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(filaHeader, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        If LCase$(Trim$(CStr(ws.Cells(filaHeader, c).Value))) = LCase$(headerName) Then
            ColByHeader = c: Exit Function
        End If
    Next c
End Function

Private Function ExtraerSecuencial(ByVal baseName As String) As String
    Dim p As Long: p = InStrRev(baseName, "-")
    If p > 0 Then ExtraerSecuencial = Mid$(baseName, p + 1)
End Function

Private Function PadLeft(ByVal s As String, ByVal totalLen As Long, ByVal ch As String) As String
    Dim t As String: t = Trim$(s)
    If Len(t) >= totalLen Then PadLeft = t Else PadLeft = String$(totalLen - Len(t), ch) & t
End Function

Private Function SoloDigitos(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then out = out & ch
    Next i
    SoloDigitos = out
End Function

Private Function SeleccionarCarpetaPDFs(ByVal titulo As String) As String
    Dim sh As Object: Set sh = CreateObject("Shell.Application")
    Dim f As Object: Set f = sh.BrowseForFolder(0, titulo, 0, 0)
    If Not f Is Nothing Then SeleccionarCarpetaPDFs = f.Self.Path & IIf(Right$(f.Self.Path, 1) = "\", "", "\")
End Function

' Devuelve la columna para escribir el estado; si no existe "estado_pdf", la crea al final.
Private Function EnsureEstadoCol(ByVal ws As Worksheet) As Long
    Dim col As Long
    col = ColByHeader(ws, "estado_pdf", 1)
    If col > 0 Then
        EnsureEstadoCol = col
        Exit Function
    End If

    ' Crear nueva columna al final de los encabezados (fila 1)
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    col = lastCol + 1
    ws.Cells(1, col).Value = "estado_pdf"
    ws.Cells(1, col).Font.Bold = True
    ws.Columns(col).NumberFormat = "@"
    EnsureEstadoCol = col
End Function

Private Function WBH() As Workbook
    If Not gTargetWb Is Nothing Then
        Set WBH = gTargetWb
    Else
        Set WBH = ThisWorkbook
    End If
End Function


