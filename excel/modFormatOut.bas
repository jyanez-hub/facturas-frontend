Attribute VB_Name = "modFormatOut"
'=====================  modFormatOut  =====================
Option Explicit

' Llama esto después de crear el libro nuevo (wbOut)
Public Sub EX_FormatTablesAndCurrency(Optional ByVal wbOut As Workbook)
    Dim wb As Workbook, ws As Worksheet, tName As String
    If wbOut Is Nothing Then Set wb = ActiveWorkbook Else Set wb = wbOut
    If wb Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    For Each ws In wb.Worksheets
        If WorksheetHasData(ws) Then
            ' 1) Convierte en Tabla
            tName = SafeTableName(ws.name)
            ConvertUsedRangeToTable ws, tName

            ' 2) Autoajuste de columnas
            ws.Cells.EntireColumn.AutoFit

            ' 3) Formatos por encabezado (moneda / fecha / porcentaje / cantidades)
            ApplyFormatsByHeader ws
        End If
    Next ws
    Application.ScreenUpdating = True
End Sub

'----- Helpers -----

Private Function WorksheetHasData(ByVal ws As Worksheet) As Boolean
    On Error Resume Next
    WorksheetHasData = (Application.WorksheetFunction.CountA(ws.UsedRange) > 0)
End Function

Private Function SafeTableName(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, " ", "_")
    t = Replace(t, ".", "_")
    t = Replace(t, "-", "_")
    t = Replace(t, "'", "")
    If Len(t) = 0 Then t = "Tabla1"
    SafeTableName = t
End Function

' Convierte UsedRange en Tabla (ListObject). Si ya es tabla, la reutiliza.
Private Sub ConvertUsedRangeToTable(ByVal ws As Worksheet, ByVal tblName As String)
    Dim lo As ListObject, rng As Range

    If ws.ListObjects.Count > 0 Then
        Set lo = ws.ListObjects(1)
        On Error Resume Next
        lo.name = tblName
        On Error GoTo 0
        Exit Sub
    End If

    Set rng = ws.UsedRange
    If rng Is Nothing Then Exit Sub
    If rng.Rows.Count < 2 Or rng.Columns.Count < 1 Then Exit Sub

    ' Asegura que la fila 1 sea encabezado
    Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    On Error Resume Next
    lo.name = tblName
    On Error GoTo 0

    ' **No** fijamos estilo aquí para no sobreescribir el corporativo.
End Sub

' Aplica formatos leyendo encabezados de la fila 1
Private Sub ApplyFormatsByHeader(ByVal ws As Worksheet)
    Dim lastCol As Long, c As Long, h As String
    Dim lo As ListObject, rngData As Range, colRng As Range

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < 1 Then Exit Sub

    ' Si hay tabla, aplicamos sobre DataBodyRange; si no, desde fila 2
    If ws.ListObjects.Count > 0 Then
        Set lo = ws.ListObjects(1)
        If lo.DataBodyRange Is Nothing Then Exit Sub
    End If

    For c = 1 To lastCol
        h = LCase$(Trim$(CStr(ws.Cells(1, c).Value)))

        If ws.ListObjects.Count > 0 Then
            Set colRng = lo.DataBodyRange.Columns(c)
        Else
            Set colRng = ws.Range(ws.Cells(2, c), ws.Cells(ws.Rows.Count, c).End(xlUp))
        End If
        If colRng Is Nothing Then GoTo NextC

        ' ----- Fecha -----
        If h = "fecha_emision" Or h Like "*fecha*" Then
            colRng.NumberFormat = "dd/mm/yyyy"
            GoTo NextC
        End If

        ' ----- Porcentajes -----
        If h = "porc_ret_iva" Or h = "porc_ret_renta" Or h Like "*porcentaje*" Then
            colRng.NumberFormat = "0.00"
            GoTo NextC
        End If

        ' ----- Cantidades (no dinero) -----
        If h Like "*cantidad*" Then
            colRng.NumberFormat = "0.0000"
            GoTo NextC
        End If

        ' ----- MONEDA ($) -----
        If IsCurrencyHeader(h) Then
            ' Usa formato en USD; si prefieres el local, cambia por: "#,##0.00 [$-$-ec]"
            colRng.NumberFormat = """$""#,##0.00"
            GoTo NextC
        End If

NextC:
    Next c
End Sub

' Reglas para detectar columnas de dinero por su encabezado
Private Function IsCurrencyHeader(ByVal h As String) As Boolean
    ' Encabezados típicos en Facturas/Detalle
    If h Like "precio*" Then IsCurrencyHeader = True: Exit Function
    If h Like "*valor*" Then IsCurrencyHeader = True: Exit Function
    If h Like "base*" Then IsCurrencyHeader = True: Exit Function
    If h Like "*subtotal*" Then IsCurrencyHeader = True: Exit Function
    If h Like "*iva*" Then IsCurrencyHeader = True: Exit Function
    If h Like "*total*" Then IsCurrencyHeader = True: Exit Function
    If h Like "*descuento*" Then IsCurrencyHeader = True: Exit Function
    If h Like "*propina*" Then IsCurrencyHeader = True: Exit Function

    ' Específicos de Retenciones / Detalle Ret.
    If h = "base_imponible" Or h = "valor_retenido" Then IsCurrencyHeader = True: Exit Function
    If h = "base_ret_iva" Or h = "valor_ret_iva" Then IsCurrencyHeader = True: Exit Function
    If h = "base_ret_renta" Or h = "valor_ret_renta" Then IsCurrencyHeader = True: Exit Function
    If h = "total_retenido" Then IsCurrencyHeader = True: Exit Function
End Function

' ======= Formatos específicos =======

' Convierte textos "YYYY-MM-DD" o "DD-MM-YYYY" (o con "/") a fecha real
Private Sub CoerceRangeToDate(ByVal rng As Range)
    On Error Resume Next
    Dim c As Range, s As String, y As Long, m As Long, d As Long, parts() As String
    For Each c In rng.Cells
        If Not IsEmpty(c.Value) Then
            If IsDate(c.Value) Then
                ' ya es fecha ? nada
            Else
                s = Trim$(CStr(c.Value))
                s = Replace(s, ".", "-")
                s = Replace(s, "/", "-")
                If Len(s) >= 8 And InStr(s, "-") > 0 Then
                    parts = Split(s, "-")
                    If UBound(parts) = 2 Then
                        If Len(parts(0)) = 4 Then
                            y = CLng(parts(0)): m = CLng(parts(1)): d = CLng(parts(2)) ' YYYY-MM-DD
                        Else
                            d = CLng(parts(0)): m = CLng(parts(1)): y = CLng(parts(2)) ' DD-MM-YYYY
                            If y < 100 Then y = 2000 + y
                        End If
                        If y >= 1900 And m >= 1 And m <= 12 And d >= 1 And d <= 31 Then
                            c.Value = DateSerial(y, m, d)
                        End If
                    End If
                End If
            End If
        End If
    Next c
    On Error GoTo 0
End Sub

' Aplica formato Contabilidad $ (separadores , y .) – Ecuador usa $ con , miles y . decimales
Private Sub ApplyAccountingUSD(ByVal rng As Range)
    ' Patrón contabilidad con símbolo $ alineado
    rng.NumberFormat = "_-[$$-409]* #,##0.00_-;_-[$$-409]* -#,##0.00_-;_-[$$-409]* ""-""??_-;_-@_-"
    ' 409 = LCID inglés (mantiene , y .). Funciona bien con configuración es-EC (coma miles, punto decimales).
End Sub
'============================================================


