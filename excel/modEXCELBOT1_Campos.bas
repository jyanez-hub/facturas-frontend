Attribute VB_Name = "modEXCELBOT1_Campos"
Option Explicit

' tipoDetalle: False=Encabezado, True=Detalle
' esRet: False = Factura/NC/ND/LiqCompra ; True = Retenciones
Public Function EX_GetHeadersDisponibles(ByVal tipoDetalle As Boolean, ByVal esRet As Boolean) As Collection
    Dim r As New Collection
    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")

    On Error Resume Next: InitHeaders: On Error GoTo 0

    If esRet = False Then
        If tipoDetalle = False Then
            EX_AddHeaders_FromSheetOrArray r, seen, "Facturas", FACTURAS_HEADERS
        Else
            EX_AddHeaders_FromSheetOrArray r, seen, "Detalle", DETALLE_HEADERS
        End If
    Else
        If tipoDetalle = False Then
            EX_AddHeaders_FromSheetOrArray r, seen, "Retenciones", RET_HEADERS
        Else
            EX_AddHeaders_FromSheetOrArray r, seen, "RetDet", RETDET_HEADERS
        End If
    End If

    Set EX_GetHeadersDisponibles = r
End Function

Private Sub EX_AddHeaders_FromSheetOrArray(ByRef r As Collection, ByRef seen As Object, _
                                           ByVal hoja As String, ByRef arr As Variant)
    Dim ws As Worksheet, lastCol As Long, c As Long, h As String
    On Error Resume Next: Set ws = ThisWorkbook.Worksheets(hoja): On Error GoTo 0

    If Not ws Is Nothing Then
        If Application.WorksheetFunction.CountA(ws.Rows(1)) > 0 Then
            lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            For c = 1 To lastCol
                h = Trim$(CStr(ws.Cells(1, c).Value))
                If Len(h) > 0 Then EX_AddUnique r, seen, h, hoja
            Next c
            Exit Sub
        End If
    End If

    Dim i As Long
    If IsArray(arr) Then
        For i = LBound(arr) To UBound(arr)
            h = Trim$(CStr(arr(i)))
            If Len(h) > 0 Then EX_AddUnique r, seen, h, hoja & " (def)"
        Next i
    End If
End Sub

' Evita duplicados (case-insensitive)
Private Sub EX_AddUnique(ByRef r As Collection, ByRef seen As Object, _
                         ByVal headerName As String, ByVal origen As String)
    Dim key As String: key = LCase$(headerName)
    If Not seen.Exists(key) Then
        seen.Add key, True
        r.Add Array(headerName, origen)
    End If
End Sub

' Preselección estándar de campos (IVA/Renta y trazabilidad)
Public Function EX_DefaultCampos_Documento() As Variant
    EX_DefaultCampos_Documento = Array( _
        "fecha_emision", "tipo_comprobante", "nro_comprobante", _
        "ruc_emisor", "razon_social_emisor", "ruc_ci_comprador", "razon_social_comprador", _
        "subtotal_iva_0", "subtotal_iva_12", "subtotal_iva_15", "subtotal_no_objeto", "subtotal_exento", _
        "iva_total", "descuento", "valor_total", _
        "doc_sust_tipo", "doc_sust_serie", "doc_sust_secuencial", "doc_sust_fecha", _
        "clave_acceso" _
    )
End Function

Public Function EX_DefaultCampos_Detalle() As Variant
    EX_DefaultCampos_Detalle = Array( _
        "clave_acceso", "codigo_principal", "codigo_auxiliar", "descripcion", _
        "cantidad", "precio_unitario", "descuento", "precio_total_sin_impuesto", _
        "impuesto_codigo", "impuesto_porcentaje_codigo", "tarifa_iva", _
        "base_imponible", "valor_iva" _
    )
End Function

Public Function EX_DefaultCampos_RetDoc() As Variant
    EX_DefaultCampos_RetDoc = Array( _
        "fecha_emision", "nro_comprobante", _
        "ruc_emisor", "razon_social_emisor", _
        "ruc_sujeto", "razon_social_sujeto", _
        "periodo_fiscal", _
        "cod_ret_iva", "porc_ret_iva", "base_ret_iva", "valor_ret_iva", _
        "cod_ret_renta", "porc_ret_renta", "base_ret_renta", "valor_ret_renta", _
        "total_retenido", _
        "doc_sust_tipo", "doc_sust_serie", "doc_sust_secuencial", "doc_sust_fecha", _
        "clave_acceso" _
    )
End Function

Public Function EX_DefaultCampos_RetDet() As Variant
    EX_DefaultCampos_RetDet = Array( _
        "clave_acceso", "impuesto", "codigo_retencion", _
        "base_imponible", "porcentaje_retener", "valor_retenido" _
    )
End Function

' Marca seleccionados en un ListBox (columna 0 = nombre del campo)
Public Sub EX_PreseleccionarEnListBox(ByRef lst As MSForms.ListBox, ByVal defs As Variant)
    On Error Resume Next
    Dim i As Long, j As Long
    For i = 0 To lst.ListCount - 1
        For j = LBound(defs) To UBound(defs)
            If StrComp(lst.List(i, 0), CStr(defs(j)), vbTextCompare) = 0 Then
                lst.Selected(i) = True: Exit For
            End If
        Next j
    Next i
    On Error GoTo 0
End Sub

