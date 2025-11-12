Attribute VB_Name = "modSetup"
Option Explicit

'##########################################################################################
'# Módulo: modSetup
'# Propósito: Crear la estructura base del libro para importar comprobantes electrónicos
'#             del SRI (facturas, notas de crédito y retenciones) y preparar tablas
'#             estructuradas con formatos coherentes.
'##########################################################################################

'--- Constantes públicas para referencia de otras rutinas ---
Public Const SHEET_FACTURAS As String = "Facturas"
Public Const TABLE_FACTURAS As String = "tblFacturas"

Public Const SHEET_FACTURAS_DETALLE As String = "Facturas_Detalle"
Public Const TABLE_FACTURAS_DETALLE As String = "tblFacturasDetalle"

Public Const SHEET_NOTAS_CREDITO As String = "NotasCredito"
Public Const TABLE_NOTAS_CREDITO As String = "tblNotasCredito"

Public Const SHEET_NOTAS_CREDITO_DETALLE As String = "NotasCredito_Detalle"
Public Const TABLE_NOTAS_CREDITO_DETALLE As String = "tblNotasCreditoDetalle"

Public Const SHEET_RETENCIONES As String = "Retenciones"
Public Const TABLE_RETENCIONES As String = "tblRetenciones"

Public Const SHEET_RETENCIONES_DETALLE As String = "Retenciones_Detalle"
Public Const TABLE_RETENCIONES_DETALLE As String = "tblRetencionesDetalle"

Public Const SHEET_PARAMETROS As String = "Parametros"
Public Const TABLE_PARAMETROS As String = "tblParametros"

Public Const SHEET_LOG As String = "LogImportacion"
Public Const TABLE_LOG As String = "tblLogImportacion"

Public Const SHEET_CONTROL As String = "Control"
Public Const SHEET_SELECTOR_CAMPOS As String = "SelectorCampos"
Public Const TABLE_SELECTOR_CAMPOS As String = "tblSelectorCampos"

Public Const MIN_SUPPORTED_EXCEL_VERSION As Double = 14# 'Excel 2010 o superior

'--- Formatos estándar ---
Private Const FORMAT_DATE As String = "dd/mm/yyyy"
Private Const FORMAT_TIME As String = "hh:mm:ss"
Private Const FORMAT_TEXT As String = "@"
Private Const FORMAT_ACCOUNTING_EC As String = _
    "$ #,##0.00;-$ #,##0.00;$ 0.00"
Private Const FORMAT_AMOUNT As String = FORMAT_ACCOUNTING_EC
Private Const FORMAT_DECIMAL As String = "#,##0.0000"
Private Const FORMAT_PERCENT As String = "0.00%"

Private Const HEADER_FONT_COLOR As Long = vbWhite
Private Const HEADER_FILL_COLOR_HEX_DEFAULT As String = "#1C0F82"
Private Const HEADER_FILL_COLOR_RGB_DEFAULT As Long = &H820F1C ' Equivalente a RGB(28, 15, 130)
Private Const CORPORATE_TABLE_STYLE As String = "TableStyleLight1"

Private mHeaderFillColor As Long
Private mHeaderFillColorInitialized As Boolean

'--- Mensajes ---
Private Const MSG_SKIP_RESET As String = _
    "Se omitió la recreación de la tabla '%s' porque contiene datos. " & _
    "Use forceReset:=True para reestructurarla."
Private Const MSG_COLUMN_ORDER_WARNING As String = _
    "La tabla '%s' no coincide con el orden de columnas esperado. Considere ejecutar " & _
    "Setup_Estructura forceReset:=True para corregirla."

'------------------------------------------------------------------------------------------
' Procedimiento principal para generar la estructura de trabajo.
'------------------------------------------------------------------------------------------
Public Sub Setup_Estructura(Optional ByVal targetWorkbook As Workbook, _
                            Optional ByVal forceReset As Boolean = False)
    Dim wb As Workbook
    Dim calculationState As XlCalculation
    Dim screenUpdatingState As Boolean
    Dim enableEventsState As Boolean
    Dim statusBarState As Variant

    On Error GoTo CleanFail

    If targetWorkbook Is Nothing Then
        Set wb = ThisWorkbook
    Else
        Set wb = targetWorkbook
    End If

    InvalidateHeaderFillColorCache

    If Val(Application.Version) < MIN_SUPPORTED_EXCEL_VERSION Then
        Debug.Print "[Setup_Estructura] Advertencia: versión de Excel " & _
                    Application.Version & " detectada. Algunas características" & _
                    " pueden variar; pruebe en Excel 2010 o superior."
    End If

    calculationState = Application.Calculation
    screenUpdatingState = Application.ScreenUpdating
    enableEventsState = Application.EnableEvents
    statusBarState = Application.StatusBar

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Configurando estructura base..."

    ' Facturas y detalle
    EnsureSheetWithTable wb, SHEET_FACTURAS, TABLE_FACTURAS, BuildFacturasColumns(), _
                         CORPORATE_TABLE_STYLE, forceReset
    EnsureSheetWithTable wb, SHEET_FACTURAS_DETALLE, TABLE_FACTURAS_DETALLE, _
                         BuildFacturasDetalleColumns(), CORPORATE_TABLE_STYLE, forceReset

    ' Notas de crédito y detalle
    EnsureSheetWithTable wb, SHEET_NOTAS_CREDITO, TABLE_NOTAS_CREDITO, _
                         BuildNotasCreditoColumns(), CORPORATE_TABLE_STYLE, forceReset
    EnsureSheetWithTable wb, SHEET_NOTAS_CREDITO_DETALLE, TABLE_NOTAS_CREDITO_DETALLE, _
                         BuildNotasCreditoDetalleColumns(), CORPORATE_TABLE_STYLE, forceReset

    ' Retenciones
    EnsureSheetWithTable wb, SHEET_RETENCIONES, TABLE_RETENCIONES, _
                         BuildRetencionesColumns(), CORPORATE_TABLE_STYLE, forceReset
    EnsureSheetWithTable wb, SHEET_RETENCIONES_DETALLE, TABLE_RETENCIONES_DETALLE, _
                         BuildRetencionesDetalleColumns(), CORPORATE_TABLE_STYLE, forceReset

    ' Parámetros, log y hoja de control
    EnsureSheetWithTable wb, SHEET_PARAMETROS, TABLE_PARAMETROS, _
                         BuildParametrosColumns(), CORPORATE_TABLE_STYLE, forceReset
    EnsureSheetWithTable wb, SHEET_LOG, TABLE_LOG, _
                         BuildLogColumns(), CORPORATE_TABLE_STYLE, forceReset
    EnsureSheetWithTable wb, SHEET_SELECTOR_CAMPOS, TABLE_SELECTOR_CAMPOS, _
                         BuildSelectorCamposColumns(), CORPORATE_TABLE_STYLE, forceReset
    PrepareControlSheet wb, forceReset
    SeedParametros wb.Worksheets(SHEET_PARAMETROS).ListObjects(TABLE_PARAMETROS)
    SeedSelectorCampos wb.Worksheets(SHEET_SELECTOR_CAMPOS).ListObjects(TABLE_SELECTOR_CAMPOS)
    InvalidateHeaderFillColorCache
    RefreshAllTableHeaders wb
    RefreshCorporateTableHeaders wb

CleanExit:
    Application.StatusBar = statusBarState
    Application.Calculation = calculationState
    Application.EnableEvents = enableEventsState
    Application.ScreenUpdating = screenUpdatingState
    Exit Sub

CleanFail:
    Debug.Print "[Setup_Estructura] Error " & Err.Number & ": " & Err.Description
    Resume CleanExit
End Sub

'------------------------------------------------------------------------------------------
' Prueba rápida para ejecutar desde el explorador de macros.
'------------------------------------------------------------------------------------------
Public Sub Run_SmokeTest()
    Setup_Estructura forceReset:=True
    Debug.Print "Run_SmokeTest completado correctamente."
End Sub

'------------------------------------------------------------------------------------------
' Construcción de columnas por tabla
'------------------------------------------------------------------------------------------
Private Function BuildFacturasColumns() As Variant
    Dim specs As Collection
    Set specs = New Collection

    With specs
        .Add ColumnSpec("TipoComprobante", FORMAT_TEXT, 16, xlCenter)
        .Add ColumnSpec("Subtipo", FORMAT_TEXT, 12, xlCenter)
        .Add ColumnSpec("FechaEmision", FORMAT_DATE, 14, xlCenter)
        .Add ColumnSpec("FechaRegistro", FORMAT_DATE, 14, xlCenter)
        .Add ColumnSpec("RUC_Emisor", FORMAT_TEXT, 16, xlLeft)
        .Add ColumnSpec("TipoIdentificacionEmisor", FORMAT_TEXT, 16, xlCenter)
        .Add ColumnSpec("Nombre_Emisor", FORMAT_TEXT, 32, xlLeft)
        .Add ColumnSpec("RUC_Receptor", FORMAT_TEXT, 16, xlLeft)
        .Add ColumnSpec("TipoIdentificacionReceptor", FORMAT_TEXT, 16, xlCenter)
        .Add ColumnSpec("Nombre_Receptor", FORMAT_TEXT, 32, xlLeft)
        .Add ColumnSpec("NumeroDocumento", FORMAT_TEXT, 20, xlLeft)
        .Add ColumnSpec("Establecimiento", FORMAT_TEXT, 10, xlCenter)
        .Add ColumnSpec("PuntoEmision", FORMAT_TEXT, 10, xlCenter)
        .Add ColumnSpec("Secuencial", FORMAT_TEXT, 12, xlCenter)
        .Add ColumnSpec("ClaveAcceso", FORMAT_TEXT, 28, xlLeft)
        .Add ColumnSpec("Moneda", FORMAT_TEXT, 10, xlCenter)
        .Add ColumnSpec("BaseIVA15", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("BaseIVA12", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("BaseIVA8", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("BaseIVA5", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("BaseIVA0", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("BaseNoObjetoIVA", FORMAT_AMOUNT, 18, xlRight)
        .Add ColumnSpec("BaseExentaIVA", FORMAT_AMOUNT, 18, xlRight)
        .Add ColumnSpec("BaseICE", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("ValorIVA15", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("ValorIVA12", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("ValorIVA8", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("ValorIVA5", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("ValorIVA0", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("ValorIVAExento", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("ValorICE", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("ValorIRBPNR", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("Propina", FORMAT_AMOUNT, 14, xlRight)
        .Add ColumnSpec("TotalSinImpuestos", FORMAT_AMOUNT, 18, xlRight)
        .Add ColumnSpec("ValorTotal", FORMAT_AMOUNT, 18, xlRight)
        .Add ColumnSpec("Estado", FORMAT_TEXT, 12, xlCenter)
        .Add ColumnSpec("OrigenArchivo", FORMAT_TEXT, 24, xlLeft)
        .Add ColumnSpec("RutaArchivo", FORMAT_TEXT, 50, xlLeft)
        .Add ColumnSpec("NombreArchivo", FORMAT_TEXT, 28, xlLeft)
        .Add ColumnSpec("EnlaceComprobante", FORMAT_TEXT, 24, xlLeft, _
                        BuildHyperlinkFormula(), True)
        .Add ColumnSpec("CreadoEn", FORMAT_DATE, 14, xlCenter)
        .Add ColumnSpec("Observaciones", FORMAT_TEXT, 50, xlLeft)
    End With

    BuildFacturasColumns = ColumnSpecsToArray(specs)
End Function

Private Function BuildFacturasDetalleColumns() As Variant
    Dim specs As Collection
    Set specs = New Collection

    With specs
        .Add ColumnSpec("NumeroDocumento", FORMAT_TEXT, 20, xlLeft)
        .Add ColumnSpec("Linea", FORMAT_TEXT, 8, xlCenter)
        .Add ColumnSpec("CodigoPrincipal", FORMAT_TEXT, 20, xlLeft)
        .Add ColumnSpec("CodigoAuxiliar", FORMAT_TEXT, 20, xlLeft)
        .Add ColumnSpec("Descripcion", FORMAT_TEXT, 48, xlLeft)
        .Add ColumnSpec("Cantidad", FORMAT_DECIMAL, 14, xlRight)
        .Add ColumnSpec("UnidadMedida", FORMAT_TEXT, 14, xlCenter)
        .Add ColumnSpec("PrecioUnitario", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("Descuento", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("PrecioTotalSinImpuesto", FORMAT_AMOUNT, 20, xlRight)
        .Add ColumnSpec("TarifaIVA", FORMAT_PERCENT, 12, xlRight)
        .Add ColumnSpec("PorcentajeIVA", FORMAT_PERCENT, 12, xlRight)
        .Add ColumnSpec("BaseIVA", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("ValorIVA", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("TarifaICE", FORMAT_PERCENT, 12, xlRight)
        .Add ColumnSpec("ValorICE", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("DetalleAdicional1", FORMAT_TEXT, 24, xlLeft)
        .Add ColumnSpec("DetalleAdicional2", FORMAT_TEXT, 24, xlLeft)
        .Add ColumnSpec("DetalleAdicional3", FORMAT_TEXT, 24, xlLeft)
    End With

    BuildFacturasDetalleColumns = ColumnSpecsToArray(specs)
End Function

Private Function BuildNotasCreditoColumns() As Variant
    Dim specs As Collection
    Set specs = New Collection

    With specs
        .Add ColumnSpec("TipoComprobante", FORMAT_TEXT, 16, xlCenter)
        .Add ColumnSpec("FechaEmision", FORMAT_DATE, 14, xlCenter)
        .Add ColumnSpec("FechaRegistro", FORMAT_DATE, 14, xlCenter)
        .Add ColumnSpec("RUC_Emisor", FORMAT_TEXT, 16, xlLeft)
        .Add ColumnSpec("TipoIdentificacionEmisor", FORMAT_TEXT, 16, xlCenter)
        .Add ColumnSpec("Nombre_Emisor", FORMAT_TEXT, 32, xlLeft)
        .Add ColumnSpec("RUC_Receptor", FORMAT_TEXT, 16, xlLeft)
        .Add ColumnSpec("TipoIdentificacionReceptor", FORMAT_TEXT, 16, xlCenter)
        .Add ColumnSpec("Nombre_Receptor", FORMAT_TEXT, 32, xlLeft)
        .Add ColumnSpec("NumeroDocumento", FORMAT_TEXT, 20, xlLeft)
        .Add ColumnSpec("DocumentoModifica", FORMAT_TEXT, 20, xlLeft)
        .Add ColumnSpec("Motivo", FORMAT_TEXT, 40, xlLeft)
        .Add ColumnSpec("ClaveAcceso", FORMAT_TEXT, 28, xlLeft)
        .Add ColumnSpec("BaseIVA15", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("BaseIVA12", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("BaseIVA0", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("BaseNoObjetoIVA", FORMAT_AMOUNT, 18, xlRight)
        .Add ColumnSpec("BaseExentaIVA", FORMAT_AMOUNT, 18, xlRight)
        .Add ColumnSpec("ValorIVA15", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("ValorIVA12", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("ValorIVA0", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("ValorTotal", FORMAT_AMOUNT, 18, xlRight)
        .Add ColumnSpec("OrigenArchivo", FORMAT_TEXT, 24, xlLeft)
        .Add ColumnSpec("RutaArchivo", FORMAT_TEXT, 50, xlLeft)
        .Add ColumnSpec("NombreArchivo", FORMAT_TEXT, 28, xlLeft)
        .Add ColumnSpec("EnlaceComprobante", FORMAT_TEXT, 24, xlLeft, _
                        BuildHyperlinkFormula(), True)
        .Add ColumnSpec("CreadoEn", FORMAT_DATE, 14, xlCenter)
    End With

    BuildNotasCreditoColumns = ColumnSpecsToArray(specs)
End Function

Private Function BuildNotasCreditoDetalleColumns() As Variant
    Dim specs As Collection
    Set specs = New Collection

    With specs
        .Add ColumnSpec("NumeroDocumento", FORMAT_TEXT, 20, xlLeft)
        .Add ColumnSpec("Linea", FORMAT_TEXT, 8, xlCenter)
        .Add ColumnSpec("CodigoPrincipal", FORMAT_TEXT, 20, xlLeft)
        .Add ColumnSpec("CodigoAuxiliar", FORMAT_TEXT, 20, xlLeft)
        .Add ColumnSpec("Descripcion", FORMAT_TEXT, 48, xlLeft)
        .Add ColumnSpec("Cantidad", FORMAT_DECIMAL, 14, xlRight)
        .Add ColumnSpec("UnidadMedida", FORMAT_TEXT, 14, xlCenter)
        .Add ColumnSpec("PrecioUnitario", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("Descuento", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("PrecioTotalSinImpuesto", FORMAT_AMOUNT, 20, xlRight)
        .Add ColumnSpec("TarifaIVA", FORMAT_PERCENT, 12, xlRight)
        .Add ColumnSpec("PorcentajeIVA", FORMAT_PERCENT, 12, xlRight)
        .Add ColumnSpec("BaseIVA", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("ValorIVA", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("TarifaICE", FORMAT_PERCENT, 12, xlRight)
        .Add ColumnSpec("ValorICE", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("DetalleAdicional1", FORMAT_TEXT, 24, xlLeft)
        .Add ColumnSpec("DetalleAdicional2", FORMAT_TEXT, 24, xlLeft)
        .Add ColumnSpec("DetalleAdicional3", FORMAT_TEXT, 24, xlLeft)
    End With

    BuildNotasCreditoDetalleColumns = ColumnSpecsToArray(specs)
End Function

Private Function BuildRetencionesColumns() As Variant
    Dim specs As Collection
    Set specs = New Collection

    With specs
        .Add ColumnSpec("TipoComprobante", FORMAT_TEXT, 16, xlCenter)
        .Add ColumnSpec("FechaEmision", FORMAT_DATE, 14, xlCenter)
        .Add ColumnSpec("FechaRegistro", FORMAT_DATE, 14, xlCenter)
        .Add ColumnSpec("RUC_Agente", FORMAT_TEXT, 16, xlLeft)
        .Add ColumnSpec("TipoIdentificacionAgente", FORMAT_TEXT, 16, xlCenter)
        .Add ColumnSpec("Nombre_Agente", FORMAT_TEXT, 32, xlLeft)
        .Add ColumnSpec("RUC_Sujeto", FORMAT_TEXT, 16, xlLeft)
        .Add ColumnSpec("TipoIdentificacionSujeto", FORMAT_TEXT, 16, xlCenter)
        .Add ColumnSpec("Nombre_Sujeto", FORMAT_TEXT, 32, xlLeft)
        .Add ColumnSpec("NumeroDocumento", FORMAT_TEXT, 20, xlLeft)
        .Add ColumnSpec("PeriodoFiscal", FORMAT_TEXT, 12, xlCenter)
        .Add ColumnSpec("DocumentoSustento", FORMAT_TEXT, 20, xlLeft)
        .Add ColumnSpec("NumeroSustento", FORMAT_TEXT, 20, xlLeft)
        .Add ColumnSpec("ClaveAcceso", FORMAT_TEXT, 28, xlLeft)
        .Add ColumnSpec("Moneda", FORMAT_TEXT, 10, xlCenter)
        .Add ColumnSpec("BaseIVA", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("PorcentajeIVA", FORMAT_PERCENT, 12, xlRight)
        .Add ColumnSpec("ValorRetenidoIVA", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("BaseRenta", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("PorcentajeRenta", FORMAT_PERCENT, 12, xlRight)
        .Add ColumnSpec("ValorRetenidoRenta", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("BaseISD", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("PorcentajeISD", FORMAT_PERCENT, 12, xlRight)
        .Add ColumnSpec("ValorRetenidoISD", FORMAT_AMOUNT, 16, xlRight)
        .Add ColumnSpec("OrigenArchivo", FORMAT_TEXT, 24, xlLeft)
        .Add ColumnSpec("RutaArchivo", FORMAT_TEXT, 50, xlLeft)
        .Add ColumnSpec("NombreArchivo", FORMAT_TEXT, 28, xlLeft)
        .Add ColumnSpec("EnlaceComprobante", FORMAT_TEXT, 24, xlLeft, _
                        BuildHyperlinkFormula(), True)
        .Add ColumnSpec("CreadoEn", FORMAT_DATE, 14, xlCenter)
    End With

    BuildRetencionesColumns = ColumnSpecsToArray(specs)
End Function

Private Function BuildRetencionesDetalleColumns() As Variant
    BuildRetencionesDetalleColumns = Array( _
        ColumnSpec("NumeroDocumento", FORMAT_TEXT, 20, xlLeft), _
        ColumnSpec("Linea", FORMAT_TEXT, 8, xlCenter), _
        ColumnSpec("Impuesto", FORMAT_TEXT, 12, xlCenter), _
        ColumnSpec("CodigoRetencion", FORMAT_TEXT, 16, xlLeft), _
        ColumnSpec("Descripcion", FORMAT_TEXT, 48, xlLeft), _
        ColumnSpec("BaseImponible", FORMAT_AMOUNT, 18, xlRight), _
        ColumnSpec("PorcentajeRetencion", FORMAT_PERCENT, 12, xlRight), _
        ColumnSpec("ValorRetenido", FORMAT_AMOUNT, 16, xlRight), _
        ColumnSpec("TipoDocumentoSustento", FORMAT_TEXT, 18, xlCenter), _
        ColumnSpec("NumeroDocumentoSustento", FORMAT_TEXT, 20, xlLeft), _
        ColumnSpec("FechaEmisionSustento", FORMAT_DATE, 14, xlCenter))
End Function

Private Function BuildParametrosColumns() As Variant
    BuildParametrosColumns = Array( _
        ColumnSpec("Parametro", FORMAT_TEXT, 28, xlLeft), _
        ColumnSpec("Valor", FORMAT_TEXT, 36, xlLeft), _
        ColumnSpec("Descripcion", FORMAT_TEXT, 60, xlLeft), _
        ColumnSpec("UltimaActualizacion", FORMAT_DATE, 14, xlCenter))
End Function

Private Function BuildLogColumns() As Variant
    BuildLogColumns = Array( _
        ColumnSpec("MarcaTiempo", FORMAT_DATE, 14, xlCenter), _
        ColumnSpec("Hora", FORMAT_TIME, 12, xlCenter), _
        ColumnSpec("Nivel", FORMAT_TEXT, 10, xlCenter), _
        ColumnSpec("Codigo", FORMAT_TEXT, 12, xlCenter), _
        ColumnSpec("Descripcion", FORMAT_TEXT, 60, xlLeft), _
        ColumnSpec("RutaArchivo", FORMAT_TEXT, 50, xlLeft), _
        ColumnSpec("Detalle", FORMAT_TEXT, 80, xlLeft))
End Function

Private Function BuildSelectorCamposColumns() As Variant
    BuildSelectorCamposColumns = Array( _
        ColumnSpec("TablaDestino", FORMAT_TEXT, 20, xlLeft), _
        ColumnSpec("Campo", FORMAT_TEXT, 28, xlLeft), _
        ColumnSpec("Descripcion", FORMAT_TEXT, 60, xlLeft), _
        ColumnSpec("Importar", FORMAT_TEXT, 12, xlCenter))
End Function

'------------------------------------------------------------------------------------------
' Helpers de estructura
'------------------------------------------------------------------------------------------
Private Function BuildHyperlinkFormula(Optional ByVal friendlyField As String = "NombreArchivo") As String
    Dim q As String
    q = Chr$(34)

    BuildHyperlinkFormula = "=IF([@RutaArchivo]=" & q & q & "," & q & q & _
                              ",HYPERLINK([@RutaArchivo],IF([@" & friendlyField & "]=" & q & q & _
                              "," & q & "Abrir XML" & q & ",[@" & friendlyField & "])))"
End Function

Private Function ColumnSpec(ByVal header As String, _
                            Optional ByVal numberFormat As String = "", _
                            Optional ByVal columnWidth As Double = 18, _
                            Optional ByVal horizontalAlignment As Long = xlLeft, _
                            Optional ByVal calculatedFormula As String = vbNullString, _
                            Optional ByVal asHyperlink As Boolean = False) As Variant
    Dim spec(0 To 5) As Variant
    spec(0) = header
    spec(1) = numberFormat
    spec(2) = columnWidth
    spec(3) = horizontalAlignment
    spec(4) = calculatedFormula
    spec(5) = asHyperlink
    ColumnSpec = spec
End Function

Private Function ColumnSpecsToArray(ByVal specs As Collection) As Variant
    Dim arr() As Variant
    Dim idx As Long

    If specs Is Nothing Then
        ColumnSpecsToArray = VBA.Array()
        Exit Function
    End If

    If specs.Count = 0 Then
        ColumnSpecsToArray = VBA.Array()
        Exit Function
    End If

    ReDim arr(0 To specs.Count - 1)

    For idx = 1 To specs.Count
        arr(idx - 1) = specs(idx)
    Next idx

    ColumnSpecsToArray = arr
End Function

Private Sub EnsureSheetWithTable(ByVal wb As Workbook, _
                                 ByVal sheetName As String, _
                                 ByVal tableName As String, _
                                 ByVal columnSpecs As Variant, _
                                 Optional ByVal tableStyle As String = CORPORATE_TABLE_STYLE, _
                                 Optional ByVal forceReset As Boolean = False)
    Dim ws As Worksheet
    Dim lo As ListObject

    Set ws = EnsureWorksheet(wb, sheetName)

    Set lo = GetListObject(ws, tableName)

    If lo Is Nothing Then
        ResetWorksheetTable ws, tableName, columnSpecs, tableStyle
    Else
        If forceReset Then
            ResetWorksheetTable ws, tableName, columnSpecs, tableStyle
        ElseIf lo.ListRows.Count > 0 Then
            Debug.Print Format$(Now, "dd/mm/yyyy hh:nn:ss") & " - " & _
                        VBA.Replace(MSG_SKIP_RESET, "%s", tableName)
            EnsureTableColumns lo, columnSpecs
            ClearTableBodyFormatting lo
            FormatTableHeader lo, wb
        Else
            ResetWorksheetTable ws, tableName, columnSpecs, tableStyle
        End If
    End If
End Sub

Private Sub ResetWorksheetTable(ByVal ws As Worksheet, _
                                ByVal tableName As String, _
                                ByVal columnSpecs As Variant, _
                                ByVal tableStyle As String)
    Dim lo As ListObject
    Dim colCount As Long
    Dim headers As Variant
    Dim idx As Long
    Dim arrIndex As Long
    Dim headerIndex As Long

    Dim headerFillColor As Long

    On Error Resume Next
    For Each lo In ws.ListObjects
        lo.Delete
    Next lo
    On Error GoTo 0

    headerFillColor = GetWorkbookHeaderFillColor(ws.Parent)

    ws.Cells.Clear
    ws.Cells.VerticalAlignment = xlCenter
    ws.Cells.WrapText = False
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 11

    colCount = UBound(columnSpecs) - LBound(columnSpecs) + 1
    ReDim headers(1 To 1, 1 To colCount)

    headerIndex = 1
    For arrIndex = LBound(columnSpecs) To UBound(columnSpecs)
        headers(1, headerIndex) = CStr(columnSpecs(arrIndex)(0))
        headerIndex = headerIndex + 1
    Next arrIndex

    With ws
        .Range("A1").Resize(1, colCount).Value = headers
        .Rows(1).Font.Bold = True
        .Rows(1).Font.Color = HEADER_FONT_COLOR
        With .Rows(1).Interior
            .Pattern = xlSolid
            .TintAndShade = 0
            .Color = headerFillColor
        End With
        .Rows(1).RowHeight = 18
        Set lo = .ListObjects.Add(SourceType:=xlSrcRange, _
                                   Source:=.Range("A1").Resize(1, colCount), _
                                   XlListObjectHasHeaders:=xlYes)
    End With

    lo.Name = tableName
    ApplyTableStyle lo, tableStyle
    lo.ShowAutoFilter = True

    ApplyColumnSpecs lo, columnSpecs
    ClearTableBodyFormatting lo
    FormatTableHeader lo, ws.Parent

    ws.Columns.AutoFit
    ws.Cells(1, 1).Select
End Sub

Private Sub EnsureTableColumns(ByVal lo As ListObject, ByVal columnSpecs As Variant)
    Dim arrIndex As Long
    Dim spec As Variant
    Dim targetColumn As ListColumn
    Dim desiredIndex As Long
    Dim mismatchFound As Boolean

    For arrIndex = LBound(columnSpecs) To UBound(columnSpecs)
        spec = columnSpecs(arrIndex)
        desiredIndex = arrIndex - LBound(columnSpecs) + 1
        Set targetColumn = GetListColumn(lo, CStr(spec(0)))

        If targetColumn Is Nothing Then
            Set targetColumn = lo.ListColumns.Add(Position:=desiredIndex)
            targetColumn.Name = CStr(spec(0))
        ElseIf targetColumn.Index <> desiredIndex Then
            mismatchFound = True
        End If

        ApplyColumnFormat targetColumn, spec
    Next arrIndex

    If mismatchFound Then
        Debug.Print Format$(Now, "dd/mm/yyyy hh:nn:ss") & " - " & _
                    VBA.Replace(MSG_COLUMN_ORDER_WARNING, "%s", lo.Name)
    End If
End Sub

Private Sub ApplyColumnSpecs(ByVal lo As ListObject, ByVal columnSpecs As Variant)
    Dim arrIndex As Long
    Dim spec As Variant
    Dim targetColumn As ListColumn

    For arrIndex = LBound(columnSpecs) To UBound(columnSpecs)
        spec = columnSpecs(arrIndex)
        Set targetColumn = lo.ListColumns(arrIndex - LBound(columnSpecs) + 1)
        targetColumn.Name = CStr(spec(0))
        ApplyColumnFormat targetColumn, spec
    Next arrIndex
End Sub

Private Sub ApplyColumnFormat(ByVal targetColumn As ListColumn, ByVal spec As Variant)
    Dim columnRange As Range

    Set columnRange = targetColumn.Range

    With columnRange
        If Len(CStr(spec(1))) > 0 Then
            .NumberFormat = CStr(spec(1))
        End If
        .HorizontalAlignment = spec(3)
        .VerticalAlignment = xlCenter
        .WrapText = False
    End With

    If spec(2) > 0 Then
        columnRange.EntireColumn.ColumnWidth = spec(2)
    End If

    EnsureCalculatedColumn targetColumn, spec
End Sub

Private Sub EnsureCalculatedColumn(ByVal targetColumn As ListColumn, ByVal spec As Variant)
    Dim formulaText As String
    Dim asHyperlink As Boolean
    Dim lo As ListObject
    Dim hadRows As Boolean
    Dim tempRow As ListRow
    Dim dataRange As Range

    If UBound(spec) >= 4 Then
        formulaText = CStr(spec(4))
    Else
        formulaText = vbNullString
    End If

    If UBound(spec) >= 5 Then
        asHyperlink = CBool(spec(5))
    Else
        asHyperlink = False
    End If

    If Len(formulaText) = 0 And Not asHyperlink Then Exit Sub

    Set lo = targetColumn.Parent
    hadRows = (lo.ListRows.Count > 0)

    If Len(formulaText) > 0 And Not hadRows Then
        Set tempRow = lo.ListRows.Add
    End If

    Set dataRange = targetColumn.DataBodyRange

    If Not dataRange Is Nothing Then
        If Len(formulaText) > 0 Then
            dataRange.Formula = formulaText
        End If
        If asHyperlink Then
            On Error Resume Next
            dataRange.Style = "Hyperlink"
            On Error GoTo 0
        End If
    End If

    If Not hadRows And Not tempRow Is Nothing Then
        tempRow.Delete
    End If
End Sub

Private Function EnsureWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    End If

    ws.Cells.VerticalAlignment = xlCenter
    ws.Cells.WrapText = False

    Set EnsureWorksheet = ws
End Function

Private Function GetListObject(ByVal ws As Worksheet, ByVal tableName As String) As ListObject
    On Error Resume Next
    Set GetListObject = ws.ListObjects(tableName)
    On Error GoTo 0
End Function

Private Function GetListColumn(ByVal lo As ListObject, ByVal headerName As String) As ListColumn
    On Error Resume Next
    Set GetListColumn = lo.ListColumns(headerName)
    On Error GoTo 0
End Function

Private Sub ApplyTableStyle(ByVal lo As ListObject, ByVal styleName As String)
    On Error GoTo Fallback

    Dim targetStyle As String

    targetStyle = CORPORATE_TABLE_STYLE

    If Len(styleName) > 0 And TableStyleExists(styleName) Then
        targetStyle = styleName
    End If

    If TableStyleExists(targetStyle) Then
        lo.TableStyle = targetStyle
    Else
        lo.TableStyle = CORPORATE_TABLE_STYLE
    End If

    lo.ShowTableStyleRowStripes = False
    lo.ShowTableStyleColumnStripes = False
    lo.ShowTableStyleFirstColumn = False
    lo.ShowTableStyleLastColumn = False

    Exit Sub

Fallback:
    On Error Resume Next
    lo.TableStyle = CORPORATE_TABLE_STYLE
    lo.ShowTableStyleRowStripes = False
    lo.ShowTableStyleColumnStripes = False
    lo.ShowTableStyleFirstColumn = False
    lo.ShowTableStyleLastColumn = False
    On Error GoTo 0
End Sub

Private Sub ClearTableBodyFormatting(ByVal lo As ListObject)
    Dim bodyRange As Range
    Dim placeholderRange As Range

    If lo Is Nothing Then Exit Sub

    On Error Resume Next
    Set bodyRange = lo.DataBodyRange
    On Error GoTo 0

    If Not bodyRange Is Nothing Then
        With bodyRange
            .Interior.ColorIndex = xlColorIndexNone
            .Font.ColorIndex = xlColorIndexAutomatic
        End With
    Else
        Set placeholderRange = GetTablePlaceholderRange(lo)
        If Not placeholderRange Is Nothing Then
            With placeholderRange
                .Interior.ColorIndex = xlColorIndexNone
                .Font.ColorIndex = xlColorIndexAutomatic
            End With
        End If
    End If

    On Error Resume Next
    If Not lo.TotalsRowRange Is Nothing Then
        With lo.TotalsRowRange
            .Interior.ColorIndex = xlColorIndexNone
            .Font.ColorIndex = xlColorIndexAutomatic
        End With
    End If
    On Error GoTo 0
End Sub

Private Function GetTablePlaceholderRange(ByVal lo As ListObject) As Range
    Dim headerRows As Long

    If lo Is Nothing Then Exit Function

    headerRows = 1
    If lo.ShowHeaders = False Then
        headerRows = 0
    End If

    On Error Resume Next
    If lo.Range.Rows.Count > headerRows Then
        Set GetTablePlaceholderRange = lo.Range.Rows(headerRows + 1)
    End If
    On Error GoTo 0
End Function

Private Sub FormatTableHeader(ByVal lo As ListObject, Optional ByVal wb As Workbook = Nothing)
    Dim targetWorkbook As Workbook
    Dim headerColor As Long
    Dim headerRange As Range

    On Error GoTo CleanExit

    If lo Is Nothing Then GoTo CleanExit

    If lo.ShowHeaders = False Then
        On Error Resume Next
        lo.ShowHeaders = True
        On Error GoTo CleanExit
    End If

    Set headerRange = Nothing

    On Error Resume Next
    Set headerRange = lo.HeaderRowRange
    On Error GoTo CleanExit

    If headerRange Is Nothing Then
        Set headerRange = lo.Range.Rows(1)
    End If

    If headerRange Is Nothing Then GoTo CleanExit

    If wb Is Nothing Then
        Set targetWorkbook = GetListObjectWorkbook(lo)
    Else
        Set targetWorkbook = wb
    End If

    headerColor = GetWorkbookHeaderFillColor(targetWorkbook)

    If headerColor <= 0 Then
        headerColor = GetDefaultHeaderFillColor()
    End If

    With headerRange
        .Interior.Pattern = xlSolid
        .Interior.TintAndShade = 0
        .Interior.Color = headerColor
        .Font.Color = HEADER_FONT_COLOR
        .Font.Bold = True
    End With

CleanExit:
    On Error GoTo 0
End Sub

Private Function GetListObjectWorkbook(ByVal lo As ListObject) As Workbook
    On Error Resume Next
    Set GetListObjectWorkbook = lo.Parent.Parent
    On Error GoTo 0
End Function

Private Function GetWorkbookHeaderFillColor(ByVal wb As Workbook) As Long
    If wb Is Nothing Then
        GetWorkbookHeaderFillColor = GetDefaultHeaderFillColor()
        Exit Function
    End If

    If Not mHeaderFillColorInitialized Then
        mHeaderFillColor = ResolveHeaderFillColor(wb)
        mHeaderFillColorInitialized = True
    End If

    GetWorkbookHeaderFillColor = mHeaderFillColor
End Function

Private Function ResolveHeaderFillColor(ByVal wb As Workbook) As Long
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim value As Variant
    Dim parsedColor As Long

    On Error GoTo Fallback

    Set ws = wb.Worksheets(SHEET_PARAMETROS)
    Set lo = ws.ListObjects(TABLE_PARAMETROS)

    value = LookupParameterValue(lo, "ColorEncabezadoHEX")
    If TryParseHexColor(value, parsedColor) Then
        ResolveHeaderFillColor = parsedColor
        Exit Function
    End If

    value = LookupParameterValue(lo, "ColorEncabezadoRGB")
    If IsNumeric(value) Then
        ResolveHeaderFillColor = CLng(value)
        Exit Function
    End If

Fallback:
    ResolveHeaderFillColor = GetDefaultHeaderFillColor()
End Function

Private Function LookupParameterValue(ByVal lo As ListObject, ByVal parametro As String) As Variant
    Dim parametroCol As ListColumn
    Dim valorCol As ListColumn
    Dim parametros As Range
    Dim valores As Range
    Dim i As Long

    If lo Is Nothing Then Exit Function

    On Error Resume Next
    Set parametroCol = lo.ListColumns("Parametro")
    Set valorCol = lo.ListColumns("Valor")
    On Error GoTo 0

    If parametroCol Is Nothing Or valorCol Is Nothing Then Exit Function

    Set parametros = parametroCol.DataBodyRange
    Set valores = valorCol.DataBodyRange

    If parametros Is Nothing Or valores Is Nothing Then Exit Function

    For i = 1 To parametros.Rows.Count
        If StrComp(CStr(parametros.Cells(i, 1).Value), parametro, vbTextCompare) = 0 Then
            LookupParameterValue = valores.Cells(i, 1).Value
            Exit Function
        End If
    Next i
End Function

Private Function TryParseHexColor(ByVal candidate As Variant, ByRef colorValue As Long) As Boolean
    Dim cleaned As String
    Dim r As Long, g As Long, b As Long
    Dim i As Long
    Dim hexDigits As String

    If IsError(candidate) Then Exit Function
    If IsNull(candidate) Then Exit Function

    cleaned = Trim$(CStr(candidate))
    If Len(cleaned) = 0 Then Exit Function

    If Left$(cleaned, 2) = "0x" Or Left$(cleaned, 2) = "0X" Then
        cleaned = Mid$(cleaned, 3)
    End If

    If Left$(cleaned, 1) = "#" Then
        cleaned = Mid$(cleaned, 2)
    End If

    If Len(cleaned) <> 6 Then Exit Function

    hexDigits = "0123456789ABCDEF"

    For i = 1 To 6
        If InStr(1, hexDigits, UCase$(Mid$(cleaned, i, 1)), vbBinaryCompare) = 0 Then
            Exit Function
        End If
    Next i

    r = CLng("&H" & Mid$(cleaned, 1, 2))
    g = CLng("&H" & Mid$(cleaned, 3, 2))
    b = CLng("&H" & Mid$(cleaned, 5, 2))

    colorValue = RGB(r, g, b)
    TryParseHexColor = True
End Function

Private Function GetDefaultHeaderFillColor() As Long
    GetDefaultHeaderFillColor = HEADER_FILL_COLOR_RGB_DEFAULT
End Function

Private Sub InvalidateHeaderFillColorCache()
    mHeaderFillColorInitialized = False
End Sub

Private Sub RefreshAllTableHeaders(ByVal wb As Workbook)
    Dim ws As Worksheet
    Dim lo As ListObject

    If wb Is Nothing Then Exit Sub

    For Each ws In wb.Worksheets
        On Error Resume Next
        For Each lo In ws.ListObjects
            ClearTableBodyFormatting lo
            FormatTableHeader lo, wb
        Next lo
        On Error GoTo 0
    Next ws
End Sub

Private Sub RefreshCorporateTableHeaders(ByVal wb As Workbook)
    If wb Is Nothing Then Exit Sub

    ApplyHeaderFormatting wb, SHEET_FACTURAS, TABLE_FACTURAS
    ApplyHeaderFormatting wb, SHEET_FACTURAS_DETALLE, TABLE_FACTURAS_DETALLE
    ApplyHeaderFormatting wb, SHEET_NOTAS_CREDITO, TABLE_NOTAS_CREDITO
    ApplyHeaderFormatting wb, SHEET_NOTAS_CREDITO_DETALLE, TABLE_NOTAS_CREDITO_DETALLE
    ApplyHeaderFormatting wb, SHEET_RETENCIONES, TABLE_RETENCIONES
    ApplyHeaderFormatting wb, SHEET_RETENCIONES_DETALLE, TABLE_RETENCIONES_DETALLE
    ApplyHeaderFormatting wb, SHEET_PARAMETROS, TABLE_PARAMETROS
    ApplyHeaderFormatting wb, SHEET_LOG, TABLE_LOG
    ApplyHeaderFormatting wb, SHEET_SELECTOR_CAMPOS, TABLE_SELECTOR_CAMPOS
End Sub

Private Sub ApplyHeaderFormatting(ByVal wb As Workbook, _
                                  ByVal sheetName As String, _
                                  ByVal tableName As String)
    Dim ws As Worksheet
    Dim lo As ListObject

    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    Set lo = GetListObject(ws, tableName)
    If lo Is Nothing Then Exit Sub

    ClearTableBodyFormatting lo
    FormatTableHeader lo, wb
End Sub

Private Function TableStyleExists(ByVal styleName As String) As Boolean
    Dim ts As TableStyle

    On Error GoTo CleanExit

    For Each ts In Application.TableStyles
        If StrComp(ts.Name, styleName, vbTextCompare) = 0 Then
            TableStyleExists = True
            Exit Function
        End If
    Next ts

CleanExit:
    On Error GoTo 0
End Function

'------------------------------------------------------------------------------------------
' Hoja de control y parámetros iniciales
'------------------------------------------------------------------------------------------
Private Sub PrepareControlSheet(ByVal wb As Workbook, ByVal forceReset As Boolean)
    Dim ws As Worksheet
    Dim needsReset As Boolean

    Set ws = EnsureWorksheet(wb, SHEET_CONTROL)

    needsReset = (ws.ListObjects.Count = 0 And WorksheetFunction.CountA(ws.Cells) = 0) Or forceReset

    If needsReset Then
        ws.Cells.Clear
        ws.Range("A1").Value = "Configuración"
        ws.Range("A1").Font.Bold = True
        ws.Range("A3").Value = "1. Ejecuta 'Setup_Estructura' para regenerar tablas."
        ws.Range("A4").Value = "2. Usa 'Importar_XML_SRI' para cargar XML." & vbLf & _
                                "   - Puedes incluir subcarpetas." & vbLf & _
                                "   - Filtra por fechas para importaciones parciales."
        ws.Range("A7").Value = "Notas"
        ws.Range("A7").Font.Bold = True
        ws.Range("A8").Value = "- Las fechas se formatean en dd/mm/yyyy." & vbLf & _
                                "- Los montos usan el formato contable $ de Ecuador." & vbLf & _
                                "- Los RUC y claves de acceso se mantienen como texto." & vbLf & _
                                "- Revisa las hojas 'Parametros' y 'SelectorCampos' para ajustar la importación." & vbLf & _
                                "- Compatible desde Excel 2010 (o superior)."
        ws.Columns("A:B").EntireColumn.AutoFit
    End If
End Sub

Private Sub SeedParametros(ByVal loParametros As ListObject)
    Dim existingRows As Long
    Dim data As Variant

    If loParametros Is Nothing Then Exit Sub

    existingRows = loParametros.ListRows.Count

    If existingRows = 0 Then
        data = Array( _
            Array("CarpetaXML", "", "Ruta base para buscar los XML del SRI"), _
            Array("IncluirSubcarpetas", "SI", "Valores permitidos: SI / NO"), _
            Array("FiltrarFechaDesde", "", "Fecha inicial (dd/mm/yyyy) para importación"), _
            Array("FiltrarFechaHasta", "", "Fecha final (dd/mm/yyyy) para importación"), _
            Array("RUC_Empresa", "", "RUC del contribuyente propietario del libro"), _
            Array("NombreComercial", "", "Nombre comercial del contribuyente"), _
            Array("ColorEncabezadoHEX", HEADER_FILL_COLOR_HEX_DEFAULT, "Color hexadecimal para los encabezados de tablas"), _
            Array("ColorEncabezadoRGB", "", "Alternativa en valor RGB decimal para encabezados"), _
            Array("UltimoProceso", "", "Fecha del último proceso de importación"), _
            Array("VersionMinimaExcel", "2010", "La plantilla está validada desde Excel 2010 en adelante"))
        AppendArrayToTable loParametros, data
    Else
        EnsureHeaderColorParametro loParametros
    End If
End Sub

Private Sub EnsureHeaderColorParametro(ByVal loParametros As ListObject)
    Dim parametroCol As ListColumn
    Dim valorCol As ListColumn
    Dim descripcionCol As ListColumn
    Dim parametros As Range
    Dim valores As Range
    Dim descripciones As Range
    Dim i As Long
    Dim cleanedValue As String
    Dim foundHex As Boolean
    Dim foundRgb As Boolean

    If loParametros Is Nothing Then Exit Sub

    On Error Resume Next
    Set parametroCol = loParametros.ListColumns("Parametro")
    Set valorCol = loParametros.ListColumns("Valor")
    Set descripcionCol = loParametros.ListColumns("Descripcion")
    On Error GoTo 0

    If parametroCol Is Nothing Or valorCol Is Nothing Then Exit Sub

    Set parametros = parametroCol.DataBodyRange
    Set valores = valorCol.DataBodyRange

    If Not descripcionCol Is Nothing Then
        Set descripciones = descripcionCol.DataBodyRange
    End If

    If parametros Is Nothing Or valores Is Nothing Then Exit Sub

    For i = 1 To parametros.Rows.Count
        If StrComp(CStr(parametros.Cells(i, 1).Value), "ColorEncabezadoHEX", vbTextCompare) = 0 Then
            cleanedValue = Trim$(CStr(valores.Cells(i, 1).Value))
            If StrComp(cleanedValue, HEADER_FILL_COLOR_HEX_DEFAULT, vbTextCompare) <> 0 Then
                valores.Cells(i, 1).Value = HEADER_FILL_COLOR_HEX_DEFAULT
            End If
            If Not descripciones Is Nothing Then
                If Len(Trim$(CStr(descripciones.Cells(i, 1).Value))) = 0 Then
                    descripciones.Cells(i, 1).Value = "Color hexadecimal para los encabezados de tablas"
                End If
            End If
            foundHex = True
        ElseIf StrComp(CStr(parametros.Cells(i, 1).Value), "ColorEncabezadoRGB", vbTextCompare) = 0 Then
            If Len(Trim$(CStr(valores.Cells(i, 1).Value))) > 0 Then
                valores.Cells(i, 1).Value = vbNullString
            End If
            If Not descripciones Is Nothing Then
                If Len(Trim$(CStr(descripciones.Cells(i, 1).Value))) = 0 Then
                    descripciones.Cells(i, 1).Value = "Alternativa en valor RGB decimal para encabezados"
                End If
            End If
            foundRgb = True
        End If
    Next i

    If Not foundHex Then
        AppendArrayToTable loParametros, Array(Array("ColorEncabezadoHEX", HEADER_FILL_COLOR_HEX_DEFAULT, _
                                                     "Color hexadecimal para los encabezados de tablas"))
    End If

    If Not foundRgb Then
        AppendArrayToTable loParametros, Array(Array("ColorEncabezadoRGB", vbNullString, _
                                                     "Alternativa en valor RGB decimal para encabezados"))
    End If
End Sub

Private Sub SeedSelectorCampos(ByVal loSelector As ListObject)
    Dim existingRows As Long
    Dim data As Variant

    If loSelector Is Nothing Then Exit Sub

    existingRows = loSelector.ListRows.Count

    If existingRows = 0 Then
        data = Array( _
            Array("Facturas", "TipoComprobante", "Código del comprobante emitido por el SRI", "SI"), _
            Array("Facturas", "Subtipo", "Detalle del tipo de comprobante si aplica", "SI"), _
            Array("Facturas", "FechaEmision", "Fecha de emisión del comprobante", "SI"), _
            Array("Facturas", "RUC_Emisor", "RUC o identificación del emisor", "SI"), _
            Array("Facturas", "RUC_Receptor", "RUC o identificación del receptor", "SI"), _
            Array("Facturas", "NumeroDocumento", "Número completo del comprobante", "SI"), _
            Array("Facturas", "ClaveAcceso", "Clave de acceso asignada por el SRI", "SI"), _
            Array("Facturas", "ValorTotal", "Valor total del comprobante", "SI"), _
            Array("Facturas", "EnlaceComprobante", "Hipervínculo directo al archivo XML", "SI"), _
            Array("Facturas", "Observaciones", "Notas internas de control", "NO"), _
            Array("Facturas_Detalle", "Descripcion", "Descripción de la línea del detalle", "SI"), _
            Array("Facturas_Detalle", "Cantidad", "Cantidad de la línea", "SI"), _
            Array("Facturas_Detalle", "PrecioUnitario", "Precio unitario del ítem", "SI"), _
            Array("NotasCredito", "NumeroDocumento", "Número de la nota de crédito", "SI"), _
            Array("NotasCredito", "DocumentoModifica", "Documento origen que modifica", "SI"), _
            Array("NotasCredito", "ValorTotal", "Total de la nota de crédito", "SI"), _
            Array("NotasCredito", "EnlaceComprobante", "Hipervínculo directo al XML de la nota", "SI"), _
            Array("Retenciones", "NumeroDocumento", "Número del comprobante de retención", "SI"), _
            Array("Retenciones", "PeriodoFiscal", "Periodo fiscal declarado", "SI"), _
            Array("Retenciones", "ValorRetenidoRenta", "Total de renta retenida", "SI"), _
            Array("Retenciones", "EnlaceComprobante", "Hipervínculo directo al XML de la retención", "SI"))
        AppendArrayToTable loSelector, data
    End If

    ApplySelectorCamposValidation loSelector
End Sub

Private Sub ApplySelectorCamposValidation(ByVal loSelector As ListObject)
    Dim importarColumn As ListColumn
    Dim targetRange As Range

    If loSelector Is Nothing Then Exit Sub

    On Error Resume Next
    Set importarColumn = loSelector.ListColumns("Importar")
    On Error GoTo 0

    If importarColumn Is Nothing Then Exit Sub

    Set targetRange = importarColumn.DataBodyRange

    If targetRange Is Nothing Then
        Set targetRange = importarColumn.Range.Offset(1, 0).Resize(1, 1)
    End If

    If targetRange Is Nothing Then Exit Sub

    On Error Resume Next
    targetRange.Validation.Delete
    On Error GoTo 0

    With targetRange.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="SI,NO"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Selecciona"
        .ErrorTitle = "Valor no válido"
        .InputMessage = "Indica SI o NO para importar el campo."
        .ErrorMessage = "Solo se permiten los valores SI o NO."
    End With
End Sub

Private Sub AppendArrayToTable(ByVal lo As ListObject, ByVal data As Variant)
    Dim i As Long
    Dim newRow As ListRow
    Dim rowData As Variant
    Dim columnCount As Long

    If lo Is Nothing Then Exit Sub
    If IsEmpty(data) Then Exit Sub

    columnCount = lo.ListColumns.Count

    For i = LBound(data) To UBound(data)
        Set newRow = lo.ListRows.Add
        rowData = ToRowVariant(data(i), columnCount)
        newRow.Range.Cells(1, 1).Resize(1, columnCount).Value = rowData
        UpdateUltimaActualizacion newRow
    Next i
End Sub

Private Sub UpdateUltimaActualizacion(ByVal loRow As ListRow)
    Dim ultimaActualizacionCol As Long

    ultimaActualizacionCol = GetColumnIndexByName(loRow.Parent, "UltimaActualizacion")

    If ultimaActualizacionCol = 0 Then Exit Sub

    With loRow.Range.Cells(1, ultimaActualizacionCol)
        .Value = Date
        .NumberFormat = FORMAT_DATE
    End With
End Sub

Private Function GetColumnIndexByName(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim col As ListColumn

    For Each col In lo.ListColumns
        If StrComp(col.Name, columnName, vbTextCompare) = 0 Then
            GetColumnIndexByName = col.Index
            Exit Function
        End If
    Next col

    GetColumnIndexByName = 0
End Function

Private Function ToRowVariant(ByVal source As Variant, ByVal columnCount As Long) As Variant
    Dim arr() As Variant
    Dim idx As Long
    Dim lower As Long
    Dim upper As Long

    ReDim arr(1 To 1, 1 To columnCount)

    If IsArray(source) Then
        lower = LBound(source)
        upper = UBound(source)
        For idx = 1 To columnCount
            If lower + idx - 1 <= upper Then
                arr(1, idx) = source(lower + idx - 1)
            Else
                arr(1, idx) = vbNullString
            End If
        Next idx
    Else
        arr(1, 1) = source
    End If

    ToRowVariant = arr
End Function

