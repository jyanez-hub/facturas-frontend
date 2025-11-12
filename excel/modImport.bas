' Attribute VB_Name = "modImport"
Option Explicit

'##########################################################################################
'# Módulo: modImport
'# Propósito: Importar comprobantes electrónicos del SRI desde archivos XML ubicados en
'#            una carpeta seleccionada por el usuario y poblar las tablas del libro.
'#            Incluye filtros por fecha, soporte para subcarpetas, registro de eventos y
'#            respeto a la selección de campos definida en la hoja SelectorCampos.
'# Compatibilidad: Diseñado para Excel 2007 o superior utilizando únicamente enlace tardío.
'##########################################################################################

'--- Constantes internas ---
Private Const FILE_DIALOG_FOLDER_PICKER As Long = 4
Private Const MSGBOX_TITLE As String = "Importar XML SRI"

'--- Tipos auxiliares ---
Private Type ImportOptions
    Carpeta As String
    IncluirSubcarpetas As Boolean
    FechaDesde As Variant
    FechaHasta As Variant
End Type

Private Type ImportStats
    TotalArchivos As Long
    Importados As Long
    Omitidos As Long
    Errores As Long
End Type

Private Type ImportTables
    Facturas As ListObject
    FacturasDetalle As ListObject
    NotasCredito As ListObject
    NotasCreditoDetalle As ListObject
    Retenciones As ListObject
    RetencionesDetalle As ListObject
    LogImportacion As ListObject
End Type

'##########################################################################################
'# API pública
'##########################################################################################

Public Sub Importar_XML_SRI()
    Dim opts As ImportOptions

    If Not PromptImportOptions(opts) Then Exit Sub
    ExecuteImport opts
End Sub

Public Sub Run_SmokeTest()
    Dim sampleXml As String
    Dim parsed As ParsedComprobante
    Dim tables As ImportTables
    Dim selectors As Object

    sampleXml = BuildSampleFacturaXML()

    tables = LoadImportTables(ThisWorkbook)
    If Not TablesReady(tables) Then
        Setup_Estructura targetWorkbook:=ThisWorkbook, forceReset:=False
        tables = LoadImportTables(ThisWorkbook)
    End If

    If Not TablesReady(tables) Then
        MsgBox "No fue posible preparar las tablas requeridas para la prueba.", vbCritical, MSGBOX_TITLE
        Exit Sub
    End If

    Set selectors = LoadSelectorPreferencias(ThisWorkbook)

    parsed = ParseComprobanteXML(sampleXml, "Prueba", vbNullString, "Factura_Demo.xml")
    parsed.Cabecera("OrigenArchivo") = "Prueba"
    parsed.Cabecera("RutaArchivo") = "Memoria"
    parsed.Cabecera("NombreArchivo") = "Factura_Demo.xml"
    parsed.Cabecera("FechaRegistro") = Date

    Call WriteParsedComprobante(tables, parsed, selectors)

    Debug.Print "Run_SmokeTest completado. Se insertó un comprobante de prueba en las tablas."
End Sub

'##########################################################################################
'# Flujo principal de importación
'##########################################################################################

Private Sub ExecuteImport(ByRef opts As ImportOptions)
    Dim wb As Workbook
    Dim tables As ImportTables
    Dim selectors As Object
    Dim files As Collection
    Dim stats As ImportStats
    Dim idx As Long
    Dim filePath As String
    Dim statusPrevious As Variant
    Dim screenUpdatingState As Boolean
    Dim enableEventsState As Boolean
    Dim calculationState As XlCalculation

    If Len(opts.Carpeta) = 0 Then Exit Sub

    Set wb = ThisWorkbook

    tables = LoadImportTables(wb)
    If Not TablesReady(tables) Then
        Setup_Estructura targetWorkbook:=wb, forceReset:=False
        tables = LoadImportTables(wb)
    End If

    If Not TablesReady(tables) Then
        MsgBox "Las tablas requeridas no están disponibles. Ejecuta Setup_Estructura e inténtalo nuevamente.", _
               vbCritical, MSGBOX_TITLE
        Exit Sub
    End If

    Set selectors = LoadSelectorPreferencias(wb)
    Set files = EnumerateXmlFiles(opts.Carpeta, opts.IncluirSubcarpetas)

    If files Is Nothing Then
        MsgBox "No se encontraron archivos XML en la carpeta seleccionada.", vbInformation, MSGBOX_TITLE
        Exit Sub
    End If

    If files.Count = 0 Then
        MsgBox "No se encontraron archivos XML en la carpeta seleccionada.", vbInformation, MSGBOX_TITLE
        Exit Sub
    End If

    stats.TotalArchivos = files.Count

    screenUpdatingState = Application.ScreenUpdating
    enableEventsState = Application.EnableEvents
    calculationState = Application.Calculation
    statusPrevious = Application.StatusBar

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo CleanFail

    For idx = 1 To files.Count
        filePath = CStr(files(idx))
        Application.StatusBar = "Importando XML (" & idx & "/" & files.Count & "): " & filePath
        ProcessXmlFile tables, selectors, opts, filePath, stats
    Next idx

    Application.StatusBar = statusPrevious

    ShowImportSummary stats

CleanExit:
    Application.Calculation = calculationState
    Application.EnableEvents = enableEventsState
    Application.ScreenUpdating = screenUpdatingState
    Application.StatusBar = statusPrevious
    Exit Sub

CleanFail:
    Application.StatusBar = statusPrevious
    AddLogEntry tables, "ERROR", "IMPORT", "Error inesperado durante la importación", filePath, Err.Description
    MsgBox "Ocurrió un error durante la importación: " & Err.Description, vbCritical, MSGBOX_TITLE
    Resume CleanExit
End Sub

Private Sub ProcessXmlFile(ByRef tables As ImportTables, _
                           ByVal selectors As Object, _
                           ByRef opts As ImportOptions, _
                           ByVal filePath As String, _
                           ByRef stats As ImportStats)
    Dim parsed As ParsedComprobante
    Dim nombreArchivo As String
    Dim fechaEmision As Variant
    Dim detalleMensaje As String

    On Error GoTo ParseFail

    nombreArchivo = ExtractFileNameFromPath(filePath)
    parsed = ParseComprobanteXML(filePath, "Archivo", filePath, nombreArchivo)

    If parsed.Cabecera Is Nothing Then GoTo ParseFail

    parsed.Cabecera("FechaRegistro") = Date
    parsed.Cabecera("OrigenArchivo") = "Archivo"
    parsed.Cabecera("RutaArchivo") = filePath
    parsed.Cabecera("NombreArchivo") = nombreArchivo

    fechaEmision = GetDictionaryValue(parsed.Cabecera, "FechaEmision")

    If FechaFueraDeRango(fechaEmision, opts) Then
        stats.Omitidos = stats.Omitidos + 1
        detalleMensaje = "FechaEmision=" & FormatDateDDMMYYYY(fechaEmision)
        AddLogEntry tables, "INFO", "FILTRO", "Comprobante omitido por filtro de fechas", filePath, detalleMensaje
        Exit Sub
    End If

    If WriteParsedComprobante(tables, parsed, selectors) Then
        stats.Importados = stats.Importados + 1
    Else
        stats.Omitidos = stats.Omitidos + 1
    End If
    Exit Sub

ParseFail:
    stats.Errores = stats.Errores + 1
    AddLogEntry tables, "ERROR", "PARSE", "No se pudo procesar el archivo XML", filePath, Err.Description
End Sub

Private Function WriteParsedComprobante(ByRef tables As ImportTables, _
                                        ByRef parsed As ParsedComprobante, _
                                        ByVal selectors As Object) As Boolean
    Select Case parsed.Tipo
        Case "01"
            AppendDictionaryToTable tables.Facturas, parsed.Cabecera, selectors, TABLE_FACTURAS
            AppendCollectionToTable tables.FacturasDetalle, parsed.Detalles, selectors, TABLE_FACTURAS_DETALLE
            WriteParsedComprobante = True
        Case "04"
            AppendDictionaryToTable tables.NotasCredito, parsed.Cabecera, selectors, TABLE_NOTAS_CREDITO
            AppendCollectionToTable tables.NotasCreditoDetalle, parsed.Detalles, selectors, TABLE_NOTAS_CREDITO_DETALLE
            WriteParsedComprobante = True
        Case "07"
            AppendDictionaryToTable tables.Retenciones, parsed.Cabecera, selectors, TABLE_RETENCIONES
            AppendCollectionToTable tables.RetencionesDetalle, parsed.Detalles, selectors, TABLE_RETENCIONES_DETALLE
            WriteParsedComprobante = True
        Case Else
            AddLogEntry tables, "ADVERTENCIA", "TIPO", "Tipo de comprobante no soportado", _
                        GetDictionaryValue(parsed.Cabecera, "RutaArchivo"), parsed.Tipo
            WriteParsedComprobante = False
    End Select
End Function

'##########################################################################################
'# Prompts y opciones de importación
'##########################################################################################

Private Function PromptImportOptions(ByRef opts As ImportOptions) As Boolean
    Dim carpeta As String
    Dim respuesta As VbMsgBoxResult
    Dim fechaInicio As Variant
    Dim fechaFin As Variant

    carpeta = PromptFolder("Selecciona la carpeta que contiene los XML autorizados del SRI")
    If Len(carpeta) = 0 Then Exit Function

    opts.Carpeta = carpeta

    respuesta = MsgBox("¿Deseas incluir también las subcarpetas?", vbQuestion + vbYesNo + vbDefaultButton1, MSGBOX_TITLE)
    opts.IncluirSubcarpetas = (respuesta = vbYes)

SolicitudFechas:
    If Not PromptOptionalDate("Ingresa la fecha inicial (dd/mm/aaaa) o deja en blanco:", fechaInicio) Then Exit Function
    If Not PromptOptionalDate("Ingresa la fecha final (dd/mm/aaaa) o deja en blanco:", fechaFin) Then Exit Function

    If IsDate(fechaInicio) And IsDate(fechaFin) Then
        If fechaFin < fechaInicio Then
            MsgBox "La fecha final no puede ser anterior a la inicial.", vbExclamation, MSGBOX_TITLE
            GoTo SolicitudFechas
        End If
    End If

    opts.FechaDesde = fechaInicio
    opts.FechaHasta = fechaFin

    PromptImportOptions = True
End Function

Private Function PromptFolder(ByVal mensaje As String) As String
    Dim fd As Object
    Dim selectedPath As String

    On Error Resume Next
    Set fd = Application.FileDialog(FILE_DIALOG_FOLDER_PICKER)
    On Error GoTo 0

    If Not fd Is Nothing Then
        With fd
            .Title = mensaje
            .AllowMultiSelect = False
            If .Show <> -1 Then Exit Function
            If .SelectedItems.Count = 0 Then Exit Function
            selectedPath = .SelectedItems(1)
        End With
    Else
        selectedPath = InputBox(mensaje & vbCrLf & vbCrLf & "Ingresa la ruta manualmente:", MSGBOX_TITLE)
    End If

    selectedPath = Trim$(selectedPath)
    If Len(selectedPath) = 0 Then Exit Function

    If Not FolderExists(selectedPath) Then
        MsgBox "La carpeta especificada no existe.", vbExclamation, MSGBOX_TITLE
        Exit Function
    End If

    PromptFolder = selectedPath
End Function

Private Function PromptOptionalDate(ByVal prompt As String, ByRef outputDate As Variant) As Boolean
    Dim respuesta As Variant
    Dim parsed As Variant

    Do
        respuesta = Application.InputBox(prompt, MSGBOX_TITLE, Type:=2)
        If VarType(respuesta) = vbBoolean And respuesta = False Then
            PromptOptionalDate = False
            Exit Function
        End If

        respuesta = Trim$(CStr(respuesta))

        If Len(respuesta) = 0 Then
            outputDate = Empty
            PromptOptionalDate = True
            Exit Function
        End If

        parsed = ParseDateInput(respuesta)
        If IsDate(parsed) Then
            outputDate = DateValue(parsed)
            PromptOptionalDate = True
            Exit Function
        Else
            MsgBox "La fecha ingresada no es válida. Usa el formato dd/mm/aaaa o deja en blanco.", _
                   vbExclamation, MSGBOX_TITLE
        End If
    Loop
End Function

Private Function ParseDateInput(ByVal valor As String) As Variant
    Dim parsed As Variant

    parsed = ParseSRIToDate(valor)
    If IsDate(parsed) Then
        ParseDateInput = DateValue(parsed)
        Exit Function
    End If

    On Error Resume Next
    parsed = CDate(valor)
    On Error GoTo 0

    If IsDate(parsed) Then
        ParseDateInput = DateValue(parsed)
    Else
        ParseDateInput = Empty
    End If
End Function

'##########################################################################################
'# Enumeración de archivos
'##########################################################################################

Private Function EnumerateXmlFiles(ByVal baseFolder As String, ByVal includeSubfolders As Boolean) As Collection
    Dim results As Collection

    If Len(baseFolder) = 0 Then Exit Function
    If Not FolderExists(baseFolder) Then Exit Function

    Set results = New Collection
    CollectXmlFiles NormalizeFolder(baseFolder), includeSubfolders, results
    Set EnumerateXmlFiles = results
End Function

Private Sub CollectXmlFiles(ByVal folderPath As String, ByVal includeSubfolders As Boolean, ByRef results As Collection)
    Dim fileName As String
    Dim subName As String
    Dim fullPath As String
    Dim attributes As Long

    fileName = Dir$(folderPath & "*.xml", vbNormal)
    Do While Len(fileName) > 0
        fullPath = folderPath & fileName
        results.Add fullPath
        fileName = Dir$
    Loop

    fileName = Dir$(folderPath & "*.XML", vbNormal)
    Do While Len(fileName) > 0
        fullPath = folderPath & fileName
        If Not ContainsPath(results, fullPath) Then results.Add fullPath
        fileName = Dir$
    Loop

    If Not includeSubfolders Then Exit Sub

    subName = Dir$(folderPath & "*", vbDirectory)
    Do While Len(subName) > 0
        If subName <> "." And subName <> ".." Then
            fullPath = folderPath & subName
            On Error Resume Next
            attributes = GetAttr(fullPath)
            On Error GoTo 0
            If (attributes And vbDirectory) = vbDirectory Then
                CollectXmlFiles NormalizeFolder(fullPath), True, results
            End If
        End If
        subName = Dir$
    Loop
End Sub

Private Function ContainsPath(ByVal col As Collection, ByVal path As String) As Boolean
    Dim item As Variant

    For Each item In col
        If StrComp(CStr(item), path, vbTextCompare) = 0 Then
            ContainsPath = True
            Exit Function
        End If
    Next item
End Function

Private Function NormalizeFolder(ByVal folderPath As String) As String
    Dim cleaned As String

    cleaned = Trim$(folderPath)
    If Len(cleaned) = 0 Then Exit Function

    cleaned = Replace(cleaned, "/", Application.PathSeparator)
    If Right$(cleaned, 1) <> Application.PathSeparator Then
        cleaned = cleaned & Application.PathSeparator
    End If

    NormalizeFolder = cleaned
End Function

Private Function FolderExists(ByVal folderPath As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(folderPath) And vbDirectory) = vbDirectory)
    On Error GoTo 0
End Function

Private Function ExtractFileNameFromPath(ByVal fullPath As String) As String
    Dim normalized As String
    Dim parts() As String

    normalized = Replace(fullPath, "/", Application.PathSeparator)
    If Right$(normalized, 1) = Application.PathSeparator Then
        normalized = Left$(normalized, Len(normalized) - 1)
    End If

    parts = Split(normalized, Application.PathSeparator)
    If UBound(parts) >= 0 Then
        ExtractFileNameFromPath = parts(UBound(parts))
    Else
        ExtractFileNameFromPath = normalized
    End If
End Function

'##########################################################################################
'# Interacción con tablas y registros
'##########################################################################################

Private Function LoadImportTables(ByVal wb As Workbook) As ImportTables
    Dim result As ImportTables

    On Error Resume Next
    Set result.Facturas = wb.Worksheets(SHEET_FACTURAS).ListObjects(TABLE_FACTURAS)
    Set result.FacturasDetalle = wb.Worksheets(SHEET_FACTURAS_DETALLE).ListObjects(TABLE_FACTURAS_DETALLE)
    Set result.NotasCredito = wb.Worksheets(SHEET_NOTAS_CREDITO).ListObjects(TABLE_NOTAS_CREDITO)
    Set result.NotasCreditoDetalle = wb.Worksheets(SHEET_NOTAS_CREDITO_DETALLE).ListObjects(TABLE_NOTAS_CREDITO_DETALLE)
    Set result.Retenciones = wb.Worksheets(SHEET_RETENCIONES).ListObjects(TABLE_RETENCIONES)
    Set result.RetencionesDetalle = wb.Worksheets(SHEET_RETENCIONES_DETALLE).ListObjects(TABLE_RETENCIONES_DETALLE)
    Set result.LogImportacion = wb.Worksheets(SHEET_LOG).ListObjects(TABLE_LOG)
    On Error GoTo 0

    LoadImportTables = result
End Function

Private Function TablesReady(ByRef tables As ImportTables) As Boolean
    TablesReady = Not (tables.Facturas Is Nothing Or _
                       tables.FacturasDetalle Is Nothing Or _
                       tables.NotasCredito Is Nothing Or _
                       tables.NotasCreditoDetalle Is Nothing Or _
                       tables.Retenciones Is Nothing Or _
                       tables.RetencionesDetalle Is Nothing Or _
                       tables.LogImportacion Is Nothing)
End Function

Private Function LoadSelectorPreferencias(Optional ByVal wb As Workbook) As Object
    Dim lo As ListObject
    Dim dict As Object
    Dim tableDict As Object
    Dim rowRange As Range
    Dim tablaDestino As String
    Dim campo As String
    Dim importarValor As String
    Dim importarCol As ListColumn
    Dim tablaCol As ListColumn
    Dim campoCol As ListColumn
    Dim i As Long

    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If

    On Error Resume Next
    Set lo = wb.Worksheets(SHEET_SELECTOR_CAMPOS).ListObjects(TABLE_SELECTOR_CAMPOS)
    On Error GoTo 0

    If lo Is Nothing Then Exit Function

    Set dict = NewDictionary()

    Set importarCol = GetListColumn(lo, "Importar")
    Set tablaCol = GetListColumn(lo, "TablaDestino")
    Set campoCol = GetListColumn(lo, "Campo")

    If tablaCol Is Nothing Or campoCol Is Nothing Then
        Set LoadSelectorPreferencias = dict
        Exit Function
    End If

    If lo.ListRows.Count = 0 Then
        Set LoadSelectorPreferencias = dict
        Exit Function
    End If

    For i = 1 To lo.ListRows.Count
        Set rowRange = lo.ListRows(i).Range
        tablaDestino = Trim$(CStr(rowRange.Cells(1, tablaCol.Index).Value))
        campo = Trim$(CStr(rowRange.Cells(1, campoCol.Index).Value))
        importarValor = "SI"
        If Not importarCol Is Nothing Then
            importarValor = Trim$(CStr(rowRange.Cells(1, importarCol.Index).Value))
            If Len(importarValor) = 0 Then importarValor = "SI"
        End If

        If Len(tablaDestino) > 0 And Len(campo) > 0 Then
            Set tableDict = EnsureTablePreference(dict, tablaDestino)
            tableDict(campo) = (StrComp(importarValor, "SI", vbTextCompare) = 0)
        End If
    Next i

    Set LoadSelectorPreferencias = dict
End Function

Private Function EnsureTablePreference(ByRef masterDict As Object, ByVal tableName As String) As Object
    Dim tableDict As Object

    If masterDict.Exists(tableName) Then
        Set tableDict = masterDict(tableName)
    Else
        Set tableDict = NewDictionary()
        masterDict.Add tableName, tableDict
    End If

    Set EnsureTablePreference = tableDict
End Function

Private Sub AppendDictionaryToTable(ByVal lo As ListObject, _
                                    ByVal data As Object, _
                                    ByVal selectors As Object, _
                                    ByVal tableName As String)
    Dim newRow As ListRow
    Dim idx As Long
    Dim lc As ListColumn
    Dim cell As Range
    Dim value As Variant

    If lo Is Nothing Then Exit Sub
    If data Is Nothing Then Exit Sub

    Set newRow = lo.ListRows.Add

    For idx = 1 To lo.ListColumns.Count
        Set lc = lo.ListColumns(idx)
        Set cell = newRow.Range.Cells(1, idx)
        If HasFormula(cell) Then
            ' Mantener fórmula existente
        ElseIf ShouldImportField(tableName, lc.Name, selectors) Then
            value = GetDictionaryValue(data, lc.Name)
            cell.Value = value
        End If
    Next idx
End Sub

Private Sub AppendCollectionToTable(ByVal lo As ListObject, _
                                    ByVal items As Collection, _
                                    ByVal selectors As Object, _
                                    ByVal tableName As String)
    Dim idx As Long
    Dim item As Object

    If lo Is Nothing Then Exit Sub
    If items Is Nothing Then Exit Sub

    For idx = 1 To items.Count
        Set item = items(idx)
        AppendDictionaryToTable lo, item, selectors, tableName
    Next idx
End Sub

Private Function ShouldImportField(ByVal tableName As String, _
                                   ByVal fieldName As String, _
                                   ByVal selectors As Object) As Boolean
    Dim tableDict As Object

    If selectors Is Nothing Then
        ShouldImportField = True
        Exit Function
    End If

    If Not selectors.Exists(tableName) Then
        ShouldImportField = True
        Exit Function
    End If

    Set tableDict = selectors(tableName)
    If tableDict.Exists(fieldName) Then
        ShouldImportField = CBool(tableDict(fieldName))
    Else
        ShouldImportField = True
    End If
End Function

Private Sub AddLogEntry(ByRef tables As ImportTables, _
                        ByVal nivel As String, _
                        ByVal codigo As String, _
                        ByVal descripcion As String, _
                        ByVal rutaArchivo As String, _
                        Optional ByVal detalle As String = "")
    Dim lo As ListObject
    Dim newRow As ListRow

    Set lo = tables.LogImportacion
    If lo Is Nothing Then Exit Sub

    Set newRow = lo.ListRows.Add

    SetRowValue newRow, "MarcaTiempo", Date
    SetRowValue newRow, "Hora", Time
    SetRowValue newRow, "Nivel", nivel
    SetRowValue newRow, "Codigo", codigo
    SetRowValue newRow, "Descripcion", descripcion
    SetRowValue newRow, "RutaArchivo", rutaArchivo
    SetRowValue newRow, "Detalle", detalle
End Sub

Private Sub SetRowValue(ByVal loRow As ListRow, ByVal columnName As String, ByVal value As Variant)
    Dim idx As Long

    idx = FindColumnIndex(loRow.Parent, columnName)
    If idx = 0 Then Exit Sub

    With loRow.Range.Cells(1, idx)
        If HasFormula(loRow.Range.Cells(1, idx)) Then
            ' mantener fórmula
        Else
            .Value = value
        End If
    End With
End Sub

Private Function FindColumnIndex(ByVal lo As ListObject, ByVal columnName As String) As Long
    Dim lc As ListColumn

    For Each lc In lo.ListColumns
        If StrComp(lc.Name, columnName, vbTextCompare) = 0 Then
            FindColumnIndex = lc.Index
            Exit Function
        End If
    Next lc
End Function

Private Function GetListColumn(ByVal lo As ListObject, ByVal name As String) As ListColumn
    Dim lc As ListColumn

    For Each lc In lo.ListColumns
        If StrComp(lc.Name, name, vbTextCompare) = 0 Then
            Set GetListColumn = lc
            Exit Function
        End If
    Next lc
End Function

Private Function HasFormula(ByVal target As Range) As Boolean
    On Error Resume Next
    HasFormula = target.HasFormula
    On Error GoTo 0
End Function

Private Function GetDictionaryValue(ByVal dict As Object, ByVal key As String) As Variant
    If dict Is Nothing Then Exit Function
    On Error Resume Next
    If dict.Exists(key) Then
        GetDictionaryValue = dict(key)
    Else
        GetDictionaryValue = Empty
    End If
    On Error GoTo 0
End Function

Private Function NewDictionary() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    dict.CompareMode = 1 ' vbTextCompare
    On Error GoTo 0
    Set NewDictionary = dict
End Function

Private Function FechaFueraDeRango(ByVal fechaValor As Variant, ByRef opts As ImportOptions) As Boolean
    If Not IsDate(opts.FechaDesde) And Not IsDate(opts.FechaHasta) Then Exit Function
    If Not IsDate(fechaValor) Then Exit Function

    If IsDate(opts.FechaDesde) Then
        If fechaValor < opts.FechaDesde Then
            FechaFueraDeRango = True
            Exit Function
        End If
    End If

    If IsDate(opts.FechaHasta) Then
        If fechaValor > opts.FechaHasta Then
            FechaFueraDeRango = True
        End If
    End If
End Function

Private Sub ShowImportSummary(ByRef stats As ImportStats)
    Dim message As String

    message = "Importación completada:" & vbCrLf & vbCrLf & _
              "Archivos encontrados: " & stats.TotalArchivos & vbCrLf & _
              "Importados: " & stats.Importados & vbCrLf & _
              "Omitidos por filtro: " & stats.Omitidos & vbCrLf & _
              "Errores: " & stats.Errores

    MsgBox message, vbInformation, MSGBOX_TITLE
End Sub

'##########################################################################################
'# Utilidades internas
'##########################################################################################

Private Function BuildSampleFacturaXML() As String
    Dim xml As String

    xml = "<?xml version=""1.0"" encoding=""UTF-8""?>" & vbLf
    xml = xml & "<factura id=""comprobante"" version=""1.1.0"">" & vbLf
    xml = xml & "<infoTributaria>" & vbLf
    xml = xml & "<razonSocial>Empresa Demo S.A.</razonSocial>" & vbLf
    xml = xml & "<nombreComercial>Empresa Demo</nombreComercial>" & vbLf
    xml = xml & "<ruc>1790012345001</ruc>" & vbLf
    xml = xml & "<claveAcceso>0123456789012345678901234567890123456789012345678</claveAcceso>" & vbLf
    xml = xml & "<codDoc>01</codDoc>" & vbLf
    xml = xml & "<estab>001</estab>" & vbLf
    xml = xml & "<ptoEmi>002</ptoEmi>" & vbLf
    xml = xml & "<secuencial>000000123</secuencial>" & vbLf
    xml = xml & "</infoTributaria>" & vbLf
    xml = xml & "<infoFactura>" & vbLf
    xml = xml & "<fechaEmision>2024-01-15</fechaEmision>" & vbLf
    xml = xml & "<tipoIdentificacionComprador>05</tipoIdentificacionComprador>" & vbLf
    xml = xml & "<razonSocialComprador>Cliente de Prueba</razonSocialComprador>" & vbLf
    xml = xml & "<identificacionComprador>0912345678</identificacionComprador>" & vbLf
    xml = xml & "<moneda>USD</moneda>" & vbLf
    xml = xml & "<totalSinImpuestos>100.00</totalSinImpuestos>" & vbLf
    xml = xml & "<totalDescuento>0.00</totalDescuento>" & vbLf
    xml = xml & "<totalConImpuestos>" & vbLf
    xml = xml & "<totalImpuesto>" & vbLf
    xml = xml & "<codigo>2</codigo>" & vbLf
    xml = xml & "<codigoPorcentaje>2</codigoPorcentaje>" & vbLf
    xml = xml & "<baseImponible>100.00</baseImponible>" & vbLf
    xml = xml & "<tarifa>12</tarifa>" & vbLf
    xml = xml & "<valor>12.00</valor>" & vbLf
    xml = xml & "</totalImpuesto>" & vbLf
    xml = xml & "</totalConImpuestos>" & vbLf
    xml = xml & "<propina>0.00</propina>" & vbLf
    xml = xml & "<importeTotal>112.00</importeTotal>" & vbLf
    xml = xml & "</infoFactura>" & vbLf
    xml = xml & "<detalles>" & vbLf
    xml = xml & "<detalle>" & vbLf
    xml = xml & "<codigoPrincipal>PROD001</codigoPrincipal>" & vbLf
    xml = xml & "<descripcion>Producto de prueba</descripcion>" & vbLf
    xml = xml & "<cantidad>1</cantidad>" & vbLf
    xml = xml & "<precioUnitario>100.00</precioUnitario>" & vbLf
    xml = xml & "<descuento>0.00</descuento>" & vbLf
    xml = xml & "<precioTotalSinImpuesto>100.00</precioTotalSinImpuesto>" & vbLf
    xml = xml & "<impuestos>" & vbLf
    xml = xml & "<impuesto>" & vbLf
    xml = xml & "<codigo>2</codigo>" & vbLf
    xml = xml & "<codigoPorcentaje>2</codigoPorcentaje>" & vbLf
    xml = xml & "<tarifa>12</tarifa>" & vbLf
    xml = xml & "<baseImponible>100.00</baseImponible>" & vbLf
    xml = xml & "<valor>12.00</valor>" & vbLf
    xml = xml & "</impuesto>" & vbLf
    xml = xml & "</impuestos>" & vbLf
    xml = xml & "</detalle>" & vbLf
    xml = xml & "</detalles>" & vbLf
    xml = xml & "<infoAdicional>" & vbLf
    xml = xml & "<campoAdicional nombre=""Observacion"">Factura de prueba</campoAdicional>" & vbLf
    xml = xml & "</infoAdicional>" & vbLf
    xml = xml & "</factura>"

    BuildSampleFacturaXML = xml
End Function

