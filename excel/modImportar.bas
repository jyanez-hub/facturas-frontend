Attribute VB_Name = "modImportar"

Option Explicit
'Public gRutaPDF As String  ' <- NUEVO: ruta base de los PDF (con \ al final)
'Public gRutaPDF As String
' ============================
'   IMPORTADOR XML SRI ? Excel (v5)
'   - Factura, Nota de Crédito (en Facturas/Detalle)
'   - Nota de Débito (en Facturas; sin detalle)
'   - Retenciones (Retenciones/RetDet)
'   - Filtros: fecha_emision (desde/hasta) e incluir subcarpetas
'   - Desencapsula <autorizacion>/<comprobante><![CDATA[...]]>
'   - Tolerante a namespaces con local-name()
'   - Columnas TEXTO y conversión numérica locale-aware
' ============================
Dim wbNuevo As Workbook

' ===== Cabeceras =====
Public FACTURAS_HEADERS As Variant
Public DETALLE_HEADERS As Variant
Public RET_HEADERS As Variant
Public RETDET_HEADERS As Variant


' ===== Columnas TEXTO =====
Private TEXT_COLS_FACTURAS As Variant
Private TEXT_COLS_DETALLE As Variant
Private TEXT_COLS_RET As Variant
Private TEXT_COLS_RETDET As Variant

' ======= Overrides desde UI (formulario) =======
' ===== Overrides que llegan del formulario =====
Public gUI_HasParams As Boolean
Public gUI_FolderPath As String
Public gUI_IncludeSubfolders As Boolean
Public gUI_FDesde As String
Public gUI_FHasta As String

' Filtro de tipo de documento elegido en el form
Public Enum EX_FiltroDoc
    exAll = 0
    exDocs = 1      ' Facturas/NC/ND/Liq
    exRet = 2       ' Retenciones
End Enum
Public gUI_FilterDocs As EX_FiltroDoc

' Ruta base para PDF/links (no obligatorio, pero evita “sub o function no definida”)
Public gRutaPDF As String
'Public Sub SetRutaPDF(ByVal p As String): gRutaPDF = p: End Sub


'***********************************************************************************

' ====== Destino de escritura ======
Public gTargetWb As Workbook
Private Function TWB() As Workbook
    If Not gTargetWb Is Nothing Then
        Set TWB = gTargetWb
    Else
        Set TWB = ThisWorkbook
    End If
End Function

'********************************************************************************************

Public Sub InitHeaders()

    ' === FACTURAS ===
    FACTURAS_HEADERS = Array( _
        "clave_acceso", "tipo_comprobante", "ambiente", "emision", _
        "ruc_emisor", "razon_social_emisor", "nombre_comercial", "dir_matriz", _
        "establecimiento", "punto_emision", "secuencial", "nro_comprobante", "fecha_emision", _
        "dir_establecimiento", "ruc_ci_comprador", "razon_social_comprador", "moneda", _
        "total_sin_impuestos", "subtotal_iva_0", "subtotal_iva_5", "subtotal_iva_12", "subtotal_iva_15", _
        "subtotal_no_objeto", "subtotal_exento", "descuento", "ice", "iva_total", _
        "propina", "valor_total", "forma_pago_1", "plazo_1", "tiempo_1", "guia_remision", _
        "doc_sust_tipo", "doc_sust_serie", "doc_sust_secuencial", "doc_sust_fecha" _
    )

    ' === DETALLE ===
    DETALLE_HEADERS = Array( _
        "clave_acceso", "codigo_principal", "codigo_auxiliar", "descripcion", "cantidad", _
        "precio_unitario", "descuento", "precio_total_sin_impuesto", _
        "impuesto_codigo", "impuesto_porcentaje_codigo", "tarifa_iva", _
        "base_imponible", "valor_iva", "ice_codigo", "ice_tarifa", "ice_valor", "irbpnr_valor" _
    )

    ' === RETENCIONES (ENCABEZADO) ===
    ' Incluye: nro_comprobante, % de retención y documento sustentado
    RET_HEADERS = Array( _
        "clave_acceso", "tipo_comprobante", "ambiente", _
        "ruc_emisor", "razon_social_emisor", "establecimiento", "punto_emision", "secuencial", "nro_comprobante", _
        "fecha_emision", "periodo_fiscal", "ruc_sujeto", "razon_social_sujeto", _
        "cod_ret_iva", "porc_ret_iva", "base_ret_iva", "valor_ret_iva", _
        "cod_ret_renta", "porc_ret_renta", "base_ret_renta", "valor_ret_renta", _
        "total_retenido", _
        "doc_sust_tipo", "doc_sust_serie", "doc_sust_secuencial", "doc_sust_fecha" _
    )

    ' === RETENCIONES (DETALLE) ===
    RETDET_HEADERS = Array( _
        "clave_acceso", "impuesto", "codigo", "codigo_retencion", _
        "base_imponible", "porcentaje_retener", "valor_retenido" _
    )
End Sub


Public Sub InitTextCols()

    TEXT_COLS_FACTURAS = Array( _
        "clave_acceso", "tipo_comprobante", "ambiente", "emision", _
        "ruc_emisor", "establecimiento", "punto_emision", "secuencial", "nro_comprobante", _
        "ruc_ci_comprador", "moneda", "forma_pago_1", "plazo_1", "tiempo_1", "guia_remision", _
        "doc_sust_tipo", "doc_sust_serie", "doc_sust_secuencial", "doc_sust_fecha" _
    )

    TEXT_COLS_DETALLE = Array( _
        "clave_acceso", "codigo_principal", "codigo_auxiliar", _
        "impuesto_codigo", "impuesto_porcentaje_codigo", "ice_codigo" _
    )

    TEXT_COLS_RET = Array( _
        "clave_acceso", "tipo_comprobante", "ambiente", _
        "ruc_emisor", "establecimiento", "punto_emision", "secuencial", "nro_comprobante", _
        "ruc_sujeto", "periodo_fiscal", _
        "doc_sust_tipo", "doc_sust_serie", "doc_sust_secuencial", "doc_sust_fecha" _
    )

    TEXT_COLS_RETDET = Array( _
        "clave_acceso", "impuesto", "codigo", "codigo_retencion" _
    )
End Sub


' ===== Macro principal (con filtros) =====
Public Sub Importar_XML_SRI()
    On Error GoTo EH

    InitHeaders
    InitTextCols

    ' -------- Filtros (UI-aware) --------
    Dim folderPath As String, includeSubfolders As Boolean
    Dim fDesde As String, fHasta As String

    If gUI_HasParams Then
        folderPath = Trim$(gUI_FolderPath)
        includeSubfolders = gUI_IncludeSubfolders
        fDesde = Trim$(gUI_FDesde)
        fHasta = Trim$(gUI_FHasta)
    Else
        folderPath = PickFolder("Elige la carpeta que contiene los XML del SRI")
        If Len(folderPath) = 0 Then Exit Sub
        includeSubfolders = (MsgBox("¿Incluir subcarpetas?", vbQuestion + vbYesNo, "Opciones de importación") = vbYes)
        fDesde = InputBox("Fecha desde (YYYY-MM-DD) o vacío:", "Filtro de fecha (opcional)")
        fHasta = InputBox("Fecha hasta (YYYY-MM-DD) o vacío:", "Filtro de fecha (opcional)")
    End If

    ' Validar existencia de carpeta (evita error 76)
    If Len(folderPath) = 0 Or dir$(folderPath, vbDirectory) = vbNullString Then
        MsgBox "La carpeta especificada no existe:" & vbCrLf & folderPath, vbExclamation
        Exit Sub
    End If

    ' Parse de fechas (admite DD-MM-YYYY o YYYY-MM-DD)
    Dim dDesde As Date, dHasta As Date, useDesde As Boolean, useHasta As Boolean
    If Len(fDesde) > 0 Then If TryParseISODate(fDesde, dDesde) Or TryParseFechaEmision(fDesde, dDesde) Then useDesde = True
    If Len(fHasta) > 0 Then If TryParseISODate(fHasta, dHasta) Or TryParseFechaEmision(fHasta, dHasta) Then useHasta = True

    ' Guardar ruta base (para PDF/links si aplica)
    On Error Resume Next
    EX_SetRutaPDF folderPath
    On Error GoTo 0

    ' -------- Preparación hojas --------
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Dim wsF As Worksheet, wsD As Worksheet, wsL As Worksheet, wsR As Worksheet, wsRD As Worksheet
    Set wsF = EnsureSheet("Facturas")
    Set wsD = EnsureSheet("Detalle")
    Set wsR = EnsureSheet("Retenciones")
    Set wsRD = EnsureSheet("RetDet")
    Set wsL = EnsureSheet("LOG")

    EnsureHeadersIfEmpty wsF, FACTURAS_HEADERS
    EnsureHeadersIfEmpty wsD, DETALLE_HEADERS
    EnsureHeadersIfEmpty wsR, RET_HEADERS
    EnsureHeadersIfEmpty wsRD, RETDET_HEADERS
    InitLogSheet wsL
    Call EnsureRetHeaders

    ' -------- Enumerar archivos --------
    Dim files As Collection: Set files = New Collection
    If includeSubfolders Then
        EnumXmlRecursive folderPath, files
    Else
        EnumXmlFlat folderPath, files
    End If

    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    Dim totalFiles As Long, okFiles As Long
    Dim i As Long
    For i = 1 To files.Count
        Dim fullPath As String: fullPath = files(i)
        If Not seen.Exists(UCase$(fullPath)) Then
            seen.Add UCase$(fullPath), True
            totalFiles = totalFiles + 1
            If ParseOneXML(fullPath, wsF, wsD, wsR, wsRD, wsL, useDesde, dDesde, useHasta, dHasta) Then okFiles = okFiles + 1
        End If
    Next i

    
    ' ===== Asegurar/actualizar TABLAS y aplicar ESTILO CORPORATIVO =====
    ' 1) Aseguramos que cada hoja tenga una ListObject que cubra el rango con datos
    Dim loF As ListObject, loD As ListObject, loR As ListObject, loRD As ListObject
    Set loF = EnsureAsTable(wsF, "tbFacturas")
    Set loD = EnsureAsTable(wsD, "tbDetalle")
    Set loR = EnsureAsTable(wsR, "tbRetenciones")
    Set loRD = EnsureAsTable(wsRD, "tbRetDet")

    ' 2) Aplicar formato azul corporativo (#1C0F82) a cada tabla creada
    If Not loF Is Nothing Then AplicarFormatoTabla_EXCELBOT loF
    If Not loD Is Nothing Then AplicarFormatoTabla_EXCELBOT loD
    If Not loR Is Nothing Then AplicarFormatoTabla_EXCELBOT loR
    If Not loRD Is Nothing Then AplicarFormatoTabla_EXCELBOT loRD

    ' (Opcional) Si tu flujo usa un libro NUEVO en variable wbNuevo, puedes formatear todo libro:
    ' On Error Resume Next
    ' AplicarFormatoATodasLasTablas wbNuevo
    ' On Error GoTo 0
    
    ' -------- AUTOFORMATO AUTOMÁTICO --------
    ' silencioso para no mostrar el segundo mensaje
    FormatearColumnas True
    Call EX_FormatTablesAndCurrency(wbNuevo)
  

    ' -------- Hipervínculos a XML/PDF --------
    Call AutoHyperlinks_AfterImport

    ' -------- Mensaje final --------
    MsgBox "Listo." & vbCrLf & _
           "XML encontrados: " & totalFiles & vbCrLf & _
           "Importados OK: " & okFiles & vbCrLf & _
           "Revisa la hoja LOG para detalles.", vbInformation

SafeExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub

EH:
    MsgBox "Error: " & Err.Description, vbExclamation
    Resume SafeExit
End Sub


' ==== Helper local: asegura que la hoja tenga una tabla y la ajusta al rango usado ====
Private Function EnsureAsTable(ByVal ws As Worksheet, ByVal tableName As String) As ListObject
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(tableName)
    On Error GoTo 0

    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastRow < 1 Or lastCol < 1 Then Exit Function

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))

    If lo Is Nothing Then
        Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
        On Error Resume Next
        lo.name = tableName
        On Error GoTo 0
    Else
        lo.Resize rng
    End If

    Set EnsureAsTable = lo
End Function


Private Sub EnsureRetHeaders()
    ' Alinea Retenciones/RetDet con los arrays globales actuales
    InitHeaders

    Dim wsR As Worksheet, wsRD As Worksheet
    On Error Resume Next
    Set wsR = ThisWorkbook.Worksheets("Retenciones")
    Set wsRD = ThisWorkbook.Worksheets("RetDet")
    On Error GoTo 0
    If wsR Is Nothing Or wsRD Is Nothing Then Exit Sub

    WriteHeadersExact wsR, RET_HEADERS
    WriteHeadersExact wsRD, RETDET_HEADERS
End Sub


' ===== Enumeración de archivos =====
Private Sub EnumXmlFlat(ByVal folderPath As String, ByRef files As Collection)
    Dim f As String
    f = dir(folderPath & "\*.xml")
    Do While Len(f) > 0
        files.Add folderPath & "\" & f
        f = dir
    Loop
    f = dir(folderPath & "\*.XML")
    Do While Len(f) > 0
        files.Add folderPath & "\" & f
        f = dir
    Loop
End Sub

Private Sub EnumXmlRecursive(ByVal rootPath As String, ByRef files As Collection)
    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim stack As Collection: Set stack = New Collection
    stack.Add rootPath
    Do While stack.Count > 0
        Dim p As String: p = stack(1): stack.Remove 1
        EnumXmlFlat p, files
        Dim fld As Object
        For Each fld In fso.GetFolder(p).SubFolders
            stack.Add fld.Path
        Next fld
    Loop
End Sub

' ===== Procesa 1 XML (versión robusta) =====
Private Function ParseOneXML(ByVal xmlPath As String, _
                             ByRef wsF As Worksheet, ByRef wsD As Worksheet, _
                             ByRef wsR As Worksheet, ByRef wsRD As Worksheet, _
                             ByRef wsL As Worksheet, _
                             ByVal useDesde As Boolean, ByVal dDesde As Date, _
                             ByVal useHasta As Boolean, ByVal dHasta As Date) As Boolean
    On Error GoTo EH

    Dim dom As Object: Set dom = CreateObject("MSXML2.DOMDocument.6.0")
    dom.async = False: dom.validateOnParse = False
    dom.SetProperty "SelectionLanguage", "XPath"

    ' 1) Cargar XML externo
    If Not dom.Load(xmlPath) Then
        Log wsL, xmlPath, "ERROR", "No se pudo cargar el XML externo."
        ParseOneXML = False: Exit Function
    End If
    If dom Is Nothing Or dom.DocumentElement Is Nothing Then
        Log wsL, xmlPath, "ERROR", "XML sin elemento raíz legible."
        ParseOneXML = False: Exit Function
    End If
    EnsureXPath dom.DocumentElement

    ' 2) Si la raíz no es factura/notaCredito/notaDebito/retención, abrir <comprobante><![CDATA[...]]>
    Dim rootName As String: rootName = LCase$(dom.DocumentElement.nodeName)
    If (InStr(rootName, "factura") = 0 And InStr(rootName, "notacredito") = 0 _
        And InStr(rootName, "notadebito") = 0 And InStr(rootName, "comprobanteretencion") = 0) Then

        Dim innerDom As Object: Set innerDom = ExtractInnerComprobante(dom)
        If Not innerDom Is Nothing Then
            Set dom = innerDom
            If dom Is Nothing Or dom.DocumentElement Is Nothing Then
                Log wsL, xmlPath, "ERROR", "Comprobante interno vacío."
                ParseOneXML = False: Exit Function
            End If
            EnsureXPath dom.DocumentElement
        Else
            ' Intento alterno: envolturas de autorización
            If Not TryLoadFromAutorizacion(xmlPath, dom) Then
                Dim rn As String
                rn = IIf(dom Is Nothing Or dom.DocumentElement Is Nothing, "(sin raíz)", LCase$(dom.DocumentElement.nodeName))
                Log wsL, xmlPath, "ERROR", "No se encontró <comprobante>. Raíz externa=" & rn
                ParseOneXML = False: Exit Function
            End If
            If dom Is Nothing Or dom.DocumentElement Is Nothing Then
                Log wsL, xmlPath, "ERROR", "Autorización sin XML interno legible."
                ParseOneXML = False: Exit Function
            End If
            EnsureXPath dom.DocumentElement
        End If
    End If

    ' 3) Identificar tipo de documento y nodos clave
    Dim docType As String
    Dim rootNode As Object, infoTrib As Object, infoNode As Object, detalles As Object
    If Not FindDocRoot(dom, docType, rootNode, infoTrib, infoNode, detalles) Then
        Dim rn2 As String
        rn2 = IIf(dom Is Nothing Or dom.DocumentElement Is Nothing, "(sin raíz)", LCase$(dom.DocumentElement.nodeName))
        Log wsL, xmlPath, "ERROR", "No se identificó tipo de comprobante. Raíz interna=" & rn2
        ParseOneXML = False: Exit Function
    End If
    EnsureXPath rootNode: EnsureXPath infoTrib: EnsureXPath infoNode: EnsureXPath detalles
    
    
    
    ' 3) Identificar tipo de documento y nodos clave
' ... (tu código actual que setea docType)

' --- FILTRO por tipo elegido en el formulario ---
If gUI_FilterDocs = exDocs And LCase$(docType) = "retencion" Then
    Log wsL, xmlPath, "SKIP", "Filtro=Documentos, se omitió Retención"
    ParseOneXML = False: Exit Function
ElseIf gUI_FilterDocs = exRet And LCase$(docType) <> "retencion" Then
    Log wsL, xmlPath, "SKIP", "Filtro=Retenciones, se omitió " & docType
    ParseOneXML = False: Exit Function
End If




' --- FILTRO por tipo elegido en el formulario ---
Select Case gUI_FilterDocs
    Case exDocs
        If LCase$(docType) = "retencion" Then
            Log wsL, xmlPath, "SKIP", "Filtro=Documentos (omitida Retención)"
            ParseOneXML = False: Exit Function
        End If
    Case exRet
        If LCase$(docType) <> "retencion" Then
            Log wsL, xmlPath, "SKIP", "Filtro=Retenciones (omitido " & docType & ")"
            ParseOneXML = False: Exit Function
        End If
End Select









    ' 4) Filtro por fecha de emisión (si se pidió)
    Dim fechaTxt As String: fechaTxt = SXL(infoNode, "./*[local-name()='fechaEmision']")
    Dim fechaEmi As Date, pasaFiltro As Boolean: pasaFiltro = True
    If TryParseFechaEmision(fechaTxt, fechaEmi) Then
        If useDesde And fechaEmi < dDesde Then pasaFiltro = False
        If useHasta And fechaEmi > dHasta Then pasaFiltro = False
    End If
    If Not pasaFiltro Then
        Log wsL, xmlPath, "SKIP", "Fuera de rango de fecha (" & fechaTxt & ")"
        ParseOneXML = False: Exit Function
    End If

    ' 5) Enrutar a parser específico
    Select Case docType
        Case "factura", "notaCredito", "notaDebito"
            ParseFacturaLike xmlPath, docType, rootNode, infoTrib, infoNode, detalles, wsF, wsD, wsL
        Case "retencion"
            ParseRetencion xmlPath, rootNode, infoTrib, infoNode, wsR, wsRD, wsL
        Case Else
            Log wsL, xmlPath, "ERROR", "Tipo de documento no soportado: " & docType
            ParseOneXML = False: Exit Function
    End Select

    ParseOneXML = True
    Exit Function

EH:
    Log wsL, xmlPath, "ERROR", Err.Description
    ParseOneXML = False
End Function

' ===== Procesa factura / notaCredito / notaDebito =====
Private Sub ParseFacturaLike(ByVal xmlPath As String, ByVal docType As String, _
                             ByVal rootNode As Object, ByVal infoTrib As Object, ByVal infoNode As Object, ByVal detalles As Object, _
                             ByRef wsF As Worksheet, ByRef wsD As Worksheet, ByRef wsL As Worksheet)

    Dim enc As Object: Set enc = CreateObject("Scripting.Dictionary")
    enc("clave_acceso") = SXL(rootNode, "./*[local-name()='infoTributaria']/*[local-name()='claveAcceso']")
    If Len(enc("clave_acceso")) = 0 Then enc("clave_acceso") = FileNameNoExt(xmlPath)

    enc("tipo_comprobante") = SXL(infoTrib, "./*[local-name()='codDoc']")
    enc("ambiente") = SXL(infoTrib, "./*[local-name()='ambiente']")
    enc("emision") = SXL(infoTrib, "./*[local-name()='tipoEmision']")
    enc("ruc_emisor") = SXL(infoTrib, "./*[local-name()='ruc']")
    enc("razon_social_emisor") = SXL(infoTrib, "./*[local-name()='razonSocial']")
    enc("nombre_comercial") = SXL(infoTrib, "./*[local-name()='nombreComercial']")
    enc("dir_matriz") = SXL(infoTrib, "./*[local-name()='dirMatriz']")
    enc("establecimiento") = SXL(infoTrib, "./*[local-name()='estab']")
    enc("punto_emision") = SXL(infoTrib, "./*[local-name()='ptoEmi']")
    enc("secuencial") = SXL(infoTrib, "./*[local-name()='secuencial']")
    
'******************************
' nro_comprobante (Facturas/NC)
enc("nro_comprobante") = FormatComprobante(enc("establecimiento"), enc("punto_emision"), enc("secuencial"))

' Si es NOTA DE CRÉDITO, completa documento sustentado
If LCase$(docType) = "notacredito" Then FillDocSust_NC infoNode, enc


'******************************

    enc("fecha_emision") = SXL(infoNode, "./*[local-name()='fechaEmision']")
    enc("dir_establecimiento") = SXL(infoNode, "./*[local-name()='dirEstablecimiento']")
    enc("ruc_ci_comprador") = SXL(infoNode, "./*[local-name()='identificacionComprador']")
    enc("razon_social_comprador") = SXL(infoNode, "./*[local-name()='razonSocialComprador']")
    enc("moneda") = SXL(infoNode, "./*[local-name()='moneda']")

    enc("total_sin_impuestos") = n(SXL(infoNode, "./*[local-name()='totalSinImpuestos']"))
    enc("descuento") = n(SXL(infoNode, "./*[local-name()='totalDescuento']"))
    enc("propina") = n(SXL(infoNode, "./*[local-name()='propina']"))

    Dim vt As Double
    vt = n(SXL(infoNode, "./*[local-name()='importeTotal']"))
    If vt = 0# Then vt = n(SXL(infoNode, "./*[local-name()='valorModificacion']")) ' notaCredito/notaDebito
    enc("valor_total") = vt

    enc("guia_remision") = SXL(infoNode, "./*[local-name()='guiaRemision']")

    ' pagos (si existen)
    Dim pagos As Object, pago As Object
    Set pagos = infoNode.SelectSingleNode("./*[local-name()='pagos']")
    If Not pagos Is Nothing Then
        Set pago = pagos.SelectSingleNode("./*[local-name()='pago']")
        If Not pago Is Nothing Then
            enc("forma_pago_1") = SXL(pago, "./*[local-name()='formaPago']")
            enc("plazo_1") = SXL(pago, "./*[local-name()='plazo']")
            enc("tiempo_1") = SXL(pago, "./*[local-name()='unidadTiempo']")
        End If
    End If
    If Not enc.Exists("forma_pago_1") Then enc("forma_pago_1") = ""
    If Not enc.Exists("plazo_1") Then enc("plazo_1") = ""
    If Not enc.Exists("tiempo_1") Then enc("tiempo_1") = ""

    ' Totales por impuestos
    Dim tot As Object: Set tot = SumarImpuestosLocalName(infoNode)
    enc("subtotal_iva_0") = tot("iva0")
    enc("subtotal_iva_5") = tot("iva5")
    enc("subtotal_iva_12") = tot("iva12")
    enc("subtotal_iva_15") = tot("iva15")
    enc("subtotal_no_objeto") = tot("noObjeto")
    enc("subtotal_exento") = tot("exento")
    enc("iva_total") = tot("ivaTotal")
    enc("ice") = tot("iceTotal")




' **********************************




FixSubtotalIVA0_FromHeader enc, "FAC " & CStr(enc("secuencial"))








' *************************************




    ' Escribir encabezado
    Dim rF As Long: rF = nextRow(wsF)
    WriteDictToRow wsF, rF, FACTURAS_HEADERS, enc

    ' DETALLE:
    ' - Factura y Nota de Crédito: tienen <detalles>
    ' - Nota de Débito: no trae ítems; dejamos solo encabezado
    If Not detalles Is Nothing And (docType = "factura" Or docType = "notaCredito") Then
        Dim det As Object, ii As Object, imps As Object
        For Each det In detalles.SelectNodes("./*[local-name()='detalle']")
            Dim rowBase As Object: Set rowBase = CreateObject("Scripting.Dictionary")
            rowBase("clave_acceso") = enc("clave_acceso")

            rowBase("codigo_principal") = SXL(det, "./*[local-name()='codigoPrincipal']")
            If Len(rowBase("codigo_principal")) = 0 Then rowBase("codigo_principal") = SXL(det, "./*[local-name()='codigoInterno']")

            rowBase("codigo_auxiliar") = SXL(det, "./*[local-name()='codigoAuxiliar']")
            If Len(rowBase("codigo_auxiliar")) = 0 Then rowBase("codigo_auxiliar") = SXL(det, "./*[local-name()='codigoAdicional']")

            Dim descTxt As String
            descTxt = SXL(det, "./*[local-name()='descripcion']")
            If Len(descTxt) = 0 Then _
                descTxt = SXL(det, "./*[local-name()='detallesAdicionales']/*[local-name()='detAdicional']/@valor")
            rowBase("descripcion") = descTxt

            rowBase("cantidad") = n(SXL(det, "./*[local-name()='cantidad']"))
            rowBase("precio_unitario") = n(SXL(det, "./*[local-name()='precioUnitario']"))
            rowBase("descuento") = n(SXL(det, "./*[local-name()='descuento']"))
            rowBase("precio_total_sin_impuesto") = n(SXL(det, "./*[local-name()='precioTotalSinImpuesto']"))

            Set imps = det.SelectSingleNode("./*[local-name()='impuestos']")
            If Not imps Is Nothing Then
                For Each ii In imps.SelectNodes("./*[local-name()='impuesto']")
                    Dim rowD As Object: Set rowD = CreateObject("Scripting.Dictionary")
                    CopyDict rowBase, rowD
                    rowD("impuesto_codigo") = SXL(ii, "./*[local-name()='codigo']")
                    rowD("impuesto_porcentaje_codigo") = SXL(ii, "./*[local-name()='codigoPorcentaje']")
                    rowD("tarifa_iva") = n(SXL(ii, "./*[local-name()='tarifa']"))
                    rowD("base_imponible") = n(SXL(ii, "./*[local-name()='baseImponible']"))
                    rowD("valor_iva") = n(SXL(ii, "./*[local-name()='valor']"))
                    rowD("ice_codigo") = "": rowD("ice_tarifa") = 0#: rowD("ice_valor") = 0#: rowD("irbpnr_valor") = 0#
                    If rowD("impuesto_codigo") = "3" Then
                        rowD("ice_codigo") = rowD("impuesto_porcentaje_codigo")
                        rowD("ice_tarifa") = rowD("tarifa_iva")
                        rowD("ice_valor") = rowD("valor_iva")
                    ElseIf rowD("impuesto_codigo") = "5" Then
                        rowD("irbpnr_valor") = rowD("valor_iva")
                    End If
                    WriteDictToRow wsD, nextRow(wsD), DETALLE_HEADERS, rowD
                Next ii
            Else
                Dim rowD2 As Object: Set rowD2 = CreateObject("Scripting.Dictionary")
                CopyDict rowBase, rowD2
                rowD2("impuesto_codigo") = "": rowD2("impuesto_porcentaje_codigo") = ""
                rowD2("tarifa_iva") = 0#: rowD2("base_imponible") = 0#: rowD2("valor_iva") = 0#
                rowD2("ice_codigo") = "": rowD2("ice_tarifa") = 0#: rowD2("ice_valor") = 0#: rowD2("irbpnr_valor") = 0#
                WriteDictToRow wsD, nextRow(wsD), DETALLE_HEADERS, rowD2
            End If
        Next det
    End If

    Log wsL, xmlPath, "OK", "Importado (" & docType & ")"
End Sub

' ===== Procesa Retención =====
' ===== Procesa Retención (COMPLETO) =====
Private Sub ParseRetencion(ByVal xmlPath As String, _
                           ByVal rootNode As Object, ByVal infoTrib As Object, ByVal infoRet As Object, _
                           ByRef wsR As Worksheet, ByRef wsRD As Worksheet, ByRef wsL As Worksheet)

    On Error GoTo EH

    Dim enc As Object: Set enc = CreateObject("Scripting.Dictionary")

    ' --- Datos básicos ---
    enc("clave_acceso") = SXL(rootNode, "./*[local-name()='infoTributaria']/*[local-name()='claveAcceso']")
    If Len(enc("clave_acceso")) = 0 Then enc("clave_acceso") = FileNameNoExt(xmlPath)

    enc("tipo_comprobante") = SXL(infoTrib, "./*[local-name()='codDoc']")
    enc("ambiente") = SXL(infoTrib, "./*[local-name()='ambiente']")
    enc("ruc_emisor") = SXL(infoTrib, "./*[local-name()='ruc']")
    enc("razon_social_emisor") = SXL(infoTrib, "./*[local-name()='razonSocial']")
    enc("establecimiento") = SXL(infoTrib, "./*[local-name()='estab']")
    enc("punto_emision") = SXL(infoTrib, "./*[local-name()='ptoEmi']")
    enc("secuencial") = SXL(infoTrib, "./*[local-name()='secuencial']")

    ' --- Nro comprobante (EEE-PPP-SSSSSSSSS) ---
    enc("nro_comprobante") = FormatComprobante(enc("establecimiento"), enc("punto_emision"), enc("secuencial"))

    ' --- Documento sustentado (tipo/serie/secuencial/fecha) ---
    FillDocSust_RET rootNode, infoRet, enc

    ' --- Fechas y sujeto retenido ---
    enc("fecha_emision") = SXL(infoRet, "./*[local-name()='fechaEmision']")
    enc("periodo_fiscal") = SXL(infoRet, "./*[local-name()='periodoFiscal']")
    enc("ruc_sujeto") = SXL(infoRet, "./*[local-name()='identificacionSujetoRetenido']")
    enc("razon_social_sujeto") = SXL(infoRet, "./*[local-name()='razonSocialSujetoRetenido']")

    ' === Acumuladores y listas ===
    Dim baseIVA As Double, valIVA As Double, codesIVA As String, porcsIVA As String
    Dim baseREN As Double, valREN As Double, codesREN As String, porcsREN As String
    baseIVA = 0#: valIVA = 0#: codesIVA = "": porcsIVA = ""
    baseREN = 0#: valREN = 0#: codesREN = "": porcsREN = ""

    Dim found As Boolean: found = False

    ' --- Variante 1: docsSustento//retenciones/retencion ---
    Dim docSustentos As Object, ds As Object, rets As Object, r As Object
    Set docSustentos = rootNode.SelectNodes(".//*[local-name()='docsSustento']/*[local-name()='docSustento']")
    If Not docSustentos Is Nothing Then
        For Each ds In docSustentos
            Set rets = ds.SelectNodes(".//*[local-name()='retenciones']/*[local-name()='retencion']")
            If Not rets Is Nothing Then
                For Each r In rets
                    found = True
                    AddRetencionLinea_AndAcc enc("clave_acceso"), r, wsRD, _
                        baseIVA, valIVA, codesIVA, porcsIVA, _
                        baseREN, valREN, codesREN, porcsREN
                Next r
            End If
        Next ds
    End If

    ' --- Variante 2: impuestos/impuesto ---
    If Not found Then
        Dim impNodes As Object, n As Object
        Set impNodes = rootNode.SelectNodes(".//*[local-name()='impuestos']/*[local-name()='impuesto']")
        If Not impNodes Is Nothing Then
            For Each n In impNodes
                found = True
                AddRetencionLinea_AndAcc enc("clave_acceso"), n, wsRD, _
                    baseIVA, valIVA, codesIVA, porcsIVA, _
                    baseREN, valREN, codesREN, porcsREN
            Next n
        End If
    End If

    ' --- Variante 3: impuestosRetenidos/impuesto ---
    If Not found Then
        Dim impRetNodes As Object, n2 As Object
        Set impRetNodes = rootNode.SelectNodes(".//*[local-name()='impuestosRetenidos']/*[local-name()='impuesto']")
        If Not impRetNodes Is Nothing Then
            For Each n2 In impRetNodes
                found = True
                AddRetencionLinea_AndAcc enc("clave_acceso"), n2, wsRD, _
                    baseIVA, valIVA, codesIVA, porcsIVA, _
                    baseREN, valREN, codesREN, porcsREN
            Next n2
        End If
    End If

    If Not found Then
        Log wsL, xmlPath, "WARN", "No se hallaron nodos de retención en este comprobante."
    End If

    ' --- Resumen que va a la hoja Retenciones ---
    enc("cod_ret_iva") = codesIVA
    enc("porc_ret_iva") = porcsIVA
    enc("base_ret_iva") = baseIVA
    enc("valor_ret_iva") = valIVA

    enc("cod_ret_renta") = codesREN
    enc("porc_ret_renta") = porcsREN
    enc("base_ret_renta") = baseREN
    enc("valor_ret_renta") = valREN

    enc("total_retenido") = valIVA + valREN

    ' --- Escribir en Retenciones usando la CABECERA global ---
    WriteDictToRow wsR, nextRow(wsR), RET_HEADERS, enc

    Log wsL, xmlPath, "OK", "Importado (retención)"
    Exit Sub

EH:
    Log wsL, xmlPath, "ERROR", Err.Description
End Sub



' ---- Escribe RetDet y acumula por IVA/RENTA (códigos + porcentajes) ----

Private Sub AddRetencionLinea_AndAcc(ByVal clave As String, ByVal nodo As Object, _
                                     ByRef wsRD As Worksheet, _
                                     ByRef baseIVA As Double, ByRef valIVA As Double, ByRef codesIVA As String, ByRef porcsIVA As String, _
                                     ByRef baseREN As Double, ByRef valREN As Double, ByRef codesREN As String, ByRef porcsREN As String)

    Dim codImp As String, codRet As String
    Dim base As Double, porc As Double, valor As Double

    codImp = SXL(nodo, "./*[local-name()='codigo']")                 ' 1=RENTA, 2=IVA
    codRet = SXL(nodo, "./*[local-name()='codigoRetencion']")
    base = n(SXL(nodo, "./*[local-name()='baseImponible']"))
    porc = n(SXL(nodo, "./*[local-name()='porcentajeRetener']"))
    valor = n(SXL(nodo, "./*[local-name()='valorRetenido']"))

    ' --- Detalle (RetDet) ---
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d("clave_acceso") = clave
    d("impuesto") = codImp
    d("codigo") = codImp
    d("codigo_retencion") = codRet
    d("base_imponible") = base
    d("porcentaje_retener") = porc
    d("valor_retenido") = valor
    WriteDictToRow wsRD, nextRow(wsRD), Array( _
        "clave_acceso", "impuesto", "codigo", "codigo_retencion", _
        "base_imponible", "porcentaje_retener", "valor_retenido" _
    ), d

    ' --- Acumular por grupo ---
    If codImp = "2" Then            ' IVA
        baseIVA = baseIVA + base
        valIVA = valIVA + valor
        If InStr(1, ";" & codesIVA & ";", ";" & codRet & ";", vbTextCompare) = 0 Then
            codesIVA = IIf(Len(codesIVA) = 0, codRet, codesIVA & " + " & codRet)
        End If
        Dim sp As String: sp = CStr(porc)
        If InStr(1, ";" & porcsIVA & ";", ";" & sp & ";", vbTextCompare) = 0 Then
            porcsIVA = IIf(Len(porcsIVA) = 0, sp, porcsIVA & " + " & sp)
        End If

    ElseIf codImp = "1" Then        ' RENTA
        baseREN = baseREN + base
        valREN = valREN + valor
        If InStr(1, ";" & codesREN & ";", ";" & codRet & ";", vbTextCompare) = 0 Then
            codesREN = IIf(Len(codesREN) = 0, codRet, codesREN & " + " & codRet)
        End If
        Dim sp2 As String: sp2 = CStr(porc)
        If InStr(1, ";" & porcsREN & ";", ";" & sp2 & ";", vbTextCompare) = 0 Then
            porcsREN = IIf(Len(porcsREN) = 0, sp2, porcsREN & " + " & sp2)
        End If
    End If
End Sub


' === Ejecuta una vez para forzar encabezados actualizados ===
Public Sub RecrearEncabezadosRet()
    Dim wsR As Worksheet, wsRD As Worksheet
    Set wsR = ThisWorkbook.Worksheets("Retenciones")
    Set wsRD = ThisWorkbook.Worksheets("RetDet")

    ' Encabezados Retenciones (resumen por comprobante)
    Dim H_R As Variant
    H_R = Array( _
        "clave_acceso", "tipo_comprobante", "ambiente", _
        "ruc_emisor", "razon_social_emisor", "establecimiento", "punto_emision", "secuencial", _
        "fecha_emision", "periodo_fiscal", "ruc_sujeto", "razon_social_sujeto", _
        "cod_ret_iva", "porc_ret_iva", "base_ret_iva", "valor_ret_iva", _
        "cod_ret_renta", "porc_ret_renta", "base_ret_renta", "valor_ret_renta", _
        "total_retenido" _
    )

    wsR.Rows(1).ClearContents
    wsR.Range(wsR.Cells(1, 1), wsR.Cells(1, UBound(H_R) + 1)).Value = H_R
    wsR.Rows(1).Font.Bold = True

    ' Encabezados RetDet (detalle línea a línea)
    Dim H_D As Variant
    H_D = Array( _
        "clave_acceso", "impuesto", "codigo", "codigo_retencion", _
        "base_imponible", "porcentaje_retener", "valor_retenido" _
    )

    wsRD.Rows(1).ClearContents
    wsRD.Range(wsRD.Cells(1, 1), wsRD.Cells(1, UBound(H_D) + 1)).Value = H_D
    wsRD.Rows(1).Font.Bold = True

    wsR.Columns.AutoFit: wsRD.Columns.AutoFit
    MsgBox "Encabezados de Retenciones y RetDet actualizados.", vbInformation
End Sub




' ===== Detecta documento y nodos clave =====
Private Function FindDocRoot(ByVal dom As Object, ByRef docType As String, _
                             ByRef rootNode As Object, ByRef infoTrib As Object, _
                             ByRef infoNode As Object, ByRef detalles As Object) As Boolean
    On Error GoTo EH

    ' factura
    Set rootNode = dom.SelectSingleNode("/*[local-name()='factura']")
    If rootNode Is Nothing Then Set rootNode = dom.SelectSingleNode("//*[local-name()='factura']")
    If Not rootNode Is Nothing Then
        docType = "factura"
        Set infoTrib = rootNode.SelectSingleNode("./*[local-name()='infoTributaria']")
        Set infoNode = rootNode.SelectSingleNode("./*[local-name()='infoFactura']")
        Set detalles = rootNode.SelectSingleNode("./*[local-name()='detalles']")
        FindDocRoot = Not (infoTrib Is Nothing Or infoNode Is Nothing)
        Exit Function
    End If

    ' notaCredito
    Set rootNode = dom.SelectSingleNode("/*[local-name()='notaCredito']")
    If rootNode Is Nothing Then Set rootNode = dom.SelectSingleNode("//*[local-name()='notaCredito']")
    If Not rootNode Is Nothing Then
        docType = "notaCredito"
        Set infoTrib = rootNode.SelectSingleNode("./*[local-name()='infoTributaria']")
        Set infoNode = rootNode.SelectSingleNode("./*[local-name()='infoNotaCredito']")
        Set detalles = rootNode.SelectSingleNode("./*[local-name()='detalles']")
        FindDocRoot = Not (infoTrib Is Nothing Or infoNode Is Nothing)
        Exit Function
    End If

    ' notaDebito
    Set rootNode = dom.SelectSingleNode("/*[local-name()='notaDebito']")
    If rootNode Is Nothing Then Set rootNode = dom.SelectSingleNode("//*[local-name()='notaDebito']")
    If Not rootNode Is Nothing Then
        docType = "notaDebito"
        Set infoTrib = rootNode.SelectSingleNode("./*[local-name()='infoTributaria']")
        Set infoNode = rootNode.SelectSingleNode("./*[local-name()='infoNotaDebito']")
        Set detalles = Nothing ' no hay ítems
        FindDocRoot = Not (infoTrib Is Nothing Or infoNode Is Nothing)
        Exit Function
    End If

    ' comprobanteRetencion
    Set rootNode = dom.SelectSingleNode("/*[local-name()='comprobanteRetencion']")
    If rootNode Is Nothing Then Set rootNode = dom.SelectSingleNode("//*[local-name()='comprobanteRetencion']")
    If Not rootNode Is Nothing Then
        docType = "retencion"
        Set infoTrib = rootNode.SelectSingleNode("./*[local-name()='infoTributaria']")
        Set infoNode = rootNode.SelectSingleNode("./*[local-name()='infoCompRetencion']")
        Set detalles = Nothing
        FindDocRoot = Not (infoTrib Is Nothing Or infoNode Is Nothing)
        Exit Function
    End If

    FindDocRoot = False
    Exit Function
EH:
    FindDocRoot = False
End Function

' ===== Extrae XML interno de <comprobante><![CDATA[...]]> =====
Private Function ExtractInnerComprobante(ByVal outerDom As Object) As Object
    On Error GoTo EH
    Dim comp As Object: Set comp = outerDom.SelectSingleNode(".//*[local-name()='comprobante']")
    If comp Is Nothing Then Set ExtractInnerComprobante = Nothing: Exit Function

    Dim inner As String: inner = Trim$(CStr(comp.Text))
    If Len(inner) = 0 Then Set ExtractInnerComprobante = Nothing: Exit Function

    Dim dom2 As Object: Set dom2 = CreateObject("MSXML2.DOMDocument.6.0")
    dom2.async = False: dom2.validateOnParse = False
    dom2.SetProperty "SelectionLanguage", "XPath"

    If dom2.LoadXML(inner) Then Set ExtractInnerComprobante = dom2: Exit Function
    inner = Replace(inner, "&lt;", "<"): inner = Replace(inner, "&gt;", ">"): inner = Replace(inner, "&amp;", "&")
    If dom2.LoadXML(inner) Then Set ExtractInnerComprobante = dom2: Exit Function

    Set ExtractInnerComprobante = Nothing
    Exit Function
EH:
    Set ExtractInnerComprobante = Nothing
End Function

' ===== Envolturas de autorización: busca <comprobante> =====
Private Function TryLoadFromAutorizacion(ByVal xmlPath As String, ByRef domOut As Object) As Boolean
    On Error GoTo EH
    Dim d As Object: Set d = CreateObject("MSXML2.DOMDocument.6.0")
    d.async = False: d.validateOnParse = False
    d.SetProperty "SelectionLanguage", "XPath"
    If Not d.Load(xmlPath) Then TryLoadFromAutorizacion = False: Exit Function

    Dim comp As Object
    Set comp = d.SelectSingleNode(".//*[local-name()='autorizaciones']/*[local-name()='autorizacion']/*[local-name()='comprobante']")
    If comp Is Nothing Then Set comp = d.SelectSingleNode(".//*[local-name()='autorizacion']/*[local-name()='comprobante']")
    If comp Is Nothing Then Set comp = d.SelectSingleNode(".//*[local-name()='respuestaAutorizacionComprobante']//*[local-name()='comprobante']")
    If comp Is Nothing Then Set comp = d.SelectSingleNode(".//*[local-name()='comprobante']")
    If comp Is Nothing Then TryLoadFromAutorizacion = False: Exit Function

    Dim inner As String: inner = Trim$(CStr(comp.Text))

    Dim dom2 As Object: Set dom2 = CreateObject("MSXML2.DOMDocument.6.0")
    dom2.async = False: dom2.validateOnParse = False
    dom2.SetProperty "SelectionLanguage", "XPath"

    If dom2.LoadXML(inner) Then Set domOut = dom2: TryLoadFromAutorizacion = True: Exit Function
    inner = Replace(inner, "&lt;", "<"): inner = Replace(inner, "&gt;", ">"): inner = Replace(inner, "&amp;", "&")
    If dom2.LoadXML(inner) Then Set domOut = dom2: TryLoadFromAutorizacion = True: Exit Function

    TryLoadFromAutorizacion = False
    Exit Function
EH:
    TryLoadFromAutorizacion = False
End Function

' ===== Utilidades =====
' Asegura XPath sin explotar si el nodo es Nothing o no tiene OwnerDocument
Private Sub EnsureXPath(ByVal anyNode As Object)
    On Error Resume Next
    If Not anyNode Is Nothing Then
        Dim doc As Object
        Set doc = anyNode.OwnerDocument
        If Not doc Is Nothing Then doc.SetProperty "SelectionLanguage", "XPath"
    End If
End Sub


' convierte string a número respetando configuración regional
Private Function n(ByVal s As String) As Double
    Dim t As String, decSep As String, thouSep As String
    decSep = Application.International(xlDecimalSeparator)
    thouSep = Application.International(xlThousandsSeparator)
    t = Trim$(s)
    If Len(t) = 0 Then n = 0#: Exit Function
    t = Replace(t, " ", "")
    t = Replace(t, thouSep, "")
    If decSep = "," Then
        t = Replace(t, ".", ",")
    Else
        t = Replace(t, ",", ".")
    End If
    On Error Resume Next
    n = CDbl(t)
    If Err.Number <> 0 Then n = 0#
    On Error GoTo 0
End Function

Private Function CleanTarifa(ByVal s As String) As String
    Dim t As String
    t = Replace(Replace(Trim$(s), ",", "."), "%", "")
    If InStr(1, t, ".") > 0 Then CleanTarifa = CStr(CLng(val(t))) Else CleanTarifa = t
End Function

Private Function FileNameNoExt(ByVal p As String) As String
    Dim nm As String: nm = Mid$(p, InStrRev(p, "\") + 1)
    If InStrRev(nm, ".") > 0 Then FileNameNoExt = Left$(nm, InStrRev(nm, ".") - 1) Else FileNameNoExt = nm
End Function

Private Function PickFolder(ByVal title As String) As String
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .title = title
        .AllowMultiSelect = False
        If .Show = -1 Then PickFolder = .SelectedItems(1) Else PickFolder = ""
    End With
End Function

Private Function EnsureSheet(ByVal name As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    'Set ws = ThisWorkbook.Worksheets(name)
    Set ws = TWB.Worksheets(name)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = TWB.Worksheets.Add(After:=Worksheets(Worksheets.Count))
        ws.name = name
    End If
    Set EnsureSheet = ws
End Function

Private Sub EnsureHeadersIfEmpty(ByVal ws As Worksheet, ByVal headers As Variant)
    If Application.WorksheetFunction.CountA(ws.Rows(1)) = 0 Then
        Dim i As Long
        For i = LBound(headers) To UBound(headers)
            ws.Cells(1, i + 1).Value = headers(i)
            ws.Cells(1, i + 1).Font.Bold = True
        Next i
        ws.Columns.AutoFit
    End If
    ApplyTextFormatByName ws, headers
End Sub

Private Sub ApplyTextFormatByName(ByVal ws As Worksheet, ByVal headers As Variant)
    Dim i As Long, name As String
    For i = LBound(headers) To UBound(headers)
        name = CStr(headers(i))
        If (ws.name = "Facturas" And IsInArray(name, TEXT_COLS_FACTURAS)) _
           Or (ws.name = "Detalle" And IsInArray(name, TEXT_COLS_DETALLE)) _
           Or (ws.name = "Retenciones" And IsInArray(name, TEXT_COLS_RET)) _
           Or (ws.name = "RetDet" And IsInArray(name, TEXT_COLS_RETDET)) Then
            ws.Columns(i + 1).NumberFormat = "@"
        End If
    Next i
End Sub

Private Function IsInArray(ByVal v As String, ByVal arr As Variant) As Boolean
    Dim j As Long
    For j = LBound(arr) To UBound(arr)
        If LCase$(arr(j)) = LCase$(v) Then IsInArray = True: Exit Function
    Next j
    IsInArray = False
End Function

Private Function nextRow(ByVal ws As Worksheet) As Long
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then
        nextRow = 2
    Else
        nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        If nextRow < 2 Then nextRow = 2
    End If
End Function

Private Sub WriteDictToRow(ByVal ws As Worksheet, ByVal rowIx As Long, ByVal headers As Variant, ByVal dict As Object)
    Dim i As Long, key As String, val As Variant, asText As Boolean

    For i = LBound(headers) To UBound(headers)
        key = CStr(headers(i))
        If dict.Exists(key) Then
            val = dict(key)

            ' ¿Esta columna es de TEXTO?
            asText = (ws.name = "Facturas" And IsInArray(key, TEXT_COLS_FACTURAS)) _
                     Or (ws.name = "Detalle" And IsInArray(key, TEXT_COLS_DETALLE)) _
                     Or (ws.name = "Retenciones" And IsInArray(key, TEXT_COLS_RET)) _
                     Or (ws.name = "RetDet" And IsInArray(key, TEXT_COLS_RETDET))

            ' --- FECHA: siempre dd-mm-yyyy ---
            If LCase$(key) = "fecha_emision" And Len(val) > 0 Then
                Dim d As Date
                If TryParseFechaEmision(CStr(val), d) Then
                    ws.Cells(rowIx, i + 1).NumberFormat = "dd-mm-yyyy"
                    ws.Cells(rowIx, i + 1).Value = d
                Else
                    ws.Cells(rowIx, i + 1).NumberFormat = "@"
                    ws.Cells(rowIx, i + 1).Value = CStr(val)
                End If

            ' --- TEXTO ---
            ElseIf asText Then
                ws.Cells(rowIx, i + 1).NumberFormat = "@"
                ws.Cells(rowIx, i + 1).Value = CStr(val)

            ' --- NUMÉRICO / GENERAL ---
            Else
                ws.Cells(rowIx, i + 1).Value = val
            End If
        End If
    Next i
End Sub


Private Sub CopyDict(ByVal src As Object, ByVal dst As Object)
    Dim k As Variant
    For Each k In src.Keys: dst(k) = src(k): Next k
End Sub

' ===== Totales por impuesto =====
Private Function SumarImpuestosLocalName(ByVal infoNode As Object) As Object
    Dim r As Object: Set r = CreateObject("Scripting.Dictionary")
    r("iva0") = 0#: r("iva5") = 0#: r("iva12") = 0#: r("iva15") = 0#
    r("noObjeto") = 0#: r("exento") = 0#
    r("ivaTotal") = 0#: r("iceTotal") = 0#
    If infoNode Is Nothing Then Set SumarImpuestosLocalName = r: Exit Function

    Dim tot As Object, ti As Object
    Set tot = infoNode.SelectSingleNode("./*[local-name()='totalConImpuestos']")
    If tot Is Nothing Then Set SumarImpuestosLocalName = r: Exit Function

    For Each ti In tot.SelectNodes("./*[local-name()='totalImpuesto']")
        Dim codigo As String, codPorc As String, tarifa As String
        Dim base As Double, valor As Double
        codigo = SXL(ti, "./*[local-name()='codigo']")
        codPorc = SXL(ti, "./*[local-name()='codigoPorcentaje']")
        tarifa = SXL(ti, "./*[local-name()='tarifa']")
        base = n(SXL(ti, "./*[local-name()='baseImponible']"))
        valor = n(SXL(ti, "./*[local-name()='valor']"))
        If codigo = "2" Then
            Select Case CleanTarifa(tarifa)
                Case "0": r("iva0") = r("iva0") + base
                Case "5": r("iva5") = r("iva5") + base
                Case "12": r("iva12") = r("iva12") + base
                Case "15": r("iva15") = r("iva15") + base
            End Select
            r("ivaTotal") = r("ivaTotal") + valor
            If codPorc = "6" Then r("noObjeto") = r("noObjeto") + base
            If codPorc = "7" Then r("exento") = r("exento") + base
        ElseIf codigo = "3" Then
            r("iceTotal") = r("iceTotal") + valor
        End If
    Next ti

    Set SumarImpuestosLocalName = r
End Function

' ===== Fechas =====
Public Function TryParseISODate(ByVal s As String, ByRef d As Date) As Boolean
    On Error GoTo EH
    Dim y As Integer, m As Integer, dD As Integer
    Dim p() As String: p = Split(Replace(Trim$(s), "/", "-"), "-")
    If UBound(p) = 2 Then
        y = CInt(p(0)): m = CInt(p(1)): dD = CInt(p(2))
        d = DateSerial(y, m, dD): TryParseISODate = True
    End If
    Exit Function
EH:
    TryParseISODate = False
End Function

' Interpreta:
' - "DD/MM/YYYY", "DD-MM-YYYY"  -> DMY (preferido)
' - "YYYY-MM-DD"                -> YMD (ISO)
Public Function TryParseFechaEmision(ByVal s As String, ByRef d As Date) As Boolean
    On Error GoTo EH
    Dim t As String: t = Trim$(s)
    If Len(t) = 0 Then TryParseFechaEmision = False: Exit Function

    Dim p() As String: p = Split(Replace(t, "/", "-"), "-")
    If UBound(p) <> 2 Then TryParseFechaEmision = False: Exit Function

    Dim a As Integer, b As Integer, c As Integer
    a = val(p(0)): b = val(p(1)): c = val(p(2))

    If Len(p(0)) = 4 Then
        ' ISO: YYYY-MM-DD
        d = DateSerial(a, b, c)
    Else
        ' Por defecto: DD-MM-YYYY (Ecuador)
        d = DateSerial(c, b, a)
    End If
    TryParseFechaEmision = True
    Exit Function
EH:
    TryParseFechaEmision = False
End Function


' ===== XPath helper (texto) =====
' ===== XPath helper (texto) ROBUSTO =====
' Intenta 4 estrategias:
' 1) node.selectSingleNode(xpath)
' 2) node.OwnerDocument.selectSingleNode(xpath)
' 3) si empieza con "./", intenta con ".//" (búsqueda profunda relativa)
' 4) si aún nada, intenta con "//" (absoluta en todo el documento)
' ===== XPath helper (texto) ROBUSTO =====
' Intenta 4 estrategias y no lanza 424 si el nodo/documento es Nothing.
Private Function SXL(ByVal node As Object, ByVal xpath As String) As String
    On Error GoTo EH
    If node Is Nothing Then SXL = "": Exit Function

    Dim doc As Object, n As Object
    On Error Resume Next
    Set doc = node.OwnerDocument
    If Not doc Is Nothing Then doc.SetProperty "SelectionLanguage", "XPath"
    On Error GoTo 0

    ' 1) relativo tal cual
    On Error Resume Next
    Set n = node.SelectSingleNode(xpath)
    On Error GoTo 0
    If Not n Is Nothing Then GoTo HAVE

    ' 2) en el documento (por si el contexto falla)
    If Not doc Is Nothing Then
        On Error Resume Next
        Set n = doc.SelectSingleNode(xpath)
        On Error GoTo 0
        If Not n Is Nothing Then GoTo HAVE
    End If

    ' 3) si empieza con "./", intenta .// (búsqueda profunda)
    If Left$(xpath, 2) = "./" Then
        On Error Resume Next
        Set n = node.SelectSingleNode(".//" & Mid$(xpath, 3))
        If n Is Nothing And Not doc Is Nothing Then Set n = doc.SelectSingleNode(".//" & Mid$(xpath, 3))
        On Error GoTo 0
        If Not n Is Nothing Then GoTo HAVE
    End If

    ' 4) absoluta en todo el documento
    If Not doc Is Nothing Then
        On Error Resume Next
        If Left$(xpath, 2) = "./" Then
            Set n = doc.SelectSingleNode("//" & Mid$(xpath, 3))
        Else
            Set n = doc.SelectSingleNode("//" & xpath)
        End If
        On Error GoTo 0
        If Not n Is Nothing Then GoTo HAVE
    End If

    SXL = ""
    Exit Function

HAVE:
    On Error Resume Next
    SXL = Trim$(CStr(n.Text))
    On Error GoTo 0
    Exit Function
EH:
    SXL = ""
End Function


' ===== LOG =====
Private Sub InitLogSheet(ByVal ws As Worksheet)
    If Application.WorksheetFunction.CountA(ws.Rows(1)) = 0 Then
        ws.Range("A1:D1").Value = Array("archivo", "estado", "detalle", "timestamp")
        ws.Rows(1).Font.Bold = True
    End If
End Sub

Private Sub Log(ByVal ws As Worksheet, ByVal archivo As String, ByVal estado As String, ByVal detalle As String)
    Dim r As Long: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = archivo
    ws.Cells(r, 2).Value = estado
    ws.Cells(r, 3).Value = detalle
    ws.Cells(r, 4).Value = Now
End Sub


' ============================
'   EXPORTAR RESULTADO A .XLSX (sin macros)
'   Cópialo al final del módulo v5 (no borres lo anterior)
' ============================

Public Sub GuardarResultado()
    On Error GoTo EH

    ' Hojas a exportar (en este orden). Si no existen, se omiten.
    Dim sheetsToExport As Variant
    sheetsToExport = Array("Facturas", "Detalle", "Retenciones", "RetDet", "LOG")

    ' Crear libro destino (en blanco)
    Dim wbOut As Workbook
    Set wbOut = Application.Workbooks.Add(xlWBATWorksheet) ' 1 hoja

    Dim i As Long, added As Boolean
    added = False

    ' Copiar cada hoja si existe (valores + formato)
    For i = LBound(sheetsToExport) To UBound(sheetsToExport)
        If SheetExists(CStr(sheetsToExport(i))) Then
            CopySheetValues ThisWorkbook.Worksheets(CStr(sheetsToExport(i))), wbOut
            added = True
        End If
    Next i

    If Not added Then
        MsgBox "No encontré hojas para exportar.", vbExclamation
        wbOut.Close SaveChanges:=False
        Exit Sub
    End If

    ' Elimina la hoja inicial vacía si no se usó
    If wbOut.Worksheets.Count > 1 Then
        If LCase$(wbOut.Worksheets(1).name) Like "hoja*" Then
            Application.DisplayAlerts = False
            wbOut.Worksheets(1).Delete
            Application.DisplayAlerts = True
        End If
    End If

    ' Sugerir nombre y pedir ruta
    Dim defaultName As String
    defaultName = "SRI_" & Format(Now, "yyyymmdd_hhnn") & ".xlsx"

    Dim targetPath As Variant
    targetPath = Application.GetSaveAsFilename(InitialFileName:=defaultName, _
                    FileFilter:="Excel Workbook (*.xlsx), *.xlsx", _
                    title:="Guardar resultado como")
    If targetPath = False Then
        wbOut.Close SaveChanges:=False
        Exit Sub
    End If

    ' Guardar como .xlsx (sin macros)
    Application.DisplayAlerts = False
    wbOut.SaveAs fileName:=targetPath, FileFormat:=xlOpenXMLWorkbook ' 51
    Application.DisplayAlerts = True
    wbOut.Close SaveChanges:=True

    MsgBox "¡Listo! Exportado a:" & vbCrLf & CStr(targetPath), vbInformation
    Exit Sub

EH:
    On Error Resume Next
    Application.DisplayAlerts = True
    If Not wbOut Is Nothing Then wbOut.Close SaveChanges:=False
    MsgBox "No se pudo guardar: " & Err.Description, vbExclamation
End Sub

' --- Helpers ---

Private Function SheetExists(ByVal sheetName As String) As Boolean
    On Error Resume Next
    SheetExists = Not ThisWorkbook.Worksheets(sheetName) Is Nothing
    On Error GoTo 0
End Function

' Copia valores + formatos de una hoja origen a una nueva hoja en wbOut
Private Sub CopySheetValues(ByVal wsSrc As Worksheet, ByVal wbOut As Workbook)
    Dim wsDst As Worksheet

    ' Crear hoja destino
    If wbOut.Worksheets.Count = 1 And LCase$(wbOut.Worksheets(1).name) Like "hoja*" Then
        Set wsDst = wbOut.Worksheets(1)
        wsDst.name = SafeSheetName(wsSrc.name)
    Else
        Set wsDst = wbOut.Worksheets.Add(After:=wbOut.Worksheets(wbOut.Worksheets.Count))
        wsDst.name = SafeSheetName(wsSrc.name)
    End If

    ' Copiar rango usado
    If Application.WorksheetFunction.CountA(wsSrc.Cells) = 0 Then Exit Sub

    Dim rng As Range
    Set rng = wsSrc.UsedRange

    ' Copiar con formato (primero formato/anchos, luego valores)
    wsDst.Cells(1, 1).Resize(rng.Rows.Count, rng.Columns.Count).NumberFormat = rng.NumberFormat
    rng.Copy
    wsDst.Range("A1").PasteSpecial xlPasteColumnWidths
    wsDst.Range("A1").PasteSpecial xlPasteFormats
    wsDst.Range("A1").PasteSpecial xlPasteValues

    Application.CutCopyMode = False
    wsDst.Cells.EntireColumn.AutoFit
End Sub

Private Function SafeSheetName(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, ":", "_")
    t = Replace(t, "\", "_")
    t = Replace(t, "/", "_")
    t = Replace(t, "?", "_")
    t = Replace(t, "*", "_")
    t = Replace(t, "[", "_")
    t = Replace(t, "]", "_")
    If Len(t) = 0 Then t = "Hoja"
    If Len(t) > 31 Then t = Left$(t, 31)
    SafeSheetName = t
End Function

'======================
' BOTÓN: LIMPIAR
'======================
Public Sub LimpiarDatos()
    Dim ws As Worksheet
    Dim hojas As Variant, h As Variant

    ' Hojas a limpiar (solo contenido desde la fila 2)
    hojas = Array("Facturas", "Detalle", "Retenciones", "RetDet")

    For Each h In hojas
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(h))
        On Error GoTo 0
        If Not ws Is Nothing Then
            If ws.UsedRange.Rows.Count > 1 Then
                ws.Rows("2:" & ws.Rows.Count).ClearContents
            End If
        End If
        Set ws = Nothing
    Next h

    ' El LOG se borra completo (deja encabezados si existen)
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("LOG")
    On Error GoTo 0
    If Not ws Is Nothing Then
        If ws.UsedRange.Rows.Count > 1 Then
            ws.Rows("2:" & ws.Rows.Count).ClearContents
        End If
    End If

    ' LIMPIAR LA HOJA LOG_Rename
    Call LimpiarLOGRename

    MsgBox "Datos limpiados (se conservaron los encabezados).", vbInformation
End Sub

'======================
' BOTÓN / AUTO: AUTOFORMATO
' - Ajusta anchos de columnas
' - Aplica formatos numéricos típicos
'======================
Public Sub FormatearColumnas(Optional ByVal silent As Boolean = False)
    Dim ws As Worksheet
    Dim hojas As Variant, h As Variant

    hojas = Array("Facturas", "Detalle", "Retenciones", "RetDet", "LOG")

    For Each h In hojas
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(h))
        On Error GoTo 0
        If Not ws Is Nothing Then
            ws.Cells.EntireColumn.AutoFit
            ApplyNumberFormats ws
        End If
        Set ws = Nothing
    Next h

    If Not silent Then
        MsgBox "Autoformato aplicado.", vbInformation
    End If
End Sub


' Helper para FormatearColumnas:
' Aplica formatos numéricos y de fecha según encabezados
Private Sub ApplyNumberFormats(ByVal ws As Worksheet)
    Dim lastCol As Long, lastRow As Long, c As Long
    If Application.WorksheetFunction.CountA(ws.Cells) = 0 Then Exit Sub

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    For c = 1 To lastCol
        Dim h As String: h = LCase$(CStr(ws.Cells(1, c).Value))

        ' ---- FECHAS ----
        If h = "fecha_emision" Then
            ws.Range(ws.Cells(2, c), ws.Cells(lastRow, c)).NumberFormat = "dd-mm-yyyy"

        ' ---- PORCENTAJES (resumen y detalle) ----
        ElseIf h = "porc_ret_iva" Or h = "porc_ret_renta" _
            Or h = "porcentaje_retener" _
            Or h Like "*porcentaje*" Then
            ws.Range(ws.Cells(2, c), ws.Cells(lastRow, c)).NumberFormat = "0.00"

        ' ---- CANTIDADES ----
        ElseIf h Like "*cantidad*" Then
            ws.Range(ws.Cells(2, c), ws.Cells(lastRow, c)).NumberFormat = "0.0000"

        ' ---- MONTOS MONETARIOS (incluye retenciones) ----
        ElseIf h Like "precio*" Or h Like "*valor*" Or h Like "base*" _
           Or h Like "*subtotal*" Or h Like "*iva*" Or h Like "*descuento*" _
           Or h Like "*total*" Or h Like "*propina*" _
           Or h = "base_imponible" Or h = "valor_retenido" _
           Or h = "base_ret_iva" Or h = "base_ret_renta" _
           Or h = "valor_ret_iva" Or h = "valor_ret_renta" _
           Or h = "total_retenido" Then
            ws.Range(ws.Cells(2, c), ws.Cells(lastRow, c)).NumberFormat = "#,##0.00"
        End If
    Next c
End Sub

Public Sub ForzarEncabezadosRet()
    Dim wsR As Worksheet, wsRD As Worksheet
    Set wsR = ThisWorkbook.Worksheets("Retenciones")
    Set wsRD = ThisWorkbook.Worksheets("RetDet")

    ' Retenciones (resumen por comprobante)
    Dim H_R As Variant
    H_R = Array( _
        "clave_acceso", "tipo_comprobante", "ambiente", _
        "ruc_emisor", "razon_social_emisor", "establecimiento", "punto_emision", "secuencial", _
        "fecha_emision", "periodo_fiscal", "ruc_sujeto", "razon_social_sujeto", _
        "cod_ret_iva", "porc_ret_iva", "base_ret_iva", "valor_ret_iva", _
        "cod_ret_renta", "porc_ret_renta", "base_ret_renta", "valor_ret_renta", _
        "total_retenido" _
    )

    wsR.Rows(1).ClearContents
    wsR.Range(wsR.Cells(1, 1), wsR.Cells(1, UBound(H_R) + 1)).Value = H_R
    wsR.Rows(1).Font.Bold = True

    ' RetDet (detalle línea a línea)
    Dim H_D As Variant
    H_D = Array( _
        "clave_acceso", "impuesto", "codigo", "codigo_retencion", _
        "base_imponible", "porcentaje_retener", "valor_retenido" _
    )
    wsRD.Rows(1).ClearContents
    wsRD.Range(wsRD.Cells(1, 1), wsRD.Cells(1, UBound(H_D) + 1)).Value = H_D
    wsRD.Rows(1).Font.Bold = True

    wsR.Cells.EntireColumn.AutoFit
    wsRD.Cells.EntireColumn.AutoFit
    MsgBox "Encabezados actualizados en Retenciones y RetDet.", vbInformation
End Sub

'===================== UTILIDADES =====================
Private Function ToD(ByVal v As Variant) As Double
    On Error GoTo f
    If IsEmpty(v) Or IsNull(v) Then
        ToD = 0#
    Else
        ToD = CDbl(Replace(CStr(v), ",", "."))
    End If
    Exit Function
f:
    ToD = 0#
End Function

Private Function GetD(ByVal enc As Object, ByVal key As String) As Double
    On Error Resume Next
    If Not enc Is Nothing And enc.Exists(key) Then
        GetD = ToD(enc(key))
    Else
        GetD = 0#
    End If
    On Error GoTo 0
End Function

Private Sub AppendLogFix(ByVal context As String, ByVal oldVal As Double, ByVal newVal As Double)
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = TWB.Worksheets("LOG")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = TWB.Worksheets.Add
        ws.name = "LOG"
        ws.Range("A1:D1").Value = Array("archivo", "estado", "detalle", "timestamp")
        ws.Rows(1).Font.Bold = True
    End If
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = context
    ws.Cells(r, 2).Value = "INFO"
    ws.Cells(r, 3).Value = "Ajuste subtotal_iva_0: de " & Format(oldVal, "0.00") & " a " & Format(newVal, "0.00")
    ws.Cells(r, 4).Value = Now
End Sub
'======================================================

'=============== AJUSTE SUBTOTAL IVA 0% ===============
' Regla principal: subtotal_iva_0 = total_sin_impuestos - (subtotal_iva_5 + _12 + _15 + no_objeto + exento)
' Regla alternativa (fallback): valor_total - (bases gravadas) - iva_total
Public Sub FixSubtotalIVA0_FromHeader(ByRef enc As Object, Optional ByVal context As String = "")
    Dim ts As Double, s0 As Double, s5 As Double, s12 As Double, s15 As Double
    Dim sNoObj As Double, sExen As Double, nuevo0 As Double
    Dim vtot As Double, iva As Double, gravadas As Double, alt0 As Double

    ' Lee encabezados que tú ya cargas en enc
    ts = GetD(enc, "total_sin_impuestos")
    s0 = GetD(enc, "subtotal_iva_0")
    s5 = GetD(enc, "subtotal_iva_5")
    s12 = GetD(enc, "subtotal_iva_12")
    s15 = GetD(enc, "subtotal_iva_15")
    sNoObj = GetD(enc, "subtotal_no_objeto")
    sExen = GetD(enc, "subtotal_exento")
    vtot = GetD(enc, "valor_total")
    iva = GetD(enc, "iva_total")

    ' Regla 1 (preferida): usa total_sin_impuestos
    nuevo0 = Round(ts - s5 - s12 - s15 - sNoObj - sExen, 2)

    ' Si total_sin_impuestos viene 0 o inconsistente, aplica Regla 2
    If (ts = 0 And (vtot > 0 Or iva > 0)) Or Abs(nuevo0) > (Abs(ts) + 0.01) Then
        gravadas = s12 + s15                          ' si manejas otras tarifas, súmalas aquí
        alt0 = Round(vtot - gravadas - iva, 2)
        ' Corrige borde: -0,01 -> 0
        If alt0 < 0 And Abs(alt0) < 0.01 Then alt0 = 0
        nuevo0 = alt0
    End If

    ' Normaliza borde de redondeo
    If nuevo0 < 0 And Abs(nuevo0) < 0.01 Then nuevo0 = 0

    ' Si difiere, actualiza y LOG
    If Abs(nuevo0 - s0) >= 0.01 Then
        enc("subtotal_iva_0") = nuevo0
        Dim ctx As String
        If Len(context) > 0 Then
            ctx = context
        ElseIf enc.Exists("clave_acceso") Then
            ctx = CStr(enc("clave_acceso"))
        Else
            ctx = "(sin clave)"
        End If
        AppendLogFix ctx, s0, nuevo0
    End If
End Sub
'======================================================

' ===== Helpers de formato y serie =====
Private Function PadLeft(ByVal s As String, ByVal n As Long, Optional ByVal ch As String = "0") As String
    s = Trim$(CStr(s))
    If Len(s) >= n Then PadLeft = Right$(s, n) Else PadLeft = String$(n - Len(s), ch) & s
End Function

Private Function FormatComprobante(ByVal estab As String, ByVal pto As String, ByVal secu As String) As String
    FormatComprobante = PadLeft(estab, 3) & "-" & PadLeft(pto, 3) & "-" & PadLeft(secu, 9)
End Function

Private Sub SplitSerieFromNumero(ByVal num As String, ByRef estab As String, ByRef pto As String, ByRef secu As String)
    Dim t() As String
    num = Trim$(num)
    If InStr(num, "-") > 0 Then
        t = Split(num, "-")
        If UBound(t) >= 2 Then estab = t(0): pto = t(1): secu = t(2)
    ElseIf Len(num) >= 15 Then
        estab = Mid$(num, 1, 3): pto = Mid$(num, 4, 3): secu = Mid$(num, 7)
    Else
        estab = "": pto = "": secu = num
    End If
End Sub

' ===== Doc. sustentado para NOTA DE CRÉDITO =====
Private Sub FillDocSust_NC(ByVal infoNC As Object, ByRef enc As Object)
    Dim tipo As String, num As String, fec As String, e As String, p As String, s As String
    tipo = SXL(infoNC, "./*[local-name()='codDocModificado']")
    num = SXL(infoNC, "./*[local-name()='numDocModificado']")
    fec = SXL(infoNC, "./*[local-name()='fechaEmisionDocSustento']")
    SplitSerieFromNumero num, e, p, s
    enc("doc_sust_tipo") = tipo
    enc("doc_sust_serie") = IIf(Len(e) Or Len(p), PadLeft(e, 3) & "-" & PadLeft(p, 3), "")
    enc("doc_sust_secuencial") = PadLeft(s, 9)
    enc("doc_sust_fecha") = fec
End Sub

' ===== Doc. sustentado para RETENCIÓN (robusto) =====
Private Sub FillDocSust_RET(ByVal rootNode As Object, ByVal infoRet As Object, ByRef enc As Object)
    Dim tipo As String, num As String, fec As String, e As String, p As String, s As String

    ' Buscar en infoRet y, si no hay, en todo el XML (PACs varían etiquetas)
    tipo = SXL(infoRet, "./*[local-name()='codDocSustento']")
    If Len(tipo) = 0 Then tipo = SXL(rootNode, ".//*[local-name()='codDocSustento' or local-name()='codDocModificado']")

    num = SXL(infoRet, "./*[local-name()='numDocSustento']")
    If Len(num) = 0 Then num = SXL(rootNode, ".//*[local-name()='numDocSustento' or local-name()='numeroDocSustento' or local-name()='numDocModificado']")

    fec = SXL(infoRet, "./*[local-name()='fechaEmisionDocSustento']")
    If Len(fec) = 0 Then fec = SXL(rootNode, ".//*[local-name()='fechaEmisionDocSustento' or local-name()='fechaEmision']")

    SplitSerieFromNumero num, e, p, s
    enc("doc_sust_tipo") = tipo
    enc("doc_sust_serie") = IIf(Len(e) Or Len(p), PadLeft(e, 3) & "-" & PadLeft(p, 3), "")
    enc("doc_sust_secuencial") = PadLeft(s, 9)
    enc("doc_sust_fecha") = fec
End Sub

' ====== UTIL: indice de columna por encabezado (fila 1) ======
Private Function ColByHeader(ByVal ws As Worksheet, ByVal header As String) As Long
    Dim c As Range
    Set c = ws.Rows(1).Find(What:=header, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If c Is Nothing Then ColByHeader = 0 Else ColByHeader = c.Column
End Function

' ====== UTIL: escribe encabezados exactos y limpia sobrantes ======
Private Sub WriteHeadersExact(ByVal ws As Worksheet, ByRef headers As Variant)
    Dim j As Long
    ws.Rows("1:1").ClearContents
    For j = LBound(headers) To UBound(headers)
        ws.Cells(1, j + 1).Value = CStr(headers(j))
    Next j
    ' Borra columnas sobrantes a la derecha de los encabezados previstos
    If ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column > UBound(headers) + 1 Then
        ws.Range(ws.Cells(1, UBound(headers) + 2), ws.Cells(1, ws.Columns.Count)).EntireColumn.Delete
    End If
    ws.Rows(1).Font.Bold = True
End Sub

' ====== UTIL: aplica formato Texto a un conjunto de campos ======
Private Sub ForceTextCols(ByVal ws As Worksheet, ByRef textCols As Variant)
    Dim lastRow As Long, k As Long, c As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then lastRow = 2
    For k = LBound(textCols) To UBound(textCols)
        c = ColByHeader(ws, CStr(textCols(k)))
        If c > 0 Then ws.Range(ws.Cells(2, c), ws.Cells(lastRow, c)).NumberFormat = "@"
    Next k
End Sub

' ====== REPARA encab. de Facturas, Detalle, Retenciones y RetDet ======
Public Sub RepararEncabezados()
    ' Inicializa arrays de encabezados/columnas texto
    InitHeaders
    InitTextCols

    On Error Resume Next
    Dim wsF As Worksheet, wsD As Worksheet, wsR As Worksheet, wsRD As Worksheet
    Set wsF = ThisWorkbook.Worksheets("Facturas")
    Set wsD = ThisWorkbook.Worksheets("Detalle")
    Set wsR = ThisWorkbook.Worksheets("Retenciones")
    Set wsRD = ThisWorkbook.Worksheets("RetDet")
    On Error GoTo 0

    If Not wsF Is Nothing Then
        WriteHeadersExact wsF, FACTURAS_HEADERS
        ForceTextCols wsF, TEXT_COLS_FACTURAS
    End If

    If Not wsD Is Nothing Then
        WriteHeadersExact wsD, DETALLE_HEADERS
        ForceTextCols wsD, TEXT_COLS_DETALLE
    End If

    If Not wsR Is Nothing Then
        WriteHeadersExact wsR, RET_HEADERS
        ForceTextCols wsR, TEXT_COLS_RET
    End If

    If Not wsRD Is Nothing Then
        WriteHeadersExact wsRD, RETDET_HEADERS
        ForceTextCols wsRD, TEXT_COLS_RETDET
    End If

    ' === AUTOFORMATO automático (silencioso) ===
    FormatearColumnas True

    MsgBox "Encabezados reparados y autoformato aplicado. Ahora ejecuta IMPORTAR.", vbInformation
End Sub


Public Sub RepararSoloFacturasHeaders()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Facturas")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "No existe la hoja 'Facturas'.", vbExclamation
        Exit Sub
    End If

    ' Reescribe exactamente los encabezados de Facturas y aplica TEXTO a columnas clave
    InitHeaders
    InitTextCols
    WriteHeadersExact ws, FACTURAS_HEADERS
    ForceTextCols ws, TEXT_COLS_FACTURAS

    ' === AUTOFORMATO automático (silencioso) ===
    FormatearColumnas True

    MsgBox "Encabezados reparados en 'Facturas' y autoformato aplicado.", vbInformation
End Sub

'========================================================
'      GENERAR PDF POR CADA COMPROBANTE
'  - Crea/usa una hoja PLANTILLA_PDF para render
'  - Busca dónde está el XML (carpeta raíz + subcarpetas)
'  - Guarda el PDF en la misma carpeta del XML
'  - Inserta hipervínculo en la celda nro_comprobante
'========================================================

Public Sub GenerarPDFsDesdeExcel()
    On Error GoTo EH

    Dim root As String
    root = PickFolder("Elige la carpeta RAÍZ donde están los XML (y donde quieres dejar los PDF)")
    If Len(root) = 0 Then Exit Sub

    Application.ScreenUpdating = False

    ' 1) Asegurar plantilla
    Dim wsT As Worksheet
    Set wsT = EnsurePdfTemplate()

    ' 2) Generar Facturas/NC/ND
    If SheetExists("Facturas") Then
        GenerarPDFs_en_Hoja ThisWorkbook.Worksheets("Facturas"), wsT, root, True
    End If

    ' 3) Generar Retenciones
    If SheetExists("Retenciones") Then
        GenerarPDFs_en_Hoja ThisWorkbook.Worksheets("Retenciones"), wsT, root, False
    End If

    Application.ScreenUpdating = True
    MsgBox "PDFs generados (donde hubo coincidencia de XML) y vinculados en nro_comprobante.", vbInformation
    Exit Sub
EH:
    Application.ScreenUpdating = True
    MsgBox "GenerarPDFsDesdeExcel: " & Err.Description, vbExclamation
End Sub

'--------------------------------------------------------
' Crea (si no existe) la hoja de plantilla para PDF.
'--------------------------------------------------------
Private Function EnsurePdfTemplate() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("PLANTILLA_PDF")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=Worksheets(Worksheets.Count))
        ws.name = "PLANTILLA_PDF"

        With ws.PageSetup
            .Orientation = xlPortrait
            .PaperSize = xlPaperA4
            .LeftMargin = Application.CentimetersToPoints(1.5)
            .RightMargin = Application.CentimetersToPoints(1.5)
            .TopMargin = Application.CentimetersToPoints(1.7)
            .BottomMargin = Application.CentimetersToPoints(1.7)
            .FitToPagesWide = 1
            .FitToPagesTall = 1
            .Zoom = False
            .CenterHorizontally = True
        End With

        ws.Cells.Font.name = "Calibri"
        ws.Cells.Font.Size = 10
        ws.Columns("A:B").ColumnWidth = 22
        ws.Columns("C:D").ColumnWidth = 45

        ' Encabezados (etiquetas fijas)
        ws.Range("A1:D1").Merge
        ws.Range("A1").Value = "COMPROBANTE"
        ws.Range("A1").Font.Size = 14
        ws.Range("A1").Font.Bold = True
        ws.Range("A1").HorizontalAlignment = xlCenter

        Dim lab As Variant
        lab = Array( _
            "Tipo", "tipo", _
            "Nro comprobante", "nro", _
            "Fecha emisión", "fec", _
            "Emisor (RUC)", "rucE", _
            "Razón social emisor", "rzE", _
            "Establecimiento", "est", _
            "Pto Emisión", "pto", _
            "Secuencial", "sec", _
            "Comprador/Sujeto", "cli", _
            "Razón social comprador/sujeto", "rzC", _
            "Valor total / Total retenido", "tot" _
        )

        Dim i As Long, fila As Long: fila = 3
        For i = LBound(lab) To UBound(lab) Step 2
            ws.Cells(fila, 1).Value = lab(i)
            ws.Cells(fila, 1).Font.Bold = True
            ws.Cells(fila, 2).name = "tag_" & lab(i + 1) ' celda de datos
            fila = fila + 1
        Next i
        ws.UsedRange.Borders.LineStyle = xlNone
        ws.Range("A3:B" & fila - 1).Borders(xlEdgeLeft).LineStyle = xlContinuous
        ws.Range("A3:B" & fila - 1).Borders(xlEdgeRight).LineStyle = xlContinuous
        ws.Range("A3:B" & fila - 1).Borders(xlEdgeTop).LineStyle = xlContinuous
        ws.Range("A3:B" & fila - 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
        ws.Range("A3:A" & fila - 1).Interior.Color = RGB(245, 245, 245)
    End If

    Set EnsurePdfTemplate = ws
End Function

'--------------------------------------------------------
' Llena la plantilla con los datos de una fila (F/NC/ND o Ret)
' y exporta a PDF en la misma carpeta del XML.
' Inserta hipervínculo en la celda nro_comprobante.
'--------------------------------------------------------
'--------------------------------------------------------
' Llena la plantilla con los datos de una fila (F/NC/ND o Ret)
' Exporta a PDF en la misma carpeta del XML y escribe LOG.
' Inserta hipervínculo en la celda nro_comprobante.
'--------------------------------------------------------
Private Sub GenerarPDFs_en_Hoja(ByVal wsSrc As Worksheet, ByVal wsT As Worksheet, _
                                ByVal root As String, ByVal esFactura As Boolean)

    Dim wsL As Worksheet
    Set wsL = EnsureSheet("LOG")   ' usaremos tu hoja LOG y tu Sub Log(...)

    Dim cNro As Long, cFec As Long, cTipo As Long, cRucE As Long
    Dim cRzE As Long, cEst As Long, cPto As Long, cSec As Long
    Dim cRucC As Long, cRzC As Long, cTotal As Long

    cNro = ColByHeader(wsSrc, "nro_comprobante")
    cFec = ColByHeader(wsSrc, "fecha_emision")
    If esFactura Then
        cTipo = ColByHeader(wsSrc, "tipo_comprobante")       ' 01/04/05
        cRucE = ColByHeader(wsSrc, "ruc_emisor")
        cRzE = ColByHeader(wsSrc, "razon_social_emisor")
        cEst = ColByHeader(wsSrc, "establecimiento")
        cPto = ColByHeader(wsSrc, "punto_emision")
        cSec = ColByHeader(wsSrc, "secuencial")
        cRucC = ColByHeader(wsSrc, "ruc_ci_comprador")
        cRzC = ColByHeader(wsSrc, "razon_social_comprador")
        cTotal = ColByHeader(wsSrc, "valor_total")
    Else
        ' Retenciones
        cTipo = 0
        cRucE = ColByHeader(wsSrc, "ruc_emisor")
        cRzE = ColByHeader(wsSrc, "razon_social_emisor")
        cEst = ColByHeader(wsSrc, "establecimiento")
        cPto = ColByHeader(wsSrc, "punto_emision")
        cSec = ColByHeader(wsSrc, "secuencial")
        cRucC = ColByHeader(wsSrc, "ruc_sujeto")
        cRzC = ColByHeader(wsSrc, "razon_social_sujeto")
        cTotal = ColByHeader(wsSrc, "total_retenido")
    End If

    If cNro = 0 Or cFec = 0 Or cRucE = 0 Or cRzE = 0 Or cSec = 0 Then Exit Sub

    Dim lastRow As Long, r As Long
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    For r = 2 To lastRow
        Dim nro As String: nro = Trim$(CStr(wsSrc.Cells(r, cNro).Value))
        If Len(nro) = 0 Then GoTo NextR

        Dim fecTxt As String: fecTxt = Trim$(CStr(wsSrc.Cells(r, cFec).Value))
        Dim tipoCod As String
        If esFactura And cTipo > 0 Then tipoCod = Trim$(CStr(wsSrc.Cells(r, cTipo).Value)) Else tipoCod = ""

        ' 1) Intentar localizar el XML
        Dim xmlPath As String
        xmlPath = FindXmlForRow(nro, fecTxt, tipoCod, root, Not esFactura)

        If Len(xmlPath) = 0 Then
            Log wsL, wsSrc.name & "!" & nro, "NO_XML", "No se halló XML bajo: " & root
            GoTo NextR
        End If

        ' 2) Llenar plantilla
        wsT.Range("tag_tipo").Value = IIf(esFactura, TipoTextoFactura(tipoCod), "Comprobante de Retención")
        wsT.Range("tag_nro").Value = nro
        wsT.Range("tag_fec").Value = SafeDateStr(fecTxt)
        wsT.Range("tag_rucE").Value = CStr(wsSrc.Cells(r, cRucE).Value)
        wsT.Range("tag_rzE").Value = CStr(wsSrc.Cells(r, cRzE).Value)
        wsT.Range("tag_est").Value = CStr(wsSrc.Cells(r, cEst).Value)
        wsT.Range("tag_pto").Value = CStr(wsSrc.Cells(r, cPto).Value)
        wsT.Range("tag_sec").Value = CStr(wsSrc.Cells(r, cSec).Value)
        wsT.Range("tag_cli").Value = CStr(wsSrc.Cells(r, cRucC).Value)
        wsT.Range("tag_rzC").Value = CStr(wsSrc.Cells(r, cRzC).Value)
        wsT.Range("tag_tot").Value = CStr(wsSrc.Cells(r, cTotal).Value)

        ' 3) Exportar PDF junto al XML
        Dim outPdf As String
        'outPdf = BuildPdfPath(xmlPath, tipoCod, nro, fecTxt, Not esFactura)
        outPdf = EX_BuildOutputPath(xmlPath, tipoCod, nro, fecTxt, Not esFactura)


        On Error Resume Next
        wsT.ExportAsFixedFormat Type:=xlTypePDF, fileName:=outPdf, _
            Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        If Err.Number = 0 And Len(dir$(outPdf)) > 0 Then
            Log wsL, xmlPath, "PDF", "Creado -> " & outPdf
        Else
            Log wsL, xmlPath, "ERROR_PDF", "No se pudo crear -> " & outPdf
        End If
        On Error GoTo 0

        ' 4) Hipervínculo en nro_comprobante
        On Error Resume Next
        wsSrc.Hyperlinks.Add Anchor:=wsSrc.Cells(r, cNro), Address:=outPdf, _
            SubAddress:="", TextToDisplay:=CStr(wsSrc.Cells(r, cNro).Value)
        If Err.Number = 0 Then
            Log wsL, xmlPath, "LINK", "Hipervínculo agregado en " & wsSrc.name & "!" & wsSrc.Cells(r, cNro).Address(0, 0)
        Else
            Log wsL, xmlPath, "ERROR_LINK", "No se pudo crear hipervínculo en " & wsSrc.name & "!" & wsSrc.Cells(r, cNro).Address(0, 0)
        End If
        On Error GoTo 0

NextR:
    Next r
End Sub


'------------------ Utilidades de ruta/nombre -------------------

' Construye:  FC- MMDDYYYY -EEE-PPP-SSSSSSSSS.pdf   (o CR-... para retenciones)
Private Function BuildPdfPath(ByVal xmlPath As String, ByVal tipoCod As String, ByVal nro As String, _
                              ByVal fecTxt As String, ByVal esRet As Boolean) As String
    Dim folder As String, pref As String, base As String

    folder = Left$(xmlPath, InStrRev(xmlPath, "\") - 1)

    If esRet Then
        pref = "CR"
    Else
        Select Case Trim$(tipoCod)
            Case "01": pref = "FC"
            Case "04": pref = "NC"
            Case "05": pref = "ND"
            Case Else: pref = "CB"
        End Select
    End If

    base = pref & "-" & FechaToken_MDY(fecTxt) & "-" & Replace$(nro, " ", "")

    ' Limpieza por si quedaran dobles guiones o un guion al final
    Do While InStr(base, "--") > 0: base = Replace(base, "--", "-"): Loop
    If Right$(base, 1) = "-" Then base = Left$(base, Len(base) - 1)

    BuildPdfPath = folder & "\" & SanitizeFileName(base) & ".pdf"
End Function


Private Function SanitizeFileName(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, "\", "_")
    t = Replace(t, "/", "_")
    t = Replace(t, ":", "_")
    t = Replace(t, "*", "_")
    t = Replace(t, "?", "_")
    t = Replace(t, """", "_")
    t = Replace(t, "<", "_")
    t = Replace(t, ">", "_")
    t = Replace(t, "|", "_")
    SanitizeFileName = t
End Function

Private Function TipoTextoFactura(ByVal cod As String) As String
    Select Case Trim$(cod)
        Case "01": TipoTextoFactura = "Factura"
        Case "04": TipoTextoFactura = "Nota de Crédito"
        Case "05": TipoTextoFactura = "Nota de Débito"
        Case Else:  TipoTextoFactura = "Comprobante"
    End Select
End Function

' Busca el XML por nombre (en raíz/subcarpetas)
' Heurísticas: nro con/ sin guiones, prefijo FC/NC/ND/CR, fecha.
Private Function FindXmlForRow(ByVal nro As String, ByVal fecTxt As String, ByVal tipoCod As String, _
                               ByVal root As String, ByVal esRet As Boolean) As String
    Dim pref As String
    If esRet Then
        pref = "CR"
    Else
        Select Case Trim$(tipoCod)
            Case "01": pref = "FC"
            Case "04": pref = "NC"
            Case "05": pref = "ND"
            Case Else:  pref = ""   ' por si acaso
        End Select
    End If

    Dim key1 As String, key2 As String, kf1 As String, kf2 As String
    key1 = LCase$(Replace$(nro, "-", ""))
    key2 = LCase$(nro)
    kf1 = FechaToken(fecTxt, True)   ' yyyymmdd
    kf2 = FechaToken(fecTxt, False)  ' ddmmyyyy

    FindXmlForRow = SearchXmlByKeys(root, pref, key1, key2, kf1, kf2)
End Function

Private Function SearchXmlByKeys(ByVal root As String, ByVal pref As String, _
                                 ByVal key1 As String, ByVal key2 As String, _
                                 ByVal kf1 As String, ByVal kf2 As String) As String
    Dim col As New Collection, st As New Collection
    Dim p As String, f As String, subf As String

    st.Add root
    Do While st.Count > 0
        p = st(1): st.Remove 1

        f = dir(p & "\*.xml")
        Do While Len(f) > 0
            Dim low As String: low = LCase$(f)
            Dim ok As Boolean: ok = False

            ' Puntuación mínima por coincidencias simples con el nombre
            Dim score As Long: score = 0
            If InStr(1, low, key1) > 0 Then score = score + 3
            If InStr(1, low, key2) > 0 Then score = score + 2
            If Len(pref) > 0 And InStr(1, low, LCase$(pref & "-")) > 0 Then score = score + 2
            If Len(kf1) > 0 And InStr(1, low, kf1) > 0 Then score = score + 1
            If Len(kf2) > 0 And InStr(1, low, kf2) > 0 Then score = score + 1

            If score >= 3 Then
                SearchXmlByKeys = p & "\" & f
                Exit Function
            End If

            f = dir
        Loop

        subf = dir(p & "\*", vbDirectory)
        Do While Len(subf) > 0
            If subf <> "." And subf <> ".." Then
                If (GetAttr(p & "\" & subf) And vbDirectory) <> 0 Then
                    st.Add p & "\" & subf
                End If
            End If
            subf = dir
        Loop
    Loop

    SearchXmlByKeys = ""
End Function

'------------------ Formatos de fecha -------------------

Private Function SafeDateStr(ByVal s As String) As String
    Dim d As Date
    If TryParseFechaEmision(s, d) Then
        SafeDateStr = Format$(d, "yyyy-mm-dd")
    Else
        SafeDateStr = s
    End If
End Function

Private Function FechaToken(ByVal s As String, ByVal ymd As Boolean) As String
    Dim d As Date
    If TryParseFechaEmision(s, d) Then
        If ymd Then FechaToken = Format$(d, "yyyymmdd") Else FechaToken = Format$(d, "ddmmyyyy")
    Else
        FechaToken = ""
    End If
End Function

' Limpia el LOG_Rename (conserva encabezados). Crea la hoja si no existe.
Public Sub LimpiarLOGRename()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("LOG_Rename")
    On Error GoTo 0

    If ws Is Nothing Then
        ' Si no existe, créala vacía con encabezados
        Set ws = ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.Count))
        ws.name = "LOG_Rename"
        ws.Range("A1:D1").Value = Array("Archivo_Origen", "Archivo_Destino", "Estado", "Mensaje")
        ws.Range("A1:D1").Font.Bold = True
        ws.Columns("A:D").ColumnWidth = 70
        Exit Sub
    End If

    With ws
        If .AutoFilterMode Then .AutoFilterMode = False
        ' Borra todo menos la fila de encabezados
        If .UsedRange.Rows.Count > 1 Then
            .Rows("2:" & .Rows.Count).ClearContents
        End If
    End With
End Sub

'************************************************************************************************************************************
'RENOMBRAR DOCUMENTOS
' Devuelve MMDDYYYY a partir de "DD/MM/YYYY", "DD-MM-YYYY" o "YYYY-MM-DD"
Private Function FechaToken_MDY(ByVal s As String) As String
    Dim d As Date
    If TryParseFechaEmision(s, d) Then
        FechaToken_MDY = Format$(d, "mmddyyyy")
    Else
        FechaToken_MDY = ""
    End If
End Function


