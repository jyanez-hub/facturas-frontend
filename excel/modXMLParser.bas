' Attribute VB_Name = "modXMLParser"
Option Explicit

'##########################################################################################
'# Módulo: modXMLParser
'# Propósito: Convertir los XML de comprobantes electrónicos del SRI en estructuras
'#            basadas en diccionarios y colecciones listas para poblar las tablas del
'#            libro configurado por modSetup.
'# Alcance:   Facturas (codDoc 01), Notas de crédito (codDoc 04) y comprobantes de
'#            retención (codDoc 07). Todo el código usa enlace tardío para maximizar la
'#            compatibilidad con Excel 2007 o superior.
'##########################################################################################

'------------------------------------------------------------------------------------------
' Tipos devueltos por el parser
'------------------------------------------------------------------------------------------
Public Type ParsedComprobante
    Tipo As String             ' Código de documento SRI (01, 04, 07)
    Cabecera As Object         ' Dictionary con campos de la tabla principal
    Detalles As Object         ' Collection de dictionaries por cada línea
    Adicional As Object        ' Dictionary con campos adicionales (infoAdicional)
    Nodo As Object             ' Referencia al nodo raíz original (DOM)
End Type

'------------------------------------------------------------------------------------------
' API pública
'------------------------------------------------------------------------------------------
Public Function ParseComprobanteXML(ByVal xmlSource As Variant, _
                                    Optional ByVal origenArchivo As String = "XML", _
                                    Optional rutaArchivo As String = "", _
                                    Optional nombreArchivo As String = "") As ParsedComprobante
    Dim dom As Object
    Dim root As Object
    Dim tipo As String
    Dim result As ParsedComprobante

    Set dom = ResolveDom(xmlSource, rutaArchivo, nombreArchivo)
    If dom Is Nothing Then Exit Function

    Set root = ResolveComprobanteRoot(dom)
    If root Is Nothing Then
        Err.Raise vbObjectError + 519, "modXMLParser.ParseComprobanteXML", _
                  "No se encontró un comprobante válido dentro del XML."
    End If

    tipo = SelectSingleNodeText(root, "infoTributaria/codDoc")
    tipo = Format$(Val(SoloDigitos(tipo)), "00")

    Select Case tipo
        Case "01"
            result = ParseFacturaDom(root, origenArchivo, rutaArchivo, nombreArchivo)
        Case "04"
            result = ParseNotaCreditoDom(root, origenArchivo, rutaArchivo, nombreArchivo)
        Case "07"
            result = ParseRetencionDom(root, origenArchivo, rutaArchivo, nombreArchivo)
        Case Else
            Err.Raise vbObjectError + 514, "modXMLParser.ParseComprobanteXML", _
                      "Tipo de comprobante no soportado: " & tipo
    End Select

    Set result.Nodo = root
    ParseComprobanteXML = result
End Function

Public Function ParseFacturaXML(ByVal xmlSource As Variant, _
                                Optional ByVal origenArchivo As String = "XML", _
                                Optional rutaArchivo As String = "", _
                                Optional nombreArchivo As String = "") As ParsedComprobante
    ParseFacturaXML = ParseComprobanteByTipo(xmlSource, "01", origenArchivo, rutaArchivo, nombreArchivo)
End Function

Public Function ParseNotaCreditoXML(ByVal xmlSource As Variant, _
                                    Optional ByVal origenArchivo As String = "XML", _
                                    Optional rutaArchivo As String = "", _
                                    Optional nombreArchivo As String = "") As ParsedComprobante
    ParseNotaCreditoXML = ParseComprobanteByTipo(xmlSource, "04", origenArchivo, rutaArchivo, nombreArchivo)
End Function

Public Function ParseRetencionXML(ByVal xmlSource As Variant, _
                                  Optional ByVal origenArchivo As String = "XML", _
                                  Optional rutaArchivo As String = "", _
                                  Optional nombreArchivo As String = "") As ParsedComprobante
    ParseRetencionXML = ParseComprobanteByTipo(xmlSource, "07", origenArchivo, rutaArchivo, nombreArchivo)
End Function

Public Sub Run_SmokeTest()
    Dim facturaXml As String
    Dim retencionXml As String
    Dim parsed As ParsedComprobante
    Dim linea As Object

    facturaXml = _
        "<factura>" & _
        "<infoTributaria>" & _
        "<codDoc>01</codDoc><ruc>1790012345001</ruc><razonSocial>Demo S.A.</razonSocial>" & _
        "<estab>001</estab><ptoEmi>002</ptoEmi><secuencial>000012345</secuencial>" & _
        "<claveAcceso>1234567890123456789012345678901234567890123456789</claveAcceso>" & _
        "</infoTributaria>" & _
        "<infoFactura>" & _
        "<fechaEmision>01/05/2024</fechaEmision><moneda>USD</moneda>" & _
        "<tipoIdentificacionComprador>05</tipoIdentificacionComprador>" & _
        "<razonSocialComprador>Cliente Prueba</razonSocialComprador>" & _
        "<identificacionComprador>0912345678</identificacionComprador>" & _
        "<totalSinImpuestos>100.00</totalSinImpuestos><propina>0.00</propina>" & _
        "<importeTotal>112.00</importeTotal>" & _
        "<totalConImpuestos><totalImpuesto><codigo>2</codigo><codigoPorcentaje>2</codigoPorcentaje>" & _
        "<baseImponible>100.00</baseImponible><tarifa>12</tarifa><valor>12.00</valor></totalImpuesto></totalConImpuestos>" & _
        "</infoFactura>" & _
        "<detalles><detalle><codigoPrincipal>A1</codigoPrincipal><descripcion>Producto demo</descripcion>" & _
        "<cantidad>1</cantidad><precioUnitario>100</precioUnitario><descuento>0</descuento>" & _
        "<precioTotalSinImpuesto>100</precioTotalSinImpuesto><impuestos><impuesto><codigo>2</codigo>" & _
        "<codigoPorcentaje>2</codigoPorcentaje><tarifa>12</tarifa><baseImponible>100</baseImponible>" & _
        "<valor>12</valor></impuesto></impuestos></detalle></detalles>" & _
        "<infoAdicional><campoAdicional nombre='Email'>demo@correo.com</campoAdicional></infoAdicional>" & _
        "</factura>"

    retencionXml = _
        "<comprobanteRetencion>" & _
        "<infoTributaria><codDoc>07</codDoc><ruc>0999999999001</ruc><razonSocial>Agente Ret</razonSocial>" & _
        "<estab>001</estab><ptoEmi>001</ptoEmi><secuencial>000000123</secuencial>" & _
        "<claveAcceso>9876543210987654321098765432109876543210987654321</claveAcceso></infoTributaria>" & _
        "<infoCompRetencion><fechaEmision>15/05/2024</fechaEmision><tipoIdentificacionSujetoRetenido>05</tipoIdentificacionSujetoRetenido>" & _
        "<razonSocialSujetoRetenido>Proveedor Demo</razonSocialSujetoRetenido>" & _
        "<identificacionSujetoRetenido>1712345678</identificacionSujetoRetenido>" & _
        "<periodoFiscal>05/2024</periodoFiscal></infoCompRetencion>" & _
        "<impuestos><impuesto><codigo>1</codigo><codigoRetencion>332</codigoRetencion><baseImponible>100</baseImponible>" & _
        "<porcentajeRetener>1</porcentajeRetener><valorRetenido>1</valorRetenido><codDocSustento>01</codDocSustento>" & _
        "<numDocSustento>001-001-000012345</numDocSustento><fechaEmisionDocSustento>01/05/2024</fechaEmisionDocSustento></impuesto>" & _
        "</impuestos></comprobanteRetencion>"

    parsed = ParseFacturaXML(facturaXml)
    Debug.Print "Factura -> " & parsed.Cabecera("NumeroDocumento") & _
                " Total: " & parsed.Cabecera("ValorTotal")
    For Each linea In parsed.Detalles
        Debug.Print "  Detalle: " & linea("Descripcion") & " IVA:" & linea("ValorIVA")
    Next linea

    parsed = ParseRetencionXML(retencionXml)
    Debug.Print "Retención -> " & parsed.Cabecera("NumeroDocumento") & _
                " Valor IVA retenido:" & parsed.Cabecera("ValorRetenidoIVA")
End Sub

'------------------------------------------------------------------------------------------
' Implementación interna
'------------------------------------------------------------------------------------------
Private Function ParseComprobanteByTipo(ByVal xmlSource As Variant, _
                                        ByVal tipoEsperado As String, _
                                        ByVal origenArchivo As String, _
                                        Optional rutaArchivo As String = "", _
                                        Optional nombreArchivo As String = "") As ParsedComprobante
    Dim dom As Object
    Dim root As Object
    Dim tipo As String
    Dim result As ParsedComprobante

    Set dom = ResolveDom(xmlSource, rutaArchivo, nombreArchivo)
    If dom Is Nothing Then Exit Function

    Set root = ResolveComprobanteRoot(dom)
    If root Is Nothing Then
        Err.Raise vbObjectError + 519, "modXMLParser.ParseComprobanteByTipo", _
                  "No se encontró un comprobante válido dentro del XML."
    End If

    tipo = SelectSingleNodeText(root, "infoTributaria/codDoc")
    tipo = Format$(Val(SoloDigitos(tipo)), "00")

    If tipo <> tipoEsperado Then
        Err.Raise vbObjectError + 515, "modXMLParser.ParseComprobanteByTipo", _
                  "El XML recibido corresponde al tipo " & tipo & _
                  ", se esperaba " & tipoEsperado
    End If

    Select Case tipoEsperado
        Case "01"
            result = ParseFacturaDom(root, origenArchivo, rutaArchivo, nombreArchivo)
        Case "04"
            result = ParseNotaCreditoDom(root, origenArchivo, rutaArchivo, nombreArchivo)
        Case "07"
            result = ParseRetencionDom(root, origenArchivo, rutaArchivo, nombreArchivo)
    End Select

    Set result.Nodo = root
    ParseComprobanteByTipo = result
End Function

Private Function ResolveDom(ByVal xmlSource As Variant, _
                            ByRef rutaArchivo As String, _
                            ByRef nombreArchivo As String) As Object
    Dim dom As Object
    Dim sourceText As String
    Dim wasPath As Boolean

    If IsObject(xmlSource) Then
        If HasProperty(xmlSource, "DocumentElement") Then
            Set dom = xmlSource
        ElseIf HasProperty(xmlSource, "ownerDocument") Then
            Set dom = xmlSource.ownerDocument
        End If
        If dom Is Nothing Then
            Err.Raise vbObjectError + 516, "modXMLParser.ResolveDom", _
                      "El objeto proporcionado no es un DOM MSXML válido."
        End If
    Else
        sourceText = CStr(xmlSource)
        wasPath = False
        If rutaArchivo = "" Then
            If IsExistingFile(sourceText) Then
                rutaArchivo = sourceText
                wasPath = True
            End If
        Else
            If nombreArchivo = "" Then
                nombreArchivo = ExtractFileName(rutaArchivo)
            End If
        End If

        Set dom = CreateDomDocument()

        If wasPath Then
            If nombreArchivo = "" Then nombreArchivo = ExtractFileName(rutaArchivo)
            On Error Resume Next
            dom.Load rutaArchivo
            If dom.parseError.errorCode <> 0 Then
                Err.Raise vbObjectError + 517, "modXMLParser.ResolveDom", dom.parseError.reason
            End If
            On Error GoTo 0
        Else
            On Error Resume Next
            dom.LoadXML sourceText
            If dom.parseError.errorCode <> 0 Then
                Err.Raise vbObjectError + 518, "modXMLParser.ResolveDom", dom.parseError.reason
            End If
            On Error GoTo 0
        End If
    End If

    Set ResolveDom = dom
End Function

Private Function ResolveComprobanteRoot(ByVal dom As Object) As Object
    Dim baseNode As Object

    If dom Is Nothing Then Exit Function

    Set baseNode = dom.DocumentElement
    If baseNode Is Nothing Then Exit Function

    Set ResolveComprobanteRoot = NormalizeComprobanteNode(baseNode)
End Function

Private Function NormalizeComprobanteNode(ByVal node As Object) As Object
    Dim nodeName As String
    Dim innerXml As String
    Dim idx As Long
    Dim child As Object
    Dim candidate As Object

    If node Is Nothing Then Exit Function

    nodeName = NodeLocalName(node)

    Select Case nodeName
        Case "factura", "notacredito", "comprobanteretencion"
            Set NormalizeComprobanteNode = node
            Exit Function
        Case "comprobante"
            innerXml = ExtractComprobanteInnerXml(node)
            If Len(innerXml) > 0 Then
                Set NormalizeComprobanteNode = LoadInnerComprobante(innerXml)
            End If
            Exit Function
        Case "autorizacion", "autorizaciones", "respuestaautorizacion", "respuestaautorizacioncomprobante"
            innerXml = ExtractComprobanteInnerXml(node)
            If Len(innerXml) > 0 Then
                Set NormalizeComprobanteNode = LoadInnerComprobante(innerXml)
                Exit Function
            End If
    End Select

    If InStr(nodeName, "autorizacion") > 0 Or InStr(nodeName, "respuesta") > 0 Then
        innerXml = ExtractComprobanteInnerXml(node)
        If Len(innerXml) > 0 Then
            Set NormalizeComprobanteNode = LoadInnerComprobante(innerXml)
            Exit Function
        End If
    End If

    If node.childNodes Is Nothing Then Exit Function

    For idx = 0 To node.childNodes.length - 1
        Set child = node.childNodes.Item(idx)
        If child.nodeType = 1 Then
            Set candidate = NormalizeComprobanteNode(child)
            If Not candidate Is Nothing Then
                Set NormalizeComprobanteNode = candidate
                Exit Function
            End If
        End If
    Next idx
End Function

Private Function ExtractComprobanteInnerXml(ByVal node As Object) As String
    Dim idx As Long
    Dim child As Object
    Dim texto As String

    If node Is Nothing Then Exit Function

    If NodeLocalName(node) = "comprobante" Then
        texto = Trim$("" & node.Text)
        If Len(texto) = 0 Then
            On Error Resume Next
            texto = Trim$("" & node.nodeTypedValue)
            On Error GoTo 0
        End If
        ExtractComprobanteInnerXml = texto
        Exit Function
    End If

    If node.childNodes Is Nothing Then Exit Function

    For idx = 0 To node.childNodes.length - 1
        Set child = node.childNodes.Item(idx)
        If child.nodeType = 1 Then
            If NodeLocalName(child) = "comprobante" Then
                texto = Trim$("" & child.Text)
                If Len(texto) = 0 Then
                    On Error Resume Next
                    texto = Trim$("" & child.nodeTypedValue)
                    On Error GoTo 0
                End If
                ExtractComprobanteInnerXml = texto
                Exit Function
            End If
        End If
    Next idx

    For idx = 0 To node.childNodes.length - 1
        Set child = node.childNodes.Item(idx)
        If child.nodeType = 1 Then
            texto = ExtractComprobanteInnerXml(child)
            If Len(texto) > 0 Then
                ExtractComprobanteInnerXml = texto
                Exit Function
            End If
        End If
    Next idx
End Function

Private Function LoadInnerComprobante(ByVal innerXml As String) As Object
    Dim doc As Object

    innerXml = Trim$(innerXml)
    If Len(innerXml) = 0 Then Exit Function

    Set doc = CreateDomDocument()

    On Error Resume Next
    doc.LoadXML innerXml
    If doc.parseError.errorCode <> 0 Then
        Err.Raise vbObjectError + 520, "modXMLParser.LoadInnerComprobante", doc.parseError.reason
    Else
        Set LoadInnerComprobante = doc.DocumentElement
    End If
    On Error GoTo 0
End Function

Private Function NodeLocalName(ByVal node As Object) As String
    Dim raw As String
    Dim pos As Long

    If node Is Nothing Then Exit Function

    On Error Resume Next
    NodeLocalName = LCase$(Trim$(CStr(node.baseName)))
    On Error GoTo 0

    If Len(NodeLocalName) = 0 Then
        raw = Trim$(CStr(node.nodeName))
        pos = InStrRev(raw, ":")
        If pos > 0 Then
            raw = Mid$(raw, pos + 1)
        End If
        NodeLocalName = LCase$(raw)
    End If
End Function

Private Function ParseFacturaDom(ByVal root As Object, _
                                 ByVal origenArchivo As String, _
                                 ByVal rutaArchivo As String, _
                                 ByVal nombreArchivo As String) As ParsedComprobante
    Dim header As Object
    Dim detalles As Collection
    Dim infoTrib As Object
    Dim infoFactura As Object
    Dim numeroDocumento As String
    Dim fechaEmision As Variant
    Dim nodoImpuestos As Object
    Dim impuesto As Object
    Dim idx As Long
    Dim lineNode As Object
    Dim result As ParsedComprobante
    Dim adicionales As Object

    Set header = InitializeFacturaHeader(origenArchivo, rutaArchivo, nombreArchivo)
    Set detalles = New Collection
    Set adicionales = CollectAdditionalFields(root.selectNodes("infoAdicional/campoAdicional"))

    Set infoTrib = root.selectSingleNode("infoTributaria")
    Set infoFactura = root.selectSingleNode("infoFactura")

    header("TipoComprobante") = "01"
    header("Subtipo") = SelectSingleNodeText(root, "infoFactura/identificacionDelIntermediario")
    header("RUC_Emisor") = SelectSingleNodeText(infoTrib, "ruc")
    header("TipoIdentificacionEmisor") = DefaultText(SelectSingleNodeText(infoTrib, "tipoIdentificacion"), "RUC")
    header("Nombre_Emisor") = DefaultText(SelectSingleNodeText(infoTrib, "razonSocial"), _
                                           SelectSingleNodeText(infoTrib, "nombreComercial"))
    header("RUC_Receptor") = SelectSingleNodeText(infoFactura, "identificacionComprador")
    header("TipoIdentificacionReceptor") = SelectSingleNodeText(infoFactura, "tipoIdentificacionComprador")
    header("Nombre_Receptor") = SelectSingleNodeText(infoFactura, "razonSocialComprador")
    header("Establecimiento") = SelectSingleNodeText(infoTrib, "estab")
    header("PuntoEmision") = SelectSingleNodeText(infoTrib, "ptoEmi")
    header("Secuencial") = SelectSingleNodeText(infoTrib, "secuencial")
    header("ClaveAcceso") = SelectSingleNodeText(infoTrib, "claveAcceso")
    header("Moneda") = DefaultText(SelectSingleNodeText(infoFactura, "moneda"), "USD")

    numeroDocumento = BuildNumeroDocumento(header("Establecimiento"), header("PuntoEmision"), header("Secuencial"))
    header("NumeroDocumento") = numeroDocumento

    fechaEmision = ParseFechaEmision(SelectSingleNodeText(infoFactura, "fechaEmision"))
    If IsDate(fechaEmision) Then
        header("FechaEmision") = fechaEmision
    End If

    header("TotalSinImpuestos") = ParseSRINumber(SelectSingleNodeText(infoFactura, "totalSinImpuestos"))
    header("Propina") = ParseSRINumber(SelectSingleNodeText(infoFactura, "propina"))
    header("ValorTotal") = ParseSRINumber(SelectSingleNodeText(infoFactura, "importeTotal"))

    Set nodoImpuestos = infoFactura.selectNodes("totalConImpuestos/totalImpuesto")
    If Not nodoImpuestos Is Nothing Then
        For idx = 0 To nodoImpuestos.length - 1
            Set impuesto = nodoImpuestos.Item(idx)
            AggregateImpuestoFactura header, impuesto
        Next idx
    End If

    Set nodoImpuestos = root.selectNodes("detalles/detalle")
    If Not nodoImpuestos Is Nothing Then
        For idx = 0 To nodoImpuestos.length - 1
            Set lineNode = nodoImpuestos.Item(idx)
            detalles.Add ParseFacturaDetalle(lineNode, numeroDocumento, idx + 1, header)
        Next idx
    End If

    header("Observaciones") = JoinAdditionalText(root.selectNodes("infoAdicional/campoAdicional"))

    result.Tipo = "01"
    Set result.Cabecera = header
    Set result.Detalles = detalles
    Set result.Adicional = adicionales
    ParseFacturaDom = result
End Function

Private Function ParseFacturaDetalle(ByVal lineNode As Object, _
                                     ByVal numeroDocumento As String, _
                                     ByVal lineaIndex As Long, _
                                     ByVal header As Object) As Object
    Dim detail As Object
    Dim impuestos As Object
    Dim impuesto As Object
    Dim idx As Long
    Dim codigo As String
    Dim codigoPorcentaje As String
    Dim tarifa As Double
    Dim baseImponible As Double
    Dim valor As Double

    Set detail = NewDictionary()

    detail("NumeroDocumento") = numeroDocumento
    detail("Linea") = Format$(lineaIndex, "0000")
    detail("CodigoPrincipal") = SelectSingleNodeText(lineNode, "codigoPrincipal")
    detail("CodigoAuxiliar") = SelectSingleNodeText(lineNode, "codigoAuxiliar")
    detail("Descripcion") = SelectSingleNodeText(lineNode, "descripcion")
    detail("Cantidad") = ParseSRINumber(SelectSingleNodeText(lineNode, "cantidad"))
    detail("UnidadMedida") = SelectSingleNodeText(lineNode, "unidadMedida")
    detail("PrecioUnitario") = ParseSRINumber(SelectSingleNodeText(lineNode, "precioUnitario"))
    detail("Descuento") = ParseSRINumber(SelectSingleNodeText(lineNode, "descuento"))
    detail("PrecioTotalSinImpuesto") = ParseSRINumber(SelectSingleNodeText(lineNode, "precioTotalSinImpuesto"))
    detail("TarifaIVA") = 0
    detail("PorcentajeIVA") = 0
    detail("BaseIVA") = 0
    detail("ValorIVA") = 0
    detail("TarifaICE") = 0
    detail("ValorICE") = 0
    detail("DetalleAdicional1") = ""
    detail("DetalleAdicional2") = ""
    detail("DetalleAdicional3") = ""

    Set impuestos = lineNode.selectNodes("impuestos/impuesto")
    If Not impuestos Is Nothing Then
        For idx = 0 To impuestos.length - 1
            Set impuesto = impuestos.Item(idx)
            codigo = SelectSingleNodeText(impuesto, "codigo")
            codigoPorcentaje = SelectSingleNodeText(impuesto, "codigoPorcentaje")
            tarifa = ParseSRINumber(SelectSingleNodeText(impuesto, "tarifa"))
            baseImponible = ParseSRINumber(SelectSingleNodeText(impuesto, "baseImponible"))
            valor = ParseSRINumber(SelectSingleNodeText(impuesto, "valor"))

            Select Case codigo
                Case "2"
                    detail("TarifaIVA") = NormalizePercent(tarifa)
                    detail("PorcentajeIVA") = NormalizePercent(tarifa)
                    detail("BaseIVA") = baseImponible
                    detail("ValorIVA") = valor
                Case "3"
                    detail("TarifaICE") = NormalizePercent(tarifa)
                    detail("ValorICE") = valor
            End Select
        Next idx
    End If

    AssignDetalleAdicionales lineNode, detail
    Set ParseFacturaDetalle = detail
End Function

Private Sub AggregateImpuestoFactura(ByVal header As Object, ByVal impuesto As Object)
    Dim codigo As String
    Dim codigoPorcentaje As String
    Dim tarifa As Double
    Dim baseImponible As Double
    Dim valor As Double

    codigo = SelectSingleNodeText(impuesto, "codigo")
    codigoPorcentaje = SelectSingleNodeText(impuesto, "codigoPorcentaje")
    tarifa = ParseSRINumber(SelectSingleNodeText(impuesto, "tarifa"))
    baseImponible = ParseSRINumber(SelectSingleNodeText(impuesto, "baseImponible"))
    valor = ParseSRINumber(SelectSingleNodeText(impuesto, "valor"))

    Select Case codigo
        Case "2"
            DistributeIVA header, codigoPorcentaje, tarifa, baseImponible, valor
        Case "3"
            header("BaseICE") = header("BaseICE") + baseImponible
            header("ValorICE") = header("ValorICE") + valor
        Case "5"
            header("ValorIRBPNR") = header("ValorIRBPNR") + valor
    End Select
End Sub

Private Sub DistributeIVA(ByVal header As Object, _
                          ByVal codigoPorcentaje As String, _
                          ByVal tarifa As Double, _
                          ByVal baseImponible As Double, _
                          ByVal valor As Double)
    Dim bucket As String

    Select Case codigoPorcentaje
        Case "6"
            bucket = "NoObjeto"
        Case "7"
            bucket = "Exento"
        Case Else
            bucket = ResolveIVABucket(tarifa, codigoPorcentaje)
    End Select

    Select Case bucket
        Case "IVA15"
            header("BaseIVA15") = header("BaseIVA15") + baseImponible
            header("ValorIVA15") = header("ValorIVA15") + valor
        Case "IVA12"
            header("BaseIVA12") = header("BaseIVA12") + baseImponible
            header("ValorIVA12") = header("ValorIVA12") + valor
        Case "IVA8"
            header("BaseIVA8") = header("BaseIVA8") + baseImponible
            header("ValorIVA8") = header("ValorIVA8") + valor
        Case "IVA5"
            header("BaseIVA5") = header("BaseIVA5") + baseImponible
            header("ValorIVA5") = header("ValorIVA5") + valor
        Case "IVA0"
            header("BaseIVA0") = header("BaseIVA0") + baseImponible
            header("ValorIVA0") = header("ValorIVA0") + valor
        Case "NoObjeto"
            header("BaseNoObjetoIVA") = header("BaseNoObjetoIVA") + baseImponible
        Case "Exento"
            header("BaseExentaIVA") = header("BaseExentaIVA") + baseImponible
        Case Else
            header("BaseIVA0") = header("BaseIVA0") + baseImponible
            header("ValorIVA0") = header("ValorIVA0") + valor
    End Select
End Sub

Private Function ResolveIVABucket(ByVal tarifa As Double, ByVal codigoPorcentaje As String) As String
    Dim normalized As Double

    If tarifa = 0 Then
        Select Case codigoPorcentaje
            Case "0", "3"
                ResolveIVABucket = "IVA0"
                Exit Function
        End Select
    End If

    normalized = Round(tarifa, 2)

    Select Case True
        Case normalized >= 14.5
            ResolveIVABucket = "IVA15"
        Case normalized >= 11 And normalized < 13.5
            ResolveIVABucket = "IVA12"
        Case normalized >= 7.5 And normalized < 8.5
            ResolveIVABucket = "IVA8"
        Case normalized >= 4.5 And normalized < 5.5
            ResolveIVABucket = "IVA5"
        Case Else
            If normalized = 0 Then
                ResolveIVABucket = "IVA0"
            Else
                ResolveIVABucket = "IVA0"
            End If
    End Select
End Function

Private Function ParseNotaCreditoDom(ByVal root As Object, _
                                     ByVal origenArchivo As String, _
                                     ByVal rutaArchivo As String, _
                                     ByVal nombreArchivo As String) As ParsedComprobante
    Dim header As Object
    Dim detalles As Collection
    Dim infoTrib As Object
    Dim infoNota As Object
    Dim numeroDocumento As String
    Dim fechaEmision As Variant
    Dim nodoImpuestos As Object
    Dim impuesto As Object
    Dim idx As Long
    Dim lineNode As Object
    Dim result As ParsedComprobante

    Set header = InitializeNotaCreditoHeader(origenArchivo, rutaArchivo, nombreArchivo)
    Set detalles = New Collection

    Set infoTrib = root.selectSingleNode("infoTributaria")
    Set infoNota = root.selectSingleNode("infoNotaCredito")

    header("TipoComprobante") = "04"
    header("RUC_Emisor") = SelectSingleNodeText(infoTrib, "ruc")
    header("TipoIdentificacionEmisor") = DefaultText(SelectSingleNodeText(infoTrib, "tipoIdentificacion"), "RUC")
    header("Nombre_Emisor") = DefaultText(SelectSingleNodeText(infoTrib, "razonSocial"), _
                                           SelectSingleNodeText(infoTrib, "nombreComercial"))
    header("RUC_Receptor") = SelectSingleNodeText(infoNota, "identificacionComprador")
    header("TipoIdentificacionReceptor") = SelectSingleNodeText(infoNota, "tipoIdentificacionComprador")
    header("Nombre_Receptor") = SelectSingleNodeText(infoNota, "razonSocialComprador")
    header("Establecimiento") = SelectSingleNodeText(infoTrib, "estab")
    header("PuntoEmision") = SelectSingleNodeText(infoTrib, "ptoEmi")
    header("Secuencial") = SelectSingleNodeText(infoTrib, "secuencial")
    header("ClaveAcceso") = SelectSingleNodeText(infoTrib, "claveAcceso")
    header("Moneda") = DefaultText(SelectSingleNodeText(infoNota, "moneda"), "USD")

    numeroDocumento = BuildNumeroDocumento(header("Establecimiento"), header("PuntoEmision"), header("Secuencial"))
    header("NumeroDocumento") = numeroDocumento
    header("DocumentoModifica") = SelectSingleNodeText(infoNota, "numDocModificado")
    header("Motivo") = SelectSingleNodeText(infoNota, "motivo")

    fechaEmision = ParseFechaEmision(SelectSingleNodeText(infoNota, "fechaEmision"))
    If IsDate(fechaEmision) Then
        header("FechaEmision") = fechaEmision
    End If

    header("ValorTotal") = ParseSRINumber(SelectSingleNodeText(infoNota, "valorModificacion"))

    Set nodoImpuestos = infoNota.selectNodes("totalConImpuestos/totalImpuesto")
    If Not nodoImpuestos Is Nothing Then
        For idx = 0 To nodoImpuestos.length - 1
            Set impuesto = nodoImpuestos.Item(idx)
            AggregateImpuestoNotaCredito header, impuesto
        Next idx
    End If

    Set nodoImpuestos = root.selectNodes("detalles/detalle")
    If Not nodoImpuestos Is Nothing Then
        For idx = 0 To nodoImpuestos.length - 1
            Set lineNode = nodoImpuestos.Item(idx)
            detalles.Add ParseNotaDetalle(lineNode, numeroDocumento, idx + 1)
        Next idx
    End If

    header("Observaciones") = JoinAdditionalText(root.selectNodes("infoAdicional/campoAdicional"))

    result.Tipo = "04"
    Set result.Cabecera = header
    Set result.Detalles = detalles
    Set adicionales = CollectAdditionalFields(root.selectNodes("infoAdicional/campoAdicional"))
    Set result.Adicional = adicionales
    ParseNotaCreditoDom = result
End Function

Private Sub AggregateImpuestoNotaCredito(ByVal header As Object, ByVal impuesto As Object)
    Dim codigo As String
    Dim codigoPorcentaje As String
    Dim tarifa As Double
    Dim baseImponible As Double
    Dim valor As Double

    codigo = SelectSingleNodeText(impuesto, "codigo")
    codigoPorcentaje = SelectSingleNodeText(impuesto, "codigoPorcentaje")
    tarifa = ParseSRINumber(SelectSingleNodeText(impuesto, "tarifa"))
    baseImponible = ParseSRINumber(SelectSingleNodeText(impuesto, "baseImponible"))
    valor = ParseSRINumber(SelectSingleNodeText(impuesto, "valor"))

    If codigo = "2" Then
        Select Case ResolveIVABucket(tarifa, codigoPorcentaje)
            Case "IVA15"
                header("BaseIVA15") = header("BaseIVA15") + baseImponible
                header("ValorIVA15") = header("ValorIVA15") + valor
            Case "IVA12"
                header("BaseIVA12") = header("BaseIVA12") + baseImponible
                header("ValorIVA12") = header("ValorIVA12") + valor
            Case "IVA0"
                header("BaseIVA0") = header("BaseIVA0") + baseImponible
                header("ValorIVA0") = header("ValorIVA0") + valor
            Case "NoObjeto"
                header("BaseNoObjetoIVA") = header("BaseNoObjetoIVA") + baseImponible
            Case "Exento"
                header("BaseExentaIVA") = header("BaseExentaIVA") + baseImponible
        End Select
    End If
End Sub

Private Function ParseNotaDetalle(ByVal lineNode As Object, _
                                  ByVal numeroDocumento As String, _
                                  ByVal lineaIndex As Long) As Object
    Dim detail As Object
    Dim impuestos As Object
    Dim impuesto As Object
    Dim idx As Long
    Dim codigo As String
    Dim tarifa As Double

    Set detail = NewDictionary()

    detail("NumeroDocumento") = numeroDocumento
    detail("Linea") = Format$(lineaIndex, "0000")
    detail("CodigoPrincipal") = SelectSingleNodeText(lineNode, "codigoPrincipal")
    detail("CodigoAuxiliar") = SelectSingleNodeText(lineNode, "codigoAuxiliar")
    detail("Descripcion") = SelectSingleNodeText(lineNode, "descripcion")
    detail("Cantidad") = ParseSRINumber(SelectSingleNodeText(lineNode, "cantidad"))
    detail("UnidadMedida") = SelectSingleNodeText(lineNode, "unidadMedida")
    detail("PrecioUnitario") = ParseSRINumber(SelectSingleNodeText(lineNode, "precioUnitario"))
    detail("Descuento") = ParseSRINumber(SelectSingleNodeText(lineNode, "descuento"))
    detail("PrecioTotalSinImpuesto") = ParseSRINumber(SelectSingleNodeText(lineNode, "precioTotalSinImpuesto"))
    detail("TarifaIVA") = 0
    detail("PorcentajeIVA") = 0
    detail("BaseIVA") = 0
    detail("ValorIVA") = 0
    detail("TarifaICE") = 0
    detail("ValorICE") = 0
    detail("DetalleAdicional1") = ""
    detail("DetalleAdicional2") = ""
    detail("DetalleAdicional3") = ""

    Set impuestos = lineNode.selectNodes("impuestos/impuesto")
    If Not impuestos Is Nothing Then
        For idx = 0 To impuestos.length - 1
            Set impuesto = impuestos.Item(idx)
            codigo = SelectSingleNodeText(impuesto, "codigo")
            tarifa = ParseSRINumber(SelectSingleNodeText(impuesto, "tarifa"))
            If codigo = "2" Then
                detail("TarifaIVA") = NormalizePercent(tarifa)
                detail("PorcentajeIVA") = NormalizePercent(tarifa)
                detail("BaseIVA") = ParseSRINumber(SelectSingleNodeText(impuesto, "baseImponible"))
                detail("ValorIVA") = ParseSRINumber(SelectSingleNodeText(impuesto, "valor"))
            ElseIf codigo = "3" Then
                detail("TarifaICE") = NormalizePercent(tarifa)
                detail("ValorICE") = ParseSRINumber(SelectSingleNodeText(impuesto, "valor"))
            End If
        Next idx
    End If

    AssignDetalleAdicionales lineNode, detail
    Set ParseNotaDetalle = detail
End Function

Private Function ParseRetencionDom(ByVal root As Object, _
                                   ByVal origenArchivo As String, _
                                   ByVal rutaArchivo As String, _
                                   ByVal nombreArchivo As String) As ParsedComprobante
    Dim header As Object
    Dim detalles As Collection
    Dim infoTrib As Object
    Dim infoRet As Object
    Dim impuestos As Object
    Dim impuesto As Object
    Dim idx As Long
    Dim numeroDocumento As String
    Dim fechaEmision As Variant
    Dim result As ParsedComprobante
    Dim adicionales As Object

    Set header = InitializeRetencionHeader(origenArchivo, rutaArchivo, nombreArchivo)
    Set detalles = New Collection

    Set infoTrib = root.selectSingleNode("infoTributaria")
    Set infoRet = root.selectSingleNode("infoCompRetencion")

    header("TipoComprobante") = "07"
    header("RUC_Agente") = SelectSingleNodeText(infoTrib, "ruc")
    header("TipoIdentificacionAgente") = DefaultText(SelectSingleNodeText(infoTrib, "tipoIdentificacion"), "RUC")
    header("Nombre_Agente") = DefaultText(SelectSingleNodeText(infoTrib, "razonSocial"), _
                                           SelectSingleNodeText(infoTrib, "nombreComercial"))
    header("RUC_Sujeto") = SelectSingleNodeText(infoRet, "identificacionSujetoRetenido")
    header("TipoIdentificacionSujeto") = SelectSingleNodeText(infoRet, "tipoIdentificacionSujetoRetenido")
    header("Nombre_Sujeto") = SelectSingleNodeText(infoRet, "razonSocialSujetoRetenido")
    header("Establecimiento") = SelectSingleNodeText(infoTrib, "estab")
    header("PuntoEmision") = SelectSingleNodeText(infoTrib, "ptoEmi")
    header("Secuencial") = SelectSingleNodeText(infoTrib, "secuencial")
    header("ClaveAcceso") = SelectSingleNodeText(infoTrib, "claveAcceso")
    header("PeriodoFiscal") = SelectSingleNodeText(infoRet, "periodoFiscal")
    header("Moneda") = DefaultText(SelectSingleNodeText(infoRet, "moneda"), header("Moneda"))

    numeroDocumento = BuildNumeroDocumento(header("Establecimiento"), header("PuntoEmision"), header("Secuencial"))
    header("NumeroDocumento") = numeroDocumento

    fechaEmision = ParseFechaEmision(SelectSingleNodeText(infoRet, "fechaEmision"))
    If IsDate(fechaEmision) Then
        header("FechaEmision") = fechaEmision
    End If

    Set impuestos = root.selectNodes("impuestos/impuesto")
    If Not impuestos Is Nothing Then
        For idx = 0 To impuestos.length - 1
            Set impuesto = impuestos.Item(idx)
            detalles.Add ParseRetencionDetalle(impuesto, numeroDocumento, idx + 1)
            If header("DocumentoSustento") = "" Then
                header("DocumentoSustento") = SelectSingleNodeText(impuesto, "codDocSustento")
            End If
            If header("NumeroSustento") = "" Then
                header("NumeroSustento") = SelectSingleNodeText(impuesto, "numDocSustento")
            End If
            AggregateRetencion header, impuesto
        Next idx
    End If

    header("Observaciones") = JoinAdditionalText(root.selectNodes("infoAdicional/campoAdicional"))

    result.Tipo = "07"
    Set result.Cabecera = header
    Set result.Detalles = detalles
    Set adicionales = CollectAdditionalFields(root.selectNodes("infoAdicional/campoAdicional"))
    Set result.Adicional = adicionales
    ParseRetencionDom = result
End Function

Private Function ParseRetencionDetalle(ByVal impuesto As Object, _
                                       ByVal numeroDocumento As String, _
                                       ByVal lineaIndex As Long) As Object
    Dim detail As Object
    Dim fechaSustento As Variant

    Set detail = NewDictionary()

    detail("NumeroDocumento") = numeroDocumento
    detail("Linea") = Format$(lineaIndex, "0000")
    detail("Impuesto") = SelectSingleNodeText(impuesto, "codigo")
    detail("CodigoRetencion") = SelectSingleNodeText(impuesto, "codigoRetencion")
    detail("Descripcion") = BuildDescripcionRetencion(detail("Impuesto"), detail("CodigoRetencion"))
    detail("BaseImponible") = ParseSRINumber(SelectSingleNodeText(impuesto, "baseImponible"))
    detail("PorcentajeRetencion") = NormalizePercent(ParseSRINumber(SelectSingleNodeText(impuesto, "porcentajeRetener")))
    detail("ValorRetenido") = ParseSRINumber(SelectSingleNodeText(impuesto, "valorRetenido"))
    detail("TipoDocumentoSustento") = SelectSingleNodeText(impuesto, "codDocSustento")
    detail("NumeroDocumentoSustento") = SelectSingleNodeText(impuesto, "numDocSustento")

    fechaSustento = ParseFechaEmision(SelectSingleNodeText(impuesto, "fechaEmisionDocSustento"))
    If IsDate(fechaSustento) Then
        detail("FechaEmisionSustento") = fechaSustento
    Else
        detail("FechaEmisionSustento") = Empty
    End If

    Set ParseRetencionDetalle = detail
End Function

Private Sub AggregateRetencion(ByVal header As Object, ByVal impuesto As Object)
    Dim codigo As String
    Dim baseImponible As Double
    Dim porcentaje As Double
    Dim valor As Double

    codigo = SelectSingleNodeText(impuesto, "codigo")
    baseImponible = ParseSRINumber(SelectSingleNodeText(impuesto, "baseImponible"))
    porcentaje = ParseSRINumber(SelectSingleNodeText(impuesto, "porcentajeRetener"))
    valor = ParseSRINumber(SelectSingleNodeText(impuesto, "valorRetenido"))

    Select Case codigo
        Case "2"
            header("BaseIVA") = header("BaseIVA") + baseImponible
            header("ValorRetenidoIVA") = header("ValorRetenidoIVA") + valor
            header("PorcentajeIVA") = NormalizePercent(porcentaje)
        Case "1"
            header("BaseRenta") = header("BaseRenta") + baseImponible
            header("ValorRetenidoRenta") = header("ValorRetenidoRenta") + valor
            header("PorcentajeRenta") = NormalizePercent(porcentaje)
        Case "6"
            header("BaseISD") = header("BaseISD") + baseImponible
            header("ValorRetenidoISD") = header("ValorRetenidoISD") + valor
            header("PorcentajeISD") = NormalizePercent(porcentaje)
    End Select
End Sub

'------------------------------------------------------------------------------------------
' Helpers de inicialización
'------------------------------------------------------------------------------------------
Private Function InitializeFacturaHeader(ByVal origenArchivo As String, _
                                         ByVal rutaArchivo As String, _
                                         ByVal nombreArchivo As String) As Object
    Dim header As Object

    Set header = NewDictionary()

    header("TipoComprobante") = ""
    header("Subtipo") = ""
    header("FechaEmision") = Empty
    header("FechaRegistro") = Date
    header("RUC_Emisor") = ""
    header("TipoIdentificacionEmisor") = ""
    header("Nombre_Emisor") = ""
    header("RUC_Receptor") = ""
    header("TipoIdentificacionReceptor") = ""
    header("Nombre_Receptor") = ""
    header("NumeroDocumento") = ""
    header("Establecimiento") = ""
    header("PuntoEmision") = ""
    header("Secuencial") = ""
    header("ClaveAcceso") = ""
    header("Moneda") = ""
    header("BaseIVA15") = 0#
    header("BaseIVA12") = 0#
    header("BaseIVA8") = 0#
    header("BaseIVA5") = 0#
    header("BaseIVA0") = 0#
    header("BaseNoObjetoIVA") = 0#
    header("BaseExentaIVA") = 0#
    header("BaseICE") = 0#
    header("ValorIVA15") = 0#
    header("ValorIVA12") = 0#
    header("ValorIVA8") = 0#
    header("ValorIVA5") = 0#
    header("ValorIVA0") = 0#
    header("ValorIVAExento") = 0#
    header("ValorICE") = 0#
    header("ValorIRBPNR") = 0#
    header("Propina") = 0#
    header("TotalSinImpuestos") = 0#
    header("ValorTotal") = 0#
    header("Estado") = "AUTORIZADO"
    header("OrigenArchivo") = origenArchivo
    header("RutaArchivo") = rutaArchivo
    header("NombreArchivo") = nombreArchivo
    header("EnlaceComprobante") = ""
    header("CreadoEn") = Date
    header("Observaciones") = ""

    Set InitializeFacturaHeader = header
End Function

Private Function InitializeNotaCreditoHeader(ByVal origenArchivo As String, _
                                             ByVal rutaArchivo As String, _
                                             ByVal nombreArchivo As String) As Object
    Dim header As Object

    Set header = NewDictionary()

    header("TipoComprobante") = ""
    header("FechaEmision") = Empty
    header("FechaRegistro") = Date
    header("RUC_Emisor") = ""
    header("TipoIdentificacionEmisor") = ""
    header("Nombre_Emisor") = ""
    header("RUC_Receptor") = ""
    header("TipoIdentificacionReceptor") = ""
    header("Nombre_Receptor") = ""
    header("NumeroDocumento") = ""
    header("DocumentoModifica") = ""
    header("Motivo") = ""
    header("ClaveAcceso") = ""
    header("BaseIVA15") = 0#
    header("BaseIVA12") = 0#
    header("BaseIVA0") = 0#
    header("BaseNoObjetoIVA") = 0#
    header("BaseExentaIVA") = 0#
    header("ValorIVA15") = 0#
    header("ValorIVA12") = 0#
    header("ValorIVA0") = 0#
    header("ValorTotal") = 0#
    header("OrigenArchivo") = origenArchivo
    header("RutaArchivo") = rutaArchivo
    header("NombreArchivo") = nombreArchivo
    header("EnlaceComprobante") = ""
    header("CreadoEn") = Date
    header("Observaciones") = ""

    Set InitializeNotaCreditoHeader = header
End Function

Private Function InitializeRetencionHeader(ByVal origenArchivo As String, _
                                           ByVal rutaArchivo As String, _
                                           ByVal nombreArchivo As String) As Object
    Dim header As Object

    Set header = NewDictionary()

    header("TipoComprobante") = ""
    header("FechaEmision") = Empty
    header("FechaRegistro") = Date
    header("RUC_Agente") = ""
    header("TipoIdentificacionAgente") = ""
    header("Nombre_Agente") = ""
    header("RUC_Sujeto") = ""
    header("TipoIdentificacionSujeto") = ""
    header("Nombre_Sujeto") = ""
    header("NumeroDocumento") = ""
    header("PeriodoFiscal") = ""
    header("DocumentoSustento") = ""
    header("NumeroSustento") = ""
    header("ClaveAcceso") = ""
    header("Moneda") = "USD"
    header("BaseIVA") = 0#
    header("PorcentajeIVA") = 0#
    header("ValorRetenidoIVA") = 0#
    header("BaseRenta") = 0#
    header("PorcentajeRenta") = 0#
    header("ValorRetenidoRenta") = 0#
    header("BaseISD") = 0#
    header("PorcentajeISD") = 0#
    header("ValorRetenidoISD") = 0#
    header("OrigenArchivo") = origenArchivo
    header("RutaArchivo") = rutaArchivo
    header("NombreArchivo") = nombreArchivo
    header("EnlaceComprobante") = ""
    header("CreadoEn") = Date
    header("Observaciones") = ""

    Set InitializeRetencionHeader = header
End Function

'------------------------------------------------------------------------------------------
' Helpers generales
'------------------------------------------------------------------------------------------
Private Function NewDictionary() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    dict.CompareMode = 1
    On Error GoTo 0
    Set NewDictionary = dict
End Function

Private Function NormalizePercent(ByVal valor As Double) As Double
    If valor = 0 Then
        NormalizePercent = 0
    Else
        NormalizePercent = valor / 100
    End If
End Function

Private Function DefaultText(ByVal primaryValue As String, _
                             Optional ByVal fallbackValue As String = "") As String
    If Trim$(primaryValue) <> "" Then
        DefaultText = Trim$(primaryValue)
    Else
        DefaultText = Trim$(fallbackValue)
    End If
End Function

Private Function ParseFechaEmision(ByVal texto As String) As Variant
    Dim parsed As Variant

    parsed = ParseSRIToDate(texto)
    If IsDate(parsed) Then
        ParseFechaEmision = parsed
        Exit Function
    End If

    On Error Resume Next
    parsed = CDate(texto)
    On Error GoTo 0

    If IsDate(parsed) Then
        ParseFechaEmision = parsed
    Else
        ParseFechaEmision = Empty
    End If
End Function

Private Function BuildNumeroDocumento(ByVal estab As String, _
                                      ByVal ptoEmi As String, _
                                      ByVal secuencial As String) As String
    If Trim$(estab) = "" And Trim$(ptoEmi) = "" And Trim$(secuencial) = "" Then
        BuildNumeroDocumento = ""
    Else
        BuildNumeroDocumento = FormatNumero(estab) & "-" & FormatNumero(ptoEmi) & _
                               "-" & FormatNumero(secuencial, 9)
    End If
End Function

Private Function FormatNumero(ByVal valor As String, Optional ByVal largo As Long = 3) As String
    Dim limpio As String
    limpio = SoloDigitos(valor)
    If limpio = "" Then limpio = "0"
    FormatNumero = Right$(String$(largo, "0") & limpio, largo)
End Function

Private Function CollectAdditionalFields(ByVal nodeList As Object) As Object
    Dim dict As Object
    Dim idx As Long
    Dim node As Object
    Dim nombre As String
    Dim valor As String
    Dim baseName As String
    Dim suffix As Long

    Set dict = NewDictionary()

    If nodeList Is Nothing Then
        Set CollectAdditionalFields = dict
        Exit Function
    End If

    For idx = 0 To nodeList.length - 1
        Set node = nodeList.Item(idx)
        nombre = GetAttributeValue(node, "nombre")
        valor = Trim$(CStr(node.text))
        If nombre = "" Then
            nombre = "Campo" & Format$(idx + 1, "00")
        End If

        baseName = nombre
        suffix = 1
        Do While dict.Exists(nombre)
            nombre = baseName & "_" & Format$(suffix, "00")
            suffix = suffix + 1
        Loop

        dict(nombre) = valor
    Next idx

    Set CollectAdditionalFields = dict
End Function

Private Function JoinAdditionalText(ByVal nodeList As Object) As String
    Dim partes() As String
    Dim idx As Long
    Dim node As Object
    Dim nombre As String
    Dim valor As String

    If nodeList Is Nothing Then Exit Function
    If nodeList.length = 0 Then Exit Function

    ReDim partes(0 To nodeList.length - 1)

    For idx = 0 To nodeList.length - 1
        Set node = nodeList.Item(idx)
        nombre = GetAttributeValue(node, "nombre")
        valor = Trim$(CStr(node.text))
        If nombre <> "" Then
            partes(idx) = nombre & ": " & valor
        Else
            partes(idx) = valor
        End If
    Next idx

    JoinAdditionalText = Join(partes, " | ")
End Function

Private Sub AssignDetalleAdicionales(ByVal lineNode As Object, ByVal detail As Object)
    Dim adicionales As Object
    Dim idx As Long
    Dim node As Object

    Set adicionales = lineNode.selectNodes("detallesAdicionales/detAdicional")

    If adicionales Is Nothing Then Exit Sub
    If adicionales.length = 0 Then Exit Sub

    For idx = 0 To adicionales.length - 1
        If idx > 2 Then Exit For
        Set node = adicionales.Item(idx)
        detail("DetalleAdicional" & (idx + 1)) = Trim$(CStr(node.text))
    Next idx
End Sub

Private Function BuildDescripcionRetencion(ByVal codigo As String, ByVal codigoRetencion As String) As String
    Dim tipo As String

    Select Case codigo
        Case "1"
            tipo = "RENTA"
        Case "2"
            tipo = "IVA"
        Case "6"
            tipo = "ISD"
        Case Else
            tipo = "IMPUESTO"
    End Select

    If Trim$(codigoRetencion) <> "" Then
        BuildDescripcionRetencion = tipo & " - Código " & codigoRetencion
    Else
        BuildDescripcionRetencion = tipo
    End If
End Function

Private Function IsExistingFile(ByVal path As String) As Boolean
    On Error Resume Next
    IsExistingFile = (Dir$(path) <> "")
    On Error GoTo 0
End Function

Private Function ExtractFileName(ByVal path As String) As String
    Dim parts() As String
    parts = Split(path, "\\")
    ExtractFileName = parts(UBound(parts))
    If InStr(ExtractFileName, "/") > 0 Then
        parts = Split(ExtractFileName, "/")
        ExtractFileName = parts(UBound(parts))
    End If
End Function

Private Function HasProperty(ByVal obj As Object, ByVal propName As String) As Boolean
    On Error Resume Next
    Dim temp As Variant
    temp = CallByName(obj, propName, VbGet)
    HasProperty = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

