' Attribute VB_Name = "modUtils"
Option Explicit

'##########################################################################################
'# Módulo: modUtils
'# Propósito: Utilidades generales para importar y validar XML del SRI.
'# Alcance:   Funciones de XPath, conversión de fechas y montos, validadores tributarios
'#            y helpers generales reutilizables por otros módulos del proyecto.
'# Notas:     Todo el código usa enlace tardío para maximizar compatibilidad con Excel 2007+
'##########################################################################################

'--- Constantes públicas ---
Public Const ISO_DATE_PATTERN As String = "yyyy-mm-dd"
Public Const ISO_DATETIME_PATTERN As String = "yyyy-mm-ddThh:nn:ss"
Public Const ISO_TZ_PATTERN As String = "yyyy-mm-ddThh:nn:ss+00:00"
Public Const DEFAULT_DECIMAL_SEPARATOR As String = "."

'--- Tipos públicos ---
Public Type XPathResult
    Exists As Boolean
    Value As String
End Type

'##########################################################################################
'# Sección: Helpers de XPath y XML
'##########################################################################################

Public Function CreateDomDocument(Optional ByVal async As Boolean = False, _
                                  Optional ByVal validateOnParse As Boolean = False) As Object
    Dim candidates As Variant
    Dim idx As Long
    Dim dom As Object

    candidates = Array("MSXML2.DOMDocument.6.0", "MSXML2.DOMDocument.3.0", "MSXML2.DOMDocument")

    For idx = LBound(candidates) To UBound(candidates)
        On Error Resume Next
        Set dom = CreateObject(candidates(idx))
        On Error GoTo 0
        If Not dom Is Nothing Then Exit For
    Next idx

    If dom Is Nothing Then
        Err.Raise vbObjectError + 513, "modUtils.CreateDomDocument", _
                  "No se pudo crear una instancia de MSXML. Instale MSXML 3.0 o superior."
    End If

    dom.async = async
    dom.validateOnParse = validateOnParse
    dom.resolveExternals = False

    On Error Resume Next
    dom.setProperty "SelectionLanguage", "XPath"
    On Error GoTo 0

    Set CreateDomDocument = dom
End Function

Public Function ApplySelectionNamespaces(ByVal dom As Object, _
                                         ByVal namespaces As Variant) As String
    Dim previousValue As String
    Dim nsLiteral As String

    If dom Is Nothing Then Exit Function

    On Error Resume Next
    previousValue = dom.getProperty("SelectionNamespaces")
    On Error GoTo 0

    nsLiteral = BuildSelectionNamespaces(namespaces)

    On Error Resume Next
    dom.setProperty "SelectionNamespaces", nsLiteral
    On Error GoTo 0

    ApplySelectionNamespaces = previousValue
End Function

Public Function BuildSelectionNamespaces(ByVal namespaces As Variant) As String
    Dim result As String

    If IsObject(namespaces) Then
        Select Case TypeName(namespaces)
            Case "Dictionary"
                result = BuildNamespacesFromDictionary(namespaces)
            Case "Collection"
                result = BuildNamespacesFromCollection(namespaces)
            Case Else
                result = CStr(namespaces)
        End Select
    ElseIf IsArray(namespaces) Then
        result = BuildNamespacesFromArray(namespaces)
    Else
        result = CStr(namespaces)
    End If

    BuildSelectionNamespaces = Trim$(result)
End Function

Private Function BuildNamespacesFromDictionary(ByVal dict As Object) As String
    Dim parts() As String
    Dim key As Variant
    Dim idx As Long

    If dict Is Nothing Then Exit Function
    If dict.Count = 0 Then Exit Function

    ReDim parts(0 To dict.Count - 1)

    For Each key In dict.Keys
        parts(idx) = "xmlns:" & CStr(key) & "='" & CStr(dict(key)) & "'"
        idx = idx + 1
    Next key

    BuildNamespacesFromDictionary = Join(parts, " ")
End Function

Private Function BuildNamespacesFromCollection(ByVal coll As Object) As String
    Dim parts() As String
    Dim idx As Long
    Dim raw As String
    Dim splitPos As Long
    Dim prefix As String
    Dim uri As String

    If coll Is Nothing Then Exit Function
    If coll.Count = 0 Then Exit Function

    ReDim parts(0 To coll.Count - 1)

    For idx = 1 To coll.Count
        raw = CStr(coll.Item(idx))
        splitPos = InStr(raw, "=")
        If splitPos > 0 Then
            prefix = Trim$(Left$(raw, splitPos - 1))
            uri = Trim$(Mid$(raw, splitPos + 1))
            parts(idx - 1) = "xmlns:" & prefix & "='" & uri & "'"
        Else
            parts(idx - 1) = Trim$(raw)
        End If
    Next idx

    BuildNamespacesFromCollection = Join(parts, " ")
End Function

Private Function BuildNamespacesFromArray(ByVal items As Variant) As String
    Dim parts() As String
    Dim idx As Long

    If Not IsArray(items) Then Exit Function

    ReDim parts(LBound(items) To UBound(items))

    For idx = LBound(items) To UBound(items)
        parts(idx) = CStr(items(idx))
    Next idx

    BuildNamespacesFromArray = Join(parts, " ")
End Function

Public Sub RestoreSelectionNamespaces(ByVal dom As Object, _
                                       ByVal previousValue As String)
    If dom Is Nothing Then Exit Sub

    On Error Resume Next
    If previousValue = vbNullString Then
        dom.setProperty "SelectionNamespaces", ""
    Else
        dom.setProperty "SelectionNamespaces", previousValue
    End If
    On Error GoTo 0
End Sub

Public Function GetOwnerDocument(ByVal contextNode As Object) As Object
    Dim dom As Object

    If contextNode Is Nothing Then Exit Function

    On Error Resume Next
    Set dom = contextNode.ownerDocument
    On Error GoTo 0

    If dom Is Nothing Then
        Set dom = contextNode
    End If

    Set GetOwnerDocument = dom
End Function

Public Function SelectSingleNodeText(ByVal contextNode As Object, _
                                     ByVal xpath As String, _
                                     Optional ByVal defaultValue As String = "", _
                                     Optional ByVal namespaces As Variant) As String
    Dim dom As Object
    Dim previousValue As String
    Dim node As Object

    If contextNode Is Nothing Then
        SelectSingleNodeText = defaultValue
        Exit Function
    End If

    Set dom = GetOwnerDocument(contextNode)

    If Not IsMissing(namespaces) Then
        previousValue = ApplySelectionNamespaces(dom, namespaces)
    End If

    On Error Resume Next
    Set node = contextNode.selectSingleNode(xpath)
    On Error GoTo 0

    If Not node Is Nothing Then
        SelectSingleNodeText = Trim$(CStr(node.text))
    Else
        SelectSingleNodeText = defaultValue
    End If

    If Not IsMissing(namespaces) Then
        RestoreSelectionNamespaces dom, previousValue
    End If
End Function

Public Function SelectNodes(ByVal contextNode As Object, _
                            ByVal xpath As String, _
                            Optional ByVal namespaces As Variant) As Object
    Dim dom As Object
    Dim previousValue As String
    Dim nodeList As Object

    If contextNode Is Nothing Then Exit Function

    Set dom = GetOwnerDocument(contextNode)

    If Not IsMissing(namespaces) Then
        previousValue = ApplySelectionNamespaces(dom, namespaces)
    End If

    On Error Resume Next
    Set nodeList = contextNode.selectNodes(xpath)
    On Error GoTo 0

    If Not IsMissing(namespaces) Then
        RestoreSelectionNamespaces dom, previousValue
    End If

    Set SelectNodes = nodeList
End Function

Public Function GetAttributeValue(ByVal contextNode As Object, _
                                  ByVal attributeName As String, _
                                  Optional ByVal defaultValue As String = "") As String
    Dim attrNode As Object

    If contextNode Is Nothing Then
        GetAttributeValue = defaultValue
        Exit Function
    End If

    On Error Resume Next
    Set attrNode = contextNode.Attributes.getNamedItem(attributeName)
    On Error GoTo 0

    If attrNode Is Nothing Then
        GetAttributeValue = defaultValue
    Else
        GetAttributeValue = Trim$(CStr(attrNode.text))
    End If
End Function

Public Function NodeExists(ByVal contextNode As Object, _
                           ByVal xpath As String, _
                           Optional ByVal namespaces As Variant) As Boolean
    Dim result As XPathResult
    result = GetXPathResult(contextNode, xpath, namespaces)
    NodeExists = result.Exists
End Function

Public Function GetXPathResult(ByVal contextNode As Object, _
                               ByVal xpath As String, _
                               Optional ByVal namespaces As Variant) As XPathResult
    Dim result As XPathResult
    Dim dom As Object
    Dim previousValue As String
    Dim node As Object

    result.Exists = False
    result.Value = ""

    If contextNode Is Nothing Then
        GetXPathResult = result
        Exit Function
    End If

    Set dom = GetOwnerDocument(contextNode)

    If Not IsMissing(namespaces) Then
        previousValue = ApplySelectionNamespaces(dom, namespaces)
    End If

    On Error Resume Next
    Set node = contextNode.selectSingleNode(xpath)
    On Error GoTo 0

    If Not node Is Nothing Then
        result.Exists = True
        result.Value = Trim$(CStr(node.text))
    End If

    If Not IsMissing(namespaces) Then
        RestoreSelectionNamespaces dom, previousValue
    End If

    GetXPathResult = result
End Function

'##########################################################################################
'# Sección: Conversión de fechas y valores
'##########################################################################################

Public Function ParseSRIToDate(ByVal fechaTexto As String) As Variant
    Dim cleaned As String
    Dim work As String
    Dim tzPos As Long
    Dim spacePos As Long
    Dim datePart As String
    Dim timePart As String
    Dim parts() As String
    Dim yearPart As Long
    Dim monthPart As Long
    Dim dayPart As Long
    Dim hourPart As Long
    Dim minutePart As Long
    Dim secondPart As Long
    Dim resultDate As Date

    cleaned = Trim$(fechaTexto)
    If cleaned = "" Then Exit Function

    work = Replace(cleaned, "T", " ")
    work = Replace(work, "Z", "")

    tzPos = InStr(12, work, "+")
    If tzPos = 0 Then tzPos = InStr(12, work, "-")
    If tzPos > 0 Then work = Left$(work, tzPos - 1)

    spacePos = InStr(work, " ")
    If spacePos > 0 Then
        datePart = Trim$(Left$(work, spacePos - 1))
        timePart = Trim$(Mid$(work, spacePos + 1))
    Else
        datePart = work
        timePart = ""
    End If

    If InStr(datePart, "-") = 0 Then Exit Function

    parts = Split(datePart, "-")
    If UBound(parts) <> 2 Then Exit Function

    On Error GoTo ParseFail
    yearPart = CLng(parts(0))
    monthPart = CLng(parts(1))
    dayPart = CLng(parts(2))

    If timePart <> "" Then
        parts = Split(timePart, ":")
        If UBound(parts) >= 0 Then hourPart = CLng(Val(parts(0)))
        If UBound(parts) >= 1 Then minutePart = CLng(Val(parts(1)))
        If UBound(parts) >= 2 Then secondPart = CLng(Val(parts(2)))
    End If

    resultDate = DateSerial(yearPart, monthPart, dayPart) + _
                 TimeSerial(hourPart, minutePart, secondPart)

    ParseSRIToDate = resultDate
    On Error GoTo 0
    Exit Function

ParseFail:
    ParseSRIToDate = Empty
    On Error GoTo 0
End Function

Public Function FormatDateDDMMYYYY(ByVal fechaValor As Variant) As String
    If IsDate(fechaValor) Then
        FormatDateDDMMYYYY = Format$(CDate(fechaValor), "dd/mm/yyyy")
    Else
        FormatDateDDMMYYYY = ""
    End If
End Function

Public Function ParseSRINumber(ByVal valorTexto As String, _
                               Optional ByVal defaultValue As Double = 0#) As Double
    Dim cleaned As String
    Dim idx As Long
    Dim ch As String
    Dim numericText As String
    Dim decimalMarker As String
    Dim decimalSeparator As String
    Dim dotPos As Long
    Dim commaPos As Long
    Dim finalText As String
    Dim converted As Double

    cleaned = Trim$(valorTexto)
    If cleaned = "" Then
        ParseSRINumber = defaultValue
        Exit Function
    End If

    For idx = 1 To Len(cleaned)
        ch = Mid$(cleaned, idx, 1)
        If (ch >= "0" And ch <= "9") Or ch = "." Or ch = "," Then
            numericText = numericText & ch
        End If
    Next idx

    dotPos = InStrRev(numericText, ".")
    commaPos = InStrRev(numericText, ",")

    If dotPos = 0 And commaPos = 0 Then
        decimalMarker = ""
    ElseIf dotPos > commaPos Then
        decimalMarker = "."
    Else
        decimalMarker = ","
    End If

    decimalSeparator = Application.DecimalSeparator

    If decimalMarker <> "" Then
        If decimalMarker = "." Then
            numericText = Replace(numericText, ",", "")
        Else
            numericText = Replace(numericText, ".", "")
        End If
        finalText = Replace(numericText, decimalMarker, decimalSeparator, 1, 1)
    Else
        finalText = numericText
        If decimalSeparator <> DEFAULT_DECIMAL_SEPARATOR Then
            finalText = Replace(finalText, DEFAULT_DECIMAL_SEPARATOR, decimalSeparator)
        End If
    End If

    On Error Resume Next
    converted = CDbl(finalText)
    If Err.Number <> 0 Then
        converted = defaultValue
    End If
    On Error GoTo 0

    ParseSRINumber = converted
End Function

Public Function ToSRINumberText(ByVal valor As Variant, Optional ByVal decimales As Long = 2) As String
    Dim formatMask As String

    If Not IsNumeric(valor) Then
        ToSRINumberText = ""
        Exit Function
    End If

    If decimales < 0 Then decimales = 0
    If decimales > 10 Then decimales = 10

    formatMask = "0"
    If decimales > 0 Then
        formatMask = formatMask & "." & String$(decimales, "0")
    End If

    ToSRINumberText = Replace(Format$(CDbl(valor), formatMask), ",", ".")
End Function

Public Function EsFechaSRIVALida(ByVal fechaTexto As String) As Boolean
    Dim parsed As Variant
    parsed = ParseSRIToDate(fechaTexto)
    EsFechaSRIVALida = IsDate(parsed)
End Function

Public Function EsTipoComprobanteValido(ByVal tipo As String) As Boolean
    Dim codigo As String

    codigo = Format$(Val(SoloDigitos(tipo)), "00")

    Select Case codigo
        Case "01", "02", "03", "04", "05", "06", "07", "08", "09", _
             "10", "11", "12", "13", "14", "15", "16", "18", "19", _
             "20", "21", "22", "23", "24", "41", "42", "43", "44"
            EsTipoComprobanteValido = True
        Case Else
            EsTipoComprobanteValido = False
    End Select
End Function

'##########################################################################################
'# Sección: Validadores SRI
'##########################################################################################

Public Function EsRucValido(ByVal numero As String) As Boolean
    Dim limpio As String
    Dim tipo As Integer

    limpio = SoloDigitos(numero)

    If Len(limpio) <> 10 And Len(limpio) <> 13 Then Exit Function

    tipo = CInt(Mid$(limpio, 3, 1))

    Select Case tipo
        Case 0 To 5
            If Not ValidarCedula(limpio) Then Exit Function
            If Len(limpio) = 13 Then
                If Right$(limpio, 3) = "000" Then Exit Function
            End If
            EsRucValido = True
        Case 6
            If Len(limpio) <> 13 Then Exit Function
            If Not ValidarRucPublico(limpio) Then Exit Function
            If Right$(limpio, 4) = "0000" Then Exit Function
            EsRucValido = True
        Case 9
            If Len(limpio) <> 13 Then Exit Function
            If Not ValidarRucPrivado(limpio) Then Exit Function
            If Right$(limpio, 3) = "000" Then Exit Function
            EsRucValido = True
        Case Else
            EsRucValido = False
    End Select
End Function

Public Function EsClaveAccesoValida(ByVal clave As String) As Boolean
    Dim limpio As String
    Dim total As Long
    Dim factor As Integer
    Dim idx As Long
    Dim digito As Integer
    Dim verificador As Integer

    limpio = SoloDigitos(clave)

    If Len(limpio) <> 49 Then Exit Function

    factor = 2

    For idx = Len(limpio) - 1 To 1 Step -1
        digito = CInt(Mid$(limpio, idx, 1))
        total = total + digito * factor
        factor = factor + 1
        If factor > 7 Then factor = 2
    Next idx

    verificador = 11 - (total Mod 11)
    If verificador = 11 Then verificador = 0
    If verificador = 10 Then verificador = 1

    EsClaveAccesoValida = (verificador = CInt(Right$(limpio, 1)))
End Function

'##########################################################################################
'# Sección: Utilidades generales
'##########################################################################################

Public Function SoloDigitos(ByVal valor As String) As String
    Dim idx As Long
    Dim ch As String
    Dim resultado As String

    For idx = 1 To Len(valor)
        ch = Mid$(valor, idx, 1)
        If ch >= "0" And ch <= "9" Then
            resultado = resultado & ch
        End If
    Next idx

    SoloDigitos = resultado
End Function

Public Function TextoONulo(ByVal valor As String) As Variant
    If Trim$(valor) = "" Then
        TextoONulo = Null
    Else
        TextoONulo = valor
    End If
End Function

Public Function EnsureTrailingPathSeparator(ByVal ruta As String) As String
    If ruta = "" Then
        EnsureTrailingPathSeparator = ""
    ElseIf Right$(ruta, 1) = Application.PathSeparator Then
        EnsureTrailingPathSeparator = ruta
    Else
        EnsureTrailingPathSeparator = ruta & Application.PathSeparator
    End If
End Function

Public Function GetExcelMajorVersion() As Long
    GetExcelMajorVersion = CLng(Val(Application.Version))
End Function

'##########################################################################################
'# Sección: Soportes privados
'##########################################################################################

Private Function ValidarCedula(ByVal numero As String) As Boolean
    Dim base10 As String
    Dim pesos As Variant
    Dim idx As Long
    Dim suma As Long
    Dim digito As Integer
    Dim producto As Integer
    Dim verificador As Integer

    base10 = Left$(numero, 10)

    If Len(base10) <> 10 Then Exit Function
    If Not ProvinciaValida(base10) Then Exit Function

    pesos = Array(2, 1, 2, 1, 2, 1, 2, 1, 2)

    For idx = 0 To 8
        digito = CInt(Mid$(base10, idx + 1, 1))
        producto = digito * pesos(idx)
        If producto >= 10 Then producto = producto - 9
        suma = suma + producto
    Next idx

    verificador = (10 - (suma Mod 10)) Mod 10

    ValidarCedula = (verificador = CInt(Mid$(base10, 10, 1)))
End Function

Private Function ValidarRucPrivado(ByVal numero As String) As Boolean
    Dim base10 As String
    Dim pesos As Variant
    Dim idx As Long
    Dim suma As Long
    Dim digito As Integer
    Dim verificador As Integer

    base10 = Left$(numero, 10)

    If Len(base10) <> 10 Then Exit Function
    If Not ProvinciaValida(base10) Then Exit Function

    pesos = Array(4, 3, 2, 7, 6, 5, 4, 3, 2)

    For idx = 0 To 8
        digito = CInt(Mid$(base10, idx + 1, 1))
        suma = suma + digito * pesos(idx)
    Next idx

    verificador = 11 - (suma Mod 11)
    If verificador = 11 Then verificador = 0
    If verificador = 10 Then Exit Function

    ValidarRucPrivado = (verificador = CInt(Mid$(numero, 10, 1)))
End Function

Private Function ValidarRucPublico(ByVal numero As String) As Boolean
    Dim base9 As String
    Dim pesos As Variant
    Dim idx As Long
    Dim suma As Long
    Dim digito As Integer
    Dim verificador As Integer

    base9 = Left$(numero, 9)

    If Len(base9) <> 9 Then Exit Function
    If Not ProvinciaValida(base9) Then Exit Function

    pesos = Array(3, 2, 7, 6, 5, 4, 3, 2)

    For idx = 0 To 7
        digito = CInt(Mid$(base9, idx + 1, 1))
        suma = suma + digito * pesos(idx)
    Next idx

    verificador = 11 - (suma Mod 11)
    If verificador = 11 Then verificador = 0
    If verificador = 10 Then Exit Function

    ValidarRucPublico = (verificador = CInt(Mid$(numero, 9, 1)))
End Function

Private Function ProvinciaValida(ByVal numero As String) As Boolean
    Dim provincia As Integer

    On Error GoTo ProvinciaFail
    provincia = CInt(Left$(numero, 2))
    ProvinciaValida = (provincia >= 1 And provincia <= 24) Or provincia = 30
    On Error GoTo 0
    Exit Function

ProvinciaFail:
    ProvinciaValida = False
    On Error GoTo 0
End Function

'##########################################################################################
'# Sección: Pruebas rápidas
'##########################################################################################

Public Sub Run_SmokeTest()
    Dim dom As Object
    Dim xmlTexto As String
    Dim nodoValor As String
    Dim fecha As Variant
    Dim numero As Double

    xmlTexto = "<Factura xmlns='urn:test'><infoFactura><fechaEmision>2024-01-05" & _
               "</fechaEmision><totalSinImpuestos>123.45</totalSinImpuestos></infoFactura></Factura>"

    Set dom = CreateDomDocument()
    dom.LoadXML xmlTexto

    nodoValor = SelectSingleNodeText(dom, "//infoFactura/fechaEmision", "", "xmlns='urn:test'")
    Debug.Print "Fecha (texto):", nodoValor

    fecha = ParseSRIToDate(nodoValor)
    Debug.Print "Fecha (formateada):", FormatDateDDMMYYYY(fecha)

    numero = ParseSRINumber("1,234.56")
    Debug.Print "Monto convertido:", numero
    Debug.Print "Monto SRI:", ToSRINumberText(numero, 2)
    Debug.Print "RUC válido 1790012345001:", EsRucValido("1790012345001")
    Debug.Print "Clave válida ejemplo:", EsClaveAccesoValida("1505202101179001324500110010010000000011234567819")
End Sub

