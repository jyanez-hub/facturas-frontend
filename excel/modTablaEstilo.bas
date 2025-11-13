Attribute VB_Name = "modTablaEstilo"
' ===================  modTablaEstilo (única versión)  ===================
Option Explicit
'Option Private Module   ' <- si quieres ocultar públicos a otros proyectos

' ---- Colores corporativos ----
Private Const HEX_AZUL As String = "#1C0F82"           ' encabezado

' === Compatibilidad: elementos de estilo de tabla (valores de Excel) ===
Private Enum teElem ' TableStyleElementType
    teWholeTable = 0
    teHeaderRow = 1
    teTotalRow = 2
    teGrandTotalRow = 3
    teFirstColumn = 4
    teLastColumn = 5
    teFirstRowStripe = 7
    teSecondRowStripe = 8
    teFirstColumnStripe = 9
    teSecondColumnStripe = 10
End Enum

' ---- Util ----
Private Function HexToRGB(ByVal hexColor As String) As Long
    Dim r As Long, g As Long, b As Long, h As String
    h = Replace(hexColor, "#", "")
    r = CLng("&H" & Mid$(h, 1, 2))
    g = CLng("&H" & Mid$(h, 3, 2))
    b = CLng("&H" & Mid$(h, 5, 2))
    HexToRGB = RGB(r, g, b)
End Function

' Crea/asegura el estilo corporativo
Private Function EnsureExcelbotStyle(ByVal wb As Workbook, _
    Optional ByVal styleName As String = "EXCELBOT_AzulCorp") As String

    Dim ts As TableStyle
    On Error Resume Next
    Set ts = wb.TableStyles(styleName)
    On Error GoTo 0

    If ts Is Nothing Then
        Set ts = wb.TableStyles.Add(styleName)

        Dim azulHdr As Long: azulHdr = HexToRGB(HEX_AZUL)
        Dim celeste1 As Long: celeste1 = RGB(233, 240, 255) ' banda clara
        Dim celeste2 As Long: celeste2 = RGB(255, 255, 255) ' blanco

        ' Encabezado
        With ts.TableStyleElements(teHeaderRow)
            .Interior.Color = azulHdr
            .Font.Color = vbWhite
            .Font.Bold = True
        End With

        ' Bandas de filas
        With ts.TableStyleElements(teFirstRowStripe)
            .Interior.Color = celeste1
            .StripeSize = 1
        End With
        With ts.TableStyleElements(teSecondRowStripe)
            .Interior.Color = celeste2
            .StripeSize = 1
        End With

        ' Bordes suaves
        With ts.TableStyleElements(teWholeTable)
            .Borders(xlEdgeLeft).Color = RGB(210, 210, 210)
            .Borders(xlEdgeTop).Color = RGB(210, 210, 210)
            .Borders(xlEdgeRight).Color = RGB(210, 210, 210)
            .Borders(xlEdgeBottom).Color = RGB(210, 210, 210)
            .Borders(xlInsideHorizontal).Color = RGB(235, 235, 235)
            .Borders(xlInsideVertical).Color = RGB(235, 235, 235)
        End With
    End If

    EnsureExcelbotStyle = styleName
End Function

' ---- Formato columnas típicas SRI ----
Private Sub AjustarEsquemaSRI(ByVal lo As ListObject)
    On Error Resume Next
    If Not lo.ListColumns("RUC") Is Nothing Then lo.ListColumns("RUC").Range.NumberFormat = "@"
    If Not lo.ListColumns("CLAVE ACCESO") Is Nothing Then lo.ListColumns("CLAVE ACCESO").Range.NumberFormat = "@"
    If Not lo.ListColumns("CLAVE DE ACCESO") Is Nothing Then lo.ListColumns("CLAVE DE ACCESO").Range.NumberFormat = "@"

    Dim nm As Variant
    For Each nm In Array("FECHA EMISION", "FECHA EMISIÓN", "F. EMISION", "F. EMISIÓN", "FECHA")
        If Not lo.ListColumns(CStr(nm)) Is Nothing Then
            lo.ListColumns(CStr(nm)).Range.NumberFormat = "dd/mm/yyyy"
        End If
    Next nm
    On Error GoTo 0
End Sub

' --- Quita hipervínculos del encabezado de una tabla ---
Private Sub LimpiarHyperlinksEncabezado(ByVal lo As ListObject)
    On Error Resume Next
    If lo.HeaderRowRange Is Nothing Then Exit Sub
    If lo.HeaderRowRange.Hyperlinks.Count > 0 Then lo.HeaderRowRange.Hyperlinks.Delete
    lo.HeaderRowRange.Font.Underline = xlUnderlineStyleNone
    On Error GoTo 0
End Sub

' --- APLICAR FORMATO (versión que blinda el encabezado) ---
Public Sub AplicarFormatoTabla_EXCELBOT(ByVal lo As ListObject, _
    Optional ByVal styleName As String = "EXCELBOT_AzulCorp")

    If lo Is Nothing Then Exit Sub

    Dim wb As Workbook: Set wb = lo.Parent.Parent
    styleName = EnsureExcelbotStyle(wb, styleName)

    lo.TableStyle = styleName
    lo.ShowTableStyleRowStripes = True
    lo.ShowAutoFilterDropDown = True

    ' 1) Encabezado sin hipervínculos
    LimpiarHyperlinksEncabezado lo

    ' 2) Forzar encabezado corporativo (fondo azul + texto blanco)
    With lo.HeaderRowRange
        On Error Resume Next
        .FormatConditions.Delete
        On Error GoTo 0

        .Interior.Color = HexToRGB(HEX_AZUL)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Font.TintAndShade = 0

        Dim c As Range
        For Each c In .Cells
            c.Font.Color = RGB(255, 255, 255)
            c.Font.Underline = xlUnderlineStyleNone
        Next c
    End With

    AjustarEsquemaSRI lo
    lo.Range.EntireColumn.AutoFit
End Sub

' Aplica a todas las tablas del libro
Public Sub AplicarFormatoATodasLasTablas(ByVal wb As Workbook, _
    Optional ByVal styleName As String = "EXCELBOT_AzulCorp")

    Dim ws As Worksheet, lo As ListObject
    Dim _style As String: _style = EnsureExcelbotStyle(wb, styleName)
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            AplicarFormatoTabla_EXCELBOT lo, _style
        Next lo
    Next ws
End Sub


