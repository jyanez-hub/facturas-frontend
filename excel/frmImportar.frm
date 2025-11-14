VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImportar 
   Caption         =   "IMPORTAR XML A EXCEL"
   ClientHeight    =   7428
   ClientLeft      =   -528
   ClientTop       =   -3744
   ClientWidth     =   7668
   OleObjectBlob   =   "frmImportar.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmImportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mCentered As Boolean
Private mEditingDate As Boolean  ' evita recursión en Change


 Sub UserForm_Initialize()
    optFacturas.Value = True                 ' Por defecto: documentos (FC/NC/ND/LIQ)
    optDetalleDocumento.Value = True
    On Error Resume Next
    chkHipervinculos.Value = True
    chkSubcarpetas.Value = True
    On Error GoTo 0

    With lstCampos
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "220 pt;120 pt"
        .MultiSelect = fmMultiSelectMulti
    End With

    CargarCampos_UI
    
End Sub

Private Sub optFacturas_Click():    CargarCampos_UI: End Sub
Private Sub optRetenciones_Click(): CargarCampos_UI: End Sub
Private Sub optDetalleDocumento_Click(): CargarCampos_UI: End Sub
Private Sub optDetalleItems_Click():     CargarCampos_UI: End Sub

'Private Sub optDetalleDocumento_Click()
 '   CargarCamposDisponibles False
'End Sub

'Private Sub optDetalleItems_Click()
'    CargarCamposDisponibles True
'End Sub

' Rellena la lista con encabezados disponibles
Private Sub CargarCamposDisponibles(ByVal esDetalle As Boolean)
    Dim col As Collection, itm As Variant
    Set col = EX_GetHeadersDisponibles(esDetalle)

    lstCampos.Clear
    For Each itm In col
        lstCampos.AddItem itm(0)                 ' Campo
        lstCampos.List(lstCampos.ListCount - 1, 1) = itm(1)   ' Origen
    Next itm

    ' Preselección por defecto
    If esDetalle Then
        EX_PreseleccionarEnListBox lstCampos, EX_DefaultCampos_Detalle
    Else
        EX_PreseleccionarEnListBox lstCampos, EX_DefaultCampos_Documento
    End If
End Sub


Private Function ControlExists(frm As Object, ctlName As String) As Boolean
    On Error Resume Next
    ControlExists = Not frm.Controls(ctlName) Is Nothing
    On Error GoTo 0
End Function


' --- Grupo Campos ---
'Private Sub cmdPredeterminados_Click()
    ' Marca en lstCampos los campos base de IVA/RENTA que sueles usar
    ' Call SeleccionarCamposPredeterminados
'End Sub

Private Sub cmdLimpiarCampos_Click()
    Dim i As Long
    For i = 0 To Me.lstCampos.ListCount - 1
        Me.lstCampos.Selected(i) = False
    Next i
End Sub

' --- Fechas ---
Private Sub cmdLimpiarFechas_Click()
    Me.txtDesde.Text = "": Me.txtHasta.Text = ""
End Sub
' --- Ruta/XML ---
Private Sub cmdExaminar_Click()
    '***********************************************************************************************************
    'ESTE BLOQUE ABRE LA VENTANA GRANDE DE EXAMINAR
    'Dim d As FileDialog
    'Set d = Application.FileDialog(msoFileDialogFolderPicker)
    'With d
    '    .Title = "Elige la carpeta que contiene los XML del SRI"
    '    If .Show = -1 Then
    '        txtRuta.Text = .SelectedItems(1)
    '    End If
    'End With
    '***********************************************************************************************************
    Dim p As String
    p = PickFolderModern(Me.txtRuta.Text, "Selecciona la carpeta con los XML/PDF")
    If Len(p) > 0 Then Me.txtRuta.Text = p
    
End Sub

Private Sub cmdImportar_Click()
    Dim useDesde As Boolean, useHasta As Boolean
    Dim dDesde As Date, dHasta As Date
    Dim p As String: p = Trim$(Me.txtRuta.Text)
    If Len(p) = 0 Or dir$(p, vbDirectory) = vbNullString Then
        MsgBox "Elige una carpeta válida con XML.", vbExclamation: Exit Sub
    End If
    ' Campos seleccionados
    Dim campos As Variant: campos = UI_GetSelectedHeaders()
    If IsEmpty(campos) Then
        MsgBox "Selecciona al menos un campo en la lista.", vbExclamation: Exit Sub
    End If

    ' 1) Libro temporal (destino real del importador)
    Dim wbTemp As Workbook
    Set wbTemp = Application.Workbooks.Add(xlWBATWorksheet)

    ' 2) Redirigir escritura al temporal
    Set gTargetWb = wbTemp
'*************************************************************************************
'4) Bloquear la importación si las fechas no pasan
If Not ValidateDateRange(False) Then Exit Sub

' normalizar ambas si están OK (asegura DD/MM/YYYY)
If Len(Trim$(txtDesde.Text)) > 0 Then txtDesde.Text = Format$(CDate(txtDesde.Text), "dd/mm/yyyy")
If Len(Trim$(txtHasta.Text)) > 0 Then txtHasta.Text = Format$(CDate(txtHasta.Text), "dd/mm/yyyy")

'*************************************************************************************

    ' 3) Parámetros UI -> globals
    gUI_HasParams = True
    gUI_FolderPath = p
    gUI_IncludeSubfolders = (chkSubcarpetas.Value = True)
    gUI_FDesde = Trim$(txtDesde.Text)
    gUI_FHasta = Trim$(txtHasta.Text)
    If optRetenciones.Value Then
        gUI_FilterDocs = exRet
    Else
        gUI_FilterDocs = exDocs
    End If
    EX_SetRutaPDF p

    On Error GoTo EH
    
    useDesde = EX_TryParseFechaUI(Trim$(txtDesde.Text), dDesde)
    useHasta = EX_TryParseFechaUI(Trim$(txtHasta.Text), dHasta)
    
    If Not ValidateDateRange(False) Then Exit Sub
    
    ' 4) Importar al temporal
    Importar_XML_SRI

    ' 5) Hipervínculos (sobre temporal)
    AutoHyperlinks_AfterImport_ConRuta EX_GetRutaPDF

    ' 6) Exportar selección a libro final (nuevo .xlsm al guardar manualmente)
    EX_ExportarSeleccion_DesdeWB wbTemp, optRetenciones.Value, optDetalleItems.Value, campos

    ' 7) Cerrar temporal sin guardar
    Application.DisplayAlerts = False
    wbTemp.Close SaveChanges:=False
    Application.DisplayAlerts = True

    ' 8) Cerrar el formulario
    Unload Me
    Exit Sub

EH:
    Application.DisplayAlerts = True
    On Error Resume Next
    wbTemp.Close SaveChanges:=False
    On Error GoTo 0
    Set gTargetWb = Nothing
    gUI_HasParams = False
    MsgBox "IMPORTAR: " & Err.Description, vbExclamation
End Sub


Private Function SafeBool(chk As Object) As Boolean
    On Error Resume Next
    SafeBool = CBool(chk.Value)
    If Err.Number <> 0 Then SafeBool = False
    On Error GoTo 0
End Function


Private Function OnErrorFalse(chk As Object) As Boolean
    On Error Resume Next
    OnErrorFalse = chk.Value
    If Err.Number <> 0 Then OnErrorFalse = False
    On Error GoTo 0
End Function


Private Sub cmdCancelar_Click()
    Unload Me
End Sub


Private Sub UserForm_Activate()
    ' Centro una sola vez cuando el form ya tiene tamaño real
    If Not mCentered Then
        CenterToExcelOwner Me
        mCentered = True
    End If
End Sub

Private Sub CenterToExcelOwner(f As MSForms.UserForm)
    On Error Resume Next
    ' Coordenadas/medidas de la ventana de Excel en puntos
    Dim appL As Double, appT As Double, appW As Double, appH As Double
    appL = Application.Left
    appT = Application.Top
    appW = Application.Width
    appH = Application.Height

    ' Seguridad: si alguna medida vino 0 (caso extraño), usa CenterOwner estándar
    If appW <= 0 Or appH <= 0 Then
        frmImportar.StartUpPosition = 1 ' CenterOwner
        Exit Sub
    End If
'frmExportRecibidos.Left = 300
    frmImportar.Left = appL + (appW - frmImportar.Width) / 2
    frmImportar.Top = appT + (appH - frmImportar.Height) / 2

    ' Clamp por si Excel está parcialmente fuera de pantalla
    If frmImportar.Left < 0 Then frmImportar.Left = 0
    If frmImportar.Top < 0 Then frmImportar.Top = 0
    
    frmImportar.Height = 498
    frmImportar.Width = 601.8
End Sub

Private Sub CargarCampos_UI()
    Dim esDetalle As Boolean, esRet As Boolean
    esDetalle = optDetalleItems.Value
    esRet = optRetenciones.Value

    Dim col As Collection, itm As Variant
    Set col = EX_GetHeadersDisponibles(esDetalle, esRet)

    lstCampos.Clear
    For Each itm In col
        lstCampos.AddItem itm(0)
        lstCampos.List(lstCampos.ListCount - 1, 1) = itm(1)
    Next itm

    ' Preselecciones según combinación
    If esRet = False And esDetalle = False Then
        EX_PreseleccionarEnListBox lstCampos, EX_DefaultCampos_Documento
    ElseIf esRet = False And esDetalle = True Then
        EX_PreseleccionarEnListBox lstCampos, EX_DefaultCampos_Detalle
    ElseIf esRet = True And esDetalle = False Then
        EX_PreseleccionarEnListBox lstCampos, EX_DefaultCampos_RetDoc
    Else
        EX_PreseleccionarEnListBox lstCampos, EX_DefaultCampos_RetDet
    End If
End Sub

Private Function UI_GetSelectedHeaders() As Variant
    Dim tmp() As String, i As Long, n As Long
    ReDim tmp(0 To 0): n = -1
    For i = 0 To lstCampos.ListCount - 1
        If lstCampos.Selected(i) Then
            n = n + 1
            ReDim Preserve tmp(0 To n)
            tmp(n) = CStr(lstCampos.List(i, 0))
        End If
    Next i
    If n = -1 Then
        UI_GetSelectedHeaders = Empty
    Else
        UI_GetSelectedHeaders = tmp
    End If
End Function

'*******************************************************************************
'CONFIGURACION DE FILTRO DE FECHAS

Private Sub txtDesde_Enter()

    With txtDesde: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

Private Sub txtHasta_Enter()

  With txtHasta: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

' Permitir solo 0-9, "/", "-", Backspace
Private Sub txtDesde_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not IsDateKey(KeyAscii) Then KeyAscii = 0
End Sub

Private Sub txtHasta_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If Not IsDateKey(KeyAscii) Then KeyAscii = 0
End Sub

Private Function IsDateKey(ByVal k As Integer) As Boolean
    IsDateKey = (k = 8) Or (k >= 48 And k <= 57) Or (k = 45) Or (k = 47)
End Function

' Auto-insertar "/" mientras se escribe (DD/MM/YYYY)
Private Sub txtDesde_Change()
    AutoSlashDateBox txtDesde
End Sub

Private Sub txtHasta_Change()
    AutoSlashDateBox txtHasta
End Sub

Private Sub AutoSlashDateBox(ByRef tb As MSForms.TextBox)
    If mEditingDate Then Exit Sub
    mEditingDate = True
    Dim raw As String, d As String
    Dim i As Long, ch As String

    ' quitar todo lo que no sea dígito
    For i = 1 To Len(tb.Text)
        ch = Mid$(tb.Text, i, 1)
        If ch Like "#" Then d = d & ch
    Next i
    ' limitar a 8 dígitos (DDMMYYYY)
    If Len(d) > 8 Then d = Left$(d, 8)

    Dim out As String
    Select Case Len(d)
        Case 0, 1, 2
            out = d
        Case 3, 4
            out = Left$(d, 2) & "/" & Mid$(d, 3)
        Case Else ' 5 a 8
            out = Left$(d, 2) & "/" & Mid$(d, 3, 2) & "/" & Mid$(d, 5)
    End Select

    Dim prevLen As Long: prevLen = Len(tb.Text)
    tb.Text = out
    ' colocar el cursor al final (simple y estable)
    tb.SelStart = Len(tb.Text)
    tb.SelLength = 0
    mEditingDate = False
End Sub

'*****************************************************************************************************************************
'2) Formateo final al salir del control (DD/MM/YYYY) + validación inmediata

Private Sub txtDesde_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = Not NormalizeOneDateBox(txtDesde)
    If Not Cancel Then Cancel = Not ValidateDateRange(False)
End Sub

Private Sub txtHasta_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Cancel = Not NormalizeOneDateBox(txtHasta)
    If Not Cancel Then Cancel = Not ValidateDateRange(False)
End Sub

Private Function NormalizeOneDateBox(ByRef tb As MSForms.TextBox) As Boolean
    Dim s As String: s = Trim$(tb.Text)
    If Len(s) = 0 Then NormalizeOneDateBox = True: Exit Function ' vacío es válido (sin filtro)
    Dim d As Date
    If EX_TryParseFechaUI(s, d) Or EX_TryParseISODate(s, d) Then
        tb.Text = Format$(d, "dd/mm/yyyy")
        NormalizeOneDateBox = True
    Else
        Beep
        MsgBox "Fecha inválida: " & s & vbCrLf & "Usa el formato DD/MM/AAAA.", vbExclamation, "Validación de fecha"
        tb.SelStart = 0: tb.SelLength = Len(tb.Text)
        NormalizeOneDateBox = False
    End If
End Function

'*******************************************************************************************************************************
'3) Validación del rango (Desde = Hasta = Hoy)
Private Function ValidateDateRange(ByVal showOKMsg As Boolean) As Boolean
    Dim hasD As Boolean, hasH As Boolean
    Dim dD As Date, dH As Date, hoy As Date: hoy = Date

    hasD = EX_TryParseFechaUI(Trim$(txtDesde.Text), dD)
    hasH = EX_TryParseFechaUI(Trim$(txtHasta.Text), dH)

    ' Si ambos vacíos o uno solo, valida límites simples
    If Not hasD And Not hasH Then ValidateDateRange = True: Exit Function

    If hasH And dH > hoy Then
        MsgBox "'Hasta' no puede ser mayor que la fecha actual (" & Format$(hoy, "dd/mm/yyyy") & ").", vbExclamation
        txtHasta.SetFocus: txtHasta.SelStart = 0: txtHasta.SelLength = Len(txtHasta.Text)
        ValidateDateRange = False: Exit Function
    End If

    If hasD And hasH And dD > dH Then
        MsgBox "'Desde' no puede ser mayor que 'Hasta'." & vbCrLf & _
               "Desde: " & Format$(dD, "dd/mm/yyyy") & "   Hasta: " & Format$(dH, "dd/mm/yyyy"), vbExclamation
        txtDesde.SetFocus: txtDesde.SelStart = 0: txtDesde.SelLength = Len(txtDesde.Text)
        ValidateDateRange = False: Exit Function
    End If

    ValidateDateRange = True
    If showOKMsg Then
        MsgBox "Rango válido.", vbInformation
    End If
End Function

'*******************************************************************************************************************************
'B) “No se selecciona todo” al llegar a txtDesde / txtHasta

' Seleccionar todo al enfocarse (TAB o clic)
Private Sub txtDesde_GotFocus()
    With txtDesde: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub
Private Sub txtHasta_GotFocus()
    With txtHasta: .SelStart = 0: .SelLength = Len(.Text): End With
End Sub

' (Opcional) también mantener en Enter por compatibilidad
'Private Sub txtDesde_Enter()
'    txtDesde_GotFocus
'End Sub
'Private Sub txtHasta_Enter()
'    txtHasta_GotFocus
'End Sub

'*******************************************************************************************************************************

' --- PARSE UI: "DD/MM/YYYY" o "DD-MM-YYYY" (preferente) y "YYYY-MM-DD" (ISO) ---
Private Function EX_TryParseFechaUI(ByVal s As String, ByRef d As Date) As Boolean
    On Error GoTo bad
    Dim t As String: t = Trim$(s)
    If Len(t) = 0 Then EX_TryParseFechaUI = False: Exit Function

    Dim p() As String: p = Split(Replace$(t, "/", "-"), "-")
    If UBound(p) <> 2 Then GoTo bad

    Dim a As Integer, b As Integer, c As Integer
    a = val(p(0)): b = val(p(1)): c = val(p(2))

    If Len(p(0)) = 4 Then
        ' ISO: YYYY-MM-DD
        d = DateSerial(a, b, c)
    Else
        ' DMY: DD-MM-YYYY (Ecuador)
        d = DateSerial(c, b, a)
    End If
    EX_TryParseFechaUI = True
    Exit Function
bad:
    EX_TryParseFechaUI = False
End Function

'**************************************************************************************************************************
'PRESENT BASICO - BOTON PREDETERMINADOS
' === Arrays de preset Básico ===

' === Preset Básico — Facturas / NC / ND / Liq. Compra (0,5,8,12,15) ===
Private Function PRESET_BASICO_FACT() As Variant
    PRESET_BASICO_FACT = Array( _
        "tipo_comprobante", "nro_comprobante", "fecha_emision", _
        "ruc_emisor", "razon_social_emisor", _
        "ruc_ci_comprador", "razon_social_comprador", _
        "subtotal_iva_0", "subtotal_iva_5", "subtotal_iva_8", "subtotal_iva_12", "subtotal_iva_15", _
        "subtotal_no_objeto", "subtotal_exento", _
        "descuento", "iva_total", "valor_total" _
    )
End Function


Private Function PRESET_BASICO_RET() As Variant
    PRESET_BASICO_RET = Array( _
        "nro_comprobante", "fecha_emision", "periodo_fiscal", _
        "ruc_emisor", "razon_social_emisor", _
        "ruc_sujeto", "razon_social_sujeto", _
        "cod_ret_iva", "porc_ret_iva", "base_ret_iva", "valor_ret_iva", _
        "cod_ret_renta", "porc_ret_renta", "base_ret_renta", "valor_ret_renta", _
        "total_retenido" _
    )
End Function

' Marca en lstCampos los elementos del preset según el tipo (Facturas/Retenciones)
Private Sub SeleccionarPresetBasico()
    Dim target As Variant, i As Long, j As Long, want As String
    
    If optRetenciones.Value Then
        target = PRESET_BASICO_RET
    Else
        target = PRESET_BASICO_FACT
    End If

    ' Limpia selección actual
    For i = 0 To lstCampos.ListCount - 1
        lstCampos.Selected(i) = False
    Next i

    ' Selecciona coincidencias (case-insensitive exacto)
    For j = LBound(target) To UBound(target)
        want = LCase$(CStr(target(j)))
        For i = 0 To lstCampos.ListCount - 1
            If LCase$(CStr(lstCampos.List(i))) = want Then
                lstCampos.Selected(i) = True
                Exit For
            End If
        Next i
    Next j
End Sub


' Usa este handler en tu botón "PREDETERMINADOS"
Private Sub cmdPredeterminados_Click()
    SeleccionarPresetBasico
End Sub
