Attribute VB_Name = "modRenombrarComprobantes"
'====================  modRenombrarComprobantes  ====================
Option Explicit

'------------------- Config -------------------
Private Const AJUSTAR_CEROS As Boolean = True ' 3-3-9
Private Const LOG_SHEET As String = "LOG_Rename"
Private Const LOG_DIAGNOSTICO As Boolean = True

' ESC para cancelar (opcional, pero muy útil en árboles grandes)
#If VBA7 Then
    Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Long
#Else
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Long
#End If

Private Const VK_ESCAPE As Long = &H1B

' ===== Progreso & contadores (dejar una sola vez en el módulo) =====
Private gTotal As Long, gDone As Long

Private cOK_XML As Long
Private cOK_PDF As Long
Private cSKIP_XML As Long
Private cSKIP_PDF As Long
Private cNO_PDF As Long
Private cERR As Long
Private cFALTAN As Long



Private Function UserCancel() As Boolean
    UserCancel = (GetAsyncKeyState(VK_ESCAPE) <> 0)
End Function

' ===== Contadores =====
Private gTotal As Long, gDone As Long
Private cOK_XML As Long, cOK_PDF As Long, cSKIP_XML As Long, cSKIP_PDF As Long
Private cNO_PDF As Long, cERR As Long, cFALTAN As Long



'=================== Entrada principal ===================
Public Sub Renombrar_Comprobantes_Desde_XML()
    On Error GoTo EH

    Dim carpeta As String
    carpeta = SeleccionarCarpeta("Selecciona la carpeta con los XML/PDF")
    If Len(carpeta) = 0 Then Exit Sub

    ' preguntar si incluye subcarpetas
    Dim incSub As VbMsgBoxResult
    incSub = MsgBox("¿Incluir subcarpetas?", vbQuestion + vbYesNo, "Renombrar comprobantes")
    Dim incluirSubcarpetas As Boolean: incluirSubcarpetas = (incSub = vbYes)

    PrepararHojaLog
    If LOG_DIAGNOSTICO Then LogLinea carpeta, "", "DIAG", "Seleccionada. Subcarpetas=" & CStr(incluirSubcarpetas)

    Dim archivos As Collection: Set archivos = New Collection
    ListarArchivos carpeta, "*.xml", archivos, incluirSubcarpetas
    If archivos.Count = 0 Then
        LogLinea "", "", "SIN_CAMBIOS", "No se encontraron XML en la carpeta seleccionada."
        MsgBox "No hay XML en la carpeta.", vbInformation
        Exit Sub
    End If

      
    ' ======= INICIO PROGRESO =======
    Progress_Init archivos.Count

    Dim i As Long, cancelado As Boolean
    For i = 1 To archivos.Count
        If UserCancel Then cancelado = True: Exit For
        ProcesarUnXML CStr(archivos(i))
        Progress_Step
    Next i

    ' ======= FIN PROGRESO =======
    Progress_Done cancelado


    MsgBox "Proceso finalizado. Revisa la hoja '" & LOG_SHEET & "'.", vbInformation
    Exit Sub

EH:
    LogLinea "", "", "ERROR", "Error general: " & Err.Number & " - " & Err.Description
    MsgBox "Ocurrió un error: " & Err.Description, vbExclamation
End Sub

'==================== Procesamiento por XML ====================
Private Sub ProcesarUnXML(ByVal xmlPath As String)
    On Error GoTo EH

    Dim tipoCod As String, estab As String, pto As String, secu As String, fec As String, esRet As Boolean
    If Not RN_GetInfoFromXML(xmlPath, tipoCod, estab, pto, secu, fec, esRet) Then
        LogLinea xmlPath, "", "FALTAN_DATOS", "No se pudo determinar tipo/serie/fecha."
        Exit Sub
    End If

    If AJUSTAR_CEROS Then
        estab = RN_Pad(estab, 3)
        pto = RN_Pad(pto, 3)
        secu = RN_Pad(SoloDigitos(secu), 9)
    End If

    ' Base: PREF–ddmmyyyy–EEE–PPP–SSSSSSSSS
    Dim nro As String, nuevoBase As String
    nro = estab & "-" & pto & "-" & secu
    nuevoBase = RN_FileBaseFrom(tipoCod, nro, fec, esRet)

    ' Carpeta del archivo (seguro)
    Dim baseDir As String
    baseDir = PathDirName(xmlPath)

    ' Rutas nuevas
    Dim xmlNuevo As String, pdfViejo As String, pdfNuevo As String
    xmlNuevo = RN_UniquePath(baseDir, nuevoBase, ".xml")
    pdfViejo = RN_ChangeExt(xmlPath, ".pdf")
    pdfNuevo = RN_UniquePath(baseDir, nuevoBase, ".pdf")

    ' ===== Renombrar XML =====
    If LCase$(xmlPath) <> LCase$(xmlNuevo) Then
        Name xmlPath As xmlNuevo
        LogLinea xmlPath, xmlNuevo, "OK_XML", "XML renombrado."
    Else
        LogLinea xmlPath, xmlNuevo, "SKIP_XML", "Ya tiene el nombre esperado."
    End If

    ' ===== Renombrar PDF (si existe) =====
    If Len(dir$(pdfViejo, vbNormal)) > 0 Then
        If LCase$(pdfViejo) <> LCase$(pdfNuevo) Then
            Name pdfViejo As pdfNuevo
            LogLinea pdfViejo, pdfNuevo, "OK_PDF", "PDF renombrado."
        Else
            LogLinea pdfViejo, pdfNuevo, "SKIP_PDF", "Ya tiene el nombre esperado."
        End If
    Else
        LogLinea pdfViejo, "", "NO_PDF", "No se encontró pareja PDF."
    End If

    Exit Sub
EH:
    LogLinea xmlPath, "", "ERROR", "ProcesarUnXML: " & Err.Number & " - " & Err.Description
End Sub

' ==== Lee datos clave desde un XML del SRI (maneja <comprobante>) ====
Private Function RN_GetInfoFromXML(ByVal xmlPath As String, _
        ByRef tipoCod As String, ByRef estab As String, ByRef pto As String, ByRef secu As String, _
        ByRef fecEmi As String, ByRef esRet As Boolean) As Boolean
    On Error GoTo EH

    Dim d As Object, root As Object, infoTrib As Object, infoNode As Object
    Set d = CreateObject("MSXML2.DOMDocument.6.0")
    d.async = False: d.validateOnParse = False
    d.SetProperty "SelectionLanguage", "XPath"
    If Not d.Load(xmlPath) Then GoTo EH

    If d.DocumentElement Is Nothing Then GoTo EH
    Set root = d.DocumentElement

    ' Si viene envuelto, extrae <comprobante><![CDATA[...]]>
    If LCase$(root.nodeName) <> "factura" And LCase$(root.nodeName) <> "notacredito" _
       And LCase$(root.nodeName) <> "notadebito" And LCase$(root.nodeName) <> "comprobanteretencion" Then
        Dim inner As Object: Set inner = RN_ExtractInnerComprobante(d)
        If Not inner Is Nothing Then Set d = inner
    End If

    Set root = d.SelectSingleNode("/*[local-name()='factura' or local-name()='notaCredito' or local-name()='notaDebito' or local-name()='comprobanteRetencion']")
    If root Is Nothing Then Set root = d.SelectSingleNode(".//*[local-name()='factura' or local-name()='notaCredito' or local-name()='notaDebito' or local-name()='comprobanteRetencion']")
    If root Is Nothing Then GoTo EH

    Set infoTrib = root.SelectSingleNode("./*[local-name()='infoTributaria']")
    If infoTrib Is Nothing Then GoTo EH

    tipoCod = RN_SXL(infoTrib, "./*[local-name()='codDoc']")
    estab = RN_SXL(infoTrib, "./*[local-name()='estab']")
    pto = RN_SXL(infoTrib, "./*[local-name()='ptoEmi']")
    secu = RN_SXL(infoTrib, "./*[local-name()='secuencial']")

    Set infoNode = root.SelectSingleNode("./*[local-name()='infoFactura' or local-name()='infoNotaCredito' or local-name()='infoNotaDebito' or local-name()='infoCompRetencion']")
    fecEmi = RN_SXL(infoNode, "./*[local-name()='fechaEmision']")
    If Len(fecEmi) = 0 Then fecEmi = RN_SXL(infoNode, "./*[local-name()='fechaEmisionDocSustento']")

    esRet = (LCase$(root.nodeName) = "comprobanteRetencion") Or (Trim$(tipoCod) = "07")

    If LOG_DIAGNOSTICO Then
        LogLinea xmlPath, "", "DIAG", "tipo=" & tipoCod & " estab=" & estab & " pto=" & pto & " sec=" & secu & " fec=" & fecEmi
    End If

    RN_GetInfoFromXML = (Len(tipoCod) > 0 And Len(estab) > 0 And Len(pto) > 0 And Len(secu) > 0)
    Exit Function
EH:
    RN_GetInfoFromXML = False
End Function

Private Function RN_SXL(ByVal node As Object, ByVal xp As String) As String
    On Error GoTo EH
    Dim n As Object
    If node Is Nothing Then RN_SXL = "": Exit Function
    node.OwnerDocument.SetProperty "SelectionLanguage", "XPath"
    Set n = node.SelectSingleNode(xp)
    If n Is Nothing Then RN_SXL = "" Else RN_SXL = Trim$(CStr(n.Text))
    Exit Function
EH:
    RN_SXL = ""
End Function

' Extrae dom interno de <comprobante> si existe
Private Function RN_ExtractInnerComprobante(ByVal outer As Object) As Object
    On Error GoTo EH
    Dim c As Object: Set c = outer.SelectSingleNode(".//*[local-name()='comprobante']")
    If c Is Nothing Then Set RN_ExtractInnerComprobante = Nothing: Exit Function
    Dim txt As String: txt = Trim$(CStr(c.Text))
    If Len(txt) = 0 Then Set RN_ExtractInnerComprobante = Nothing: Exit Function

    Dim d2 As Object: Set d2 = CreateObject("MSXML2.DOMDocument.6.0")
    d2.async = False: d2.validateOnParse = False
    d2.SetProperty "SelectionLanguage", "XPath"
    If d2.LoadXML(txt) Then Set RN_ExtractInnerComprobante = d2 Else Set RN_ExtractInnerComprobante = Nothing
    Exit Function
EH:
    Set RN_ExtractInnerComprobante = Nothing
End Function

' EEE-PPP-SSSSSSSSS helpers
Private Function RN_Pad(ByVal s As String, ByVal n As Long) As String
    s = Trim$(CStr(s))
    If Len(s) >= n Then RN_Pad = Right$(s, n) Else RN_Pad = String$(n - Len(s), "0") & s
End Function
Private Function SoloDigitos(ByVal s As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then out = out & ch
    Next i
    SoloDigitos = out
End Function

' === Construye nombre base con FECHA ddmmyyyy ===
' PREF–ddmmyyyy–EEE–PPP–SSSSSSSSS
Private Function RN_FileBaseFrom(ByVal tipoCod As String, _
                                 ByVal nro As String, _
                                 ByVal fecTxt As String, _
                                 ByVal esRet As Boolean) As String
    Dim pref As String
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

    RN_FileBaseFrom = pref & "-" & RN_DateTokenDMY(fecTxt) & "-" & Replace$(nro, " ", "")
End Function

' --- Parseo robusto de fecha a ddmmyyyy ---
Private Function RN_DateTokenDMY(ByVal s As String) As String
    Dim d As Date
    If RN_TryParseFecha(s, d) Then
        RN_DateTokenDMY = Format$(d, "ddmmyyyy")
    Else
        RN_DateTokenDMY = ""
    End If
End Function

Private Function RN_TryParseFecha(ByVal s As String, ByRef d As Date) As Boolean
    On Error GoTo EH
    Dim t As String: t = Trim$(s)
    If Len(t) = 0 Then Exit Function
    Dim p() As String
    t = Replace$(t, "/", "-")
    p = Split(t, "-")
    If UBound(p) <> 2 Then Exit Function

    Dim a As Integer, b As Integer, c As Integer
    a = val(p(0)): b = val(p(1)): c = val(p(2))
    If Len(p(0)) = 4 Then
        d = DateSerial(a, b, c)          ' YYYY-MM-DD
    Else
        d = DateSerial(c, b, a)          ' DD-MM-YYYY
    End If
    RN_TryParseFecha = True
    Exit Function
EH:
    RN_TryParseFecha = False
End Function

'================== Recorrer carpeta (FSO, robusto) ==================
Private Sub ListarArchivos(ByVal carpeta As String, ByVal patron As String, _
                           ByRef col As Collection, ByVal subcarpetas As Boolean)
    On Error GoTo Fin

    Dim fso As Object, f As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Right$(carpeta, 1) <> "\" Then carpeta = carpeta & "\"
    If Not fso.FolderExists(carpeta) Then Exit Sub

    Set f = fso.GetFolder(carpeta)

    ' Recorre archivos de la carpeta actual (*.xml)
    Dim ar As Object
    For Each ar In f.files
        ' más rápido/seguro que Like: solo nos interesa *.xml
        If LCase$(Right$(ar.name, 4)) = ".xml" Then
            col.Add ar.Path
        End If
        If UserCancel Then Exit Sub
    Next ar

    If Not subcarpetas Then GoTo Fin

    ' Recorre subcarpetas (recursivo, pero sin Dir)
    Dim subf As Object
    For Each subf In f.SubFolders
        ' Saltar reparse points o carpetas con permisos problemáticos
        On Error Resume Next
        Dim attr As Long: attr = fso.GetFolder(subf.Path).Attributes
        If Err.Number <> 0 Then
            Err.Clear
        Else
            ' 1024 = ReparsePoint (puntos de montaje, OneDrive, symlinks)
            If (attr And 1024) = 0 Then
                On Error GoTo Fin
                ListarArchivos subf.Path, patron, col, True
                On Error GoTo Fin
            End If
        End If
        On Error GoTo Fin
        DoEvents
        If UserCancel Then Exit Sub
    Next subf

Fin:
    ' Nada: salimos silenciosos si hay error de permisos o similares
End Sub


'================== LOG ==================
Private Sub PrepararHojaLog()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.Count))
        ws.name = LOG_SHEET
        ws.Range("A1:D1").Value = Array("Archivo_Origen", "Archivo_Destino", "Estado", "Mensaje")
        ws.Range("A1:D1").Font.Bold = True
        ws.Columns("A:D").ColumnWidth = 70
    Else
        ' conservar encabezados, limpiar datos antiguos (opcional)
        If ws.UsedRange.Rows.Count > 1 Then ws.Rows("2:" & ws.Rows.Count).ClearContents
    End If
End Sub

Private Sub LogLinea(ByVal origen As String, ByVal destino As String, _
                     ByVal estado As String, ByVal mensaje As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(LOG_SHEET)
    Dim r As Long: r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(r, 1).Value = origen
    ws.Cells(r, 2).Value = destino
    ws.Cells(r, 3).Value = estado
    ws.Cells(r, 4).Value = mensaje

    ' <-- NUEVO: acumula contadores según el estado
    Tally estado
End Sub

'================== Utilidades de ruta ==================
Private Function PathDirName(ByVal fullPath As String) As String
    Dim p As Long: p = InStrRev(fullPath, "\")
    If p > 0 Then PathDirName = Left$(fullPath, p - 1) Else PathDirName = CurDir$
End Function

Private Function RN_ChangeExt(ByVal fullPath As String, ByVal newExt As String) As String
    Dim p As Long: p = InStrRev(fullPath, ".")
    If p = 0 Then RN_ChangeExt = fullPath & newExt _
               Else RN_ChangeExt = Left$(fullPath, p - 1) & newExt
End Function

Private Function RN_UniquePath(ByVal baseDir As String, ByVal base As String, ByVal ext As String) As String
    Dim n As Long, pth As String
    pth = baseDir & IIf(Right$(baseDir, 1) = "\", "", "\") & base & ext
    If Len(dir$(pth, vbNormal)) = 0 Then RN_UniquePath = pth: Exit Function
    n = 1
    Do
        pth = baseDir & IIf(Right$(baseDir, 1) = "\", "", "\") & base & " (" & n & ")" & ext
        If Len(dir$(pth, vbNormal)) = 0 Then RN_UniquePath = pth: Exit Function
        n = n + 1
    Loop
End Function

'================== UI carpeta (simple) ==================
Private Function SeleccionarCarpeta(ByVal titulo As String) As String
    On Error Resume Next
    Dim sh As Object: Set sh = CreateObject("Shell.Application")
    Dim f As Object: Set f = sh.BrowseForFolder(0, titulo, 0, 0)
    If Not f Is Nothing Then
        SeleccionarCarpeta = f.Self.Path
        If Right$(SeleccionarCarpeta, 1) <> "\" Then SeleccionarCarpeta = SeleccionarCarpeta & "\"
    End If
End Function

Private Sub Progress_Init(ByVal totalItems As Long)
    gTotal = totalItems: gDone = 0
    cOK_XML = 0: cOK_PDF = 0: cSKIP_XML = 0: cSKIP_PDF = 0
    cNO_PDF = 0: cERR = 0: cFALTAN = 0
    On Error Resume Next
    frmProgress.InitUI
    frmProgress.Show vbModeless
    Application.StatusBar = "Renombrando… 0% (0/" & gTotal & ")"
    DoEvents
End Sub

Private Sub Progress_Step()
    gDone = gDone + 1
    Dim pct As Double: If gTotal > 0 Then pct = gDone / gTotal
    Dim msg As String
    msg = "Procesando " & gDone & " / " & gTotal & _
          "   |  OK_XML:" & cOK_XML & "  OK_PDF:" & cOK_PDF & _
          "  SKIP:" & (cSKIP_XML + cSKIP_PDF) & _
          "  NO_PDF:" & cNO_PDF & "  ERR:" & cERR
    On Error Resume Next
    frmProgress.UpdateUI pct, msg
    Application.StatusBar = "Renombrando… " & Format(pct, "0%") & _
                            "  (" & gDone & "/" & gTotal & ")"
    DoEvents
End Sub

Private Sub Progress_Done(Optional ByVal canceled As Boolean = False)
    On Error Resume Next
    Unload frmProgress
    Application.StatusBar = False
    Dim resumen As String
    resumen = IIf(canceled, "Proceso CANCELADO por el usuario." & vbCrLf & vbCrLf, "") & _
              "XML renombrados: " & cOK_XML & vbCrLf & _
              "PDF renombrados: " & cOK_PDF & vbCrLf & _
              "Omitidos XML: " & cSKIP_XML & "   |   Omitidos PDF: " & cSKIP_PDF & vbCrLf & _
              "Sin PDF: " & cNO_PDF & vbCrLf & _
              "Errores: " & cERR
    MsgBox resumen, vbInformation, "Resumen de renombrado"
End Sub


' ====== SUMARIZADOR DE ESTADOS (2.3) ======
Private Sub Tally(ByVal estado As String)
    Select Case UCase$(Trim$(estado))
        Case "OK_XML":      cOK_XML = cOK_XML + 1
        Case "OK_PDF":      cOK_PDF = cOK_PDF + 1
        Case "SKIP_XML":    cSKIP_XML = cSKIP_XML + 1
        Case "SKIP_PDF":    cSKIP_PDF = cSKIP_PDF + 1
        Case "NO_PDF":      cNO_PDF = cNO_PDF + 1
        Case "FALTAN_DATOS": cFALTAN = cFALTAN + 1
        Case "ERROR", "ERROR_REN_XML", "ERROR_REN_PDF": cERR = cERR + 1
        Case Else
            ' DIAG / SIN_CAMBIOS no suman
    End Select
End Sub

