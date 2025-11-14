Attribute VB_Name = "modExportSeleccionWB"

Option Explicit

Public Sub EX_ExportarSeleccion_DesdeWB( _
    ByVal srcWb As Workbook, _
    ByVal esRet As Boolean, _
    ByVal esDetalle As Boolean, _
    ByVal headers As Variant)

    On Error GoTo EH

    Dim srcName As String
    srcName = IIf(esRet, IIf(esDetalle, "RetDet", "Retenciones"), IIf(esDetalle, "Detalle", "Facturas"))

    Dim wsSrc As Worksheet
    On Error Resume Next
    Set wsSrc = srcWb.Worksheets(srcName)
    On Error GoTo 0
    If wsSrc Is Nothing Then
        MsgBox "No existe la hoja origen: " & srcName, vbExclamation
        Exit Sub
    End If

    Dim lastRow As Long: lastRow = wsSrc.Cells(wsSrc.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No hay filas para exportar en " & srcName & ".", vbExclamation
        Exit Sub
    End If

    Dim wbOut As Workbook, wsDst As Worksheet
    Set wbOut = Application.Workbooks.Add(xlWBATWorksheet)
    Set wsDst = wbOut.Worksheets(1)
    On Error Resume Next: wsDst.name = "Reporte_" & srcName: On Error GoTo 0

    Dim i As Long, cSrc As Long, cOut As Long: cOut = 1
    For i = LBound(headers) To UBound(headers)
        cSrc = ColByHeaderLocal(wsSrc, CStr(headers(i)))
        If cSrc > 0 Then
            wsSrc.Range(wsSrc.Cells(1, cSrc), wsSrc.Cells(lastRow, cSrc)).Copy _
                Destination:=wsDst.Cells(1, cOut)
            cOut = cOut + 1
        End If
    Next i

    If cOut = 1 Then
        wbOut.Close SaveChanges:=False
        MsgBox "Ninguno de los campos seleccionados existe en " & srcName & ".", vbExclamation
        Exit Sub
    End If

    wsDst.Cells.EntireColumn.AutoFit
    wsDst.Rows(2).Select: ActiveWindow.FreezePanes = True
    wbOut.Activate
    Exit Sub
EH:
    MsgBox "Exportar selección: " & Err.Description, vbExclamation
End Sub

Private Function ColByHeaderLocal(ByVal ws As Worksheet, ByVal header As String) As Long
    Dim c As Range
    Set c = ws.Rows(1).Find(What:=header, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If c Is Nothing Then ColByHeaderLocal = 0 Else ColByHeaderLocal = c.Column
End Function


