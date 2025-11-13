Attribute VB_Name = "modNombresArchivos"
Option Explicit

' Prefijo según tipo SRI
Public Function EX_PrefijoTipo(ByVal tipoCod As String, ByVal esRet As Boolean) As String
    If esRet Then
        EX_PrefijoTipo = "CR"                 ' Comprobante de Retención
    Else
        Select Case Trim$(tipoCod)
            Case "01": EX_PrefijoTipo = "FC"  ' Factura
            Case "04": EX_PrefijoTipo = "NC"  ' Nota de Crédito
            Case "05": EX_PrefijoTipo = "ND"  ' Nota de Débito
            Case "03": EX_PrefijoTipo = "LC"  ' Liquidación de compra (si aplica)
            Case Else: EX_PrefijoTipo = "CB"  ' Genérico
        End Select
    End If
End Function

' Devuelve MMDDYYYY a partir de "DD/MM/YYYY", "DD-MM-YYYY" o "YYYY-MM-DD"
Public Function FechaToken_MDY(ByVal s As String) As String
    Dim d As Date
    If TryParseFechaEmision(s, d) Then
        FechaToken_MDY = Format$(d, "mmddyyyy")
    Else
        FechaToken_MDY = ""
    End If
End Function

' Base de nombre: PREFIJO-MMDDYYYY-EEE-PPP-SSSSSSSSS (sin extensión)
Public Function EX_FileBaseFrom(ByVal tipoCod As String, ByVal nro As String, ByVal fecTxt As String, ByVal esRet As Boolean) As String
    Dim pref As String, serie As String, base As String
    pref = EX_PrefijoTipo(tipoCod, esRet)
    serie = Replace$(nro, " ", "")
    base = pref & "-" & FechaToken_MDY(fecTxt) & "-" & serie
    base = Replace(base, "--", "-")                      ' sanea dobles guiones
    If Right$(base, 1) = "-" Then base = Left$(base, Len(base) - 1)
    EX_FileBaseFrom = SanitizeFileName(base)
End Function

' Construye ruta del PDF junto al XML
Public Function EX_BuildOutputPath(ByVal xmlPath As String, ByVal tipoCod As String, _
                                   ByVal nro As String, ByVal fecTxt As String, ByVal esRet As Boolean) As String
    EX_BuildOutputPath = Left$(xmlPath, InStrRev(xmlPath, "\") - 1) & "\" & _
                         EX_FileBaseFrom(tipoCod, nro, fecTxt, esRet) & ".pdf"
End Function


