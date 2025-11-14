Attribute VB_Name = "modDateHelpersUI"
'================= modDateHelpersUI =================
Option Explicit

' Acepta: DD/MM/YYYY, DD-MM-YYYY, YYYY-MM-DD
Public Function EX_TryParseFechaUI(ByVal s As String, ByRef d As Date) As Boolean
    On Error GoTo EH
    Dim t As String: t = Trim$(s)
    If Len(t) = 0 Then EX_TryParseFechaUI = False: Exit Function

    ' normaliza separadores
    t = Replace(t, ".", "/")
    t = Replace(t, "-", "/")

    Dim p() As String: p = Split(t, "/")
    If UBound(p) <> 2 Then GoTo EH

    Dim a As Integer, b As Integer, c As Integer
    a = val(p(0)): b = val(p(1)): c = val(p(2))

    If Len(p(0)) = 4 Then
        ' YYYY/MM/DD
        d = DateSerial(a, b, c)
    Else
        ' DD/MM/YYYY (Ecuador)
        d = DateSerial(c, b, a)
    End If
    EX_TryParseFechaUI = True
    Exit Function
EH:
    EX_TryParseFechaUI = False
End Function

' Alias “ISO” por compatibilidad con tu código del form
Public Function EX_TryParseISODate(ByVal s As String, ByRef d As Date) As Boolean
    EX_TryParseISODate = EX_TryParseFechaUI(s, d)
End Function
'====================================================

