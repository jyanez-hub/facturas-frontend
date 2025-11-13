Attribute VB_Name = "modConfig"
'================ modConfig ================
Option Explicit

Private mRutaPDF As String   ' <— variable interna (no pública)

Public Function GetRutaPDF() As String
    GetRutaPDF = mRutaPDF
End Function

Public Sub SetRutaPDF(ByVal ruta As String)
    If Len(ruta) = 0 Then Exit Sub
    If Right$(ruta, 1) <> "\" Then ruta = ruta & "\"
    mRutaPDF = ruta
End Sub

