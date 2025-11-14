Attribute VB_Name = "modRutaPDF"
Option Explicit
Private mRutaPDF As String

Public Sub EX_SetRutaPDF(ByVal p As String)
    If Len(p) = 0 Then
        mRutaPDF = ""
    Else
        mRutaPDF = IIf(Right$(p, 1) = "\", p, p & "\")
    End If
End Sub

Public Function EX_GetRutaPDF() As String
    EX_GetRutaPDF = mRutaPDF
End Function

