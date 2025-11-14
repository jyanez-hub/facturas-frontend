Attribute VB_Name = "modUI_Dialogs"
Option Explicit

' Diálogo moderno para escoger carpeta (con barra de navegación y cuadro de edición)
Public Function PickFolderModern(Optional ByVal initPath As String = "", _
                                 Optional ByVal title As String = "Selecciona la carpeta con los XML/PDF") As String
    Dim sh As Object, fld As Object
    Const BIF_RETURNONLYFSDIRS As Long = &H1
    Const BIF_NEWDIALOGSTYLE   As Long = &H40
    Const BIF_EDITBOX          As Long = &H10

    Set sh = CreateObject("Shell.Application")
    On Error Resume Next
    ' Si pasas initPath, abre allí
    If Len(initPath) > 0 Then
        Set fld = sh.BrowseForFolder(0, title, BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE Or BIF_EDITBOX, initPath)
    Else
        Set fld = sh.BrowseForFolder(0, title, BIF_RETURNONLYFSDIRS Or BIF_NEWDIALOGSTYLE Or BIF_EDITBOX)
    End If
    On Error GoTo 0

    If Not fld Is Nothing Then
        On Error Resume Next
        PickFolderModern = fld.Self.Path
        If Len(PickFolderModern) = 0 Then PickFolderModern = fld.Items.Item.Path
        On Error GoTo 0
    Else
        PickFolderModern = ""
    End If
End Function


