Attribute VB_Name = "modRibbon"
'Callback for botonFacturas onAction
Sub RibbonX_ImpFacturas(control As IRibbonControl)
   frmImportar.Show
End Sub

'Callback for botonRenombrar onAction
Sub RibbonX_Renombrar(control As IRibbonControl)
   Renombrar_Comprobantes_Desde_XML
End Sub

'Callback for botonLimpiar onAction
Sub RibbonX_Limpiar(control As IRibbonControl)
    LimpiarDatos
End Sub
