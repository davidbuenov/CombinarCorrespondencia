Attribute VB_Name = "DBVMacros"
Option Explicit
' Autor: David Bueno Vallejo
' twitter: @davidbuenov  linkedin: davidbueno
' Este codigo se encuentra actualizado en: https://github.com/davidbuenov/CombinarCorrespondencia/
' Se distribuye bajo licencia GPL v3.0 https://github.com/davidbuenov/CombinarCorrespondencia/blob/main/LICENSE


' Guarda la selección como un documento manteniendo los estilos
' Crea el archivo en la misma carpeta anadiendo .sel al nombre. prueba.docx-->prueba.sel.docx
Public Sub GuardarSeleccion()
' Autor: David Bueno Vallejo
' twitter: @davidbuenov  linkedin: davidbueno
' Este codigo se encuentra actualizado en: https://github.com/davidbuenov/CombinarCorrespondencia/
' Se distribuye bajo licencia GPL v3.0 https://github.com/davidbuenov/CombinarCorrespondencia/blob/main/LICENSE

    Dim nombreDocs As String    'Sera en nombre que tendran las partes del documento
    Dim nuevoDoc As Document
    Dim nombreDoc As String
    Dim carpeta As String
    
    'carpeta = "d:\temp\generados"  'si queremos marcar ruta especifica
    carpeta = ActiveDocument.Path 'si queremos que se guarde en la misma carpeta
    
    On Error GoTo ControlErrores
    'quita al nombre del documento la extension. Quedaria Todosv4.
    nombreDocs = Mid(ActiveDocument.Name, 1, Len(ActiveDocument.Name) - 4)
    
    Selection.Copy
    
    Set nuevoDoc = Documents.Add
    Selection.PasteAndFormat (wdFormatOriginalFormatting)
    nuevoDoc.SaveAs2 FileName:=carpeta & "\" & nombreDocs & "sel" & ".docx"
    nuevoDoc.Close
    Exit Sub ' Salida Normal
ControlErrores:
     Dim mensaje As String
     Select Case Err.Number
        Case 4605
            mensaje = "EDBV. Pero que quiere que guarde si no has seleccionado nada :-)"
            
        Case Else
            mensaje = "EDBV. Se ha producido el error: " & Err.Number & " - " & Err.Description
    End Select
    MsgBox mensaje
End Sub