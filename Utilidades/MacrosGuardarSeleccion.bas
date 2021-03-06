Attribute VB_Name = "DBVMacros"
Option Explicit
' Autor: David Bueno Vallejo
' twitter: @davidbuenov  linkedin: davidbueno
' Este codigo se encuentra actualizado en: https://github.com/davidbuenov/CombinarCorrespondencia/
' Se distribuye bajo licencia GPL v3.0 https://github.com/davidbuenov/CombinarCorrespondencia/blob/main/LICENSE

' Guarda la seleccion como un documento manteniendo los estilos
' Crea el archivo en la misma carpeta anadiendo .sel al nombre. prueba.docx-->prueba.sel.docx
Public Sub GuardarSeleccion()
' Autor: David Bueno Vallejo
' twitter: @davidbuenov  linkedin: davidbueno
' Este codigo se encuentra actualizado en: https://github.com/davidbuenov/CombinarCorrespondencia/
' Se distribuye bajo licencia GPL v3.0 https://github.com/davidbuenov/CombinarCorrespondencia/blob/main/LICENSE

    Dim docPrincipal As String  'Automaticamente guardara el nombre del documento que se quiere dividir. Por ejemplo: Todosv4.docx
    Dim nombreDocs As String    'Sera en nombre que tendran las partes del documento
    Dim nuevoDoc As Document
    Dim nombreDoc As String
    Dim carpeta As String
    
    'carpeta = "d:\temp\generados"  'si queremos marcar ruta especifica
    carpeta = ActiveDocument.Path 'si queremos que se guarde en la misma carpeta
    docPrincipal = ActiveDocument.Name
    On Error GoTo ControlErrores
    'quita al nombre del documento la extension. Quedaria Todosv4.
    nombreDocs = Mid(ActiveDocument.Name, 1, Len(ActiveDocument.Name) - 4)
        
    Selection.Copy
    
    ' hay que guardar el documento original como plantilla para asociarsela al nuevo documento
    ActiveDocument.SaveAs2 FileName:=carpeta & "\" & nombreDocs & "dotx", _
        FileFormat:=wdFormatXMLTemplate, LockComments:=False, _
        AddToRecentFiles:=False, ReadOnlyRecommended:=False, _
        EmbedTrueTypeFonts:=False, SaveNativePictureFormat:=False, SaveFormsData _
        :=False, SaveAsAOCELetter:=False, CompatibilityMode:=15
    'Al guardarlo como el documento original .docx se cierra. Asi lo abrimos de nuevo
    Documents.Open (carpeta & "\" & docPrincipal)
    Documents(docPrincipal).Activate
    Documents(carpeta & "\" & nombreDocs & "dotx").Close
    
    ' Creamos el documento
    Set nuevoDoc = Documents.Add
    
    'Le asociamos el formato y los estilos del original
    With ActiveDocument
        .UpdateStylesOnOpen = True
        .AttachedTemplate = carpeta & "\" & nombreDocs & "dotx"
    End With
    
    
    ' Selection.PasteAndFormat (wdFormatOriginalFormatting)
    Selection.PasteAndFormat (wdUseDestinationStylesRecovery)
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