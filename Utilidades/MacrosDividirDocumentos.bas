Attribute VB_Name = "DBVMacros"
Option Explicit
' Autor: David Bueno Vallejo
' twitter: @davidbuenov  linkedin: davidbueno
' Este codigo se encuentra actualizado en: https://github.com/davidbuenov/CombinarCorrespondencia/
' Se distribuye bajo licencia GPL v3.0 https://github.com/davidbuenov/CombinarCorrespondencia/blob/main/LICENSE

' Divide el documento actual en multiples documentos word
' Guarda los otros documentos en la misma carpeta anadiendo .1, .2, .3 a cada uno
' Para que funcione el documento debe estar dividido en secciones y se creara un
' nuevo documento para cada seccion.
' Importante: Al final del documento se debe insertar un salto de seccion
Public Sub DividirDocumento()
' Autor: David Bueno Vallejo
' twitter: @davidbuenov  linkedin: davidbueno
' Este codigo se encuentra actualizado en: https://github.com/davidbuenov/CombinarCorrespondencia/
' Se distribuye bajo licencia GPL v3.0 https://github.com/davidbuenov/CombinarCorrespondencia/blob/main/LICENSE


    Dim docPrincipal As String  'Automaticamente guardara el nombre del documento que se quiere dividir. Por ejemplo: Todosv4.docx
    Dim nombreDocs As String    'Sera en nombre que tendran las partes del documento
    Dim nuevoDoc As Document
    Dim seccion As Section
    Dim carpeta As String
    Dim numSeccion As Integer
    
    'carpeta = "d:\temp\generados"  'si queremos marcar ruta especifica
    carpeta = ActiveDocument.Path 'si queremos que se guarde en la misma carpeta
    On Error GoTo ControlErrores
    docPrincipal = ActiveDocument.Name
    nombreDocs = Mid(docPrincipal, 1, Len(docPrincipal) - 4) 'quita al nombre del documento la extension. Quedaria Todosv4.
    numSeccion = 1
    
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
    
    ' Recorre todas las secciones ignorando el ultimo salto de seccion
    For numSeccion = 1 To Documents(docPrincipal).Sections.Count - 1
        With Documents(docPrincipal).Sections(numSeccion)
            Documents(docPrincipal).Range(Start:=.Range.Start, End:=.Range.End - 1).Select
            Selection.Copy
            Set nuevoDoc = Documents.Add
            
            'Le asociamos el formato y los estilos del original
            With ActiveDocument
                .UpdateStylesOnOpen = True
                .AttachedTemplate = carpeta & "\" & nombreDocs & "dotx"
            End With
            Selection.PasteAndFormat (wdUseDestinationStylesRecovery)
            nuevoDoc.SaveAs2 FileName:=carpeta & "\" & nombreDocs & numSeccion & ".docx" 'genera path completo. Ej. d:\temp\Todosv4.1.docx
            nuevoDoc.Close
        End With
    Next numSeccion
    If numSeccion - 1 = 0 Then
        MsgBox "DBV. Recuerda que al final del documento debes insertar un salto de sección"
    Else
        MsgBox "Se han generado: " & numSeccion - 1 & " documentos"
    End If
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