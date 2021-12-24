# Control de Versiones

24/12/2021 - V2.0

        - Permite trabajar en dos fases. Una de generación de los documentos, que pueden ser firmados externamente, y una segunda que permite enviar los documentos generados.
       
        - Posibilidad de Retrasar el envío. Con dos modalidades, retrasar todos los envíos a una fecha, o que cada documento se envíe en una fecha específica (por ejemplo para felicitaciones de cumpleaños o caducidades de seguros)
       
        - Aumento de la Seguridad. La macro está autofirmada para que se pueda verificar que es original.
  
6/8/2021 - V1.3
                  
        -Cambio completo del etiquetado para nombrar archivos, correo, contraseña. Antes se utilizaban las etiquetas _archivo- çcorreoç y `contraseña`.  Esto generaba conflictos cuando el documento tenía alguno de esos caracteres. Se han cambiado las etiquetas de apertura por DBVNOMBRE, DBVCORREO y DBVCONTRA y la de cierre siempre es DBVFIN. Con esto ya se podrá usar cualquier caracter en los documentos.
                  
        - Se ha eliminado la alerta que aparecía cuando se terminaba la generación y se ha incorporado como mensaje en la consola.

        - Se ha cambiado el uso de CInt por un contador para evitar problemas en documentos con varias páginas (gracias Erik Rosas)

        - Incorporado pie de página en correos.

        - Actualizada la carpeta de Ejemplos a versión v5 (compatibles con la versión 1.3 de CombinarCorrespondencia)

18/4/2021 - V1.2

        - En esta versión se añade la posibilidad de que el cuerpo del mensaje sea HTML, permitiendo enriquecer enormemente el contenido de los mensajes que hasta ahora solo se podía con texto. 

        - Se añade también en la web un apartado de opiniones y empresas que utilizan la aplicación. 

5/2/2021 - V1.1

        - Esta versión permite generar archivos PDF/A (estos no se pueden proteger con contraseña por definición del estándar)

19/1/2021 - V1.0.1 

        - Pequeño bug al enviar correos con contraseña en formato .docx

18/1/2021 - V1.0

        - Se puede seleccionar tipo de documento generado .docx o .pdf

        - En los dos casos se puede poner contraseña a los documentos generados

        - Esta actualización permite usar la herramienta para dividir documento en varios manteniendo el formato.


17/12/2020 - V.0.0.4

        - Añadida selección de carpeta destino con FileDialog
                  
        - Añadida la posibilidad de adjuntar varios archivos comunes a los correos
                    
        - Iconos para Borrar consola y las nuevas opciones
15/12/2020 - V.0.0.3 
                     
        - Incluido Asunto en los mensajes.
                     
        - Añadidas Ayudas (Tips) en todas las funcionalidades.
                     
        - Mejora comentarios del código.                    
6/12/2020 - v.0.0.2  
                      
        - Mejora selección cuenta de envío.
5/12/2020 - v.0.0.1  
                      
        - Publicación de la primera versión.
                      
## Ir a Inicio
[Ir a inicio](README.md#combinar-correspondencia)

