# Informes de Biopsias

El Sistema de Informes de Biopsias surge por la necesidad de exportar rapidamente valores de un fichero Excel sin formato a un fichero Word personalizado

## Empezando

Estas instrucciones le proporcionarán una copia del proyecto en funcionamiento en su máquina local para fines de desarrollo y prueba. Consulte implementación para obtener notas sobre cómo implementar el proyecto en un sistema en vivo.

### Requisitos previos

Esta versión del sistema funciona con PHP 8
La validación de los usuarios funciona contra un Directorio Activo y Grupos de Usuarios, configurable en el login del sistema
El sistema está basado en importar ficheros Excel y exportar a documentos de Word

### Instalación

Colocar la carpeta biopsias en un servidor web dentro de la red.
Si el servidor es basado en Linux es importante dale permisos 775 a la carpeta para que el sistema pueda manipular los ficheros Excel y Word.
Crear una base de datos en mysql llamada parte, además crear un usuario biopsias con la contraseña biopsias2012\*/
Una vez realizados estos pasos puede iniciar el sistema.

## Autor

MSc. Eric Enrique Sedeño Estrada<br>
Especialista informático Sucursal Camagüey<br>
Teléfono: 59938830<br>
Correo: informaticosmc.cmw@infomed.sld.cu
