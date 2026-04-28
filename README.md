IntuneDocKit
Herramienta PowerShell que genera automáticamente la documentación técnica completa de un tenant de Microsoft Intune en un archivo Word (.docx). Basado en el proyecto original de @matbe, modernizado y mejorado con automatización de App Registration, permisos y Admin Consent vía Microsoft Graph API.

Descripción
IntuneDocKit resuelve uno de los problemas más comunes en la administración de entornos Microsoft Intune: la documentación técnica. Documentar manualmente un tenant implica revisar decenas de configuraciones, copiar valores, organizar secciones y mantener el documento actualizado. Este proceso es lento, propenso a errores y difícil de estandarizar entre proyectos.
IntuneDocKit automatiza todo ese proceso. Se conecta al tenant usando Microsoft Graph API con una App Registration dedicada, extrae la configuración completa de Intune y genera un documento Word estructurado, listo para entregar al cliente o para uso interno como respaldo técnico.
Todas las operaciones son de solo lectura. El script nunca modifica ninguna configuración del tenant.

Qué documenta
IntuneDocKit extrae y documenta más de 20 secciones de Intune, entre ellas:
Políticas de cumplimiento de dispositivos, configuraciones de dispositivos, configuraciones de enrollment, perfiles de Windows Autopilot, aplicaciones móviles y políticas de protección, políticas de Acceso Condicional, roles RBAC y sus asignaciones, scripts de dispositivos, perfiles de Windows Quality Update y Feature Update, plantillas de mensajes de notificación, Scope Tags, categorías de dispositivos, tokens VPP de Apple, certificado Apple Push Notification y resumen general del inventario de dispositivos administrados.

Requisitos
PowerShell 5.1 en Windows. No es compatible con PowerShell Core ni PowerShell 7.
Cuenta de Azure con permisos para crear App Registrations y otorgar Admin Consent.
Cuenta de Intune con permisos de lectura para autenticarse durante la exportación.
Conexión a internet para instalar módulos desde PSGallery y autenticarse contra Microsoft Graph.

Instalación de módulos
El script instala automáticamente todos los módulos necesarios durante el primer paso de ejecución. No es necesario instalarlos manualmente. Los módulos requeridos son PSWriteWord, PSWriteHTML, ImportExcel, Microsoft.Graph.Authentication, Microsoft.Graph.Applications, Microsoft.Graph.Identity.SignIns, AzureAD versión 2.0.2.140 y MSAL.PS.
En algunos entornos puede requerirse ejecutar PowerShell como administrador para instalar módulos en el scope AllUsers.

Componentes del proyecto
IntuneDocKit.ps1 es la interfaz gráfica principal construida con Windows Forms. Desde aquí el usuario ingresa el Tenant ID, el nombre del cliente, el usuario UPN y selecciona el modo de ejecución. También incluye una consola integrada que muestra el progreso en tiempo real y una barra de pasos visual.
MasterExport-MEM.ps1 es el orquestador maestro. Coordina los cinco pasos del proceso y maneja la lógica de cada modo de ejecución. Recibe los parámetros desde la GUI o directamente desde la línea de comandos.
Export-MEMConfiguration.ps1 es el motor de exportación. Se encarga de autenticar al usuario mediante MSAL Device Code Flow, realizar todas las llamadas a Microsoft Graph API y generar el documento Word usando la plantilla incluida.
Export-MEMConfiguration.xml es el archivo de configuración. Almacena el Tenant ID, el Client ID de la App Registration y controla qué secciones de Intune se incluyen en la documentación. Este archivo se actualiza automáticamente durante el Paso 4.
MEMDocumentationTempl.docx es la plantilla base del documento Word generado. Define la estructura, estilos y portada del archivo de salida.
IntuneDocKit.exe es el ejecutable compilado para facilitar la distribución sin necesidad de ejecutar PowerShell directamente.

Modos de ejecución
Modo 1: Crear App Registration nueva (primera vez)
Este es el modo para la primera ejecución en un tenant nuevo. El script crea el App Registration en Azure, le asigna todos los permisos necesarios, otorga el Admin Consent de forma automática, actualiza el XML con el TenantId y el AppId obtenido, y luego procede a exportar la documentación completa.
powershell.\MasterExport-MEM.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -ClientName "Nombre del Cliente"
Modo 2: Usar App Registration existente
Para ejecuciones posteriores cuando la App Registration ya fue creada. El script omite los pasos de creación y configuración de Azure, lee el ClientId existente desde el XML y procede directamente a la exportación.
powershell.\MasterExport-MEM.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -ClientName "Nombre del Cliente" -SkipAppCreation
Modo 3: Solo crear App Registration
Útil cuando se quiere preparar el acceso al tenant pero realizar la exportación en otro momento. El script crea la App Registration, asigna permisos y actualiza el XML, pero no genera ningún documento.
powershell.\MasterExport-MEM.ps1 -TenantId "xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx" -OnlyCreateApp

Los cinco pasos del proceso
Paso 1: Módulos
Verifica que todos los módulos de PowerShell necesarios estén instalados. Si alguno falta, lo instala automáticamente desde PSGallery. También configura el proveedor NuGet y establece PSGallery como repositorio de confianza.
Paso 2: App Registration
Crea el App Registration llamado "IntuneDocKit" en el tenant de Azure usando Microsoft.Graph. Si ya existe una App con ese nombre, la reutiliza sin crear duplicados. También habilita los Public Client Flows y crea el Service Principal correspondiente. Este paso solo se ejecuta en los modos 1 y 3.
Paso 3: Permisos y Admin Consent
Asigna los permisos delegados necesarios al App Registration y otorga el Admin Consent de forma automática para todos los usuarios del tenant. Los permisos asignados son todos de solo lectura y cubren las áreas de dispositivos, aplicaciones, usuarios, grupos, directorio, políticas y auditoría de Intune. Este paso solo se ejecuta en los modos 1 y 3.
Paso 4: Actualización del XML
Escribe el Tenant ID y el Client ID obtenido en el archivo Export-MEMConfiguration.xml para que la exportación pueda autenticarse correctamente. Si se usa el modo 2, simplemente verifica que los valores ya existan en el XML.
Paso 5: Exportación
Lanza una sesión limpia de PowerShell para evitar conflictos entre los módulos de Microsoft.Graph y MSAL.PS. En esta sesión se autentica al usuario mediante el Device Code Flow de MSAL, se realizan todas las llamadas a Microsoft Graph API en versión Beta y se genera el archivo Word con la documentación completa del tenant.

Autenticación Device Code
Durante el Paso 5 el script solicita al usuario que se autentique. El flujo es el siguiente:
El script muestra un código de uso único en la consola. El usuario abre un browser y navega a https://microsoft.com/devicelogin. Ingresa el código que aparece en la consola. Se autentica con la cuenta de Intune que tenga permisos de lectura. Una vez completada la autenticación, el proceso continúa automáticamente.
La cuenta utilizada para autenticarse en este paso no necesita permisos de Azure. Solo requiere permisos de lectura en Intune.

Permisos de Microsoft Graph
Los siguientes permisos delegados se asignan automáticamente a la App Registration durante el Paso 3:
openid, profile, offline_access, User.Read, DeviceManagementApps.Read.All, DeviceManagementConfiguration.Read.All, DeviceManagementManagedDevices.Read.All, DeviceManagementRBAC.Read.All, DeviceManagementServiceConfig.Read.All, DeviceManagementScripts.Read.All, Directory.Read.All, User.Read.All, Group.Read.All, Organization.Read.All, Application.Read.All, Policy.Read.All, AuditLog.Read.All, Agreement.Read.All, CloudPC.Read.All y Policy.ReadWrite.ConditionalAccess.
Todos son permisos de solo lectura a excepción de Policy.ReadWrite.ConditionalAccess, el cual se requiere para leer las políticas de Acceso Condicional mediante la API delegada disponible en Graph.

Archivo de configuración XML
El archivo Export-MEMConfiguration.xml controla el comportamiento de la exportación. Los campos más relevantes son los siguientes:
Tenant almacena el Tenant ID del cliente. Se actualiza automáticamente en el Paso 4.
ClientId almacena el App ID de la App Registration creada. Se actualiza automáticamente en el Paso 4.
Document habilita la generación del archivo Word. El valor predeterminado es True.
DocumentName define el nombre del archivo generado. El patrón predeterminado es IntuneDoc_{ClientName}_{fecha}.docx.
Export habilita la exportación adicional en formato JSON por cada sección. El valor predeterminado es False.
ExportPath define la ruta donde se guardan los archivos exportados. El valor predeterminado es C:\temp\export_{fecha}.
graphApiVersion define la versión de la API de Graph. El valor predeterminado es Beta.
La sección Process del XML contiene un elemento por cada sección de Intune con valor True o False para incluirla o excluirla de la documentación.

Resultado
Al finalizar el proceso se genera un archivo Word en la ruta configurada, normalmente C:\temp. El archivo se abre automáticamente al completarse. El documento incluye una portada con el nombre del cliente y la fecha, y una sección por cada área de Intune configurada en el XML, con tablas de valores y propiedades extraídas directamente del tenant.

Uso desde la GUI
Al abrir IntuneDocKit.exe o ejecutar IntuneDocKit.ps1 se presenta una interfaz gráfica con los siguientes campos:
Nombre del cliente, que aparecerá en la portada del Word generado. Tenant ID, que identifica el tenant de Azure del cliente. Usuario UPN, que sirve como pista para la autenticación en el Paso 5. Modo de ejecución, donde se selecciona entre los tres modos descritos anteriormente.
La GUI también incluye una barra de progreso con los cinco pasos, una consola integrada que muestra el output en tiempo real con colores diferenciados por tipo de mensaje, y botones para copiar el comando de PowerShell equivalente, limpiar la consola y copiar el Device Code durante la autenticación.

Créditos
Este proyecto está basado en el trabajo original de matbe disponible en https://github.com/matbe/MEMDocumentAndExporter.
El proyecto original fue retomado, modernizado y extendido con automatización del proceso de App Registration, asignación de permisos, Admin Consent y una interfaz gráfica completa.
