# ODM Migration

## Script Disable Mailbox Access

### Requisitos

Tener instalados los módulos de:

- ExchangeOnlineManagement
- AzureAD
- Usuario administrador con roles de administrador de Exchange Online y AzureAD

### Información relevante a la ejecución

El script requiere tener activas conexiones contra los servicios:

- Exchange Online
- Azure AD

En caso de no tenerlas activas, se puede añadir el comando `-CreateConnection` para crear las conexiones necesarias.

### Parámetros del script

| Param | Values | Description |
| --- | --- | --- |
| -Action (Obligatorio) | Disable, Enable | Especificar si se desea habilitar o deshabilitar los clientes seleccionados |
| -AccesType (Obligatorio) | All, DesktopClients, MobileClients, WebClients | Indicar el tipo de acceso que se quiere habilitar o deshabilitar |
| -RefreshToken |  | Forzar refresco del token de usuario. Necesario para cliente pesado y clientes móviles |
| -CreateConnection |  | Fuerza la creación de las conexiones necesarias para que el script se ejecute correctamente |

## Uso

1. Cargar todos los usuarios sobre los que se quiere realizar la configuración en el fichero CSV: `Users.csv` en la raiz del script.
    
    Ejemplo:
    
    ```powershell
    SourceUserPrincipalName,TargetUserPrincipalName
    flozada_src@365enespanol.com,flozada_trg@M365x59909532.onmicrosoft.com
    ```
    
2. Guardar el CSV
3. Ejecutar el script: `DisableMailboxAccess.ps1` con los parámetros necesarios

## Ejemplos de ejecución

### 1.

```powershell
.\DisableMailboxAccess.ps1 -AccessType All -Action Disable -RefreshToken
```

Eel resultado sera que se deshabilitarán todos los protocolos y forzara que el usuario tenga que volver a iniciar sesión para los usuarios cargados en el fichero `‘Users.csv’`

### 2.

```powershell
.\DisableMailboxAccess.ps1 -AccessType All -Action Enable
```

Se activaran todos los protocolos para poder acceder al buzón del usuario (sin forzar inicio de sesión)

### 3.

```powershell
.\DisableMailboxAccess.ps1 -AccessType WebClients -Action Disable
```

Bloquear el acceso a usuarios por Outlook Web Access