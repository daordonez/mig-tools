# Get One Drive Status

## Requisitos

Tener instalado el módulo de SharePoint Online:

- Microsoft.Online.SharePoint.PowerShell
- Disponer de al menos el rol de lector global del Tenant

Haber configurado el fichero: `project_config.json` con los parametros adecuados a los tenants con los que se está trabajando.

```powershell
{
  "params" : {
      "sharePoint":{
        "logsFileName": "OneDriveStatus.log",
        "adminUrl": "https://targetdomain-admin.sharepoint.com",
        "baseODUrl": "https://targetdomain-my.sharepoint.com/personal"
      }
  }
}
```

### Información relevante a la ejecución

- El solo realiza consultas, en ningún casó realiza ninguna configuración
- Es posible realizar una conexión limpia o bien reutilizar una conexión previa que se haya hecho de formal manual al servicio de SharePoint Online
- Cada ejecución dejará un fichero CSV en la carpeta : `export` ubicada en la raíz del script

## Parámetros del script

| Param | Values | Description |
| --- | --- | --- |
| -CreateConnection |  | Creará una conexión contra el tenant destino (indicado en el fichero: project_config.json) |

## Uso

1. Cargar los usuarios en el CSV: `Users.csv`
    
    Ejemplo:
    
    ```powershell
    UserPrincipalName
    user.name@domain.com
    ```
    
2. Guardar
3. Ejecutar el script: `GetOneDriveStatus.ps1`

## Ejemplos de ejecución

### 1.

```powershell
.\GetOneDriveStatus.ps1 -CreateConnection
```

El script solicitará credenciales para crear la conexión y a continuación obtendra un reporte de los usuarios cargados en el fichero `‘Users.csv’`

### 2.

```powershell
.\GetOneDriveStatus.ps1
```

Se generará un reporte con los usuarios cargados en el fichero `‘Users.csv’`

### 

## Ejemplo salida por consola de la consulta

```powershell
2023/01/25 16:46:45 INFO Script Start
2023/01/25 16:46:45 INFO Total users:3
2023/01/25 16:46:45 INFO Getting user information
        Getting:flozada_trg@M365x59909532.onmicrosoft.com        OK Active
        Getting:cnaranjo_trg@M365x59909532.onmicrosoft.com       OK Active
        Getting:mcano_trg@M365x59909532.onmicrosoft.com  OK Active
2023/01/25 16:46:47 INFO Summary: Total:3,Active:,3,Pending:0
2023/01/25 16:46:47 INFO Exporting results. PATH:C:\GetOneDriveStatus\export

DisplayName    UserPrincipalName                          Status UserOneDriveSite
-----------    -----------------                          ------ ----------------
Fazzio Lozada  flozada_trg@M365x59909532.onmicrosoft.com  Active https://M365x59909532-my.sharepoint.com/personal/flozada_trg_M365x59909532_onmicrosoft_com
Celest Naranjo cnaranjo_trg@M365x59909532.onmicrosoft.com Active https://M365x59909532-my.sharepoint.com/personal/cnaranjo_trg_M365x59909532_onmicrosoft_com
Mirta Cano     mcano_trg@M365x59909532.onmicrosoft.com    Active https://M365x59909532-my.sharepoint.com/personal/mcano_trg_M365x59909532_onmicrosoft_com
```