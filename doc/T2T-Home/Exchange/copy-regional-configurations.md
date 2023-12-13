# Copy Regional configuration

### Requisitos

Tener instalados los modulos de:

- ExchangeOnlineManagement
- Usuarios con rol administrador de Exchange en ambos tenants de la migración T2T

### Información relevante a la ejecución

- Para el correcto funcionamiento del script es obligatorio definir los parámetros adecuados en el fichero :  `project_config.json`
    
    ```json
    {
      "params" : {
          "sourceTenantDomain" : "sourcedomain.net",
          "targetTenantDomain" : "targetdomain.com",
          "sourceConnectionPrefix": "SRC",
          "targetConnectionPrefix" : "TRG"
      }
    }
    ```
    
- El script cerrará las conexiones existentes a Exchange Online y creará nuevas conexiones incluyendo los prefijos definidos en:
    
    ```json
    {
       "sourceConnectionPrefix": "SRC",
       "targetConnectionPrefix" : "TRG"
    }
    ```
    

### Parámetros del script

| Param | Values | Description |
| --- | --- | --- |
| -Confirmation |  | Omite la enumeración de usuarios en modo lista en tiempo de ejecución del script |
| -DisplayCurrentStatus |  | Realiza una consulta del estado actual de la configuración regional del listado de usuarios cargados en el csv |

## Uso

1. Cargar todos los usuarios en el CSV de entrada: `Users.csv` respetando las cabeceras definidas:
    
    Ejemplo:
    
    ```powershell
    SourceUserPrincipalName,TargetUserPrincipalName
    flozada_src@365enespanol.com,flozada_trg@M365x59909532.onmicrosoft.com
    ```
    
2. Guardar
3. Lanzar el script: `CopyRegionalSettings.ps1`

## Ejemplos de ejecución

### 1.

```powershell
.\CopyRegionalSettings.ps1
```

Se copiaran todas las configuraciones regionales para los usuarios definidos en el CSV `‘Users.csv’` siguiendo la relación: SourcePrincipalName —> TargetUserPrincipalName

> Si un buzón no tiene configuración regional en origen, el script no copiara ninguna configuración regional en destino.
> 

### 2.

```powershell
.\CopyRegionalSettings.ps1 -Confirmation
```

Se procederá con la copia de configuraciones regionales sin realizar un output de información para confirmar el listado de usuarios que se ha introducido por CSV

### 3.

```powershell
.\CopyRegionalSettings.ps1 -DisplayCurrentStatus
```

No se copiara ningún tipo de configuración, simplemente se realiza una enumeración de la configuración actual de todos los usuarios introducidos por CSV mediante el comando de Exchange Online: `Get-MailboxRegionalConfiguration`