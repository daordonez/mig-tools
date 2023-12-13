# Generate ODM Files

### Requisitos

Ninguno

### Información relevante a la ejecución

- Este script generará los ficheros CSV para:
    - Tarea de Discovery
        
        Ejemplo de fichero salida:
        
        ```
        "UserPrincipalname","Type"
        "flozada_src@365enespanol.com","User"
        "cnaranjo_src@365enespanol.com","User"
        "mcano_src@365enespanol.com","User"
        ```
        
    - Tarea de Mapping
        
        ```
        "UserPrincipalName","UserPrincipalName"
        "flozada_src@365enespanol.com","flozada_trg@M365x59909532.onmicrosoft.com"
        ```
        
    - Tarea de creación de colección
        
        ```
        "UserPrincipalName","Collection"
        "flozada_src@365enespanol.com","Collection1_Contoso"
        "cnaranjo_src@365enespanol.com","Collection1_Contoso"
        "mcano_src@365enespanol.com","Collection1_Contoso"
        ```
        

## Parámetros del script

| Param | Values | Description |
| --- | --- | --- |
| -CollectionName (obligatorio) | Nombre de la colección | Establecerá el nombre de la colección y el nombre los ficheros a crear |

> No usar espacios en los nombres de las colecciones, preferiblemente usar guiones bajos en su lugar
> 

## Uso

1. Cargar los usuarios en el `CSV` de Entrada
    
    Ejemplo:
    
    ```powershell
    SourceUserPrincipalName,TargetUserPrincipalName
    flozada_src@365enespanol.com,flozada_trg@M365x59909532.onmicrosoft.com
    ```
    
2. Guardar
3. Ejecutar el script: `GenerateODMFiles.ps1`

## Ejemplos de ejecución

### 1.

```powershell
.\GenerateODMFiles.ps1 -CollectionName "DocuCollection"
```

Se generarán los ficheros necesario para las tareas de: Discovery, mapping y Collection, para los usuarios dentro del fichero: `‘Users.csv’`. La colección se llamará para este ejemplo “DocuCollection”