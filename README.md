# Tenant to Tenant O365 migrations.

## Introducción

---

Este repositorio recoge los scripts utilizados para una migración T2T en entornos basados en tecnología Microsoft. También se recogen herramientas de conexión a la API de la herramienta de Quest : On Demand Migration.

## Dependencias

---

Los siguientes modules de Powershell deberán ser descargados de la PowerShell Gallery para poder hacer uso de los comandos utilizados:

- ExchangeOnlineManagement
- AzureAD
- SPOService
- MSOL Service
- MS Graph
    
    ```powershell
    Install-Module -Name ExchangeOnlineManagement
    Install-Module -Name AzureAD
    Install-Module -Name Microsoft.Online.SharePoint.PowerShell
    Install-Module -Name MSOnline
    Install-Module -Name Microsoft.Graph.Authentication
    ```
    

> Para módulos adicionales: [PowerShell Gallery](https://www.powershellgallery.com/)
> 

## Antes de empezar

---
Antes de lanzar cualquier de los scripts recogidos en este repositorio, es imprescindible revisar los ficheros de configuracion : ```project_config.json``` , así como el fichero de entrada de usuarios ```Users.csv```

## Inicio simple

---

Si nunca has trabajado previamente con repositorios git, deberás instalarlo en tu maquina ( [Descargar Git](https://git-scm.com/downloads))

Una vez instalado , procedemos a descargar una versión local de este repositorio:

1. Abrir consola de PowerShell
    
    > Se recomienda el uso de [Windows Terminal](https://aka.ms/terminal) para trabajar de manera más ágil
    > 
2. Lanzamos el siguiente comando:
    
    ```powershell
    git clone https://O365Migrations@dev.azure.com/O365Migrations/Unishare/_git/Unishare
    ```
    
3. Nos movemos dentro de la carpeta que se nos acaba de crear 
    
    ```powershell
    cd .\Unishare
    ```
    
4. Veremos los directorios que contienen los scripts
    
    ```powershell
    
    ls
    
    LastWriteTime       length Name
    -------------       ------ ----
    24/01/2023 10:46:45        CheckPreMigrationStatus
    24/01/2023 10:58:48        Diagrams
    25/01/2023 12:45:28        doc
    24/01/2023 10:46:46        EXO
    24/01/2023 10:46:48        ODM
    24/01/2023 10:46:51        OneDrive
    25/01/2023 13:58:33 51     .gitignore
    24/01/2023 12:33:22 2198   README.md
    ```
    

## Contribuir

---

Para un correcto mantenimiento del repositorio se recomienda realizar los cambios en ramas de desarrollo, y una vez validados , proceder a realizar una Pull Request sobre la rama `master`

### En caso de no haber trabajado con repositorios git

1. Descargar el repositorio en local
    
    ```powershell
    git clone https://O365Migrations@dev.azure.com/O365Migrations/Unishare/_git/Unishare
    ```
    
2. Moverse dentro de la carpeta que se descarga
3. Incluiremos el script que deseamos añadir en el repositorio en la carpeta de la tecnología correspondiente
4. Abrimos una terminal de Windows Powershell
5. Creamos una nueva rama. Ejemplo:
    
    ```powershell
    git checkout -b "dev/[NuevoScriptPararepo]"
    ```
    
    Esto creara una nueva rama en local con el nombre “dev/[NuevoScriptParerepo]”
    
6. Añadimos todos los ficheros al repositorio local:
    
    ```powershell
    git add .
    ```
    
7. Realizamos un commit con un mensaje que describa lo que estamos incluyendo en el repositorio:
    
    ```powershell
    git commit -m "Esto script se usa para conectarse al servicio de MSOL"
    ```
    
8. Enviamos el código al repositorio central (remoto)
    
    ```powershell
    git push origin dev/NuevoScriptPararepo
    ```
    
9. Comentamos el cambio con alguno de los contribuidores principales para incluir dicho script en la rama `master` de este repositorio

## Principales contribuidores

---

- Diego Ordoñez - Sr Analyst Modern Workplace  - Infra TC