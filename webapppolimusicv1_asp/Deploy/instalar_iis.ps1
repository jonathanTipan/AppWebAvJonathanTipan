<#
===============================================================================
 PoliMusic - Instalador para IIS
===============================================================================
 Copia la aplicacion ASP clasico a C:\inetpub\wwwroot\PoliMusic, habilita
 ASP clasico si hace falta, ajusta el tamano maximo de peticion (necesario
 para subir archivos MP3), concede permisos a la carpeta uploads y deja el
 sitio listo para abrir en el navegador.

 USO (desde una consola de PowerShell):
     powershell -ExecutionPolicy Bypass -File .\instalar_iis.ps1

 Si no se ejecuta como administrador, el script se relanza solo y pedira
 confirmacion de UAC.
===============================================================================
#>

[CmdletBinding()]
param(
    [string] $Origen  = (Join-Path (Split-Path -Parent $PSScriptRoot) 'PoliMusic'),
    [string] $Destino = 'C:\inetpub\wwwroot\PoliMusic',
    [string] $Sitio   = 'Default Web Site',
    [int]    $MaxMB   = 100
)

$ErrorActionPreference = 'Stop'

# -----------------------------------------------------------------------------
#  Elevacion automatica
# -----------------------------------------------------------------------------
function Test-Admin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $p  = New-Object Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-Admin)) {
    Write-Host ''
    Write-Host 'Se necesitan permisos de administrador para escribir en C:\inetpub\wwwroot.' -ForegroundColor Yellow
    Write-Host 'Se solicitara confirmacion (UAC)...' -ForegroundColor Yellow
    $argList = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', ('"{0}"' -f $PSCommandPath),
        '-Origen',  ('"{0}"' -f $Origen),
        '-Destino', ('"{0}"' -f $Destino),
        '-Sitio',   ('"{0}"' -f $Sitio),
        '-MaxMB',   $MaxMB
    )
    try {
        Start-Process -FilePath 'powershell.exe' -ArgumentList $argList -Verb RunAs -Wait
    } catch {
        Write-Host ''
        Write-Host 'No se pudo elevar automaticamente (no hay una sesion interactiva donde mostrar el aviso de UAC).' -ForegroundColor Red
        Write-Host ''
        Write-Host 'Haz lo siguiente:' -ForegroundColor Yellow
        Write-Host ('  1) Clic derecho en  ' + (Join-Path $PSScriptRoot 'desplegar_como_administrador.bat')) -ForegroundColor White
        Write-Host '  2) Elige "Ejecutar como administrador"' -ForegroundColor White
        Write-Host ''
        Write-Host 'O bien abre PowerShell como administrador y ejecuta:' -ForegroundColor Yellow
        Write-Host ('  powershell -ExecutionPolicy Bypass -File "{0}"' -f $PSCommandPath) -ForegroundColor Cyan
    }
    return
}

function Write-Step($text) { Write-Host ('[*] ' + $text) -ForegroundColor Cyan }
function Write-Ok($text)   { Write-Host ('    OK  ' + $text) -ForegroundColor Green }
function Write-Warn2($text){ Write-Host ('    !   ' + $text) -ForegroundColor Yellow }

Write-Host ''
Write-Host '==========================================================' -ForegroundColor DarkCyan
Write-Host '  PoliMusic - Instalacion en IIS' -ForegroundColor White
Write-Host '==========================================================' -ForegroundColor DarkCyan
Write-Host ''
Write-Host ('Origen : ' + $Origen)
Write-Host ('Destino: ' + $Destino)
Write-Host ''

# -----------------------------------------------------------------------------
#  1. Comprobaciones previas
# -----------------------------------------------------------------------------
Write-Step 'Comprobando el origen'
if (-not (Test-Path -LiteralPath $Origen)) {
    throw ('No se encontro la carpeta de la aplicacion: ' + $Origen)
}
if (-not (Test-Path -LiteralPath (Join-Path $Origen 'default.asp'))) {
    throw ('La carpeta de origen no contiene default.asp: ' + $Origen)
}
Write-Ok ('Aplicacion encontrada (' + (Get-ChildItem -LiteralPath $Origen -Recurse -File | Measure-Object).Count + ' archivos)')

Write-Step 'Comprobando el servicio IIS (W3SVC)'
$svc = Get-Service -Name 'W3SVC' -ErrorAction SilentlyContinue
if ($null -eq $svc) {
    throw 'IIS no esta instalado en este equipo. Instala el rol "Servidor web (IIS)".'
}
if ($svc.Status -ne 'Running') {
    Start-Service -Name 'W3SVC'
    Write-Ok 'Servicio W3SVC iniciado'
} else {
    Write-Ok 'Servicio W3SVC en ejecucion'
}

# -----------------------------------------------------------------------------
#  2. ASP clasico
# -----------------------------------------------------------------------------
Write-Step 'Verificando ASP clasico'
$aspOk = $false
try {
    $feat = Get-WindowsOptionalFeature -Online -FeatureName 'IIS-ASP' -ErrorAction Stop
    if ($feat.State -eq 'Enabled') {
        $aspOk = $true
        Write-Ok 'La caracteristica IIS-ASP ya esta habilitada'
    } else {
        Write-Warn2 'Habilitando la caracteristica IIS-ASP...'
        Enable-WindowsOptionalFeature -Online -FeatureName 'IIS-ASP' -All -NoRestart | Out-Null
        $aspOk = $true
        Write-Ok 'IIS-ASP habilitado'
    }
} catch {
    Write-Warn2 ('No se pudo consultar/habilitar IIS-ASP: ' + $_.Exception.Message)
}

if (-not (Test-Path -LiteralPath (Join-Path $env:windir 'System32\inetsrv\asp.dll'))) {
    Write-Warn2 'No se encontro asp.dll. Habilita manualmente: Caracteristicas de Windows > IIS > Desarrollo de aplicaciones > ASP.'
} else {
    Write-Ok 'asp.dll presente en el servidor'
}

# -----------------------------------------------------------------------------
#  3. Copia de archivos
# -----------------------------------------------------------------------------
Write-Step 'Copiando la aplicacion'
if (-not (Test-Path -LiteralPath (Split-Path -Parent $Destino))) {
    New-Item -ItemType Directory -Path (Split-Path -Parent $Destino) -Force | Out-Null
}

$roboArgs = @(
    $Origen, $Destino,
    '/MIR',
    '/NFL', '/NDL', '/NJH', '/NJS', '/NC', '/NS', '/NP',
    '/XF', '*.tmp',
    # No se replica la carpeta uploads: asi una nueva publicacion no borra los
    # archivos que hayan subido los usuarios. Las subcarpetas se crean despues.
    '/XD', 'uploads'
)
& robocopy @roboArgs | Out-Null
$rc = $LASTEXITCODE
if ($rc -ge 8) {
    throw ('robocopy devolvio el codigo ' + $rc)
}
Write-Ok ('Archivos copiados en ' + $Destino)

# Asegura que existan las carpetas de subida
$subidas = @('uploads', 'uploads\songFiles', 'uploads\userImages', 'uploads\playlistImages', 'uploads\albumImages', 'uploads\artistImages')
foreach ($s in $subidas) {
    $p = Join-Path $Destino $s
    if (-not (Test-Path -LiteralPath $p)) {
        New-Item -ItemType Directory -Path $p -Force | Out-Null
    }
}
Write-Ok 'Carpetas de subida verificadas'

# -----------------------------------------------------------------------------
#  4. Permisos
# -----------------------------------------------------------------------------
Write-Step 'Concediendo permisos a IIS_IUSRS'
& icacls $Destino /grant 'IIS_IUSRS:(OI)(CI)(RX)' /T /C | Out-Null
& icacls (Join-Path $Destino 'uploads') /grant 'IIS_IUSRS:(OI)(CI)(M)' /T /C | Out-Null
Write-Ok 'Lectura concedida en la aplicacion y escritura en uploads'

# -----------------------------------------------------------------------------
#  5. Ajustes de ASP en IIS
# -----------------------------------------------------------------------------
Write-Step 'Ajustando ASP clasico en IIS'
$appcmd = Join-Path $env:windir 'System32\inetsrv\appcmd.exe'
$maxBytes = $MaxMB * 1024 * 1024

if (Test-Path -LiteralPath $appcmd) {

    # Tamano maximo de peticion (necesario para subir MP3)
    try {
        & $appcmd set config -section:system.webServer/asp ("/limits.maxRequestEntityAllowed:{0}" -f $maxBytes) | Out-Null
        Write-Ok ('maxRequestEntityAllowed = {0} MB' -f $MaxMB)
    } catch {
        Write-Warn2 ('No se pudo ajustar maxRequestEntityAllowed: ' + $_.Exception.Message)
    }

    # Errores detallados mientras se pone a punto
    try {
        & $appcmd set config -section:system.webServer/asp '/scriptErrorSentToBrowser:true' | Out-Null
        Write-Ok 'scriptErrorSentToBrowser = true'
    } catch {
        Write-Warn2 'No se pudo activar scriptErrorSentToBrowser'
    }

    # Limite de filtrado de peticiones
    try {
        & $appcmd set config -section:system.webServer/security/requestFiltering '/requestLimits.maxAllowedContentLength:{0}' -f $maxBytes | Out-Null
        Write-Ok ('maxAllowedContentLength = {0} MB' -f $MaxMB)
    } catch {
        Write-Warn2 'No se pudo ajustar maxAllowedContentLength'
    }

    # Habilitar errores detallados en la aplicacion
    try {
        & $appcmd set config ($Sitio + '/PoliMusic') -section:system.webServer/httpErrors '/errorMode:Detailed' | Out-Null
        Write-Ok 'httpErrors en modo detallado'
    } catch {
        Write-Warn2 'No se pudo configurar httpErrors (no es critico)'
    }
} else {
    Write-Warn2 'No se encontro appcmd.exe; revisa los limites de ASP manualmente'
}

# -----------------------------------------------------------------------------
#  6. Estado del sitio
# -----------------------------------------------------------------------------
Write-Step 'Comprobando el sitio de IIS'
try {
    $sitios = & $appcmd list site 2>$null
    $existe = $false
    foreach ($s in $sitios) {
        if ($s -match [regex]::Escape($Sitio)) { $existe = $true }
    }
    if ($existe) {
        & $appcmd start site ("/site.name:{0}" -f $Sitio) 2>$null | Out-Null
        Write-Ok ('Sitio "' + $Sitio + '" disponible')
    } else {
        Write-Warn2 ('No se encontro el sitio "' + $Sitio + '". La carpeta ya esta en wwwroot, asi que se servira con el sitio predeterminado.')
    }
} catch {
    Write-Warn2 'No se pudo consultar el estado del sitio'
}

# -----------------------------------------------------------------------------
#  7. Resumen
# -----------------------------------------------------------------------------
$url = 'http://localhost/PoliMusic/'
Write-Host ''
Write-Host '==========================================================' -ForegroundColor DarkCyan
Write-Host '  Instalacion finalizada' -ForegroundColor Green
Write-Host '==========================================================' -ForegroundColor DarkCyan
Write-Host ''
Write-Host ('  URL      : ' + $url) -ForegroundColor White
Write-Host  '  Admin    : userAdmin / 12345678' -ForegroundColor White
Write-Host  '  Usuario  : userNormal / 12345678' -ForegroundColor White
Write-Host  '  Usuario  : epn.fis / 12345678' -ForegroundColor White
Write-Host ''
Write-Host '  Si la pagina no abre, comprueba que el servicio "World Wide Web"'
Write-Host '  este iniciado y que ASP clasico este habilitado.'
Write-Host ''

if ($Host.Name -eq 'ConsoleHost' -and $MyInvocation.InvocationName -ne '.' -and -not [Console]::IsInputRedirected) {
    Read-Host 'Pulsa ENTER para cerrar'
}
