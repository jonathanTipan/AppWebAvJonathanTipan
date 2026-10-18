# PoliMusic — notas permanentes del proyecto

Clon de Spotify en **ASP clásico (VBScript)** + SQL Server, con tema **azul** (no verde).
Trabajo de la asignatura; el usuario es hispanohablante y espera comentarios en español.

## Convenciones del proyecto
- **Raíz web**: `PoliMusic/`. Todo lo que no sea la app (scripts SQL, tests, despliegue) va
  fuera de `PoliMusic/` para no publicarlo.
- **Inclusión única**: `inc/boot.asp` encadena config → db → util → icons → partials → upload → auth.
  SSI no tiene search path, por eso existe ese único punto de entrada.
- **UTF-8**: `<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>` y `Response.CodePage = 65001`.
- **SQL siempre parametrizado** con `@p0, @p1…` en el texto (nunca concatenar valores).
  `MakeCmd` los traduce a `?` posicionales; no cambiar los `@pN` de las consultas.
- **Nombres de variables**: VBScript no distingue mayúsculas y las funciones son globales.
  Evitar nombres de una letra que choquen con utilidades: `N`, `H`, `F`, `C`, `P`, `Q`.
  Y nunca declarar una variable llamada `inStr` (tapa `InStr`).
- **No existe `Optional`** en VBScript: las funciones tienen aridad fija.
- Los `Dim` de nivel de página van **al principio** del bloque `<% %>`.

## Cuentas y roles
- Roles: `0` = usuario normal, `1` = administrador (`TBL_CATALOG` cod 1).
- Admin `userAdmin/12345678`; normales `userNormal/12345678` y `epn.fis/12345678`.
- Sólo el admin hace CRUD de usuarios, canciones, artistas y álbumes.
- El usuario normal gestiona **sus** playlists (crear/editar/borrar/añadir canciones).
- La propiedad de una playlist se comprueba **siempre contra la base de datos**.

## Despliegue
- Destino: `C:\inetpub\wwwroot\PoliMusic` → URL `http://localhost/PoliMusic/`.
- La app detecta sola su carpeta base (`EnsureAppRoot`), funciona en raíz o en subcarpeta.
- Lanzar `Deploy\desplegar_como_administrador.bat` (Ejecutar como administrador).
- `robocopy /MIR` excluye `uploads/` para no borrar los archivos subidos por los usuarios.
