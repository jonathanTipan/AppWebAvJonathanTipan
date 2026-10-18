<%
' =============================================================================
'  PoliMusic - Configuracion global
'  Proyecto: Clon de Spotify (tema azul) en ASP clasico + SQL Server
' =============================================================================
Option Explicit

' Todas las paginas se sirven como UTF-8
Response.CodePage = 65001
Response.CharSet = "utf-8"
Response.Buffer = True

' -----------------------------------------------------------------------------
'  Identidad de la aplicacion
' -----------------------------------------------------------------------------
Const APP_NAME        = "PoliMusic"
Const APP_TAGLINE     = "Música sin límites"
Const APP_VERSION     = "1.0.0"

' -----------------------------------------------------------------------------
'  Roles de usuario (TBL_CATALOG / 'USER TYPE')
' -----------------------------------------------------------------------------
Const ROLE_USER       = 0      ' Usuario normal
Const ROLE_ADMIN      = 1      ' Administrador

' -----------------------------------------------------------------------------
'  Tipos de playlist (TBL_CATALOG / 'PLAYLIST TYPE')
' -----------------------------------------------------------------------------
Const PL_PUBLIC       = 0
Const PL_PRIVATE      = 1

' -----------------------------------------------------------------------------
'  Carpetas de recursos (relativas a la raiz del sitio)
' -----------------------------------------------------------------------------
Const DIR_SONGS       = "assets/songFiles/"
Const DIR_USERS       = "assets/userImages/"
Const DIR_PLAYLISTS   = "assets/playlistImages/"
Const DIR_ALBUMS      = "assets/albumImages/"
Const DIR_ARTISTS     = "assets/artistImages/"
Const DIR_IMG         = "assets/img/"
Const DIR_UPLOADS     = "uploads/"

' -----------------------------------------------------------------------------
'  Cadena de conexion a SQL Server.
'  Se intentan en orden y se usa la primera que funcione, de modo que la
'  aplicacion arranca tanto con MSOLEDBSQL (recomendado) como con SQLOLEDB
'  (incluido de serie en Windows).
' -----------------------------------------------------------------------------
Const DB_SERVER       = "localhost"
Const DB_NAME         = "BDD_PoliMusic"
Const DB_USER         = "usr_polimusic"
Const DB_PASS         = "Politecnica1"

Dim g_ConnStrings
g_ConnStrings = Array( _
    "Provider=MSOLEDBSQL;Server=" & DB_SERVER & ";Database=" & DB_NAME & ";UID=" & DB_USER & ";PWD=" & DB_PASS & ";Encrypt=Optional;TrustServerCertificate=Yes;", _
    "Provider=MSOLEDBSQL;Server=" & DB_SERVER & ";Database=" & DB_NAME & ";UID=" & DB_USER & ";PWD=" & DB_PASS & ";", _
    "Provider=SQLNCLI11;Server=" & DB_SERVER & ";Database=" & DB_NAME & ";UID=" & DB_USER & ";PWD=" & DB_PASS & ";", _
    "Provider=SQLOLEDB;Server=" & DB_SERVER & ";Database=" & DB_NAME & ";UID=" & DB_USER & ";PWD=" & DB_PASS & ";" _
)

' -----------------------------------------------------------------------------
'  Estado de la peticion (se rellena en inc/auth.asp)
' -----------------------------------------------------------------------------
Dim g_Conn          ' ADODB.Connection activa
Dim g_FSO           ' Scripting.FileSystemObject reutilizable
Dim g_UserID        ' Id del usuario autenticado (0 = anonimo)
Dim g_Username
Dim g_UserType
Dim g_UserPhoto
Dim g_UserEmail
Dim g_IsLogged
Dim g_IsAdmin
Dim g_PageTitle
Dim g_ActiveNav
Dim g_AppRoot       ' URL base de la aplicacion, p.ej. "/PoliMusic/"
%>
