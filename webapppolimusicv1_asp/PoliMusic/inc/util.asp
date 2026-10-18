<%
' =============================================================================
'  PoliMusic - Utilidades generales
' =============================================================================

' -----------------------------------------------------------------------------
'  Conversion y saneamiento
' -----------------------------------------------------------------------------
' Normaliza cualquier valor a cadena.
' OJO: en IIS, Request.Form(...), Request.QueryString(...) y
' Request.ServerVariables(...) devuelven objetos "IStringList" (IsObject = True)
' que se convierten a texto mediante su propiedad predeterminada. Por eso NO se
' puede descartar todo lo que sea objeto: hay que intentar la conversion.
Function N(v)
    Dim s
    If IsNull(v) Or IsEmpty(v) Then
        N = ""
        Exit Function
    End If
    If IsObject(v) Then
        s = ""
        On Error Resume Next
        s = "" & v
        If Err.Number <> 0 Then s = ""
        Err.Clear
        On Error GoTo 0
        N = s
    Else
        N = CStr(v)
    End If
End Function

Function H(s)
    H = Server.HTMLEncode(N(s))
End Function

Function ToIntDef(v, d)
    Dim s, r
    r = d
    s = Trim(N(v))
    If Len(s) > 0 Then
        If IsNumeric(s) Then
            On Error Resume Next
            r = CLng(CDbl(s))
            If Err.Number <> 0 Then
                Err.Clear
                r = d
            End If
            On Error GoTo 0
        End If
    End If
    ToIntDef = r
End Function

Function ToInt(v)
    ToInt = ToIntDef(v, 0)
End Function

' Devuelve Null si la cadena esta vacia (util para columnas numericas opcionales).
Function IIfNull(v)
    Dim s
    s = Trim(N(v))
    If Len(s) = 0 Then
        IIfNull = Null
    Else
        IIfNull = s
    End If
End Function

' Operador condicional simple.
Function IIf(cond, valorSi, valorNo)
    If cond Then
        IIf = valorSi
    Else
        IIf = valorNo
    End If
End Function

Function Trunc(s, maxLen)
    Dim t
    t = N(s)
    If Len(t) > maxLen Then t = Left(t, maxLen - 1) & Chr(8230)
    Trunc = t
End Function

' -----------------------------------------------------------------------------
'  Entrada de datos (Request)
' -----------------------------------------------------------------------------
Function IsPost()
    IsPost = (UCase(N(Request.ServerVariables("REQUEST_METHOD"))) = "POST")
End Function

Function PStr(name)
    PStr = Trim(N(Request.Form(name)))
End Function

Function PInt(name)
    PInt = ToIntDef(Request.Form(name), 0)
End Function

Function PNum(name)
    Dim v
    v = Trim(N(Request.Form(name)))
    If Len(v) = 0 Then
        PNum = Null
    ElseIf IsNumeric(v) Then
        PNum = CDbl(v)
    Else
        PNum = Null
    End If
End Function

Function PDate(name)
    Dim v
    v = Trim(N(Request.Form(name)))
    If Len(v) = 0 Then
        PDate = Null
    ElseIf IsDate(v) Then
        PDate = CDate(v)
    Else
        PDate = Null
    End If
End Function

Function PChk(name)
    Dim v
    v = LCase(Trim(N(Request.Form(name))))
    PChk = (v = "1" Or v = "on" Or v = "true" Or v = "yes" Or v = "si")
End Function

Function QStr(name)
    QStr = Trim(N(Request.QueryString(name)))
End Function

Function QInt(name)
    QInt = ToIntDef(Request.QueryString(name), 0)
End Function

' -----------------------------------------------------------------------------
'  Rutas de la aplicacion
' -----------------------------------------------------------------------------
' Deduce la URL base de la aplicacion (p.ej. "/PoliMusic/") para que todos los
' enlaces, hojas de estilo, imagenes y formularios funcionen tanto si el
' aplicativo esta publicado en la raiz del sitio como dentro de una carpeta
' (caso normal: C:\inetpub\wwwroot\PoliMusic -> http://servidor/PoliMusic/).
'
' OJO: APPL_MD_PATH solo contiene la ruta de la aplicacion cuando la carpeta es
' una aplicacion o un directorio virtual de IIS. Si es un simple subdirectorio
' de wwwroot, APPL_MD_PATH vale "/LM/W3SVC/1/ROOT" y no aporta nada, por lo que
' en ese caso hay que deducir la carpeta a partir de SCRIPT_NAME.
Sub EnsureAppRoot()
    Dim p, i, app, sn
    If Not IsEmpty(g_AppRoot) Then Exit Sub

    app = ""
    p = N(Request.ServerVariables("APPL_MD_PATH"))
    i = InStr(LCase(p), "/root")
    If i > 0 Then app = Mid(p, i + 5)

    ' SCRIPT_NAME es "/PoliMusic/login.asp" -> se quita el nombre del archivo
    sn = N(Request.ServerVariables("SCRIPT_NAME"))
    i = InStrRev(sn, "/")
    If i > 0 Then sn = Left(sn, i - 1) Else sn = ""

    If Len(app) = 0 Then app = sn

    p = app
    If Len(p) > 0 Then
        If Left(p, 1) <> "/" Then p = "/" & p
        If Right(p, 1) <> "/" Then p = p & "/"
    End If
    g_AppRoot = p
End Sub

' Convierte una ruta relativa al sitio en una URL absoluta de la aplicacion.
Function AppUrl(rel)
    Dim r
    EnsureAppRoot
    r = N(rel)
    Do While Len(r) > 0
        If Left(r, 1) = "/" Then r = Mid(r, 2) Else Exit Do
    Loop
    If Len(g_AppRoot) = 0 Then
        AppUrl = "/" & r
    Else
        AppUrl = g_AppRoot & r
    End If
End Function

' -----------------------------------------------------------------------------
'  Recursos (imagenes y audio)
' -----------------------------------------------------------------------------
Function FSObj()
    If Not IsObject(g_FSO) Then
        Set g_FSO = Server.CreateObject("Scripting.FileSystemObject")
    End If
    Set FSObj = g_FSO
End Function

Function DefaultAsset(kind)
    Select Case LCase(kind)
        Case "song"     : DefaultAsset = AppUrl("assets/songFiles/bensound-adventure.mp3")
        Case "user"     : DefaultAsset = AppUrl("assets/userImages/default.jpg")
        Case "playlist" : DefaultAsset = AppUrl("assets/playlistImages/default.jpg")
        Case "album"    : DefaultAsset = AppUrl("assets/img/default-album.svg")
        Case "artist"   : DefaultAsset = AppUrl("assets/img/default-artist.svg")
        Case Else       : DefaultAsset = AppUrl("assets/img/logo.svg")
    End Select
End Function

' Traduce la ruta guardada en la base de datos (p.ej. "~/songFiles/x.mp3")
' a una URL real, comprobando que el archivo exista en disco. Si no existe,
' devuelve una imagen/audio por defecto.
Function AssetUrl(stored, kind)
    Dim raw, seg, fname, folder, base, cands, i, root, fso, dotPos, prefix

    raw = Replace(N(stored), "\", "/")
    raw = Trim(raw)
    If Left(raw, 2) = "~/" Then raw = Mid(raw, 3)

    ' Los archivos subidos por el administrador se guardan bajo "uploads/"
    If LCase(Left(raw, 8)) = "uploads/" Then
        prefix = "uploads/"
        raw = Mid(raw, 9)
    Else
        prefix = "assets/"
        If LCase(Left(raw, 7)) = "assets/" Then raw = Mid(raw, 8)
    End If

    seg = ""
    fname = ""
    If InStr(raw, "/") > 0 Then
        seg = LCase(Left(raw, InStrRev(raw, "/") - 1))
        fname = Mid(raw, InStrRev(raw, "/") + 1)
    Else
        fname = raw
    End If

    Select Case seg
        Case "songfiles"                    : folder = "songFiles"
        Case "userimages"                   : folder = "userImages"
        Case "playlistimages"               : folder = "playlistImages"
        Case "albumimages"                  : folder = "albumImages"
        Case "artistimages"                 : folder = "artistImages"
        Case "img"                          : folder = "img"
        Case Else                           : folder = ""
    End Select

    If Len(folder) = 0 Or Len(fname) = 0 Then
        AssetUrl = DefaultAsset(kind)
        Exit Function
    End If

    dotPos = InStrRev(fname, ".")
    If dotPos > 0 Then base = Left(fname, dotPos - 1) Else base = fname

    cands = Array(fname, base & ".png", base & ".jpg", base & ".jpeg", _
                  base & ".svg", base & ".gif", base & ".webp", base & ".mp3")

    On Error Resume Next
    Set fso = FSObj()
    root = Server.MapPath(prefix & folder & "/")
    If Err.Number = 0 Then
        ' Server.MapPath NO garantiza la barra final: en IIS devuelve
        ' "...\assets\songFiles" sin "\" al final. Sin anadirla, la comprobacion
        ' FileExists daria siempre Falso y se usaria la imagen/audio por defecto.
        If Len(root) > 0 Then
            If Right(root, 1) <> "\" And Right(root, 1) <> "/" Then root = root & "\"
        End If
        For i = 0 To UBound(cands)
            If fso.FileExists(root & cands(i)) Then
                On Error GoTo 0
                AssetUrl = AppUrl(prefix & folder & "/" & cands(i))
                Exit Function
            End If
        Next
    End If
    Err.Clear
    On Error GoTo 0

    AssetUrl = DefaultAsset(kind)
End Function

' -----------------------------------------------------------------------------
'  Fechas
' -----------------------------------------------------------------------------
Function FmtDate(v)
    Dim d
    FmtDate = ""
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    If Not IsDate(v) Then Exit Function
    d = CDate(v)
    FmtDate = Right("0" & Day(d), 2) & "/" & Right("0" & Month(d), 2) & "/" & Year(d)
End Function

Function FmtDateTime(v)
    Dim d
    FmtDateTime = ""
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    If Not IsDate(v) Then Exit Function
    d = CDate(v)
    FmtDateTime = FmtDate(d) & " " & Right("0" & Hour(d), 2) & ":" & Right("0" & Minute(d), 2)
End Function

' Valor para un <input type="date"> a partir de un datetime de SQL Server.
Function IsoDate(v)
    Dim d
    IsoDate = ""
    If IsNull(v) Or IsEmpty(v) Then Exit Function
    If Not IsDate(v) Then Exit Function
    d = CDate(v)
    IsoDate = Year(d) & "-" & Right("0" & Month(d), 2) & "-" & Right("0" & Day(d), 2)
End Function

' Fecha en formato AAAAMMDD (siempre interpretable por SQL Server) o Null.
Function SqlDateOnly(v)
    Dim d
    If IsNull(v) Or IsEmpty(v) Then
        SqlDateOnly = Null
        Exit Function
    End If
    If Not IsDate(v) Then
        SqlDateOnly = Null
        Exit Function
    End If
    d = CDate(v)
    SqlDateOnly = Year(d) & Right("0" & Month(d), 2) & Right("0" & Day(d), 2)
End Function

' -----------------------------------------------------------------------------
'  Validaciones
' -----------------------------------------------------------------------------
Function IsEmail(s)
    Dim t, at, dot
    t = Trim(N(s))
    at = InStr(t, "@")
    dot = InStrRev(t, ".")
    IsEmail = (Len(t) >= 6 And at > 1 And dot > at + 1 And dot < Len(t) And InStr(t, " ") = 0)
End Function

Function SqlLike(s)
    Dim t
    t = Replace(N(s), "[", "[[]")
    t = Replace(t, "%", "[%]")
    t = Replace(t, "_", "[_]")
    SqlLike = t
End Function

' -----------------------------------------------------------------------------
'  Mensajes flash (se muestran una sola vez tras un redirect)
' -----------------------------------------------------------------------------
Sub Flash(msg, kind)
    Session("flash_msg") = N(msg)
    Session("flash_kind") = N(kind)
End Sub

Function RenderFlash()
    Dim m, k
    m = N(Session("flash_msg"))
    k = N(Session("flash_kind"))
    Session("flash_msg") = ""
    Session("flash_kind") = ""
    If Len(m) = 0 Then
        RenderFlash = ""
        Exit Function
    End If
    If Len(k) = 0 Then k = "ok"
    If k <> "ok" And k <> "error" And k <> "warn" Then k = "ok"
    RenderFlash = "<div class=""alert alert--" & k & """ role=""status"">" & _
                  "<span class=""alert__dot""></span><span>" & H(m) & "</span></div>"
End Function

' -----------------------------------------------------------------------------
'  Redireccion
' -----------------------------------------------------------------------------
Sub Go(url)
    Response.Redirect url
    Response.End
End Sub

' -----------------------------------------------------------------------------
'  Catalogos (TBL_CATALOG)
' -----------------------------------------------------------------------------
Function CatalogName(parentCode, code)
    CatalogName = N(DB_Scalar("SELECT VALUE FROM TBL_CATALOG WHERE COD_CATALOG_PARENT = @p0 AND COD_CATALOG = @p1", _
                              Array(parentCode, CStr(code)), ""))
End Function

Function RoleName(t)
    Dim v
    v = CatalogName("1", CStr(t))
    If Len(v) = 0 Then
        If ToInt(t) = ROLE_ADMIN Then v = "ADMIN" Else v = "USER"
    End If
    RoleName = v
End Function

Function PlaylistTypeName(t)
    Dim v
    v = CatalogName("2", CStr(t))
    If Len(v) = 0 Then
        If ToInt(t) = PL_PRIVATE Then v = "Private" Else v = "Public"
    End If
    PlaylistTypeName = v
End Function

' Devuelve el ID_USER propietario real de una playlist consultando la base de
' datos. Nunca debe tomarse el propietario de un campo del formulario, porque el
' cliente puede falsificarlo y editar/borrar playlists de otros usuarios.
Function PlaylistOwnerId(id)
    If ToInt(id) <= 0 Then
        PlaylistOwnerId = 0
    Else
        PlaylistOwnerId = ToIntDef(DB_Scalar("SELECT ID_USER FROM TBL_PLAYLIST WHERE ID_PLAYLIST = @p0", Array(ToInt(id)), 0), 0)
    End If
End Function

' -----------------------------------------------------------------------------
'  Texto de apoyo
' -----------------------------------------------------------------------------
Function NavClass(key)
    If LCase(N(g_ActiveNav)) = LCase(N(key)) Then
        NavClass = "is-active"
    Else
        NavClass = ""
    End If
End Function

Function PlaysLabel(n)
    Dim v
    v = ToInt(n)
    If v = 1 Then
        PlaysLabel = "1 reproducción"
    Else
        PlaysLabel = v & " reproducciones"
    End If
End Function

Function DurLabel(seconds)
    Dim s, m
    s = ToInt(seconds)
    If s < 0 Then s = 0
    m = s \ 60
    s = s Mod 60
    DurLabel = m & ":" & Right("0" & s, 2)
End Function
%>
