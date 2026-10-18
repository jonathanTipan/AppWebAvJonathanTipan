<%
' =============================================================================
'  PoliMusic - Autenticacion y control de acceso por roles
' =============================================================================

' Carga el usuario autenticado desde la sesion.
Sub Auth_Load()
    g_UserID    = ToIntDef(Session("uid"), 0)
    g_Username  = N(Session("uname"))
    g_UserType  = ToIntDef(Session("utype"), ROLE_USER)
    g_UserPhoto = N(Session("uphoto"))
    g_UserEmail = N(Session("uemail"))
    g_IsLogged  = (g_UserID > 0)
    g_IsAdmin   = (g_IsLogged And g_UserType = ROLE_ADMIN)
End Sub

Sub InitPage(pageTitle, activeNav)
    g_PageTitle = pageTitle
    g_ActiveNav = activeNav
End Sub

' -----------------------------------------------------------------------------
'  Hash de contrasena (SHA1) calculado por SQL Server.
'  Se usa CAST a varchar para que el hash corresponda a los bytes ASCII, que es
'  el formato que ya existe en la tabla TBL_USER.
' -----------------------------------------------------------------------------
Function Sha1Sql(paramName)
    Sha1Sql = "CONVERT(varchar(40), HASHBYTES('SHA1', CAST(" & paramName & " AS varchar(200))), 2)"
End Function

Function HashPassword(pwd)
    HashPassword = N(DB_Scalar("SELECT " & Sha1Sql("@p0"), Array(N(pwd)), ""))
End Function

' -----------------------------------------------------------------------------
'  Intenta autenticar. Acepta contrasenas guardadas en texto plano (datos
'  heredados) o como hash SHA1 en hexadecimal.
'  Devuelve True y rellena la sesion si las credenciales son correctas.
' -----------------------------------------------------------------------------
Function TryLogin(username, password, ByRef errMsg)
    Dim sql, rs, ok, uid

    ok = False
    errMsg = ""

    sql = "SELECT TOP 1 ID_USER, USERNAME, USER_TYPE, USER_PHOTO, EMAIL, PASSWORD " & _
          "FROM TBL_USER WHERE USERNAME = @p0 AND (" & _
          "  PASSWORD = @p1 OR " & _
          "  UPPER(PASSWORD) = UPPER(" & Sha1Sql("@p1") & "))"

    Set rs = DB_RS(sql, Array(N(username), N(password)))

    If IsObject(rs) Then
        If Not (rs Is Nothing) Then
            If Not rs.EOF Then
                uid = ToInt(rs("ID_USER"))
                Session("uid")    = uid
                Session("uname")  = N(rs("USERNAME"))
                Session("utype")  = ToInt(rs("USER_TYPE"))
                Session("uphoto") = N(rs("USER_PHOTO"))
                Session("uemail") = N(rs("EMAIL"))
                Session("login_time") = Now()
                ok = True
            End If
            rs.Close
        End If
    End If
    Set rs = Nothing

    If ok Then
        Auth_Load
    Else
        errMsg = "Usuario o contraseña incorrectos."
    End If

    TryLogin = ok
End Function

Sub LogoutUser()
    Session.Contents.RemoveAll
    Session.Abandon
End Sub

' -----------------------------------------------------------------------------
'  Guardas de acceso
' -----------------------------------------------------------------------------
Sub RequireLogin()
    Auth_Load
    If Not g_IsLogged Then
        Go AppUrl("login.asp")
    End If
End Sub

Sub RequireAdmin()
    RequireLogin
    If Not g_IsAdmin Then
        Flash "Esa sección es exclusiva del usuario administrador.", "error"
        Go AppUrl("home.asp")
    End If
End Sub

' Recarga los datos del usuario en la sesion (tras editar el perfil).
Sub RefreshSessionUser()
    Dim rs
    Set rs = DB_RS("SELECT USERNAME, USER_TYPE, USER_PHOTO, EMAIL FROM TBL_USER WHERE ID_USER = @p0", Array(g_UserID))
    If IsObject(rs) Then
        If Not (rs Is Nothing) Then
            If Not rs.EOF Then
                Session("uname")  = N(rs("USERNAME"))
                Session("utype")  = ToInt(rs("USER_TYPE"))
                Session("uphoto") = N(rs("USER_PHOTO"))
                Session("uemail") = N(rs("EMAIL"))
            End If
            rs.Close
        End If
    End If
    Set rs = Nothing
    Auth_Load
End Sub

' -----------------------------------------------------------------------------
'  Ayudas de sesion
' -----------------------------------------------------------------------------
Function CurrentPhotoUrl()
    CurrentPhotoUrl = AssetUrl(g_UserPhoto, "user")
End Function

Function UserPhotoUrl(photoPath)
    UserPhotoUrl = AssetUrl(photoPath, "user")
End Function
%>
