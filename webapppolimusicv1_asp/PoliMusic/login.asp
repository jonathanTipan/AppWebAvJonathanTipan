<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="inc/boot.asp"-->
<%
' =============================================================================
'  PoliMusic - Inicio de sesion
' =============================================================================
Dim lg_user, lg_pass, lg_err, lg_frm, lg_bye

Auth_Load
If g_IsLogged Then Go AppUrl("home.asp")

lg_err = ""
lg_user = ""
lg_bye = (QStr("bye") = "1")

If IsPost() Then
    Set lg_frm = New clsForm
    lg_frm.Load

    lg_user = Trim(lg_frm.Val("username"))
    lg_pass = lg_frm.Val("password")

    If Len(lg_user) = 0 Or Len(lg_pass) = 0 Then
        lg_err = "Escribe tu usuario y tu contraseña."
    ElseIf Len(lg_user) > 50 Then
        lg_err = "El nombre de usuario es demasiado largo."
    ElseIf TryLogin(lg_user, lg_pass, lg_err) Then
        Go AppUrl("home.asp")
    End If
End If
%><!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="theme-color" content="#0a0f1c">
<title>Iniciar sesión &middot; <%=APP_NAME%></title>
<link rel="icon" type="image/svg+xml" href="<%=AppUrl("assets/img/favicon.svg")%>">
<link rel="stylesheet" href="<%=AppUrl("assets/css/polimusic.css")%>">
</head>
<body>
<div class="auth">

  <section class="auth__side">
    <div class="auth__brand">
      <img src="<%=AppUrl("assets/img/logo.svg")%>" alt="">
      <b><%=APP_NAME%></b>
    </div>
    <h1>Toda tu música,<br>en un solo lugar.</h1>
    <p>Explora canciones, álbumes y artistas. Crea tus propias listas de reproducción y escuchalas donde quieras.</p>
    <div class="auth__features">
      <div class="auth__feature"><%=Icon("music")%><span>Reproductor continuo mientras navegas</span></div>
      <div class="auth__feature"><%=Icon("playlist")%><span>Playlists públicas y privadas</span></div>
      <div class="auth__feature"><%=Icon("disc")%><span>Catálogo de álbumes y artistas</span></div>
      <div class="auth__feature"><%=Icon("users")%><span>Administración completa para el rol admin</span></div>
    </div>
  </section>

  <section class="auth__panel">
    <div class="auth__card">
      <h2>Iniciar sesión</h2>
      <p class="sub">Accede con tu cuenta de <%=APP_NAME%> para continuar.</p>

      <% If lg_bye Then %>
      <div class="alert alert--ok"><span class="alert__dot"></span><span>Tu sesión se cerró correctamente. Hasta pronto.</span></div>
      <% End If %>

      <% If Len(lg_err) > 0 Then %>
      <div class="alert alert--error"><span class="alert__dot"></span><span><%=H(lg_err)%></span></div>
      <% End If %>

      <form class="auth__form" method="post" action="<%=AppUrl("login.asp")%>" autocomplete="on">
        <div class="field">
          <label for="username">Usuario</label>
          <input class="input" type="text" id="username" name="username" value="<%=H(lg_user)%>" maxlength="50" required autofocus autocomplete="username" placeholder="p. ej. userAdmin">
        </div>
        <div class="field">
          <label for="password">Contraseña</label>
          <input class="input" type="password" id="password" name="password" maxlength="50" required autocomplete="current-password" placeholder="Tu contraseña">
        </div>
        <button class="btn btn--primary btn--block" type="submit"><%=Icon("login")%>Entrar</button>
      </form>

      <div class="auth__demo">
        <b>Cuentas de prueba</b><br>
        Administrador: <b>userAdmin</b> / 12345678<br>
        Usuario normal: <b>userNormal</b> / 12345678<br>
        Usuario normal: <b>epn.fis</b> / 12345678
      </div>

      <div class="auth__foot">
        ¿No tienes cuenta? <a href="<%=AppUrl("register.asp")%>">Regístrate aquí</a>
      </div>
    </div>
  </section>

</div>
</body>
</html>
