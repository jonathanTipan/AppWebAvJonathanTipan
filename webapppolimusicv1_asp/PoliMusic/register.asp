<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="inc/boot.asp"-->
<%
' =============================================================================
'  PoliMusic - Registro (enrolamiento) de usuarios
'  Crea siempre usuarios con rol USER (USER_TYPE = 0).
' =============================================================================
Dim rg_frm, rg_user, rg_mail, rg_p1, rg_p2, rg_birth, rg_err, rg_photo, rg_newId, rg_dummy

Auth_Load
If g_IsLogged Then Go AppUrl("home.asp")

rg_err   = ""
rg_user  = ""
rg_mail  = ""
rg_birth = ""
rg_photo = "~/userImages/default.jpg"

If IsPost() Then
    Set rg_frm = New clsForm
    rg_frm.Load

    rg_user  = Trim(rg_frm.Val("username"))
    rg_mail  = Trim(rg_frm.Val("email"))
    rg_p1    = rg_frm.Val("password")
    rg_p2    = rg_frm.Val("password2")
    rg_birth = Trim(rg_frm.Val("birthday"))

    If Len(rg_user) < 3 Then
        rg_err = "El usuario debe tener al menos 3 caracteres."
    ElseIf Len(rg_user) > 50 Then
        rg_err = "El usuario no puede superar los 50 caracteres."
    ElseIf InStr(rg_user, " ") > 0 Then
        rg_err = "El usuario no puede contener espacios."
    ElseIf Not IsEmail(rg_mail) Then
        rg_err = "Escribe un correo electrónico válido."
    ElseIf Len(rg_p1) < 6 Then
        rg_err = "La contraseña debe tener al menos 6 caracteres."
    ElseIf rg_p1 <> rg_p2 Then
        rg_err = "Las contraseñas no coinciden."
    ElseIf Len(rg_birth) > 0 And Not IsDate(rg_birth) Then
        rg_err = "La fecha de nacimiento no es válida."
    ElseIf DB_Count("SELECT COUNT(*) FROM TBL_USER WHERE USERNAME = @p0", Array(rg_user)) > 0 Then
        rg_err = "Ese nombre de usuario ya está registrado."
    ElseIf DB_Count("SELECT COUNT(*) FROM TBL_USER WHERE EMAIL = @p0", Array(rg_mail)) > 0 Then
        rg_err = "Ese correo electrónico ya está registrado."
    Else
        ' Fotografia opcional
        If rg_frm.HasFile("photo") Then
            rg_photo = rg_frm.SaveFile("photo", DIR_USERS, "user", "|.jpg|.jpeg|.png|.gif|.webp|", 4 * 1024 * 1024)
            If Len(rg_photo) = 0 Then
                rg_photo = "~/userImages/default.jpg"
                If Len(rg_frm.Errors) > 0 Then rg_err = rg_frm.Errors
            End If
        End If

        If Len(rg_err) = 0 Then
            rg_newId = DB_Insert( _
                "INSERT INTO TBL_USER (USERNAME, PASSWORD, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE) " & _
                "VALUES (@p0, CONVERT(varchar(40), HASHBYTES('SHA1', CAST(@p1 AS varchar(200))), 2), @p2, @p3, @p4, @p5)", _
                Array(rg_user, rg_p1, rg_mail, SqlDateOnly(rg_birth), rg_photo, CStr(ROLE_USER)))

            If rg_newId > 0 Then
                rg_dummy = ""
                If TryLogin(rg_user, rg_p1, rg_dummy) Then
                    Go AppUrl("home.asp")
                Else
                    Go AppUrl("login.asp")
                End If
            Else
                rg_err = "No se pudo crear la cuenta. Intentalo de nuevo."
            End If
        End If
    End If
End If
%><!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="theme-color" content="#0a0f1c">
<title>Crear cuenta &middot; <%=APP_NAME%></title>
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
    <h1>Crea tu cuenta<br>y empieza a sonar.</h1>
    <p>El registro es inmediato. Podras escuchar el catálogo completo, consultar álbumes y artistas y armar tus propias playlists.</p>
    <div class="auth__features">
      <div class="auth__feature"><%=Icon("check")%><span>Acceso al catálogo completo de canciones</span></div>
      <div class="auth__feature"><%=Icon("check")%><span>Playlists públicas o privadas</span></div>
      <div class="auth__feature"><%=Icon("check")%><span>Tu foto y datos de perfil</span></div>
    </div>
  </section>

  <section class="auth__panel">
    <div class="auth__card">
      <h2>Crear cuenta</h2>
      <p class="sub">Completa tus datos para registrarte en <%=APP_NAME%>.</p>

      <% If Len(rg_err) > 0 Then %>
      <div class="alert alert--error"><span class="alert__dot"></span><span><%=H(rg_err)%></span></div>
      <% End If %>

      <form class="auth__form" method="post" action="<%=AppUrl("register.asp")%>" enctype="multipart/form-data">
        <div class="field">
          <label for="username">Usuario <span class="hint">(sin espacios)</span></label>
          <input class="input" type="text" id="username" name="username" value="<%=H(rg_user)%>" maxlength="50" required autofocus placeholder="p. ej. epn.fis">
        </div>
        <div class="field">
          <label for="email">Correo electrónico</label>
          <input class="input" type="email" id="email" name="email" value="<%=H(rg_mail)%>" maxlength="50" required placeholder="usuario@epn.edu.ec">
        </div>
        <div class="field">
          <label for="password">Contraseña <span class="hint">(mínimo 6 caracteres)</span></label>
          <input class="input" type="password" id="password" name="password" maxlength="50" required autocomplete="new-password">
        </div>
        <div class="field">
          <label for="password2">Repetir contraseña</label>
          <input class="input" type="password" id="password2" name="password2" maxlength="50" required autocomplete="new-password">
        </div>
        <div class="field">
          <label for="birthday">Fecha de nacimiento <span class="hint">(opcional)</span></label>
          <input class="input" type="date" id="birthday" name="birthday" value="<%=H(rg_birth)%>">
        </div>
        <div class="field">
          <label for="photo">Fotografía <span class="hint">(opcional, max. 4 MB)</span></label>
          <input class="input" type="file" id="photo" name="photo" accept="image/*">
        </div>
        <button class="btn btn--primary btn--block" type="submit"><%=Icon("user")%>Crear mi cuenta</button>
      </form>

      <div class="auth__foot">
        Ya tienes una cuenta? <a href="<%=AppUrl("login.asp")%>">Inicia sesión</a>
      </div>
    </div>
  </section>

</div>
</body>
</html>
