<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="inc/boot.asp"-->
<%
' =============================================================================
'  PoliMusic - Perfil del usuario conectado
' =============================================================================
Dim pf_rs, pf_frm, pf_act, pf_err, pf_ok, pf_user, pf_mail, pf_birth, pf_photo
Dim pf_pass, pf_pass2, pf_up, pf_nPl, pf_nSongs, pf_saved

RequireLogin
InitPage "Mi perfil", "profile"

pf_err  = ""
pf_ok   = ""
pf_user = g_Username
pf_mail = g_UserEmail
pf_birth = ""
pf_photo = g_UserPhoto
pf_saved = False

' Datos actuales
Set pf_rs = DB_RS("SELECT USERNAME, EMAIL, BIRTHDAY, USER_PHOTO FROM TBL_USER WHERE ID_USER = @p0", Array(g_UserID))
If IsObject(pf_rs) Then
    If Not (pf_rs Is Nothing) Then
        If Not pf_rs.EOF Then
            pf_user  = N(pf_rs("USERNAME"))
            pf_mail  = N(pf_rs("EMAIL"))
            pf_birth = IsoDate(pf_rs("BIRTHDAY"))
            pf_photo = N(pf_rs("USER_PHOTO"))
        End If
        pf_rs.Close
    End If
End If
Set pf_rs = Nothing

If IsPost() Then
    Set pf_frm = New clsForm
    pf_frm.Load
    pf_act = LCase(pf_frm.Val("action"))

    If pf_act = "profile" Then
        pf_user  = Trim(pf_frm.Val("username"))
        pf_mail  = Trim(pf_frm.Val("email"))
        pf_birth = Trim(pf_frm.Val("birthday"))
        pf_photo = Trim(pf_frm.Val("user_photo"))
        pf_pass  = pf_frm.Val("password")
        pf_pass2 = pf_frm.Val("password2")

        If Len(pf_user) < 3 Then
            pf_err = "El usuario debe tener al menos 3 caracteres."
        ElseIf Len(pf_user) > 50 Then
            pf_err = "El usuario no puede superar los 50 caracteres."
        ElseIf Not IsEmail(pf_mail) Then
            pf_err = "Escribe un correo electrónico válido."
        ElseIf Len(pf_birth) > 0 And Not IsDate(pf_birth) Then
            pf_err = "La fecha de nacimiento no es válida."
        ElseIf DB_Count("SELECT COUNT(*) FROM TBL_USER WHERE USERNAME = @p0 AND ID_USER <> @p1", Array(pf_user, g_UserID)) > 0 Then
            pf_err = "Ese nombre de usuario ya está en uso."
        ElseIf DB_Count("SELECT COUNT(*) FROM TBL_USER WHERE EMAIL = @p0 AND ID_USER <> @p1", Array(pf_mail, g_UserID)) > 0 Then
            pf_err = "Ese correo ya está en uso."
        ElseIf Len(pf_pass) > 0 And Len(pf_pass) < 6 Then
            pf_err = "La nueva contraseña debe tener al menos 6 caracteres."
        ElseIf Len(pf_pass) > 0 And pf_pass <> pf_pass2 Then
            pf_err = "Las contraseñas no coinciden."
        Else
            If pf_frm.HasFile("photo_file") Then
                pf_up = pf_frm.SaveFile("photo_file", "uploads/userImages", "user", "|.jpg|.jpeg|.png|.gif|.webp|.svg|", 4 * 1024 * 1024)
                If Len(pf_up) > 0 Then
                    pf_photo = pf_up
                ElseIf Len(pf_frm.Errors) > 0 Then
                    pf_err = pf_frm.Errors
                End If
            End If

            If Len(pf_err) = 0 Then
                If Len(pf_photo) = 0 Then pf_photo = "~/userImages/default.jpg"
                If Len(pf_pass) > 0 Then
                    DB_Run "UPDATE TBL_USER SET USERNAME = @p0, EMAIL = @p1, BIRTHDAY = @p2, USER_PHOTO = @p3, " & _
                           "PASSWORD = CONVERT(varchar(40), HASHBYTES('SHA1', CAST(@p4 AS varchar(200))), 2) WHERE ID_USER = @p5", _
                           Array(pf_user, pf_mail, SqlDateOnly(pf_birth), pf_photo, pf_pass, g_UserID)
                Else
                    DB_Run "UPDATE TBL_USER SET USERNAME = @p0, EMAIL = @p1, BIRTHDAY = @p2, USER_PHOTO = @p3 WHERE ID_USER = @p4", _
                           Array(pf_user, pf_mail, SqlDateOnly(pf_birth), pf_photo, g_UserID)
                End If
                RefreshSessionUser
                Flash "Perfil actualizado correctamente.", "ok"
                Go AppUrl("profile.asp")
            End If
        End If
    End If
End If

pf_nPl    = DB_Count("SELECT COUNT(*) FROM TBL_PLAYLIST WHERE ID_USER = @p0", Array(g_UserID))
pf_nSongs = DB_Count("SELECT COUNT(DISTINCT ID_SONG) FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST IN (SELECT ID_PLAYLIST FROM TBL_PLAYLIST WHERE ID_USER = @p0)", Array(g_UserID))
%>
<!--#include file="inc/layout_head.asp"-->

<header class="detail hero--profile">
  <img class="hero__avatar" src="<%=H(AssetUrl(pf_photo, "user"))%>" alt="<%=H(pf_user)%>">
  <div class="detail__info">
    <div class="detail__kind">Perfil</div>
    <h1 class="detail__title"><%=H(g_Username)%></h1>
    <div class="detail__meta">
      <span class="badge <%=IIf(g_IsAdmin, "badge--admin", "badge--user")%>"><%=H(RoleName(g_UserType))%></span>
      <span class="dot"></span><span><%=H(g_UserEmail)%></span>
    </div>
  </div>
</header>

<div class="stats mb-24">
  <div class="stat"><div class="stat__num"><%=pf_nPl%></div><div class="stat__lbl">Playlists</div></div>
  <div class="stat"><div class="stat__num"><%=pf_nSongs%></div><div class="stat__lbl">Canciones guardadas</div></div>
  <div class="stat"><div class="stat__num"><%=IIf(g_IsAdmin, "ADMIN", "USER")%></div><div class="stat__lbl">Rol</div></div>
</div>

<% If Len(pf_err) > 0 Then %>
<div class="alert alert--error"><span class="alert__dot"></span><span><%=H(pf_err)%></span></div>
<% End If %>

<section class="panel panel--narrow">
  <div class="panel__head">
    <h3>Editar mis datos</h3>
    <p>Puedes cambiar tu nombre de usuario, correo, fecha de nacimiento y fotografía. La contraseña es opcional.</p>
  </div>
  <form method="post" action="<%=AppUrl("profile.asp")%>" enctype="multipart/form-data">
    <input type="hidden" name="action" value="profile">
    <div class="form-grid">
      <div class="field">
        <label for="username">Usuario *</label>
        <input class="input" type="text" id="username" name="username" value="<%=H(pf_user)%>" maxlength="50" required>
      </div>
      <div class="field">
        <label for="email">Correo electrónico *</label>
        <input class="input" type="email" id="email" name="email" value="<%=H(pf_mail)%>" maxlength="50" required>
      </div>
      <div class="field">
        <label for="birthday">Fecha de nacimiento</label>
        <input class="input" type="date" id="birthday" name="birthday" value="<%=H(pf_birth)%>">
      </div>
      <div class="field">
        <label for="user_photo">Ruta de la fotografía</label>
        <input class="input" type="text" id="user_photo" name="user_photo" value="<%=H(pf_photo)%>" maxlength="255">
      </div>
      <div class="field field--full">
        <label for="photo_file">Subir nueva fotografía <span class="hint">(opcional, max. 4 MB)</span></label>
        <input class="input" type="file" id="photo_file" name="photo_file" accept="image/*">
      </div>
      <div class="field">
        <label for="password">Nueva contraseña <span class="hint">(opcional)</span></label>
        <input class="input" type="password" id="password" name="password" maxlength="50" autocomplete="new-password">
      </div>
      <div class="field">
        <label for="password2">Repetir nueva contraseña</label>
        <input class="input" type="password" id="password2" name="password2" maxlength="50" autocomplete="new-password">
      </div>
    </div>
    <div class="form-actions">
      <button class="btn btn--primary" type="submit"><%=Icon("check")%>Guardar cambios</button>
      <a class="btn btn--ghost" href="<%=AppUrl("home.asp")%>">Cancelar</a>
      <a class="btn btn--danger" href="<%=AppUrl("logout.asp")%>"><%=Icon("logout")%>Cerrar sesión</a>
    </div>
  </form>
</section>

<!--#include file="inc/layout_foot.asp"-->
