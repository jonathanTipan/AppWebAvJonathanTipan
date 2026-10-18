<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="inc/boot.asp"-->
<%
' =============================================================================
'  PoliMusic - Gestion de usuarios (exclusivo del administrador)
' =============================================================================
Dim us_rs, us_frm, us_act, us_id, us_err, us_row, us_new, us_acts
Dim us_user, us_mail, us_pass, us_birth, us_photo, us_type, us_up, us_q, us_where, us_prms, us_n

RequireAdmin
InitPage "Usuarios", "users"

us_act  = ""
us_err  = ""
us_id   = QInt("id")
us_q    = QStr("q")
us_new  = (LCase(QStr("action")) = "new")
us_photo = "~/userImages/default.jpg"
us_type = ROLE_USER
us_where = ""
us_prms = Array()

If IsPost() Then
    Set us_frm = New clsForm
    us_frm.Load
    us_act = LCase(us_frm.Val("action"))

    If us_act = "save" Then
        us_new   = True
        us_id    = ToIntDef(us_frm.Val("id"), 0)
        us_user  = Trim(us_frm.Val("username"))
        us_mail  = Trim(us_frm.Val("email"))
        us_pass  = us_frm.Val("password")
        us_birth = Trim(us_frm.Val("birthday"))
        us_photo = Trim(us_frm.Val("user_photo"))
        us_type  = ToIntDef(us_frm.Val("user_type"), ROLE_USER)

        If Len(us_user) < 3 Then
            us_err = "El usuario debe tener al menos 3 caracteres."
        ElseIf Len(us_user) > 50 Then
            us_err = "El usuario no puede superar los 50 caracteres."
        ElseIf Not IsEmail(us_mail) Then
            us_err = "Escribe un correo electrónico válido."
        ElseIf Len(us_birth) > 0 And Not IsDate(us_birth) Then
            us_err = "La fecha de nacimiento no es válida."
        ElseIf us_type <> ROLE_USER And us_type <> ROLE_ADMIN Then
            us_err = "Rol de usuario no válido."
        ElseIf us_id = 0 And Len(us_pass) < 6 Then
            us_err = "La contraseña debe tener al menos 6 caracteres."
        ElseIf Len(us_pass) > 0 And Len(us_pass) < 6 Then
            us_err = "La nueva contraseña debe tener al menos 6 caracteres."
        ElseIf DB_Count("SELECT COUNT(*) FROM TBL_USER WHERE USERNAME = @p0 AND ID_USER <> @p1", Array(us_user, us_id)) > 0 Then
            us_err = "Ya existe otro usuario con ese nombre."
        ElseIf DB_Count("SELECT COUNT(*) FROM TBL_USER WHERE EMAIL = @p0 AND ID_USER <> @p1", Array(us_mail, us_id)) > 0 Then
            us_err = "Ya existe otro usuario con ese correo."
        Else
            If us_frm.HasFile("photo_file") Then
                us_up = us_frm.SaveFile("photo_file", "uploads/userImages", "user", "|.jpg|.jpeg|.png|.gif|.webp|.svg|", 4 * 1024 * 1024)
                If Len(us_up) > 0 Then
                    us_photo = us_up
                ElseIf Len(us_frm.Errors) > 0 Then
                    us_err = us_frm.Errors
                End If
            End If

            If Len(us_err) = 0 Then
                If Len(us_photo) = 0 Then us_photo = "~/userImages/default.jpg"
                If us_id > 0 Then
                    If Len(us_pass) > 0 Then
                        DB_Run "UPDATE TBL_USER SET USERNAME = @p0, EMAIL = @p1, BIRTHDAY = @p2, USER_PHOTO = @p3, USER_TYPE = @p4, " & _
                               "PASSWORD = CONVERT(varchar(40), HASHBYTES('SHA1', CAST(@p5 AS varchar(200))), 2) WHERE ID_USER = @p6", _
                               Array(us_user, us_mail, SqlDateOnly(us_birth), us_photo, us_type, us_pass, us_id)
                    Else
                        DB_Run "UPDATE TBL_USER SET USERNAME = @p0, EMAIL = @p1, BIRTHDAY = @p2, USER_PHOTO = @p3, USER_TYPE = @p4 WHERE ID_USER = @p5", _
                               Array(us_user, us_mail, SqlDateOnly(us_birth), us_photo, us_type, us_id)
                    End If
                    Flash "Usuario actualizado correctamente.", "ok"
                    If us_id = g_UserID Then RefreshSessionUser
                Else
                    DB_Run "INSERT INTO TBL_USER (USERNAME, PASSWORD, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE) " & _
                           "VALUES (@p0, CONVERT(varchar(40), HASHBYTES('SHA1', CAST(@p1 AS varchar(200))), 2), @p2, @p3, @p4, @p5)", _
                           Array(us_user, us_pass, us_mail, SqlDateOnly(us_birth), us_photo, us_type)
                    Flash "Usuario creado correctamente.", "ok"
                End If
                Go AppUrl("users.asp")
            End If
        End If

    ElseIf us_act = "delete" Then
        us_id = ToIntDef(us_frm.Val("id"), 0)
        If us_id = g_UserID Then
            Flash "No puedes eliminar tu propia cuenta mientras la estás usando.", "error"
        ElseIf us_id > 0 Then
            DB_Run "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST IN (SELECT ID_PLAYLIST FROM TBL_PLAYLIST WHERE ID_USER = @p0)", Array(us_id)
            DB_Run "DELETE FROM TBL_PLAYLIST WHERE ID_USER = @p0", Array(us_id)
            DB_Run "DELETE FROM TBL_USER WHERE ID_USER = @p0", Array(us_id)
            Flash "Usuario eliminado.", "ok"
        End If
        Go AppUrl("users.asp")
    End If
End If

If us_id > 0 And Len(us_err) = 0 Then
    us_new = True
    Set us_row = DB_RS("SELECT USERNAME, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE FROM TBL_USER WHERE ID_USER = @p0", Array(us_id))
    If IsObject(us_row) Then
        If Not (us_row Is Nothing) Then
            If us_row.EOF Then
                us_row.Close
                Set us_row = Nothing
                Flash "El usuario solicitado no existe.", "error"
                Go AppUrl("users.asp")
            Else
                us_user  = N(us_row("USERNAME"))
                us_mail  = N(us_row("EMAIL"))
                us_birth = IsoDate(us_row("BIRTHDAY"))
                us_photo = N(us_row("USER_PHOTO"))
                us_type  = ToInt(us_row("USER_TYPE"))
            End If
            us_row.Close
        End If
    End If
    Set us_row = Nothing
End If

If IsEmpty(us_user) Then us_user = ""
If IsEmpty(us_mail) Then us_mail = ""
If IsEmpty(us_birth) Then us_birth = ""
%>
<!--#include file="inc/layout_head.asp"-->

<div class="page-head">
  <div class="page-head__txt">
    <h2>Usuarios</h2>
    <p>Gestión de cuentas y roles. El rol <b>ADMIN</b> tiene acceso a todas las operaciones del sistema.</p>
  </div>
  <div class="page-head__actions">
    <form class="searchbox" action="<%=AppUrl("users.asp")%>" method="get" role="search">
      <span class="searchbox__ic"><%=Icon("search")%></span>
      <input type="search" name="q" value="<%=H(us_q)%>" placeholder="Buscar usuario o correo">
    </form>
    <a class="btn btn--primary" href="<%=AppUrl("users.asp?action=new")%>"><%=Icon("plus")%>Nuevo usuario</a>
  </div>
</div>

<% If Len(us_err) > 0 Then %>
<div class="alert alert--error"><span class="alert__dot"></span><span><%=H(us_err)%></span></div>
<% End If %>

<% If us_new Then %>
<section class="panel mb-24">
  <div class="panel__head">
    <h3><% If us_id > 0 Then %>Editar usuario<% Else %>Nuevo usuario<% End If %></h3>
    <p><% If us_id > 0 Then %>Deja la contraseña en blanco para conservar la actual.<% Else %>La contraseña se guarda cifrada con SHA1.<% End If %></p>
  </div>
  <form method="post" action="<%=AppUrl("users.asp")%>" enctype="multipart/form-data">
    <input type="hidden" name="action" value="save">
    <input type="hidden" name="id" value="<%=us_id%>">
    <div class="form-grid">
      <div class="field">
        <label for="username">Usuario *</label>
        <input class="input" type="text" id="username" name="username" value="<%=H(us_user)%>" maxlength="50" required>
      </div>
      <div class="field">
        <label for="email">Correo electrónico *</label>
        <input class="input" type="email" id="email" name="email" value="<%=H(us_mail)%>" maxlength="50" required>
      </div>
      <div class="field">
        <label for="password">Contraseña <% If us_id > 0 Then %><span class="hint">(en blanco = sin cambios)</span><% Else %>*<% End If %></label>
        <input class="input" type="password" id="password" name="password" maxlength="50" <% If us_id = 0 Then %>required<% End If %> autocomplete="new-password">
      </div>
      <div class="field">
        <label for="birthday">Fecha de nacimiento</label>
        <input class="input" type="date" id="birthday" name="birthday" value="<%=H(us_birth)%>">
      </div>
      <div class="field">
        <label for="user_type">Rol</label>
        <select class="select" id="user_type" name="user_type">
          <option value="0" <% If ToInt(us_type) = ROLE_USER Then %>selected<% End If %>>USER - usuario normal</option>
          <option value="1" <% If ToInt(us_type) = ROLE_ADMIN Then %>selected<% End If %>>ADMIN - administrador</option>
        </select>
      </div>
      <div class="field">
        <label for="user_photo">Ruta de la fotografía</label>
        <input class="input" type="text" id="user_photo" name="user_photo" value="<%=H(us_photo)%>" maxlength="255" placeholder="~/userImages/default.jpg">
      </div>
      <div class="field field--full">
        <label for="photo_file">Subir fotografía <span class="hint">(opcional, max. 4 MB)</span></label>
        <input class="input" type="file" id="photo_file" name="photo_file" accept="image/*">
      </div>
    </div>
    <div class="form-actions">
      <button class="btn btn--primary" type="submit"><%=Icon("check")%>Guardar</button>
      <a class="btn btn--ghost" href="<%=AppUrl("users.asp")%>">Cancelar</a>
    </div>
  </form>
</section>
<% End If %>

<%
us_q = Trim(us_q)
If Len(us_q) > 0 Then
    us_where = "WHERE (u.USERNAME LIKE @p0 OR u.EMAIL LIKE @p0)"
    us_prms = Array("%" & SqlLike(us_q) & "%")
End If

Set us_rs = DB_RS("SELECT u.ID_USER, u.USERNAME, u.EMAIL, u.BIRTHDAY, u.USER_PHOTO, u.USER_TYPE, " & _
                  "(SELECT COUNT(*) FROM TBL_PLAYLIST p WHERE p.ID_USER = u.ID_USER) AS N_PL " & _
                  "FROM TBL_USER u " & us_where & " ORDER BY u.ID_USER", us_prms)
us_n = 0
If IsObject(us_rs) Then
    If Not (us_rs Is Nothing) Then
        If us_rs.EOF Then
            Response.Write EmptyState("users", "Sin usuarios", "No hay usuarios que coincidan con la búsqueda.", "")
        Else
            Response.Write "<div class=""tablewrap""><table class=""table""><thead><tr>" & _
                "<th>Usuario</th><th>Correo</th><th>Nacimiento</th><th>Rol</th><th>Playlists</th><th style=""text-align:right"">Acciones</th>" & _
                "</tr></thead><tbody>"
            Do While Not us_rs.EOF
                us_n = us_n + 1
                Response.Write "<tr>" & _
                    "<td><div class=""cell-main""><img class=""thumb thumb--round"" src=""" & H(AssetUrl(us_rs("USER_PHOTO"), "user")) & """ alt="""">" & _
                    "<span class=""cell-main__txt""><b>" & H(us_rs("USERNAME")) & "</b><small>#" & ToInt(us_rs("ID_USER")) & "</small></span></div></td>" & _
                    "<td>" & H(us_rs("EMAIL")) & "</td>" & _
                    "<td class=""nowrap"">" & IIf(Len(FmtDate(us_rs("BIRTHDAY"))) = 0, "-", FmtDate(us_rs("BIRTHDAY"))) & "</td>" & _
                    "<td><span class=""badge " & IIf(ToInt(us_rs("USER_TYPE")) = ROLE_ADMIN, "badge--admin", "badge--user") & """>" & H(RoleName(us_rs("USER_TYPE"))) & "</span></td>" & _
                    "<td>" & ToInt(us_rs("N_PL")) & "</td>" & _
                    "<td><div class=""table__actions"">" & _
                        EditLink(AppUrl("users.asp?action=new&id=" & ToInt(us_rs("ID_USER"))), "Editar usuario") & _
                        DeleteForm(AppUrl("users.asp"), "delete", "id", ToInt(us_rs("ID_USER")), "¿Eliminar este usuario y sus playlists?") & _
                    "</div></td></tr>"
                us_rs.MoveNext
            Loop
            Response.Write "</tbody></table></div>"
            Response.Write "<p class=""muted small mt-16"">" & us_n & " usuarios registrados.</p>"
        End If
        us_rs.Close
    End If
End If
Set us_rs = Nothing
%>

<!--#include file="inc/layout_foot.asp"-->
