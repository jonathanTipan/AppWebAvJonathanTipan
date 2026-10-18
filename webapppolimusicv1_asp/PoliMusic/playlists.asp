<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="inc/boot.asp"-->
<%
' =============================================================================
'  PoliMusic - Playlists
'  El usuario normal gestiona SUS playlists (crear, consultar, actualizar,
'  eliminar). El administrador puede gestionar las de todos los usuarios.
' =============================================================================
Dim pl_rs, pl_frm, pl_act, pl_id, pl_err, pl_row, pl_new, pl_acts
Dim pl_title, pl_type, pl_photo, pl_up, pl_owner, pl_count, pl_can

RequireLogin
InitPage "Mis playlists", "playlists"

pl_act  = ""
pl_err  = ""
pl_id   = QInt("id")
pl_new  = (LCase(QStr("action")) = "new")
pl_type = PL_PUBLIC
pl_photo = "~/playlistImages/default.jpg"

If IsPost() Then
    Set pl_frm = New clsForm
    pl_frm.Load
    pl_act = LCase(pl_frm.Val("action"))

    If pl_act = "save" Then
        pl_new   = True
        pl_id    = ToIntDef(pl_frm.Val("id"), 0)
        pl_title = Trim(pl_frm.Val("title"))
        pl_type  = ToIntDef(pl_frm.Val("pl_type"), PL_PUBLIC)
        pl_photo = Trim(pl_frm.Val("photo_path"))

        ' El propietario se obtiene SIEMPRE de la base de datos, nunca del
        ' formulario: asi un usuario no puede editar playlists ajenas.
        pl_owner = PlaylistOwnerId(pl_id)

        ' Permisos: el dueno real o el administrador
        pl_can = (pl_id = 0) Or (pl_owner = g_UserID) Or g_IsAdmin

        If Not pl_can Then
            pl_err = "No tienes permisos sobre esa playlist."
        ElseIf Len(pl_title) = 0 Then
            pl_err = "El título de la playlist es obligatorio."
        ElseIf Len(pl_title) > 50 Then
            pl_err = "El título no puede superar los 50 caracteres."
        ElseIf pl_type <> PL_PUBLIC And pl_type <> PL_PRIVATE Then
            pl_err = "Tipo de playlist no válido."
        Else
            If pl_frm.HasFile("photo_file") Then
                pl_up = pl_frm.SaveFile("photo_file", "uploads/playlistImages", "playlist", "|.jpg|.jpeg|.png|.gif|.webp|.svg|", 6 * 1024 * 1024)
                If Len(pl_up) > 0 Then
                    pl_photo = pl_up
                ElseIf Len(pl_frm.Errors) > 0 Then
                    pl_err = pl_frm.Errors
                End If
            End If

            If Len(pl_err) = 0 Then
                If Len(pl_photo) = 0 Then pl_photo = "~/playlistImages/default.jpg"
                If pl_id > 0 Then
                    DB_Run "UPDATE TBL_PLAYLIST SET TITLE = @p0, TYPE = @p1, PHOTO = @p2 WHERE ID_PLAYLIST = @p3", _
                           Array(pl_title, pl_type, pl_photo, pl_id)
                    Flash "Playlist actualizada correctamente.", "ok"
                Else
                    DB_Run "INSERT INTO TBL_PLAYLIST (ID_USER, TITLE, CREATION_DATE, TYPE, PHOTO) " & _
                           "VALUES (@p0, @p1, GETDATE(), @p2, @p3)", _
                           Array(g_UserID, pl_title, pl_type, pl_photo)
                    Flash "Playlist creada correctamente.", "ok"
                End If
                Go AppUrl("playlists.asp")
            End If
        End If

    ElseIf pl_act = "delete" Then
        pl_id    = ToIntDef(pl_frm.Val("id"), 0)
        pl_owner = PlaylistOwnerId(pl_id)
        If pl_id > 0 Then
            If pl_owner = g_UserID Or g_IsAdmin Then
                DB_Run "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = @p0", Array(pl_id)
                DB_Run "DELETE FROM TBL_PLAYLIST WHERE ID_PLAYLIST = @p0", Array(pl_id)
                Flash "Playlist eliminada.", "ok"
            Else
                Flash "No tienes permisos para eliminar esa playlist.", "error"
            End If
        End If
        Go AppUrl("playlists.asp")
    End If
End If

' ---------------------------------------------------------------------------
'  Carga para edicion (solo el dueno o el administrador)
' ---------------------------------------------------------------------------
If pl_id > 0 And Len(pl_err) = 0 Then
    Set pl_row = DB_RS("SELECT ID_USER, TITLE, TYPE, PHOTO FROM TBL_PLAYLIST WHERE ID_PLAYLIST = @p0", Array(pl_id))
    If IsObject(pl_row) Then
        If Not (pl_row Is Nothing) Then
            If pl_row.EOF Then
                pl_row.Close
                Set pl_row = Nothing
                Flash "La playlist solicitada no existe.", "error"
                Go AppUrl("playlists.asp")
            Else
                pl_owner = ToInt(pl_row("ID_USER"))
                If pl_owner = g_UserID Or g_IsAdmin Then
                    pl_new   = True
                    pl_title = N(pl_row("TITLE"))
                    pl_type  = ToInt(pl_row("TYPE"))
                    pl_photo = N(pl_row("PHOTO"))
                Else
                    pl_err = "No tienes permisos para editar esa playlist."
                End If
            End If
            pl_row.Close
        End If
    End If
    Set pl_row = Nothing
End If

If IsEmpty(pl_title) Then pl_title = ""
If IsEmpty(pl_photo) Then pl_photo = "~/playlistImages/default.jpg"
%>
<!--#include file="inc/layout_head.asp"-->

<div class="page-head">
  <div class="page-head__txt">
    <h2><% If g_IsAdmin Then %>Playlists<% Else %>Mis playlists<% End If %></h2>
    <p>Crea, edita y elimina tus listas de reproducción. Puedes hacerlas públicas o privadas.</p>
  </div>
  <div class="page-head__actions">
    <a class="btn btn--primary" href="<%=AppUrl("playlist.asp?action=new")%>"><%=Icon("plus")%>Nueva playlist</a>
  </div>
</div>

<% If Len(pl_err) > 0 Then %>
<div class="alert alert--error"><span class="alert__dot"></span><span><%=H(pl_err)%></span></div>
<% End If %>

<% If pl_new Then %>
<section class="panel panel--narrow mb-24">
  <div class="panel__head">
    <h3><% If pl_id > 0 Then %>Editar playlist<% Else %>Nueva playlist<% End If %></h3>
    <p>Las playlists privadas solo son visibles para ti; las públicas pueden verlas todos los usuarios.</p>
  </div>
  <form method="post" action="<%=AppUrl("playlists.asp")%>" enctype="multipart/form-data">
    <input type="hidden" name="action" value="save">
    <input type="hidden" name="id" value="<%=pl_id%>">
    <div class="form-grid">
      <div class="field">
        <label for="title">Título *</label>
        <input class="input" type="text" id="title" name="title" value="<%=H(pl_title)%>" maxlength="50" required>
      </div>
      <div class="field">
        <label for="pl_type">Visibilidad</label>
        <select class="select" id="pl_type" name="pl_type">
          <option value="0" <% If ToInt(pl_type) = PL_PUBLIC Then %>selected<% End If %>>Pública - visible para todos</option>
          <option value="1" <% If ToInt(pl_type) = PL_PRIVATE Then %>selected<% End If %>>Privada - solo para ti</option>
        </select>
      </div>
      <div class="field field--full">
        <label for="photo_path">Ruta de la portada</label>
        <input class="input" type="text" id="photo_path" name="photo_path" value="<%=H(pl_photo)%>" maxlength="255" placeholder="~/playlistImages/playlist1.png">
      </div>
      <div class="field field--full">
        <label for="photo_file">Subir portada <span class="hint">(opcional, max. 6 MB)</span></label>
        <input class="input" type="file" id="photo_file" name="photo_file" accept="image/*">
      </div>
    </div>
    <div class="form-actions">
      <button class="btn btn--primary" type="submit"><%=Icon("check")%>Guardar</button>
      <a class="btn btn--ghost" href="<%=AppUrl("playlists.asp")%>">Cancelar</a>
    </div>
  </form>
</section>
<% End If %>

<section class="section">
  <div class="section__head"><h3>Tus playlists</h3></div>
  <%
  Set pl_rs = DB_RS("SELECT p.ID_PLAYLIST, p.TITLE, p.TYPE, p.PHOTO, p.ID_USER, p.CREATION_DATE, " & _
                    "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG s WHERE s.ID_PLAYLIST = p.ID_PLAYLIST) AS N_SONGS " & _
                    "FROM TBL_PLAYLIST p WHERE p.ID_USER = @p0 ORDER BY p.CREATION_DATE DESC, p.ID_PLAYLIST DESC", Array(g_UserID))
  If IsObject(pl_rs) Then
      If Not (pl_rs Is Nothing) Then
          If pl_rs.EOF Then
              Response.Write EmptyState("playlist", "Todavía no tienes playlists", "Crea tu primera lista de reproducción para organizar tus canciones.", "<a class=""btn btn--primary"" href=""" & AppUrl("playlist.asp?action=new") & """>" & Icon("plus") & "Crear playlist</a>")
          Else
              Response.Write "<div class=""grid"">"
              Do While Not pl_rs.EOF
                  pl_count = ToInt(pl_rs("N_SONGS"))
                  pl_acts = "<div class=""card__actions"">" & _
                            EditLink(AppUrl("playlists.asp?action=new&id=" & ToInt(pl_rs("ID_PLAYLIST"))), "Editar playlist") & _
                            DeleteForm(AppUrl("playlists.asp"), "delete", "id", ToInt(pl_rs("ID_PLAYLIST")), "¿Eliminar la playlist seleccionada?") & _
                            "</div>"
                  Response.Write "<article class=""card"">" & _
                      "<a class=""card__link"" href=""" & AppUrl("playlist.asp?id=" & ToInt(pl_rs("ID_PLAYLIST"))) & """ aria-label=""" & H(pl_rs("TITLE")) & """></a>" & _
                      "<div class=""card__art""><img src=""" & H(AssetUrl(pl_rs("PHOTO"), "playlist")) & """ alt="""" loading=""lazy""></div>" & _
                      "<div class=""card__title"">" & H(pl_rs("TITLE")) & "</div>" & _
                      "<div class=""card__sub"">" & H(PlaylistTypeName(pl_rs("TYPE"))) & " &middot; " & pl_count & " canciones</div>" & _
                      pl_acts & "</article>"
                  pl_rs.MoveNext
              Loop
              Response.Write "</div>"
          End If
          pl_rs.Close
      End If
  End If
  Set pl_rs = Nothing
  %>
</section>

<section class="section">
  <div class="section__head"><h3>Playlists públicas de otros usuarios</h3></div>
  <%
  Set pl_rs = DB_RS("SELECT TOP 12 p.ID_PLAYLIST, p.TITLE, p.TYPE, p.PHOTO, u.USERNAME, " & _
                    "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG s WHERE s.ID_PLAYLIST = p.ID_PLAYLIST) AS N_SONGS " & _
                    "FROM TBL_PLAYLIST p INNER JOIN TBL_USER u ON u.ID_USER = p.ID_USER " & _
                    "WHERE p.TYPE = 0 AND p.ID_USER <> @p0 ORDER BY p.CREATION_DATE DESC", Array(g_UserID))
  If IsObject(pl_rs) Then
      If Not (pl_rs Is Nothing) Then
          If pl_rs.EOF Then
              Response.Write "<p class=""muted small"">Todavía no hay playlists públicas de otros usuarios.</p>"
          Else
              Response.Write "<div class=""grid"">"
              Do While Not pl_rs.EOF
                  Response.Write PlaylistCardHtml(pl_rs("ID_PLAYLIST"), pl_rs("TITLE"), pl_rs("TYPE"), pl_rs("PHOTO"), pl_rs("USERNAME"))
                  pl_rs.MoveNext
              Loop
              Response.Write "</div>"
          End If
          pl_rs.Close
      End If
  End If
  Set pl_rs = Nothing
  %>
</section>

<% If g_IsAdmin Then %>
<section class="section">
  <div class="section__head"><h3>Todas las playlists (administración)</h3></div>
  <div class="tablewrap">
    <table class="table">
      <thead>
        <tr><th>Playlist</th><th>Propietario</th><th>Tipo</th><th>Creada</th><th>Canciones</th><th style="text-align:right">Acciones</th></tr>
      </thead>
      <tbody>
        <%
        Set pl_rs = DB_RS("SELECT p.ID_PLAYLIST, p.TITLE, p.TYPE, p.PHOTO, p.CREATION_DATE, u.USERNAME, " & _
                          "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG s WHERE s.ID_PLAYLIST = p.ID_PLAYLIST) AS N_SONGS " & _
                          "FROM TBL_PLAYLIST p INNER JOIN TBL_USER u ON u.ID_USER = p.ID_USER " & _
                          "ORDER BY p.ID_PLAYLIST DESC", Array())
        If IsObject(pl_rs) Then
            If Not (pl_rs Is Nothing) Then
                Do While Not pl_rs.EOF
        %>
        <tr>
          <td>
            <div class="cell-main">
              <img class="thumb" src="<%=AssetUrl(pl_rs("PHOTO"), "playlist")%>" alt="">
              <span class="cell-main__txt"><b><%=H(pl_rs("TITLE"))%></b><small>#<%=ToInt(pl_rs("ID_PLAYLIST"))%></small></span>
            </div>
          </td>
          <td><%=H(pl_rs("USERNAME"))%></td>
          <td><span class="badge <%=IIf(ToInt(pl_rs("TYPE")) = PL_PRIVATE, "badge--priv", "badge--pub")%>"><%=H(PlaylistTypeName(pl_rs("TYPE")))%></span></td>
          <td class="nowrap"><%=FmtDate(pl_rs("CREATION_DATE"))%></td>
          <td><%=ToInt(pl_rs("N_SONGS"))%></td>
          <td>
            <div class="table__actions">
              <a class="icon-btn" href="<%=AppUrl("playlist.asp?id=" & ToInt(pl_rs("ID_PLAYLIST")))%>" title="Abrir"><%=Icon("chevron-right")%></a>
              <a class="icon-btn" href="<%=AppUrl("playlists.asp?action=new&id=" & ToInt(pl_rs("ID_PLAYLIST")))%>" title="Editar"><%=Icon("edit")%></a>
              <%=DeleteForm(AppUrl("playlists.asp"), "delete", "id", ToInt(pl_rs("ID_PLAYLIST")), "¿Eliminar esta playlist?")%>
            </div>
          </td>
        </tr>
        <%
                pl_rs.MoveNext
                Loop
                pl_rs.Close
            End If
        End If
        Set pl_rs = Nothing
        %>
      </tbody>
    </table>
  </div>
</section>
<% End If %>

<!--#include file="inc/layout_foot.asp"-->
