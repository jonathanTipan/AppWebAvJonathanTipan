<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="inc/boot.asp"-->
<%
' =============================================================================
'  PoliMusic - Detalle de una playlist
'  El propietario (y el administrador) puede anadir y quitar canciones.
' =============================================================================
Dim pt_id, pt_frm, pt_act, pt_err, pt_rs, pt_row, pt_n, pt_acts, pt_song
Dim pt_title, pt_type, pt_photo, pt_owner, pt_ownerName, pt_created, pt_count
Dim pt_canView, pt_canEdit, pt_addq, pt_like, pt_added, pt_where, pt_prms

RequireLogin
InitPage "Playlist", "playlists"

pt_id   = QInt("id")
pt_err  = ""
pt_addq = Trim(QStr("addq"))
pt_canView = False
pt_canEdit = False

' ---------------------------------------------------------------------------
'  Acciones
' ---------------------------------------------------------------------------
If IsPost() Then
    Set pt_frm = New clsForm
    pt_frm.Load
    pt_act = LCase(pt_frm.Val("action"))
    pt_id  = ToIntDef(pt_frm.Val("playlist"), pt_id)
    pt_song = ToIntDef(pt_frm.Val("song"), 0)

    ' Comprueba propiedad
    Set pt_row = DB_RS("SELECT ID_USER FROM TBL_PLAYLIST WHERE ID_PLAYLIST = @p0", Array(pt_id))
    If IsObject(pt_row) Then
        If Not (pt_row Is Nothing) Then
            If Not pt_row.EOF Then
                pt_owner = ToInt(pt_row("ID_USER"))
                pt_canEdit = (pt_owner = g_UserID) Or g_IsAdmin
            End If
            pt_row.Close
        End If
    End If
    Set pt_row = Nothing

    If Not pt_canEdit Then
        pt_err = "Solo el propietario de la playlist o el administrador pueden modificarla."
    ElseIf pt_act = "addsong" Then
        If pt_song <= 0 Then
            pt_err = "Selecciona una canción."
        ElseIf DB_Count("SELECT COUNT(*) FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = @p0 AND ID_SONG = @p1", Array(pt_id, pt_song)) > 0 Then
            Flash "Esa canción ya está en la playlist.", "warn"
            Go AppUrl("playlist.asp?id=" & pt_id)
        Else
            DB_Run "INSERT INTO TBL_PLAYLIST_SONG (ID_PLAYLIST, ID_SONG) VALUES (@p0, @p1)", Array(pt_id, pt_song)
            Flash "Canción agregada a la playlist.", "ok"
            Go AppUrl("playlist.asp?id=" & pt_id)
        End If

    ElseIf pt_act = "removesong" Then
        DB_Run "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = @p0 AND ID_SONG = @p1", Array(pt_id, pt_song)
        Flash "Canción quitada de la playlist.", "ok"
        Go AppUrl("playlist.asp?id=" & pt_id)

    ElseIf pt_act = "delete" Then
        DB_Run "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = @p0", Array(pt_id)
        DB_Run "DELETE FROM TBL_PLAYLIST WHERE ID_PLAYLIST = @p0", Array(pt_id)
        Flash "Playlist eliminada.", "ok"
        Go AppUrl("playlists.asp")

    ElseIf pt_act = "clear" Then
        DB_Run "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = @p0", Array(pt_id)
        Flash "Se vació la playlist.", "ok"
        Go AppUrl("playlist.asp?id=" & pt_id)
    End If
End If

If pt_id <= 0 Then
    Flash "Playlist no especificada.", "error"
    Go AppUrl("playlists.asp")
End If

' ---------------------------------------------------------------------------
'  Carga de la playlist
' ---------------------------------------------------------------------------
Set pt_row = DB_RS("SELECT p.ID_PLAYLIST, p.ID_USER, p.TITLE, p.TYPE, p.PHOTO, p.CREATION_DATE, u.USERNAME " & _
                   "FROM TBL_PLAYLIST p INNER JOIN TBL_USER u ON u.ID_USER = p.ID_USER " & _
                   "WHERE p.ID_PLAYLIST = @p0", Array(pt_id))
If IsObject(pt_row) Then
    If Not (pt_row Is Nothing) Then
        If pt_row.EOF Then
            pt_row.Close
            Set pt_row = Nothing
            Flash "La playlist solicitada no existe.", "error"
            Go AppUrl("playlists.asp")
        Else
            pt_title     = N(pt_row("TITLE"))
            pt_type      = ToInt(pt_row("TYPE"))
            pt_photo     = N(pt_row("PHOTO"))
            pt_owner     = ToInt(pt_row("ID_USER"))
            pt_ownerName = N(pt_row("USERNAME"))
            pt_created   = pt_row("CREATION_DATE")
        End If
        pt_row.Close
    End If
End If
Set pt_row = Nothing

pt_canEdit = (pt_owner = g_UserID) Or g_IsAdmin
pt_canView = pt_canEdit Or (pt_type = PL_PUBLIC)

If Not pt_canView Then
    Flash "Esa playlist es privada.", "error"
    Go AppUrl("playlists.asp")
End If

pt_count = DB_Count("SELECT COUNT(*) FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = @p0", Array(pt_id))
InitPage pt_title, "playlists"
%>
<!--#include file="inc/layout_head.asp"-->

<a class="btn btn--ghost btn--sm mb-16" href="<%=AppUrl("playlists.asp")%>"><%=Icon("arrow-left")%>Volver a playlists</a>

<% If Len(pt_err) > 0 Then %>
<div class="alert alert--error"><span class="alert__dot"></span><span><%=H(pt_err)%></span></div>
<% End If %>

<div data-queue>
  <header class="detail">
    <img class="detail__art" src="<%=H(AssetUrl(pt_photo, "playlist"))%>" alt="<%=H(pt_title)%>">
    <div class="detail__info">
      <div class="detail__kind">Playlist &middot; <%=H(PlaylistTypeName(pt_type))%></div>
      <h1 class="detail__title"><%=H(pt_title)%></h1>
      <div class="detail__meta">
        <span><%=H(pt_ownerName)%></span><span class="dot"></span>
        <span><%=pt_count%> canciones</span><span class="dot"></span>
        <span>Creada el <%=FmtDate(pt_created)%></span>
      </div>
      <div class="detail__actions">
        <% If pt_count > 0 Then %>
        <button type="button" class="btn btn--primary" data-playall><%=Icon("play")%>Reproducir</button>
        <% End If %>
        <% If pt_canEdit Then %>
        <a class="btn btn--ghost" href="<%=AppUrl("playlists.asp?action=new&id=" & pt_id)%>"><%=Icon("edit")%>Editar datos</a>
        <form class="inline-form" method="post" action="<%=AppUrl("playlist.asp?id=" & pt_id)%>" onsubmit="return confirm('¿Vaciar la playlist? Se quitarán todas las canciones.');">
          <input type="hidden" name="action" value="clear">
          <input type="hidden" name="playlist" value="<%=pt_id%>">
          <button class="btn btn--ghost btn--sm" type="submit">Vaciar</button>
        </form>
        <%=DeleteForm(AppUrl("playlist.asp"), "delete", "playlist", pt_id, "¿Eliminar definitivamente esta playlist?")%>
        <% End If %>
      </div>
    </div>
  </header>

  <section class="section">
    <div class="section__head"><h3>Canciones</h3></div>
    <%
    Set pt_rs = DB_RS(SongSql(0, "WHERE s.ID_SONG IN (SELECT ID_SONG FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = @p0)", "ORDER BY s.SONG_NAME"), Array(pt_id))
    If IsObject(pt_rs) Then
        If Not (pt_rs Is Nothing) Then
            If pt_rs.EOF Then
                Response.Write EmptyState("music", "Esta playlist está vacía", "Agrega canciones desde el buscador de abajo o con el boton + de cualquier lista.", "")
            Else
                Response.Write TracklistHead()
                Response.Write "<div class=""tracklist"">"
                pt_n = 0
                Do While Not pt_rs.EOF
                    pt_n = pt_n + 1
                    pt_acts = AddToPlaylistBtn(pt_rs("ID_SONG"), pt_rs("SONG_NAME"))
                    If pt_canEdit Then
                        pt_acts = pt_acts & DeleteForm(AppUrl("playlist.asp"), "removesong", "song", ToInt(pt_rs("ID_SONG")), "¿Quitar esta canción de la playlist?")
                    End If
                    Response.Write TrackRowHtml(pt_rs, pt_n, True, True, pt_acts)
                    pt_rs.MoveNext
                Loop
                Response.Write "</div>"
            End If
            pt_rs.Close
        End If
    End If
    Set pt_rs = Nothing
    %>
  </section>
</div>

<% If pt_canEdit Then %>
<section class="panel mt-24">
  <div class="panel__head">
    <h3>Agregar canciones</h3>
    <p>Busca en el catálogo y agrega las canciones que quieras a esta playlist.</p>
  </div>

  <form class="filters" method="get" action="<%=AppUrl("playlist.asp")%>">
    <input type="hidden" name="id" value="<%=pt_id%>">
    <div class="field" style="flex:1 1 320px">
      <label for="addq">Buscar canción</label>
      <input class="input" type="search" id="addq" name="addq" value="<%=H(pt_addq)%>" placeholder="Nombre de la canción o artista">
    </div>
    <button class="btn btn--ghost" type="submit"><%=Icon("search")%>Buscar</button>
  </form>

  <%
  pt_where = "WHERE s.ID_SONG NOT IN (SELECT ID_SONG FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = @p0)"
  pt_prms = Array(pt_id)
  If Len(pt_addq) > 0 Then
      pt_like = "%" & SqlLike(pt_addq) & "%"
      pt_where = pt_where & " AND (s.SONG_NAME LIKE @p1 OR ar.ARTIST_NAME LIKE @p1)"
      pt_prms = Array(pt_id, pt_like)
  End If

  Set pt_rs = DB_RS(SongSql(25, pt_where, "ORDER BY s.SONG_NAME"), pt_prms)
  If IsObject(pt_rs) Then
      If Not (pt_rs Is Nothing) Then
          If pt_rs.EOF Then
              Response.Write "<p class=""muted small"">No hay canciones disponibles para agregar con ese criterio.</p>"
          Else
              pt_added = 0
              Response.Write "<div class=""tablewrap""><table class=""table""><thead><tr><th>Canción</th><th>Artista</th><th>Álbum</th><th style=""text-align:right"">Acción</th></tr></thead><tbody>"
              Do While Not pt_rs.EOF
                  pt_added = pt_added + 1
                  Response.Write "<tr>" & _
                      "<td><div class=""cell-main""><img class=""thumb"" src=""" & H(AssetUrl(pt_rs("ALBUM_IMAGE_PATH"), "album")) & """ alt=""""><span class=""cell-main__txt""><b>" & H(pt_rs("SONG_NAME")) & "</b><small>" & PlaysLabel(pt_rs("PLAYS")) & "</small></span></div></td>" & _
                      "<td>" & H(pt_rs("ARTIST_NAME")) & "</td>" & _
                      "<td>" & H(pt_rs("TITLE_ALBUM")) & "</td>" & _
                      "<td><div class=""table__actions"">" & _
                          "<form class=""inline-form"" method=""post"" action=""" & AppUrl("playlist.asp?id=" & pt_id) & """>" & _
                          "<input type=""hidden"" name=""action"" value=""addsong"">" & _
                          "<input type=""hidden"" name=""playlist"" value=""" & pt_id & """>" & _
                          "<input type=""hidden"" name=""song"" value=""" & ToInt(pt_rs("ID_SONG")) & """>" & _
                          "<button class=""btn btn--soft btn--sm"" type=""submit"">" & Icon("plus") & "Agregar</button>" & _
                          "</form></div></td></tr>"
                  pt_rs.MoveNext
              Loop
              Response.Write "</tbody></table></div>"
              If pt_added = 25 Then
                  Response.Write "<p class=""muted small mt-16"">Mostrando las primeras 25 coincidencias. Afina la búsqueda para ver más.</p>"
              End If
          End If
          pt_rs.Close
      End If
  End If
  Set pt_rs = Nothing
  %>
</section>
<% End If %>

<!--#include file="inc/layout_foot.asp"-->
