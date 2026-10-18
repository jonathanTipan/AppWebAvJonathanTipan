<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="inc/boot.asp"-->
<%
' =============================================================================
'  PoliMusic - Detalle de un artista
' =============================================================================
Dim at_id, at_frm, at_act, at_err, at_rs, at_row, at_n, at_acts, at_song
Dim at_name, at_last, at_img, at_count, at_av, at_avail

RequireLogin
InitPage "Artista", "artists"

at_id  = QInt("id")
at_err = ""

If at_id <= 0 Then
    Flash "Artista no especificado.", "error"
    Go AppUrl("artists.asp")
End If

If IsPost() Then
    Set at_frm = New clsForm
    at_frm.Load
    at_act = LCase(at_frm.Val("action"))
    at_id  = ToIntDef(at_frm.Val("artist"), at_id)

    If at_act = "addsong" Then
        If Not g_IsAdmin Then
            at_err = "Solo el administrador puede modificar el repertorio del artista."
        Else
            at_song = ToIntDef(at_frm.Val("song"), 0)
            If at_song <= 0 Then
                at_err = "Selecciona una canción."
            ElseIf DB_Count("SELECT COUNT(*) FROM TBL_PLAYER WHERE ID_ARTIST = @p0 AND ID_SONG = @p1", Array(at_id, at_song)) > 0 Then
                at_err = "Esa canción ya está asociada al artista."
            Else
                DB_Run "INSERT INTO TBL_PLAYER (ID_SONG, ID_ARTIST) VALUES (@p0, @p1)", Array(at_song, at_id)
                Flash "Canción asociada al artista.", "ok"
                Go AppUrl("artist.asp?id=" & at_id)
            End If
        End If

    ElseIf at_act = "removesong" Then
        If Not g_IsAdmin Then
            at_err = "Solo el administrador puede modificar el repertorio del artista."
        Else
            at_song = ToIntDef(at_frm.Val("song"), 0)
            DB_Run "DELETE FROM TBL_PLAYER WHERE ID_ARTIST = @p0 AND ID_SONG = @p1", Array(at_id, at_song)
            Flash "Canción desasociada del artista.", "ok"
            Go AppUrl("artist.asp?id=" & at_id)
        End If

    ElseIf at_act = "delete" Then
        If Not g_IsAdmin Then
            at_err = "Solo el administrador puede eliminar artistas."
        Else
            DB_Run "DELETE FROM TBL_PLAYER WHERE ID_ARTIST = @p0", Array(at_id)
            DB_Run "DELETE FROM TBL_ARTIST WHERE ID_ARTIST = @p0", Array(at_id)
            Flash "Artista eliminado.", "ok"
            Go AppUrl("artists.asp")
        End If
    End If
End If

Set at_row = DB_RS("SELECT ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST WHERE ID_ARTIST = @p0", Array(at_id))
If IsObject(at_row) Then
    If Not (at_row Is Nothing) Then
        If at_row.EOF Then
            at_row.Close
            Set at_row = Nothing
            Flash "El artista solicitado no existe.", "error"
            Go AppUrl("artists.asp")
        Else
            at_name = N(at_row("ARTIST_NAME"))
            at_last = N(at_row("ARTIST_LASTNAME"))
            at_img  = N(at_row("ARTIST_IMAGE"))
        End If
        at_row.Close
    End If
End If
Set at_row = Nothing

at_count = DB_Count("SELECT COUNT(*) FROM TBL_PLAYER WHERE ID_ARTIST = @p0", Array(at_id))
InitPage Trim(at_name & " " & at_last), "artists"
%>
<!--#include file="inc/layout_head.asp"-->

<a class="btn btn--ghost btn--sm mb-16" href="<%=AppUrl("artists.asp")%>"><%=Icon("arrow-left")%>Volver a artistas</a>

<% If Len(at_err) > 0 Then %>
<div class="alert alert--error"><span class="alert__dot"></span><span><%=H(at_err)%></span></div>
<% End If %>

<div data-queue>
  <header class="detail">
    <img class="detail__art detail__art--round" src="<%=H(AssetUrl(at_img, "artist"))%>" alt="<%=H(at_name)%>">
    <div class="detail__info">
      <div class="detail__kind">Artista</div>
      <h1 class="detail__title"><%=H(Trim(at_name & " " & at_last))%></h1>
      <div class="detail__meta"><span><%=at_count%> canciones en el catálogo</span></div>
      <div class="detail__actions">
        <% If at_count > 0 Then %>
        <button type="button" class="btn btn--primary" data-playall><%=Icon("play")%>Reproducir</button>
        <% End If %>
        <% If g_IsAdmin Then %>
        <a class="btn btn--ghost" href="<%=AppUrl("artists.asp?action=new&id=" & at_id)%>"><%=Icon("edit")%>Editar artista</a>
        <%=DeleteForm(AppUrl("artist.asp"), "delete", "artist", at_id, "¿Eliminar este artista y sus asociaciones con canciones?")%>
        <% End If %>
      </div>
    </div>
  </header>

  <section class="section">
    <div class="section__head"><h3>Canciones</h3></div>
    <%
    Set at_rs = DB_RS(SongSql(0, "WHERE s.ID_SONG IN (SELECT ID_SONG FROM TBL_PLAYER WHERE ID_ARTIST = @p0)", "ORDER BY s.SONG_NAME"), Array(at_id))
    If IsObject(at_rs) Then
        If Not (at_rs Is Nothing) Then
            If at_rs.EOF Then
                Response.Write EmptyState("music", "Sin canciones asociadas", "Este artista todavía no tiene canciones en el catálogo.", "")
            Else
                Response.Write TracklistHead()
                Response.Write "<div class=""tracklist"">"
                at_n = 0
                Do While Not at_rs.EOF
                    at_n = at_n + 1
                    at_acts = AddToPlaylistBtn(at_rs("ID_SONG"), at_rs("SONG_NAME"))
                    If g_IsAdmin Then
                        at_acts = at_acts & DeleteForm(AppUrl("artist.asp"), "removesong", "song", ToInt(at_rs("ID_SONG")), "¿Quitar esta canción del artista?")
                    End If
                    Response.Write TrackRowHtml(at_rs, at_n, False, True, at_acts)
                    at_rs.MoveNext
                Loop
                Response.Write "</div>"
            End If
            at_rs.Close
        End If
    End If
    Set at_rs = Nothing
    %>
  </section>
</div>

<section class="section">
  <div class="section__head"><h3>Álbumes con sus canciones</h3></div>
  <%
  Set at_rs = DB_RS("SELECT DISTINCT al.ID_ALBUM, al.TITLE_ALBUM, al.RELEASE_YEAR, al.ALBUM_IMAGE_PATH " & _
                    "FROM TBL_ALBUM al " & _
                    "INNER JOIN TBL_TRACKLIST t ON t.ID_ALBUM = al.ID_ALBUM " & _
                    "INNER JOIN TBL_PLAYER p ON p.ID_SONG = t.ID_SONG " & _
                    "WHERE p.ID_ARTIST = @p0 ORDER BY al.ID_ALBUM", Array(at_id))
  If IsObject(at_rs) Then
      If Not (at_rs Is Nothing) Then
          If at_rs.EOF Then
              Response.Write "<p class=""muted small"">Este artista no aparece en ningún álbum.</p>"
          Else
              Response.Write "<div class=""grid"">"
              Do While Not at_rs.EOF
                  Response.Write AlbumCardHtml(at_rs("ID_ALBUM"), at_rs("TITLE_ALBUM"), at_rs("RELEASE_YEAR"), at_rs("ALBUM_IMAGE_PATH"))
                  at_rs.MoveNext
              Loop
              Response.Write "</div>"
          End If
          at_rs.Close
      End If
  End If
  Set at_rs = Nothing
  %>
</section>

<% If g_IsAdmin Then %>
<section class="panel panel--narrow mt-24">
  <div class="panel__head">
    <h3>Asociar una canción a este artista</h3>
    <p>La relacion se guarda en la tabla TBL_PLAYER.</p>
  </div>
  <form method="post" action="<%=AppUrl("artist.asp?id=" & at_id)%>">
    <input type="hidden" name="action" value="addsong">
    <input type="hidden" name="artist" value="<%=at_id%>">
    <div class="form-grid">
      <div class="field field--full">
        <label for="song">Canción</label>
        <select class="select" id="song" name="song" required>
          <option value="">-- Selecciona una canción --</option>
          <%
          Set at_av = DB_RS("SELECT ID_SONG, SONG_NAME FROM TBL_SONG " & _
                            "WHERE ID_SONG NOT IN (SELECT ID_SONG FROM TBL_PLAYER WHERE ID_ARTIST = @p0) " & _
                            "ORDER BY SONG_NAME", Array(at_id))
          at_avail = 0
          If IsObject(at_av) Then
              If Not (at_av Is Nothing) Then
                  Do While Not at_av.EOF
                      at_avail = at_avail + 1
          %>
          <option value="<%=ToInt(at_av("ID_SONG"))%>"><%=H(at_av("SONG_NAME"))%></option>
          <%
                      at_av.MoveNext
                  Loop
                  at_av.Close
              End If
          End If
          Set at_av = Nothing
          %>
        </select>
        <% If at_avail = 0 Then %><span class="hint">Todas las canciones del catálogo ya están asociadas.</span><% End If %>
      </div>
    </div>
    <div class="form-actions">
      <button class="btn btn--primary" type="submit" <% If at_avail = 0 Then %>disabled<% End If %>><%=Icon("plus")%>Asociar</button>
    </div>
  </form>
</section>
<% End If %>

<!--#include file="inc/layout_foot.asp"-->
