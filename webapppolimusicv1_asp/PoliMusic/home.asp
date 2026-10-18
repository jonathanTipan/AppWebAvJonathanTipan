<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="inc/boot.asp"-->
<%
' =============================================================================
'  PoliMusic - Pantalla principal
' =============================================================================
Dim hm_rs, hm_i, hm_nSongs, hm_nAlbums, hm_nArtists, hm_nPlaylists, hm_nUsers, hm_n

RequireLogin
InitPage "Inicio", "home"

hm_nSongs     = DB_Count("SELECT COUNT(*) FROM TBL_SONG", Array())
hm_nAlbums    = DB_Count("SELECT COUNT(*) FROM TBL_ALBUM", Array())
hm_nArtists   = DB_Count("SELECT COUNT(*) FROM TBL_ARTIST", Array())
hm_nPlaylists = DB_Count("SELECT COUNT(*) FROM TBL_PLAYLIST WHERE ID_USER = @p0", Array(g_UserID))
hm_nUsers     = DB_Count("SELECT COUNT(*) FROM TBL_USER", Array())
%>
<!--#include file="inc/layout_head.asp"-->

<div data-queue>

  <section class="hero">
    <div class="hero__eyebrow">Bienvenido de nuevo</div>
    <h2>Hola, <%=H(g_Username)%>.</h2>
    <p>
      <% If g_IsAdmin Then %>
      Tienes el rol <b>Administrador</b>: puedes gestionar usuarios, canciones, álbumes, artistas y playlists.
      <% Else %>
      Explora el catálogo, reproduce tus canciones favoritas y organiza tu música en playlists.
      <% End If %>
    </p>
    <div class="hero__actions">
      <button type="button" class="btn btn--primary" data-playall><%=Icon("play")%>Reproducir todo</button>
      <a class="btn btn--ghost" href="<%=AppUrl("songs.asp")%>"><%=Icon("music")%>Ver canciones</a>
      <% If g_IsAdmin Then %>
      <a class="btn btn--soft" href="<%=AppUrl("users.asp")%>"><%=Icon("users")%>Administrar usuarios</a>
      <% Else %>
      <a class="btn btn--soft" href="<%=AppUrl("playlists.asp")%>"><%=Icon("playlist")%>Mis playlists</a>
      <% End If %>
    </div>
  </section>

  <div class="stats mb-24">
    <div class="stat"><div class="stat__num"><%=hm_nSongs%></div><div class="stat__lbl">Canciones</div></div>
    <div class="stat"><div class="stat__num"><%=hm_nAlbums%></div><div class="stat__lbl">Álbumes</div></div>
    <div class="stat"><div class="stat__num"><%=hm_nArtists%></div><div class="stat__lbl">Artistas</div></div>
    <div class="stat"><div class="stat__num"><%=hm_nPlaylists%></div><div class="stat__lbl">Tus playlists</div></div>
    <% If g_IsAdmin Then %>
    <div class="stat"><div class="stat__num"><%=hm_nUsers%></div><div class="stat__lbl">Usuarios</div></div>
    <% End If %>
  </div>

  <section class="section">
    <div class="section__head">
      <h3>Canciones más escuchadas</h3>
      <a href="<%=AppUrl("songs.asp")%>">Ver todas</a>
    </div>
    <%
    Set hm_rs = DB_RS(SongSql(6, "", "ORDER BY s.PLAYS DESC, s.ID_SONG"), Array())
    hm_n = 0
    If IsObject(hm_rs) Then
        If Not (hm_rs Is Nothing) Then
            If hm_rs.EOF Then
                Response.Write EmptyState("music", "Todavía no hay canciones", "El catálogo está vacío. Inicia sesión como administrador para agregar canciones.", "")
            Else
                Response.Write TracklistHead()
                Response.Write "<div class=""tracklist"">"
                Do While Not hm_rs.EOF
                    hm_n = hm_n + 1
                    Response.Write TrackRowHtml(hm_rs, hm_n, True, True, AddToPlaylistBtn(hm_rs("ID_SONG"), hm_rs("SONG_NAME")))
                    hm_rs.MoveNext
                Loop
                Response.Write "</div>"
            End If
            hm_rs.Close
        End If
    End If
    Set hm_rs = Nothing
    %>
  </section>

</div>

<section class="section">
  <div class="section__head">
    <h3>Álbumes</h3>
    <a href="<%=AppUrl("albums.asp")%>">Ver todos</a>
  </div>
  <div class="grid">
    <%
    Set hm_rs = DB_RS("SELECT TOP 5 ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM ORDER BY ID_ALBUM", Array())
    If IsObject(hm_rs) Then
        If Not (hm_rs Is Nothing) Then
            Do While Not hm_rs.EOF
                Response.Write AlbumCardHtml(hm_rs("ID_ALBUM"), hm_rs("TITLE_ALBUM"), hm_rs("RELEASE_YEAR"), hm_rs("ALBUM_IMAGE_PATH"))
                hm_rs.MoveNext
            Loop
            hm_rs.Close
        End If
    End If
    Set hm_rs = Nothing
    %>
  </div>
</section>

<section class="section">
  <div class="section__head">
    <h3>Artistas</h3>
    <a href="<%=AppUrl("artists.asp")%>">Ver todos</a>
  </div>
  <div class="grid">
    <%
    Set hm_rs = DB_RS("SELECT TOP 5 ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST ORDER BY ID_ARTIST", Array())
    If IsObject(hm_rs) Then
        If Not (hm_rs Is Nothing) Then
            Do While Not hm_rs.EOF
                Response.Write ArtistCardHtml(hm_rs("ID_ARTIST"), Trim(N(hm_rs("ARTIST_NAME")) & " " & N(hm_rs("ARTIST_LASTNAME"))), hm_rs("ARTIST_IMAGE"))
                hm_rs.MoveNext
            Loop
            hm_rs.Close
        End If
    End If
    Set hm_rs = Nothing
    %>
  </div>
</section>

<section class="section">
  <div class="section__head">
    <h3>Tus playlists</h3>
    <a href="<%=AppUrl("playlists.asp")%>">Gestionar</a>
  </div>
  <%
  Set hm_rs = DB_RS("SELECT TOP 5 ID_PLAYLIST, TITLE, TYPE, PHOTO FROM TBL_PLAYLIST WHERE ID_USER = @p0 ORDER BY CREATION_DATE DESC, ID_PLAYLIST DESC", Array(g_UserID))
  If IsObject(hm_rs) Then
      If Not (hm_rs Is Nothing) Then
          If hm_rs.EOF Then
              Response.Write EmptyState("playlist", "Aún no tienes playlists", "Crea tu primera lista de reproducción y organiza tu música como quieras.", "<a class=""btn btn--primary"" href=""" & AppUrl("playlist.asp?action=new") & """>" & Icon("plus") & "Crear playlist</a>")
          Else
              Response.Write "<div class=""grid"">"
              Do While Not hm_rs.EOF
                  Response.Write PlaylistCardHtml(hm_rs("ID_PLAYLIST"), hm_rs("TITLE"), hm_rs("TYPE"), hm_rs("PHOTO"), g_Username)
                  hm_rs.MoveNext
              Loop
              Response.Write "</div>"
          End If
          hm_rs.Close
      End If
  End If
  Set hm_rs = Nothing
  %>
</section>

<!--#include file="inc/layout_foot.asp"-->
