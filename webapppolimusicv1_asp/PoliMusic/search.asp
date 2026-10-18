<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="inc/boot.asp"-->
<%
' =============================================================================
'  PoliMusic - Busqueda global (canciones, artistas, albumes y playlists)
' =============================================================================
Dim se_q, se_like, se_rs, se_n, se_acts

RequireLogin
InitPage "Buscar", "search"

se_q = Trim(QStr("q"))
%>
<!--#include file="inc/layout_head.asp"-->

<div class="page-head">
  <div class="page-head__txt">
    <h2>Buscar</h2>
    <p>Encuentra canciones, artistas, álbumes y playlists.</p>
  </div>
</div>

<form class="searchbox mb-24" action="<%=AppUrl("search.asp")%>" method="get" role="search" style="width:100%;max-width:560px">
  <span class="searchbox__ic"><%=Icon("search")%></span>
  <input type="search" name="q" value="<%=H(se_q)%>" placeholder="¿Qué quieres escuchar hoy?" autofocus>
</form>

<% If Len(se_q) = 0 Then %>
  <%=EmptyState("search", "Empieza a escribir", "Busca por el nombre de una canción, un artista, un álbum o una playlist.", "")%>
<% Else %>
  <%
  se_like = "%" & SqlLike(se_q) & "%"
  %>

  <section class="section">
    <div class="section__head"><h3>Canciones</h3></div>
    <%
    Set se_rs = DB_RS(SongSql(10, "WHERE (s.SONG_NAME LIKE @p0 OR ar.ARTIST_NAME LIKE @p0)", "ORDER BY s.SONG_NAME"), Array(se_like))
    If IsObject(se_rs) Then
        If Not (se_rs Is Nothing) Then
            If se_rs.EOF Then
                Response.Write "<p class=""muted small"">Ninguna canción coincide con la búsqueda.</p>"
            Else
                Response.Write TracklistHead()
                Response.Write "<div class=""tracklist"" data-queue>"
                se_n = 0
                Do While Not se_rs.EOF
                    se_n = se_n + 1
                    Response.Write TrackRowHtml(se_rs, se_n, True, True, AddToPlaylistBtn(se_rs("ID_SONG"), se_rs("SONG_NAME")))
                    se_rs.MoveNext
                Loop
                Response.Write "</div>"
            End If
            se_rs.Close
        End If
    End If
    Set se_rs = Nothing
    %>
  </section>

  <section class="section">
    <div class="section__head"><h3>Artistas</h3></div>
    <%
    Set se_rs = DB_RS("SELECT TOP 10 ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST " & _
                      "WHERE (ARTIST_NAME LIKE @p0 OR ARTIST_LASTNAME LIKE @p0) ORDER BY ARTIST_NAME", Array(se_like))
    If IsObject(se_rs) Then
        If Not (se_rs Is Nothing) Then
            If se_rs.EOF Then
                Response.Write "<p class=""muted small"">Ningún artista coincide con la búsqueda.</p>"
            Else
                Response.Write "<div class=""grid"">"
                Do While Not se_rs.EOF
                    Response.Write ArtistCardHtml(se_rs("ID_ARTIST"), Trim(N(se_rs("ARTIST_NAME")) & " " & N(se_rs("ARTIST_LASTNAME"))), se_rs("ARTIST_IMAGE"))
                    se_rs.MoveNext
                Loop
                Response.Write "</div>"
            End If
            se_rs.Close
        End If
    End If
    Set se_rs = Nothing
    %>
  </section>

  <section class="section">
    <div class="section__head"><h3>Álbumes</h3></div>
    <%
    Set se_rs = DB_RS("SELECT TOP 10 ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM " & _
                      "WHERE TITLE_ALBUM LIKE @p0 ORDER BY TITLE_ALBUM", Array(se_like))
    If IsObject(se_rs) Then
        If Not (se_rs Is Nothing) Then
            If se_rs.EOF Then
                Response.Write "<p class=""muted small"">Ningún álbum coincide con la búsqueda.</p>"
            Else
                Response.Write "<div class=""grid"">"
                Do While Not se_rs.EOF
                    Response.Write AlbumCardHtml(se_rs("ID_ALBUM"), se_rs("TITLE_ALBUM"), se_rs("RELEASE_YEAR"), se_rs("ALBUM_IMAGE_PATH"))
                    se_rs.MoveNext
                Loop
                Response.Write "</div>"
            End If
            se_rs.Close
        End If
    End If
    Set se_rs = Nothing
    %>
  </section>

  <section class="section">
    <div class="section__head"><h3>Playlists</h3></div>
    <%
    Set se_rs = DB_RS("SELECT TOP 10 p.ID_PLAYLIST, p.TITLE, p.TYPE, p.PHOTO, u.USERNAME " & _
                      "FROM TBL_PLAYLIST p INNER JOIN TBL_USER u ON u.ID_USER = p.ID_USER " & _
                      "WHERE (p.TYPE = 0 OR p.ID_USER = @p0) AND p.TITLE LIKE @p1 ORDER BY p.TITLE", _
                      Array(g_UserID, se_like))
    If IsObject(se_rs) Then
        If Not (se_rs Is Nothing) Then
            If se_rs.EOF Then
                Response.Write "<p class=""muted small"">Ninguna playlist coincide con la búsqueda.</p>"
            Else
                Response.Write "<div class=""grid"">"
                Do While Not se_rs.EOF
                    Response.Write PlaylistCardHtml(se_rs("ID_PLAYLIST"), se_rs("TITLE"), se_rs("TYPE"), se_rs("PHOTO"), se_rs("USERNAME"))
                    se_rs.MoveNext
                Loop
                Response.Write "</div>"
            End If
            se_rs.Close
        End If
    End If
    Set se_rs = Nothing
    %>
  </section>
<% End If %>

<!--#include file="inc/layout_foot.asp"-->
