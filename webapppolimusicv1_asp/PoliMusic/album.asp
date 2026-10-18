<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="inc/boot.asp"-->
<%
' =============================================================================
'  PoliMusic - Detalle de un album (con gestion de su lista de canciones)
' =============================================================================
Dim al_id, al_frm, al_act, al_err, al_rs, al_row, al_n, al_acts, al_song
Dim al_title, al_year, al_img, al_count, al_av, al_avail

RequireLogin
InitPage "Album", "albums"

al_id  = QInt("id")
al_err = ""

If al_id <= 0 Then
    Flash "Álbum no especificado.", "error"
    Go AppUrl("albums.asp")
End If

' ---------------------------------------------------------------------------
'  Acciones
' ---------------------------------------------------------------------------
If IsPost() Then
    Set al_frm = New clsForm
    al_frm.Load
    al_act = LCase(al_frm.Val("action"))
    al_id  = ToIntDef(al_frm.Val("album"), al_id)

    If al_act = "addsong" Then
        If Not g_IsAdmin Then
            al_err = "Solo el administrador puede modificar la lista de canciones del álbum."
        Else
            al_song = ToIntDef(al_frm.Val("song"), 0)
            If al_song <= 0 Then
                al_err = "Selecciona una canción."
            ElseIf DB_Count("SELECT COUNT(*) FROM TBL_TRACKLIST WHERE ID_ALBUM = @p0 AND ID_SONG = @p1", Array(al_id, al_song)) > 0 Then
                al_err = "Esa canción ya forma parte del álbum."
            Else
                DB_Run "INSERT INTO TBL_TRACKLIST (ID_ALBUM, ID_SONG) VALUES (@p0, @p1)", Array(al_id, al_song)
                Flash "Canción agregada al álbum.", "ok"
                Go AppUrl("album.asp?id=" & al_id)
            End If
        End If

    ElseIf al_act = "removesong" Then
        If Not g_IsAdmin Then
            al_err = "Solo el administrador puede modificar la lista de canciones del álbum."
        Else
            al_song = ToIntDef(al_frm.Val("song"), 0)
            DB_Run "DELETE FROM TBL_TRACKLIST WHERE ID_ALBUM = @p0 AND ID_SONG = @p1", Array(al_id, al_song)
            Flash "Canción quitada del álbum.", "ok"
            Go AppUrl("album.asp?id=" & al_id)
        End If

    ElseIf al_act = "delete" Then
        If Not g_IsAdmin Then
            al_err = "Solo el administrador puede eliminar álbumes."
        Else
            DB_Run "DELETE FROM TBL_TRACKLIST WHERE ID_ALBUM = @p0", Array(al_id)
            DB_Run "DELETE FROM TBL_ALBUM WHERE ID_ALBUM = @p0", Array(al_id)
            Flash "Álbum eliminado.", "ok"
            Go AppUrl("albums.asp")
        End If
    End If
End If

' ---------------------------------------------------------------------------
'  Datos del album
' ---------------------------------------------------------------------------
Set al_row = DB_RS("SELECT TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM WHERE ID_ALBUM = @p0", Array(al_id))
If IsObject(al_row) Then
    If Not (al_row Is Nothing) Then
        If al_row.EOF Then
            al_row.Close
            Set al_row = Nothing
            Flash "El álbum solicitado no existe.", "error"
            Go AppUrl("albums.asp")
        Else
            al_title = N(al_row("TITLE_ALBUM"))
            al_year  = N(al_row("RELEASE_YEAR"))
            al_img   = N(al_row("ALBUM_IMAGE_PATH"))
        End If
        al_row.Close
    End If
End If
Set al_row = Nothing

al_count = DB_Count("SELECT COUNT(*) FROM TBL_TRACKLIST WHERE ID_ALBUM = @p0", Array(al_id))
InitPage al_title, "albums"
%>
<!--#include file="inc/layout_head.asp"-->

<a class="btn btn--ghost btn--sm mb-16" href="<%=AppUrl("albums.asp")%>"><%=Icon("arrow-left")%>Volver a álbumes</a>

<% If Len(al_err) > 0 Then %>
<div class="alert alert--error"><span class="alert__dot"></span><span><%=H(al_err)%></span></div>
<% End If %>

<div data-queue>
  <header class="detail">
    <img class="detail__art" src="<%=H(AssetUrl(al_img, "album"))%>" alt="<%=H(al_title)%>">
    <div class="detail__info">
      <div class="detail__kind">Álbum</div>
      <h1 class="detail__title"><%=H(al_title)%></h1>
      <div class="detail__meta">
        <% If Len(al_year) > 0 Then %><span><%=H(al_year)%></span><span class="dot"></span><% End If %>
        <span><%=al_count%> canciones</span>
      </div>
      <div class="detail__actions">
        <% If al_count > 0 Then %>
        <button type="button" class="btn btn--primary" data-playall><%=Icon("play")%>Reproducir álbum</button>
        <% End If %>
        <% If g_IsAdmin Then %>
        <a class="btn btn--ghost" href="<%=AppUrl("albums.asp?action=new&id=" & al_id)%>"><%=Icon("edit")%>Editar álbum</a>
        <%=DeleteForm(AppUrl("album.asp"), "delete", "album", al_id, "¿Eliminar este álbum y su lista de canciones?")%>
        <% End If %>
      </div>
    </div>
  </header>

  <section class="section">
    <%
    Set al_rs = DB_RS(SongSql(0, "WHERE s.ID_SONG IN (SELECT ID_SONG FROM TBL_TRACKLIST WHERE ID_ALBUM = @p0)", "ORDER BY s.ID_SONG"), Array(al_id))
    If IsObject(al_rs) Then
        If Not (al_rs Is Nothing) Then
            If al_rs.EOF Then
                Response.Write EmptyState("music", "Este álbum no tiene canciones", "El administrador puede agregar canciones a este álbum desde el panel de abajo.", "")
            Else
                Response.Write TracklistHead()
                Response.Write "<div class=""tracklist"">"
                al_n = 0
                Do While Not al_rs.EOF
                    al_n = al_n + 1
                    al_acts = AddToPlaylistBtn(al_rs("ID_SONG"), al_rs("SONG_NAME"))
                    If g_IsAdmin Then
                        al_acts = al_acts & DeleteForm(AppUrl("album.asp"), "removesong", "song", ToInt(al_rs("ID_SONG")), "¿Quitar esta canción del álbum?")
                    End If
                    Response.Write TrackRowHtml(al_rs, al_n, True, False, al_acts)
                    al_rs.MoveNext
                Loop
                Response.Write "</div>"
            End If
            al_rs.Close
        End If
    End If
    Set al_rs = Nothing
    %>
  </section>
</div>

<% If g_IsAdmin Then %>
<section class="panel panel--narrow mt-24">
  <div class="panel__head">
    <h3>Agregar canción al álbum</h3>
    <p>Selecciona una canción del catálogo para incorporarla a este álbum.</p>
  </div>
  <form method="post" action="<%=AppUrl("album.asp?id=" & al_id)%>">
    <input type="hidden" name="action" value="addsong">
    <input type="hidden" name="album" value="<%=al_id%>">
    <div class="form-grid">
      <div class="field field--full">
        <label for="song">Canción</label>
        <select class="select" id="song" name="song" required>
          <option value="">-- Selecciona una canción --</option>
          <%
          Set al_av = DB_RS("SELECT ID_SONG, SONG_NAME FROM TBL_SONG " & _
                            "WHERE ID_SONG NOT IN (SELECT ID_SONG FROM TBL_TRACKLIST WHERE ID_ALBUM = @p0) " & _
                            "ORDER BY SONG_NAME", Array(al_id))
          al_avail = 0
          If IsObject(al_av) Then
              If Not (al_av Is Nothing) Then
                  Do While Not al_av.EOF
                      al_avail = al_avail + 1
          %>
          <option value="<%=ToInt(al_av("ID_SONG"))%>"><%=H(al_av("SONG_NAME"))%></option>
          <%
                      al_av.MoveNext
                  Loop
                  al_av.Close
              End If
          End If
          Set al_av = Nothing
          %>
        </select>
        <% If al_avail = 0 Then %><span class="hint">Todas las canciones del catálogo ya están en este álbum.</span><% End If %>
      </div>
    </div>
    <div class="form-actions">
      <button class="btn btn--primary" type="submit" <% If al_avail = 0 Then %>disabled<% End If %>><%=Icon("plus")%>Agregar</button>
    </div>
  </form>
</section>
<% End If %>

<!--#include file="inc/layout_foot.asp"-->
