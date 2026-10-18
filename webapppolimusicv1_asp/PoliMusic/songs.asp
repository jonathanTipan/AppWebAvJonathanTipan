<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="inc/boot.asp"-->
<%
' =============================================================================
'  PoliMusic - Canciones
'  Todos los usuarios pueden consultar y reproducir.
'  Solo el administrador puede crear, editar y eliminar.
' =============================================================================
Dim sg_rs, sg_frm, sg_act, sg_id, sg_err, sg_q, sg_n, sg_row, sg_new
Dim sg_name, sg_path, sg_plays, sg_upPath, sg_acts, sg_where, sg_prms

RequireLogin
InitPage "Canciones", "songs"

sg_act = ""
sg_err = ""
sg_id  = QInt("id")
sg_q   = QStr("q")
sg_new = (LCase(QStr("action")) = "new")
sg_path = "~/songFiles/"
sg_where = ""
sg_prms = Array()

' ---------------------------------------------------------------------------
'  Acciones (solo administrador)
' ---------------------------------------------------------------------------
If IsPost() Then
    Set sg_frm = New clsForm
    sg_frm.Load
    sg_act = LCase(sg_frm.Val("action"))

    If Not g_IsAdmin Then
        sg_err = "Solo el administrador puede modificar el catálogo de canciones."
    ElseIf sg_act = "save" Then
        sg_new   = True
        sg_id    = ToIntDef(sg_frm.Val("id"), 0)
        sg_name  = Trim(sg_frm.Val("song_name"))
        sg_path  = Trim(sg_frm.Val("song_path"))
        sg_plays = sg_frm.Val("plays")

        If Len(sg_name) = 0 Then
            sg_err = "El nombre de la canción es obligatorio."
        ElseIf Len(sg_name) > 50 Then
            sg_err = "El nombre no puede superar los 50 caracteres."
        Else
            If sg_frm.HasFile("song_file") Then
                sg_upPath = sg_frm.SaveFile("song_file", "uploads/songFiles", "song", "|.mp3|.wav|.ogg|.m4a|.aac|.wma|", 40 * 1024 * 1024)
                If Len(sg_upPath) > 0 Then
                    sg_path = sg_upPath
                ElseIf Len(sg_frm.Errors) > 0 Then
                    sg_err = sg_frm.Errors
                End If
            End If

            If Len(sg_err) = 0 And Len(sg_path) = 0 Then
                sg_err = "Indica la ruta del archivo de audio o sube un archivo."
            End If

            If Len(sg_err) = 0 Then
                If Len(sg_path) > 255 Then
                    sg_err = "La ruta del archivo es demasiado larga (máximo 255 caracteres)."
                ElseIf ToIntDef(sg_plays, 0) < 0 Then
                    sg_err = "El número de reproducciones no puede ser negativo."
                ElseIf sg_id > 0 Then
                    DB_Run "UPDATE TBL_SONG SET SONG_NAME = @p0, SONG_PATH = @p1, PLAYS = @p2 WHERE ID_SONG = @p3", _
                           Array(sg_name, sg_path, ToIntDef(sg_plays, 0), sg_id)
                    Flash "Canción actualizada correctamente.", "ok"
                    Go AppUrl("songs.asp")
                Else
                    DB_Run "INSERT INTO TBL_SONG (SONG_NAME, SONG_PATH, PLAYS) VALUES (@p0, @p1, @p2)", _
                           Array(sg_name, sg_path, ToIntDef(sg_plays, 0))
                    Flash "Canción creada correctamente.", "ok"
                    Go AppUrl("songs.asp")
                End If
            End If
        End If

    ElseIf sg_act = "delete" Then
        sg_id = ToIntDef(sg_frm.Val("id"), 0)
        If sg_id > 0 Then
            DB_Run "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_SONG = @p0", Array(sg_id)
            DB_Run "DELETE FROM TBL_TRACKLIST WHERE ID_SONG = @p0", Array(sg_id)
            DB_Run "DELETE FROM TBL_PLAYER WHERE ID_SONG = @p0", Array(sg_id)
            DB_Run "DELETE FROM TBL_SONG WHERE ID_SONG = @p0", Array(sg_id)
            Flash "Canción eliminada.", "ok"
        End If
        Go AppUrl("songs.asp")
    End If
End If

' ---------------------------------------------------------------------------
'  Formulario de edicion
' ---------------------------------------------------------------------------
If sg_id > 0 And g_IsAdmin And Len(sg_err) = 0 Then
    sg_new = True
    Set sg_row = DB_RS("SELECT SONG_NAME, SONG_PATH, PLAYS FROM TBL_SONG WHERE ID_SONG = @p0", Array(sg_id))
    If IsObject(sg_row) Then
        If Not (sg_row Is Nothing) Then
            If Not sg_row.EOF Then
                sg_name  = N(sg_row("SONG_NAME"))
                sg_path  = N(sg_row("SONG_PATH"))
                sg_plays = ToInt(sg_row("PLAYS"))
            End If
            sg_row.Close
        End If
    End If
    Set sg_row = Nothing
End If

If IsEmpty(sg_plays) Then sg_plays = 0
%>
<!--#include file="inc/layout_head.asp"-->

<div class="page-head">
  <div class="page-head__txt">
    <h2>Canciones</h2>
    <p>Catálogo completo de audio de <%=APP_NAME%>.</p>
  </div>
  <div class="page-head__actions">
    <form class="searchbox" action="<%=AppUrl("songs.asp")%>" method="get" role="search">
      <span class="searchbox__ic"><%=Icon("search")%></span>
      <input type="search" name="q" value="<%=H(sg_q)%>" placeholder="Buscar por canción o artista">
    </form>
    <% If g_IsAdmin Then %>
    <a class="btn btn--primary" href="<%=AppUrl("songs.asp?action=new")%>"><%=Icon("plus")%>Nueva canción</a>
    <% End If %>
  </div>
</div>

<% If Not g_IsAdmin Then %>
<div class="alert alert--warn"><span class="alert__dot"></span><span>Estás en modo consulta. La gestión del catálogo es exclusiva del usuario administrador.</span></div>
<% End If %>

<% If Len(sg_err) > 0 Then %>
<div class="alert alert--error"><span class="alert__dot"></span><span><%=H(sg_err)%></span></div>
<% End If %>

<% If sg_new And g_IsAdmin Then %>
<section class="panel panel--narrow mb-24">
  <div class="panel__head">
    <h3><% If sg_id > 0 Then %>Editar canción<% Else %>Nueva canción<% End If %></h3>
    <p>El nombre es obligatorio. Puedes indicar la ruta de un archivo existente o subir uno nuevo.</p>
  </div>
  <form method="post" action="<%=AppUrl("songs.asp")%>" enctype="multipart/form-data">
    <input type="hidden" name="action" value="save">
    <input type="hidden" name="id" value="<%=sg_id%>">
    <div class="form-grid">
      <div class="field">
        <label for="song_name">Nombre de la canción *</label>
        <input class="input" type="text" id="song_name" name="song_name" value="<%=H(sg_name)%>" maxlength="50" required>
      </div>
      <div class="field">
        <label for="plays">Reproducciones</label>
        <input class="input" type="number" id="plays" name="plays" value="<%=H(sg_plays)%>" min="0" step="1">
      </div>
      <div class="field field--full">
        <label for="song_path">Ruta del archivo de audio</label>
        <input class="input" type="text" id="song_path" name="song_path" value="<%=H(sg_path)%>" maxlength="255" placeholder="~/songFiles/bensound-adventure.mp3">
        <span class="hint">Ejemplo: <span class="mono">~/songFiles/bensound-adventure.mp3</span>. Se ignora si subes un archivo abajo.</span>
      </div>
      <div class="field field--full">
        <label for="song_file">Subir archivo de audio <span class="hint">(opcional, max. 40 MB)</span></label>
        <input class="input" type="file" id="song_file" name="song_file" accept="audio/*">
      </div>
    </div>
    <div class="form-actions">
      <button class="btn btn--primary" type="submit"><%=Icon("check")%>Guardar</button>
      <a class="btn btn--ghost" href="<%=AppUrl("songs.asp")%>">Cancelar</a>
    </div>
  </form>
</section>
<% End If %>

<%
sg_q = Trim(sg_q)
If Len(sg_q) > 0 Then
    sg_where = "WHERE (s.SONG_NAME LIKE @p0 OR ar.ARTIST_NAME LIKE @p0)"
    sg_prms = Array("%" & SqlLike(sg_q) & "%")
End If

Set sg_rs = DB_RS(SongSql(0, sg_where, "ORDER BY s.ID_SONG"), sg_prms)
If IsObject(sg_rs) Then
    If Not (sg_rs Is Nothing) Then
        If sg_rs.EOF Then
            Response.Write EmptyState("music", "Sin resultados", "No se encontraron canciones con ese criterio de búsqueda.", "")
        Else
            Response.Write TracklistHead()
            Response.Write "<div class=""tracklist"" data-queue>"
            sg_n = 0
            Do While Not sg_rs.EOF
                sg_n = sg_n + 1
                If g_IsAdmin Then
                    sg_acts = EditLink(AppUrl("songs.asp?action=new&id=" & ToInt(sg_rs("ID_SONG"))), "Editar canción") & _
                              DeleteForm(AppUrl("songs.asp"), "delete", "id", ToInt(sg_rs("ID_SONG")), "¿Eliminar la canción seleccionada?")
                Else
                    sg_acts = AddToPlaylistBtn(sg_rs("ID_SONG"), sg_rs("SONG_NAME"))
                End If
                Response.Write TrackRowHtml(sg_rs, sg_n, True, True, sg_acts)
                sg_rs.MoveNext
            Loop
            Response.Write "</div>"
            Response.Write "<p class=""muted small mt-16"">" & sg_n & " canciones encontradas.</p>"
        End If
        sg_rs.Close
    End If
End If
Set sg_rs = Nothing
%>

<!--#include file="inc/layout_foot.asp"-->
