<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="inc/boot.asp"-->
<%
' =============================================================================
'  PoliMusic - Albumes
'  Consulta para todos los usuarios; CRUD exclusivo del administrador.
' =============================================================================
Dim ab_rs, ab_frm, ab_act, ab_id, ab_err, ab_q, ab_row, ab_new, ab_acts
Dim ab_title, ab_year, ab_img, ab_up, ab_count, ab_where, ab_prms

RequireLogin
InitPage "Álbumes", "albums"

ab_act = ""
ab_err = ""
ab_id  = QInt("id")
ab_q   = QStr("q")
ab_new = (LCase(QStr("action")) = "new")
ab_img = "~/albumImages/"
ab_where = ""
ab_prms = Array()

If IsPost() Then
    Set ab_frm = New clsForm
    ab_frm.Load
    ab_act = LCase(ab_frm.Val("action"))

    If Not g_IsAdmin Then
        ab_err = "Solo el administrador puede modificar los álbumes."
    ElseIf ab_act = "save" Then
        ab_new   = True
        ab_id    = ToIntDef(ab_frm.Val("id"), 0)
        ab_title = Trim(ab_frm.Val("title_album"))
        ab_year  = Trim(ab_frm.Val("release_year"))
        ab_img   = Trim(ab_frm.Val("album_image_path"))

        If Len(ab_title) = 0 Then
            ab_err = "El título del álbum es obligatorio."
        ElseIf Len(ab_title) > 50 Then
            ab_err = "El título no puede superar los 50 caracteres."
        ElseIf Len(ab_year) > 0 And Not IsNumeric(ab_year) Then
            ab_err = "El año de lanzamiento debe ser numérico."
        ElseIf Len(ab_year) > 0 And (ToInt(ab_year) < 1900 Or ToInt(ab_year) > 2100) Then
            ab_err = "El año de lanzamiento debe estar entre 1900 y 2100."
        Else
            If ab_frm.HasFile("album_image") Then
                ab_up = ab_frm.SaveFile("album_image", "uploads/albumImages", "album", "|.jpg|.jpeg|.png|.gif|.webp|.svg|", 6 * 1024 * 1024)
                If Len(ab_up) > 0 Then
                    ab_img = ab_up
                ElseIf Len(ab_frm.Errors) > 0 Then
                    ab_err = ab_frm.Errors
                End If
            End If

            If Len(ab_err) = 0 Then
                If Len(ab_img) = 0 Then ab_img = "~/albumImages/album1.png"
                If ab_id > 0 Then
                    DB_Run "UPDATE TBL_ALBUM SET TITLE_ALBUM = @p0, RELEASE_YEAR = @p1, ALBUM_IMAGE_PATH = @p2 WHERE ID_ALBUM = @p3", _
                           Array(ab_title, IIfNull(ab_year), ab_img, ab_id)
                    Flash "Álbum actualizado correctamente.", "ok"
                Else
                    DB_Run "INSERT INTO TBL_ALBUM (TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH) VALUES (@p0, @p1, @p2)", _
                           Array(ab_title, IIfNull(ab_year), ab_img)
                    Flash "Álbum creado correctamente.", "ok"
                End If
                Go AppUrl("albums.asp")
            End If
        End If

    ElseIf ab_act = "delete" Then
        ab_id = ToIntDef(ab_frm.Val("id"), 0)
        If ab_id > 0 Then
            DB_Run "DELETE FROM TBL_TRACKLIST WHERE ID_ALBUM = @p0", Array(ab_id)
            DB_Run "DELETE FROM TBL_ALBUM WHERE ID_ALBUM = @p0", Array(ab_id)
            Flash "Álbum eliminado.", "ok"
        End If
        Go AppUrl("albums.asp")
    End If
End If

If ab_id > 0 And g_IsAdmin And Len(ab_err) = 0 Then
    ab_new = True
    Set ab_row = DB_RS("SELECT TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM WHERE ID_ALBUM = @p0", Array(ab_id))
    If IsObject(ab_row) Then
        If Not (ab_row Is Nothing) Then
            If Not ab_row.EOF Then
                ab_title = N(ab_row("TITLE_ALBUM"))
                ab_year  = N(ab_row("RELEASE_YEAR"))
                ab_img   = N(ab_row("ALBUM_IMAGE_PATH"))
            End If
            ab_row.Close
        End If
    End If
    Set ab_row = Nothing
End If
%>
<!--#include file="inc/layout_head.asp"-->

<div class="page-head">
  <div class="page-head__txt">
    <h2>Álbumes</h2>
    <p>Explora la discografía disponible en el catálogo.</p>
  </div>
  <div class="page-head__actions">
    <form class="searchbox" action="<%=AppUrl("albums.asp")%>" method="get" role="search">
      <span class="searchbox__ic"><%=Icon("search")%></span>
      <input type="search" name="q" value="<%=H(ab_q)%>" placeholder="Buscar álbum">
    </form>
    <% If g_IsAdmin Then %>
    <a class="btn btn--primary" href="<%=AppUrl("albums.asp?action=new")%>"><%=Icon("plus")%>Nuevo álbum</a>
    <% End If %>
  </div>
</div>

<% If Len(ab_err) > 0 Then %>
<div class="alert alert--error"><span class="alert__dot"></span><span><%=H(ab_err)%></span></div>
<% End If %>

<% If ab_new And g_IsAdmin Then %>
<section class="panel panel--narrow mb-24">
  <div class="panel__head">
    <h3><% If ab_id > 0 Then %>Editar álbum<% Else %>Nuevo álbum<% End If %></h3>
    <p>Puedes conservar la imagen actual o subir una nueva.</p>
  </div>
  <form method="post" action="<%=AppUrl("albums.asp")%>" enctype="multipart/form-data">
    <input type="hidden" name="action" value="save">
    <input type="hidden" name="id" value="<%=ab_id%>">
    <div class="form-grid">
      <div class="field">
        <label for="title_album">Título del álbum *</label>
        <input class="input" type="text" id="title_album" name="title_album" value="<%=H(ab_title)%>" maxlength="50" required>
      </div>
      <div class="field">
        <label for="release_year">Año de lanzamiento</label>
        <input class="input" type="number" id="release_year" name="release_year" value="<%=H(ab_year)%>" min="1900" max="2100" step="1">
      </div>
      <div class="field field--full">
        <label for="album_image_path">Ruta de la portada</label>
        <input class="input" type="text" id="album_image_path" name="album_image_path" value="<%=H(ab_img)%>" maxlength="255" placeholder="~/albumImages/album1.png">
      </div>
      <div class="field field--full">
        <label for="album_image">Subir portada <span class="hint">(opcional, max. 6 MB)</span></label>
        <input class="input" type="file" id="album_image" name="album_image" accept="image/*">
      </div>
    </div>
    <div class="form-actions">
      <button class="btn btn--primary" type="submit"><%=Icon("check")%>Guardar</button>
      <a class="btn btn--ghost" href="<%=AppUrl("albums.asp")%>">Cancelar</a>
    </div>
  </form>
</section>
<% End If %>

<%
ab_q = Trim(ab_q)
If Len(ab_q) > 0 Then
    ab_where = "WHERE TITLE_ALBUM LIKE @p0"
    ab_prms = Array("%" & SqlLike(ab_q) & "%")
End If

Set ab_rs = DB_RS("SELECT ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH, " & _
                  "(SELECT COUNT(*) FROM TBL_TRACKLIST t WHERE t.ID_ALBUM = TBL_ALBUM.ID_ALBUM) AS N_SONGS " & _
                  "FROM TBL_ALBUM " & ab_where & " ORDER BY ID_ALBUM", ab_prms)
If IsObject(ab_rs) Then
    If Not (ab_rs Is Nothing) Then
        If ab_rs.EOF Then
            Response.Write EmptyState("disc", "Sin álbumes", "No hay álbumes que coincidan con la búsqueda.", "")
        Else
            Response.Write "<div class=""grid"">"
            Do While Not ab_rs.EOF
                ab_count = ToInt(ab_rs("N_SONGS"))
                ab_acts = ""
                If g_IsAdmin Then
                    ab_acts = "<div class=""card__actions"">" & _
                              EditLink(AppUrl("albums.asp?action=new&id=" & ToInt(ab_rs("ID_ALBUM"))), "Editar álbum") & _
                              DeleteForm(AppUrl("albums.asp"), "delete", "id", ToInt(ab_rs("ID_ALBUM")), "¿Eliminar el álbum seleccionado?") & _
                              "</div>"
                End If
                Response.Write "<article class=""card"">" & _
                    "<a class=""card__link"" href=""" & AppUrl("album.asp?id=" & ToInt(ab_rs("ID_ALBUM"))) & """ aria-label=""" & H(ab_rs("TITLE_ALBUM")) & """></a>" & _
                    "<div class=""card__art""><img src=""" & H(AssetUrl(ab_rs("ALBUM_IMAGE_PATH"), "album")) & """ alt="""" loading=""lazy""></div>" & _
                    "<div class=""card__title"">" & H(ab_rs("TITLE_ALBUM")) & "</div>" & _
                    "<div class=""card__sub"">" & H(ab_rs("RELEASE_YEAR")) & " &middot; " & ab_count & " canciones</div>" & _
                    ab_acts & "</article>"
                ab_rs.MoveNext
            Loop
            Response.Write "</div>"
        End If
        ab_rs.Close
    End If
End If
Set ab_rs = Nothing
%>

<!--#include file="inc/layout_foot.asp"-->
