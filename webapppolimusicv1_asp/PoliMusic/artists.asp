<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001" %>
<!--#include file="inc/boot.asp"-->
<%
' =============================================================================
'  PoliMusic - Artistas
'  Consulta para todos los usuarios; CRUD exclusivo del administrador.
' =============================================================================
Dim ar_rs, ar_frm, ar_act, ar_id, ar_err, ar_q, ar_row, ar_new, ar_acts
Dim ar_name, ar_last, ar_img, ar_up, ar_count, ar_where, ar_prms

RequireLogin
InitPage "Artistas", "artists"

ar_act = ""
ar_err = ""
ar_id  = QInt("id")
ar_q   = QStr("q")
ar_new = (LCase(QStr("action")) = "new")
ar_img = "~/ArtistImages/"
ar_where = ""
ar_prms = Array()

If IsPost() Then
    Set ar_frm = New clsForm
    ar_frm.Load
    ar_act = LCase(ar_frm.Val("action"))

    If Not g_IsAdmin Then
        ar_err = "Solo el administrador puede modificar los artistas."
    ElseIf ar_act = "save" Then
        ar_new  = True
        ar_id   = ToIntDef(ar_frm.Val("id"), 0)
        ar_name = Trim(ar_frm.Val("artist_name"))
        ar_last = Trim(ar_frm.Val("artist_lastname"))
        ar_img  = Trim(ar_frm.Val("artist_image"))

        If Len(ar_name) = 0 Then
            ar_err = "El nombre del artista es obligatorio."
        ElseIf Len(ar_name) > 100 Then
            ar_err = "El nombre no puede superar los 100 caracteres."
        ElseIf Len(ar_last) > 100 Then
            ar_err = "El apellido no puede superar los 100 caracteres."
        Else
            If ar_frm.HasFile("artist_image_file") Then
                ar_up = ar_frm.SaveFile("artist_image_file", "uploads/artistImages", "artist", "|.jpg|.jpeg|.png|.gif|.webp|.svg|", 6 * 1024 * 1024)
                If Len(ar_up) > 0 Then
                    ar_img = ar_up
                ElseIf Len(ar_frm.Errors) > 0 Then
                    ar_err = ar_frm.Errors
                End If
            End If

            If Len(ar_err) = 0 Then
                If Len(ar_img) = 0 Then ar_img = "~/ArtistImages/artist1.jpg"
                If ar_id > 0 Then
                    DB_Run "UPDATE TBL_ARTIST SET ARTIST_NAME = @p0, ARTIST_LASTNAME = @p1, ARTIST_IMAGE = @p2 WHERE ID_ARTIST = @p3", _
                           Array(ar_name, IIfNull(ar_last), ar_img, ar_id)
                    Flash "Artista actualizado correctamente.", "ok"
                Else
                    DB_Run "INSERT INTO TBL_ARTIST (ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE) VALUES (@p0, @p1, @p2)", _
                           Array(ar_name, IIfNull(ar_last), ar_img)
                    Flash "Artista creado correctamente.", "ok"
                End If
                Go AppUrl("artists.asp")
            End If
        End If

    ElseIf ar_act = "delete" Then
        ar_id = ToIntDef(ar_frm.Val("id"), 0)
        If ar_id > 0 Then
            DB_Run "DELETE FROM TBL_PLAYER WHERE ID_ARTIST = @p0", Array(ar_id)
            DB_Run "DELETE FROM TBL_ARTIST WHERE ID_ARTIST = @p0", Array(ar_id)
            Flash "Artista eliminado.", "ok"
        End If
        Go AppUrl("artists.asp")
    End If
End If

If ar_id > 0 And g_IsAdmin And Len(ar_err) = 0 Then
    ar_new = True
    Set ar_row = DB_RS("SELECT ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST WHERE ID_ARTIST = @p0", Array(ar_id))
    If IsObject(ar_row) Then
        If Not (ar_row Is Nothing) Then
            If Not ar_row.EOF Then
                ar_name = N(ar_row("ARTIST_NAME"))
                ar_last = N(ar_row("ARTIST_LASTNAME"))
                ar_img  = N(ar_row("ARTIST_IMAGE"))
            End If
            ar_row.Close
        End If
    End If
    Set ar_row = Nothing
End If
%>
<!--#include file="inc/layout_head.asp"-->

<div class="page-head">
  <div class="page-head__txt">
    <h2>Artistas</h2>
    <p>Conoce a los interpretes que forman parte del catálogo.</p>
  </div>
  <div class="page-head__actions">
    <form class="searchbox" action="<%=AppUrl("artists.asp")%>" method="get" role="search">
      <span class="searchbox__ic"><%=Icon("search")%></span>
      <input type="search" name="q" value="<%=H(ar_q)%>" placeholder="Buscar artista">
    </form>
    <% If g_IsAdmin Then %>
    <a class="btn btn--primary" href="<%=AppUrl("artists.asp?action=new")%>"><%=Icon("plus")%>Nuevo artista</a>
    <% End If %>
  </div>
</div>

<% If Len(ar_err) > 0 Then %>
<div class="alert alert--error"><span class="alert__dot"></span><span><%=H(ar_err)%></span></div>
<% End If %>

<% If ar_new And g_IsAdmin Then %>
<section class="panel panel--narrow mb-24">
  <div class="panel__head">
    <h3><% If ar_id > 0 Then %>Editar artista<% Else %>Nuevo artista<% End If %></h3>
    <p>Registra el nombre artístico y, si lo deseas, una fotografía.</p>
  </div>
  <form method="post" action="<%=AppUrl("artists.asp")%>" enctype="multipart/form-data">
    <input type="hidden" name="action" value="save">
    <input type="hidden" name="id" value="<%=ar_id%>">
    <div class="form-grid">
      <div class="field">
        <label for="artist_name">Nombre artístico *</label>
        <input class="input" type="text" id="artist_name" name="artist_name" value="<%=H(ar_name)%>" maxlength="100" required>
      </div>
      <div class="field">
        <label for="artist_lastname">Apellido</label>
        <input class="input" type="text" id="artist_lastname" name="artist_lastname" value="<%=H(ar_last)%>" maxlength="100">
      </div>
      <div class="field field--full">
        <label for="artist_image">Ruta de la fotografía</label>
        <input class="input" type="text" id="artist_image" name="artist_image" value="<%=H(ar_img)%>" maxlength="255" placeholder="~/ArtistImages/artist1.jpg">
      </div>
      <div class="field field--full">
        <label for="artist_image_file">Subir fotografía <span class="hint">(opcional, max. 6 MB)</span></label>
        <input class="input" type="file" id="artist_image_file" name="artist_image_file" accept="image/*">
      </div>
    </div>
    <div class="form-actions">
      <button class="btn btn--primary" type="submit"><%=Icon("check")%>Guardar</button>
      <a class="btn btn--ghost" href="<%=AppUrl("artists.asp")%>">Cancelar</a>
    </div>
  </form>
</section>
<% End If %>

<%
ar_q = Trim(ar_q)
If Len(ar_q) > 0 Then
    ar_where = "WHERE (ARTIST_NAME LIKE @p0 OR ARTIST_LASTNAME LIKE @p0)"
    ar_prms = Array("%" & SqlLike(ar_q) & "%")
End If

Set ar_rs = DB_RS("SELECT ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE, " & _
                  "(SELECT COUNT(*) FROM TBL_PLAYER p WHERE p.ID_ARTIST = TBL_ARTIST.ID_ARTIST) AS N_SONGS " & _
                  "FROM TBL_ARTIST " & ar_where & " ORDER BY ID_ARTIST", ar_prms)
If IsObject(ar_rs) Then
    If Not (ar_rs Is Nothing) Then
        If ar_rs.EOF Then
            Response.Write EmptyState("artist", "Sin artistas", "No hay artistas que coincidan con la búsqueda.", "")
        Else
            Response.Write "<div class=""grid"">"
            Do While Not ar_rs.EOF
                ar_count = ToInt(ar_rs("N_SONGS"))
                ar_acts = ""
                If g_IsAdmin Then
                    ar_acts = "<div class=""card__actions"">" & _
                              EditLink(AppUrl("artists.asp?action=new&id=" & ToInt(ar_rs("ID_ARTIST"))), "Editar artista") & _
                              DeleteForm(AppUrl("artists.asp"), "delete", "id", ToInt(ar_rs("ID_ARTIST")), "¿Eliminar el artista seleccionado?") & _
                              "</div>"
                End If
                Response.Write "<article class=""card"">" & _
                    "<a class=""card__link"" href=""" & AppUrl("artist.asp?id=" & ToInt(ar_rs("ID_ARTIST"))) & """ aria-label=""" & H(ar_rs("ARTIST_NAME")) & """></a>" & _
                    "<div class=""card__art card__art--round""><img src=""" & H(AssetUrl(ar_rs("ARTIST_IMAGE"), "artist")) & """ alt="""" loading=""lazy""></div>" & _
                    "<div class=""card__title"">" & H(Trim(N(ar_rs("ARTIST_NAME")) & " " & N(ar_rs("ARTIST_LASTNAME")))) & "</div>" & _
                    "<div class=""card__sub"">" & ar_count & " canciones</div>" & _
                    ar_acts & "</article>"
                ar_rs.MoveNext
            Loop
            Response.Write "</div>"
        End If
        ar_rs.Close
    End If
End If
Set ar_rs = Nothing
%>

<!--#include file="inc/layout_foot.asp"-->
