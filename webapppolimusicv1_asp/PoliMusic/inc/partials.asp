<%
' =============================================================================
'  PoliMusic - Consultas y fragmentos de interfaz reutilizables
' =============================================================================

Function SqlTop(n)
    If ToInt(n) > 0 Then
        SqlTop = "TOP " & ToInt(n) & " "
    Else
        SqlTop = ""
    End If
End Function

' -----------------------------------------------------------------------------
'  Consulta base de canciones: incluye el artista principal y el album principal
'  de cada cancion (una cancion puede pertenecer a varios).
' -----------------------------------------------------------------------------
Function SongSql(topN, whereSql, orderSql)
    SongSql = "SELECT " & SqlTop(topN) & _
        "s.ID_SONG, s.SONG_NAME, s.SONG_PATH, s.PLAYS, " & _
        "ar.ARTIST_ID, ar.ARTIST_NAME, ar.ARTIST_IMAGE, " & _
        "al.ID_ALBUM, al.TITLE_ALBUM, al.ALBUM_IMAGE_PATH " & _
        "FROM TBL_SONG s " & _
        "OUTER APPLY (SELECT TOP 1 a.ID_ARTIST AS ARTIST_ID, " & _
        "    a.ARTIST_NAME + ISNULL(' ' + NULLIF(a.ARTIST_LASTNAME, ''), '') AS ARTIST_NAME, " & _
        "    a.ARTIST_IMAGE " & _
        "  FROM TBL_PLAYER p INNER JOIN TBL_ARTIST a ON a.ID_ARTIST = p.ID_ARTIST " & _
        "  WHERE p.ID_SONG = s.ID_SONG ORDER BY a.ID_ARTIST) ar " & _
        "OUTER APPLY (SELECT TOP 1 al2.ID_ALBUM, al2.TITLE_ALBUM, al2.ALBUM_IMAGE_PATH " & _
        "  FROM TBL_TRACKLIST t INNER JOIN TBL_ALBUM al2 ON al2.ID_ALBUM = t.ID_ALBUM " & _
        "  WHERE t.ID_SONG = s.ID_SONG ORDER BY al2.ID_ALBUM) al " & _
        whereSql & " " & orderSql
End Function

' -----------------------------------------------------------------------------
'  Cabecera de una lista de canciones
' -----------------------------------------------------------------------------
Function TracklistHead()
    TracklistHead = "<div class=""tracklist__head"">" & _
        "<span>#</span><span>Título</span><span class=""th-artist"">Artista</span>" & _
        "<span class=""th-album"">Álbum</span><span class=""th-plays"">Reprod.</span><span></span></div>"
End Function

' -----------------------------------------------------------------------------
'  Fila de cancion. Si showArtist/showAlbum son False se deja la celda vacia
'  para que la rejilla siga alineada.
' -----------------------------------------------------------------------------
Function TrackRowHtml(rs, num, showArtist, showAlbum, actionsHtml)
    Dim sid, title, artist, artistId, albumId, albumTitle, src, cover, plays, html

    sid         = ToInt(rs("ID_SONG"))
    title       = N(rs("SONG_NAME"))
    artist      = N(rs("ARTIST_NAME"))
    artistId    = ToInt(rs("ARTIST_ID"))
    albumId     = ToInt(rs("ID_ALBUM"))
    albumTitle  = N(rs("TITLE_ALBUM"))
    plays       = ToInt(rs("PLAYS"))
    src         = AssetUrl(rs("SONG_PATH"), "song")
    cover       = AssetUrl(rs("ALBUM_IMAGE_PATH"), "album")
    If Len(artist) = 0 Then artist = "Artista desconocido"

    html = "<div class=""track"" data-id=""" & sid & """ data-src=""" & H(src) & _
           """ data-title=""" & H(title) & """ data-artist=""" & H(artist) & _
           """ data-cover=""" & H(cover) & """>"

    html = html & "<div class=""track__idx""><span class=""track__num"">" & num & "</span>" & _
           "<button type=""button"" class=""icon-btn track__play js-play"" title=""Reproducir " & H(title) & """>" & Icon("play") & "</button></div>"

    html = html & "<div class=""track__main"">" & _
           "<img class=""track__img"" src=""" & H(cover) & """ alt="""">" & _
           "<span class=""track__txt""><span class=""track__name"">" & H(title) & "</span>" & _
           "<span class=""track__sub"">" & H(artist) & "</span></span></div>"

    If showArtist Then
        If artistId > 0 Then
            html = html & "<a class=""track__artist"" href=""" & AppUrl("artist.asp?id=" & artistId) & """ title=""" & H(artist) & """>" & H(artist) & "</a>"
        Else
            html = html & "<span class=""track__artist"">" & H(artist) & "</span>"
        End If
    Else
        html = html & "<span class=""track__artist""></span>"
    End If

    If showAlbum Then
        If albumId > 0 Then
            html = html & "<a class=""track__album"" href=""" & AppUrl("album.asp?id=" & albumId) & """ title=""" & H(albumTitle) & """>" & H(albumTitle) & "</a>"
        Else
            html = html & "<span class=""track__album""></span>"
        End If
    Else
        html = html & "<span class=""track__album""></span>"
    End If

    html = html & "<span class=""track__plays"">" & plays & "</span>"
    html = html & "<span class=""track__acts js-noplay"">" & N(actionsHtml) & "</span>"
    html = html & "</div>"

    TrackRowHtml = html
End Function

' -----------------------------------------------------------------------------
'  Boton "añadir a playlist" (solo usuarios con sesion)
' -----------------------------------------------------------------------------
Function AddToPlaylistBtn(songId, songTitle)
    If Not g_IsLogged Then
        AddToPlaylistBtn = ""
        Exit Function
    End If
    AddToPlaylistBtn = "<button type=""button"" class=""icon-btn"" data-addsong=""" & ToInt(songId) & _
                       """ data-songtitle=""" & H(songTitle) & """ title=""Añadir a una playlist"">" & Icon("plus") & "</button>"
End Function

' -----------------------------------------------------------------------------
'  Tarjetas
' -----------------------------------------------------------------------------
Function AlbumCardHtml(id, title, year, imgPath)
    Dim url
    url = AppUrl("album.asp?id=" & ToInt(id))
    AlbumCardHtml = "<article class=""card"">" & _
        "<a class=""card__link"" href=""" & url & """ aria-label=""" & H(title) & """></a>" & _
        "<div class=""card__art""><img src=""" & H(AssetUrl(imgPath, "album")) & """ alt="""" loading=""lazy""></div>" & _
        "<div class=""card__title"">" & H(title) & "</div>" & _
        "<div class=""card__sub"">" & H(N(year)) & "</div>" & _
        "</article>"
End Function

Function ArtistCardHtml(id, name, imgPath)
    Dim url
    url = AppUrl("artist.asp?id=" & ToInt(id))
    ArtistCardHtml = "<article class=""card"">" & _
        "<a class=""card__link"" href=""" & url & """ aria-label=""" & H(name) & """></a>" & _
        "<div class=""card__art card__art--round""><img src=""" & H(AssetUrl(imgPath, "artist")) & """ alt="""" loading=""lazy""></div>" & _
        "<div class=""card__title"">" & H(name) & "</div>" & _
        "<div class=""card__sub"">Artista</div>" & _
        "</article>"
End Function

Function PlaylistCardHtml(id, title, plType, imgPath, ownerName)
    Dim url
    url = AppUrl("playlist.asp?id=" & ToInt(id))
    PlaylistCardHtml = "<article class=""card"">" & _
        "<a class=""card__link"" href=""" & url & """ aria-label=""" & H(title) & """></a>" & _
        "<div class=""card__art""><img src=""" & H(AssetUrl(imgPath, "playlist")) & """ alt="""" loading=""lazy""></div>" & _
        "<div class=""card__title"">" & H(title) & "</div>" & _
        "<div class=""card__sub"">" & H(PlaylistTypeName(plType)) & " &middot; " & H(ownerName) & "</div>" & _
        "</article>"
End Function

' -----------------------------------------------------------------------------
'  Acciones de administracion en linea
' -----------------------------------------------------------------------------
Function EditLink(url, title)
    EditLink = "<a class=""icon-btn"" href=""" & N(url) & """ title=""" & H(title) & """>" & Icon("edit") & "</a>"
End Function

Function DeleteForm(url, actionValue, fieldName, fieldValue, confirmMsg)
    Dim msg
    msg = N(confirmMsg)
    If Len(msg) = 0 Then msg = "¿Seguro que deseas eliminar este registro?"
    msg = Replace(msg, "'", " ")
    DeleteForm = "<form class=""inline-form"" method=""post"" action=""" & N(url) & """ " & _
                 "onsubmit=""return confirm('" & msg & "');"">" & _
                 "<input type=""hidden"" name=""action"" value=""" & H(actionValue) & """>" & _
                 "<input type=""hidden"" name=""" & N(fieldName) & """ value=""" & H(fieldValue) & """>" & _
                 "<button type=""submit"" class=""icon-btn"" title=""Eliminar"">" & Icon("trash") & "</button></form>"
End Function

' -----------------------------------------------------------------------------
'  Estado vacio
' -----------------------------------------------------------------------------
Function EmptyState(iconName, title, text, actionHtml)
    EmptyState = "<div class=""empty"">" & Icon(iconName) & _
        "<h4>" & H(title) & "</h4><p>" & H(text) & "</p>" & N(actionHtml) & "</div>"
End Function
%>
