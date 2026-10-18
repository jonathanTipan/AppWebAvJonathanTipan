<%
' =============================================================================
'  PoliMusic - Cabecera comun de la aplicacion (barra lateral + barra superior)
'  Se incluye solo en paginas que requieren sesion iniciada.
' =============================================================================
Dim lh_rs, lh_pl, lh_title, lh_userPhoto
lh_title = N(g_PageTitle)
If Len(lh_title) = 0 Then lh_title = "Inicio"
lh_userPhoto = CurrentPhotoUrl()
%><!DOCTYPE html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="theme-color" content="#0a0f1c">
<title><%=H(lh_title)%> &middot; <%=APP_NAME%></title>
<link rel="icon" type="image/svg+xml" href="<%=AppUrl("assets/img/favicon.svg")%>">
<link rel="stylesheet" href="<%=AppUrl("assets/css/polimusic.css")%>">
</head>
<body class="app">
<a class="skip" href="#content">Saltar al contenido</a>
<div class="layout">

  <aside class="sidebar">
    <a class="brand" href="<%=AppUrl("home.asp")%>">
      <img class="brand__logo" src="<%=AppUrl("assets/img/logo.svg")%>" alt="">
      <span class="brand__name"><%=APP_NAME%></span>
    </a>

    <nav class="nav" aria-label="Navegación principal">
      <a class="nav__item <%=NavClass("home")%>" href="<%=AppUrl("home.asp")%>"><%=Icon("home")%><span>Inicio</span></a>
      <a class="nav__item <%=NavClass("search")%>" href="<%=AppUrl("search.asp")%>"><%=Icon("search")%><span>Buscar</span></a>
      <a class="nav__item <%=NavClass("songs")%>" href="<%=AppUrl("songs.asp")%>"><%=Icon("music")%><span>Canciones</span></a>
      <a class="nav__item <%=NavClass("albums")%>" href="<%=AppUrl("albums.asp")%>"><%=Icon("disc")%><span>Álbumes</span></a>
      <a class="nav__item <%=NavClass("artists")%>" href="<%=AppUrl("artists.asp")%>"><%=Icon("artist")%><span>Artistas</span></a>
      <a class="nav__item <%=NavClass("playlists")%>" href="<%=AppUrl("playlists.asp")%>"><%=Icon("playlist")%><span>Mis playlists</span></a>
      <% If g_IsAdmin Then %>
      <a class="nav__item <%=NavClass("users")%>" href="<%=AppUrl("users.asp")%>"><%=Icon("users")%><span>Usuarios</span><em class="pill">admin</em></a>
      <% End If %>
    </nav>

    <div class="side-block">
      <div class="side-block__head">
        <span>Tu biblioteca</span>
        <a class="mini-btn" href="<%=AppUrl("playlist.asp?action=new")%>" title="Crear playlist"><%=Icon("plus")%></a>
      </div>
      <div class="side-block__list">
        <%
        Set lh_rs = DB_RS("SELECT TOP 30 ID_PLAYLIST, TITLE, PHOTO, TYPE FROM TBL_PLAYLIST WHERE ID_USER = @p0 ORDER BY CREATION_DATE DESC, ID_PLAYLIST DESC", Array(g_UserID))
        If IsObject(lh_rs) Then
            If Not (lh_rs Is Nothing) Then
                Do While Not lh_rs.EOF
                    lh_pl = ToInt(lh_rs("ID_PLAYLIST"))
        %>
        <a class="side-link" href="<%=AppUrl("playlist.asp?id=" & lh_pl)%>">
          <img class="side-link__img" src="<%=AssetUrl(lh_rs("PHOTO"), "playlist")%>" alt="" loading="lazy">
          <span class="side-link__txt">
            <b><%=H(lh_rs("TITLE"))%></b>
            <small><%=H(PlaylistTypeName(lh_rs("TYPE")))%></small>
          </span>
        </a>
        <%
                    lh_rs.MoveNext
                Loop
                lh_rs.Close
            End If
        End If
        Set lh_rs = Nothing
        %>
        <a class="side-link side-link--more" href="<%=AppUrl("playlists.asp")%>"><%=Icon("chevron-right")%><span>Ver todas mis playlists</span></a>
      </div>
    </div>

    <div class="side-foot">
      <a class="side-foot__link <%=NavClass("profile")%>" href="<%=AppUrl("profile.asp")%>"><%=Icon("settings")%><span>Mi perfil</span></a>
      <a class="side-foot__link" href="<%=AppUrl("logout.asp")%>"><%=Icon("logout")%><span>Cerrar sesión</span></a>
    </div>
  </aside>

  <div class="main">
    <header class="topbar">
      <div class="topbar__left">
        <h1 class="topbar__title"><%=H(lh_title)%></h1>
      </div>
      <div class="topbar__right">
        <form class="searchbox" action="<%=AppUrl("search.asp")%>" method="get" role="search">
          <span class="searchbox__ic"><%=Icon("search")%></span>
          <input type="search" name="q" value="<%=H(QStr("q"))%>" placeholder="Buscar canciones, artistas o álbumes" aria-label="Buscar">
        </form>
        <a class="userchip" href="<%=AppUrl("profile.asp")%>">
          <img class="userchip__img" src="<%=lh_userPhoto%>" alt="">
          <span class="userchip__txt">
            <b><%=H(g_Username)%></b>
            <small><%=H(RoleName(g_UserType))%></small>
          </span>
        </a>
        <a class="icon-btn" href="<%=AppUrl("logout.asp")%>" title="Cerrar sesión"><%=Icon("logout")%></a>
      </div>
    </header>

    <main class="content" id="content">
      <%=RenderFlash()%>
