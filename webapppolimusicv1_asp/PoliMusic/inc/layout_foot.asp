<%
' =============================================================================
'  PoliMusic - Pie comun: reproductor persistente y scripts
' =============================================================================
EnsureAppRoot
%>
    </main>
  </div>
</div>

<div class="player" id="player" data-approot="<%=g_AppRoot%>">
  <div class="player__track">
    <div class="player__coverwrap">
      <img class="player__cover" id="plCover" src="<%=AppUrl("assets/img/default-album.svg")%>" alt="">
    </div>
    <div class="player__meta">
      <span class="player__name" id="plTitle">Sin reproducción</span>
      <span class="player__artist" id="plArtist">Elige una canción de la biblioteca</span>
    </div>
  </div>

  <div class="player__mid">
    <div class="player__buttons">
      <button type="button" class="icon-btn" id="plShuffle" title="Aleatorio"><%=Icon("shuffle")%></button>
      <button type="button" class="icon-btn" id="plPrev" title="Anterior"><%=Icon("prev")%></button>
      <button type="button" class="playbtn" id="plPlay" title="Reproducir"><%=Icon("play")%></button>
      <button type="button" class="icon-btn" id="plNext" title="Siguiente"><%=Icon("next")%></button>
      <button type="button" class="icon-btn" id="plRepeat" title="Repetir"><%=Icon("repeat")%></button>
    </div>
    <div class="player__bar">
      <span class="t" id="plCur">0:00</span>
      <input class="range" id="plSeek" type="range" min="0" max="1000" value="0" step="1" aria-label="Progreso de la canción">
      <span class="t" id="plDur">0:00</span>
    </div>
  </div>

  <div class="player__right">
    <button type="button" class="icon-btn" id="plMute" title="Silenciar"><%=Icon("volume")%></button>
    <input class="range range--vol" id="plVol" type="range" min="0" max="100" value="100" step="1" aria-label="Volumen">
    <span class="queue-pill" id="plCount" title="Canciones en la cola">0</span>
  </div>

  <audio id="audio" preload="metadata"></audio>
</div>

<div class="toasts" id="toasts" aria-live="polite"></div>

<%
' ---------------------------------------------------------------------------
'  Dialogo "añadir canción a una playlist"
' ---------------------------------------------------------------------------
Dim lf_rs, lf_n
lf_n = 0
%>
<div class="modal" id="pmModal" hidden>
  <div class="modal__backdrop" data-close="1"></div>
  <div class="modal__box" role="dialog" aria-modal="true" aria-labelledby="pmModalTitle">
    <div class="modal__head">
      <h4 id="pmModalTitle">Añadir a una playlist</h4>
      <button type="button" class="icon-btn" data-close="1" title="Cerrar"><%=Icon("close")%></button>
    </div>
    <p class="modal__sub" id="pmModalSub">Elige la playlist de destino.</p>
    <form method="post" action="<%=AppUrl("playlist.asp")%>">
      <input type="hidden" name="action" value="addsong">
      <input type="hidden" name="song" id="pmSongId" value="">
      <div class="field">
        <label for="pmPlaylist">Playlist</label>
        <select class="select" name="playlist" id="pmPlaylist" required>
          <option value="">-- Selecciona una playlist --</option>
          <%
          Set lf_rs = DB_RS("SELECT ID_PLAYLIST, TITLE, TYPE FROM TBL_PLAYLIST WHERE ID_USER = @p0 ORDER BY TITLE", Array(g_UserID))
          If IsObject(lf_rs) Then
              If Not (lf_rs Is Nothing) Then
                  Do While Not lf_rs.EOF
                      lf_n = lf_n + 1
          %>
          <option value="<%=ToInt(lf_rs("ID_PLAYLIST"))%>"><%=H(lf_rs("TITLE"))%> (<%=H(PlaylistTypeName(lf_rs("TYPE")))%>)</option>
          <%
                      lf_rs.MoveNext
                  Loop
                  lf_rs.Close
              End If
          End If
          Set lf_rs = Nothing
          %>
        </select>
        <% If lf_n = 0 Then %>
        <span class="hint">Todavía no tienes playlists. <a href="<%=AppUrl("playlist.asp?action=new")%>">Crea la primera</a>.</span>
        <% End If %>
      </div>
      <div class="form-actions">
        <button class="btn btn--primary" type="submit"><%=Icon("plus")%>Añadir</button>
        <button class="btn btn--ghost" type="button" data-close="1">Cancelar</button>
      </div>
    </form>
  </div>
</div>

<script src="<%=AppUrl("assets/js/polimusic.js")%>"></script>
</body>
</html>
