/* ===========================================================================
   PoliMusic - Reproductor y utilidades de interfaz
   La cola de reproduccion se guarda en localStorage para que la musica
   continue al navegar entre las paginas del sitio.
   =========================================================================== */
(function () {
  'use strict';

  var STORE = 'polimusic.player.v1';

  var audio = document.getElementById('audio');
  if (!audio) { return; }

  var ui = {
    player:   document.getElementById('player'),
    cover:    document.getElementById('plCover'),
    title:    document.getElementById('plTitle'),
    artist:   document.getElementById('plArtist'),
    play:     document.getElementById('plPlay'),
    prev:     document.getElementById('plPrev'),
    next:     document.getElementById('plNext'),
    shuffle:  document.getElementById('plShuffle'),
    repeat:   document.getElementById('plRepeat'),
    seek:     document.getElementById('plSeek'),
    cur:      document.getElementById('plCur'),
    dur:      document.getElementById('plDur'),
    vol:      document.getElementById('plVol'),
    mute:     document.getElementById('plMute'),
    count:    document.getElementById('plCount'),
    toasts:   document.getElementById('toasts')
  };

  var state = {
    queue: [],
    index: -1,
    time: 0,
    playing: false,
    shuffle: false,
    repeat: 'off',   /* off | all | one */
    volume: 1,
    muted: false
  };

  var pendingResume = false;
  var ICONS = {
    play:  '<svg class="ic" viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M8 5.2v13.6L19.2 12z"/></svg>',
    pause: '<svg class="ic" viewBox="0 0 24 24" aria-hidden="true"><path fill="currentColor" d="M7.5 5h3.2v14H7.5zM13.3 5h3.2v14h-3.2z"/></svg>',
    volume:'<svg class="ic" viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><path d="M4 9.5h3l4-3.5v12l-4-3.5H4z"/><path d="M14.5 9.4a3.6 3.6 0 0 1 0 5.2M17.2 6.7a7.4 7.4 0 0 1 0 10.6"/></svg>',
    mute:  '<svg class="ic" viewBox="0 0 24 24" aria-hidden="true" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round"><path d="M4 9.5h3l4-3.5v12l-4-3.5H4z"/><path d="m15 9.5 5 5M20 9.5l-5 5"/></svg>'
  };

  /* ---------------------------------------------------------------- estado */
  function loadState() {
    try {
      var raw = localStorage.getItem(STORE);
      if (!raw) { return; }
      var saved = JSON.parse(raw);
      if (saved && typeof saved === 'object') {
        state.queue   = Array.isArray(saved.queue) ? saved.queue : [];
        state.index   = typeof saved.index === 'number' ? saved.index : -1;
        state.time    = typeof saved.time === 'number' ? saved.time : 0;
        state.playing = !!saved.playing;
        state.shuffle = !!saved.shuffle;
        state.repeat  = saved.repeat || 'off';
        state.volume  = typeof saved.volume === 'number' ? saved.volume : 1;
        state.muted   = !!saved.muted;
      }
    } catch (e) { /* almacenamiento no disponible */ }
  }

  function saveState() {
    try {
      localStorage.setItem(STORE, JSON.stringify({
        queue: state.queue, index: state.index, time: state.time,
        playing: state.playing, shuffle: state.shuffle, repeat: state.repeat,
        volume: state.volume, muted: state.muted
      }));
    } catch (e) { /* ignorar */ }
  }

  /* ---------------------------------------------------------------- ayuda */
  function fmt(sec) {
    if (!isFinite(sec) || sec < 0) { sec = 0; }
    var m = Math.floor(sec / 60), s = Math.floor(sec % 60);
    return m + ':' + (s < 10 ? '0' : '') + s;
  }

  function paintRange(input) {
    if (!input) { return; }
    var min = parseFloat(input.min || 0), max = parseFloat(input.max || 100);
    var val = parseFloat(input.value || 0);
    var pct = max > min ? ((val - min) / (max - min)) * 100 : 0;
    input.style.background =
      'linear-gradient(90deg, #c77dff ' + pct + '%, rgba(255,255,255,.14) ' + pct + '%)';
  }

  function toast(msg, type) {
    if (!ui.toasts) { return; }
    var d = document.createElement('div');
    d.className = 'toast' + (type === 'error' ? ' toast--error' : '');
    d.textContent = msg;
    ui.toasts.appendChild(d);
    setTimeout(function () {
      d.style.transition = 'opacity .3s';
      d.style.opacity = '0';
      setTimeout(function () { if (d.parentNode) { d.parentNode.removeChild(d); } }, 320);
    }, 3600);
  }

  /* ------------------------------------------------------------ reproduccion */
  function current() {
    if (state.index < 0 || state.index >= state.queue.length) { return null; }
    return state.queue[state.index];
  }

  function render() {
    var t = current();
    if (ui.cover)  { ui.cover.src = t && t.cover ? t.cover : ''; }
    if (ui.title)  { ui.title.textContent = t ? t.title : 'Sin reproduccion'; }
    if (ui.artist) { ui.artist.textContent = t ? (t.artist || 'Artista desconocido') : 'Elige una cancion de la biblioteca'; }
    if (ui.play)   { ui.play.innerHTML = state.playing ? ICONS.pause : ICONS.play; }
    if (ui.player) { ui.player.classList.toggle('is-playing', state.playing); }
    if (ui.count)  { ui.count.textContent = String(state.queue.length); }

    if (ui.shuffle) { ui.shuffle.classList.toggle('is-on', state.shuffle); }
    if (ui.repeat)  { ui.repeat.classList.toggle('is-on', state.repeat !== 'off'); }
    if (ui.mute)    { ui.mute.innerHTML = (state.muted || state.volume === 0) ? ICONS.mute : ICONS.volume; }
    if (ui.vol)     { ui.vol.value = Math.round((state.muted ? 0 : state.volume) * 100); paintRange(ui.vol); }

    document.title = t && state.playing
      ? t.title + ' - ' + (t.artist || '') + ' | PoliMusic'
      : document.title.replace(/^.* \| PoliMusic$/, function (m) { return m; });

    /* Resalta la fila que se esta reproduciendo */
    var rows = document.querySelectorAll('[data-src]');
    for (var i = 0; i < rows.length; i++) {
      var on = t && rows[i].getAttribute('data-src') === t.src &&
               rows[i].getAttribute('data-title') === t.title;
      rows[i].classList.toggle('is-playing', !!on);
    }

    /* Estado de las barras de progreso */
    if (ui.seek && audio.duration && isFinite(audio.duration)) {
      ui.seek.value = Math.round((audio.currentTime / audio.duration) * 1000);
    }
    paintRange(ui.seek);
    if (ui.cur) { ui.cur.textContent = fmt(audio.currentTime); }
    if (ui.dur) { ui.dur.textContent = fmt(audio.duration); }
  }

  function countPlay(track) {
    if (!track || !track.id) { return; }
    var root = (ui.player && ui.player.getAttribute('data-approot')) || '';
    var url = root + 'play.asp?id=' + encodeURIComponent(track.id);
    try {
      if (navigator.sendBeacon) {
        navigator.sendBeacon(url);
      } else {
        fetch(url, { method: 'POST', keepalive: true });
      }
    } catch (e) { /* el contador es informativo */ }
  }

  function setTrack(index, autoplay) {
    if (index < 0 || index >= state.queue.length) { return; }
    state.index = index;
    var t = state.queue[index];
    audio.src = t.src;
    try { audio.currentTime = 0; } catch (e) { /* aun sin metadatos */ }
    state.time = 0;
    if (autoplay !== false) {
      countPlay(t);
      play();
    }
    render();
    saveState();
  }

  function play() {
    var t = current();
    if (!t) { toast('Selecciona una cancion para reproducir'); return; }
    var p = audio.play();
    if (p && p.catch) {
      p.then(function () {
        state.playing = true;
        pendingResume = false;
        render(); saveState();
      }).catch(function () {
        /* El navegador exige una interaccion previa del usuario */
        state.playing = false;
        pendingResume = true;
        render(); saveState();
      });
    } else {
      state.playing = true;
      render(); saveState();
    }
  }

  function pause() {
    audio.pause();
    state.playing = false;
    state.time = audio.currentTime || 0;
    render(); saveState();
  }

  function toggle() {
    if (audio.paused) { play(); } else { pause(); }
  }

  function step(delta) {
    if (!state.queue.length) { return; }
    var i = state.index + delta;
    if (state.shuffle && delta > 0) {
      i = Math.floor(Math.random() * state.queue.length);
      if (state.queue.length > 1 && i === state.index) { i = (i + 1) % state.queue.length; }
    }
    if (i >= state.queue.length) {
      if (state.repeat === 'all') { i = 0; } else { pause(); return; }
    }
    if (i < 0) { i = state.queue.length - 1; }
    setTrack(i, true);
  }

  function playQueue(items, index) {
    if (!items || !items.length) { return; }
    state.queue = items;
    setTrack(index >= 0 && index < items.length ? index : 0, true);
    toast('Reproduciendo: ' + items[state.index].title);
  }

  function handlePlayFrom(node) {
    var container = node.closest('[data-queue]');
    var scope = container || document;
    var nodes = scope.querySelectorAll('[data-src]');
    var items = [], idx = 0, found = false;

    for (var i = 0; i < nodes.length; i++) {
      items.push({
        id:     nodes[i].getAttribute('data-id') || '',
        src:    nodes[i].getAttribute('data-src'),
        title:  nodes[i].getAttribute('data-title') || 'Pista',
        artist: nodes[i].getAttribute('data-artist') || '',
        cover:  nodes[i].getAttribute('data-cover') || ''
      });
      if (nodes[i] === node) { idx = i; found = true; }
    }
    if (!found) { idx = 0; }
    playQueue(items, idx);
  }

  /* ------------------------------------------------------------- eventos UI */
  if (ui.play)    { ui.play.addEventListener('click', toggle); }
  if (ui.prev)    { ui.prev.addEventListener('click', function () {
                      if (audio.currentTime > 3) { audio.currentTime = 0; }
                      else { step(-1); }
                    }); }
  if (ui.next)    { ui.next.addEventListener('click', function () { step(1); }); }
  if (ui.shuffle) { ui.shuffle.addEventListener('click', function () {
                      state.shuffle = !state.shuffle; render(); saveState();
                      toast(state.shuffle ? 'Aleatorio activado' : 'Aleatorio desactivado');
                    }); }
  if (ui.repeat)  { ui.repeat.addEventListener('click', function () {
                      state.repeat = state.repeat === 'off' ? 'all' : (state.repeat === 'all' ? 'one' : 'off');
                      render(); saveState();
                      toast(state.repeat === 'off' ? 'Repeticion desactivada'
                          : state.repeat === 'all' ? 'Repetir toda la cola' : 'Repetir esta cancion');
                    }); }
  if (ui.mute)    { ui.mute.addEventListener('click', function () {
                      state.muted = !state.muted;
                      audio.muted = state.muted;
                      render(); saveState();
                    }); }
  if (ui.vol)     { ui.vol.addEventListener('input', function () {
                      state.volume = parseFloat(ui.vol.value) / 100;
                      state.muted = state.volume === 0;
                      audio.volume = state.volume;
                      audio.muted = false;
                      render(); saveState();
                    }); }
  if (ui.seek)    { ui.seek.addEventListener('input', function () {
                      if (audio.duration && isFinite(audio.duration)) {
                        audio.currentTime = (parseFloat(ui.seek.value) / 1000) * audio.duration;
                        paintRange(ui.seek);
                      }
                    }); }

  audio.addEventListener('timeupdate', function () {
    state.time = audio.currentTime;
    if (ui.seek && audio.duration && isFinite(audio.duration)) {
      ui.seek.value = Math.round((audio.currentTime / audio.duration) * 1000);
      paintRange(ui.seek);
    }
    if (ui.cur) { ui.cur.textContent = fmt(audio.currentTime); }
  });

  audio.addEventListener('loadedmetadata', function () {
    if (ui.dur) { ui.dur.textContent = fmt(audio.duration); }
    if (state.time > 0 && state.time < audio.duration - 1) {
      try { audio.currentTime = state.time; } catch (e) { /* ignorar */ }
    }
  });

  audio.addEventListener('play',  function () { state.playing = true;  render(); });
  audio.addEventListener('pause', function () { state.playing = false; render(); });

  audio.addEventListener('ended', function () {
    if (state.repeat === 'one') {
      audio.currentTime = 0;
      play();
      return;
    }
    var last = state.index >= state.queue.length - 1;
    if (last && state.repeat !== 'all') {
      state.playing = false;
      state.time = 0;
      render(); saveState();
      return;
    }
    step(1);
  });

  audio.addEventListener('error', function () {
    if (audio.src) { toast('No se pudo cargar el audio de esta pista', 'error'); }
  });

  /* Delegacion: botones de reproducir y filas de canciones */
  document.addEventListener('click', function (e) {
    var playBtn = e.target.closest ? e.target.closest('.js-play') : null;
    if (playBtn) {
      e.preventDefault();
      var node = playBtn.closest('[data-src]');
      if (node) { handlePlayFrom(node); }
      return;
    }
    var allBtn = e.target.closest ? e.target.closest('[data-playall]') : null;
    if (allBtn) {
      e.preventDefault();
      var container = allBtn.closest('[data-queue]') || document;
      var nodes = container.querySelectorAll('[data-src]');
      if (!nodes.length) { toast('No hay canciones para reproducir', 'error'); return; }
      handlePlayFrom(nodes[0]);
      return;
    }
    var row = e.target.closest ? e.target.closest('[data-src]') : null;
    if (row && !e.target.closest('a, button, input, select, textarea, .js-noplay')) {
      e.preventDefault();
      handlePlayFrom(row);
    }
  });

  /* Guarda el progreso antes de salir */
  window.addEventListener('pagehide', function () { state.time = audio.currentTime || 0; saveState(); });
  window.addEventListener('beforeunload', function () { state.time = audio.currentTime || 0; saveState(); });

  /* ------------------------------------------------------------ restauracion */
  function restore() {
    loadState();
    audio.volume = state.volume;
    audio.muted = state.muted;

    if (!state.queue.length || state.index < 0) { render(); return; }

    var t = state.queue[state.index];
    audio.src = t.src;
    render();

    if (state.playing) {
      var p = audio.play();
      if (p && p.catch) {
        p.catch(function () {
          pendingResume = true;
          state.playing = false;
          render(); saveState();
        });
      }
    }
  }

  /* Si el navegador bloquea la reanudacion automatica, se reanuda en el
     primer clic o tecla del usuario. */
  function armResume() {
    function once() {
      if (!pendingResume) { return; }
      pendingResume = false;
      var p = audio.play();
      if (p && p.catch) { p.catch(function () { /* sigue bloqueado */ }); }
    }
    document.addEventListener('click', once, true);
    document.addEventListener('keydown', once, true);
  }

  /* Atajos de teclado */
  document.addEventListener('keydown', function (e) {
    var tag = (e.target.tagName || '').toLowerCase();
    if (tag === 'input' || tag === 'textarea' || tag === 'select') { return; }
    if (e.code === 'Space')      { e.preventDefault(); toggle(); }
    else if (e.code === 'ArrowRight' && e.altKey) { step(1); }
    else if (e.code === 'ArrowLeft'  && e.altKey) { step(-1); }
  });

  /* ------------------------------------------------------------- dialogos */
  var modal = document.getElementById('pmModal');

  function closeModal() {
    if (modal) { modal.hidden = true; }
  }

  function openAddToPlaylist(songId, songTitle) {
    if (!modal) {
      toast('Debes iniciar sesion para usar playlists', 'error');
      return;
    }
    var idf = document.getElementById('pmSongId');
    var sub = document.getElementById('pmModalSub');
    var sel = document.getElementById('pmPlaylist');
    if (idf) { idf.value = songId; }
    if (sub) { sub.textContent = 'Anade "' + songTitle + '" a la playlist:'; }
    if (sel) { sel.value = ''; }
    modal.hidden = false;
    if (sel) { sel.focus(); }
  }

  document.addEventListener('click', function (e) {
    var opener = e.target.closest ? e.target.closest('[data-addsong]') : null;
    if (opener) {
      e.preventDefault();
      e.stopPropagation();
      openAddToPlaylist(opener.getAttribute('data-addsong'), opener.getAttribute('data-songtitle') || '');
      return;
    }
    if (e.target.closest && e.target.closest('[data-close]')) {
      closeModal();
    }
  });

  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape') { closeModal(); }
  });

  /* API publica minima */
  window.PoliMusic = {
    toast: toast,
    play: function (items, index) { playQueue(items, index || 0); }
  };

  restore();
  armResume();
  paintRange(ui.seek);
  paintRange(ui.vol);
  if (ui.seek) { ui.seek.value = 0; }
  setInterval(saveState, 4000);
})();
