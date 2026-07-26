/* Music Player - Alpine.js audio playback + folder browser control */

function audioPlayer() {
  return {
    playing: false,
    currentTime: 0,
    duration: 0,
    title: 'No tracks found',
    sub: '',
    hasTracks: false,
    trackCount: 0,
    idx: 0,
    vol: 70,
    _t: null,
    _a: null,
    get progress() { return this.duration > 0 ? this.currentTime / this.duration * 100 : 0; },
    init() {
      this._a = document.getElementById('app-audio');
      this.idx = parseInt(this.$el.dataset.si) || 0;
      this.vol = parseInt(this.$el.dataset.sv) || 70;
      this._a.volume = this.vol / 100;
      this.refreshFromDOM();
      var act = document.querySelector('.playlist-item.active');
      if (act) this._a.src = act.dataset.url;
      this._a.addEventListener('timeupdate', function() {
        this.currentTime = this._a.currentTime;
        this.duration = this._a.duration || 0;
      }.bind(this));
      this._a.addEventListener('ended', function() { this.next(); }.bind(this));
    },
    refreshFromDOM() {
      var items = document.querySelectorAll('.playlist-item');
      this.trackCount = items.length;
      this.hasTracks = items.length > 0;
      var act = document.querySelector('.playlist-item.active');
      if (act) {
        this.title = act.querySelector('.item-name').textContent;
        this.sub = 'Track ' + (this.idx + 1) + ' of ' + items.length;
        if (!this._a.src) {
          this._a.src = act.dataset.url;
        }
      } else {
        this.title = 'No tracks found';
        this.sub = '';
      }
    },
    play(el) {
      document.querySelectorAll('.playlist-item.active,.playlist-item.playing')
        .forEach(function(e) { e.classList.remove('active', 'playing'); });
      el.classList.add('active', 'playing');
      this.idx = parseInt(el.dataset.idx);
      this._a.src = el.dataset.url;
      this.title = el.querySelector('.item-name').textContent;
      this.sub = 'Track ' + (this.idx + 1) + ' of ' + document.querySelectorAll('.playlist-item').length;
      this._a.play().catch(function() {});
      this.playing = true;
      this._save();
    },
    toggle() {
      if (!this._a.src) return;
      if (this.playing) {
        this._a.pause();
        document.querySelectorAll('.playlist-item.playing').forEach(function(e) { e.classList.remove('playing'); });
      } else {
        this._a.play().catch(function() {});
        var a = document.querySelector('.playlist-item.active');
        if (a) a.classList.add('playing');
      }
      this.playing = !this.playing;
    },
    next() {
      var items = document.querySelectorAll('.playlist-item');
      if (items.length < 2) return;
      this.play(items[(this.idx + 1) % items.length]);
    },
    prev() {
      var items = document.querySelectorAll('.playlist-item');
      if (items.length < 2) return;
      this.play(items[(this.idx - 1 + items.length) % items.length]);
    },
    seek(e) {
      var r = e.currentTarget.getBoundingClientRect();
      this._a.currentTime = (e.clientX - r.left) / r.width * this.duration;
    },
    setVolume(v) {
      this.vol = parseInt(v);
      this._a.volume = this.vol / 100;
      this._save();
    },
    _save() {
      clearTimeout(this._t);
      var self = this;
      this._t = setTimeout(function() {
        fetch('', {method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},
          body:'action=save_state&idx='+self.idx+'&vol='+self.vol}).catch(function(){});
      }, 500);
    },
    fmtTime(s) {
      if (!s || isNaN(s)) return '0:00';
      return Math.floor(s/60) + ':' + String(Math.floor(s%60)).padStart(2,'0');
    }
  };
}

/* Folder browser modal control */
function openFolderBrowser() {
  document.getElementById('folder-browser-overlay').style.display = 'flex';
  htmx.ajax('GET', '/api/browse.asp', '#folder-browser-body');
}

function closeFolderBrowser() {
  document.getElementById('folder-browser-overlay').style.display = 'none';
}
