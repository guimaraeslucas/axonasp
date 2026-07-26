# Music Player - AxonHTA Demo

A music player built with **HTMX + Alpine.js + ASP (VBScript)**, running on AxonHTA. Fully offline — no CDN dependencies.

## Architecture

| Layer | Technology | Responsibility |
|-------|-----------|----------------|
| Playlist rendering | ASP (VBScript) + FSO | Server-side directory scan, HTML fragment output |
| Folder browser | ASP + HTMX | Visual directory tree navigation, no manual path input |
| Audio playback | Alpine.js (~100 lines) | play/pause/seek/volume/next/prev |
| State persistence | ASP + FSO | `data/state.dat` (idx+vol), `data/path_aliases.dat` (music dir) |
| File serving | AxonHTA path alias | `/music/` → any local directory |

## File Structure

```
axonhta-music-player/
├── index.hta          Minimal entry: HTA tag + save_state handler + HTML shell
├── app.js             Alpine.js player component + folder browser control
├── style.css          Dark theme stylesheet
├── vendor/            Local JS libraries (offline, no CDN)
│   ├── htmx.min.js       HTMX 2.0.4
│   └── alpine.min.js     Alpine.js 3.x
├── lib/               Reusable ASP functions
│   ├── helpers.asp       JSEsc, FmtSize, IsAudioFile, HtmlEsc
│   ├── music.asp         GetMusicDir, ScanMusicFolder
│   └── state.asp         LoadState, SaveState, SaveMusicDir
├── api/               HTMX endpoints (return HTML fragments)
│   ├── playlist.asp      GET: render playlist / POST: save dir + rescan
│   └── browse.asp        GET: render folder browser modal
├── music/             Default music folder (put .mp3/.wav/.flac here)
├── data/              Runtime data (auto-created)
│   ├── state.dat         Saved playback state (idx|vol)
│   └── path_aliases.dat  Virtual path config (/music/|D:\Music)
└── README.md
```

## Usage

1. Put audio files in `music/` folder, OR click "Folder" button to browse and select a directory
2. Run:
   ```bash
   axonhta.exe --app ./axonhta-music-player
   ```

## Custom Music Directory

Click the "Folder" button in the playlist header to open a visual directory browser. Navigate the folder tree, then click "Select This Folder" to save. The path is stored in `data/path_aliases.dat` and the `/music/` URL prefix is mapped to that directory by AxonHTA's path alias system.

## Supported Formats

.mp3, .wav, .ogg, .flac, .m4a, .wma, .aac
