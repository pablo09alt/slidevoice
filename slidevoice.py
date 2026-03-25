#!/usr/bin/env python3
"""
SlideVoice — Avancement automatique de slides PowerPoint par reconnaissance vocale
Usage: python slidevoice.py script.txt
"""

import sys
import re
import time
import threading
import json
import webbrowser
from http.server import HTTPServer, BaseHTTPRequestHandler
from pathlib import Path

# ── Dépendances optionnelles ──────────────────────────────────────────────────
try:
    import speech_recognition as sr
    SR_OK = True
except ImportError:
    SR_OK = False

try:
    import whisper as whisper_lib
    import numpy as np
    WHISPER_OK = True
except ImportError:
    WHISPER_OK = False

try:
    import pyautogui
    PYAUTOGUI_OK = True
except ImportError:
    PYAUTOGUI_OK = False

# ── État global partagé ───────────────────────────────────────────────────────
state = {
    "slides": [],          # liste de {"keyword": str, "text": str}
    "current": 0,          # index de la slide actuelle
    "listening": False,
    "last_heard": "",
    "last_action": "",
    "log": [],
    "demo_mode": not (SR_OK and PYAUTOGUI_OK),  # mode demo si dépendances manquantes
}

# ── Parsing du script balisé ──────────────────────────────────────────────────
def parse_script(path: str) -> list[dict]:
    """
    Format attendu :
        [INTRO] Bonjour à tous, aujourd'hui je vais vous parler...
        [CONTEXTE] Pour commencer, le contexte est...
        [CONCLUSION] En conclusion...

    Le mot entre crochets est le mot-clé vocal. Quand il est détecté,
    on passe à la slide suivante.
    """
    text = Path(path).read_text(encoding="utf-8")
    slides = []
    pattern = re.compile(r"\[([^\]]+)\](.*?)(?=\[[^\]]+\]|$)", re.DOTALL)
    for match in pattern.finditer(text):
        keyword = match.group(1).strip().lower()
        content = match.group(2).strip()
        slides.append({"keyword": keyword, "text": content})
    return slides


# ── Avancement de slide ───────────────────────────────────────────────────────
def find_powerpoint_hwnd():
    """Cherche la fenêtre PowerPoint et retourne son handle."""
    try:
        import win32gui
        hwnds = []
        def callback(hwnd, results):
            title = win32gui.GetWindowText(hwnd)
            if any(k in title for k in ["PowerPoint", "Diaporama", "Slide Show"]):
                results.append(hwnd)
        win32gui.EnumWindows(callback, hwnds)
        return hwnds[0] if hwnds else None
    except Exception:
        return None

def advance_slide():
    """Envoie la touche → directement à la fenêtre PowerPoint via PostMessage."""
    if not PYAUTOGUI_OK:
        log("→ [DEMO] Touche flèche simulée (pyautogui non installé)")
        return
    try:
        import win32gui, win32con, win32api
        hwnd = find_powerpoint_hwnd()
        if hwnd:
            # VK_RIGHT = 0x27 — envoie directement à la fenêtre, pas besoin de focus
            win32api.PostMessage(hwnd, win32con.WM_KEYDOWN, 0x27, 0)
            time.sleep(0.05)
            win32api.PostMessage(hwnd, win32con.WM_KEYUP, 0x27, 0)
            log("→ Touche flèche envoyée directement à PowerPoint")
        else:
            # Fallback : focus + pyautogui
            log("⚠️  PowerPoint introuvable, tentative avec focus...")
            pyautogui.press("right")
            log("→ Touche flèche envoyée (fallback)")
    except ImportError:
        # win32 pas dispo — fallback pyautogui
        pyautogui.press("right")
        log("→ Touche flèche envoyée (pyautogui)")
    except Exception as e:
        log(f"❌ Erreur advance_slide: {e}")


def log(msg: str):
    timestamp = time.strftime("%H:%M:%S")
    entry = f"[{timestamp}] {msg}"
    state["log"].append(entry)
    if len(state["log"]) > 50:
        state["log"] = state["log"][-50:]
    print(entry)


# ── Reconnaissance vocale ─────────────────────────────────────────────────────
def normalize(text: str) -> str:
    """Retire accents et met en minuscule pour comparaison souple."""
    import unicodedata
    nfkd = unicodedata.normalize("NFKD", text.lower())
    return "".join(c for c in nfkd if not unicodedata.combining(c))


# Modèle Whisper chargé une seule fois
_whisper_model = None

def get_whisper_model():
    global _whisper_model
    if _whisper_model is None:
        log("⏳ Chargement du modèle Whisper 'base' (première fois uniquement)...")
        _whisper_model = whisper_lib.load_model("tiny")
        log("✅ Modèle Whisper prêt !")
    return _whisper_model

def voice_loop():
    if WHISPER_OK:
        voice_loop_whisper()
    elif SR_OK:
        voice_loop_google()
    else:
        log("⚠️  Aucun moteur de reconnaissance installé — mode démo actif")

def voice_loop_whisper():
    import sounddevice as sd
    model = get_whisper_model()
    SAMPLE_RATE = 16000
    CHUNK_SECONDS = 2  # écoute par tranches de 2 secondes

    log("🎙️  Écoute Whisper démarrée")
    while state["listening"]:
        try:
            audio_chunk = sd.rec(
                int(CHUNK_SECONDS * SAMPLE_RATE),
                samplerate=SAMPLE_RATE,
                channels=1,
                dtype="float32"
            )
            sd.wait()
            if not state["listening"]:
                break
            audio_flat = audio_chunk.flatten()
            # Ignore les chunks trop silencieux
            if np.abs(audio_flat).mean() < 0.001:
                continue
            result = model.transcribe(
                audio_flat,
                language="fr",
                fp16=False,
                condition_on_previous_text=False
            )
            text = result["text"].strip()
            if text:
                state["last_heard"] = text
                log(f"🗣  Entendu : « {text} »")
                check_keywords(text)
        except Exception as e:
            log(f"❌ Erreur Whisper : {e}")
            time.sleep(1)

def voice_loop_google():
    recognizer = sr.Recognizer()
    recognizer.energy_threshold = 300
    recognizer.dynamic_energy_threshold = True

    log("🎙️  Écoute Google Speech démarrée")
    with sr.Microphone() as source:
        recognizer.adjust_for_ambient_noise(source, duration=1)
        while state["listening"]:
            try:
                audio = recognizer.listen(source, timeout=5, phrase_time_limit=6)
                text = recognizer.recognize_google(audio, language="fr-FR")
                state["last_heard"] = text
                log(f"🗣  Entendu : « {text} »")
                check_keywords(text)
            except sr.WaitTimeoutError:
                pass
            except sr.UnknownValueError:
                pass
            except sr.RequestError as e:
                log(f"❌ Erreur API Google : {e}")
                time.sleep(2)


def keyword_matches(keyword: str, heard: str) -> bool:
    """
    Matching souple : tous les mots du mot-clé doivent apparaître dans heard.
    Ex: keyword="domaines compétence" matche "de compétences" ou "domaine de compétences"
    On ignore les petits mots (de, le, la, les, des, un, une, et, en, à)
    """
    stopwords = {"de", "le", "la", "les", "des", "un", "une", "et", "en", "a", "au", "aux"}
    kw_words = [w for w in keyword.split() if w not in stopwords and len(w) > 2]
    if not kw_words:
        return keyword in heard
    # Chaque mot du mot-clé doit apparaître quelque part dans heard
    # On accepte aussi les racines (ex: "compétence" dans "compétences")
    for kw_word in kw_words:
        # Cherche le mot ou sa racine (premiers 5 caractères)
        root = kw_word[:5] if len(kw_word) >= 5 else kw_word
        if root not in heard:
            return False
    return True

def check_keywords(heard: str):
    heard_norm = normalize(heard)
    slides = state["slides"]
    current = state["current"]

    if current + 1 >= len(slides):
        log("✅ Dernière slide atteinte")
        return

    # On cherche dans TOUTES les slides après la position actuelle
    for idx in range(current + 1, len(slides)):
        keyword = normalize(slides[idx]["keyword"])
        if keyword_matches(keyword, heard_norm):
            steps = idx - current
            for _ in range(steps):
                advance_slide()
            state["current"] = idx
            state["last_action"] = f"Slide {idx + 1} — mot-clé « {slides[idx]['keyword']} » détecté"
            if steps > 1:
                log(f"⏩ Saut de {steps} slides → {state['last_action']}")
            else:
                log(f"✅ {state['last_action']}")
            return


# ── Serveur HTTP local (interface web) ────────────────────────────────────────
HTML_PAGE = """<!DOCTYPE html>
<html lang="fr">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>SlideVoice</title>
<style>
  @import url('https://fonts.googleapis.com/css2?family=JetBrains+Mono:wght@400;700&family=Syne:wght@400;700;800&display=swap');

  :root {
    --bg: #0d0d0f;
    --surface: #16161a;
    --border: #2a2a32;
    --accent: #00e5a0;
    --accent2: #7b61ff;
    --text: #e8e8f0;
    --muted: #6b6b80;
    --danger: #ff4d6d;
  }

  * { box-sizing: border-box; margin: 0; padding: 0; }

  body {
    font-family: 'Syne', sans-serif;
    background: var(--bg);
    color: var(--text);
    min-height: 100vh;
    display: grid;
    grid-template-rows: auto 1fr;
  }

  header {
    padding: 1.5rem 2rem;
    border-bottom: 1px solid var(--border);
    display: flex;
    align-items: center;
    gap: 1rem;
  }

  .logo {
    font-size: 1.4rem;
    font-weight: 800;
    letter-spacing: -0.03em;
  }
  .logo span { color: var(--accent); }

  .status-pill {
    margin-left: auto;
    padding: 0.3rem 0.9rem;
    border-radius: 999px;
    font-size: 0.78rem;
    font-family: 'JetBrains Mono', monospace;
    font-weight: 700;
    letter-spacing: 0.05em;
    text-transform: uppercase;
    background: var(--surface);
    border: 1px solid var(--border);
    color: var(--muted);
    transition: all 0.3s;
  }
  .status-pill.active {
    background: rgba(0,229,160,0.1);
    border-color: var(--accent);
    color: var(--accent);
    box-shadow: 0 0 12px rgba(0,229,160,0.2);
  }

  main {
    display: grid;
    grid-template-columns: 1fr 340px;
    gap: 0;
    height: 100%;
  }

  .panel {
    padding: 2rem;
    border-right: 1px solid var(--border);
    overflow-y: auto;
  }

  .section-title {
    font-size: 0.72rem;
    font-family: 'JetBrains Mono', monospace;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: var(--muted);
    margin-bottom: 1rem;
  }

  /* Slides timeline */
  .slides-list { display: flex; flex-direction: column; gap: 0.5rem; }

  .slide-item {
    display: flex;
    gap: 1rem;
    align-items: flex-start;
    padding: 1rem;
    border-radius: 8px;
    border: 1px solid var(--border);
    background: var(--surface);
    transition: all 0.25s;
    opacity: 0.5;
  }
  .slide-item.active {
    opacity: 1;
    border-color: var(--accent);
    background: rgba(0,229,160,0.05);
    box-shadow: 0 0 20px rgba(0,229,160,0.1);
  }
  .slide-item.done {
    opacity: 0.35;
    border-color: transparent;
  }

  .slide-num {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.75rem;
    color: var(--muted);
    min-width: 2rem;
    padding-top: 0.1rem;
  }
  .slide-item.active .slide-num { color: var(--accent); }

  .slide-keyword {
    font-size: 0.85rem;
    font-weight: 700;
    color: var(--accent2);
    font-family: 'JetBrains Mono', monospace;
    margin-bottom: 0.35rem;
  }
  .slide-text {
    font-size: 0.88rem;
    color: var(--muted);
    line-height: 1.5;
    display: -webkit-box;
    -webkit-line-clamp: 2;
    -webkit-box-orient: vertical;
    overflow: hidden;
  }
  .slide-item.active .slide-text { color: var(--text); }

  /* Right sidebar */
  .sidebar {
    display: flex;
    flex-direction: column;
    gap: 0;
  }

  .sidebar-section {
    padding: 1.5rem;
    border-bottom: 1px solid var(--border);
  }

  /* Waveform / listening indicator */
  .wave-container {
    display: flex;
    align-items: center;
    justify-content: center;
    gap: 4px;
    height: 48px;
    margin: 1rem 0;
  }
  .wave-bar {
    width: 4px;
    border-radius: 2px;
    background: var(--border);
    height: 8px;
    transition: height 0.1s;
  }
  .listening .wave-bar {
    background: var(--accent);
    animation: wave 1s ease-in-out infinite;
  }
  .wave-bar:nth-child(1) { animation-delay: 0s; }
  .wave-bar:nth-child(2) { animation-delay: 0.1s; }
  .wave-bar:nth-child(3) { animation-delay: 0.2s; }
  .wave-bar:nth-child(4) { animation-delay: 0.3s; }
  .wave-bar:nth-child(5) { animation-delay: 0.2s; }
  .wave-bar:nth-child(6) { animation-delay: 0.1s; }
  .wave-bar:nth-child(7) { animation-delay: 0s; }

  @keyframes wave {
    0%, 100% { height: 8px; }
    50% { height: 32px; }
  }

  .last-heard {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.82rem;
    color: var(--text);
    background: var(--bg);
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 0.6rem 0.8rem;
    min-height: 2.5rem;
    word-break: break-word;
  }

  /* Controls */
  .btn {
    width: 100%;
    padding: 0.8rem 1.2rem;
    border-radius: 8px;
    border: none;
    font-family: 'Syne', sans-serif;
    font-size: 0.9rem;
    font-weight: 700;
    cursor: pointer;
    transition: all 0.2s;
    margin-bottom: 0.5rem;
  }
  .btn-start {
    background: var(--accent);
    color: #000;
  }
  .btn-start:hover { filter: brightness(1.1); box-shadow: 0 0 16px rgba(0,229,160,0.3); }

  .btn-stop {
    background: transparent;
    border: 1px solid var(--danger);
    color: var(--danger);
  }
  .btn-stop:hover { background: rgba(255,77,109,0.1); }

  .btn-next {
    background: var(--surface);
    border: 1px solid var(--border);
    color: var(--text);
  }
  .btn-next:hover { border-color: var(--accent2); color: var(--accent2); }

  /* Log */
  .log-box {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.72rem;
    color: var(--muted);
    line-height: 1.7;
    max-height: 220px;
    overflow-y: auto;
    background: var(--bg);
    border: 1px solid var(--border);
    border-radius: 6px;
    padding: 0.75rem;
  }
  .log-entry { border-bottom: 1px solid var(--border); padding: 0.2rem 0; }
  .log-entry:last-child { border-bottom: none; color: var(--text); }

  /* Demo banner */
  .demo-banner {
    background: rgba(123,97,255,0.1);
    border: 1px solid var(--accent2);
    border-radius: 8px;
    padding: 0.75rem 1rem;
    font-size: 0.8rem;
    color: var(--accent2);
    margin-bottom: 1rem;
    line-height: 1.5;
  }

  /* Progress bar */
  .progress-bar {
    height: 3px;
    background: var(--border);
    border-radius: 2px;
    margin-top: 0.5rem;
    overflow: hidden;
  }
  .progress-fill {
    height: 100%;
    background: linear-gradient(90deg, var(--accent), var(--accent2));
    border-radius: 2px;
    transition: width 0.4s ease;
  }
</style>
</head>
<body>
<header>
  <div class="logo">Slide<span>Voice</span></div>
  <div class="status-pill" id="statusPill">En attente</div>
</header>

<main>
  <div class="panel" id="mainPanel">
    <div class="section-title">Script balisé — timeline</div>
    <div class="slides-list" id="slidesList">
      <div style="color: var(--muted); font-size: 0.9rem;">Chargement…</div>
    </div>
  </div>

  <div class="sidebar">
    <div class="sidebar-section">
      <div class="section-title">Microphone</div>
      <div class="wave-container" id="waveContainer">
        <div class="wave-bar"></div><div class="wave-bar"></div>
        <div class="wave-bar"></div><div class="wave-bar"></div>
        <div class="wave-bar"></div><div class="wave-bar"></div>
        <div class="wave-bar"></div>
      </div>
      <div class="section-title" style="margin-top: 0.5rem; margin-bottom: 0.4rem;">Dernier mot entendu</div>
      <div class="last-heard" id="lastHeard">—</div>
    </div>

    <div class="sidebar-section">
      <div class="section-title">Contrôles</div>
      <div id="demoBanner" class="demo-banner" style="display:none">
        ⚠️ Mode démo — les dépendances Python ne sont pas toutes installées. Les slides ne s'avanceront pas réellement.
      </div>
      <button class="btn btn-start" onclick="startListening()">▶ Démarrer l'écoute</button>
      <button class="btn btn-stop" onclick="stopListening()">■ Arrêter</button>
      <button class="btn btn-next" onclick="manualNext()">→ Slide suivante (manuel)</button>
      <div style="display:flex; gap: 0.5rem; margin-bottom: 0.5rem; align-items: center;">
        <input id="syncInput" type="number" min="1" placeholder="N°" style="width:64px;padding:0.55rem 0.6rem;border-radius:6px;border:1px solid var(--border);background:var(--bg);color:var(--text);font-family:'JetBrains Mono',monospace;font-size:0.85rem;text-align:center;">
        <button class="btn btn-next" style="margin-bottom:0;flex:1;" onclick="syncSlide()">⟳ Sync — je suis à la slide N°</button>
      </div>
      <div class="progress-bar">
        <div class="progress-fill" id="progressFill" style="width: 0%"></div>
      </div>
    </div>

    <div class="sidebar-section" style="flex: 1; overflow: hidden; display: flex; flex-direction: column;">
      <div class="section-title">Journal</div>
      <div class="log-box" id="logBox"></div>
    </div>
  </div>
</main>

<script>
let pollInterval = null;

async function fetchState() {
  const r = await fetch('/api/state');
  return r.json();
}

async function postAction(action, data = {}) {
  await fetch('/api/action', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({action, ...data})
  });
}

function renderSlides(slides, current) {
  const list = document.getElementById('slidesList');
  if (!slides.length) {
    list.innerHTML = '<div style="color: var(--muted); font-size: 0.9rem;">Aucune slide chargée.</div>';
    return;
  }
  list.innerHTML = slides.map((s, i) => `
    <div class="slide-item ${i === current ? 'active' : i < current ? 'done' : ''}">
      <div class="slide-num">${String(i + 1).padStart(2, '0')}</div>
      <div>
        <div class="slide-keyword">[${s.keyword}]</div>
        <div class="slide-text">${s.text || '<em>pas de texte</em>'}</div>
      </div>
    </div>
  `).join('');

  // Scroll active slide into view
  const active = list.querySelector('.active');
  if (active) active.scrollIntoView({behavior: 'smooth', block: 'nearest'});
}

function updateUI(s) {
  renderSlides(s.slides, s.current);

  // Status pill
  const pill = document.getElementById('statusPill');
  pill.textContent = s.listening ? '🎙 Écoute' : 'En attente';
  pill.className = 'status-pill' + (s.listening ? ' active' : '');

  // Wave animation
  document.getElementById('waveContainer').className =
    'wave-container' + (s.listening ? ' listening' : '');

  // Last heard
  document.getElementById('lastHeard').textContent = s.last_heard || '—';

  // Progress
  const pct = s.slides.length > 1
    ? (s.current / (s.slides.length - 1)) * 100
    : 0;
  document.getElementById('progressFill').style.width = pct + '%';

  // Log
  const logBox = document.getElementById('logBox');
  logBox.innerHTML = [...s.log].reverse()
    .map(l => `<div class="log-entry">${l}</div>`).join('');

  // Demo banner
  document.getElementById('demoBanner').style.display = s.demo_mode ? 'block' : 'none';
}

async function startListening() {
  await postAction('start');
}

async function stopListening() {
  await postAction('stop');
}

async function manualNext() {
  await postAction('next');
}

async function syncSlide() {
  const n = parseInt(document.getElementById('syncInput').value);
  if (!n || n < 1) return;
  await postAction('sync', {slide: n});
}

async function poll() {
  try {
    const s = await fetchState();
    updateUI(s);
  } catch(e) {}
}

poll();
pollInterval = setInterval(poll, 800);
</script>
</body>
</html>
"""


class Handler(BaseHTTPRequestHandler):
    def log_message(self, *args):
        pass  # Silence les logs HTTP

    def do_GET(self):
        if self.path == "/" or self.path == "/index.html":
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(HTML_PAGE.encode("utf-8"))
        elif self.path == "/api/state":
            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(json.dumps(state).encode("utf-8"))
        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        if self.path == "/api/action":
            length = int(self.headers.get("Content-Length", 0))
            body = json.loads(self.rfile.read(length))
            action = body.get("action")

            if action == "start" and not state["listening"]:
                state["listening"] = True
                t = threading.Thread(target=voice_loop, daemon=True)
                t.start()
                log("▶ Écoute démarrée")

            elif action == "stop":
                state["listening"] = False
                log("■ Écoute arrêtée")

            elif action == "next":
                if state["current"] + 1 < len(state["slides"]):
                    state["current"] += 1
                    advance_slide()
                    log(f"→ Avancement manuel vers slide {state['current'] + 1}")

            elif action == "sync":
                n = body.get("slide", 1)
                idx = max(0, min(int(n) - 1, len(state["slides"]) - 1))
                state["current"] = idx
                log(f"⟳ Synchronisé sur la slide {idx + 1} — en attente du mot-clé suivant")

            self.send_response(200)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(b'{"ok":true}')


# ── Point d'entrée ────────────────────────────────────────────────────────────
def main():
    script_path = sys.argv[1] if len(sys.argv) > 1 else None

    if script_path:
        slides = parse_script(script_path)
        state["slides"] = slides
        log(f"📄 Script chargé : {len(slides)} slide(s)")
    else:
        # Exemple de démonstration
        state["slides"] = [
            {"keyword": "bonjour", "text": "Bonjour à tous, aujourd'hui je vais vous présenter…"},
            {"keyword": "contexte", "text": "Pour comprendre le sujet, voici le contexte économique…"},
            {"keyword": "analyse", "text": "Notre analyse montre que les données indiquent…"},
            {"keyword": "conclusion", "text": "En conclusion, nous pouvons retenir trois points clés…"},
        ]
        log("ℹ️  Aucun script fourni — données de démonstration chargées")

    state["demo_mode"] = not (PYAUTOGUI_OK and (WHISPER_OK or SR_OK))

    port = 8742
    server = HTTPServer(("localhost", port), Handler)
    url = f"http://localhost:{port}"
    log(f"🌐 Interface disponible sur {url}")

    threading.Timer(0.5, lambda: webbrowser.open(url)).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        log("👋 SlideVoice arrêté")


if __name__ == "__main__":
    main()