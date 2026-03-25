# slidevoice
Slidevoice was made to soluce a problem : 
When giving a presentation, you should always focus on the text, but don’t forget to move on to the next slide.
With SlideVoice, you no longer need to worry about PowerPoint. 

## Is slidevoice vibe coded ? Can I contribute ?
Yes, I'm not a pro developer, but feel free to copy this project and do it better on your own. You can open a pull request or fork my repo.

## SlideVoice 

**Automatically advance your PowerPoint slides by saying a keyword out loud.**

SlideVoice listens to your microphone while you present. When it hears the keyword you've assigned to the next slide, it instantly sends the right arrow key directly to PowerPoint — no clicking, no remote, no interruption.

It also exposes a phone-friendly remote control page on your local network, so you can tap your phone as a backup if voice recognition misses something.

---

## How it works

1. You write your presentation script in a `.txt` file, with `[keyword]` tags marking each slide transition
2. SlideVoice listens continuously using **Whisper** (local, offline, no API key needed)
3. When a keyword is detected, it sends the `→` key directly to PowerPoint via Windows API — no focus switching needed
4. A local web interface shows you the current slide, the transcript, and live logs

---

## Installation

```bash
pip install openai-whisper pyautogui sounddevice numpy pywin32
```

> On Windows you may also need:
> ```bash
> pip install pyaudio
> ```

---

## Script format

Create a `.txt` file with `[keyword]` tags. Each tag marks the word you'll say naturally in your speech to trigger the next slide.

**The first tag is your starting slide — its keyword is never triggered.** Every subsequent tag is triggered by saying its keyword.

```
[intro]
Good morning everyone. Today I'll be talking about...

[context]
To understand the topic, let's look at the context first...

[analysis]
Moving on to the analysis, the data shows that...

[conclusion]
To wrap things up, here are the three key takeaways...
```

**Tips for good keywords:**
- Use **short, single words** — they're recognized far more reliably than long phrases
- Pick words you'd say **naturally** in your transition sentence
- Avoid common filler words like "so", "now", "then" — they'll trigger too easily

---

## Usage

### 1. Start the script

```bash
python slidevoice.py my_script.txt
```

The web interface opens automatically in your browser. The first run will download the Whisper model (~140 MB for `base`).

### 2. Open PowerPoint in presentation mode

Press `F5` in PowerPoint to start the slideshow.

### 3. Start listening

Click **▶ Start listening** in the SlideVoice interface.

### 4. Present normally

Say your keywords naturally as you speak. SlideVoice will detect them and advance the slide automatically.

---

## Phone remote control

When the script starts, it prints a URL in the terminal:

```
📱 Phone remote: http://192.168.1.x:8742/remote
```

Open this URL on your phone (same Wi-Fi network). You get a full-screen remote with a large **Next slide** button and a smaller **Previous** button — useful as a backup when voice recognition misses a word.

---

## Web interface

The browser interface (at `http://localhost:8742`) shows:

- **Timeline** — all your slides with the current one highlighted
- **Waveform** — visual indicator when the mic is active
- **Last heard** — live transcript of what Whisper recognized
- **Sync** — if you start mid-presentation, type your current slide number and click Sync to re-align the keyword detector
- **Manual controls** — next/previous buttons and the sync field

---

## Fuzzy keyword matching

SlideVoice doesn't require an exact match. It ignores stopwords (de, le, la, les…) and matches on word roots, so `[competences]` will match "his competence", "their competencies", etc.

---

## Dependencies

| Package | Role |
|---|---|
| `openai-whisper` | Local speech recognition (offline) |
| `pyautogui` | Fallback key simulation |
| `pywin32` | Direct keypress to PowerPoint window (Windows) |
| `sounddevice` | Microphone access |
| `numpy` | Audio processing |

---

## Whisper model sizes

You can change the model in `slidevoice.py` (`tiny`, `base`, `small`, `medium`):

| Model | Size | Speed | Quality |
|---|---|---|---|
| `tiny` | 40 MB | fastest | good |
| `base` | 140 MB | fast | better ← default |
| `small` | 460 MB | medium | great |
| `medium` | 1.5 GB | slow | excellent |

---

## Requirements

- Windows (uses `pywin32` for direct window targeting)
- Python 3.9+
- PowerPoint open in slideshow mode (`F5`)
- Microphone
