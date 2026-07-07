# Excel Tutor — Live Coach

A PWA that teaches Excel from beginner to intermediate. Nothing is pre-written —
every lesson, tip, and quiz question is generated live by Groq's API based on
the topic you pick or the question you ask.

## Run it locally

Service workers (needed for offline/installable behavior) don't work over
`file://`, so serve the folder over localhost. Any of these work:

```bash
# Option 1: Python
python3 -m http.server 8000

# Option 2: Node
npx serve .
```

Then open `http://localhost:8000` in your browser. Chrome/Edge will offer an
"Install" icon in the address bar once it's running — that installs it as a
real app (PWA).

## Get a free Groq API key

1. Go to https://console.groq.com/keys
2. Sign up (no credit card needed)
3. Create a key (starts with `gsk_...`)
4. Paste it into the app's setup screen — it's stored only in this browser's
   localStorage, never sent anywhere except directly to Groq's API.

Free tier is generous enough for personal tutoring use: roughly 30 requests/
minute and up to ~14,400/day depending on model, per Groq's published limits.

## How it's structured

- `index.html` / `styles.css` / `app.js` — the whole app (no build step, no framework)
- `manifest.json` / `sw.js` — PWA install + offline app-shell caching
- **Learn tab** — a skill tree (Foundations → Everyday functions → Level up).
  Clicking a topic asks Groq to teach that topic live, tailored to your
  current level and what you've already mastered, ending in a quick quiz.
- **Chat tab** — open-ended Q&A with the same tutor persona.
- **Live visuals** — the tutor can render inline mini spreadsheet grids (with
  highlighted ranges), charts (bar/line/pie via Chart.js), and step-flow
  diagrams whenever they'd clarify a topic. These are generated as structured
  data by the model and drawn by the app's own code, so they render reliably
  every time instead of relying on the model to draw pixels.
- **Practice sandbox** — some lessons include an editable mini-grid where you
  actually type a formula into a real cell. A small built-in formula engine
  (`formula-engine.js`, no external dependency) evaluates it live — SUM,
  AVERAGE, COUNT, MIN/MAX, IF, AND/OR, ROUND, CONCAT, ranges, cell refs — and
  checks it against the expected answer, with a hint button if you get stuck.
- **Export to real Excel** — any grid or completed practice challenge has a
  "Download as Excel" button (via SheetJS) that writes an actual `.xlsx` file
  you can open in Excel, formulas included.
- **Weak-spot tracking** — every missed quiz/challenge answer is logged
  (topic + question) in your local progress. The Progress tab surfaces a
  "Worth revisiting" list with a one-tap "Review weak spots with tutor"
  button that asks for a targeted refresher. The tutor also gets a running
  summary of your mastered topics and weak spots on every request, so it
  naturally references your history instead of starting cold each time.
- **Session continuity** — your last ~20 chat messages are saved locally and
  restored on reload, so refreshing mid-conversation doesn't lose context.
- **Optional server-side key proxy** — see `server/README.md`. By default the
  app runs in bring-your-own-key mode; deploying the included Cloudflare
  Worker removes the key-entry step entirely for anyone you share the app with.
- **Progress tab** — tracked entirely in `localStorage` on this device:
  topics mastered, quiz accuracy, and a day streak. Nothing is sent to a server.

## Deploying so it's installable from anywhere

Push this folder to GitHub Pages, Netlify, Vercel, or any static host — it's
plain HTML/CSS/JS, no backend required. HTTPS (which those all provide) is
required for the service worker and install prompt to work outside of
localhost.
