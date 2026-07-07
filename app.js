/* ============================================================
   Excel Tutor — live AI coach powered by Groq
   No canned lessons: every explanation is generated on request.
   ============================================================ */

const GROQ_MODEL = 'llama-3.3-70b-versatile';
const GROQ_URL = 'https://api.groq.com/openai/v1/chat/completions';

const SYSTEM_PROMPT = `You are "Excel Tutor," a live, patient Excel coach helping a learner progress from beginner to intermediate. You have no pre-written curriculum — every explanation is generated fresh for exactly what's asked.

Rules:
- Be conversational, concrete, and example-driven. Short paragraphs, no filler, no long intros.
- Prefer the easiest/fastest real way to do something in Excel (shortcuts, built-in tools) over convoluted formulas — call out the quicker route when one exists.
- Use inline code formatting for cell references, formulas, and function names, e.g. \`=VLOOKUP(A2,Table1,2,FALSE)\`.
- When teaching a specific topic (not a casual question), structure the reply as: a short plain-English explanation, then one concrete worked example, then a single comprehension check-in.
- You can render live visuals inline, in place of describing a table/chart/process in words. Use them whenever they'd make the concept clearer — not on every message. Insert them as fenced blocks exactly where they belong in your explanation:

1) A small spreadsheet grid — for cell references, formulas, ranges, sorting, lookups, conditional formatting, anything tied to specific cells:
\`\`\`grid
{"cols": 3, "rows": 4, "cells": {"A1": "Item", "B1": "Qty", "C1": "Price", "A2": "Pen", "B2": "12", "C2": "1.20", "A3": "Mug", "B3": "5", "C3": "6.00"}, "highlight": ["C2", "C3"], "formulaCell": "C2", "formulaValue": "=B2*1.2", "caption": "optional one-line caption"}
\`\`\`
Keep grids small: max 6 cols (letters A-F only), max 8 rows. "highlight" and "formulaCell" are optional.

2) A chart — only for genuinely chart-related topics (trends, comparisons, the Charts & Visualization topic itself):
\`\`\`chart
{"type": "bar", "labels": ["Q1", "Q2", "Q3", "Q4"], "data": [12, 19, 14, 23], "label": "Revenue", "caption": "optional one-line caption"}
\`\`\`
"type" is one of: bar, line, pie.

3) A flow diagram — for step-based or logical-order topics (nested IF, VLOOKUP argument order, Power Query steps):
\`\`\`diagram
{"nodes": [{"label": "Start"}, {"label": "Step 2"}, {"label": "Step 3"}], "caption": "optional one-line caption"}
\`\`\`
Nodes render left to right in the order given — keep it to 2-5 short-labeled nodes.

4) An interactive practice challenge — use this instead of a grid (never both) when the learner should actually write a formula themselves rather than just observe one. Give a small starter grid, mark exactly the cell(s) they should fill in as editable, and specify one target cell with the value you expect:
\`\`\`challenge
{"description": "Work out the total cost of both items.", "cols": 3, "rows": 3, "cells": {"A1": "Item", "B1": "Qty", "C1": "Price", "A2": "Pen", "B2": "12", "C2": "1.20", "A3": "Mug", "B3": "5", "C3": "6.00"}, "editableCells": ["C4"], "targetCell": "C4", "expectedValue": 44.4, "tolerance": 0.05, "hint": "Try =B2*C2+B3*C3, or =SUMPRODUCT(B2:B3,C2:C3)."}
\`\`\`
Reference cells outside the given grid (like a totals row) are fine as long as they're within the max 6x8 size. Offer a challenge like this whenever it fits naturally, especially for formula-writing topics — reading about a formula and writing one are different skills, and the learner benefits from doing both.

Use at most one grid, one chart, one diagram, and one challenge per message — never more than one of the same type. Place each block right where it's discussed in your prose, not all bunched at the end.
- When a comprehension check-in fits, end your ENTIRE reply with exactly one fenced block like this, after any grid/chart/diagram, with nothing after it:
\`\`\`quiz
{"question": "...", "options": ["...", "...", "...", "..."], "correctIndex": 0, "explanation": "..."}
\`\`\`
- For casual open questions that aren't a full topic lesson, you may omit the quiz block, but still use grid/chart/diagram blocks whenever they'd clarify your answer.
- Adapt depth and vocabulary to the learner's stated level and progress summary.
- Each request includes a separate system message summarizing what the learner has mastered and previously struggled with. Weave in brief, natural refreshers on struggled areas when relevant — don't force it if it doesn't fit.
- Never claim you lack real-time knowledge — you always answer live from your own Excel expertise.`;

const SKILL_TREE = [
  {
    unit: 'Foundations',
    level: 'beginner',
    topics: [
      { id: 'interface-nav', name: 'Interface & navigation' },
      { id: 'entering-data', name: 'Entering & editing data' },
      { id: 'basic-formulas', name: 'Basic formulas (+ − × ÷)' },
      { id: 'cell-refs', name: 'Cell references ($ absolute vs relative)' },
      { id: 'formatting', name: 'Formatting cells & numbers' },
      { id: 'sort-filter', name: 'Sorting & filtering' },
    ],
  },
  {
    unit: 'Everyday functions',
    level: 'beginner',
    topics: [
      { id: 'sum-avg-count', name: 'SUM, AVERAGE, COUNT, MIN/MAX' },
      { id: 'if-basics', name: 'IF function basics' },
      { id: 'text-functions', name: 'Text functions (CONCAT, LEFT/RIGHT, TRIM)' },
      { id: 'date-time', name: 'Dates & time basics' },
    ],
  },
  {
    unit: 'Level up',
    level: 'intermediate',
    topics: [
      { id: 'lookup', name: 'VLOOKUP / XLOOKUP' },
      { id: 'nested-logic', name: 'Nested IF, AND/OR' },
      { id: 'cond-format', name: 'Conditional formatting' },
      { id: 'data-validation', name: 'Data validation' },
      { id: 'named-ranges', name: 'Named ranges' },
      { id: 'pivot-tables', name: 'PivotTables' },
      { id: 'charts', name: 'Charts & visualization' },
      { id: 'shortcuts', name: 'Keyboard shortcuts & speed' },
      { id: 'power-query', name: 'Intro to Power Query' },
    ],
  },
];

const COL_LETTERS = ['A', 'B', 'C', 'D', 'E'];

/* ---------------- State ---------------- */
const state = {
  apiKey: localStorage.getItem('excelTutor.apiKey') || '',
  progress: loadProgress(),
  currentTab: 'learn',
  currentTopicId: null,
  chatHistory: [], // {role, content}
  streaming: false,
};

function loadProgress() {
  try {
    const raw = JSON.parse(localStorage.getItem('excelTutor.progress') || '{}');
    return {
      completed: raw.completed || [],
      quizCorrect: raw.quizCorrect || 0,
      quizTotal: raw.quizTotal || 0,
      streak: raw.streak || 0,
      lastVisit: raw.lastVisit || null,
      weakSpots: raw.weakSpots || [], // [{topicId, topicName, question, ts}]
    };
  } catch {
    return { completed: [], quizCorrect: 0, quizTotal: 0, streak: 0, lastVisit: null, weakSpots: [] };
  }
}

function saveProgress() {
  localStorage.setItem('excelTutor.progress', JSON.stringify(state.progress));
}

function saveChatHistory() {
  localStorage.setItem('excelTutor.chatHistory', JSON.stringify(state.chatHistory.slice(-20)));
}

function loadChatHistory() {
  try {
    return JSON.parse(localStorage.getItem('excelTutor.chatHistory') || '[]');
  } catch {
    return [];
  }
}

function updateStreak() {
  const today = new Date().toISOString().slice(0, 10);
  const last = state.progress.lastVisit;
  if (last === today) return;
  if (last) {
    const diffDays = Math.round((new Date(today) - new Date(last)) / 86400000);
    state.progress.streak = diffDays === 1 ? state.progress.streak + 1 : 1;
  } else {
    state.progress.streak = 1;
  }
  state.progress.lastVisit = today;
  saveProgress();
}

function allTopics() {
  return SKILL_TREE.flatMap((u) => u.topics.map((t) => ({ ...t, level: u.level, unit: u.unit })));
}

function computeLevel() {
  const beginnerTopics = allTopics().filter((t) => t.level === 'beginner');
  const doneBeginner = beginnerTopics.filter((t) => state.progress.completed.includes(t.id)).length;
  return doneBeginner >= Math.ceil(beginnerTopics.length * 0.6) ? 'intermediate' : 'beginner';
}

function buildLearnerContextMessage() {
  const level = computeLevel();
  const completedNames = allTopics()
    .filter((t) => state.progress.completed.includes(t.id))
    .map((t) => t.name);
  const weakNames = [...new Set(state.progress.weakSpots.map((w) => w.topicName).filter(Boolean))].slice(-5);
  return {
    role: 'system',
    content: `Learner context — level: ${level}. Topics mastered: ${
      completedNames.join(', ') || 'none yet'
    }. Areas previously missed in quizzes: ${weakNames.join(', ') || 'none yet'}.`,
  };
}

/* ---------------- DOM refs ---------------- */
const onboardView = document.getElementById('onboardView');
const mainView = document.getElementById('mainView');
const apiKeyInput = document.getElementById('apiKeyInput');
const saveKeyBtn = document.getElementById('saveKeyBtn');
const skillTree = document.getElementById('skillTree');
const thread = document.getElementById('thread');
const learnPanel = document.getElementById('learnPanel');
const chatPanel = document.getElementById('chatPanel');
const progressPanel = document.getElementById('progressPanel');
const sheetTabs = document.getElementById('sheetTabs');
const composer = document.getElementById('composer');
const chatInput = document.getElementById('chatInput');
const sendBtn = document.getElementById('sendBtn');
const cellRef = document.getElementById('cellRef');
const breadcrumb = document.getElementById('breadcrumb');
const levelPill = document.getElementById('levelPill');
const levelFill = document.getElementById('levelFill');
const levelText = document.getElementById('levelText');
const statTopics = document.getElementById('statTopics');
const statStreak = document.getElementById('statStreak');
const statQuiz = document.getElementById('statQuiz');
const progressList = document.getElementById('progressList');
const resetProgressBtn = document.getElementById('resetProgressBtn');
const gearBtn = document.getElementById('gearBtn');
const settingsOverlay = document.getElementById('settingsOverlay');
const settingsClose = document.getElementById('settingsClose');
const settingsBackendMode = document.getElementById('settingsBackendMode');
const settingsKeyRow = document.getElementById('settingsKeyRow');
const apiKeyInputSettings = document.getElementById('apiKeyInputSettings');
const saveKeySettingsBtn = document.getElementById('saveKeySettingsBtn');
const forgetKeyBtn = document.getElementById('forgetKeyBtn');
const weakSpotsBlock = document.getElementById('weakSpotsBlock');
const weakSpotsList = document.getElementById('weakSpotsList');
const reviewWeakBtn = document.getElementById('reviewWeakBtn');

const USE_BACKEND = typeof BACKEND_URL !== 'undefined' && BACKEND_URL.trim().length > 0;

/* ---------------- Onboarding ---------------- */
function boot() {
  if (USE_BACKEND || state.apiKey) {
    onboardView.classList.add('hidden');
    mainView.classList.remove('hidden');
    sheetTabs.classList.remove('hidden');
    composer.classList.remove('hidden');
    updateStreak();
    renderSkillTree();
    renderProgress();
    updateHeader();
    restorePersistedChat();
  } else {
    onboardView.classList.remove('hidden');
    mainView.classList.add('hidden');
  }
}

function restorePersistedChat() {
  const saved = loadChatHistory();
  if (!saved.length || state.chatHistory.length) return;
  state.chatHistory = saved;
  clearThread();
  saved.forEach((m) => {
    if (m.role === 'user') {
      addUserBubble(m.content);
    } else if (m.role === 'assistant') {
      const bubble = addTutorBubble();
      finalRenderBubble(bubble, m.content, null);
    }
  });
}

saveKeyBtn.addEventListener('click', () => {
  const key = apiKeyInput.value.trim();
  if (!key) return;
  state.apiKey = key;
  localStorage.setItem('excelTutor.apiKey', key);
  boot();
});

/* ---------------- Settings modal ---------------- */
function openSettings() {
  if (USE_BACKEND) {
    settingsBackendMode.textContent = 'Connected through a shared backend — no personal API key needed.';
    settingsKeyRow.classList.add('hidden');
    forgetKeyBtn.classList.add('hidden');
  } else {
    settingsBackendMode.textContent = 'Running in bring-your-own-key mode. Your key is stored only on this device.';
    settingsKeyRow.classList.remove('hidden');
    forgetKeyBtn.classList.remove('hidden');
    apiKeyInputSettings.value = '';
  }
  settingsOverlay.classList.remove('hidden');
}
function closeSettings() {
  settingsOverlay.classList.add('hidden');
}
gearBtn.addEventListener('click', openSettings);
settingsClose.addEventListener('click', closeSettings);
settingsOverlay.addEventListener('click', (e) => {
  if (e.target === settingsOverlay) closeSettings();
});
saveKeySettingsBtn.addEventListener('click', () => {
  const key = apiKeyInputSettings.value.trim();
  if (!key) return;
  state.apiKey = key;
  localStorage.setItem('excelTutor.apiKey', key);
  closeSettings();
});
forgetKeyBtn.addEventListener('click', () => {
  if (!confirm('Remove the saved API key from this device? Your progress will stay intact.')) return;
  localStorage.removeItem('excelTutor.apiKey');
  state.apiKey = '';
  closeSettings();
  boot();
});

/* ---------------- Tabs ---------------- */
sheetTabs.addEventListener('click', (e) => {
  const tab = e.target.closest('.sheet-tab');
  if (!tab) return;
  switchTab(tab.dataset.tab);
});

function switchTab(name) {
  state.currentTab = name;
  [...sheetTabs.children].forEach((el) => el.classList.toggle('active', el.dataset.tab === name));
  learnPanel.classList.toggle('hidden', name !== 'learn');
  chatPanel.classList.toggle('hidden', name !== 'chat');
  progressPanel.classList.toggle('hidden', name !== 'progress');
  composer.classList.toggle('hidden', name === 'progress');
  if (name === 'progress') renderProgress();
  updateHeader();
}

function updateHeader() {
  const level = computeLevel();
  levelPill.textContent = level === 'intermediate' ? 'Intermediate' : 'Beginner';
  if (state.currentTab === 'learn') {
    cellRef.textContent = 'A1';
    breadcrumb.textContent = 'Learn > Pick a topic to start';
  } else if (state.currentTab === 'chat') {
    const t = allTopics().find((x) => x.id === state.currentTopicId);
    cellRef.textContent = t ? refFor(t.id) : 'B1';
    breadcrumb.textContent = t ? `Learn > ${t.unit} > ${t.name}` : 'Chat > Ask anything';
  } else {
    cellRef.textContent = 'C1';
    breadcrumb.textContent = 'Progress > Your journey';
  }
}

function refFor(topicId) {
  for (let u = 0; u < SKILL_TREE.length; u++) {
    const idx = SKILL_TREE[u].topics.findIndex((t) => t.id === topicId);
    if (idx > -1) return `${COL_LETTERS[u % COL_LETTERS.length]}${idx + 1}`;
  }
  return 'A1';
}

/* ---------------- Render: skill tree ---------------- */
function renderSkillTree() {
  skillTree.innerHTML = '';
  SKILL_TREE.forEach((unit, uIdx) => {
    const unitEl = document.createElement('div');
    unitEl.className = 'unit';
    unitEl.innerHTML = `<div class="unit-title">${unit.unit} · ${unit.level}</div>`;
    const grid = document.createElement('div');
    grid.className = 'topic-grid';
    unit.topics.forEach((topic, tIdx) => {
      const done = state.progress.completed.includes(topic.id);
      const card = document.createElement('button');
      card.className = 'topic-card' + (done ? ' done' : '');
      card.innerHTML = `<span class="ref">${COL_LETTERS[uIdx % COL_LETTERS.length]}${tIdx + 1}</span><span class="name">${topic.name}</span>`;
      card.addEventListener('click', () => startTopic(topic, unit));
      grid.appendChild(card);
    });
    unitEl.appendChild(grid);
    skillTree.appendChild(unitEl);
  });
}

/* ---------------- Render: progress ---------------- */
function renderProgress() {
  const topics = allTopics();
  const done = topics.filter((t) => state.progress.completed.includes(t.id)).length;
  const pct = Math.round((done / topics.length) * 100);
  levelFill.style.width = pct + '%';
  levelText.textContent = `${done} of ${topics.length} topics mastered`;
  statTopics.textContent = done;
  statStreak.textContent = state.progress.streak;
  const acc = state.progress.quizTotal ? Math.round((state.progress.quizCorrect / state.progress.quizTotal) * 100) : 0;
  statQuiz.textContent = acc + '%';

  progressList.innerHTML = '';
  SKILL_TREE.forEach((unit, uIdx) => {
    const unitEl = document.createElement('div');
    unitEl.className = 'unit';
    unitEl.innerHTML = `<div class="unit-title">${unit.unit}</div>`;
    const grid = document.createElement('div');
    grid.className = 'topic-grid';
    unit.topics.forEach((topic, tIdx) => {
      const isDone = state.progress.completed.includes(topic.id);
      const card = document.createElement('div');
      card.className = 'topic-card' + (isDone ? ' done' : '');
      card.innerHTML = `<span class="ref">${COL_LETTERS[uIdx % COL_LETTERS.length]}${tIdx + 1}</span><span class="name">${topic.name}</span>`;
      grid.appendChild(card);
    });
    unitEl.appendChild(grid);
    progressList.appendChild(unitEl);
  });

  renderWeakSpots();
}

function renderWeakSpots() {
  const spots = state.progress.weakSpots.slice(-6).reverse();
  if (!spots.length) {
    weakSpotsBlock.classList.add('hidden');
    return;
  }
  weakSpotsBlock.classList.remove('hidden');
  weakSpotsList.innerHTML = '';
  spots.forEach((w) => {
    const item = document.createElement('div');
    item.className = 'weak-spot-item';
    item.innerHTML = `<span class="weak-spot-topic">${escapeHtml(w.topicName || 'General')}</span><span class="weak-spot-q">${escapeHtml(
      w.question || ''
    )}</span>`;
    weakSpotsList.appendChild(item);
  });
}

reviewWeakBtn.addEventListener('click', () => {
  const topics = [...new Set(state.progress.weakSpots.map((w) => w.topicName).filter(Boolean))].slice(-5);
  switchTab('chat');
  if (thread.querySelector('.empty-state')) clearThread();
  state.currentTopicId = null;
  const prompt = topics.length
    ? `I'd like a quick refresher on things I've struggled with: ${topics.join(', ')}. Give me a short recap and a fresh check-in question for the trickiest one.`
    : `Give me a quick review of anything I've gotten wrong in quizzes recently, with a fresh check-in question.`;
  sendToTutor(prompt, { isTopicLesson: false, showUserBubble: true });
});

resetProgressBtn.addEventListener('click', () => {
  if (!confirm('Reset all progress on this device? This cannot be undone.')) return;
  state.progress = {
    completed: [],
    quizCorrect: 0,
    quizTotal: 0,
    streak: state.progress.streak,
    lastVisit: state.progress.lastVisit,
    weakSpots: [],
  };
  saveProgress();
  renderSkillTree();
  renderProgress();
});

/* ---------------- Topic lesson flow ---------------- */
function startTopic(topic, unit) {
  state.currentTopicId = topic.id;
  state.chatHistory = [];
  switchTab('chat');
  clearThread();
  addSystemNote(`${unit.unit} · ${topic.name}`);
  const prompt = `Teach me: "${topic.name}" in Excel. Give me the lesson now.`;
  sendToTutor(prompt, { isTopicLesson: true });
}

/* ---------------- Chat composer ---------------- */
chatInput.addEventListener('input', () => {
  sendBtn.disabled = !chatInput.value.trim() || state.streaming;
  chatInput.style.height = 'auto';
  chatInput.style.height = Math.min(chatInput.scrollHeight, 90) + 'px';
});
chatInput.addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && !e.shiftKey) {
    e.preventDefault();
    if (!sendBtn.disabled) submitChat();
  }
});
sendBtn.addEventListener('click', submitChat);

function submitChat() {
  const text = chatInput.value.trim();
  if (!text || state.streaming) return;
  if (state.currentTab !== 'chat') switchTab('chat');
  if (thread.querySelector('.empty-state')) clearThread();
  chatInput.value = '';
  chatInput.style.height = 'auto';
  sendBtn.disabled = true;
  sendToTutor(text, { isTopicLesson: false, showUserBubble: true });
}

/* ---------------- Rendering helpers ---------------- */
function clearThread() {
  thread.innerHTML = '';
}
function addSystemNote(text) {
  const el = document.createElement('div');
  el.className = 'msg system-note';
  el.textContent = text;
  thread.appendChild(el);
  thread.scrollTop = thread.scrollHeight;
}
function addUserBubble(text) {
  const el = document.createElement('div');
  el.className = 'msg user';
  el.textContent = text;
  thread.appendChild(el);
  thread.scrollTop = thread.scrollHeight;
}
function addTutorBubble() {
  const el = document.createElement('div');
  el.className = 'msg tutor';
  el.innerHTML = '<span class="typing-dots"><span></span><span></span><span></span></span>';
  thread.appendChild(el);
  thread.scrollTop = thread.scrollHeight;
  return el;
}

function escapeHtml(s) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function renderMarkdownLite(text) {
  let html = escapeHtml(text);
  html = html.replace(/```([a-z]*)\n([\s\S]*?)```/g, (m, lang, code) => `<pre><code>${code}</code></pre>`);
  html = html.replace(/`([^`]+)`/g, '<code>$1</code>');
  html = html.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
  html = html
    .split(/\n{2,}/)
    .map((p) => `<p>${p.replace(/\n/g, '<br>')}</p>`)
    .join('');
  return html;
}

/* During streaming: show only the prose that precedes the first fenced block,
   so half-formed JSON never flashes on screen. */
function liveRenderDuringStream(bubbleEl, fullText) {
  const fenceIdx = fullText.indexOf('```');
  const visible = fenceIdx > -1 ? fullText.slice(0, fenceIdx) : fullText;
  bubbleEl.innerHTML =
    renderMarkdownLite(visible) + (fenceIdx > -1 ? '<div class="prep-note mono">preparing visual…</div>' : '');
}

/* Split full response text into ordered segments: plain prose and typed
   fenced blocks (grid / chart / diagram / quiz), preserving their order. */
function parseSegments(fullText) {
  const blockRegex = /```(grid|chart|diagram|challenge|quiz)\s*([\s\S]*?)```/g;
  const segments = [];
  let lastIndex = 0;
  let match;
  while ((match = blockRegex.exec(fullText)) !== null) {
    const [whole, type, body] = match;
    if (match.index > lastIndex) {
      segments.push({ kind: 'text', content: fullText.slice(lastIndex, match.index) });
    }
    let data = null;
    try {
      data = JSON.parse(body.trim());
    } catch {
      data = null;
    }
    if (data) segments.push({ kind: type, data });
    lastIndex = match.index + whole.length;
  }
  if (lastIndex < fullText.length) {
    segments.push({ kind: 'text', content: fullText.slice(lastIndex) });
  }
  return segments;
}

/* Final render: builds the bubble content from ordered segments, wiring up
   quiz interactivity and instantiating any charts. */
function finalRenderBubble(bubbleEl, fullText, topicId) {
  bubbleEl.innerHTML = '';
  const segments = parseSegments(fullText);
  segments.forEach((seg) => {
    if (seg.kind === 'text') {
      const trimmed = seg.content.trim();
      if (!trimmed) return;
      const div = document.createElement('div');
      div.innerHTML = renderMarkdownLite(trimmed);
      bubbleEl.appendChild(div);
    } else if (seg.kind === 'grid') {
      bubbleEl.appendChild(renderGrid(seg.data));
    } else if (seg.kind === 'chart') {
      bubbleEl.appendChild(renderChart(seg.data));
    } else if (seg.kind === 'diagram') {
      bubbleEl.appendChild(renderDiagram(seg.data));
    } else if (seg.kind === 'challenge') {
      bubbleEl.appendChild(renderChallenge(seg.data, topicId));
    } else if (seg.kind === 'quiz') {
      renderQuiz(bubbleEl, seg.data, topicId);
    }
  });
  if (!bubbleEl.children.length) {
    bubbleEl.innerHTML = renderMarkdownLite(fullText.trim() || "Sorry, I didn't catch that — try asking again.");
  }
}

const COL_LETTERS_FULL = ['A', 'B', 'C', 'D', 'E', 'F'];

function parseCellRef(ref) {
  const m = ref.trim().match(/^([A-Fa-f])(\d+)$/);
  if (!m) return null;
  return { col: m[1].toUpperCase(), row: parseInt(m[2], 10) };
}

function expandHighlight(refs) {
  const out = new Set();
  (refs || []).forEach((r) => {
    if (r.includes(':')) {
      const [a, b] = r.split(':');
      const ra = parseCellRef(a);
      const rb = parseCellRef(b);
      if (!ra || !rb) return;
      const c1 = COL_LETTERS_FULL.indexOf(ra.col);
      const c2 = COL_LETTERS_FULL.indexOf(rb.col);
      for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) {
        for (let row = Math.min(ra.row, rb.row); row <= Math.max(ra.row, rb.row); row++) {
          out.add(COL_LETTERS_FULL[c] + row);
        }
      }
    } else {
      out.add(r.trim().toUpperCase());
    }
  });
  return out;
}

/* ---- Excel export via SheetJS ---- */
function exportSheetToXlsx(sheetObj, cols, rows, filename) {
  if (typeof XLSX === 'undefined') {
    alert('The Excel export library is still loading — give it a second and try again.');
    return;
  }
  const wsData = [];
  for (let r = 1; r <= rows; r++) {
    const rowArr = [];
    for (let c = 0; c < cols; c++) {
      rowArr.push(sheetObj[COL_LETTERS_FULL[c] + r] ?? '');
    }
    wsData.push(rowArr);
  }
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  Object.keys(ws).forEach((addr) => {
    if (addr[0] === '!') return;
    const cell = ws[addr];
    if (typeof cell.v === 'string' && cell.v.trim().startsWith('=')) {
      cell.f = cell.v.trim().slice(1);
      delete cell.v;
    }
  });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  XLSX.writeFile(wb, filename);
}

function addExportButton(wrap, getSheetFn, cols, rows, filename) {
  const btn = document.createElement('button');
  btn.className = 'btn btn-ghost export-btn';
  btn.textContent = '⬇ Download as Excel';
  btn.addEventListener('click', () => exportSheetToXlsx(getSheetFn(), cols, rows, filename));
  wrap.appendChild(btn);
  return btn;
}

/* ---- Visual: mini spreadsheet grid ---- */
function renderGrid(data) {
  const wrap = document.createElement('div');
  wrap.className = 'visual-wrap';

  const cols = Math.min(Math.max(data.cols || 3, 1), 6);
  const rows = Math.min(Math.max(data.rows || 3, 1), 8);
  const cells = data.cells || {};
  const highlight = expandHighlight(data.highlight);

  if (data.formulaCell && data.formulaValue) {
    const strip = document.createElement('div');
    strip.className = 'grid-formula-strip mono';
    strip.innerHTML = `<span class="cell-ref">${escapeHtml(data.formulaCell.toUpperCase())}</span><span class="fx-icon">fx</span><span>${escapeHtml(
      data.formulaValue
    )}</span>`;
    wrap.appendChild(strip);
  }

  const table = document.createElement('table');
  table.className = 'mini-grid';

  const headRow = document.createElement('tr');
  headRow.appendChild(document.createElement('th'));
  for (let c = 0; c < cols; c++) {
    const th = document.createElement('th');
    th.textContent = COL_LETTERS_FULL[c];
    headRow.appendChild(th);
  }
  table.appendChild(headRow);

  for (let r = 1; r <= rows; r++) {
    const tr = document.createElement('tr');
    const rowHead = document.createElement('th');
    rowHead.textContent = r;
    rowHead.className = 'row-head';
    tr.appendChild(rowHead);
    for (let c = 0; c < cols; c++) {
      const ref = COL_LETTERS_FULL[c] + r;
      const td = document.createElement('td');
      if (highlight.has(ref)) td.classList.add('hl');
      td.textContent = cells[ref] || '';
      tr.appendChild(td);
    }
    table.appendChild(tr);
  }
  wrap.appendChild(table);

  if (data.caption) {
    const cap = document.createElement('div');
    cap.className = 'visual-caption';
    cap.textContent = data.caption;
    wrap.appendChild(cap);
  }

  addExportButton(wrap, () => cells, cols, rows, 'excel-tutor-example.xlsx');
  return wrap;
}

/* ---- Visual: interactive practice challenge ---- */
function renderChallenge(data, topicId) {
  const wrap = document.createElement('div');
  wrap.className = 'visual-wrap challenge-wrap';

  const cols = Math.min(Math.max(data.cols || 3, 1), 6);
  const rows = Math.min(Math.max(data.rows || 3, 1), 8);
  const startCells = { ...(data.cells || {}) };
  const editable = new Set((data.editableCells || []).map((r) => r.toUpperCase()));
  const targetCell = (data.targetCell || '').toUpperCase();
  const sheet = { ...startCells };

  if (data.description) {
    const desc = document.createElement('div');
    desc.className = 'challenge-desc';
    desc.textContent = data.description;
    wrap.appendChild(desc);
  }

  const table = document.createElement('table');
  table.className = 'mini-grid';
  const inputs = {};

  const headRow = document.createElement('tr');
  headRow.appendChild(document.createElement('th'));
  for (let c = 0; c < cols; c++) {
    const th = document.createElement('th');
    th.textContent = COL_LETTERS_FULL[c];
    headRow.appendChild(th);
  }
  table.appendChild(headRow);

  function getEngineValue(ref) {
    return sheet[ref];
  }

  function recompute() {
    const engine = new FormulaEngine(getEngineValue);
    for (let r = 1; r <= rows; r++) {
      for (let c = 0; c < cols; c++) {
        const ref = COL_LETTERS_FULL[c] + r;
        const cellEl = table.querySelector(`[data-static-ref="${ref}"]`);
        if (!cellEl) continue;
        const raw = sheet[ref];
        if (raw !== undefined && String(raw).trim().startsWith('=')) {
          try {
            cellEl.textContent = formatFormulaResult(engine.valueAt(ref));
          } catch (e) {
            cellEl.textContent = '#ERROR!';
          }
        } else {
          cellEl.textContent = raw || '';
        }
      }
    }
  }

  for (let r = 1; r <= rows; r++) {
    const tr = document.createElement('tr');
    const rowHead = document.createElement('th');
    rowHead.textContent = r;
    rowHead.className = 'row-head';
    tr.appendChild(rowHead);
    for (let c = 0; c < cols; c++) {
      const ref = COL_LETTERS_FULL[c] + r;
      const td = document.createElement('td');
      if (editable.has(ref)) {
        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'grid-cell-input mono';
        input.placeholder = ref === targetCell ? '=...' : '';
        input.value = sheet[ref] || '';
        input.addEventListener('input', () => {
          sheet[ref] = input.value;
          recompute();
        });
        inputs[ref] = input;
        td.appendChild(input);
      } else {
        td.dataset.staticRef = ref;
        td.textContent = sheet[ref] || '';
      }
      tr.appendChild(td);
    }
    table.appendChild(tr);
  }
  wrap.appendChild(table);
  recompute();

  const actions = document.createElement('div');
  actions.className = 'challenge-actions';

  const checkBtn = document.createElement('button');
  checkBtn.className = 'btn btn-primary';
  checkBtn.textContent = 'Check answer';
  actions.appendChild(checkBtn);

  if (data.hint) {
    const hintBtn = document.createElement('button');
    hintBtn.className = 'btn btn-ghost';
    hintBtn.textContent = 'Show hint';
    let shown = false;
    const hintBox = document.createElement('div');
    hintBox.className = 'challenge-hint hidden';
    hintBox.textContent = data.hint;
    hintBtn.addEventListener('click', () => {
      shown = !shown;
      hintBox.classList.toggle('hidden', !shown);
      hintBtn.textContent = shown ? 'Hide hint' : 'Show hint';
    });
    actions.appendChild(hintBtn);
    wrap.appendChild(actions);
    wrap.appendChild(hintBox);
  } else {
    wrap.appendChild(actions);
  }

  const resultBox = document.createElement('div');
  resultBox.className = 'challenge-result';
  wrap.appendChild(resultBox);

  checkBtn.addEventListener('click', () => {
    const engine = new FormulaEngine(getEngineValue);
    let actual;
    let errored = false;
    try {
      actual = engine.valueAt(targetCell);
    } catch (e) {
      errored = true;
    }
    const tolerance = data.tolerance ?? 0.01;
    const expected = data.expectedValue;
    const correct =
      !errored &&
      typeof actual === 'number' &&
      typeof expected === 'number' &&
      Math.abs(actual - expected) <= tolerance;

    state.progress.quizTotal += 1;
    if (correct) {
      state.progress.quizCorrect += 1;
      if (topicId && !state.progress.completed.includes(topicId)) {
        state.progress.completed.push(topicId);
      }
    } else {
      const topic = allTopics().find((t) => t.id === topicId);
      state.progress.weakSpots.push({
        topicId: topicId || null,
        topicName: topic ? topic.name : null,
        question: data.description || 'Practice challenge',
        ts: Date.now(),
      });
      state.progress.weakSpots = state.progress.weakSpots.slice(-25);
    }
    saveProgress();
    renderSkillTree();
    updateHeader();

    resultBox.className = 'challenge-result ' + (correct ? 'ok' : 'bad');
    resultBox.textContent = correct
      ? `✓ Correct — ${targetCell} = ${formatFormulaResult(actual)}`
      : errored
      ? `That formula has an error — check your syntax and try again.`
      : `Not quite — ${targetCell} is currently ${
          actual === undefined ? 'empty' : formatFormulaResult(actual)
        }. Keep trying, or check the hint.`;
  });

  addExportButton(wrap, () => sheet, cols, rows, 'excel-tutor-practice.xlsx');
  return wrap;
}

/* ---- Visual: chart via Chart.js ---- */
let chartCounter = 0;
function renderChart(data) {
  const wrap = document.createElement('div');
  wrap.className = 'visual-wrap chart-wrap';
  const canvas = document.createElement('canvas');
  canvas.id = 'chart-' + ++chartCounter;
  canvas.height = 200;
  wrap.appendChild(canvas);

  if (data.caption) {
    const cap = document.createElement('div');
    cap.className = 'visual-caption';
    cap.textContent = data.caption;
    wrap.appendChild(cap);
  }

  const draw = () => {
    if (typeof Chart === 'undefined') {
      wrap.querySelector('canvas').replaceWith(document.createTextNode('Chart could not load (offline?).'));
      return;
    }
    const palette = ['#1E5F4A', '#E3A72E', '#5B8A76', '#C88E1C', '#8AA898', '#B4432F'];
    new Chart(canvas.getContext('2d'), {
      type: data.type === 'pie' ? 'pie' : data.type === 'line' ? 'line' : 'bar',
      data: {
        labels: data.labels || [],
        datasets: [
          {
            label: data.label || '',
            data: data.data || [],
            backgroundColor: data.type === 'line' ? 'rgba(30,95,74,0.15)' : palette,
            borderColor: '#1E5F4A',
            borderWidth: data.type === 'line' ? 2 : 1,
            tension: 0.3,
          },
        ],
      },
      options: {
        responsive: true,
        plugins: { legend: { display: !!data.label && data.type !== 'pie' } },
        scales: data.type === 'pie' ? {} : { y: { beginAtZero: true } },
      },
    });
  };
  if (typeof Chart === 'undefined') {
    window.addEventListener('load', () => setTimeout(draw, 200));
  } else {
    setTimeout(draw, 0);
  }
  return wrap;
}

/* ---- Visual: linear flow diagram ---- */
function renderDiagram(data) {
  const wrap = document.createElement('div');
  wrap.className = 'visual-wrap';
  const flow = document.createElement('div');
  flow.className = 'diagram-flow';
  const nodes = (data.nodes || []).slice(0, 5);
  nodes.forEach((node, idx) => {
    const box = document.createElement('div');
    box.className = 'diagram-node';
    box.textContent = node.label || '';
    flow.appendChild(box);
    if (idx < nodes.length - 1) {
      const arrow = document.createElement('div');
      arrow.className = 'diagram-arrow';
      arrow.textContent = '→';
      flow.appendChild(arrow);
    }
  });
  wrap.appendChild(flow);
  if (data.caption) {
    const cap = document.createElement('div');
    cap.className = 'visual-caption';
    cap.textContent = data.caption;
    wrap.appendChild(cap);
  }
  return wrap;
}

function renderQuiz(bubbleEl, quiz, topicId) {
  const wrap = document.createElement('div');
  wrap.className = 'quiz-options';
  quiz.options.forEach((opt, idx) => {
    const btn = document.createElement('button');
    btn.className = 'quiz-opt';
    btn.textContent = opt;
    btn.addEventListener('click', () => {
      const correct = idx === quiz.correctIndex;
      [...wrap.children].forEach((c, i) => {
        c.classList.remove('correct', 'wrong');
        if (i === quiz.correctIndex) c.classList.add('correct');
      });
      if (!correct) btn.classList.add('wrong');
      [...wrap.children].forEach((c) => (c.disabled = true));

      state.progress.quizTotal += 1;
      if (correct) {
        state.progress.quizCorrect += 1;
      } else {
        const topic = allTopics().find((t) => t.id === topicId);
        state.progress.weakSpots.push({
          topicId: topicId || null,
          topicName: topic ? topic.name : null,
          question: quiz.question,
          ts: Date.now(),
        });
        state.progress.weakSpots = state.progress.weakSpots.slice(-25);
      }
      if (correct && topicId && !state.progress.completed.includes(topicId)) {
        state.progress.completed.push(topicId);
      }
      saveProgress();
      renderSkillTree();
      updateHeader();

      const note = document.createElement('div');
      note.style.marginTop = '8px';
      note.style.fontSize = '12.5px';
      note.style.color = correct ? '#1E5F4A' : '#B4432F';
      note.textContent = (correct ? '✓ Correct — ' : '✗ Not quite — ') + quiz.explanation;
      wrap.appendChild(note);
    });
    wrap.appendChild(btn);
  });
  bubbleEl.appendChild(wrap);
}

/* ---------------- Groq API call (streaming) ---------------- */
async function sendToTutor(userText, opts = {}) {
  if (!navigator.onLine) {
    addSystemNote('You appear to be offline — live tutoring needs an internet connection.');
    return;
  }
  if (opts.showUserBubble) addUserBubble(userText);
  state.chatHistory.push({ role: 'user', content: userText });
  saveChatHistory();

  const bubble = addTutorBubble();
  state.streaming = true;
  sendBtn.disabled = true;

  const messages = [
    { role: 'system', content: SYSTEM_PROMPT },
    buildLearnerContextMessage(),
    ...state.chatHistory.slice(-12),
  ];

  const endpoint = USE_BACKEND ? `${BACKEND_URL.replace(/\/$/, '')}/chat` : GROQ_URL;
  const headers = { 'Content-Type': 'application/json' };
  if (!USE_BACKEND) headers.Authorization = `Bearer ${state.apiKey}`;

  let full = '';
  try {
    const res = await fetch(endpoint, {
      method: 'POST',
      headers,
      body: JSON.stringify({
        model: GROQ_MODEL,
        messages,
        stream: true,
        temperature: 0.6,
      }),
    });

    if (!res.ok) {
      const errBody = await res.text();
      throw new Error(`Tutor service error ${res.status}: ${errBody.slice(0, 200)}`);
    }

    const reader = res.body.getReader();
    const decoder = new TextDecoder();
    let buffer = '';
    bubble.innerHTML = '';

    while (true) {
      const { done, value } = await reader.read();
      if (done) break;
      buffer += decoder.decode(value, { stream: true });
      const lines = buffer.split('\n');
      buffer = lines.pop();
      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed.startsWith('data:')) continue;
        const payload = trimmed.slice(5).trim();
        if (payload === '[DONE]') continue;
        try {
          const json = JSON.parse(payload);
          const delta = json.choices?.[0]?.delta?.content;
          if (delta) {
            full += delta;
            liveRenderDuringStream(bubble, full);
            thread.scrollTop = thread.scrollHeight;
          }
        } catch {
          /* ignore partial JSON chunk */
        }
      }
    }

    finalRenderBubble(bubble, full, opts.isTopicLesson ? state.currentTopicId : null);
    thread.scrollTop = thread.scrollHeight;
    state.chatHistory.push({ role: 'assistant', content: full });
    saveChatHistory();
  } catch (err) {
    const hint = USE_BACKEND
      ? 'The backend proxy may be unreachable — check its deployment.'
      : "Double-check your Groq API key in Settings (gear icon).";
    bubble.innerHTML = renderMarkdownLite(`Couldn't reach the tutor: ${err.message}. ${hint}`);
  } finally {
    state.streaming = false;
    sendBtn.disabled = !chatInput.value.trim();
  }
}

/* ---------------- PWA: service worker ---------------- */
if ('serviceWorker' in navigator) {
  window.addEventListener('load', () => {
    navigator.serviceWorker.register('sw.js').catch(() => {
      /* offline shell simply won't be available */
    });
  });
}

boot();
