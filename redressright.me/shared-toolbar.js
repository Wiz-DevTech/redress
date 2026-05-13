/**
 * RedressRight Shared Toolbar
 * ─────────────────────────────────────────────────────
 * Provides two services on EVERY page:
 *   1. Text-to-Speech (for visually impaired) — start / pause / stop
 *      anywhere the user decides.
 *   2. Incremental Browser Memory — every <input>, <select>, and
 *      <textarea> value is persisted to localStorage keyed per-page,
 *      so each page independently remembers the user's progress.
 * ─────────────────────────────────────────────────────
 * Add ONE script tag to each page, just before </body>:
 *   <script src="shared-toolbar.js"></script>
 */

(function () {
  'use strict';

  /* ─── CONFIG ─────────────────────────────────────── */
  const PAGE_KEY   = 'RR_FORM_' + window.location.pathname.replace(/\W/g, '_');
  const SAVE_DELAY = 800; // ms debounce before writing to localStorage

  /* ─── INJECT TOOLBAR STYLES ─────────────────────── */
  const css = `
    #rr-toolbar {
      position: fixed;
      bottom: 28px;
      left: 20px;
      z-index: 99999;
      display: flex;
      flex-direction: column;
      align-items: flex-start;
      gap: 6px;
      font-family: 'Inter', Arial, sans-serif;
    }
    #rr-toolbar-toggle {
      width: 52px;
      height: 52px;
      border-radius: 50%;
      background: linear-gradient(135deg,#1e3c72,#2a5298);
      border: 2px solid #ffd700;
      color: #ffd700;
      font-size: 22px;
      cursor: pointer;
      display: flex;
      align-items: center;
      justify-content: center;
      box-shadow: 0 4px 18px rgba(0,0,0,.45);
      transition: transform .2s;
    }
    #rr-toolbar-toggle:hover { transform: scale(1.1); }
    #rr-panel {
      background: rgba(15,23,42,.96);
      border: 1px solid rgba(255,215,0,.3);
      border-radius: 12px;
      padding: 14px 16px;
      min-width: 240px;
      box-shadow: 0 8px 30px rgba(0,0,0,.5);
      display: none;
    }
    #rr-panel.open { display: block; }
    #rr-panel h4 {
      color: #ffd700;
      font-size: 11px;
      font-weight: 700;
      letter-spacing: 1.5px;
      text-transform: uppercase;
      margin: 0 0 10px 0;
    }
    .rr-btn-row {
      display: flex;
      gap: 6px;
      margin-bottom: 8px;
    }
    .rr-btn {
      flex: 1;
      padding: 7px 4px;
      border-radius: 6px;
      border: none;
      font-size: 11px;
      font-weight: 600;
      cursor: pointer;
      letter-spacing: .5px;
      transition: opacity .2s;
    }
    .rr-btn:disabled { opacity: .35; cursor: not-allowed; }
    .rr-btn-play  { background: #10b981; color: #fff; }
    .rr-btn-pause { background: #f59e0b; color: #000; }
    .rr-btn-stop  { background: #ef4444; color: #fff; }
    .rr-btn-read  { background: #3b82f6; color: #fff; }
    .rr-status {
      font-size: 10px;
      color: #94a3b8;
      margin-bottom: 8px;
      min-height: 14px;
    }
    .rr-divider {
      border: none;
      border-top: 1px solid rgba(255,255,255,.1);
      margin: 8px 0;
    }
    .rr-mem-row {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 8px;
    }
    .rr-mem-label {
      font-size: 10px;
      color: #94a3b8;
    }
    .rr-mem-dot {
      width: 8px;
      height: 8px;
      border-radius: 50%;
      background: #10b981;
      flex-shrink: 0;
    }
    .rr-mem-dot.empty { background: #4b5563; }
    .rr-btn-clear {
      background: transparent;
      border: 1px solid rgba(239,68,68,.5);
      color: #ef4444;
      font-size: 10px;
      padding: 3px 8px;
      border-radius: 4px;
      cursor: pointer;
    }
    .rr-btn-clear:hover { background: rgba(239,68,68,.15); }
    #rr-speed {
      width: 100%;
      margin-top: 6px;
      accent-color: #ffd700;
    }
    .rr-speed-label {
      font-size: 10px;
      color: #94a3b8;
      display: flex;
      justify-content: space-between;
    }
  `;

  const styleEl = document.createElement('style');
  styleEl.textContent = css;
  document.head.appendChild(styleEl);

  /* ─── BUILD TOOLBAR HTML ─────────────────────────── */
  const toolbar = document.createElement('div');
  toolbar.id = 'rr-toolbar';
  toolbar.innerHTML = `
    <div id="rr-panel">
      <h4>♿ Accessibility</h4>

      <div class="rr-status" id="rr-status">Ready — select text or click Read Page</div>

      <div class="rr-btn-row">
        <button class="rr-btn rr-btn-read"  id="rr-read"  title="Read entire page aloud">▶ Read Page</button>
        <button class="rr-btn rr-btn-read"  id="rr-sel"   title="Read selected text only">📌 Selection</button>
      </div>
      <div class="rr-btn-row">
        <button class="rr-btn rr-btn-pause" id="rr-pause" title="Pause / Resume" disabled>⏸ Pause</button>
        <button class="rr-btn rr-btn-stop"  id="rr-stop"  title="Stop reading"   disabled>⏹ Stop</button>
      </div>

      <div class="rr-speed-label"><span>Speed</span><span id="rr-speed-val">1×</span></div>
      <input type="range" id="rr-speed" min="0.5" max="2.5" step="0.1" value="1">

      <hr class="rr-divider">

      <h4>💾 Page Memory</h4>
      <div class="rr-mem-row">
        <div class="rr-mem-dot ${hasSavedData() ? '' : 'empty'}" id="rr-mem-dot"></div>
        <span class="rr-mem-label" id="rr-mem-label">${hasSavedData() ? 'Form data saved' : 'No saved data'}</span>
        <button class="rr-btn-clear" id="rr-clear">Clear</button>
      </div>
    </div>

    <button id="rr-toolbar-toggle" title="Accessibility &amp; Memory Tools" aria-label="Open accessibility toolbar">
      ♿
    </button>
  `;
  document.body.appendChild(toolbar);

  /* ─── PANEL TOGGLE ───────────────────────────────── */
  document.getElementById('rr-toolbar-toggle').addEventListener('click', () => {
    document.getElementById('rr-panel').classList.toggle('open');
  });

  /* ─── TEXT-TO-SPEECH ─────────────────────────────── */
  const synth = window.speechSynthesis;
  let utterance = null;
  let isPaused   = false;

  function setStatus(msg) {
    document.getElementById('rr-status').textContent = msg;
  }

  function setControls(playing) {
    document.getElementById('rr-pause').disabled = !playing;
    document.getElementById('rr-stop').disabled  = !playing;
  }

  function getPageText() {
    // Collect visible text nodes, skipping scripts/styles/toolbar itself
    const walker = document.createTreeWalker(
      document.body,
      NodeFilter.SHOW_TEXT,
      {
        acceptNode(node) {
          const p = node.parentElement;
          if (!p) return NodeFilter.FILTER_REJECT;
          if (['SCRIPT','STYLE','NOSCRIPT'].includes(p.tagName)) return NodeFilter.FILTER_REJECT;
          if (p.closest('#rr-toolbar')) return NodeFilter.FILTER_REJECT;
          if (getComputedStyle(p).display === 'none') return NodeFilter.FILTER_REJECT;
          const txt = node.textContent.trim();
          return txt ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP;
        }
      }
    );
    const parts = [];
    let node;
    while ((node = walker.nextNode())) {
      parts.push(node.textContent.trim());
    }
    return parts.join('. ');
  }

  function speak(text) {
    if (!synth) { setStatus('Speech not supported in this browser.'); return; }
    synth.cancel();
    isPaused = false;
    if (!text || !text.trim()) { setStatus('No text found to read.'); return; }

    utterance = new SpeechSynthesisUtterance(text);
    utterance.rate  = parseFloat(document.getElementById('rr-speed').value) || 1;
    utterance.pitch = 1;

    utterance.onstart = () => {
      setStatus('Reading…');
      setControls(true);
      document.getElementById('rr-pause').textContent = '⏸ Pause';
    };
    utterance.onpause = () => {
      setStatus('Paused — click Resume to continue.');
      document.getElementById('rr-pause').textContent = '▶ Resume';
    };
    utterance.onresume = () => {
      setStatus('Reading…');
      document.getElementById('rr-pause').textContent = '⏸ Pause';
    };
    utterance.onend = () => {
      setStatus('Finished.');
      setControls(false);
      utterance = null;
      isPaused = false;
    };
    utterance.onerror = (e) => {
      if (e.error !== 'interrupted') setStatus('Error: ' + e.error);
      setControls(false);
    };

    synth.speak(utterance);
  }

  document.getElementById('rr-read').addEventListener('click', () => {
    speak(getPageText());
  });

  document.getElementById('rr-sel').addEventListener('click', () => {
    const sel = window.getSelection ? window.getSelection().toString().trim() : '';
    if (sel) {
      speak(sel);
    } else {
      setStatus('Highlight text on the page first, then click Selection.');
    }
  });

  document.getElementById('rr-pause').addEventListener('click', () => {
    if (!synth) return;
    if (isPaused) {
      synth.resume();
      isPaused = false;
    } else {
      synth.pause();
      isPaused = true;
    }
  });

  document.getElementById('rr-stop').addEventListener('click', () => {
    if (synth) { synth.cancel(); }
    utterance = null;
    isPaused  = false;
    setStatus('Stopped.');
    setControls(false);
    document.getElementById('rr-pause').textContent = '⏸ Pause';
  });

  document.getElementById('rr-speed').addEventListener('input', function () {
    document.getElementById('rr-speed-val').textContent = this.value + '×';
    if (utterance) { utterance.rate = parseFloat(this.value); }
  });

  /* ─── BROWSER MEMORY — FORM AUTO-SAVE ───────────── */
  function hasSavedData() {
    try { return !!localStorage.getItem(PAGE_KEY); } catch (e) { return false; }
  }

  function updateMemIndicator() {
    const dot   = document.getElementById('rr-mem-dot');
    const label = document.getElementById('rr-mem-label');
    if (hasSavedData()) {
      dot.classList.remove('empty');
      label.textContent = 'Form data saved';
    } else {
      dot.classList.add('empty');
      label.textContent = 'No saved data';
    }
  }

  function collectFormState() {
    const state = {};
    document.querySelectorAll('input:not([type="password"]):not([type="file"]), select, textarea').forEach(el => {
      if (el.closest('#rr-toolbar')) return; // skip toolbar's own range input
      const key = el.id || el.name;
      if (!key) return;
      if (el.type === 'checkbox' || el.type === 'radio') {
        state[key] = el.checked;
      } else {
        state[key] = el.value;
      }
    });
    return state;
  }

  function applyFormState(state) {
    if (!state) return;
    Object.entries(state).forEach(([key, val]) => {
      const el = document.getElementById(key) || document.querySelector(`[name="${key}"]`);
      if (!el) return;
      if (el.type === 'checkbox' || el.type === 'radio') {
        el.checked = !!val;
      } else {
        el.value = val;
        // Fire change event so frameworks / GAAP app re-renders
        el.dispatchEvent(new Event('change', { bubbles: true }));
      }
    });
  }

  let saveTimer = null;
  function scheduleSave() {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => {
      try {
        const state = collectFormState();
        if (Object.keys(state).length > 0) {
          localStorage.setItem(PAGE_KEY, JSON.stringify(state));
          updateMemIndicator();
        }
      } catch (e) { /* storage full — silent */ }
    }, SAVE_DELAY);
  }

  // Restore on load (after short delay to let frameworks hydrate)
  window.addEventListener('load', () => {
    setTimeout(() => {
      try {
        const raw = localStorage.getItem(PAGE_KEY);
        if (raw) {
          applyFormState(JSON.parse(raw));
          setStatus('Form data restored from memory.');
          setTimeout(() => setStatus('Ready'), 3000);
        }
      } catch (e) { /* ignore parse errors */ }
      updateMemIndicator();
    }, 600);
  });

  // Listen for any input change on the page
  document.addEventListener('input',  scheduleSave, true);
  document.addEventListener('change', scheduleSave, true);

  // Clear memory button
  document.getElementById('rr-clear').addEventListener('click', () => {
    try { localStorage.removeItem(PAGE_KEY); } catch (e) {}
    updateMemIndicator();
    setStatus('Page memory cleared.');
    setTimeout(() => setStatus('Ready'), 2000);
  });

  // Also save before leaving the page
  window.addEventListener('beforeunload', () => {
    clearTimeout(saveTimer);
    try {
      const state = collectFormState();
      if (Object.keys(state).length > 0) {
        localStorage.setItem(PAGE_KEY, JSON.stringify(state));
      }
    } catch (e) {}
  });

})();
