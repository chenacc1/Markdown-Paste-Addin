// MarkdownPasteAddin v3.0 Popup
// Chrome & Edge compatible. Uses browser-polyfill.js for cross-browser support.

const BRIDGE_URL = 'http://localhost:9876';

document.addEventListener('DOMContentLoaded', () => {
  loadSettings();
  checkHealth();

  document.getElementById('btnPasteExport').addEventListener('click', pasteAndExport);
  document.getElementById('btnExportPage').addEventListener('click', exportPage);
  document.getElementById('btnExportSelection').addEventListener('click', exportSelection);
  document.getElementById('linkOptions').addEventListener('click', () => chrome.runtime.openOptionsPage());

  // Save options on change
  ['optToc', 'optCaptions', 'optFormat', 'optCover'].forEach(id => {
    document.getElementById(id).addEventListener('change', saveSettings);
  });
  document.getElementById('presetSelect').addEventListener('change', saveSettings);
});

// --- Settings ----------------------------------------------------------------──

function getOptions() {
  return {
    preset: document.getElementById('presetSelect').value || 'business',
    toc: document.getElementById('optToc').checked,
    auto_captions: document.getElementById('optCaptions').checked,
    format_after: document.getElementById('optFormat').checked,
    add_cover: document.getElementById('optCover').checked,
    image_width: 5.5,
  };
}

function saveSettings() {
  chrome.storage.local.set(getOptions());
}

function loadSettings() {
  chrome.storage.local.get({
    preset: 'business', toc: false, auto_captions: true,
    format_after: true, add_cover: false, image_width: 5.5
  }, (items) => {
    document.getElementById('optToc').checked = items.toc;
    document.getElementById('optCaptions').checked = items.auto_captions;
    document.getElementById('optFormat').checked = items.format_after;
    document.getElementById('optCover').checked = items.add_cover;
    document.getElementById('presetSelect').value = items.preset;
  });
}

// --- Health -------------------------------------------------------------------─

async function checkHealth() {
  const dot = document.getElementById('statusDot');
  const text = document.getElementById('statusText');
  const hint = document.getElementById('serverHint');

  try {
    const resp = await fetch(`${BRIDGE_URL}/api/health`, { signal: AbortSignal.timeout(2000) });
    if (resp.ok) {
      const data = await resp.json();
      dot.className = 'status-dot online';
      text.textContent = `Connected · v${data.version}`;
      hint.classList.add('hidden');
      return;
    }
  } catch {}
  dot.className = 'status-dot offline';
  text.textContent = 'Bridge offline';
  hint.classList.remove('hidden');
}

// --- Paste & Export (reads clipboard from page) --------------------------------

async function pasteAndExport() {
  setStatus('converting', 'Reading clipboard...');

  try {
    // Read clipboard directly from popup context (requires clipboardRead permission)
    const text = await navigator.clipboard.readText();

    if (text && text.trim()) {
      try {
        await sendToBridge(text.trim(), 'text');
      } catch (e) {
        setStatus('error', 'Bridge offline? ' + e.message);
      }
    } else {
      setStatus('error', 'Clipboard empty. Copy from DeepSeek first (Ctrl+C), then click this button.');
    }
  } catch (err) {
    setStatus('error', 'Cannot read clipboard. Try: right-click extension icon → Manage extension → enable clipboard permission.');
  }
}

// --- Export -------------------------------------------------------------------─

async function exportPage() {
  setStatus('converting', 'Exporting page...');
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.id) return;

  try {
    const results = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: () => {
        const main = document.querySelector('main, article, [role="main"], .prose, .content') || document.body;
        let md = '';
        function walk(el, d) {
          if (d > 25 || !el) return;
          if (['SCRIPT','STYLE','NAV','FOOTER','HEADER'].includes(el.tagName)) return;
          if (el.nodeType === Node.TEXT_NODE) { const t = el.textContent.trim(); if (t) md += t + '\n'; return; }
          if (el.nodeType !== Node.ELEMENT_NODE) return;
          switch (el.tagName) {
            case 'H1': md += `# ${el.textContent.trim()}\n\n`; break;
            case 'H2': md += `## ${el.textContent.trim()}\n\n`; break;
            case 'H3': md += `### ${el.textContent.trim()}\n\n`; break;
            case 'H4': md += `#### ${el.textContent.trim()}\n\n`; break;
            case 'P': md += el.textContent.trim() + '\n\n'; break;
            case 'PRE': md += '```\n' + el.textContent.trim() + '\n```\n\n'; break;
            case 'LI': md += `- ${el.textContent.trim()}\n`; break;
            case 'HR': md += '---\n\n'; break;
            case 'TABLE': {
              const rows = el.querySelectorAll('tr');
              if (rows.length) {
                const hdr = rows[0].querySelectorAll('th,td');
                md += '| ' + [...hdr].map(c => c.textContent.trim()).join(' | ') + ' |\n';
                md += '| ' + [...hdr].map(() => '---').join(' | ') + ' |\n';
                for (let i = 1; i < rows.length; i++) {
                  md += '| ' + [...rows[i].querySelectorAll('td,th')].map(c => c.textContent.trim()).join(' | ') + ' |\n';
                }
                md += '\n';
              }
              break;
            }
            default: for (const c of el.childNodes) walk(c, d + 1);
          }
        }
        walk(main, 0);
        return { content: md, format: 'markdown' };
      }
    });

    if (results?.[0]?.result) {
      await sendToBridge(results[0].result.content, results[0].result.format);
    } else {
      setStatus('error', 'No content detected');
    }
  } catch (err) {
    setStatus('error', err.message);
  }
}

async function exportSelection() {
  setStatus('converting', 'Exporting selection...');
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  if (!tab?.id) return;

  try {
    const results = await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: () => {
        const sel = window.getSelection();
        if (!sel || sel.isCollapsed) return null;
        const div = document.createElement('div');
        for (let i = 0; i < sel.rangeCount; i++)
          div.appendChild(sel.getRangeAt(i).cloneContents());
        return { content: div.innerText?.trim() || sel.toString().trim(), format: 'text' };
      }
    });

    if (results?.[0]?.result) {
      await sendToBridge(results[0].result.content, results[0].result.format);
    } else {
      setStatus('error', 'No text selected');
    }
  } catch (err) {
    setStatus('error', err.message);
  }
}

async function sendToBridge(content, format) {
  const options = getOptions();
  const resp = await fetch(`${BRIDGE_URL}/api/convert`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ content, format, options })
  });

  if (!resp.ok) {
    const err = await resp.text();
    throw new Error('Server ' + resp.status + ': ' + err.substring(0, 200));
  }

  // Convert blob to base64 data URL, then use downloads API with saveAs dialog
  const blob = await resp.blob();
  const reader = new FileReader();
  const dataUrl = await new Promise((resolve, reject) => {
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });

  await chrome.downloads.download({
    url: dataUrl,
    filename: 'export.docx',
    saveAs: true
  });

  setStatus('done', 'Choose save location...');
  setTimeout(() => window.close(), 2000);
}

async function openSidePanel() {
  try {
    await chrome.runtime.sendMessage({ action: 'openSidePanel' });
  } catch {
    // Fallback: open popup itself is fine
  }
}

// --- Status -------------------------------------------------------------------─

function setStatus(state, msg) {
  const dot = document.getElementById('statusDot');
  const text = document.getElementById('statusText');
  text.textContent = msg;
  switch (state) {
    case 'converting': dot.className = 'status-dot online'; break;
    case 'done':       dot.className = 'status-dot online'; break;
    case 'error':      dot.className = 'status-dot offline'; break;
    default:           dot.className = 'status-dot'; break;
  }
}
