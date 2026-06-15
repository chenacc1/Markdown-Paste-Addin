// MarkdownPasteAddin v3.2 — Service Worker
// Offline-first: generates .docx locally when bridge server is unavailable.

importScripts('lib/docx-builder.js');

const BRIDGE_URL = 'http://localhost:9876';

// --- Extension lifecycle -------------------------------------------------------
self.addEventListener('activate', () => {
  chrome.contextMenus.removeAll(() => {
    chrome.contextMenus.create({
      id: 'export-selected',
      title: 'Export selected as Word',
      contexts: ['selection']
    });
  });
  try {
    if (chrome.sidePanel) {
      chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: false });
    }
  } catch (e) { /* sidePanel not available */ }
  console.log('MarkdownPasteAddin v3.2 ready');
});

// --- Context menu --------------------------------------------------------------
chrome.contextMenus.onClicked.addListener((info, tab) => {
  if (info.menuItemId === 'export-selected' && tab?.id) {
    chrome.tabs.sendMessage(tab.id, { action: 'exportSelection' });
  }
});

// --- Message router ------------------------------------------------------------
chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  switch (message.action) {
    case 'checkHealth':
      checkBridgeHealth().then(sendResponse);
      return true;

    case 'convert':
      convertToDocx(message.content, message.format, message.options)
        .then(sendResponse)
        .catch(err => sendResponse({ success: false, error: err.message }));
      return true;

    case 'openSidePanel':
      openSidePanel(sender.tab?.windowId).then(sendResponse);
      return true;
  }
});

// --- Bridge communication ------------------------------------------------------
async function checkBridgeHealth() {
  try {
    const resp = await fetch(`${BRIDGE_URL}/api/health`, {
      method: 'GET',
      signal: AbortSignal.timeout(2000)
    });
    if (resp.ok) {
      const data = await resp.json();
      return { online: true, version: data.version || 'unknown' };
    }
    return { online: false, local: true };
  } catch {
    return { online: false, local: true };
  }
}

async function convertToDocx(content, format = 'markdown', options = {}) {
  // Step 1: Try bridge server for full-featured conversion
  try {
    const resp = await fetch(`${BRIDGE_URL}/api/convert`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ content, format, options }),
      signal: AbortSignal.timeout(15000)
    });

    if (resp.ok) {
      const disposition = resp.headers.get('Content-Disposition') || '';
      const filenameMatch = disposition.match(/filename="?([^"]+)"?/);
      const filename = filenameMatch ? filenameMatch[1] : 'export.docx';

      const arrayBuffer = await resp.arrayBuffer();
      const bytes = new Uint8Array(arrayBuffer);
      const base64 = arrayBufferToBase64(bytes);
      const dataUrl = 'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' + base64;

      await chrome.downloads.download({ url: dataUrl, filename: filename, saveAs: true });
      return { success: true, filename: filename, via: 'bridge' };
    }
  } catch {
    // Bridge unavailable — fall through to local generation
  }

  // Step 2: Bridge offline — generate .docx locally in the browser
  return convertToDocxLocal(content, format, options);
}

// --- Local (offline) conversion -------------------------------------------------
function convertToDocxLocal(content, format, options) {
  try {
    // Convert HTML to plain text / markdown if needed
    let mdContent = content;
    if (format === 'html') {
      // Basic HTML-to-text: strip tags, preserve structure
      const temp = document.createElement('div');
      temp.innerHTML = content;
      mdContent = htmlToMarkdownApprox(temp);
    }

    // Markdown is already parsed as markdown
    const docxBytes = DocxBuilder.buildDocx(mdContent, options);
    const base64 = arrayBufferToBase64(docxBytes);
    const title = (options.title || 'export').replace(/[<>:"/\\|?*]/g, '_').substring(0, 80) || 'export';
    const dataUrl = 'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' + base64;

    return chrome.downloads.download({
      url: dataUrl,
      filename: title + '.docx',
      saveAs: true
    }).then(() => ({ success: true, filename: title + '.docx', via: 'local' }));
  } catch (e) {
    throw new Error('Local conversion failed: ' + e.message);
  }
}

function htmlToMarkdownApprox(root) {
  let md = '';
  function walk(el, depth) {
    if (depth > 30 || !el) return;
    if (el.nodeType === Node.TEXT_NODE) {
      const t = el.textContent.trim();
      if (t) md += t + '\n';
      return;
    }
    if (el.nodeType !== Node.ELEMENT_NODE) return;
    const tag = el.tagName;
    if (['SCRIPT', 'STYLE', 'NAV', 'FOOTER', 'HEADER'].includes(tag)) return;
    switch (tag) {
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
      default: for (const c of el.childNodes) walk(c, depth + 1);
    }
  }
  walk(root, 0);
  return md;
}

// --- Utilities -----------------------------------------------------------------
function arrayBufferToBase64(bytes) {
  const chunkSize = 8192;
  let base64 = '';
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.subarray(i, Math.min(i + chunkSize, bytes.length));
    let binary = '';
    for (let j = 0; j < chunk.length; j++) {
      binary += String.fromCharCode(chunk[j]);
    }
    base64 += btoa(binary);
  }
  return base64;
}

async function openSidePanel(windowId) {
  try {
    if (chrome.sidePanel) {
      await chrome.sidePanel.open({ windowId: windowId || -1 });
      return { success: true };
    }
  } catch (e) {
    return { success: false, error: 'sidePanel not supported' };
  }
}
