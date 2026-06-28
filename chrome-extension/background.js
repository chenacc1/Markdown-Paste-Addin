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
  const MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';

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
      return { success: true, base64, filename, via: 'bridge', mimeType: MIME };
    }
  } catch {
    // Bridge unavailable — fall through to local generation
  }

  // Step 2: Bridge offline — generate .docx locally in the browser
  return convertToDocxLocal(content, format, options);
}

// --- Local (offline) conversion -------------------------------------------------
function convertToDocxLocal(content, format, options) {
  const MIME = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  try {
    let mdContent = content;
    if (format === 'html') {
      mdContent = stripHtml(content);
    }

    const docxBytes = DocxBuilder.buildDocx(mdContent, options);
    const base64 = arrayBufferToBase64(docxBytes);
    const title = (options.title || 'export').replace(/[<>:"/\\|?*]/g, '_').substring(0, 80) || 'export';
    return { success: true, base64, filename: title + '.docx', via: 'local', mimeType: MIME };
  } catch (e) {
    throw new Error('Local conversion failed: ' + e.message);
  }
}

function stripHtml(html) {
  // DOM-free HTML-to-text: remove tags, decode entities, preserve line breaks
  return html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n\n')
    .replace(/<\/h[1-6]>/gi, '\n\n')
    .replace(/<\/li>/gi, '\n')
    .replace(/<\/tr>/gi, '\n')
    .replace(/<\/div>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, ' ')
    .split('\n').map(s => s.trim()).join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
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
