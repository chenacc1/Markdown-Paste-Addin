// --- Service Worker for Chrome & Edge ----------------------------------------──
// Polyfill is loaded via manifest "background" → lib/browser-polyfill.js
// The actual background logic initializes after the polyfill

const BRIDGE_URL = 'http://localhost:9876';

// --- Extension lifecycle -------------------------------------------------------

self.addEventListener('activate', () => {
  // Set up context menu — remove all then create fresh (avoids duplicate errors)
  chrome.contextMenus.removeAll(() => {
    chrome.contextMenus.create({
      id: 'export-selected',
      title: '导出选中内容为Word',
      contexts: ['selection']
    });
  });

  // Configure side panel (Edge + Chrome 114+)
  try {
    if (chrome.sidePanel) {
      chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: false });
    }
  } catch (e) { /* sidePanel not available in this version */ }

  console.log('MarkdownPasteAddin v2.0 ready');
});

// --- Context menu -------------------------------------------------------------─

chrome.contextMenus.onClicked.addListener((info, tab) => {
  if (info.menuItemId === 'export-selected' && tab?.id) {
    chrome.tabs.sendMessage(tab.id, { action: 'exportSelection' });
  }
});

// --- Message router ----------------------------------------------------------──

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

// --- Action (toolbar icon) click ----------------------------------------------─

chrome.action.onClicked.addListener((tab) => {
  // Try side panel first (Edge), fall back to popup is default behavior
});

// --- Bridge communication ----------------------------------------------------──

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
    return { online: false };
  } catch {
    return { online: false };
  }
}

async function convertToDocx(content, format = 'markdown', options = {}) {
  const resp = await fetch(`${BRIDGE_URL}/api/convert`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ content, format, options })
  });

  if (!resp.ok) {
    const err = await resp.text();
    throw new Error('Server error ' + resp.status + ': ' + err);
  }

  const disposition = resp.headers.get('Content-Disposition') || '';
  const filenameMatch = disposition.match(/filename="?([^"]+)"?/);
  const filename = filenameMatch ? filenameMatch[1] : 'export.docx';

  // Service workers don't have URL.createObjectURL — convert binary to base64 in chunks
  const arrayBuffer = await resp.arrayBuffer();
  const bytes = new Uint8Array(arrayBuffer);
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
  const dataUrl = 'data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,' + base64;

  await chrome.downloads.download({ url: dataUrl, filename: filename, saveAs: true });

  return { success: true, filename: filename };
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
