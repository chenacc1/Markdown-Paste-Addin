// --- Side Panel (Edge + Chrome) ----------------------------------------------──
// Edge supports sidePanel natively since v102, Chrome since v114 via sidePanel API

const BRIDGE_URL = 'http://localhost:9876';

document.addEventListener('DOMContentLoaded', () => {
  checkHealth();

  document.getElementById('btnExportPage').addEventListener('click', exportPage);
  document.getElementById('btnExportSelection').addEventListener('click', exportSelection);
});

async function checkHealth() {
  const dot = document.getElementById('dot');
  const text = document.getElementById('statusText');
  const hint = document.getElementById('serverHint');

  try {
    const resp = await fetch(`${BRIDGE_URL}/api/health`, {
      signal: AbortSignal.timeout(2000)
    });
    if (resp.ok) {
      const data = await resp.json();
      dot.className = 'dot online';
      text.textContent = `桥接服务已连接 v${data.version}`;
      hint.classList.add('hidden');
      return;
    }
  } catch {}
  dot.className = 'dot online';
  text.textContent = 'Ready (offline mode)';
  hint.classList.add('hidden');
}

async function exportPage() {
  setStatus('正在导出...');

  const [tab] = await browser.tabs.query({ active: true, currentWindow: true });
  if (!tab?.id) return;

  try {
    const results = await browser.scripting.executeScript({
      target: { tabId: tab.id },
      func: () => {
        const main = document.querySelector('main, article, [role="main"], .content, .prose') || document.body;
        let md = '';
        function walk(el) {
          if (['SCRIPT','STYLE','NAV','FOOTER','HEADER'].includes(el.tagName)) return;
          if (el.nodeType === Node.TEXT_NODE) { const t = el.textContent.trim(); if (t) md += t + '\n'; return; }
          if (el.nodeType !== Node.ELEMENT_NODE) return;
          switch (el.tagName) {
            case 'H1': md += `# ${el.textContent.trim()}\n\n`; break;
            case 'H2': md += `## ${el.textContent.trim()}\n\n`; break;
            case 'H3': md += `### ${el.textContent.trim()}\n\n`; break;
            case 'P': md += el.textContent.trim() + '\n\n'; break;
            case 'PRE': md += '```\n' + el.textContent.trim() + '\n```\n\n'; break;
            case 'LI': md += `- ${el.textContent.trim()}\n`; break;
            default: for (const c of el.childNodes) walk(c);
          }
        }
        walk(main);
        return { content: md, format: 'markdown' };
      }
    });

    if (results?.[0]?.result) {
      await sendToBridge(results[0].result.content, results[0].result.format);
    }
  } catch (err) {
    setStatus('错误: ' + err.message);
  }
}

async function exportSelection() {
  const [tab] = await browser.tabs.query({ active: true, currentWindow: true });
  if (!tab?.id) return;

  try {
    const results = await browser.scripting.executeScript({
      target: { tabId: tab.id },
      func: () => {
        const sel = window.getSelection();
        return sel?.toString() || '';
      }
    });

    if (results?.[0]?.result) {
      await sendToBridge(results[0].result, 'text');
    } else {
      setStatus('未选中任何内容');
    }
  } catch (err) {
    setStatus('错误: ' + err.message);
  }
}

async function sendToBridge(content, format) {
  // Route through background service worker (avoids URL.createObjectURL issues)
  return new Promise((resolve) => {
    chrome.runtime.sendMessage({
      action: 'convert',
      content: content,
      format: format,
      options: {
        toc: document.getElementById('optToc')?.checked || false,
        auto_captions: document.getElementById('optCaptions')?.checked || false,
        format_after: document.getElementById('optFormat')?.checked || false,
        image_width: 5.5,
        title: document.title || ''
      }
    }, (response) => {
      if (response?.success) {
        setStatus('导出完成！');
        resolve();
      } else {
        setStatus('Conversion failed: ' + (response?.error || 'unknown error'));
        resolve();
      }
    });
  });
}

function setStatus(msg) {
  document.getElementById('statusText').textContent = msg;
}
