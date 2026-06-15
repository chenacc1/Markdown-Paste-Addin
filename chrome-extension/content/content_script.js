// --- Content Script: injected into every page ---------------------------------─
// Detects structured content, adds floating export button, handles selections

(function () {
  'use strict';

  // --- Extension context safety ------------------------------------------------
  function isExtensionValid() {
    try {
      return !!(typeof chrome !== 'undefined' && chrome && chrome.runtime && chrome.runtime.id);
    } catch(e) { return false; }
  }

  function safeSendMessage(msg, callback) {
    if (!isExtensionValid()) {
      console.debug('Extension context invalidated, skipping message:', msg.action);
      return;
    }
    try {
      chrome.runtime.sendMessage(msg, function(response) {
        if (chrome.runtime.lastError) {
          console.debug('Runtime error:', chrome.runtime.lastError.message);
          return;
        }
        if (callback) callback(response);
      });
    } catch(e) {
      console.debug('sendMessage failed:', e.message);
    }
  }

  // --- Floating button ----------------------------------------------------------

  let floatingBtn = null;

  function createFloatingButton() {
    if (floatingBtn) return;

    floatingBtn = document.createElement('div');
    floatingBtn.id = 'mdpaste-float-btn';
    floatingBtn.innerHTML = `
      <button id="mdpaste-export-btn" title="导出为Word">
        <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="white" stroke-width="2">
          <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
          <polyline points="14 2 14 8 20 8"/>
          <line x1="12" y1="18" x2="12" y2="12"/>
          <polyline points="9 15 12 18 15 15"/>
        </svg>
      </button>
      <span id="mdpaste-status"></span>
    `;

    document.body.appendChild(floatingBtn);

    document.getElementById('mdpaste-export-btn').addEventListener('click', onExportClick);

    // Check bridge health (only if extension context is valid)
    if (isExtensionValid()) checkHealth();
  }

  function onExportClick() {
    const content = capturePageContent();
    if (!content) {
      showToast('未检测到可导出的内容');
      return;
    }

    // Send to background for conversion
    safeSendMessage({
      action: 'convert',
      content: content.text,
      format: content.format,
      options: {
        title: document.title || '',
        toc: false,
        format_after: true,
        auto_captions: true,
        image_width: 5.5
      }
    }, (response) => {
      if (response?.success) {
        showToast('导出完成: ' + (response.filename || 'export.docx'));
      } else {
        showToast('转换失败: ' + (response?.error || '桥接服务未启动'));
      }
    });
  }

  function checkHealth() {
    safeSendMessage({ action: 'checkHealth' }, (response) => {
      const statusEl = document.getElementById('mdpaste-status');
      if (statusEl) {
        if (response?.online) {
          statusEl.className = 'mdpaste-online';
          statusEl.title = 'Bridge connected v' + (response.version || '?') + ' | Full features';
        } else {
          statusEl.className = 'mdpaste-online';
          statusEl.title = 'Offline mode | Click to export as Word';
        }
      }
    });
  }

  // Periodic health check (stops when extension context is invalidated)
  setInterval(function() { if (isExtensionValid()) checkHealth(); }, 30000);

  // --- Content capture ----------------------------------------------------------

  function capturePageContent() {
    // Strategy 1: Find AI conversation containers
    const conversations = findConversations();
    if (conversations.length > 0) {
      return { text: conversations, format: 'markdown' };
    }

    // Strategy 2: Find main content area with Markdown elements
    const markdownContent = findMarkdownContent();
    if (markdownContent) {
      return { text: markdownContent, format: 'markdown' };
    }

    // Strategy 3: Get article/main body
    const article = document.querySelector('article, main, [role="main"], .content, .post-content');
    if (article) {
      return { text: htmlToMarkdown(article), format: 'markdown' };
    }

    // Strategy 4: Whole body
    if (document.body && document.body.innerText && document.body.innerText.trim().length > 100) {
      return { text: htmlToMarkdown(document.body), format: 'markdown' };
    }

    return null;
  }

  function findConversations() {
    // Generic detection: find Q&A patterns in the DOM
    // Works with DeepSeek, ChatGPT, Claude, etc.
    let markdown = '';

    // Try common AI chat container selectors
    const selectors = [
      '.conversation', '.chat-container', '.messages-container',
      '[class*="conversation"]', '[class*="chat-message"]',
      '[class*="message-bubble"]', '[class*="assistant"]',
      '.markdown', '.prose', '[data-message-author-role]',
      '.ds-markdown', // DeepSeek
      '.agent-turn',  // various
    ];

    for (const sel of selectors) {
      const elements = document.querySelectorAll(sel);
      if (elements.length > 0) {
        elements.forEach((el, idx) => {
          const text = el.innerText?.trim();
          if (text && text.length > 20) {
            // Detect if it's a user or assistant message
            const isAssistant = el.classList.contains('assistant') ||
              el.getAttribute('data-message-author-role') === 'assistant' ||
              el.closest('[class*="assistant"]') !== null;

            if (isAssistant) {
              markdown += `\n\n## AI 回答 ${idx + 1}\n\n${text}\n`;
            } else {
              markdown += `\n\n## 问题 ${idx + 1}\n\n${text}\n`;
            }
            markdown += '\n---\n';
          }
        });
        if (markdown) break;
      }
    }

    return markdown || null;
  }

  function findMarkdownContent() {
    // Find code blocks, tables, mermaid diagrams in the page
    const codeBlocks = document.querySelectorAll('pre code, .highlight, [class*="language-"]');
    const tables = document.querySelectorAll('table:not([class*="nav"])');
    const mermaidBlocks = document.querySelectorAll('[class*="mermaid"], code.language-mermaid');

    if (codeBlocks.length > 0 || tables.length > 2 || mermaidBlocks.length > 0) {
      // Extract from main area
      const main = document.querySelector('main, article, .main-content, #content') || document.body;
      return htmlToMarkdown(main);
    }
    return null;
  }

  function htmlToMarkdown(element) {
    // Clone to avoid modifying the DOM
    const clone = element.cloneNode(true);

    // Remove nav, scripts, styles, hidden elements
    clone.querySelectorAll('script, style, nav, footer, header, .nav, .sidebar, .hidden, [aria-hidden="true"]').forEach(el => el.remove());

    // Convert to markdown using basic rules
    return elementToMarkdown(clone);
  }

  function elementToMarkdown(el) {
    let md = '';
    const tag = (el.tagName || '').toLowerCase();

    for (const child of el.childNodes) {
      if (child.nodeType === Node.TEXT_NODE) {
        const text = child.textContent.trim();
        if (text) md += text + '\n';
        continue;
      }
      if (child.nodeType !== Node.ELEMENT_NODE) continue;

      const childTag = (child.tagName || '').toLowerCase();

      switch (childTag) {
        case 'h1': md += `# ${child.textContent.trim()}\n\n`; break;
        case 'h2': md += `## ${child.textContent.trim()}\n\n`; break;
        case 'h3': md += `### ${child.textContent.trim()}\n\n`; break;
        case 'h4': md += `#### ${child.textContent.trim()}\n\n`; break;
        case 'h5':
        case 'h6': md += `##### ${child.textContent.trim()}\n\n`; break;

        case 'p':
          md += child.textContent.trim() + '\n\n';
          break;

        case 'pre': {
          const code = child.querySelector('code');
          const lang = code ? (code.className.match(/language-(\w+)/)?.[1] || '') : '';
          const text = code ? code.textContent : child.textContent;
          md += `\`\`\`${lang}\n${text.trim()}\n\`\`\`\n\n`;
          break;
        }

        case 'table':
          md += convertTable(child) + '\n\n';
          break;

        case 'ul':
          child.querySelectorAll('li').forEach(li => {
            md += `- ${li.textContent.trim()}\n`;
          });
          md += '\n';
          break;

        case 'ol':
          child.querySelectorAll('li').forEach((li, i) => {
            md += `${i + 1}. ${li.textContent.trim()}\n`;
          });
          md += '\n';
          break;

        case 'blockquote':
          child.textContent.trim().split('\n').forEach(line => {
            md += `> ${line.trim()}\n`;
          });
          md += '\n';
          break;

        case 'hr':
          md += '---\n\n';
          break;

        case 'img': {
          const src = child.src || child.getAttribute('data-src') || '';
          const alt = child.alt || '';
          if (src) md += `![${alt}](${src})\n\n`;
          break;
        }

        case 'a': {
          const href = child.href || '';
          const text = child.textContent.trim();
          if (href && href !== text) md += `[${text}](${href})`;
          else md += text;
          break;
        }

        case 'br':
          md += '\n';
          break;

        case 'strong':
        case 'b':
          md += `**${child.textContent.trim()}**`;
          break;

        case 'em':
        case 'i':
          md += `*${child.textContent.trim()}*`;
          break;

        case 'code':
          md += `\`${child.textContent.trim()}\``;
          break;

        case 'div':
        case 'section':
        case 'article':
        case 'span':
          md += elementToMarkdown(child);
          break;

        default:
          md += child.textContent.trim() + '\n';
          break;
      }
    }

    return md;
  }

  function convertTable(table) {
    const rows = table.querySelectorAll('tr');
    if (rows.length === 0) return '';

    let md = '';
    const allRows = [];

    rows.forEach((row, idx) => {
      const cells = [];
      row.querySelectorAll('th, td').forEach(cell => {
        cells.push(cell.textContent.trim().replace(/\|/g, '\\|'));
      });
      if (cells.length > 0) allRows.push(cells);
    });

    if (allRows.length === 0) return '';

    const colCount = Math.max(...allRows.map(r => r.length));
    allRows.forEach(r => {
      while (r.length < colCount) r.push('');
    });

    // Header
    md += '| ' + allRows[0].join(' | ') + ' |\n';
    // Separator
    md += '| ' + Array(colCount).fill('---').join(' | ') + ' |\n';
    // Body
    for (let i = 1; i < allRows.length; i++) {
      md += '| ' + allRows[i].join(' | ') + ' |\n';
    }

    return md;
  }

  // --- Export selection handler -------------------------------------------------

  chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
    if (message.action === 'exportSelection') {
      const selection = window.getSelection();
      if (selection && selection.rangeCount > 0 && !selection.isCollapsed) {
        const container = document.createElement('div');
        for (let i = 0; i < selection.rangeCount; i++) {
          container.appendChild(selection.getRangeAt(i).cloneContents());
        }
        const markdown = htmlToMarkdown(container);

        safeSendMessage({
          action: 'convert',
          content: markdown,
          format: 'markdown',
          options: {
            title: document.title || '',
            toc: false,
            format_after: true,
            auto_captions: true,
            image_width: 5.5
          }
        });
      }
    }
    return false;
  });

  // --- Toast notification -------------------------------------------------------

  function showToast(message) {
    let toast = document.getElementById('mdpaste-toast');
    if (!toast) {
      toast = document.createElement('div');
      toast.id = 'mdpaste-toast';
      document.body.appendChild(toast);
    }
    toast.textContent = message;
    toast.className = 'mdpaste-toast-show';

    clearTimeout(toast._timeout);
    toast._timeout = setTimeout(() => {
      toast.className = 'mdpaste-toast-hide';
    }, 3000);
  }

  // --- Initialize -------------------------------------------------------------──

  function init() {
    if (!isExtensionValid()) {
      setTimeout(function() {
        if (isExtensionValid()) createFloatingButton();
      }, 1000);
      return;
    }
    createFloatingButton();
  }

  if (document.readyState === 'complete') {
    init();
  } else {
    window.addEventListener('load', init);
  }

})();
