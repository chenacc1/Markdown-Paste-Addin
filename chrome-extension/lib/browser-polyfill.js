// --- Cross-browser polyfill ----------------------------------------------------
// Normalizes chrome.* (Chrome/Edge/Opera) and browser.* (Firefox) APIs.
//
// Edge supports chrome.* natively since it's Chromium-based.
// Firefox needs browser.* + promises.
// This polyfill ensures both work, with Edge having zero runtime cost.

(function (global) {
  'use strict';

  // If browser is already defined, no polyfill needed (Firefox)
  if (typeof global.browser !== 'undefined') return;

  const hasChrome = typeof global.chrome !== 'undefined';
  if (!hasChrome) return;

  // Edge detection (for analytics only, not behavior)
  const isEdge = navigator.userAgent.includes('Edg/');
  const isChrome = navigator.userAgent.includes('Chrome/') && !isEdge;
  const isFirefox = navigator.userAgent.includes('Firefox/');

  // Store browser info for conditional logic
  global.__BROWSER__ = {
    edge: isEdge,
    chrome: isChrome,
    firefox: isFirefox,
    name: isEdge ? 'edge' : isChrome ? 'chrome' : isFirefox ? 'firefox' : 'chromium'
  };

  // ── API wrapping -------------------------------------------------------------─

  const wrap = {
    runtime: {
      sendMessage(extensionId, message, options) {
        // Normalize arguments (chrome.runtime.sendMessage has multiple signatures)
        if (typeof extensionId === 'string') {
          return chrome.runtime.sendMessage(extensionId, message, options);
        }
        return chrome.runtime.sendMessage(extensionId);
      },

      sendNativeMessage(application, message) {
        return chrome.runtime.sendNativeMessage(application, message);
      },

      getManifest() {
        return chrome.runtime.getManifest();
      },

      getURL(path) {
        return chrome.runtime.getURL(path);
      },

      openOptionsPage() {
        return chrome.runtime.openOptionsPage();
      },

      onMessage: chrome.runtime.onMessage,
      onInstalled: chrome.runtime.onInstalled,
      onConnect: chrome.runtime.onConnect,
      lastError: null,
    },

    tabs: {
      query(queryInfo) {
        return new Promise((resolve) => {
          chrome.tabs.query(queryInfo, resolve);
        });
      },

      create(createProperties) {
        return new Promise((resolve) => {
          chrome.tabs.create(createProperties, resolve);
        });
      },

      sendMessage(tabId, message) {
        return new Promise((resolve, reject) => {
          chrome.tabs.sendMessage(tabId, message, (response) => {
            if (chrome.runtime.lastError) {
              reject(chrome.runtime.lastError);
            } else {
              resolve(response);
            }
          });
        });
      },

      executeScript(tabId, details) {
        return chrome.scripting.executeScript({
          target: { tabId: tabId },
          ...details,
        });
      },
    },

    storage: {
      local: {
        get(keys) {
          return new Promise((resolve) => {
            chrome.storage.local.get(keys, resolve);
          });
        },
        set(items) {
          return new Promise((resolve) => {
            chrome.storage.local.set(items, resolve);
          });
        },
        remove(keys) {
          return new Promise((resolve) => {
            chrome.storage.local.remove(keys, resolve);
          });
        },
        clear() {
          return new Promise((resolve) => {
            chrome.storage.local.clear(resolve);
          });
        },
        onChanged: chrome.storage.local.onChanged,
      },
      sync: {
        get(keys) {
          return new Promise((resolve) => {
            chrome.storage.sync.get(keys, resolve);
          });
        },
        set(items) {
          return new Promise((resolve) => {
            chrome.storage.sync.set(items, resolve);
          });
        },
      },
    },

    scripting: {
      executeScript(injection) {
        return new Promise((resolve, reject) => {
          if (typeof chrome.scripting !== 'undefined') {
            chrome.scripting.executeScript(injection, (results) => {
              if (chrome.runtime.lastError) {
                reject(chrome.runtime.lastError);
              } else {
                resolve(results);
              }
            });
          } else {
            // Fallback for older browsers
            chrome.tabs.executeScript(
              injection.target.tabId,
              {
                file: injection.files?.[0],
                code: injection.func ? `(${injection.func.toString()})()` : undefined,
              },
              (results) => {
                if (chrome.runtime.lastError) {
                  reject(chrome.runtime.lastError);
                } else {
                  resolve([{ result: results?.[0] }]);
                }
              }
            );
          }
        });
      },
    },

    contextMenus: {
      create(createProperties) {
        return chrome.contextMenus.create(createProperties);
      },
      onClicked: chrome.contextMenus.onClicked,
    },

    downloads: {
      download(options) {
        return new Promise((resolve, reject) => {
          chrome.downloads.download(options, (downloadId) => {
            if (chrome.runtime.lastError) {
              reject(chrome.runtime.lastError);
            } else {
              resolve(downloadId);
            }
          });
        });
      },
    },

    sidePanel: {
      open(options) {
        if (typeof chrome.sidePanel !== 'undefined') {
          return chrome.sidePanel.open(options);
        }
        return Promise.reject(new Error('sidePanel not supported'));
      },
      setOptions(options) {
        if (typeof chrome.sidePanel !== 'undefined') {
          return chrome.sidePanel.setOptions(options);
        }
        return Promise.reject(new Error('sidePanel not supported'));
      },
    },
  };

  // Handle runtime.lastError
  Object.defineProperty(wrap.runtime, 'lastError', {
    get() {
      return chrome.runtime.lastError;
    },
  });

  global.browser = wrap;

})(typeof globalThis !== 'undefined' ? globalThis : typeof window !== 'undefined' ? window : typeof global !== 'undefined' ? global : this);
