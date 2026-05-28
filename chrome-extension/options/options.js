// --- Options page logic -------------------------------------------------------─

const DEFAULTS = {
  defaultToc: false,
  defaultCaptions: true,
  defaultFormat: true,
  defaultImageWidth: 5.5,
  bridgeUrl: 'http://localhost:9876'
};

document.addEventListener('DOMContentLoaded', () => {
  // Load saved settings
  chrome.storage.local.get(DEFAULTS, (items) => {
    document.getElementById('defaultToc').checked = items.defaultToc;
    document.getElementById('defaultCaptions').checked = items.defaultCaptions;
    document.getElementById('defaultFormat').checked = items.defaultFormat;
    document.getElementById('defaultImageWidth').value = items.defaultImageWidth;
    document.getElementById('bridgeUrl').value = items.bridgeUrl;
  });

  // Save
  document.getElementById('btnSave').addEventListener('click', () => {
    chrome.storage.local.set({
      defaultToc: document.getElementById('defaultToc').checked,
      defaultCaptions: document.getElementById('defaultCaptions').checked,
      defaultFormat: document.getElementById('defaultFormat').checked,
      defaultImageWidth: parseFloat(document.getElementById('defaultImageWidth').value) || 5.5,
      bridgeUrl: document.getElementById('bridgeUrl').value.trim() || 'http://localhost:9876'
    }, () => {
      const indicator = document.getElementById('savedIndicator');
      indicator.classList.add('show');
      setTimeout(() => indicator.classList.remove('show'), 2000);
    });
  });

  // Reset
  document.getElementById('btnReset').addEventListener('click', () => {
    chrome.storage.local.set(DEFAULTS, () => {
      document.getElementById('defaultToc').checked = DEFAULTS.defaultToc;
      document.getElementById('defaultCaptions').checked = DEFAULTS.defaultCaptions;
      document.getElementById('defaultFormat').checked = DEFAULTS.defaultFormat;
      document.getElementById('defaultImageWidth').value = DEFAULTS.defaultImageWidth;
      document.getElementById('bridgeUrl').value = DEFAULTS.bridgeUrl;

      const indicator = document.getElementById('savedIndicator');
      indicator.classList.add('show');
      setTimeout(() => indicator.classList.remove('show'), 2000);
    });
  });
});
