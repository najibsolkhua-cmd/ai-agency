// Service worker: screenshots, downloads, injection fallback and the table page.

const WB = /^https:\/\/([a-z0-9-]+\.)?wildberries\.ru\//i;

// Make sure the content script is running in the tab (it is missing in tabs
// that were already open when the extension was installed or updated).
async function ensureContent(tabId) {
  try {
    const r = await chrome.tabs.sendMessage(tabId, { type: 'ping' });
    if (r && r.ok) return true;
  } catch (_) { /* not injected yet */ }
  await chrome.scripting.insertCSS({ target: { tabId }, files: ['content.css'] });
  await chrome.scripting.executeScript({ target: { tabId }, files: ['content.js'] });
  return true;
}

chrome.commands.onCommand.addListener(async (command, tab) => {
  if (command !== 'save-seller' || !tab || !WB.test(tab.url || '')) return;
  await ensureContent(tab.id);
  chrome.tabs.sendMessage(tab.id, { type: 'save', withShot: true }).catch(() => {});
});

chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
  if (msg.type === 'capture') {
    chrome.tabs
      .captureVisibleTab(sender.tab.windowId, { format: 'png' })
      .then((dataUrl) => sendResponse({ ok: true, dataUrl }))
      .catch((e) => sendResponse({ ok: false, error: String(e) }));
    return true;
  }
  if (msg.type === 'download') {
    chrome.downloads
      .download({ url: msg.url, filename: msg.filename, conflictAction: 'uniquify', saveAs: false })
      .then(() => sendResponse({ ok: true }))
      .catch((e) => sendResponse({ ok: false, error: String(e) }));
    return true;
  }
  if (msg.type === 'ensure') {
    ensureContent(msg.tabId)
      .then(() => sendResponse({ ok: true }))
      .catch((e) => sendResponse({ ok: false, error: String(e) }));
    return true;
  }
  if (msg.type === 'openTable') {
    chrome.tabs.create({ url: chrome.runtime.getURL('table.html') });
  }
  return false;
});
