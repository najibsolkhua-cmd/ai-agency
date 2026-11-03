// Service worker: screenshots, downloads and opening the table page.

function saveWithShot(tab) {
  if (!tab || !tab.id || !/^https:\/\/www\.wildberries\.ru\//.test(tab.url || '')) return;
  chrome.tabs.sendMessage(tab.id, { type: 'save', withShot: true }).catch(() => {});
}

// Clicking the toolbar icon or pressing the shortcut grants activeTab,
// which is what allows captureVisibleTab below.
chrome.action.onClicked.addListener(saveWithShot);
chrome.commands.onCommand.addListener((command, tab) => {
  if (command === 'save-seller') saveWithShot(tab);
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
  if (msg.type === 'openTable') {
    chrome.tabs.create({ url: chrome.runtime.getURL('table.html') });
  }
  return false;
});
