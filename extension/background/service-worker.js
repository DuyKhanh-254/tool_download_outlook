/**
 * OneDrive & SharePoint PDF Downloader - Background Service Worker
 * Quản lý download file, lưu trữ trạng thái tiến trình và trao đổi message giữa popup và content script.
 */

// Lưu trạng thái tiến trình tải hiện tại cho từng tab
const tabStates = {};

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  const tabId = sender.tab ? sender.tab.id : message.tabId;

  switch (message.type) {
    case 'UPDATE_PROGRESS':
      if (tabId) {
        tabStates[tabId] = {
          status: message.status,
          current: message.current,
          total: message.total,
          percent: message.percent,
          filename: message.filename,
          isDone: message.isDone || false,
          error: message.error || null,
          timestamp: Date.now()
        };
      }
      sendResponse({ received: true });
      break;

    case 'GET_PROGRESS':
      if (tabId && tabStates[tabId]) {
        sendResponse(tabStates[tabId]);
      } else {
        sendResponse({ status: 'IDLE', current: 0, total: 0, percent: 0 });
      }
      break;

    case 'CLEAR_PROGRESS':
      if (tabId && tabStates[tabId]) {
        delete tabStates[tabId];
      }
      sendResponse({ cleared: true });
      break;

    case 'DOWNLOAD_PDF':
      if (message.dataUrl) {
        const filename = message.filename || 'sharepoint_document.pdf';
        chrome.downloads.download({
          url: message.dataUrl,
          filename: filename,
          saveAs: false,
          conflictAction: 'uniquify'
        }, (downloadId) => {
          if (chrome.runtime.lastError) {
            console.error('Download error:', chrome.runtime.lastError.message);
            sendResponse({ success: false, error: chrome.runtime.lastError.message });
          } else {
            console.log('Download started, ID:', downloadId);
            sendResponse({ success: true, downloadId: downloadId });
          }
        });
        return true; // Async response
      }
      break;

    default:
      break;
  }
  return true;
});

// Xóa state khi tab bị đóng
chrome.tabs.onRemoved.addListener((tabId) => {
  if (tabStates[tabId]) {
    delete tabStates[tabId];
  }
});
