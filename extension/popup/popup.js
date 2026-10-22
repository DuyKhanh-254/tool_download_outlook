/**
 * OneDrive & SharePoint PDF Downloader - Popup Script
 * Điều khiển tương tác giữa người dùng và Content Script trên tab đang hoạt động.
 */

document.addEventListener('DOMContentLoaded', async () => {
  // Elements
  const docTitleEl = document.getElementById('doc-title');
  const docPagesEl = document.getElementById('doc-pages');
  const envBadgeEl = document.getElementById('env-badge');
  const pageStartInput = document.getElementById('page-start');
  const pageEndInput = document.getElementById('page-end');
  const autoScrollCheck = document.getElementById('auto-scroll');
  const pageDelaySelect = document.getElementById('page-delay');
  const imgQualitySelect = document.getElementById('img-quality');
  const btnStart = document.getElementById('btn-start');
  const btnCancel = document.getElementById('btn-cancel');
  const progressSection = document.getElementById('progress-section');
  const progressStatus = document.getElementById('progress-status');
  const progressPercent = document.getElementById('progress-percent');
  const progressFill = document.getElementById('progress-fill');
  const progressCounter = document.getElementById('progress-counter');
  const statusBox = document.getElementById('status-box');
  const statusText = document.getElementById('status-text');

  let activeTab = null;

  // Lấy active tab hiện tại
  try {
    const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
    activeTab = tabs[0];
  } catch (err) {
    console.error('Không thể lấy active tab:', err);
  }

  if (!activeTab || !activeTab.url) {
    setUnsupportedState('Không thể kết nối với tab hiện tại.');
    return;
  }

  const url = activeTab.url.toLowerCase();
  const isSharePoint =
    url.includes('sharepoint.com') ||
    url.includes('onedrive.live.com') ||
    url.includes('office.com') ||
    url.includes('officeapps.live.com');

  if (!isSharePoint) {
    setUnsupportedState('Tiện ích này hoạt động trên trang xem tài liệu SharePoint hoặc OneDrive.');
    return;
  }

  // Khôi phục tiến trình từ background (nếu đang tải dở)
  checkExistingProgress();

  // Yêu cầu Content Script gửi thông tin tài liệu
  requestDocumentInfo();

  /**
   * Kiểm tra xem tab có đang tải dở không
   */
  function checkExistingProgress() {
    chrome.runtime.sendMessage({ type: 'GET_PROGRESS', tabId: activeTab.id }, (response) => {
      if (response && response.status && response.status !== 'IDLE') {
        if (!response.isDone) {
          showProgressUI();
          updateProgressUI(response);
        }
      }
    });
  }

  /**
   * Lấy thông tin tài liệu từ Content Script
   */
  function requestDocumentInfo() {
    chrome.tabs.sendMessage(activeTab.id, { action: 'GET_PAGE_INFO' }, (response) => {
      if (chrome.runtime.lastError) {
        docTitleEl.textContent = activeTab.title || 'Tài liệu SharePoint';
        docPagesEl.textContent = 'Trang đang xem';
        return;
      }

      if (response) {
        if (response.title) {
          docTitleEl.textContent = response.title;
          docTitleEl.title = response.title;
        }
        if (response.pageCount && response.pageCount > 1) {
          docPagesEl.textContent = `Ước tính: ~${response.pageCount} trang`;
          pageEndInput.value = response.pageCount;
        } else {
          docPagesEl.textContent = 'Trang hiện tại (Quét cuộn)';
        }
        if (response.isRunning) {
          showProgressUI();
        }
      }
    });
  }

  /**
   * Xử lý khi trang không phải SharePoint/OneDrive
   */
  function setUnsupportedState(msg) {
    envBadgeEl.textContent = 'Inactive';
    envBadgeEl.style.background = 'rgba(248, 81, 73, 0.15)';
    envBadgeEl.style.color = '#f85149';
    envBadgeEl.style.borderColor = 'rgba(248, 81, 73, 0.3)';

    docTitleEl.textContent = 'Chưa phát hiện tài liệu';
    docPagesEl.textContent = 'N/A';
    btnStart.disabled = true;
    statusText.textContent = msg;
  }

  /**
   * Bấm nút "Bắt đầu tải PDF ngay"
   */
  btnStart.addEventListener('click', () => {
    const pageDelay = parseInt(pageDelaySelect.value, 10) || 1500;
    const qualityVal = imgQualitySelect.value;
    const imageFormat = qualityVal === 'png-max' ? 'image/png' : 'image/jpeg';
    const imageQuality = qualityVal === 'png-max' ? 1.0 : 0.95;
    const startPage = parseInt(pageStartInput.value, 10) || 1;
    const endPage = parseInt(pageEndInput.value, 10) || 1;
    const autoScanToEnd = autoScrollCheck.checked;

    showProgressUI();
    progressStatus.textContent = `Đang chuẩn bị quét từ trang ${startPage}...`;
    progressPercent.textContent = '0%';
    progressFill.style.width = '0%';
    progressCounter.textContent = 'Bắt đầu...';
    statusText.textContent = 'Đang tự động cuộn và trích xuất từng trang sang PDF. Vui lòng giữ tab mở.';

    chrome.tabs.sendMessage(
      activeTab.id,
      {
        action: 'START_DOWNLOAD',
        pageDelay: pageDelay,
        imageFormat: imageFormat,
        imageQuality: imageQuality,
        startPage: startPage,
        endPage: endPage,
        autoScanToEnd: autoScanToEnd
      },
      (response) => {
        if (chrome.runtime.lastError) {
          console.warn('Lỗi gửi message đến content script:', chrome.runtime.lastError.message);
          statusText.textContent = 'Hãy F5 (tải lại) tab SharePoint để nạp extension, sau đó thử lại.';
          resetUI();
        }
      }
    );
  });

  /**
   * Bấm nút "Dừng lại"
   */
  btnCancel.addEventListener('click', () => {
    chrome.tabs.sendMessage(activeTab.id, { action: 'CANCEL_DOWNLOAD' }, () => {
      statusText.textContent = 'Đã gửi yêu cầu dừng quá trình tải.';
      resetUI();
    });
  });

  /**
   * Lắng nghe cập nhật tiến trình từ background hoặc content script
   */
  chrome.runtime.onMessage.addListener((message) => {
    if (message.type === 'UPDATE_PROGRESS') {
      updateProgressUI(message);
    }
  });

  function updateProgressUI(data) {
    showProgressUI();

    if (data.filename) {
      docTitleEl.textContent = data.filename;
    }

    if (data.status) {
      progressStatus.textContent = data.status;
    }

    const percent = data.percent || 0;
    progressPercent.textContent = `${percent}%`;
    progressFill.style.width = `${percent}%`;

    if (data.total > 0) {
      progressCounter.textContent = `${data.current} / ${data.total} trang`;
    }

    if (data.isDone) {
      if (data.error) {
        statusText.textContent = `❌ Lỗi: ${data.error}`;
      } else {
        statusText.textContent = `🎉 Hoàn tất! File đã được lưu tự động vào thư mục Downloads của bạn.`;
      }
      resetUI();
    }
  }

  function showProgressUI() {
    progressSection.style.display = 'block';
    btnStart.style.display = 'none';
    btnCancel.style.display = 'flex';
  }

  function resetUI() {
    btnStart.style.display = 'flex';
    btnStart.disabled = false;
    btnStart.innerHTML = `
      <svg viewBox="0 0 24 24"><path d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/></svg>
      <span>Tải lại</span>
    `;
    btnCancel.style.display = 'none';
  }

  // Help link
  document.getElementById('link-help').addEventListener('click', (e) => {
    e.preventDefault();
    alert(
      'HƯỚNG DẪN SỬ DỤNG:\n\n' +
        '1. Mở file PDF/Word/PowerPoint trên SharePoint hoặc OneDrive.\n' +
        '2. Bạn có thể chỉnh "Từ trang" và "Đến trang" (ví dụ: Từ 1 đến 88) hoặc để chế độ "Quét cuộn liên tục đến hết file".\n' +
        '3. Bấm "Bắt đầu tải PDF ngay". Extension sẽ tự động đưa tài liệu về trang 1, cuộn quét từng trang cho đến trang cuối và lưu file PDF hoàn chỉnh!'
    );
  });
});
