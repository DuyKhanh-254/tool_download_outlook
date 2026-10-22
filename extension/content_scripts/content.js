/**
 * OneDrive & SharePoint PDF Downloader - Content Script
 * Tự động nhận diện tài liệu, cuộn từ trang 1 đến trang cuối (kể cả khi viewer lazy-load),
 * trích xuất raw canvas từng trang và đóng gói thành file PDF hoàn chỉnh qua jsPDF.
 */

(() => {
  if (window.__SP_PDF_DOWNLOADER_INJECTED__) return;
  window.__SP_PDF_DOWNLOADER_INJECTED__ = true;

  // Trạng thái hoạt động
  const state = {
    isRunning: false,
    isCancelled: false,
    currentPage: 1,
    startPage: 1,
    endPage: 1,
    totalPages: 1,
    docTitle: 'tailieu.pdf',
    capturedImages: [],
    pageDelay: 1500, // ms
    imageFormat: 'image/jpeg',
    imageQuality: 0.95,
    autoScanToEnd: true
  };

  const sleep = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  /**
   * Trích xuất tên file từ URL, Title hoặc DOM SharePoint
   */
  function extractDocumentTitle() {
    try {
      const urlObj = new URL(window.location.href);
      const idParam = urlObj.searchParams.get('id') || urlObj.searchParams.get('file');
      if (idParam) {
        const decoded = decodeURIComponent(idParam);
        const parts = decoded.split('/');
        const candidate = parts[parts.length - 1];
        if (candidate && candidate.trim().length > 0) {
          let name = candidate.trim();
          if (!name.toLowerCase().endsWith('.pdf')) name += '.pdf';
          return name;
        }
      }

      // Kiểm tra tiêu đề trang
      let title = document.title;
      if (title) {
        title = title
          .replace(/\s*-\s*SharePoint.*$/i, '')
          .replace(/\s*-\s*OneDrive.*$/i, '')
          .replace(/\s*-\s*Word.*$/i, '')
          .replace(/\s*-\s*PowerPoint.*$/i, '')
          .trim();
        if (title.length > 0) {
          if (!title.toLowerCase().endsWith('.pdf')) title += '.pdf';
          return title;
        }
      }
    } catch (e) {
      console.warn('Lỗi trích xuất tên file:', e);
    }
    return 'SharePoint_Document.pdf';
  }

  state.docTitle = extractDocumentTitle();

  /**
   * Phát hiện container cuộn chính của Viewer
   */
  function findScrollContainer() {
    const selectors = [
      '#viewerContainer',
      '.pdfViewer',
      '#main-content',
      '#WACViewPanel_ScrollViewer',
      'div[role="region"][aria-label*="Document" i]',
      'div[role="main"]',
      '.od-ItemContent-content',
      '.CanvasComponent'
    ];

    for (const sel of selectors) {
      const el = document.querySelector(sel);
      if (el && el.scrollHeight > el.clientHeight) return el;
    }

    // Tìm phần tử bất kỳ có thanh cuộn dọc
    const allDivs = document.querySelectorAll('div');
    for (const div of allDivs) {
      const style = window.getComputedStyle(div);
      if (
        (style.overflowY === 'auto' || style.overflowY === 'scroll') &&
        div.scrollHeight > div.clientHeight + 50
      ) {
        return div;
      }
    }

    return document.scrollingElement || document.documentElement || window;
  }

  /**
   * Quét và phát hiện số trang hiện tại & tổng số trang từ nhiều nguồn khác nhau
   */
  function scanDocumentPages() {
    let total = 0;
    let current = 1;

    // 1. Kiểm tra input số trang (PDF.js / SharePoint viewer)
    const inputs = document.querySelectorAll(
      'input[type="number"], input[role="spinbutton"], input.pageNumber, input[aria-label*="page" i], input[aria-label*="trang" i]'
    );
    for (const inp of inputs) {
      const max = parseInt(
        inp.getAttribute('aria-valuemax') || inp.max || inp.getAttribute('max'),
        10
      );
      const val = parseInt(
        inp.getAttribute('aria-valuenow') || inp.value,
        10
      );
      if (max && max > 1) total = Math.max(total, max);
      if (val && val >= 1) current = val;
    }

    // 2. Kiểm tra thuộc tính aria-setsize (Fluent UI / Office Online)
    const setSizeEls = document.querySelectorAll('[aria-setsize]');
    for (const el of setSizeEls) {
      const size = parseInt(el.getAttribute('aria-setsize'), 10);
      if (size && size > 1) total = Math.max(total, size);
    }

    // 3. Quét các thẻ thumbnails (hình thu nhỏ)
    const thumbs = document.querySelectorAll(
      '.thumbnail, [data-thumbnail-index], [data-page-index], [role="tablist"] [role="tab"]'
    );
    if (thumbs.length > 1) {
      total = Math.max(total, thumbs.length);
    }

    // 4. Quét các phần tử trang đã nạp trong DOM
    const pageDivs = document.querySelectorAll(
      'div.page[data-page-number], div.page, .pdf-page, .WacPage, div[id*="pageContainer" i]'
    );
    if (pageDivs.length > 1) {
      total = Math.max(total, pageDivs.length);
    }

    // 5. Quét text trên thanh công cụ / header tìm mẫu "X of Y" hoặc "X / Y"
    const textEls = document.querySelectorAll('span, div, p, label');
    for (const el of textEls) {
      if (el.children.length === 0 && el.textContent) {
        const text = el.textContent.trim();
        // Ví dụ: "1 of 88", "1 / 88", "Trang 1 / 88", "1 trên 88"
        const m = text.match(/(?:Trang|Page)?\s*(\d+)\s*(?:\/|of|của|trên)\s*(\d+)/i);
        if (m) {
          current = parseInt(m[1], 10) || current;
          total = Math.max(total, parseInt(m[2], 10));
          break;
        }
        // Hoặc mẫu riêng lẻ "/ 88"
        const m2 = text.match(/^\s*(?:\/|of|của|trên)\s*(\d+)\s*$/i);
        if (m2) {
          total = Math.max(total, parseInt(m2[1], 10));
        }
      }
    }

    // Nếu không tìm thấy, tối thiểu là 1 trang
    if (total < 1) total = 1;
    return { total, current };
  }

  /**
   * Mở thanh Thumbnail bên trái nếu có (để nạp danh sách toàn bộ trang)
   */
  function tryOpenThumbnails() {
    try {
      const toggleButtons = document.querySelectorAll(
        'button[aria-label*="thumbnail" i], button[aria-label*="hình thu nhỏ" i], button[title*="thumbnail" i], button[aria-label*="sidebar" i], button[aria-label*="mục lục" i]'
      );
      for (const btn of toggleButtons) {
        if (btn.offsetParent !== null && !btn.classList.contains('active')) {
          btn.click();
          break;
        }
      }
    } catch (e) {
      // bỏ qua
    }
  }

  /**
   * Chuyển đến trang cụ thể (1-indexed)
   */
  async function navigateToPage(pageNum, totalPages, scrollContainer) {
    // Cách 1: Nhập vào ô số trang nếu có
    const pageInputs = document.querySelectorAll(
      'input[type="number"], input[role="spinbutton"], input.pageNumber, input[aria-label*="page" i], input[aria-label*="trang" i]'
    );
    for (const inp of pageInputs) {
      if (inp.offsetParent !== null) {
        inp.focus();
        inp.value = pageNum;
        inp.dispatchEvent(new Event('input', { bubbles: true }));
        inp.dispatchEvent(new Event('change', { bubbles: true }));
        inp.dispatchEvent(
          new KeyboardEvent('keydown', { key: 'Enter', keyCode: 13, which: 13, bubbles: true })
        );
        return;
      }
    }

    // Cách 2: Bấm vào thumbnail trang nếu có
    const thumbs = document.querySelectorAll(
      '.thumbnail, [data-thumbnail-index], [data-page-index]'
    );
    if (thumbs && thumbs.length >= pageNum) {
      const targetThumb = thumbs[pageNum - 1];
      targetThumb.scrollIntoView({ behavior: 'instant', block: 'center' });
      targetThumb.click();
      return;
    }

    // Cách 3: Tìm phần tử trang trong DOM và cuộn tới
    const pageEl = document.querySelector(
      `div.page[data-page-number="${pageNum}"], [data-page-index="${pageNum - 1}"], #pageContainer${pageNum}`
    );
    if (pageEl) {
      pageEl.scrollIntoView({ behavior: 'instant', block: 'center' });
      return;
    }

    // Cách 4: Bấm nút Next trang nếu chuyển sang trang tiếp theo
    const nextBtn = document.querySelector(
      'button[aria-label*="next" i], button[aria-label*="sau" i], button[title*="next" i], .nextPageButton'
    );
    if (nextBtn && !nextBtn.disabled && nextBtn.offsetParent !== null) {
      nextBtn.click();
      return;
    }

    // Cách 5: Cuộn container & giả lập phím PageDown
    if (scrollContainer) {
      if (pageNum === 1) {
        scrollContainer.scrollTop = 0;
      } else if (totalPages > 1 && scrollContainer.scrollHeight > scrollContainer.clientHeight) {
        const step = (scrollContainer.scrollHeight - scrollContainer.clientHeight) / (totalPages - 1 || 1);
        scrollContainer.scrollTop = (pageNum - 1) * step;
      } else {
        scrollContainer.scrollBy(0, scrollContainer.clientHeight * 0.9);
      }

      // Phát sự kiện phím PageDown
      const keyEvt = new KeyboardEvent('keydown', {
        key: 'PageDown',
        code: 'PageDown',
        keyCode: 34,
        which: 34,
        bubbles: true
      });
      document.dispatchEvent(keyEvt);
      if (scrollContainer.dispatchEvent) scrollContainer.dispatchEvent(keyEvt);
    }
  }

  /**
   * Tìm canvas trang hiện tại đang hiển thị trong màn hình
   */
  function findCurrentVisibleCanvas() {
    const canvases = Array.from(document.querySelectorAll('canvas')).filter((c) => {
      if (c.width < 150 || c.height < 150 || c.style.display === 'none') return false;
      const rect = c.getBoundingClientRect();
      // Canvas phải có phần nằm trong viewport
      return rect.bottom > 0 && rect.top < window.innerHeight && rect.width > 100;
    });

    if (canvases.length === 0) return null;

    // Chọn canvas có diện tích hiển thị lớn nhất
    canvases.sort((a, b) => {
      const rectA = a.getBoundingClientRect();
      const rectB = b.getBoundingClientRect();
      const visibleHeightA = Math.min(rectA.bottom, window.innerHeight) - Math.max(rectA.top, 0);
      const visibleHeightB = Math.min(rectB.bottom, window.innerHeight) - Math.max(rectB.top, 0);
      return visibleHeightB * rectB.width - visibleHeightA * rectA.width;
    });

    return canvases[0];
  }

  /**
   * Chờ canvas của trang render hoàn tất
   */
  async function waitForCanvasReady(maxWaitMs = 4000, previousDataUrl = null) {
    const startTime = Date.now();
    let bestCanvas = null;

    while (Date.now() - startTime < maxWaitMs) {
      if (state.isCancelled) return null;

      const canvas = findCurrentVisibleCanvas();
      if (canvas && canvas.width > 150 && canvas.height > 150) {
        bestCanvas = canvas;
        const currentData = extractImageFromCanvas(canvas);
        // Nếu canvas có nội dung và khác với trang trước đó
        if (currentData && (!previousDataUrl || currentData.dataUrl !== previousDataUrl)) {
          return currentData;
        }
      }
      await sleep(200);
    }

    if (bestCanvas) {
      return extractImageFromCanvas(bestCanvas);
    }
    return null;
  }

  /**
   * Trích xuất hình ảnh từ thẻ canvas
   */
  function extractImageFromCanvas(canvas) {
    try {
      const parent = canvas.parentElement;
      const siblingCanvases = parent ? Array.from(parent.querySelectorAll('canvas')) : [];

      // Ghép nếu có nhiều tiles nhỏ
      if (siblingCanvases.length > 1) {
        let totalWidth = 0;
        let totalHeight = 0;
        siblingCanvases.forEach((c) => {
          totalWidth = Math.max(totalWidth, c.offsetLeft + c.width);
          totalHeight = Math.max(totalHeight, c.offsetTop + c.height);
        });

        if (totalWidth > 0 && totalHeight > 0) {
          const offscreen = document.createElement('canvas');
          offscreen.width = totalWidth;
          offscreen.height = totalHeight;
          const ctx = offscreen.getContext('2d');
          siblingCanvases.forEach((c) => {
            ctx.drawImage(c, c.offsetLeft, c.offsetTop);
          });
          return {
            dataUrl: offscreen.toDataURL(state.imageFormat, state.imageQuality),
            width: offscreen.width,
            height: offscreen.height
          };
        }
      }

      return {
        dataUrl: canvas.toDataURL(state.imageFormat, state.imageQuality),
        width: canvas.width,
        height: canvas.height
      };
    } catch (err) {
      console.error('Lỗi khi toDataURL từ canvas:', err);
      return null;
    }
  }

  /**
   * Thử Direct Download qua cookie (Method 2)
   */
  async function tryDirectDownload() {
    try {
      const url = new URL(window.location.href);
      if (!url.searchParams.has('download')) {
        const testUrl = new URL(window.location.href);
        testUrl.searchParams.set('download', '1');
        const resp = await fetch(testUrl.toString(), {
          method: 'HEAD',
          credentials: 'include'
        });
        const contentType = resp.headers.get('content-type') || '';
        if (resp.ok && contentType.includes('pdf')) {
          return testUrl.toString();
        }
      }
    } catch (e) {
      // Bỏ qua nếu không hỗ trợ
    }
    return null;
  }

  /**
   * QUY TRÌNH CHÍNH: Bắt đầu quét từ startPage đến endPage (hoặc đến hết file)
   */
  async function startCapture(options = {}) {
    if (state.isRunning) return;

    state.isRunning = true;
    state.isCancelled = false;
    state.capturedImages = [];
    state.pageDelay = options.pageDelay || 1500;
    state.imageFormat = options.imageFormat || 'image/jpeg';
    state.imageQuality = options.imageQuality || 0.95;
    state.docTitle = extractDocumentTitle();

    // Cập nhật thông số trang
    const detected = scanDocumentPages();
    state.startPage = options.startPage ? parseInt(options.startPage, 10) : 1;
    state.endPage = options.endPage ? parseInt(options.endPage, 10) : Math.max(detected.total, 1);
    state.totalPages = Math.max(state.endPage, detected.total);
    state.autoScanToEnd = options.autoScanToEnd !== undefined ? options.autoScanToEnd : true;

    updateStatusUI('Đang kiểm tra tài liệu và phương thức tối ưu...', 0, state.totalPages);

    // Bước 1: Thử Direct Download URL
    const directUrl = await tryDirectDownload();
    if (directUrl) {
      updateStatusUI('Phát hiện URL tải trực tiếp! Đang tải...', 1, 1);
      const a = document.createElement('a');
      a.href = directUrl;
      a.download = state.docTitle;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      finishCapture(true, 'Tải file trực tiếp thành công!');
      return;
    }

    // Bước 2: Chuẩn bị viewer
    tryOpenThumbnails();
    await sleep(300);

    const scrollContainer = findScrollContainer();

    // Điều hướng về startPage (mặc định là trang 1)
    updateStatusUI(`Đang điều hướng về trang ${state.startPage}...`, 0, state.totalPages);
    await navigateToPage(state.startPage, state.totalPages, scrollContainer);
    await sleep(state.pageDelay);

    // Bước 3: Vòng lặp quét từng trang tuần tự
    let lastDataUrl = null;
    let pageNum = state.startPage;
    let consecutiveDuplicateCount = 0;

    while (pageNum <= state.endPage || state.autoScanToEnd) {
      if (state.isCancelled) {
        finishCapture(false, 'Đã dừng quá trình tải theo yêu cầu.');
        return;
      }

      state.currentPage = pageNum;
      updateStatusUI(
        `Đang xử lý trang ${pageNum} / ${state.totalPages}...`,
        pageNum,
        state.totalPages
      );

      // Chuyển tới trang cần chụp
      await navigateToPage(pageNum, state.totalPages, scrollContainer);
      await sleep(state.pageDelay);

      // Chờ canvas nạp và lấy dữ liệu
      const imgData = await waitForCanvasReady(3000, lastDataUrl);

      if (imgData) {
        // Kiểm tra trùng lặp với trang vừa chụp
        if (lastDataUrl && imgData.dataUrl === lastDataUrl) {
          consecutiveDuplicateCount++;
          // Nếu trùng lặp quá 2 lần liên tiếp ở cuối tài liệu -> đã đến trang cuối
          if (consecutiveDuplicateCount >= 2 && pageNum >= state.endPage) {
            console.log('[SP Downloader] Đã chạm đến trang cuối cùng của tài liệu.');
            break;
          }
        } else {
          consecutiveDuplicateCount = 0;
          lastDataUrl = imgData.dataUrl;
          state.capturedImages.push(imgData);
        }
      } else {
        console.warn(`[SP Downloader] Không tìm thấy canvas ở trang ${pageNum}`);
        consecutiveDuplicateCount++;
        if (consecutiveDuplicateCount >= 3 && pageNum >= state.endPage) {
          break;
        }
      }

      // Kiểm tra nếu đã đến trang endPage được chỉ định
      if (!state.autoScanToEnd && pageNum >= state.endPage) {
        break;
      }

      // Nếu đang bật autoScanToEnd và vượt qua endPage ban đầu, tự động tăng endPage
      if (pageNum >= state.endPage && state.autoScanToEnd) {
        state.endPage++;
        state.totalPages = state.endPage;
      }

      pageNum++;
    }

    // Bước 4: Đóng gói thành file PDF qua jsPDF
    if (state.capturedImages.length === 0) {
      finishCapture(
        false,
        'Không thu thập được trang nào. Hãy kiểm tra xem file đã được mở xem trước trên màn hình chưa.'
      );
      return;
    }

    updateStatusUI(
      `Đang đóng gói ${state.capturedImages.length} trang thành file PDF...`,
      state.capturedImages.length,
      state.capturedImages.length
    );
    await sleep(200);

    try {
      assembleAndDownloadPDF();
    } catch (err) {
      console.error('Lỗi khi xuất PDF:', err);
      finishCapture(false, `Lỗi tạo PDF: ${err.message}`);
    }
  }

  /**
   * Ghép các ảnh thành file PDF bằng jsPDF
   */
  function assembleAndDownloadPDF() {
    const jspdfLib = window.jspdf || (window.jspdf_umd ? window.jspdf_umd.jsPDF : null);
    if (!jspdfLib || !jspdfLib.jsPDF) {
      throw new Error('Thư viện jsPDF chưa được nạp!');
    }

    const { jsPDF } = jspdfLib;
    let pdf = null;
    const images = state.capturedImages;

    for (let i = 0; i < images.length; i++) {
      const item = images[i];
      const orientation = item.width > item.height ? 'landscape' : 'portrait';

      if (i === 0) {
        pdf = new jsPDF({
          orientation: orientation,
          unit: 'px',
          format: [item.width, item.height],
          compress: true
        });
        pdf.addImage(item.dataUrl, 'JPEG', 0, 0, item.width, item.height, undefined, 'FAST');
      } else {
        pdf.addPage([item.width, item.height], orientation);
        pdf.addImage(item.dataUrl, 'JPEG', 0, 0, item.width, item.height, undefined, 'FAST');
      }
    }

    pdf.save(state.docTitle);
    finishCapture(
      true,
      `Hoàn thành! Đã tải xuống ${images.length} trang vào file: ${state.docTitle}`
    );
  }

  /**
   * Kết thúc tiến trình
   */
  function finishCapture(success, message) {
    state.isRunning = false;
    notifyProgress({
      status: success ? 'DONE' : 'FAILED',
      current: state.capturedImages.length,
      total: state.capturedImages.length || state.totalPages,
      percent: 100,
      filename: state.docTitle,
      isDone: true,
      error: success ? null : message
    });

    showToast(message, success ? 'success' : 'error');
    resetCardUI(message);
  }

  function notifyProgress(data) {
    try {
      chrome.runtime.sendMessage({
        type: 'UPDATE_PROGRESS',
        ...data
      });
    } catch (e) {
      // bỏ qua khi popup đóng
    }
  }

  function updateStatusUI(statusText, current, total) {
    const percent = total > 0 ? Math.min(100, Math.round((current / total) * 100)) : 0;
    notifyProgress({
      status: statusText,
      current: current,
      total: total,
      percent: percent,
      filename: state.docTitle,
      isDone: false
    });

    const card = document.getElementById('sp-downloader-card');
    if (card) {
      const statusEl = card.querySelector('.sp-status-text');
      const progressBar = card.querySelector('.sp-progress-bar');
      const metaPages = card.querySelector('.sp-progress-pages');
      const metaPercent = card.querySelector('.sp-progress-percent');

      if (statusEl) statusEl.textContent = statusText;
      if (progressBar) progressBar.style.width = `${percent}%`;
      if (metaPages) metaPages.textContent = total > 0 ? `${current}/${total} trang` : '';
      if (metaPercent) metaPercent.textContent = `${percent}%`;
    }
  }

  function resetCardUI(msg) {
    const card = document.getElementById('sp-downloader-card');
    if (card) {
      const btnStart = card.querySelector('#sp-btn-start');
      const btnStop = card.querySelector('#sp-btn-stop');
      const statusEl = card.querySelector('.sp-status-text');

      if (btnStart) {
        btnStart.disabled = false;
        btnStart.textContent = '📥 Tải lại';
      }
      if (btnStop) btnStop.style.display = 'none';
      if (statusEl) statusEl.textContent = msg || 'Sẵn sàng tải xuống.';
    }
  }

  function showToast(message, type = 'info') {
    const existing = document.querySelector('.sp-toast');
    if (existing) existing.remove();

    const toast = document.createElement('div');
    toast.className = 'sp-toast';
    toast.style.borderColor =
      type === 'success' ? '#2ea043' : type === 'error' ? '#f85149' : '#58a6ff';
    toast.innerHTML = `
      <span>${type === 'success' ? '✅' : type === 'error' ? '❌' : 'ℹ️'}</span>
      <span>${message}</span>
    `;
    document.body.appendChild(toast);

    setTimeout(() => {
      if (toast.parentElement) toast.remove();
    }, 4500);
  }

  /**
   * Tạo Floating Action Button (FAB) & Card điều khiển trên trang
   */
  function injectFloatingWidget() {
    if (document.getElementById('sp-downloader-widget')) return;

    const widget = document.createElement('div');
    widget.id = 'sp-downloader-widget';

    const fab = document.createElement('div');
    fab.className = 'sp-downloader-fab';
    fab.title = 'Bấm để mở công cụ tải PDF SharePoint';
    fab.innerHTML = `
      <div class="sp-downloader-fab-icon">
        <svg viewBox="0 0 24 24"><path d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/></svg>
      </div>
      <span>Tải PDF</span>
    `;

    const card = document.createElement('div');
    card.id = 'sp-downloader-card';
    card.innerHTML = `
      <div class="sp-card-header">
        <div class="sp-card-title">
          <span>📄</span>
          <span>SharePoint PDF Downloader</span>
        </div>
        <button class="sp-card-close" id="sp-card-close-btn">&times;</button>
      </div>
      <div class="sp-card-docname" title="${state.docTitle}">${state.docTitle}</div>
      
      <!-- Cấu hình số trang trực tiếp trên Card -->
      <div style="display: flex; align-items: center; justify-content: space-between; gap: 8px; margin-bottom: 12px; font-size: 12px; background: rgba(255,255,255,0.05); padding: 8px 10px; border-radius: 8px;">
        <span>Số trang quét:</span>
        <div style="display: flex; align-items: center; gap: 4px;">
          <span>Từ</span>
          <input type="number" id="sp-card-start-page" value="1" min="1" style="width: 48px; background: #0d1117; border: 1px solid rgba(255,255,255,0.2); color: #fff; border-radius: 4px; padding: 2px 4px; font-size: 11px; text-align: center;">
          <span>đến</span>
          <input type="number" id="sp-card-end-page" value="1" min="1" style="width: 48px; background: #0d1117; border: 1px solid rgba(255,255,255,0.2); color: #fff; border-radius: 4px; padding: 2px 4px; font-size: 11px; text-align: center;">
        </div>
      </div>

      <div class="sp-progress-container">
        <div class="sp-progress-bar"></div>
      </div>
      <div class="sp-progress-meta">
        <span class="sp-progress-pages">Sẵn sàng</span>
        <span class="sp-progress-percent">0%</span>
      </div>
      <div class="sp-status-text">Tự động cuộn từ trang 1 đến trang cuối để ghép thành file PDF.</div>
      <div class="sp-card-actions">
        <button class="sp-btn-primary" id="sp-btn-start">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="currentColor"><path d="M19 9h-4V3H9v6H5l7 7 7-7zM5 18v2h14v-2H5z"/></svg>
          Bắt đầu tải PDF
        </button>
        <button class="sp-btn-secondary" id="sp-btn-stop" style="display: none;">Dừng lại</button>
      </div>
    `;

    widget.appendChild(fab);
    document.body.appendChild(widget);
    document.body.appendChild(card);

    function refreshDetectedPages() {
      state.docTitle = extractDocumentTitle();
      const detected = scanDocumentPages();
      const docNameEl = card.querySelector('.sp-card-docname');
      const startInp = card.querySelector('#sp-card-start-page');
      const endInp = card.querySelector('#sp-card-end-page');

      if (docNameEl) docNameEl.textContent = state.docTitle;
      if (startInp && !startInp.matches(':focus')) startInp.value = 1;
      if (endInp && !endInp.matches(':focus')) {
        endInp.value = detected.total > 1 ? detected.total : (parseInt(endInp.value, 10) > 1 ? endInp.value : 1);
      }
    }

    fab.addEventListener('click', () => {
      refreshDetectedPages();
      card.style.display = card.style.display === 'block' ? 'none' : 'block';
    });

    const closeBtn = card.querySelector('#sp-card-close-btn');
    if (closeBtn) {
      closeBtn.addEventListener('click', () => {
        card.style.display = 'none';
      });
    }

    const btnStart = card.querySelector('#sp-btn-start');
    const btnStop = card.querySelector('#sp-btn-stop');

    if (btnStart) {
      btnStart.addEventListener('click', () => {
        btnStart.disabled = true;
        if (btnStop) btnStop.style.display = 'inline-block';
        const sPage = parseInt(card.querySelector('#sp-card-start-page').value, 10) || 1;
        const ePage = parseInt(card.querySelector('#sp-card-end-page').value, 10) || 1;

        startCapture({
          pageDelay: state.pageDelay,
          startPage: sPage,
          endPage: ePage,
          autoScanToEnd: true
        });
      });
    }

    if (btnStop) {
      btnStop.addEventListener('click', () => {
        state.isCancelled = true;
        btnStop.style.display = 'none';
        if (btnStart) btnStart.disabled = false;
        updateStatusUI('Đang dừng lại...', state.currentPage, state.totalPages);
      });
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', injectFloatingWidget);
  } else {
    injectFloatingWidget();
  }

  // Lắng nghe Message từ Extension Popup
  chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    switch (request.action) {
      case 'GET_PAGE_INFO':
        state.docTitle = extractDocumentTitle();
        const info = scanDocumentPages();
        sendResponse({
          title: state.docTitle,
          pageCount: info.total,
          currentPage: info.current,
          isRunning: state.isRunning
        });
        break;

      case 'START_DOWNLOAD':
        if (!state.isRunning) {
          const card = document.getElementById('sp-downloader-card');
          if (card) card.style.display = 'block';
          const btnStart = document.getElementById('sp-btn-start');
          const btnStop = document.getElementById('sp-btn-stop');
          if (btnStart) btnStart.disabled = true;
          if (btnStop) btnStop.style.display = 'inline-block';

          startCapture({
            pageDelay: request.pageDelay || 1500,
            imageQuality: request.imageQuality || 0.95,
            imageFormat: request.imageFormat || 'image/jpeg',
            startPage: request.startPage || 1,
            endPage: request.endPage || 1,
            autoScanToEnd: request.autoScanToEnd !== undefined ? request.autoScanToEnd : true
          });
          sendResponse({ success: true, message: 'Đã bắt đầu tải!' });
        } else {
          sendResponse({ success: false, message: 'Đang trong quá trình tải!' });
        }
        break;

      case 'CANCEL_DOWNLOAD':
        state.isCancelled = true;
        sendResponse({ success: true, message: 'Đã gửi lệnh hủy.' });
        break;

      default:
        break;
    }
    return true;
  });
})();
