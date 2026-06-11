/**
 * 幻灯片渲染器 — 加载、缓存和自适应缩放显示幻灯片 HTML。
 */
const slideRenderer = {
  cache: {},       // { `${pptId}:${idx}`: htmlString }
  currentPptId: null,
  currentSlideIdx: 0,
  totalSlides: 0,

  /** 加载并渲染指定幻灯片 */
  async loadSlide(pptId, slideIdx, targetContainerId = 'slide-container') {
    const container = document.getElementById(targetContainerId);
    if (!container) return;
    const key = `${pptId}:${slideIdx}`;

    // 显示加载状态
    container.innerHTML = '<div class="empty-state"><div class="spinner"></div><p>加载中...</p></div>';

    try {
      let html = this.cache[key];
      if (!html) {
        html = await api.getSlide(pptId, slideIdx);
        this.cache[key] = html;
        this._prefetch(pptId, slideIdx);
      }

      container.innerHTML = html;
      this.currentPptId = pptId;
      this.currentSlideIdx = slideIdx;

      // 自适应缩放幻灯片
      this._scaleToFit(container);

      // 通知各模块幻灯片已渲染
      window.dispatchEvent(new CustomEvent('slide-rendered', {
        detail: { pptId, slideIdx, totalSlides: this.totalSlides }
      }));

      container.scrollTop = 0;

    } catch (err) {
      container.innerHTML = `<div class="empty-state"><p>加载失败: ${err.message}</p></div>`;
    }
  },

  /** 将容器内的 .slide-wrapper 缩放到适合容器宽度 */
  _scaleToFit(container) {
    if (!container) return;
    const wrapper = container.querySelector('.slide-wrapper');
    if (!wrapper) return;

    const naturalWidth = parseFloat(wrapper.style.width) || 960;

    // 可用宽度 = 容器内容宽度 - 左右 padding
    const style = window.getComputedStyle(container);
    const paddingH = parseFloat(style.paddingLeft || 0) + parseFloat(style.paddingRight || 0);
    const availableWidth = container.clientWidth - paddingH;

    // 容器宽度还没准备好（reflow 未完成），跳过本次，由 rAF 重试
    if (availableWidth <= 0) return;

    let scale = availableWidth / naturalWidth;
    if (scale > 1) scale = 1;  // 不放大，只缩小

    // 用 zoom 缩放：会真实重排布局，滚动/高度都正确（Chromium/Edge 原生支持）
    wrapper.style.zoom = scale;
  },

  /** 重新缩放所有可见的幻灯片面板（布局变化后调用，等待 reflow） */
  rescaleAll() {
    // 用两帧 rAF 确保 CSS 布局已完成 reflow 再测量宽度
    requestAnimationFrame(() => {
      requestAnimationFrame(() => {
        document.querySelectorAll('.slide-container-inner').forEach(c => this._scaleToFit(c));
      });
    });
  },

  /** 获取当前幻灯片中的纯文本 */
  getCurrentSlideText() {
    const wrapper = document.querySelector('.slide-wrapper');
    if (!wrapper) return '';
    return wrapper.textContent.trim();
  },

  /** 预加载相邻幻灯片 */
  _prefetch(pptId, slideIdx) {
    const keys = [
      `${pptId}:${slideIdx - 1}`,
      `${pptId}:${slideIdx + 1}`,
    ];
    keys.forEach(async (key) => {
      if (this.cache[key]) return;
      const [, idxStr] = key.split(':');
      const idx = parseInt(idxStr);
      if (idx >= 0 && idx < this.totalSlides) {
        try {
          const html = await api.getSlide(pptId, idx);
          this.cache[key] = html;
        } catch (e) { /* 静默失败 */ }
      }
    });
  },

  /** 清除缓存 */
  clearCache() {
    this.cache = {};
    this.currentPptId = null;
    this.currentSlideIdx = 0;
    this.totalSlides = 0;
  },

  /** 设置幻灯片总数 */
  setTotalSlides(count) {
    this.totalSlides = count;
  },
};

// ── 窗口大小变化时重新缩放 ──────────────────────────────
let _resizeTimer = null;
window.addEventListener('resize', () => {
  clearTimeout(_resizeTimer);
  _resizeTimer = setTimeout(() => slideRenderer.rescaleAll(), 150);
});
