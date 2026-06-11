/**
 * 双语对照模块 — 左侧原文、右侧译文，同步翻页和滚动。
 *
 * 这是用户最看重的核心功能。
 */
const bilingual = {
  enabled: false,
  translating: false,
  translatedCache: {},  // { `${pptId}:${slideIdx}`: htmlString }

  /** 开启/切换双语对照模式 */
  async toggle() {
    if (!slideRenderer.currentPptId) {
      showToast('请先上传 PPT 文件', 'error');
      return;
    }

    this.enabled = !this.enabled;
    const btn = document.getElementById('btn-bilingual');
    const translatedPane = document.getElementById('slide-pane-translated');
    const slideArea = document.getElementById('slide-area');
    const mainLayout = document.getElementById('main-layout');

    if (this.enabled) {
      // 开启双语模式：显示译文面板，切换布局（保留右侧聊天栏，隐藏缩略图）
      btn.classList.add('active');
      btn.textContent = '📖 退出对照';
      translatedPane.classList.remove('hidden');
      slideArea.classList.add('bilingual');
      mainLayout.classList.add('bilingual');

      // 布局变化后立即重缩放左侧原文，让它适配一半宽度
      slideRenderer.rescaleAll();

      // 检查翻译状态
      let status = { translated: false };
      try {
        const resp = await fetch(`/api/ppt/${slideRenderer.currentPptId}/translation-status`);
        status = await resp.json();
      } catch (e) {}

      if (!status.translated) {
        // 触发全文翻译
        this.translating = true;
        document.getElementById('slide-container-translated').innerHTML =
          '<div class="empty-state"><div class="spinner"></div><p>正在翻译全文，请稍候...</p></div>';

        try {
          const resp = await fetch(`/api/ppt/${slideRenderer.currentPptId}/translate-all`, { method: 'POST' });
          if (!resp.ok) throw new Error('翻译请求失败');
          const result = await resp.json();
          showToast(`全文翻译完成 (${result.slide_count} 页)`, 'success');
        } catch (e) {
          showToast('翻译失败: ' + e.message, 'error');
          this._exitBilingual();
          return;
        }
        this.translating = false;
      }

      // 加载当前页的译文
      await this._loadTranslatedSlide(slideRenderer.currentSlideIdx);
      // 两侧都重缩放一次，确保宽度一致
      slideRenderer.rescaleAll();
      this._syncScroll();

      showStatus('双语对照模式');

    } else {
      this._exitBilingual();
      showStatus('已退出双语对照模式');
    }
  },

  /** 退出双语模式（恢复布局） */
  _exitBilingual() {
    this.enabled = false;
    const btn = document.getElementById('btn-bilingual');
    btn.classList.remove('active');
    btn.textContent = '📖 双语对照';
    document.getElementById('slide-pane-translated').classList.add('hidden');
    document.getElementById('slide-area').classList.remove('bilingual');
    document.getElementById('main-layout').classList.remove('bilingual');
    // 恢复全宽后重缩放原文
    slideRenderer.rescaleAll();
  },

  /** 加载翻译后的幻灯片 */
  async _loadTranslatedSlide(slideIdx) {
    const pptId = slideRenderer.currentPptId;
    const cacheKey = `${pptId}:${slideIdx}`;

    const container = document.getElementById('slide-container-translated');
    container.innerHTML = '<div class="empty-state"><div class="spinner"></div><p>加载译文...</p></div>';

    try {
      let html = this.translatedCache[cacheKey];
      if (!html) {
        const resp = await fetch(`/api/ppt/${pptId}/slides/${slideIdx}/translated`);
        if (!resp.ok) {
          const err = await resp.json();
          throw new Error(err.detail || '加载译文失败');
        }
        html = await resp.text();

        // 检查是否包含翻译失败标记
        if (!html.includes('翻译失败')) {
          this.translatedCache[cacheKey] = html;
        }
      }
      container.innerHTML = html;

      // 自适应缩放译文幻灯片（等待 reflow）
      slideRenderer.rescaleAll();

      // 预加载相邻页译文
      this._prefetchTranslated(pptId, slideIdx);
    } catch (e) {
      container.innerHTML = `<div class="empty-state"><p>⚠️ 译文加载失败: ${e.message}</p></div>`;
    }
  },

  /** 预加载相邻译文 */
  async _prefetchTranslated(pptId, slideIdx) {
    for (const offset of [-1, 1]) {
      const idx = slideIdx + offset;
      if (idx < 0 || idx >= slideRenderer.totalSlides) continue;
      const key = `${pptId}:${idx}`;
      if (this.translatedCache[key]) continue;
      try {
        const resp = await fetch(`/api/ppt/${pptId}/slides/${idx}/translated`);
        if (resp.ok) {
          this.translatedCache[key] = await resp.text();
        }
      } catch (e) { /* ignore */ }
    }
  },

  /** 同步两侧滚动 */
  _syncScroll() {
    const origContainer = document.getElementById('slide-container');
    const transContainer = document.getElementById('slide-container-translated');

    if (!origContainer || !transContainer) return;

    let syncing = false;

    function sync(source, target) {
      if (syncing) return;
      syncing = true;
      // 按比例同步滚动
      const sourceMax = source.scrollHeight - source.clientHeight;
      const targetMax = target.scrollHeight - target.clientHeight;
      if (sourceMax > 0 && targetMax > 0) {
        const ratio = source.scrollTop / sourceMax;
        target.scrollTop = ratio * targetMax;
      }
      syncing = false;
    }

    // 移除旧监听器（用新方式）
    origContainer.onscroll = () => sync(origContainer, transContainer);
    transContainer.onscroll = () => sync(transContainer, origContainer);
  },

  /** 在双语模式下加载指定页（由导航模块调用） */
  async loadSlide(slideIdx) {
    if (this.enabled) {
      await this._loadTranslatedSlide(slideIdx);
      this._syncScroll();
    }
  },

  /** 清除缓存 */
  clearCache() {
    this.translatedCache = {};
    this.translating = false;
    this._exitBilingual();
  },
};

// ── 双语按钮 ──────────────────────────────────────────────
document.getElementById('btn-bilingual').addEventListener('click', () => bilingual.toggle());

// ── 监听幻灯片渲染，同步加载译文 ─────────────────────────
window.addEventListener('slide-rendered', async (e) => {
  const { slideIdx } = e.detail;
  if (bilingual.enabled) {
    await bilingual.loadSlide(slideIdx);
  }
});
