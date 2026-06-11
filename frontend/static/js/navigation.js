/**
 * 导航模块 — 缩略图条、前后翻页按钮、键盘快捷键。
 */
const navigation = {
  currentPptId: null,
  slideCount: 0,

  /** 初始化导航（上传 PPT 后调用） */
  init(pptId, slideCount) {
    this.currentPptId = pptId;
    this.slideCount = slideCount;
    this._renderThumbnails();
    this._setActiveThumbnail(0);
    this._updateControls(0);
  },

  /** 跳转到指定页 */
  goToSlide(slideIdx) {
    if (slideIdx < 0 || slideIdx >= this.slideCount) return;
    this._setActiveThumbnail(slideIdx);
    this._updateControls(slideIdx);
    slideRenderer.loadSlide(this.currentPptId, slideIdx);
  },

  /** 下一页 */
  next() {
    const nextIdx = slideRenderer.currentSlideIdx + 1;
    if (nextIdx < this.slideCount) this.goToSlide(nextIdx);
  },

  /** 上一页 */
  prev() {
    const prevIdx = slideRenderer.currentSlideIdx - 1;
    if (prevIdx >= 0) this.goToSlide(prevIdx);
  },

  /** 渲染缩略图列表 */
  _renderThumbnails() {
    const list = document.getElementById('thumbnail-list');
    list.innerHTML = '';
    for (let i = 0; i < this.slideCount; i++) {
      const item = document.createElement('div');
      item.className = 'thumbnail-item';
      item.textContent = `📄 ${i + 1}`;
      item.dataset.slideIdx = i;
      item.addEventListener('click', () => this.goToSlide(i));
      list.appendChild(item);
    }
  },

  /** 高亮当前缩略图 */
  _setActiveThumbnail(idx) {
    document.querySelectorAll('.thumbnail-item').forEach((el, i) => {
      el.classList.toggle('active', i === idx);
    });

    // 滚动缩略图到可见区域
    const active = document.querySelector('.thumbnail-item.active');
    if (active) {
      active.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
    }
  },

  /** 更新前后翻页按钮和页码指示器 */
  _updateControls(idx) {
    document.getElementById('btn-prev').disabled = idx <= 0;
    document.getElementById('btn-next').disabled = idx >= this.slideCount - 1;
    document.getElementById('slide-indicator').textContent =
      this.slideCount > 0 ? `${idx + 1} / ${this.slideCount}` : '';
  },

  /** 重置 */
  reset() {
    this.currentPptId = null;
    this.slideCount = 0;
    document.getElementById('thumbnail-list').innerHTML = '';
    document.getElementById('btn-prev').disabled = true;
    document.getElementById('btn-next').disabled = true;
    document.getElementById('slide-indicator').textContent = '';
  },
};

// ── 键盘快捷键 ────────────────────────────────
document.addEventListener('keydown', (e) => {
  // 不在输入框内时才响应
  if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;

  if (e.key === 'ArrowRight' || e.key === 'ArrowDown') {
    e.preventDefault();
    navigation.next();
  } else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp') {
    e.preventDefault();
    navigation.prev();
  } else if (e.key === 'Escape') {
    // 关闭各种弹窗
    document.getElementById('vision-modal').classList.add('hidden');
    document.getElementById('result-popup').classList.add('hidden');
    document.getElementById('highlight-picker').classList.add('hidden');
  }
});

// ── 按钮绑定 ───────────────────────────────────
document.getElementById('btn-prev').addEventListener('click', () => navigation.prev());
document.getElementById('btn-next').addEventListener('click', () => navigation.next());

// ── 监听幻灯片渲染完成，更新导航状态 ──────────
window.addEventListener('slide-rendered', (e) => {
  const { slideIdx, totalSlides } = e.detail;
  navigation.slideCount = totalSlides;
  navigation._setActiveThumbnail(slideIdx);
  navigation._updateControls(slideIdx);
});
