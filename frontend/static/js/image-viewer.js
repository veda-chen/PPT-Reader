/**
 * 图片识图模块 — 点击幻灯片中的图片 → 调用豆包模型识图。
 */
const imageViewer = {
  /** 初始化：绑定幻灯片区域的图片点击事件 */
  init() {
    document.getElementById('slide-container').addEventListener('click', (e) => {
      const img = e.target.closest('img[data-image-path]');
      if (!img) return;

      e.stopPropagation();
      this._describeImage(img);
    });
  },

  /** 调用识图 API 并显示结果 */
  async _describeImage(img) {
    if (!slideRenderer.currentPptId) return;

    const imagePath = img.dataset.imagePath;
    const modal = document.getElementById('vision-modal');
    const imageEl = document.getElementById('vision-image');
    const resultEl = document.getElementById('vision-result');

    // 显示模态框
    modal.classList.remove('hidden');
    imageEl.src = img.src || '';
    imageEl.alt = '加载中...';
    resultEl.innerHTML = '<div class="spinner"></div><span>豆包正在识别图片...</span>';

    // 先尝试找到正确的图片 URL
    // 图片可能在 slides HTML 中还没有 src（渲染时只设了 data-image-path）
    if (!img.src || img.src.endsWith('/undefined')) {
      // 尝试加载实际图片
      try {
        const slideHtml = document.querySelector('.slide-wrapper')?.innerHTML || '';
        // 尝试查找匹配的图片文件
        const match = slideHtml.match(new RegExp(`src="([^"]*${imagePath}[^"]*)"`, 'i'));
        if (match) {
          imageEl.src = match[1];
        }
      } catch (e) { /* ignore */ }
    }

    try {
      const resp = await api.describeImage(slideRenderer.currentPptId, imagePath);
      if (resp.description.startsWith('未找到图片文件')) {
        resultEl.innerHTML = `<span style="color:#999;">⚠️ ${resp.description}</span>`;
      } else {
        resultEl.innerHTML = `<p style="line-height:1.8;">${_escapeHtml(resp.description)}</p>`;
      }
    } catch (e) {
      resultEl.innerHTML = `<span style="color:#d32f2f;">识图出错: ${e.message}</span>`;
    }
  },
};

// ── 初始化 ────────────────────────────────────────────────
imageViewer.init();

// ── 模态框关闭按钮 ────────────────────────────────────────
document.querySelector('#vision-modal .modal-close').addEventListener('click', () => {
  document.getElementById('vision-modal').classList.add('hidden');
});
document.getElementById('vision-modal').addEventListener('click', (e) => {
  if (e.target.classList.contains('modal-backdrop')) {
    e.currentTarget.classList.add('hidden');
  }
});

function _escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}
