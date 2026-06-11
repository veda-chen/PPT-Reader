/**
 * LLM 功能面板 — 翻译和总结按钮事件。
 */
const llmPanel = {
  /** 翻译选中文本 */
  async translateSelection() {
    const text = window.getSelection()?.toString()?.trim();
    if (!text) {
      showToast('请先在幻灯片中选中要翻译的文本', 'error');
      return;
    }

    showPopup('🌐 翻译中...', '<div class="spinner"></div>');
    try {
      const resp = await api.translate(text);
      showPopup('🌐 翻译结果', _escapeHtml(resp.translated_text));
    } catch (e) {
      showPopup('翻译失败', e.message);
    }
  },

  /** 总结当前幻灯片 */
  async summarizeSlide() {
    const text = slideRenderer.getCurrentSlideText();
    if (!text) {
      showToast('当前幻灯片无文本内容', 'error');
      return;
    }

    showPopup('📝 总结中...', '<div class="spinner"></div>');
    try {
      const resp = await api.summarize(text, 'slide', 'brief');
      showPopup('📝 本页总结', _escapeHtml(resp.summary));
    } catch (e) {
      showPopup('总结失败', e.message);
    }
  },

  /** 总结全文 */
  async summarizeDocument() {
    if (!slideRenderer.currentPptId) {
      showToast('请先上传 PPT 文件', 'error');
      return;
    }

    showPopup('📄 全文总结中...', '<div class="spinner"></div>');
    try {
      const data = await api.getFullText(slideRenderer.currentPptId);
      const resp = await api.summarize(data.full_text, 'document', 'detailed');
      showPopup('📄 全文总结', _escapeHtml(resp.summary));
    } catch (e) {
      showPopup('总结失败', e.message);
    }
  },
};

// ── 按钮绑定 ──────────────────────────────────────────────
document.getElementById('btn-translate-text').addEventListener('click', () => llmPanel.translateSelection());
document.getElementById('btn-summarize-slide').addEventListener('click', () => llmPanel.summarizeSlide());
document.getElementById('btn-summarize-doc').addEventListener('click', () => llmPanel.summarizeDocument());

// ── Popup 辅助 ────────────────────────────────────────────
function showPopup(title, bodyHtml) {
  const popup = document.getElementById('result-popup');
  document.getElementById('result-popup-title').textContent = title;
  document.getElementById('result-popup-body').innerHTML = bodyHtml;
  popup.classList.remove('hidden');
}

document.querySelector('.result-popup-close').addEventListener('click', () => {
  document.getElementById('result-popup').classList.add('hidden');
});

// 点击背景关闭
document.getElementById('result-popup').addEventListener('click', (e) => {
  if (e.target === e.currentTarget) {
    e.currentTarget.classList.add('hidden');
  }
});

function _escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}
