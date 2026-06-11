/**
 * 主控制器 — 应用启动、状态管理、模块编排。
 */
const app = {
  currentPptId: null,
  currentSlideIdx: 0,

  /** 初始化 */
  init() {
    this._bindUpload();
    this._bindSidebarTabs();
    this._bindEmptyStateUpload();

    // 初始状态：禁用工具栏
    this._setToolbarEnabled(false);

    console.log('📊 智能PPT阅读器已就绪');
    showStatus('就绪');
  },

  /** 上传 PPT 文件 */
  async uploadPPT(file) {
    showStatus('正在上传和解析 PPT...');
    this._setToolbarEnabled(false);

    try {
      const info = await api.uploadPPT(file);

      this.currentPptId = info.id;
      this.currentSlideIdx = 0;

      // 清除旧缓存
      slideRenderer.clearCache();
      bilingual.clearCache();

      // 更新 UI
      document.getElementById('header-filename').textContent = info.original_filename;
      document.getElementById('empty-state')?.remove();
      this._setToolbarEnabled(true);

      // 初始化导航
      slideRenderer.setTotalSlides(info.slide_count);
      navigation.init(info.id, info.slide_count);

      // 加载第一页
      await slideRenderer.loadSlide(info.id, 0);

      // 加载对话历史
      chatPanel.loadHistory();

      // 触发 PPT 加载事件
      window.dispatchEvent(new CustomEvent('ppt-loaded', {
        detail: { pptId: info.id, slideCount: info.slide_count }
      }));

      showToast(`已加载: ${info.original_filename} (${info.slide_count} 页)`, 'success');
      showStatus(`已加载: ${info.original_filename}`);

    } catch (e) {
      showToast('上传失败: ' + e.message, 'error');
      showStatus('上传失败');
    }
  },

  /** 启用/禁用工具栏按钮 */
  _setToolbarEnabled(enabled) {
    ['btn-bilingual', 'btn-translate-text', 'btn-summarize-slide', 'btn-summarize-doc'].forEach(id => {
      document.getElementById(id).disabled = !enabled;
    });
  },

  /** 绑定上传按钮 */
  _bindUpload() {
    const fileInput = document.getElementById('file-input');
    const uploadBtn = document.getElementById('btn-upload');

    uploadBtn.addEventListener('click', () => fileInput.click());
    fileInput.addEventListener('change', (e) => {
      const file = e.target.files[0];
      if (file) {
        this.uploadPPT(file);
        fileInput.value = ''; // 允许重复上传同一文件
      }
    });
  },

  /** 绑定空状态页的上传按钮 */
  _bindEmptyStateUpload() {
    const largeBtn = document.getElementById('btn-upload-large');
    if (largeBtn) {
      largeBtn.addEventListener('click', () => document.getElementById('file-input').click());
    }
  },

  /** 绑定侧边栏 Tab 切换 */
  _bindSidebarTabs() {
    document.querySelectorAll('.sidebar-tab').forEach(tab => {
      tab.addEventListener('click', () => {
        // 切换 active tab 样式
        document.querySelectorAll('.sidebar-tab').forEach(t => t.classList.remove('active'));
        tab.classList.add('active');

        // 切换面板
        const target = tab.dataset.tab;
        document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
        document.getElementById(`tab-${target}`).classList.add('active');
      });
    });
  },
};

// ── Toast 通知 ─────────────────────────────────────────────
function showToast(msg, type = '') {
  const toast = document.getElementById('toast');
  toast.textContent = msg;
  toast.className = `toast ${type}`;
  toast.classList.remove('hidden');

  clearTimeout(toast._timeout);
  toast._timeout = setTimeout(() => {
    toast.classList.add('hidden');
  }, 3000);
}

// ── 状态栏 ─────────────────────────────────────────────────
function showStatus(msg) {
  document.getElementById('status-text').textContent = msg;
}

// ── 启动 ───────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => app.init());
