/**
 * 对话面板 — 文档 Q&A 聊天界面。
 */
const chatPanel = {
  /** 发送消息 */
  async sendMessage() {
    const input = document.getElementById('chat-input');
    const message = input.value.trim();
    if (!message) return;
    if (!slideRenderer.currentPptId) {
      showToast('请先上传 PPT 文件', 'error');
      return;
    }

    const btn = document.getElementById('btn-send');
    input.disabled = true;
    btn.disabled = true;

    // 添加用户消息到界面
    this._appendMessage('user', message);
    input.value = '';

    // 添加加载指示
    const loadingId = this._appendLoading();

    try {
      const resp = await api.chat(slideRenderer.currentPptId, message);
      this._removeLoading(loadingId);
      this._appendMessage('assistant', resp.reply);
    } catch (e) {
      this._removeLoading(loadingId);
      this._appendMessage('assistant', '对话出错: ' + e.message);
    }

    input.disabled = false;
    btn.disabled = false;
    input.focus();
  },

  /** 加载历史消息 */
  async loadHistory() {
    const container = document.getElementById('chat-messages');
    container.innerHTML = '';

    if (!slideRenderer.currentPptId) {
      container.innerHTML = '<div class="chat-empty">开始与文档对话，提问关于 PPT 内容的任何问题</div>';
      return;
    }

    try {
      const messages = await api.getChatHistory(slideRenderer.currentPptId);
      if (messages.length === 0) {
        container.innerHTML = '<div class="chat-empty">开始与文档对话，提问关于 PPT 内容的任何问题</div>';
        return;
      }
      messages.forEach(m => this._appendMessage(m.role, m.content));
    } catch (e) {
      container.innerHTML = '<div class="chat-empty">开始与文档对话，提问关于 PPT 内容的任何问题</div>';
    }
  },

  /** 清空历史 */
  async clearHistory() {
    if (!slideRenderer.currentPptId) return;
    if (!confirm('确定清空对话历史？')) return;
    await api.clearChatHistory(slideRenderer.currentPptId);
    document.getElementById('chat-messages').innerHTML =
      '<div class="chat-empty">对话已清空。开始提问吧</div>';
    showToast('对话历史已清空', 'success');
  },

  /** 添加消息到界面 */
  _appendMessage(role, content) {
    const container = document.getElementById('chat-messages');
    // 移除空状态
    const empty = container.querySelector('.chat-empty');
    if (empty) empty.remove();

    const div = document.createElement('div');
    div.className = `chat-msg ${role}`;
    div.textContent = content;
    container.appendChild(div);
    container.scrollTop = container.scrollHeight;
  },

  /** 添加加载动画 */
  _appendLoading() {
    const container = document.getElementById('chat-messages');
    const div = document.createElement('div');
    div.className = 'chat-msg assistant';
    div.innerHTML = '<div class="spinner"></div>';
    div.id = 'chat-loading';
    container.appendChild(div);
    container.scrollTop = container.scrollHeight;
    return 'chat-loading';
  },

  /** 移除加载动画 */
  _removeLoading(id) {
    const el = document.getElementById(id);
    if (el) el.remove();
  },
};

// ── 事件绑定 ──────────────────────────────────────────────
document.getElementById('btn-send').addEventListener('click', () => chatPanel.sendMessage());
document.getElementById('chat-input').addEventListener('keydown', (e) => {
  if (e.key === 'Enter' && !e.shiftKey) {
    e.preventDefault();
    chatPanel.sendMessage();
  }
});

// 切换 PPT 时重新加载对话历史
window.addEventListener('ppt-loaded', () => chatPanel.loadHistory());
