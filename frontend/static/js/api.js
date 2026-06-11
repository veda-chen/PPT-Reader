/**
 * API 层 — 所有后端请求的 fetch() 封装。
 */
const api = {
  // ── PPT ──────────────────────────────────────
  async uploadPPT(file) {
    const formData = new FormData();
    formData.append('file', file);
    const res = await fetch('/api/upload', { method: 'POST', body: formData });
    if (!res.ok) throw new Error((await res.json()).detail || '上传失败');
    return res.json();
  },

  async getPPT(pptId) {
    const res = await fetch(`/api/ppt/${pptId}`);
    if (!res.ok) throw new Error('获取 PPT 信息失败');
    return res.json();
  },

  async getSlide(pptId, slideIdx) {
    const res = await fetch(`/api/ppt/${pptId}/slides/${slideIdx}`);
    if (!res.ok) throw new Error('获取幻灯片失败');
    return res.text();
  },

  async getFullText(pptId) {
    const res = await fetch(`/api/ppt/${pptId}/text`);
    if (!res.ok) throw new Error('获取文本失败');
    return res.json();
  },

  // ── Highlights ───────────────────────────────
  async createHighlight(data) {
    const res = await fetch('/api/highlights', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data),
    });
    if (!res.ok) throw new Error('创建高亮失败');
    return res.json();
  },

  async getHighlights(pptId, slideIdx) {
    const res = await fetch(`/api/highlights/${pptId}/${slideIdx}`);
    if (!res.ok) return [];
    return res.json();
  },

  async updateHighlight(id, data) {
    const res = await fetch(`/api/highlights/${id}`, {
      method: 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data),
    });
    if (!res.ok) throw new Error('更新高亮失败');
    return res.json();
  },

  async deleteHighlight(id) {
    const res = await fetch(`/api/highlights/${id}`, { method: 'DELETE' });
    if (!res.ok) throw new Error('删除高亮失败');
    return res.json();
  },

  // ── LLM ──────────────────────────────────────
  async translate(text, targetLang = 'zh') {
    const res = await fetch('/api/llm/translate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ text, target_lang: targetLang }),
    });
    if (!res.ok) throw new Error('翻译失败');
    return res.json();
  },

  async summarize(text, scope = 'slide', style = 'brief') {
    const res = await fetch('/api/llm/summarize', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ text, scope, style }),
    });
    if (!res.ok) throw new Error('总结失败');
    return res.json();
  },

  async chat(pptId, message) {
    const res = await fetch('/api/chat', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ppt_id: pptId, message }),
    });
    if (!res.ok) throw new Error('对话失败');
    return res.json();
  },

  async getChatHistory(pptId) {
    const res = await fetch(`/api/chat/${pptId}/history`);
    if (!res.ok) return [];
    return res.json();
  },

  async clearChatHistory(pptId) {
    const res = await fetch(`/api/chat/${pptId}/history`, { method: 'DELETE' });
    return res.json();
  },

  // ── Vision ───────────────────────────────────
  async describeImage(pptId, imagePath) {
    const res = await fetch('/api/vision/describe', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ ppt_id: pptId, image_path: imagePath }),
    });
    if (!res.ok) throw new Error('识图失败');
    return res.json();
  },
};
