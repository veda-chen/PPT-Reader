/**
 * 高亮系统 — 文本选区捕获、高亮创建/恢复/删除。
 *
 * 核心流程:
 *   选中文本 → mouseup → 显示颜色选择器 → 选颜色 → 计算 segments → POST API → 包裹 <mark>
 *   恢复: 加载幻灯片后 → GET highlights → 按 segments 查找 DOM → Range.surroundContents(<mark>)
 */
const highlight = {
  selectedColor: '#FFEB3B',
  pendingSelection: null,  // 最近一次选区的捕获结果
  activeHighlightId: null, // 当前被点击的高亮 ID（用于编辑/删除）

  /** 在幻灯片渲染后恢复所有高亮 */
  async restoreAll(pptId, slideIdx) {
    try {
      const highlights = await api.getHighlights(pptId, slideIdx);
      for (const h of highlights) {
        this._renderSingle(h);
      }
    } catch (e) {
      console.warn('恢复高亮失败:', e);
    }
  },

  /** 渲染单个高亮到 DOM */
  _renderSingle(h) {
    let segments;
    try {
      segments = JSON.parse(h.segments_json);
    } catch { return; }

    const container = document.getElementById('slide-container');

    for (const seg of segments) {
      const span = container.querySelector(
        `[data-shape-id="${seg.shape_id}"][data-para-idx="${seg.para_idx}"][data-run-idx="${seg.run_idx}"]`
      );
      if (!span) continue;

      const textNode = span.firstChild;
      if (!textNode || textNode.nodeType !== Node.TEXT_NODE) continue;

      const charStart = Math.min(seg.char_start, textNode.length);
      const charEnd = Math.min(seg.char_end, textNode.length);
      if (charStart >= charEnd) continue;

      try {
        const range = document.createRange();
        range.setStart(textNode, charStart);
        range.setEnd(textNode, charEnd);

        const mark = document.createElement('mark');
        mark.className = 'ppt-highlight';
        if (h.note) mark.classList.add('has-note');
        mark.style.backgroundColor = h.color;
        mark.dataset.highlightId = h.id;
        if (h.note) mark.title = h.note;

        range.surroundContents(mark);
      } catch (e) {
        // fallback: extractContents + insertNode
        try {
          const range = document.createRange();
          range.setStart(textNode, charStart);
          range.setEnd(textNode, charEnd);
          const mark = document.createElement('mark');
          mark.className = 'ppt-highlight';
          if (h.note) mark.classList.add('has-note');
          mark.style.backgroundColor = h.color;
          mark.dataset.highlightId = h.id;
          mark.appendChild(range.extractContents());
          range.insertNode(mark);
        } catch (e2) {
          console.warn('高亮恢复失败:', e2);
        }
      }
    }
  },

  /** 捕获当前文本选区，返回 segments 信息 */
  captureSelection() {
    const sel = window.getSelection();
    if (!sel.rangeCount || sel.isCollapsed) return null;

    const range = sel.getRangeAt(0);
    const container = document.getElementById('slide-container');
    if (!container || !container.contains(range.commonAncestorContainer)) return null;

    if (!sel.toString().trim()) return null;

    // 遍历选区内的所有文本节点
    const walker = document.createTreeWalker(
      container.querySelector('.slide-wrapper') || container,
      NodeFilter.SHOW_TEXT,
      null,
      false
    );

    const segments = [];
    const textParts = [];   // 用于重建可读文本
    let node = walker.nextNode();

    while (node) {
      if (range.intersectsNode(node)) {
        const runSpan = node.parentElement?.closest?.('[data-run-idx]');
        if (!runSpan) { node = walker.nextNode(); continue; }

        const shapeId = parseInt(runSpan.dataset.shapeId);
        const paraIdx = parseInt(runSpan.dataset.paraIdx);
        const runIdx = parseInt(runSpan.dataset.runIdx);
        if (isNaN(shapeId) || isNaN(paraIdx) || isNaN(runIdx)) {
          node = walker.nextNode(); continue;
        }

        const isStart = (node === range.startContainer);
        const isEnd = (node === range.endContainer);
        const charStart = isStart ? range.startOffset : 0;
        const charEnd = isEnd ? range.endOffset : node.textContent.length;

        if (charStart < charEnd) {
          segments.push({
            shape_id: shapeId,
            para_idx: paraIdx,
            run_idx: runIdx,
            char_start: charStart,
            char_end: charEnd,
          });

          // 重建文本：不同文本框/不同行（shape_id 或 para_idx 变化）之间插入空格分隔，
          // 同一行内的多个 run 直接拼接（高保真模式下各行是独立绝对定位 span，
          // 浏览器原生 toString 不会自动加分隔，会把单词连在一起）
          const sub = node.textContent.substring(charStart, charEnd);
          // 不同 span 之间智能插入空格分隔，避免 PyMuPDF 提取的 span 无尾部空格时单词粘连
          if (textParts.length > 0) {
            const prevEnd = textParts[textParts.length - 1];
            // 上一个不以空格/换行结尾，且当前不以空格/换行开头 → 需插入空格
            if (!(/[ \n]$/.test(prevEnd)) && !(/^[ \n]/.test(sub))) {
              textParts.push(' ');
            }
          }
          textParts.push(sub);
        }
      }
      node = walker.nextNode();
    }

    if (segments.length === 0) return null;

    // 合并多余空白，得到正常句式
    const text = textParts.join('').replace(/\s+/g, ' ').trim();
    return { text, segments };
  },

  /** 创建高亮并渲染 */
  async createHighlight(color) {
    if (!this.pendingSelection) return null;
    const { text, segments } = this.pendingSelection;

    try {
      const h = await api.createHighlight({
        ppt_id: slideRenderer.currentPptId,
        slide_idx: slideRenderer.currentSlideIdx,
        highlighted_text: text,
        segments: segments,
        color: color || this.selectedColor,
      });
      // 清除选区后渲染
      window.getSelection().removeAllRanges();
      this._renderSingle(h);
      this.pendingSelection = null;
      return h;
    } catch (e) {
      console.error('创建高亮失败:', e);
      showToast('创建高亮失败: ' + e.message, 'error');
      return null;
    }
  },

  /** 更新高亮笔记 */
  async updateNote(highlightId, note) {
    try {
      return await api.updateHighlight(highlightId, { note });
    } catch (e) {
      showToast('更新笔记失败: ' + e.message, 'error');
      return null;
    }
  },

  /** 删除高亮 */
  async deleteHighlight(highlightId) {
    try {
      await api.deleteHighlight(highlightId);
      // 从 DOM 移除对应的 <mark> 元素
      document.querySelectorAll(`mark[data-highlight-id="${highlightId}"]`).forEach(m => {
        const parent = m.parentNode;
        while (m.firstChild) {
          parent.insertBefore(m.firstChild, m);
        }
        parent.removeChild(m);
      });
      document.getElementById('highlight-picker').classList.add('hidden');
    } catch (e) {
      showToast('删除高亮失败: ' + e.message, 'error');
    }
  },

  /** 刷新笔记面板 */
  async refreshNotesPanel() {
    if (!slideRenderer.currentPptId) return;
    try {
      const highlights = await api.getHighlights(slideRenderer.currentPptId, slideRenderer.currentSlideIdx);
      const withNotes = highlights.filter(h => h.note);
      _renderNotesList(withNotes);
    } catch (e) {
      console.warn('刷新笔记失败:', e);
    }
  },
};

// ── 全局: mouseup 事件 ──────────────────────────────────────
document.addEventListener('mouseup', (e) => {
  // 延迟执行，确保 selection 已稳定
  setTimeout(() => {
    const container = document.getElementById('slide-container');
    if (!container) return;

    const sel = window.getSelection();
    if (!sel || sel.isCollapsed) return;
    if (!container.contains(sel.anchorNode)) return;

    const result = highlight.captureSelection();
    if (!result) return;

    highlight.pendingSelection = result;

    // 显示浮动颜色选择器
    const picker = document.getElementById('highlight-picker');
    const range = sel.getRangeAt(0);
    const rect = range.getBoundingClientRect();

    picker.style.left = `${rect.left + rect.width / 2 - picker.offsetWidth / 2}px`;
    picker.style.top = `${rect.bottom + 8}px`;
    picker.classList.remove('hidden');
    highlight.activeHighlightId = null;
  }, 10);
});

// ── 高亮颜色选择器按钮 ──────────────────────────────────────
document.getElementById('highlight-picker').addEventListener('click', async (e) => {
  const btn = e.target.closest('button');
  if (!btn) return;

  const picker = document.getElementById('highlight-picker');

  if (btn.dataset.color) {
    highlight.selectedColor = btn.dataset.color;
    await highlight.createHighlight(btn.dataset.color);
    picker.classList.add('hidden');
    highlight.refreshNotesPanel();
  } else if (btn.dataset.action === 'note') {
    // 对当前活跃高亮添加笔记
    if (highlight.activeHighlightId) {
      const note = prompt('请输入笔记内容:');
      if (note !== null) {
        await highlight.updateNote(highlight.activeHighlightId, note);
        // 更新 DOM
        document.querySelectorAll(`mark[data-highlight-id="${highlight.activeHighlightId}"]`).forEach(m => {
          m.title = note;
          if (note) m.classList.add('has-note');
          else m.classList.remove('has-note');
        });
        highlight.refreshNotesPanel();
      }
    } else if (highlight.pendingSelection) {
      // 先创建高亮，再添加笔记
      const h = await highlight.createHighlight(highlight.selectedColor);
      if (h) {
        const note = prompt('请输入笔记内容:');
        if (note !== null) {
          await highlight.updateNote(h.id, note);
          document.querySelectorAll(`mark[data-highlight-id="${h.id}"]`).forEach(m => {
            m.title = note;
            m.classList.add('has-note');
          });
          highlight.refreshNotesPanel();
        }
      }
    }
    picker.classList.add('hidden');
  } else if (btn.dataset.action === 'delete') {
    if (highlight.activeHighlightId) {
      if (confirm('确定删除此高亮？')) {
        await highlight.deleteHighlight(highlight.activeHighlightId);
        highlight.refreshNotesPanel();
      }
    }
    picker.classList.add('hidden');
  }
});

// ── 点击已有高亮标记 → 显示编辑选项 ────────────────────────
document.getElementById('slide-container').addEventListener('click', (e) => {
  const mark = e.target.closest('mark.ppt-highlight');
  if (!mark) {
    // 点击非高亮区域 → 隐藏选择器
    if (!window.getSelection()?.isCollapsed) return;
    document.getElementById('highlight-picker').classList.add('hidden');
    return;
  }

  e.stopPropagation();
  const highlightId = parseInt(mark.dataset.highlightId);
  if (isNaN(highlightId)) return;

  highlight.activeHighlightId = highlightId;

  const picker = document.getElementById('highlight-picker');
  const rect = mark.getBoundingClientRect();
  picker.style.left = `${rect.left + rect.width / 2 - picker.offsetWidth / 2}px`;
  picker.style.top = `${rect.bottom + 8}px`;
  picker.classList.remove('hidden');

  // 更新当前颜色
  highlight.selectedColor = mark.style.backgroundColor || '#FFEB3B';
});

// ── 幻灯片渲染后恢复高亮 ──────────────────────────────────
window.addEventListener('slide-rendered', (e) => {
  const { pptId, slideIdx } = e.detail;
  highlight.restoreAll(pptId, slideIdx);
  highlight.refreshNotesPanel();
});

// ── 渲染笔记列表 ───────────────────────────────────────────
function _renderNotesList(highlights) {
  const notesList = document.getElementById('notes-list');
  if (highlights.length === 0) {
    notesList.innerHTML = '<div class="notes-empty">暂无笔记。选中幻灯片文字创建高亮笔记</div>';
    return;
  }
  notesList.innerHTML = highlights.map(h => `
    <div class="note-item" data-highlight-id="${h.id}" style="border-left-color:${h.color}">
      <div class="note-highlight-preview" style="background:${h.color}33">${_escapeHtml(h.highlighted_text)}</div>
      ${h.note ? `<div class="note-content">${_escapeHtml(h.note)}</div>` : ''}
      <div class="note-meta">
        <span>页 ${h.slide_idx + 1}</span>
        <span>${h.updated_at?.substring(0, 16) || ''}</span>
      </div>
    </div>
  `).join('');

  // 点击笔记条目 → 跳转到对应页
  notesList.querySelectorAll('.note-item').forEach(item => {
    item.addEventListener('click', () => {
      const hId = parseInt(item.dataset.highlightId);
      const h = highlights.find(x => x.id === hId);
      if (h && h.slide_idx !== slideRenderer.currentSlideIdx) {
        navigation.goToSlide(h.slide_idx);
      }
    });
  });
}

function _escapeHtml(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}
