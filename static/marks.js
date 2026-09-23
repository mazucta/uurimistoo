// Раскраска поверх оформленного текста: что вставлено и где пометки об ошибках.
// Позиции считаются по символам textContent, ровно так же, как их считал редактор ученика.
(() => {
  const box = document.getElementById('text');
  if (!box) return;
  const text = box.textContent;
  const note = document.getElementById('paintWarning');
  if (typeof ORIGINS === 'string' && ORIGINS.length !== text.length) {
    if (note) note.hidden = false;  // текст и лог разошлись, лучше ничего не красить
    return;
  }

  const cls = new Array(text.length).fill('');
  for (let i = 0; i < ORIGINS.length;) {
    let j = i;
    while (j < ORIGINS.length && ORIGINS[j] === ORIGINS[i]) j++;
    const kind = ORIGINS[i] === 'p' ? (j - i >= 150 ? 'paste-big' : 'paste') : ORIGINS[i] === '?' ? 'unk' : '';
    if (kind) cls.fill(kind, i, j);
    i = j;
  }

  const markAt = new Array(text.length).fill(null);
  // замечание руководителя рисуется поверх предложения ИИ
  for (const m of [...MARKS].sort((a, b) => (a.status !== 'suggested') - (b.status !== 'suggested'))) {
    markAt.fill(m, Math.max(0, m.start), Math.min(text.length, m.end));
  }

  const nodes = [];
  const walker = document.createTreeWalker(box, NodeFilter.SHOW_TEXT);
  for (let pos = 0; walker.nextNode(); pos += walker.currentNode.data.length) nodes.push([walker.currentNode, pos]);

  for (const [node, offset] of nodes.reverse()) {
    const parts = [];
    for (let start = 0, i = 1; i <= node.data.length; i++) {
      const border = i === node.data.length
        || cls[offset + i] !== cls[offset + start] || markAt[offset + i] !== markAt[offset + start];
      if (border) {
        parts.push([start, i, cls[offset + start], markAt[offset + start]]);
        start = i;
      }
    }
    for (const [from, to, kind, mark] of parts.reverse()) {
      if (!kind && !mark) continue;
      if (to < node.data.length) node.splitText(to);
      const piece = from ? node.splitText(from) : node;
      const span = document.createElement('span');
      span.className = [kind, mark && 'mark', mark && mark.status === 'suggested' && 'ai'].filter(Boolean).join(' ');
      if (mark) {
        span.dataset.mark = mark.id;
        span.style.setProperty('--c', mark.color);
        span.title = (mark.status === 'suggested' ? T.suggests + ' ' : '') + (mark.kind ? mark.kind + ': ' : '') + mark.text;
      }
      piece.parentNode.replaceChild(span, piece);
      span.append(piece);
    }
  }
})();
