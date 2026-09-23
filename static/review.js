// Руководитель выделяет фрагмент работы и пишет замечание.
const box = document.getElementById('text'), info = document.getElementById('selInfo');
const field = document.getElementById('commentText'), addButton = document.getElementById('addComment');
const hint = info.textContent;
let sel = null;

// Позиции считаются по символам видимого текста, подписи замечаний рисуются через CSS и в счёт не идут.
document.addEventListener('selectionchange', () => {
  if (document.activeElement.matches('input, textarea')) return;
  const s = getSelection();
  if (!s.rangeCount || s.isCollapsed) return;
  const r = s.getRangeAt(0);
  if (!box.contains(r.commonAncestorContainer)) { sel = null; info.textContent = hint; return; }
  const pre = document.createRange();
  pre.selectNodeContents(box);
  pre.setEnd(r.startContainer, r.startOffset);
  const start = pre.toString().length, quote = r.toString();
  sel = {start, end: start + quote.length};
  info.textContent = '«' + (quote.length > 60 ? quote.slice(0, 60) + '…' : quote) + '»';
});

addButton.addEventListener('mousedown', e => e.preventDefault());  // не терять выделение
addButton.addEventListener('click', async () => {
  if (!sel) return info.textContent = T.select;
  if (!field.value.trim()) return info.textContent = T.needText;
  const r = await fetch(`/api/works/${WORK_ID}/comments`, {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({...sel, text: field.value}),
  });
  if (r.ok) location.reload(); else info.textContent = T.failed;
});

// Клик по замечанию в тексте показывает его в списке справа.
box.addEventListener('click', e => {
  const id = e.target.dataset.mark;
  if (!id || !getSelection().isCollapsed) return;
  const li = document.getElementById('m' + id);
  if (!li) return;
  li.scrollIntoView({block: 'nearest', behavior: 'smooth'});
  li.classList.add('flash');
  setTimeout(() => li.classList.remove('flash'), 1200);
});
