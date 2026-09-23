// Рабочее место ученика: текст работы, структура, объём и источники.
const rich = document.getElementById('editor');
const status = document.getElementById('status'), count = document.getElementById('count');
let queue = [], chain = Promise.resolve(), last = '', pendingType = null, caret = null, dirty = false;

const log = (type, extra) => queue.push({t: Date.now(), type, ...extra});
const api = (path, body) => fetch(`/api/works/${WORK.id}${path}`, {
  method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify(body),
  keepalive: document.visibilityState === 'hidden',  // ponytail: keepalive ограничен 64 КБ, огромная вставка перед закрытием вкладки может не уйти
});

// Разница старого и нового текста: общий префикс и суффикс. Ловит любой ввод: IME, автозамену, undo.
function diff(a, b) {
  const max = Math.min(a.length, b.length);
  let p = 0; while (p < max && a[p] === b[p]) p++;
  let s = 0; while (s < max - p && a[a.length - 1 - s] === b[b.length - 1 - s]) s++;
  return {pos: p, del: a.slice(p, a.length - s), ins: b.slice(p, b.length - s)};
}

// ---------- объём и структура ----------

function refreshSidebar() {
  const text = rich.textContent;
  count.textContent = `${(text.match(/\S+/g) || []).length} ${T.words} · ${text.length} ${T.chars}`;
  const share = Math.min(100, Math.round(100 * text.length / WORK.target));
  document.getElementById('volume').textContent = `${share}% ${T.ofTarget}`;
  document.getElementById('volumeBar').style.width = share + '%';
  document.getElementById('volumeHint').textContent = text.length >= WORK.target
    ? T.enough : `${T.left} ${WORK.target - text.length} ${T.chars}`;

  // Номера как в Word: введение, выводы, список источников, приложения и подразделы под ними без номера.
  // То же правило на сервере в richtext.mark_headings.
  const headings = [...rich.querySelectorAll('h2, h3, h4')];
  const outline = document.getElementById('outline');
  outline.innerHTML = '';
  let numbered = false, apps = 0;
  const num = [0, 0, 0];
  headings.forEach((h, i) => {
    const level = +h.tagName[1] - 1, text = h.textContent.trim();
    let label = '';
    if (level === 1) {
      numbered = !h.classList.contains('appendix') && !WORK.unnumbered.some(k => text.toLowerCase().startsWith(k));
      if (h.classList.contains('appendix')) label = `${T.appendix} ${++apps}. `;
    }
    h.classList.toggle('nonum', !numbered && !h.classList.contains('appendix'));
    if (numbered) {
      num[level - 1]++;
      num.fill(0, level);
      label = num.slice(0, level).join('.') + ' ';
    }
    h.id = h.id || 'h' + i;
    const li = document.createElement('li');
    li.className = level > 1 ? 'sub' : '';
    const a = document.createElement('a');
    a.href = '#' + h.id;
    a.textContent = label + (text || '…');
    li.append(a);
    outline.append(li);
  });
  if (!headings.length) {
    const li = document.createElement('li');
    li.className = 'empty';
    li.textContent = T.noHeadings;
    outline.append(li);
  }

  const titles = headings.map(h => h.textContent.toLowerCase());
  const list = document.getElementById('checklist');
  list.innerHTML = '';
  for (const section of SECTIONS) {
    const found = titles.some(t => section.keys.some(k => t.includes(k)));
    const li = document.createElement('li');
    li.className = found ? 'on' : '';
    li.textContent = (found ? '✓ ' : '· ') + section.label;
    list.append(li);
  }
  // На каждое приложение должна быть ссылка в основном тексте.
  for (let n = 1; n <= apps; n++) {
    const found = new RegExp(`(приложен|додат|lisa)\\S*\\s+${n}(?!\\d)`, 'i').test(text);
    const li = document.createElement('li');
    li.className = found ? 'on' : '';
    li.textContent = `${found ? '✓' : '·'} ${T.appendix} ${n}: ${found ? T.appendixRef : T.appendixNoRef}`;
    list.append(li);
  }
}

// ---------- запись работы ----------

rich.innerHTML = WORK.html || '<p><br></p>';
last = rich.textContent;
refreshSidebar();

if (!WORK.locked) {
  document.execCommand('defaultParagraphSeparator', false, 'p');
  rich.addEventListener('beforeinput', e => pendingType = e.inputType);
  rich.addEventListener('input', () => {
    const now = rich.textContent, d = diff(last, now);
    if (d.del || d.ins) log(pendingType || 'insertText', d);
    last = now; pendingType = null; dirty = true;
    refreshSidebar();
  });
  rich.addEventListener('focus', () => log('focus'));
  rich.addEventListener('blur', () => log('blur'));
  log('resume', {ins: last});
}

const tools = [...document.querySelectorAll('.tool[data-cmd]')];
const refreshTools = () => tools.forEach(b => {
  try {
    b.classList.toggle('on', b.dataset.arg
      ? document.queryCommandValue('formatBlock').toLowerCase() === b.dataset.arg
      : document.queryCommandState(b.dataset.cmd));
  } catch { /* браузер не знает команду */ }
});
tools.forEach(b => {
  b.addEventListener('mousedown', e => e.preventDefault());  // не терять выделение
  b.addEventListener('click', () => {
    if (b.dataset.cmd === 'createLink') {
      const url = prompt(T.link, 'https://');
      if (url) document.execCommand('createLink', false, url);
    } else {
      document.execCommand(b.dataset.cmd, false, b.dataset.arg || null);
    }
    rich.focus();
    refreshTools();
  });
});

const restoreCaret = () => {
  rich.focus();
  if (caret) {
    const sel = getSelection();
    sel.removeAllRanges();
    sel.addRange(caret);
  }
};

// Приложение: заголовок раздела с классом appendix, номер «ПРИЛОЖЕНИЕ N» ставит оформление.
document.getElementById('appendixBtn')?.addEventListener('mousedown', e => e.preventDefault());
document.getElementById('appendixBtn')?.addEventListener('click', () => {
  restoreCaret();
  const node = getSelection().anchorNode;
  const h = (node?.nodeType === 1 ? node : node?.parentElement)?.closest('h2');
  if (h && rich.contains(h)) h.classList.toggle('appendix');
  else {
    document.execCommand('formatBlock', false, 'h2');
    const n = getSelection().anchorNode;
    (n?.nodeType === 1 ? n : n?.parentElement)?.closest('h2')?.classList.add('appendix');
  }
  dirty = true;
  refreshSidebar();
});

const photoFile = document.getElementById('photoFile');
document.getElementById('photoBtn')?.addEventListener('mousedown', e => e.preventDefault());
document.getElementById('photoBtn')?.addEventListener('click', () => photoFile.click());
photoFile?.addEventListener('change', async () => {
  const file = photoFile.files[0];
  photoFile.value = '';
  if (!file) return;
  const data = new FormData();
  data.append('file', file);
  const r = await fetch(`/api/works/${WORK.id}/images`, {method: 'POST', body: data});
  const res = await r.json().catch(() => ({}));
  if (!r.ok) return alert(res.error || T.photoFailed);
  restoreCaret();
  document.execCommand('insertHTML', false, `<figure><img src="${res.url}"><figcaption><br></figcaption></figure><p><br></p>`);
  dirty = true;
});

const cover = document.getElementById('coverForm');
cover?.addEventListener('change', () => api('/cover', Object.fromEntries(new FormData(cover))));

document.getElementById('docxBtn').addEventListener('click', async e => {
  if (WORK.locked) return;
  const href = e.currentTarget.href;  // сначала сохраняем последние правки, потом скачиваем
  e.preventDefault();
  await flush();
  location.href = href;
});
document.addEventListener('selectionchange', () => {
  const sel = getSelection();
  if (!sel.rangeCount || !rich.contains(sel.getRangeAt(0).commonAncestorContainer)) return;
  caret = sel.getRangeAt(0).cloneRange();
  refreshTools();
});

const title = document.getElementById('title');
title.addEventListener('change', () => api('/title', {title: title.value}));

// ---------- источники ----------

// ---------- условия к работе ----------

function renderChecks(checks) {
  CHECKS = checks;
  const box = document.getElementById('checks');
  box.innerHTML = '';
  for (const c of checks) {
    const li = document.createElement('li');
    li.className = c.status;
    const mark = document.createElement('span');
    mark.className = 'mark';
    mark.textContent = c.status === 'pass' ? '✓' : c.status === 'fail' ? '✗' : '·';
    const text = document.createElement('span');
    text.className = 'grow';
    text.textContent = c.text + (c.note ? ` ${c.note}` : (c.auto ? '' : ` — ${T.willCheck}`));
    li.append(mark, text);
    box.append(li);
  }
  const passed = checks.filter(c => c.status === 'pass').length;
  document.getElementById('checksDone').textContent = checks.length ? `${passed} / ${checks.length} ${T.checksDone}` : '';
}
renderChecks(CHECKS);

const list = document.getElementById('sourceList'), form = document.getElementById('sourceForm');

function cite(src) {
  const label = [src.author || src.title, src.year].filter(Boolean).join(', ');
  restoreCaret();
  document.execCommand('insertText', false, ` (${label})`);
}

function renderSources(sources) {
  list.innerHTML = '';
  for (const src of sources) {
    const li = document.createElement('li');
    const head = document.createElement('div');
    head.className = 'source-head';
    head.textContent = [src.author, src.title && `«${src.title}»`, src.year].filter(Boolean).join(' ');
    li.append(head);
    if (src.url) {
      const a = document.createElement('a');
      a.href = src.url; a.target = '_blank'; a.rel = 'noopener noreferrer'; a.className = 'small';
      a.textContent = src.url.replace(/^https?:\/\//, '').slice(0, 40);
      li.append(a);
    }
    if (!WORK.locked) {
      const actions = document.createElement('div');
      actions.className = 'source-actions';
      const insert = document.createElement('button');
      insert.type = 'button'; insert.className = 'btn small'; insert.textContent = T.cite;
      insert.onclick = () => cite(src);
      const drop = document.createElement('button');
      drop.type = 'button'; drop.className = 'btn-link danger'; drop.title = T.drop; drop.textContent = '✕';
      drop.onclick = async () => {
        const r = await fetch(`/api/sources/${src.id}/delete`, {
          method: 'POST', headers: {'Content-Type': 'application/json'}, body: '{}'});
        if (r.ok) renderSources((await r.json()).sources);
      };
      actions.append(insert, drop);
      li.append(actions);
    }
    list.append(li);
  }
  document.getElementById('sourceCount').textContent = sources.length || '';
}
renderSources(SOURCES);

form?.addEventListener('submit', async e => {
  e.preventDefault();
  const r = await api('/sources', Object.fromEntries(new FormData(form)));
  if (!r.ok) return alert(T.sourceFailed);
  renderSources((await r.json()).sources);
  form.reset();
});

// ---------- сохранение и сдача ----------

// Отправки идут строго по очереди, иначе события на сервере перемешаются.
const flush = () => chain = chain.then(async () => {
  if (!queue.length && !dirty) return;
  const events = queue; queue = []; dirty = false;
  try {
    const r = await api('/events', {events, html: rich.innerHTML});
    if (!r.ok) throw new Error(r.status);
    renderChecks((await r.json()).checks);
    status.textContent = `${T.saved} ${new Date().toTimeString().slice(0, 5)}`;
  } catch {
    queue = events.concat(queue); dirty = true;
    status.textContent = T.offline;
  }
});

if (!WORK.locked) {
  setInterval(flush, 10000);
  document.addEventListener('visibilitychange', () => document.hidden && flush());
  document.getElementById('submitBtn').onclick = async e => {
    if (!rich.textContent.trim()) return alert(T.empty);
    const warn = [title.value.trim() ? '' : T.noTitle, list.children.length ? '' : T.noSources].filter(Boolean).join(' ');
    if (!confirm((warn ? warn + ' ' : '') + T.confirm)) return;
    e.target.disabled = true;
    await flush();
    const r = queue.length ? null : await api('/submit', {html: rich.innerHTML});
    if (r?.ok) return location.reload();
    e.target.disabled = false;
    alert(T.failed);
  };
}
