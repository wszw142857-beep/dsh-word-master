/* Word Master 审阅界面 —— 前端逻辑。
 *
 * 交互骨架照 ppt-master 的 svg_editor（实时预览编辑器）：
 *   点画布上的元素 → 右侧显示「已选」→ 写修改说明 → 记下 → 提交
 *
 * 三个面板各有分工：
 *   左：段落目录（相当于它的幻灯片列表），带 P0/P1 标记
 *   中：正文纸面（相当于它的画布），点选段落
 *   右：标注面板（相当于它的 Annotations），chips + 自由文本
 *
 * 与它的差别只在实体：它改 SVG 元素属性，我们改 Markdown 段落内容 + 给 AI 的指令。
 */

const state = {
    blocks: [],
    issues: [],
    decisions: [],      // [{line, quote, verdicts, note}]
    catalogs: null,
    recommendations: {},
    selected: null,     // 当前选中的 block
    selection: [],      // 浮层里选中的 verdict ids
    dirty: new Set(),   // 已直接改过文字的段落行号
    history: [],        // 撤销栈（只记 decisions 的操作）
    originalDecisions: [],
    mtime: null,        // 上次渲染时的 draft.md mtime
    pollTimer: null,
    driftPending: false,
};

const POLL_MS = 2000;  // 对齐 ppt-master svg_editor 的轮询节奏

// ── 工具 ──────────────────────────────────────────────────────────────────

function toast(msg, ms = 2200) {
    const node = document.getElementById('toast');
    node.textContent = msg;
    node.classList.add('show');
    clearTimeout(node._t);
    node._t = setTimeout(() => node.classList.remove('show'), ms);
}

function esc(s) {
    return String(s).replace(/[&<>"']/g, (c) => ({
        '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;',
    }[c]));
}

function el(tag, cls, text) {
    const node = document.createElement(tag);
    if (cls) node.className = cls;
    if (text != null) node.textContent = text;
    return node;
}

function inline(raw) {
    return esc(raw)
        .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
        .replace(/(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)/g, '<em>$1</em>')
        .replace(/`(.+?)`/g, '<code>$1</code>');
}

async function post(url, body) {
    try {
        const res = await fetch(url, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(body || {}),
        });
        return await res.json();
    } catch (e) {
        return { ok: false, error: String(e) };
    }
}

function verdictLabel(id) {
    const found = (state.catalogs?.verdicts || []).find((x) => x.id === id);
    return found ? found.label_zh : id;
}

function decisionFor(line) {
    return state.decisions.find((d) => d.line === line);
}

function issuesFor(line) {
    return state.issues.filter((i) => i.line === line);
}

// ── 左栏：段落目录 ────────────────────────────────────────────────────────

function renderOutline() {
    const box = document.getElementById('outline');
    box.innerHTML = '';
    if (state.blocks.length === 0) {
        box.innerHTML = '<div class="outline-empty">没有正文。先写出 draft.md。</div>';
        return;
    }

    const heads = state.blocks.filter((b) => b.kind === 'heading');
    const source = heads.length > 0 ? heads : state.blocks;
    if (heads.length === 0) {
        box.appendChild(el('div', 'outline-empty', '（无标题，列出全部段落）'));
    }

    for (const b of source) {
        const lineIssues = issuesFor(b.line);
        const hasP0 = lineIssues.some((i) => i.severity === 'P0');
        const hasP1 = lineIssues.some((i) => i.severity === 'P1');
        const cls = ['outline-item'];
        if (b.level === 2) cls.push('h2');
        if (hasP0) cls.push('has-p0');
        else if (hasP1) cls.push('has-p1');
        if (state.selected && state.selected.line === b.line) cls.push('active');

        const item = el('div', cls.join(' '), b.text.slice(0, 26) || '(空)');
        item.title = b.text;
        item.addEventListener('click', () => selectBlock(b.line, true));
        box.appendChild(item);
    }
}

// ── 中栏：正文纸面 ────────────────────────────────────────────────────────

function renderDoc() {
    const content = document.getElementById('doc-content');
    const placeholder = document.getElementById('doc-placeholder');
    content.innerHTML = '';

    if (state.blocks.length === 0) {
        placeholder.textContent = '没有正文。先写出 draft.md。';
        placeholder.style.display = 'block';
        content.style.display = 'none';
        return;
    }
    placeholder.style.display = 'none';
    content.style.display = 'block';

    for (const b of state.blocks) {
        const div = el('div', 'block');
        div.dataset.line = b.line;
        div.dataset.kind = b.kind;

        const readonly = b.kind === 'math' || b.kind === 'table';
        div.contentEditable = readonly ? 'false' : 'true';
        div.spellcheck = false;

        if (b.kind === 'heading') {
            const tag = 'h' + Math.min(Math.max(b.level || 1, 1), 4);
            div.innerHTML = `<${tag}>${inline(b.raw.replace(/^#+\s*/, ''))}</${tag}>`;
        } else if (b.kind === 'list') {
            div.innerHTML = inline(b.raw.trim());
        } else if (b.kind === 'quote') {
            div.innerHTML = inline(b.raw.replace(/^>\s*/, ''));
        } else if (readonly) {
            div.textContent = b.raw.trim();
        } else {
            div.innerHTML = inline(b.raw);
        }

        div.appendChild(el('span', 'ln', String(b.line)));

        const lineIssues = issuesFor(b.line);
        if (lineIssues.some((i) => i.severity === 'P0')) div.classList.add('hl-p0');
        else if (lineIssues.some((i) => i.severity === 'P1')) div.classList.add('hl-p1');

        if (decisionFor(b.line)) div.classList.add('annotated');
        if (state.selected && state.selected.line === b.line) div.classList.add('selected');

        // 点选：选中该段并展示到右栏
        div.addEventListener('click', (e) => {
            if (e.target.isContentEditable && document.activeElement === div) return;
            selectBlock(b.line, false);
        });
        div.addEventListener('input', () => {
            if (!state.dirty.has(b.line)) {
                state.dirty.add(b.line);
                // 记一条撤销项；同一段反复输入不重复入栈
                if (!state.history.some((h) => h.kind === 'stage' && h.line === b.line)) {
                    state.history.push({ kind: 'stage', line: b.line });
                }
                updateUndo();
            }
        });
        div.addEventListener('focus', () => selectBlock(b.line, false));

        content.appendChild(div);
    }
}

function selectBlock(line, scroll) {
    const block = state.blocks.find((b) => b.line === line);
    if (!block) return;
    state.selected = block;

    for (const node of document.querySelectorAll('.block')) {
        node.classList.toggle('selected', Number(node.dataset.line) === line);
    }
    for (const node of document.querySelectorAll('.outline-item')) node.classList.remove('active');

    renderOutline();
    renderSelectionPanel(line);
    updateNavLabel();

    if (scroll) {
        const node = document.querySelector(`.block[data-line="${line}"]`);
        if (node) node.scrollIntoView({ behavior: 'smooth', block: 'center' });
    }
}

function renderSelectionPanel(line) {
    const block = state.blocks.find((b) => b.line === line);
    if (!block) return;

    const box = document.getElementById('selected-element');
    box.classList.remove('empty');
    box.textContent = block.text.slice(0, 160) || '(空段落)';

    const props = document.getElementById('element-props');
    const lineIssues = issuesFor(block.line);
    const kindLabel = {
        heading: `标题 h${block.level || 1}`,
        paragraph: '正文段落',
        list: '列表项',
        quote: '引用',
        table: '表格',
        math: '公式',
    }[block.kind] || block.kind;

    props.style.display = 'block';
    props.innerHTML =
        `<span>第 ${block.line} 行</span>` +
        `<span>${esc(kindLabel)}</span>` +
        (lineIssues.length > 0 ? `<span>${lineIssues.length} 个问题</span>` : '');

    // 已选 chips 回填
    const existing = decisionFor(block.line);
    state.selection = existing ? [...(existing.verdicts || [])] : [];
    document.getElementById('annotation-text').value = existing ? (existing.note || '') : '';

    const readonly = block.kind === 'math' || block.kind === 'table';
    document.getElementById('annotation-input').style.display = readonly ? 'none' : 'block';
    if (!readonly) paintChips();
}

function paintChips() {
    const parent = document.getElementById('verdict-chips');
    parent.innerHTML = '';
    const options = state.catalogs?.verdicts || [];
    const rec = state.recommendations.verdict;

    for (const opt of options) {
        const on = state.selection.includes(opt.id);
        const chip = el('button', 'chip' + (on ? ' selected' : ''));
        chip.type = 'button';
        chip.title = opt.desc_zh || '';
        chip.appendChild(el('span', null, opt.label_zh || opt.id));
        if (opt.id === rec && !on) {
            chip.classList.add('recommended');
            chip.appendChild(el('span', 'rec-badge', '★'));
        }
        chip.addEventListener('click', (e) => {
            e.stopPropagation();
            state.selection = on
                ? state.selection.filter((x) => x !== opt.id)
                : [...state.selection, opt.id];
            paintChips();
        });
        parent.appendChild(chip);
    }
}

// ── 右栏：已记录标注 ──────────────────────────────────────────────────────

function renderAnnotations() {
    const box = document.getElementById('annotations');
    box.innerHTML = '';
    if (state.decisions.length === 0) {
        box.innerHTML = '<div class="empty-sm">还没有记录。点正文里的段落开始。</div>';
        return;
    }

    const sorted = [...state.decisions].sort((a, b) => (a.line || 0) - (b.line || 0));
    for (const d of sorted) {
        const card = el('div', 'annotation-item');
        const labels = (d.verdicts || []).map(verdictLabel);
        const meta = el('div', 'meta');
        meta.innerHTML = `<b>第 ${d.line} 行</b>` + (labels.length ? ' · ' + esc(labels.join(' · ')) : '');
        card.appendChild(meta);
        if (d.note) card.appendChild(el('div', 'body', d.note));

        const del = el('button', 'del', '×');
        del.type = 'button';
        del.title = '删除这条';
        del.addEventListener('click', async (e) => {
            e.stopPropagation();
            await post('/api/decision/delete', { line: d.line });
            await loadDecisions();
            renderDoc();
            renderOutline();
            if (state.selected && state.selected.line === d.line) renderSelectionPanel(d.line);
            updateUndo();
        });
        card.appendChild(del);

        card.addEventListener('click', () => selectBlock(d.line, true));
        box.appendChild(card);
    }
}

// ── 记录 / 撤销 / 提交 ────────────────────────────────────────────────────

async function addAnnotation() {
    if (!state.selected) { toast('先点一段正文'); return; }
    const note = document.getElementById('annotation-text').value.trim();
    if (state.selection.length === 0 && note === '') {
        toast('选一个处理方式，或写一句具体要求');
        return;
    }
    const res = await post('/api/decision', {
        line: state.selected.line,
        quote: state.selected.text.slice(0, 120),
        verdicts: state.selection,
        note,
    });
    if (!res || !res.ok) {
        toast('记不下：' + ((res && res.error) || '未知'));
        return;
    }
    state.history.push({ kind: 'decision', line: state.selected.line });
    await loadDecisions();
    renderDoc();
    renderOutline();
    renderSelectionPanel(state.selected.line);
    updateUndo();
    toast('已记下');
}

/**
 * 撤销上一条。
 *
 * 两类操作分开回退，顺序是「先撤最近的」：
 *   - 暂存编辑（本地文字改动）→ /api/stage/undo
 *   - 标注（给 AI 的决定）    → /api/decision/delete
 */
async function undoLast() {
    const last = state.history.pop();
    if (!last) return;

    if (last.kind === 'stage') {
        const res = await post('/api/stage/undo', {});
        if (res && res.ok) {
            state.dirty.delete(last.line);
            renderDoc();
            selectBlock(last.line, false);
            toast('已撤销该段改动');
        }
    } else {
        await post('/api/decision/delete', { line: last.line });
        await loadDecisions();
        renderDoc();
        renderOutline();
        if (state.selected) renderSelectionPanel(state.selected.line);
        toast('已删除该条标注');
    }
    updateUndo();
}

function updateUndo() {
    const btn = document.getElementById('btn-undo');
    btn.disabled = state.history.length === 0;
    btn.textContent = state.history.length > 0 ? `撤销上一条 (${state.history.length})` : '撤销上一条';
}

/**
 * 提交标注：把暂存的正文改动一次性落盘，再写交接记录。
 *
 * 顺序很重要 —— 先落盘正文，再写 review-result.json。反过来会出现"交接记录
 * 说改完了，但磁盘上还是旧稿"。
 */
async function submitAnnotations() {
    // 1. 把 DOM 里改过的段落推进暂存队列（不落盘）
    for (const line of state.dirty) {
        const block = state.blocks.find((b) => b.line === line);
        if (!block) continue;
        const after = readBlockFromDom(block);
        if (after == null || after === block.raw) continue;
        await post('/api/stage', { line, after });
    }

    const staged = await fetch('/api/staged').then((r) => r.json()).catch(() => null);
    const stagedCount = (staged && staged.staged) ? staged.staged.length : 0;

    // 2. 一次性落盘
    if (stagedCount > 0) {
        const saved = await post('/api/save', {});
        if (!saved || !saved.ok) {
            toast('正文保存失败：' + describeFail(saved), 4000);
            return;
        }
        if (Array.isArray(saved.rejected) && saved.rejected.length > 0) {
            toast(`有 ${saved.rejected.length} 段没能写入：` + describeFail(saved), 5000);
        }
        state.dirty.clear();
    }

    // 3. 写交接记录
    const res = await post('/api/confirm', {});
    if (!res || !res.ok) { toast('提交失败'); return; }
    const rec = res.record;

    let msg = `已提交 ${rec.decisions} 条标注`;
    if (stagedCount > 0) msg += `、${stagedCount} 段正文改动`;
    if (rec.p0 > 0) {
        msg += `。\n\n还有 ${rec.p0} 条 P0 问题没解决，AI 会先处理这些再改稿。`;
    }
    msg += '\n\n现在可以回到对话里说「继续」。';

    showModal(msg, async () => {
        window.location.href = '/done';
    });
}

function describeFail(res) {
    if (!res) return '无响应';
    if (res.error) return res.error;
    if (Array.isArray(res.rejected) && res.rejected.length) {
        return res.rejected.map((r) => `${r.line}:${r.reason}`).join(', ');
    }
    return '未知';
}

function readBlockFromDom(block) {
    const node = document.querySelector(`.block[data-line="${block.line}"]`);
    if (!node) return null;
    if (block.kind === 'heading') {
        const h = node.querySelector('h1,h2,h3,h4');
        return '#'.repeat(block.level || 1) + ' ' + ((h ? h.textContent : '') || '').trim();
    }
    if (block.kind === 'list') {
        const prefix = (block.raw.match(/^\s*(?:[-*]|\d+[.)])\s*/) || [''])[0];
        return prefix + node.textContent.trim();
    }
    if (block.kind === 'quote') return '> ' + node.textContent.trim();
    return node.textContent.trim();
}

// ── 导航 ──────────────────────────────────────────────────────────────────

function updateNavLabel() {
    const counts = { p0: 0, p1: 0 };
    for (const i of state.issues) {
        if (i.severity === 'P0') counts.p0 += 1;
        else if (i.severity === 'P1') counts.p1 += 1;
    }
    document.getElementById('stat-p0').textContent = `P0 ${counts.p0}`;
    document.getElementById('stat-p1').textContent = `P1 ${counts.p1}`;

    const badge = document.getElementById('verdict-badge');
    if (counts.p0 > 0) { badge.className = 'badge p0'; badge.textContent = '有阻断项'; }
    else if (counts.p1 > 0) { badge.className = 'badge p1'; badge.textContent = '可交付'; }
    else { badge.className = 'badge ok'; badge.textContent = '全部通过'; }

    document.getElementById('nav-counter').textContent = `${state.blocks.length} 段`;
    document.getElementById('nav-name').textContent = state.selected
        ? `第 ${state.selected.line} 行`
        : '未选中';
}

function jumpNextIssue() {
    if (state.issues.length === 0) { toast('没有问题了'); return; }
    const currentLine = state.selected ? state.selected.line : -1;
    const ahead = state.issues
        .filter((i) => i.line != null && i.line > currentLine)
        .sort((a, b) => a.line - b.line);
    const target = ahead.length > 0 ? ahead[0] : [...state.issues].filter((i) => i.line != null).sort((a, b) => a.line - b.line)[0];
    if (target) selectBlock(target.line, true);
}

// ── Modal ─────────────────────────────────────────────────────────────────

function showModal(message, onConfirm) {
    document.getElementById('modal-message').textContent = message;
    document.getElementById('modal-overlay').style.display = 'flex';
    document.getElementById('modal-confirm').onclick = () => {
        document.getElementById('modal-overlay').style.display = 'none';
        if (onConfirm) onConfirm();
    };
    document.getElementById('modal-cancel').onclick = () => {
        document.getElementById('modal-overlay').style.display = 'none';
    };
}

// ── 磁盘漂移检测（学 svg_editor 的 mtime 轮询 + reload banner）─────────────

/** 在页面顶部插一条"文件已更新"横幅。 */
function showDriftBanner() {
    if (state.driftPending) return;
    state.driftPending = true;

    const old = document.getElementById('reload-banner');
    if (old) old.remove();

    const banner = el('div', null, '');
    banner.id = 'reload-banner';
    banner.innerHTML =
        '<span>draft.md 已在磁盘上更新（AI 那边改过了）</span>' +
        '<button id="reload-now">重新加载</button>' +
        '<button id="reload-dismiss" title="忽略">×</button>';

    banner.querySelector('#reload-now').addEventListener('click', async (e) => {
        e.stopPropagation();
        hideDriftBanner();
        await reloadAll();
    });
    banner.querySelector('#reload-dismiss').addEventListener('click', (e) => {
        e.stopPropagation();
        hideDriftBanner();
    });

    document.body.appendChild(banner);
}

function hideDriftBanner() {
    state.driftPending = false;
    const old = document.getElementById('reload-banner');
    if (old) old.remove();
}

/** 重新拉取正文、发现与决定（丢弃本地未保存改动前会提示）。 */
async function reloadAll() {
    if (state.dirty.size > 0) {
        showModal(
            `有 ${state.dirty.size} 段本地改动还没保存。\n\n重新加载会丢弃这些改动，确定吗？`,
            async () => {
                state.dirty.clear();
                await loadDecisions();
                await loadCheck();
                restoreSelection();
                startPolling();
                toast('已重新加载');
            },
        );
        return;
    }
    await loadDecisions();
    await loadCheck();
    restoreSelection();
    startPolling();
    toast('已重新加载');
}

/** 重载后尽量保持原选中位置。 */
function restoreSelection() {
    if (state.selected) {
        const still = state.blocks.find((b) => b.line === state.selected.line);
        if (still) {
            selectBlock(still.line, false);
            return;
        }
    }
    const firstP0 = state.issues.find((i) => i.severity === 'P0' && i.line != null);
    if (firstP0) selectBlock(firstP0.line, false);
    else if (state.blocks.length > 0) selectBlock(state.blocks[0].line, false);
}

/**
 * 开始轮询 mtime。
 *
 * 三个前置条件缺一不可：
 *   1. 不是正在输入（contenteditable 聚焦时重绘会打断打字）
 *   2. 没有未保存的本地改动（那种情况下用户正在编辑，别打扰）
 *   3. 弹层没开着
 *
 * 触发时**只弹横幅，不自动重绘** —— 自动重绘会吞掉用户正在看的位置和选中态。
 * 让用户自己决定什么时候加载，这是 svg_editor 的做法，也确实更稳。
 */
function startPolling() {
    if (state.pollTimer) clearInterval(state.pollTimer);
    state.pollTimer = setInterval(async () => {
        if (state.dirty.size > 0) return;
        if (document.getElementById('modal-overlay').style.display === 'flex') return;
        const active = document.activeElement;
        if (active && active.isContentEditable) return;

        try {
            const res = await fetch('/api/mtime').then((r) => r.json());
            if (!res || !res.ok) return;
            if (state.mtime != null && res.mtime != null && res.mtime !== state.mtime) {
                showDriftBanner();
            }
        } catch {
            /* 网络抖动忽略，下一轮再试 */
        }
    }, POLL_MS);
}

async function loadDecisions() {
    const res = await fetch('/api/decisions').then((r) => r.json()).catch(() => null);
    state.decisions = (res && res.decisions) || [];
    renderAnnotations();
}

async function loadCheck() {
    const res = await fetch('/api/check').then((r) => r.json()).catch(() => null);
    if (!res || !res.ok) { toast('检查失败'); return; }
    state.blocks = res.blocks;
    state.issues = res.issues;
    // 记录本次渲染对应的磁盘状态，后续轮询据此判断漂移
    state.mtime = res.mtime != null ? res.mtime : null;
    hideDriftBanner();
    renderDoc();
    renderOutline();
    updateNavLabel();
}

// ── 事件 ──────────────────────────────────────────────────────────────────

document.getElementById('btn-add-annotation').addEventListener('click', addAnnotation);
document.getElementById('btn-undo').addEventListener('click', undoLast);
document.getElementById('btn-save').addEventListener('click', submitAnnotations);
document.getElementById('btn-exit-preview').addEventListener('click', () => {
    showModal('退出预览不会丢失已记录的决定（都写在 review-notes.json 里）。\n\n确定退出？', async () => {
        await post('/api/shutdown', {});
        document.body.innerHTML =
            '<div style="display:grid;place-items:center;height:100vh;color:#8a8aa8;font-size:14px">' +
            '预览已关闭。回到对话里说「继续」。</div>';
    });
});
document.getElementById('nav-top').addEventListener('click', () => {
    document.getElementById('doc-container').scrollTo({ top: 0, behavior: 'smooth' });
});
document.getElementById('nav-issue').addEventListener('click', jumpNextIssue);

document.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
        e.preventDefault();
        addAnnotation();
    }
});

(async function boot() {
    const cat = await fetch('/api/catalogs').then((r) => r.json()).catch(() => null);
    state.catalogs = cat && cat.verdicts ? cat : { verdicts: [] };
    state.recommendations = (cat && cat.recommendations) || {};

    await loadDecisions();
    state.originalDecisions = [...state.decisions];
    await loadCheck();

    // 自动选中第一个有 P0 的段落，让用户一进来就知道从哪开始
    const firstP0 = state.issues.find((i) => i.severity === 'P0' && i.line != null);
    if (firstP0) selectBlock(firstP0.line, false);
    else if (state.blocks.length > 0) selectBlock(state.blocks[0].line, false);

    startPolling();
})();
