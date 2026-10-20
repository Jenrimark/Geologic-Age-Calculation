/* ============================================================================
 *  年代表年龄转换 · 文件上传版前端
 *  - 上传 Excel -> /inspect-file 探测结构 -> 选工作表/列 -> /convert-file -> 下载
 *  - 手动输入年龄 -> /convert
 *  - 每 5 秒 /ping 心跳，供服务端判断页面是否已关闭
 * ========================================================================== */
(function () {
  'use strict';

  const $id = (s) => document.getElementById(s);
  const els = {
    apiStatus: $id('apiStatus'),
    stVersions: $id('stVersions'), stUnits: $id('stUnits'),
    dropZone: $id('dropZone'), fileInput: $id('fileInput'), fileInfo: $id('fileInfo'),
    sheetField: $id('sheetField'), sheetSel: $id('sheetSel'),
    colField: $id('colField'), colSel: $id('colSel'), colMeta: $id('colMeta'),
    srcSel2: $id('srcSel2'), dstSel2: $id('dstSel2'),
    srcMeta2: $id('srcMeta2'), dstMeta2: $id('dstMeta2'),
    swapBtn2: $id('swapBtn2'), goFileBtn: $id('goFileBtn'),
    fileSummary: $id('fileSummary'), fileResultBody: $id('fileResultBody'), dlBtn: $id('dlBtn'),
    srcSel: $id('srcSel'), dstSel: $id('dstSel'),
    srcMeta: $id('srcMeta'), dstMeta: $id('dstMeta'),
    swapBtn: $id('swapBtn'), ageInput: $id('ageInput'), goBtn: $id('goBtn'),
    resultBody: $id('resultBody'), flowSrc: $id('flowSrc'), flowDst: $id('flowDst'),
    flowMethod: $id('flowMethod'), axisBox: $id('axisBox'), axisSvg: $id('axisSvg'),
    coverSvg: $id('coverSvg')
  };

  let online = false;
  let VERSIONS = [];
  let TOTAL_UNITS = 0;
  let fileB64 = null;
  let fileName = '';
  let fileMeta = null;
  let currentSheet = null;

  /* ---------------- 工具 ---------------- */
  function setStatus(cls, text) {
    els.apiStatus.className = 'status ' + cls;
    els.apiStatus.textContent = text;
  }
  function show(el) { el.classList.remove('hidden'); }
  function hide(el) { el.classList.add('hidden'); }
  function fileNote(cls, html) {
    els.fileInfo.className = 'fileinfo' + (cls ? ' ' + cls : '');
    els.fileInfo.innerHTML = html;
    show(els.fileInfo);
  }
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, (c) =>
      ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  }
  function fmtNum(v) {
    if (v === null || v === undefined || v === '') return '';
    const n = Number(v);
    return Number.isNaN(n) ? String(v) : String(v);
  }

  /* ---------------- 心跳 ---------------- */
  setInterval(() => { if (online) fetch('/ping').catch(() => {}); }, 5000);

  /* ---------------- 连接检测 ---------------- */
  async function checkHealth() {
    try {
      const h = await (await fetch('/health')).json();
      online = h.status === 'ok';
      setStatus(online ? 'on' : 'off',
        online ? `R 服务已连接 · ${h.versions} 版本` : 'R 服务响应异常');
      if (online && !VERSIONS.length) {
        els.stVersions.textContent = h.versions;
        const v = await (await fetch('/versions')).json();
        VERSIONS = v.versions || [];
        // 用各版本覆盖数的最大值作为地层单元总数，避免把表头行计入
        TOTAL_UNITS = VERSIONS.reduce((m, x) => Math.max(m, x.cover || 0), 0);
        els.stUnits.textContent = TOTAL_UNITS;
        fillVersionSelects();
        drawCoverage();
        els.goBtn.disabled = false;
      }
      if (online) refreshGoFile();
    } catch (e) {
      online = false;
      setStatus('off', '未连接 R 服务，请先启动 R 脚本');
      els.goBtn.disabled = true;
      els.goFileBtn.disabled = true;
      els.ageInput.placeholder = '请先启动 R 服务后刷新本页';
    }
  }

  function fillVersionSelects() {
    [[els.srcSel, els.dstSel], [els.srcSel2, els.dstSel2]].forEach(([s, d]) => {
      [s, d].forEach((sel) => {
        if (!sel) return;
        sel.innerHTML = '';
        VERSIONS.forEach((v) => {
          const o = document.createElement('option');
          o.value = v.name;
          o.textContent = `${String(v.index).padStart(2, '0')}　${v.name}　(${v.cover} 单元)`;
          sel.appendChild(o);
        });
      });
    });
    const best = [...VERSIONS].sort((a, b) => b.cover - a.cover)[0];
    if (best) {
      if (els.dstSel) els.dstSel.value = best.name;
      if (els.dstSel2) els.dstSel2.value = best.name;
      const first = VERSIONS[0] ? VERSIONS[0].name : '';
      if (els.srcSel) els.srcSel.value = first;
      if (els.srcSel2) els.srcSel2.value = first;
    }
    syncMeta();
  }

  function metaOf(name) {
    const v = VERSIONS.find((x) => x.name === name);
    if (!v) return '';
    const pct = TOTAL_UNITS ? Math.round(v.cover * 100 / TOTAL_UNITS) : 0;
    return `覆盖 ${v.cover}/${TOTAL_UNITS} 单元（${pct}%）· 序号 ${v.index}`;
  }
  function syncMeta() {
    if (els.srcMeta) els.srcMeta.textContent = metaOf(els.srcSel.value);
    if (els.dstMeta) els.dstMeta.textContent = metaOf(els.dstSel.value);
    if (els.srcMeta2) els.srcMeta2.textContent = metaOf(els.srcSel2.value);
    if (els.dstMeta2) els.dstMeta2.textContent = metaOf(els.dstSel2.value);
  }

  /* ---------------- 页签 ---------------- */
  document.querySelectorAll('.tab').forEach((btn) => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab').forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');
      const isFile = btn.dataset.mode === 'file';
      if (isFile) { hide($id('modeManual')); show($id('modeFile')); }
      else { hide($id('modeFile')); show($id('modeManual')); }
    });
  });

  /* ---------------- 交换 ---------------- */
  if (els.swapBtn) els.swapBtn.addEventListener('click', () => {
    const t = els.srcSel.value; els.srcSel.value = els.dstSel.value; els.dstSel.value = t; syncMeta();
  });
  if (els.swapBtn2) els.swapBtn2.addEventListener('click', () => {
    const t = els.srcSel2.value; els.srcSel2.value = els.dstSel2.value; els.dstSel2.value = t;
    syncMeta(); refreshGoFile();
  });
  [els.srcSel, els.dstSel, els.srcSel2, els.dstSel2].forEach((s) => {
    if (s) s.addEventListener('change', () => { syncMeta(); refreshGoFile(); });
  });

  /* =======================================================================
   *  文件上传
   * ===================================================================== */
  els.dropZone.addEventListener('click', () => els.fileInput.click());
  els.fileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files[0]) handleFile(e.target.files[0]);
  });
  ['dragenter', 'dragover'].forEach((ev) =>
    els.dropZone.addEventListener(ev, (e) => { e.preventDefault(); els.dropZone.classList.add('over'); }));
  ['dragleave', 'drop'].forEach((ev) =>
    els.dropZone.addEventListener(ev, (e) => { e.preventDefault(); els.dropZone.classList.remove('over'); }));
  els.dropZone.addEventListener('drop', (e) => {
    const f = e.dataTransfer.files && e.dataTransfer.files[0];
    if (f) handleFile(f);
  });
  document.addEventListener('paste', (e) => {
    const items = e.clipboardData && e.clipboardData.files;
    const fm = $id('modeFile');
    if (items && items.length && fm && !fm.classList.contains('hidden')) handleFile(items[0]);
  });

  function readFileAsBase64(file) {
    return new Promise((resolve, reject) => {
      const r = new FileReader();
      r.onload = () => {
        const s = String(r.result || '');
        const i = s.indexOf(',');
        resolve(i >= 0 ? s.slice(i + 1) : s);
      };
      r.onerror = () => reject(new Error('读取文件失败'));
      r.readAsDataURL(file);
    });
  }

  function refreshGoFile() {
    if (!els.goFileBtn) return;
    const ok = online && !!fileB64 && !!els.colSel.value &&
      !!els.srcSel2.value && !!els.dstSel2.value && els.srcSel2.value !== els.dstSel2.value;
    els.goFileBtn.disabled = !ok;
  }

  async function handleFile(file) {
    hide(els.dlBtn);
    els.fileResultBody.innerHTML = '';
    els.fileSummary.innerHTML = '';
    const ext = (file.name.split('.').pop() || '').toLowerCase();
    if (ext !== 'xlsx' && ext !== 'xls') {
      fileNote('err', `不支持 <b>${esc(ext || '该')}</b> 格式，仅支持 .xlsx / .xls`);
      fileB64 = null; refreshGoFile(); return;
    }
    if (file.size > 40 * 1024 * 1024) {
      fileNote('err', `文件约 ${(file.size / 1024 / 1024).toFixed(1)} MB，超过 40 MB 上限`);
      fileB64 = null; refreshGoFile(); return;
    }

    fileNote('busy', `正在读取 <b>${esc(file.name)}</b>（${(file.size / 1024).toFixed(1)} KB）…`);
    try {
      fileB64 = await readFileAsBase64(file);
      fileName = file.name;
    } catch (e) {
      fileNote('err', esc(e.message || '读取文件失败'));
      fileB64 = null; refreshGoFile(); return;
    }

    fileNote('busy', '正在解析工作表与列名…');
    let j;
    try {
      const r = await fetch('/inspect-file', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ filename: fileName, content: fileB64 })
      });
      j = await r.json();
    } catch (e) {
      fileNote('err', '无法连接 R 服务，请确认服务已启动'); return;
    }
    if (j.error) {
      fileNote('err', `<b>${esc(j.error)}</b>${j.hint ? '　— ' + esc(j.hint) : ''}`);
      fileB64 = null; refreshGoFile(); return;
    }

    fileMeta = j.sheets || [];
    if (!fileMeta.length) { fileNote('err', '未解析到任何工作表'); fileB64 = null; return; }

    els.sheetSel.innerHTML = '';
    fileMeta.forEach((s, i) => {
      const o = document.createElement('option');
      o.value = String(i);
      o.textContent = `${s.sheet}　(${s.columns.length} 列 / ${s.rows} 行)`;
      els.sheetSel.appendChild(o);
    });
    show(els.sheetField);
    pickSheet(0);
    fileNote('', `<b>${esc(file.name)}</b>　已就绪 · ${fileMeta.length} 个工作表`);
    refreshGoFile();
  }

  els.sheetSel.addEventListener('change', () => pickSheet(Number(els.sheetSel.value)));

  function pickSheet(idx) {
    const s = fileMeta[idx];
    if (!s) return;
    currentSheet = s;
    els.colSel.innerHTML = '';
    s.columns.forEach((c) => {
      const o = document.createElement('option');
      o.value = c; o.textContent = c;
      els.colSel.appendChild(o);
    });
    show(els.colField);
    const guess = s.columns.find((c) => /年龄|age|年代|Ma$/i.test(c));
    if (guess) els.colSel.value = guess;
    updateColMeta();
    refreshGoFile();
  }

  els.colSel.addEventListener('change', () => { updateColMeta(); refreshGoFile(); });

  function updateColMeta() {
    const s = currentSheet;
    if (!s) { els.colMeta.textContent = ''; return; }
    els.colMeta.textContent = `工作表「${s.sheet}」共 ${s.columns.length} 列 / ${s.rows} 行；` +
      (els.colSel.value ? `已选「${els.colSel.value}」` : '请选择年龄列');
  }

  /* ---------- 执行文件转换 ---------- */
  els.goFileBtn.addEventListener('click', async () => {
    if (!online) { alert('未连接 R 服务，请先启动 R 脚本'); return; }
    if (!fileB64) { alert('请先上传 Excel 文件'); return; }
    const sheetName = currentSheet ? currentSheet.sheet : '';
    const ageCol = els.colSel.value;
    const src = els.srcSel2.value, dst = els.dstSel2.value;
    if (!ageCol) { alert('请选择年龄所在列'); return; }
    if (!src || !dst) { alert('请选择源版本与目标版本'); return; }
    if (src === dst) { alert('源版本与目标版本相同，无需转换'); return; }

    els.goFileBtn.disabled = true;
    els.goFileBtn.textContent = '转换中…';
    fileNote('busy', `正在转换「${esc(sheetName)}」的「${esc(ageCol)}」列：${esc(src)} → ${esc(dst)}`);
    els.fileSummary.innerHTML = ''; els.fileResultBody.innerHTML = ''; hide(els.dlBtn);

    let j;
    try {
      const r = await fetch('/convert-file', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          filename: fileName, content: fileB64,
          sheet: sheetName, age_column: ageCol, source: src, target: dst
        })
      });
      j = await r.json();
    } catch (e) {
      fileNote('err', '请求失败，请确认 R 服务仍在运行');
      els.goFileBtn.disabled = false; els.goFileBtn.textContent = '开始转换'; return;
    }

    els.goFileBtn.disabled = false; els.goFileBtn.textContent = '开始转换';
    refreshGoFile();

    if (j.error) {
      fileNote('err', `<b>${esc(j.error)}</b>${j.message ? '<br>' + esc(j.message) : ''}` +
        (j.hint ? '<br>' + esc(j.hint) : ''));
      els.fileSummary.innerHTML = '';
      return;
    }

    fileNote('', `<b>${esc(fileName)}</b> · 工作表「${esc(j.sheet)}」· 年龄列「${esc(j.age_column)}」`);

    const pct = j.rows ? Math.round(j.converted * 100 / j.rows) : 0;
    const fmtCls = j.comments_lost || !j.format_preserved ? 'warn' : 'ok';
    els.fileSummary.innerHTML =
      `共 <b>${j.rows}</b> 行 → <span class="ok">成功 ${j.converted}</span>（${pct}%）` +
      (j.failed ? `　·　<span class="warn">无法转换 ${j.failed}</span>` : '') +
      (j.skipped ? `　·　<span class="warn">空/非数值 ${j.skipped}</span>` : '') +
      `<br>${esc(j.source)} → ${esc(j.target)}<br>` +
      `<span class="${fmtCls}">${esc(j.format_note || '')}</span>` +
      (j.dangling_cleaned ? '　·　已修正表格兼容性' : '') +
      `<br>新增 ${j.new_columns.length} 列：<b>${esc(j.new_columns.join('　·　'))}</b>` +
      ((j.preview || []).length < (j.converted + j.failed)
        ? `<br><span class="warn">下表仅预览前 ${j.preview.length} 条，完整结果见下载文件</span>` : '');

    const rows = j.preview || [];
    els.fileResultBody.innerHTML = rows.map((r) => {
      const bad = r.output === null || r.output === undefined || r.output === '';
      const cls = bad ? 'fail' : (r.method === '精确匹配' ? 'exact' : 'interp');
      return `<tr>
        <td class="n">${esc(r.row)}</td>
        <td class="n">${fmtNum(r.input)}</td>
        <td class="n">${bad ? '—' : fmtNum(r.output)}</td>
        <td><span class="badge ${cls}">${esc(r.method)}</span></td>
        <td>${esc(r.stage || '')}</td></tr>`;
    }).join('');

    if (j.content) {
      const bin = atob(j.content);
      const u8 = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
      const blob = new Blob([u8], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      });
      els.dlBtn.href = URL.createObjectURL(blob);
      els.dlBtn.download = j.filename || '转换结果.xlsx';
      els.dlBtn.textContent = `⬇ 下载 ${j.filename || '转换结果.xlsx'}`;
      show(els.dlBtn);
    }
  });

  /* =======================================================================
   *  手动输入
   * ===================================================================== */
  async function convertMany(ages) {
    const r = await fetch('/convert', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ source: els.srcSel.value, target: els.dstSel.value, ages: ages })
    });
    const j = await r.json();
    if (j.error) throw new Error(j.error + (j.message ? '：' + j.message : ''));
    return j.results || [];
  }

  els.goBtn.addEventListener('click', run);
  els.ageInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) run();
  });

  async function run() {
    if (!online) { alert('未连接 R 服务，请先启动 R 脚本'); return; }
    const tokens = els.ageInput.value.split(/[,，;；、\s]+/).filter(Boolean);
    const ages = tokens.map(Number).filter((n) => !Number.isNaN(n));
    if (!ages.length) { alert('请输入有效年龄数字'); return; }
    els.goBtn.disabled = true; els.goBtn.textContent = '转换中…';
    try {
      const results = await convertMany(ages);
      if (!results.length) { alert('这些年龄超出所选版本的控制点范围，无法转换（不外推）'); return; }
      renderManual(results);
    } catch (err) {
      alert('转换失败：' + err.message);
    } finally {
      els.goBtn.disabled = false; els.goBtn.textContent = '转换';
    }
  }

  function renderManual(results) {
    els.flowSrc.textContent = els.srcSel.value;
    els.flowDst.textContent = els.dstSel.value;
    const nExact = results.filter((r) => r.method === '精确匹配').length;
    els.flowMethod.textContent =
      `共 ${results.length} 条 · 精确匹配 ${nExact} · 插值 ${results.length - nExact}`;
    els.resultBody.innerHTML = results.map((r) => {
      const cls = r.method === '精确匹配' ? 'exact' : 'interp';
      return `<tr><td class="n">${fmtNum(r.input)}</td><td class="n">${fmtNum(r.output)}</td>
        <td><span class="badge ${cls}">${esc(r.method)}</span></td>
        <td>${esc(r.stage || '')}</td></tr>`;
    }).join('');
    const it = results.find((r) => r.method !== '精确匹配');
    if (it) drawAxis(it); else hide(els.axisBox);
  }

  /* ---------- 插值示意 ---------- */
  function drawAxis(r) {
    if (!r.lower || !r.upper) { hide(els.axisBox); return; }
    show(els.axisBox);
    const lo = r.lower, hi = r.upper;
    const W = 720, H = 170, P = { l: 56, r: 20, t: 22, b: 40 };
    const xmin = Math.min(lo.x, hi.x, r.input) - 0.5;
    const xmax = Math.max(lo.x, hi.x, r.input) + 0.5;
    const ymin = Math.min(lo.y, hi.y, r.output) - 0.5;
    const ymax = Math.max(lo.y, hi.y, r.output) + 0.5;
    const X = (v) => P.l + (v - xmin) / (xmax - xmin) * (W - P.l - P.r);
    const Y = (v) => H - P.b - (v - ymin) / (ymax - ymin) * (H - P.t - P.b);
    let s = `<rect x="${P.l}" y="${P.t}" width="${W - P.l - P.r}" height="${H - P.t - P.b}" fill="#fff" stroke="#e2e8f0"/>`;
    s += `<line x1="${X(lo.x)}" y1="${Y(lo.y)}" x2="${X(hi.x)}" y2="${Y(hi.y)}" stroke="#2563eb" stroke-width="2.5"/>`;
    [[lo, '下控制点'], [hi, '上控制点']].forEach(([p, lab]) => {
      s += `<circle cx="${X(p.x)}" cy="${Y(p.y)}" r="5" fill="#2563eb"/>`;
      s += `<text x="${X(p.x)}" y="${Y(p.y) - 10}" font-size="11" fill="#64748b" text-anchor="middle">${lab} ${p.x}→${p.y}</text>`;
    });
    s += `<circle cx="${X(r.input)}" cy="${Y(r.output)}" r="6" fill="#dc2626"/>`;
    s += `<text x="${X(r.input)}" y="${Y(r.output) + 20}" font-size="12" fill="#dc2626" text-anchor="middle" font-weight="bold">${r.input} → ${r.output}</text>`;
    s += `<text x="${W / 2}" y="${H - 8}" font-size="11" fill="#94a3b8" text-anchor="middle">源版本年龄 (Ma) → 目标版本年龄 (Ma)</text>`;
    els.axisSvg.innerHTML = s;
  }

  /* ---------- 覆盖度柱状图 ---------- */
  function drawCoverage() {
    if (!VERSIONS.length) return;
    const W = 360, H = 300, P = { l: 34, r: 8, t: 12, b: 42 };
    const n = VERSIONS.length;
    const bw = (W - P.l - P.r) / n;
    const ymax = TOTAL_UNITS || 1;
    let s = '';
    [0, .25, .5, .75, 1].forEach((f) => {
      const y = H - P.b - f * (H - P.t - P.b);
      s += `<line x1="${P.l}" y1="${y}" x2="${W - P.r}" y2="${y}" stroke="#e2e8f0" stroke-width="1"/>`;
      s += `<text x="${P.l - 5}" y="${y + 4}" font-size="9" fill="#94a3b8" text-anchor="end">${Math.round(f * ymax)}</text>`;
    });
    VERSIONS.forEach((v, i) => {
      const h = (v.cover / ymax) * (H - P.t - P.b);
      const x = P.l + i * bw, y = H - P.b - h;
      const col = v.cover >= 90 ? '#0e7490' : (v.cover >= 40 ? '#f59e0b' : '#94a3b8');
      s += `<rect x="${x + bw * .12}" y="${y}" width="${Math.max(bw * .76, 1)}" height="${Math.max(h, .5)}" fill="${col}" rx="1"><title>${esc(v.name)}：覆盖 ${v.cover} 单元</title></rect>`;
    });
    s += `<text x="${W / 2}" y="${H - 14}" font-size="10" fill="#64748b" text-anchor="middle">版本顺序 1 → ${n}（悬停查看名称）</text>`;
    els.coverSvg.innerHTML = s;
  }

  /* ---------------- 启动 ---------------- */
  checkHealth();
  setInterval(checkHealth, 8000);
})();
