/* 不同版本地质年代自动对齐模型 · 前端（纯 R 计算）
 * 页面由 R 服务同源托管；版本列表、覆盖度图、转换结果全部来自 R 接口，浏览器不做任何计算。 */
(function () {
  const $id = (id) => document.getElementById(id);

  const els = {
    status: $id('apiStatus'),
    srcSel: $id('srcSel'), dstSel: $id('dstSel'),
    srcMeta: $id('srcMeta'), dstMeta: $id('dstMeta'),
    ageInput: $id('ageInput'), goBtn: $id('goBtn'),
    resultBody: $id('resultBody'),
    flowSrc: $id('flowSrc'), flowDst: $id('flowDst'), flowMethod: $id('flowMethod'),
    axisBox: $id('axisBox'), axisSvg: $id('axisSvg'),
    coverSvg: $id('coverSvg'),
    swapBtn: $id('swapBtn')
  };

  let versions = [];      // R /versions 返回
  let totalUnits = 118;   // 单元总数（来自 /health.rows）
  let online = false;

  /* ---------- 工具 ---------- */
  function fmt(n) {
    if (n == null) return '—';
    return Number(n).toFixed(4).replace(/\.?0+$/, '') || '0';
  }
  function setStatus(mode, text) {
    els.status.textContent = text;
    els.status.className = 'status ' + mode;
  }

  /* ---------- 启动：健康检查 + 拉取版本 ---------- */
  async function boot() {
    try {
      const h = await fetch('/health');
      if (!h.ok) throw 0;
      const hj = await h.json();
      online = true;
      totalUnits = hj.rows || 118;
      setStatus('on', 'R 服务已连接');
      const v = await (await fetch('/versions')).json();
      versions = v.versions || [];
      fillSelect(els.srcSel, versions);
      fillSelect(els.dstSel, versions);
      selectDefault(els.srcSel, 'GTS2004');
      selectDefault(els.dstSel, 'GTS2020');
      refreshMeta();
      drawCoverChart();
      els.goBtn.disabled = false;
    } catch (e) {
      online = false;
      setStatus('off', '未连接 R 服务，请先启动 R 脚本');
      els.goBtn.disabled = true;
      els.ageInput.placeholder = '请先在终端启动 R 服务后刷新本页';
    }
  }

  function fillSelect(sel, list) {
    sel.innerHTML = '';
    list.forEach(v => {
      const opt = document.createElement('option');
      opt.value = v.name;
      opt.textContent = `${v.name}（覆盖 ${v.cover}/${totalUnits}）`;
      sel.appendChild(opt);
    });
  }
  function selectDefault(sel, name) {
    const hit = [...sel.options].find(o => o.text.includes(name));
    if (hit) sel.value = hit.value;
  }
  function refreshMeta() {
    const s = versions.find(v => v.name === els.srcSel.value);
    const t = versions.find(v => v.name === els.dstSel.value);
    els.srcMeta.textContent = s ? `覆盖 ${s.cover}/${totalUnits} 单元` : '';
    els.dstMeta.textContent = t ? `覆盖 ${t.cover}/${totalUnits} 单元` : '';
  }

  /* ---------- 转换：只调 R /convert ---------- */
  async function convertMany(ages) {
    const params = new URLSearchParams({
      source: els.srcSel.value, target: els.dstSel.value, ages: ages.join(',')
    });
    const r = await fetch('/convert?' + params.toString());
    const j = await r.json();
    if (!r.ok || j.error) throw new Error(j.error || '转换失败');
    return (j.results || []).map(x => ({
      input: x.input, output: x.output,
      method: x.method === '精确匹配' ? '精确匹配' : '插值',
      stage: x.stage || '—',
      l: x.lower ? [x.lower.x, x.lower.y] : null,
      u: x.upper ? [x.upper.x, x.upper.y] : null,
      ratio: null
    }));
  }

  /* ---------- 渲染 ---------- */
  function render(results) {
    els.resultBody.innerHTML = '';
    let lastInterp = null;
    results.forEach(res => {
      const tr = document.createElement('tr');
      const tagCls = res.method === '精确匹配' ? 'exact' : 'interp';
      tr.innerHTML = `
        <td>${fmt(res.input)}</td>
        <td><b>${fmt(res.output)}</b></td>
        <td><span class="tag ${tagCls}">${res.method}</span></td>
        <td>${res.stage}</td>`;
      els.resultBody.appendChild(tr);
      if (res.method === '插值' && res.l) lastInterp = res;
    });
    const first = results[0];
    if (first) {
      els.flowSrc.textContent = `${fmt(first.input)} Ma`;
      els.flowDst.textContent = `${fmt(first.output)} Ma`;
      els.flowMethod.textContent = first.method === '精确匹配' ? '精确吸附控制点' : '一步直达 · 至多一次插值';
    }
    if (lastInterp) drawAxis(lastInterp); else els.axisBox.classList.remove('show');
  }

  // 注：插值数轴的上下控制点由 R 返回；此处仅在 R 提供 l/u 时绘制
  function drawAxis(res) {
    els.axisBox.classList.add('show');
    const svg = els.axisSvg;
    const W = 720, H = 170, padL = 70, padR = 70;
    const x0 = padL, x1 = W - padR;
    const lo = res.l[0], hi = res.u[0];
    const xp = v => x0 + (v - lo) / (hi - lo) * (x1 - x0);
    const y = 80;
    const xi = xp(res.input);
    svg.innerHTML = `
      <line x1="${x0}" y1="${y}" x2="${x1}" y2="${y}" stroke="#94a3b8" stroke-width="2"/>
      <circle cx="${xp(lo)}" cy="${y}" r="7" fill="#94a3b8"/>
      <circle cx="${xp(hi)}" cy="${y}" r="7" fill="#94a3b8"/>
      <circle cx="${xi}" cy="${y}" r="9" fill="#1d4ed8"/>
      <text x="${xp(lo)}" y="${y-16}" font-size="13" fill="#475569">下界 ${fmt(lo)} Ma → ${fmt(res.l[1])}</text>
      <text x="${xp(hi)}" y="${y-16}" font-size="13" fill="#475569" text-anchor="end">上界 ${fmt(hi)} Ma → ${fmt(res.u[1])}</text>
      <text x="${xi}" y="${y+38}" font-size="13" fill="#1d4ed8" text-anchor="middle">输入 ${fmt(res.input)}</text>
      <text x="${xi}" y="${y+58}" font-size="14" font-weight="700" fill="#0f766e" text-anchor="middle">输出 ${fmt(res.output)} Ma（目标版本）</text>
      <text x="${(x0+x1)/2}" y="26" font-size="13" fill="#6b7280" text-anchor="middle">分段线性插值：相邻控制点间按比例换算，不越过控制点范围</text>`;
  }

  /* ---------- 右侧覆盖度图（R /versions 真实数据） ---------- */
  function drawCoverChart() {
    const svg = els.coverSvg;
    const W = 360, H = 300, padL = 16, padR = 16, padT = 18, padB = 30;
    const n = versions.length;
    const plotW = W - padL - padR;
    const step = plotW / n;
    const bw = step * 0.62;
    let bars = '';
    versions.forEach((v, i) => {
      const x = padL + i * step + (step - bw) / 2;
      const h = (v.cover / totalUnits) * (H - padT - padB);
      const y = H - padB - h;
      const color = v.cover >= 90 ? '#0f766e' : v.cover >= 40 ? '#5eb5ae' : '#c7e2e0';
      bars += `<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${bw.toFixed(1)}" height="${h.toFixed(1)}" rx="1.5" fill="${color}"><title>${v.name}：覆盖 ${v.cover}/${totalUnits}</title></rect>`;
    });
    svg.innerHTML = `
      <line x1="${padL}" y1="${H-padB}" x2="${W-padR}" y2="${H-padB}" stroke="#cbd5e1"/>
      ${bars}
      <text x="${padL}" y="${H-padB+18}" font-size="11" fill="#6b7280">1917</text>
      <text x="${W-padR}" y="${H-padB+18}" font-size="11" fill="#6b7280" text-anchor="end">2026</text>
      <text x="${padL}" y="${padT-4}" font-size="11" fill="#6b7280">单元数 →</text>`;
  }

  /* ---------- 事件 ---------- */
  async function run() {
    if (!online) { alert('未连接 R 服务，请先启动 R 脚本'); return; }
    const tokens = els.ageInput.value.split(/[,，;；、\s]+/).filter(Boolean);
    const ages = tokens.map(Number).filter(n => !isNaN(n));
    if (!ages.length) { alert('请输入有效年龄数字'); return; }
    els.goBtn.textContent = '转换中…';
    try {
      const results = await convertMany(ages);
      if (!results.length) { alert('这些年龄超出所选版本的控制点范围，无法转换（不外推）'); return; }
      render(results);
    } catch (err) {
      alert(err.message || '转换失败');
    } finally {
      els.goBtn.textContent = '转换';
    }
  }

  els.goBtn.addEventListener('click', run);
  els.ageInput.addEventListener('keydown', e => { if (e.key === 'Enter') run(); });
  els.swapBtn.addEventListener('click', () => {
    const a = els.srcSel.value;
    els.srcSel.value = els.dstSel.value;
    els.dstSel.value = a;
    refreshMeta();
  });
  els.srcSel.addEventListener('change', refreshMeta);
  els.dstSel.addEventListener('change', refreshMeta);

  /* ---------- 关闭页面：通知 R 停止并清理 ---------- */
  window.addEventListener('pagehide', (e) => {
    if (e.persisted) return;               // 进入 bfcache（前进/后退）不关闭
    try { navigator.sendBeacon('/shutdown'); } catch (_) {}
  });

  boot();
})();
