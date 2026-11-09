/* UI for benchmark features. Loaded after app.js; uses its state, commit, render, dialog, toast helpers. */
'use strict';
function suppliersOfShip() {
  return [...new Set(state.rows.filter(r => ConfBench.SHIP.includes(r.meta.status)).map(ConfBench.supplierOf))].sort();
}
function templateFor(supplier) { return (state.templates && state.templates[supplier]) || null; }
function safeFile(name) { return String(name).replace(/[\\/:*?"<>|]/g, '_').slice(0, 80); }
function xlsxDownload(aoa, sheetName, filename) {
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), sheetName);
  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  download(out, filename, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
}
function benchView() {
  const sums = ConfBench.supplierSummary(state.rows, state.masters);
  const sups = suppliersOfShip();
  const shipRows = state.rows.filter(r => ConfBench.SHIP.includes(r.meta.status));
  const supRows = sups.length ? sups.map((s, i) => {
    const n = shipRows.filter(r => ConfBench.supplierOf(r) === s).length;
    return `<tr><td>${esc(s)}</td><td class="right">${fmt(n)}</td><td>${templateFor(s) ? '사용자 양식' : '기본 양식'}</td><td><div class="stack"><button class="row-btn" data-sup-xlsx="${i}">발주서 엑셀</button><button class="row-btn" data-sup-tpl="${i}">양식 편집</button></div></td></tr>`;
  }).join('') : '<tr><td colspan="4" class="muted">출고마감 또는 송장연결 행이 아직 없습니다.</td></tr>';
  const setRows = sums.length ? sums.map(s => `<tr><td><b>${esc(s.supplier)}</b></td><td class="right">${fmt(s.orders)}</td><td class="right">${fmt(s.pieces)}</td><td class="money">${fmt(s.sales)}</td><td class="money">${fmt(s.settle)}</td><td class="money">${fmt(s.fee)}${s.feeRate !== null ? ' <span class="small">(' + (s.feeRate * 100).toFixed(1) + '%)</span>' : ''}</td><td class="money">${s.cost === null ? '<span class="badge warn">원가 미등록 ' + fmt(s.costMissing) + '행</span>' : fmt(s.cost)}</td><td class="money">${s.csCount ? fmt(s.csDeduct) + ' <span class="small">(' + fmt(s.csCount) + '건)</span>' : '-'}</td><td class="money">${s.net === null ? '계산 불가' : fmt(s.net)}</td></tr>`).join('') : '<tr><td colspan="9" class="muted">검토완료 이후 행이 있어야 정산이 계산됩니다.</td></tr>';
  const csRows = state.rows.filter(r => r.meta.cs).slice(0, 50);
  const csHtml = csRows.length ? csRows.map(r => `<div class="timeline">${esc(r.values['주문고유번호'] || '(번호 없음)')} · <span class="badge">${esc(r.meta.cs.type)}</span> ${esc(new Date(r.meta.cs.at).toLocaleString('ko-KR'))}<br>${esc(r.meta.cs.reason || '')}</div>`).join('') : '<p class="small">기록된 반품·교환·취소가 없습니다.</p>';
  return `<div class="notice warn">발주서와 정산 요약은 이 PC에서 엑셀로 만들어집니다. 메일 발송이나 거래처 전송은 하지 않습니다. 발주서에는 수취인 정보가 들어 있으므로 전달 전에 개인정보 취급 기준을 확인하세요.</div>
<section class="panel"><div class="panel-head"><div><h3>거래처별 발주서</h3><span class="small">출고마감·송장연결 행을 거래처 양식대로 만듭니다.</span></div></div><div class="table-wrap"><table><thead><tr><th>거래처</th><th class="right">출고 행</th><th>양식</th><th></th></tr></thead><tbody>${supRows}</tbody></table></div></section>
<section class="panel"><div class="panel-head"><div><h3>거래처별 정산 요약</h3><span class="small">검토완료 이후 행 기준. 반품·취소 기록은 정산가에서 차감합니다. 교환은 차감하지 않습니다.</span></div><button id="settle-xlsx">정산 요약 엑셀</button></div><div class="table-wrap"><table><thead><tr><th>거래처</th><th class="right">주문 수</th><th class="right">낱알</th><th class="right">판매가 합</th><th class="right">정산가 합</th><th class="right">수수료</th><th class="right">원가</th><th class="right">반품·취소 차감</th><th class="right">예상 순이익</th></tr></thead><tbody>${setRows}</tbody></table></div><div class="cardbody small">순이익은 정산가 합에서 반품·취소 차감과 원가(낱알 × 원가)를 뺀 값입니다. 원가가 없는 상품이 있으면 계산하지 않습니다.</div></section>
<section class="panel"><div class="panel-head"><div><h3>반품·교환·취소 기록</h3><span class="small">주문번호가 같은 모든 행에 기록합니다.</span></div></div><div class="cardbody"><div class="stack"><input id="cs-order" placeholder="주문고유번호" aria-label="CS 주문번호"><select id="cs-type" aria-label="CS 유형"><option>반품</option><option>교환</option><option>취소</option></select><input id="cs-reason" maxlength="200" placeholder="사유" aria-label="CS 사유" style="min-width:260px"><button id="cs-save" class="primary">기록</button></div><div style="margin-top:14px">${csHtml}</div></div></section>`;
}
function bindBench() {
  document.querySelectorAll('[data-sup-xlsx]').forEach(b => b.onclick = () => exportSupplier(Number(b.dataset.supXlsx)));
  document.querySelectorAll('[data-sup-tpl]').forEach(b => b.onclick = () => openTemplate(Number(b.dataset.supTpl)));
  $('#settle-xlsx')?.addEventListener('click', exportSettlement);
  $('#cs-save')?.addEventListener('click', saveCS);
}
function exportSupplier(index) {
  try {
    const s = suppliersOfShip()[index];
    if (!s) throw Error('거래처를 찾을 수 없습니다.');
    const aoa = ConfBench.supplierOrderSheet(state.rows, s, templateFor(s) || ConfBench.DEFAULT_TEMPLATE);
    if (aoa.length <= 1) throw Error('이 거래처의 출고 행이 없습니다.');
    xlsxDownload(aoa, '발주서', safeFile('발주서_' + s + '_' + new Date().toLocaleDateString('en-CA') + '.xlsx'));
    toast(s + ' 발주서 ' + (aoa.length - 1) + '행을 내려받았습니다. 메일은 직접 보내세요.');
  } catch (e) { toast(e.message); }
}
function exportSettlement() {
  try {
    const sums = ConfBench.supplierSummary(state.rows, state.masters);
    if (!sums.length) throw Error('정산할 행이 없습니다.');
    const head = ['거래처', '주문 수', '낱알', '판매가 합', '정산가 합', '수수료', '수수료율', '원가', '반품·취소 차감', '예상 순이익', '원가 미등록 행'];
    const body = sums.map(s => [s.supplier, s.orders, s.pieces, s.sales, s.settle, s.fee, s.feeRate, s.cost, s.csDeduct, s.net, s.costMissing]);
    xlsxDownload([head, ...body], '정산요약', safeFile('정산요약_' + new Date().toLocaleDateString('en-CA') + '.xlsx'));
    toast('정산 요약을 내려받았습니다. 원가 미등록 행은 순이익 계산에서 제외되어 있습니다.');
  } catch (e) { toast(e.message); }
}
function openTemplate(index) {
  const s = suppliersOfShip()[index];
  if (!s) { toast('거래처를 찾을 수 없습니다.'); return; }
  const current = (templateFor(s) || ConfBench.DEFAULT_TEMPLATE).map(p => p[0] + '=' + p[1]).join('\n');
  const d = dialog('발주서 양식 · ' + esc(s), `<p class="small">한 줄에 “발주서 헤더=원장 열”을 씁니다. 원장 열은 기존 엑셀 열 이름과 같아야 합니다. 상품명은 <b>CONF_제품명</b>을 씁니다.</p><textarea id="tpl-text" rows="12" style="width:100%" aria-label="발주서 양식">${esc(current)}</textarea><p class="small">사용 가능한 원장 열: ${esc(L_COLUMNS_TEXT())}</p>`, '<span class="small">저장하면 이 거래처의 발주서에 바로 적용됩니다.</span><button class="primary" id="tpl-save">양식 저장</button>');
  $('#tpl-save').onclick = async () => {
    try {
      const pairs = ConfBench.parseTemplateText($('#tpl-text').value);
      await commit({ ...state, templates: { ...(state.templates || {}), [s]: pairs }, history: [logEvent('발주서 양식 수정', [], s + ' · ' + pairs.length + '열'), ...state.history].slice(0, 3000) });
      d.close(); render(); toast(s + ' 발주서 양식을 저장했습니다.');
    } catch (e) { toast(e.message); }
  };
}
function L_COLUMNS_TEXT() { return ConfLedger.COLUMNS.join(', '); }
function openTrackingPaste() {
  const d = dialog('송장 붙여넣기', '<p class="small">한 줄에 “주문번호, 택배사, 송장번호”를 씁니다(쉼표 또는 탭). 출고마감 행에만 적용하며, 이미 다른 송장이 있는 행은 덮어쓰지 않습니다. 적용된 행은 송장연결로 넘어갑니다.</p><textarea id="trk-text" rows="10" style="width:100%" aria-label="송장 목록" placeholder="20261109-0000001, 롯데택배, 268812345678"></textarea><div id="trk-preview" class="validation" hidden></div>', '<button id="trk-check">미리보기</button><button class="primary" id="trk-apply" disabled>적용</button>');
  const preview = () => {
    const p = ConfBench.parseTracking($('#trk-text').value);
    const shipIds = new Set(state.rows.filter(r => r.meta.status === '출고마감').map(r => String(r.values['주문고유번호'] || '').trim()));
    const matched = [...p.entries.keys()].filter(id => shipIds.has(id));
    const unmatched = [...p.entries.keys()].filter(id => !shipIds.has(id));
    const box = $('#trk-preview');
    box.hidden = false;
    box.innerHTML = `<b>적용 대상 ${matched.length}건</b> · 원장에 없거나 출고마감이 아닌 주문 ${unmatched.length}건 · 형식 오류 ${p.errors.length}줄` + (p.errors.length ? '<ul>' + p.errors.slice(0, 20).map(e => '<li>' + e.line + '줄: ' + esc(e.reason) + '</li>').join('') + '</ul>' : '') + (unmatched.length ? '<details><summary>대상이 아닌 주문번호</summary><ul>' + unmatched.slice(0, 50).map(x => '<li>' + esc(x) + '</li>').join('') + '</ul></details>' : '');
    $('#trk-apply').disabled = !matched.length;
    return p;
  };
  $('#trk-check').onclick = preview;
  $('#trk-apply').onclick = async () => {
    try {
      const p = ConfBench.parseTracking($('#trk-text').value);
      if (p.errors.length) throw Error('형식 오류 ' + p.errors.length + '줄을 먼저 고치세요.');
      const res = ConfBench.applyTracking(state.rows, p.entries, state.masters);
      if (!res.applied) throw Error('적용할 출고마감 행이 없습니다.');
      await commit(withHistory(res.rows, logEvent('송장 붙여넣기', [], res.applied + '행 송장연결 · 건너뜀 ' + res.skipped.length + ' · 원장에 없음 ' + res.unmatched.length)));
      selected.clear(); d.close(); render();
      toast(res.applied + '행을 송장연결로 넘겼습니다.' + (res.skipped.length ? '\n덮어쓰지 않고 건너뜀 ' + res.skipped.length + '건' : '') + (res.unmatched.length ? '\n원장에 없는 주문 ' + res.unmatched.length + '건' : ''));
    } catch (e) { toast(e.message); }
  };
}
function editMasters() {
  const list = (state.masters || []).map(m => ({ ...m, aliasText: (m.aliases || []).join(', '), costText: m.cost == null ? '' : String(m.cost), isUse: APPROVED.includes(m.status) }));
  const d = dialog('상품 기준 관리', `<p class="small">‘상세상품’ 이름으로 입수를 연결합니다. 별칭은 거래처가 쓰는 다른 상품명입니다(쉼표 구분). 원가는 낱알 1개당 매입가이며, 비우면 순이익을 계산하지 않습니다.</p><div class="table-wrap"><table><thead><tr><th>상세상품</th><th>입수</th><th>상태</th><th>별칭</th><th>낱알 원가(원)</th></tr></thead><tbody>${list.map((m, i) => `<tr><td><input data-m="name" data-i="${i}" value="${esc(m.name)}"></td><td><input data-m="pack" data-i="${i}" type="number" min="1" step="1" value="${esc(m.pack)}"></td><td><select data-m="status" data-i="${i}"><option ${m.isUse ? 'selected' : ''}>사용</option><option ${m.isUse ? '' : 'selected'}>검토필요</option></select></td><td><input data-m="aliases" data-i="${i}" value="${esc(m.aliasText)}"></td><td><input data-m="cost" data-i="${i}" type="number" min="0" step="0.01" value="${esc(m.costText)}"></td></tr>`).join('')}</tbody></table></div><div class="field-grid" style="margin-top:14px"><div class="field"><label>새 상세상품</label><input id="nm-name"></div><div class="field"><label>입수</label><input id="nm-pack" type="number" min="1" step="1"></div><div class="field"><label>별칭</label><input id="nm-alias"></div></div>`, '<span class="small">‘사용’은 실제 포장 입수를 확인했다는 뜻입니다.</span><button class="primary" id="mst-save">기준 저장</button>');
  $('#mst-save').onclick = async () => {
    try {
      const get = (k, i) => document.querySelector(`[data-m="${k}"][data-i="${i}"]`).value.trim();
      const out = [];
      list.forEach((m, i) => {
        const name = get('name', i);
        if (!name) return;
        const pack = Number(get('pack', i));
        if (!Number.isSafeInteger(pack) || pack <= 0) throw Error(name + ': 입수는 1 이상의 정수여야 합니다.');
        const c = get('cost', i);
        const cost = c === '' ? undefined : Number(c);
        if (cost !== undefined && (!Number.isFinite(cost) || cost < 0)) throw Error(name + ': 원가는 0 이상의 숫자여야 합니다.');
        const aliases = get('aliases', i).split(',').map(s => s.trim()).filter(Boolean);
        const status = get('status', i) === '사용' ? '사용' : '검토필요';
        out.push({ name, pack, status, aliases, ...(cost !== undefined ? { cost } : {}) });
      });
      const nn = $('#nm-name').value.trim();
      if (nn) {
        const pack = Number($('#nm-pack').value);
        if (!Number.isSafeInteger(pack) || pack <= 0) throw Error('새 상품 입수를 확인하세요.');
        out.push({ name: nn, pack, status: '사용', aliases: $('#nm-alias').value.split(',').map(s => s.trim()).filter(Boolean) });
      }
      const seen = new Set();
      for (const m of out) for (const k of [m.name, ...m.aliases]) {
        if (seen.has(k)) throw Error('상품명 또는 별칭이 중복됩니다: ' + k);
        seen.add(k);
      }
      await commit({ ...state, masters: out, history: [logEvent('상품 기준 수정', [], out.length + '개 기준 저장'), ...state.history].slice(0, 3000) });
      d.close(); render(); toast('상품 기준을 저장했습니다.');
    } catch (e) { toast(e.message); }
  };
}
async function saveCS() {
  try {
    const id = $('#cs-order').value.trim(), type = $('#cs-type').value, reason = $('#cs-reason').value.trim();
    const res = ConfBench.addCS(state.rows, id, type, reason, new Date().toISOString());
    await commit(withHistory(res.rows, logEvent('CS 기록', [], id + ' / ' + type + ' / ' + res.hit + '행' + (reason ? ' / ' + reason : ''))));
    render();
    toast(res.hit + '행에 ' + type + ' 기록을 남겼습니다.');
  } catch (e) { toast(e.message); }
}
