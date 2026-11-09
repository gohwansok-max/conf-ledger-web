/* Local benchmark features borrowed from order-distribution SaaS patterns:
 * supplier order sheets, supplier templates, alias matching, tracking paste,
 * settlement summary, and CS (return/exchange/cancel) records.
 * Pure functions. No network, no mail sending, no carrier issuance. */
(function (root, factory) {
  'use strict';
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./ledger-core.js'));
  else root.ConfBench = factory(root.ConfLedger);
})(typeof globalThis !== 'undefined' ? globalThis : this, function (L) {
  'use strict';
  const SHIP = ['출고마감', '송장연결'];
  const SETTLE_BASE = ['검토완료', '출고마감', '송장연결', '회신준비'];
  const CS_TYPES = ['반품', '교환', '취소'];
  const DEFAULT_TEMPLATE = [
    ['주문번호', '주문고유번호'], ['수취인', '수취인명'], ['연락처', '수취인 연락처'], ['우편번호', '우편번호'],
    ['주소', '주소'], ['상품', '상세상품'], ['낱알수량', '수량(알)'], ['배송메세지', '배송메세지'],
    ['택배사', '택배사'], ['송장번호', '송장번호'],
  ];
  const txt = v => (v === null || v === undefined) ? '' : String(v).trim();
  const isNum = v => typeof v === 'number' && Number.isFinite(v);
  const supplierOf = r => txt(r.values['구매처']) || txt(r.values['상세구매처']) || '미지정';
  const productOf = r => txt(r.values['상세상품']) || txt(r.values['제품명']);
  function masterFor(row, masters) {
    const name = productOf(row);
    return (masters || []).find(m => m && (txt(m.name) === name || (Array.isArray(m.aliases) && m.aliases.map(txt).includes(name)))) || null;
  }
  function validTemplate(template) {
    const rows = Array.isArray(template) && template.length ? template : DEFAULT_TEMPLATE;
    if (rows.length > 40) throw new Error('발주서 열은 40개 이하여야 합니다.');
    for (const pair of rows) {
      if (!Array.isArray(pair) || pair.length !== 2) throw new Error('양식은 [헤더, 원장열] 쌍이어야 합니다.');
      const [header, source] = pair.map(txt);
      if (!header || header.length > 40) throw new Error('발주서 헤더는 1~40자여야 합니다.');
      if (!L.COLUMNS.includes(source) && source !== 'CONF_제품명') throw new Error('원장에 없는 열을 양식에 넣을 수 없습니다: ' + source);
    }
    return rows.map(p => [txt(p[0]), txt(p[1])]);
  }
  function parseTemplateText(text) {
    const pairs = [];
    for (const raw of String(text || '').split(/\r?\n/)) {
      const line = raw.trim();
      if (!line) continue;
      const i = line.indexOf('=');
      if (i < 1) throw new Error('양식 줄은 "발주서 헤더=원장 열" 형식이어야 합니다: ' + line.slice(0, 40));
      pairs.push([line.slice(0, i).trim(), line.slice(i + 1).trim()]);
    }
    return validTemplate(pairs);
  }
  function supplierSummary(rows, masters) {
    const map = new Map();
    for (const r of rows) {
      if (!SETTLE_BASE.includes(r.meta.status)) continue;
      const key = supplierOf(r);
      if (!map.has(key)) map.set(key, { supplier: key, rows: 0, orders: new Set(), pieces: 0, sales: 0, settle: 0, shipFee: 0, cost: 0, costMissing: 0, unknownSales: 0, csDeduct: 0, csCount: 0 });
      const s = map.get(key);
      s.rows++;
      if (txt(r.values['주문고유번호'])) s.orders.add(txt(r.values['주문고유번호']));
      const pieces = r.values['수량(알)'], sales = r.values['판매가'], settle = r.values['정산가'], ship = r.values['배송비'];
      if (isNum(pieces)) s.pieces += pieces;
      if (isNum(sales)) s.sales += sales; else s.unknownSales++;
      if (isNum(settle)) s.settle += settle;
      if (isNum(ship)) s.shipFee += ship;
      const m = masterFor(r, masters);
      if (m && isNum(m.cost) && isNum(pieces)) s.cost += pieces * m.cost; else s.costMissing++;
      if (r.meta.cs && (r.meta.cs.type === '반품' || r.meta.cs.type === '취소') && isNum(settle)) { s.csDeduct += settle; s.csCount++; }
    }
    return [...map.values()].map(s => {
      const fee = s.sales - s.settle;
      const net = s.costMissing ? null : s.settle - s.csDeduct - s.cost;
      return {
        supplier: s.supplier, rows: s.rows, orders: s.orders.size, pieces: s.pieces, sales: s.sales, settle: s.settle,
        fee, feeRate: s.sales > 0 ? fee / s.sales : null, shipFee: s.shipFee, cost: s.costMissing ? null : s.cost,
        costMissing: s.costMissing, unknownSales: s.unknownSales, csCount: s.csCount, csDeduct: s.csDeduct,
        net,
      };
    }).sort((a, b) => b.sales - a.sales);
  }
  function supplierOrderSheet(rows, supplier, template) {
    const pairs = validTemplate(template);
    const header = pairs.map(p => p[0]);
    const body = rows.filter(r => SHIP.includes(r.meta.status) && supplierOf(r) === supplier).map(r => pairs.map(([, src]) => {
      const v = src === 'CONF_제품명' ? productOf(r) : r.values[src];
      if (v === null || v === undefined) return '';
      return typeof v === 'string' && /^[=+\-@]/.test(v.trimStart()) ? "'" + v : v;
    }));
    return [header, ...body];
  }
  function parseTracking(text) {
    const entries = new Map(), errors = [];
    const lines = String(text || '').split(/\r?\n/);
    lines.forEach((raw, i) => {
      const line = raw.trim();
      if (!line) return;
      const parts = line.split(/\t|,/).map(txt);
      if (parts[0] === '주문번호' || parts[0] === '주문고유번호') return;
      const [orderId, carrier, tracking] = parts;
      if (!orderId || !carrier || !tracking) errors.push({ line: i + 1, reason: '주문번호, 택배사, 송장번호 3칸이 필요합니다.' });
      else if (!/^[0-9A-Za-z-]{6,30}$/.test(tracking)) errors.push({ line: i + 1, reason: '송장번호 형식을 확인하세요(영문·숫자·하이픈 6~30자).' });
      else if (entries.has(orderId) && (entries.get(orderId).tracking !== tracking || entries.get(orderId).carrier !== carrier)) errors.push({ line: i + 1, reason: '같은 주문번호에 다른 송장이 두 번 입력되었습니다.' });
      else entries.set(orderId, { carrier, tracking });
    });
    return { entries, errors };
  }
  function applyTracking(rows, entries, masters) {
    const result = { rows: [], applied: 0, shipped: 0, skipped: [], unmatched: new Set(entries.keys()) };
    for (const r of rows) {
      const id = txt(r.values['주문고유번호']);
      if (!entries.has(id) || r.meta.status !== '출고마감') { result.rows.push(r); continue; }
      result.unmatched.delete(id);
      const e = entries.get(id);
      const existing = txt(r.values['송장번호']);
      if (existing && existing !== e.tracking) { result.skipped.push({ id, reason: '이미 다른 송장번호가 있어 덮어쓰지 않았습니다.' }); result.rows.push(r); continue; }
      try {
        const values = { ...r.values, '택배사': e.carrier, '송장번호': e.tracking };
        let next = L.update(r, values, '송장 붙여넣기 입력', masters).row;
        next = L.transition(next, '송장연결', masters);
        result.rows.push(next); result.applied++; result.shipped++;
      } catch (err) {
        result.skipped.push({ id, reason: err.message });
        result.rows.push(r);
      }
    }
    result.unmatched = [...result.unmatched];
    return result;
  }
  function addCS(rows, orderId, type, reason, at) {
    if (!CS_TYPES.includes(type)) throw new Error('CS 유형은 반품·교환·취소 중 하나여야 합니다.');
    const note = txt(reason).slice(0, 200);
    const id = txt(orderId);
    if (!id) throw new Error('주문번호를 입력하세요.');
    let hit = 0;
    const out = rows.map(r => {
      if (txt(r.values['주문고유번호']) !== id) return r;
      hit++;
      return { ...r, meta: { ...r.meta, revision: r.meta.revision + 1, updatedAt: at, cs: { type, reason: note, at } } };
    });
    if (!hit) throw new Error('같은 주문번호의 원장 행이 없습니다.');
    return { rows: out, hit };
  }
  return { SHIP, SETTLE_BASE, CS_TYPES, DEFAULT_TEMPLATE, supplierOf, productOf, masterFor, validTemplate, parseTemplateText, supplierSummary, supplierOrderSheet, parseTracking, applyTracking, addCS };
});
