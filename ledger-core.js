/* CONF ledger core. No IO, XLSX dependency, issuance or carrier integration.
 * Values stay typed: numeric strings are not silently coerced.
 * Master status approved / 승인 / 확정 / 사용 is approved; all others require review.
 * Master name matches nonblank 상세상품. 제품명 is fallback only when 상세상품
 * is blank; an unknown detail name never silently falls back to a product group.
 * Money inputs: nonnegative, at most 2 decimal places, safe integer minor units.
 * Derived unit price and fee are unrounded quotients, not new input prices.
 * Dates accept valid Date objects, Excel date serials 1..2958465, or calendar
 * strings YYYY-MM-DD / YYYY.MM.DD / YYYY/MM/DD (optional T or space time
 * suffix). Calendar and time components are checked without timezone conversion.
 * summary: rows = all detail rows, active = not hold/cancel, ready = validated
 * active rows. Only ready rows contribute units/sales/settlement. issues counts
 * all rows with either errors or warnings. ordersKnown counts distinct ACTIVE
 * [스토어, 구매처, 상세구매처, 주문고유번호]; ordersUnknown counts ACTIVE detail
 * rows without order number or channel. Orders are not shipment counts.
 * fromOriginal maps AI (index 34, originally duplicate 비고) to 품목번호 ONLY
 * for an actual string. Other nonblank AI values stay in meta.notes, untouched.
 * recalc is separate from update. Invalid derived inputs produce null, never
 * guessed packs/prices. Existing tracking is retained and flagged on invalidation.
 */
(function (root, factory) {
  'use strict';
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.ConfLedger = factory();
})(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';
  const COLUMNS = Object.freeze(['원본시트','스토어','월','주','발주일','주문일','유,무상','구매처','상세구매처','제품명','상세상품','수량','수량(알)','무상출고','정산가','판매가','배송비','단가','수수료율','행사','비고','벌크','띠지','구매자명','구매자연락처','수취인명','수취인 연락처','우편번호','주소','배송메세지','증빙 여부','택배사','송장번호','주문고유번호','품목번호']);
  const STATES = Object.freeze(['접수','검토완료','출고마감','송장연결','회신준비','보류','취소']);
  const FLOW = STATES.slice(0, 5);
  const NUMERIC = new Set(['수량','수량(알)','정산가','판매가','배송비','단가','수수료율']);
  const MONEY = new Set(['정산가','판매가','배송비']);
  const IMPORTANT = new Set(['스토어','구매처','상세구매처','발주일','주문일','유,무상','제품명','상세상품','수량','수량(알)','무상출고','정산가','판매가','배송비','단가','수수료율','행사','벌크','띠지','구매자명','구매자연락처','수취인명','수취인 연락처','우편번호','주소','배송메세지','주문고유번호','품목번호']);
  const MAX = Number.MAX_SAFE_INTEGER;
  const MAX_MINOR = BigInt(MAX);
  const blank = value => value === null || value === undefined || (typeof value === 'string' && value.trim() === '');
  const text = value => typeof value === 'string' ? value.trim() : (typeof value === 'number' && Number.isFinite(value) ? String(value) : '');
  const positive = value => Number.isSafeInteger(value) && value > 0;
  const bounded = value => typeof value === 'number' && Number.isFinite(value) && Math.abs(value) <= MAX;
  const formula = value => typeof value === 'string' && /^[=+@]/.test(value.trimStart());
  const active = row => row?.meta?.status !== '보류' && row?.meta?.status !== '취소';
  function clone(value) {
    if (Array.isArray(value)) return value.map(clone);
    if (value instanceof Date) return new Date(value.getTime());
    if (value && typeof value === 'object') {
      const out = {};
      for (const key of Object.keys(value)) Object.defineProperty(out, key, { value: clone(value[key]), writable: true, enumerable: true, configurable: true });
      return out;
    }
    return value;
  }
  function minor(value) {
    if (!bounded(value) || value < 0) return null;
    const match = String(value).match(/^(\d+)(?:\.(\d+))?(?:e([+-]?\d+))?$/i);
    if (!match) return null;
    const fraction = match[2] || '';
    const scale = fraction.length - Number(match[3] || 0);
    let result = BigInt(match[1] + fraction);
    if (scale > 2) {
      const divisor = 10n ** BigInt(scale - 2);
      if (result % divisor !== 0n) return null;
      result /= divisor;
    } else result *= 10n ** BigInt(2 - scale);
    return result <= MAX_MINOR ? result : null;
  }
  function validDate(value) {
    if (value instanceof Date) return Number.isFinite(value.getTime());
    if (typeof value === 'number') return Number.isSafeInteger(value) && value >= 1 && value <= 2958465;
    if (typeof value !== 'string') return false;
    const match = value.trim().match(/^(\d{4})([-/.])(\d{1,2})\2(\d{1,2})(?:[T ](\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,3}))?)?(Z|[+-]\d{2}:\d{2})?)?$/);
    if (!match) return false;
    const year = Number(match[1]), month = Number(match[3]), day = Number(match[4]);
    const leap = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
    const days = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
    if (!(year >= 1 && month >= 1 && month <= 12 && day >= 1 && day <= days[month - 1])) return false;
    if (match[5] !== undefined && (Number(match[5]) > 23 || Number(match[6]) > 59 || Number(match[7] || 0) > 59)) return false;
    const zone = match[9];
    if (zone && zone !== 'Z') {
      const zoneHour = Number(zone.slice(1, 3)), zoneMinute = Number(zone.slice(4, 6));
      if (zoneHour > 14 || zoneMinute > 59 || (zoneHour === 14 && zoneMinute !== 0)) return false;
    }
    return true;
  }
  function approved(master) { return ['approved','승인','확정','사용'].includes(master?.status); }
  function resolveMaster(values, masters) {
    const name = text(blank(values['상세상품']) ? values['제품명'] : values['상세상품']);
    const found = Array.isArray(masters) ? masters.filter(master => master && name !== '' && (text(master.name) === name || (Array.isArray(master.aliases) && master.aliases.map(text).includes(name)))) : [];
    if (!found.length) return { error: '제품 마스터가 없습니다. 등록 및 승인 후 검토하세요.' };
    if (found.length !== 1) return { error: '동일 제품 마스터가 중복되었습니다. 하나의 승인 기준으로 정리하세요.' };
    if (!approved(found[0])) return { error: '제품 마스터가 검토 중이거나 미승인입니다.' };
    if (!positive(found[0].pack)) return { error: '승인 마스터의 입수는 안전한 양의 정수여야 합니다.' };
    return { master: found[0] };
  }
  function isFree(values) {
    return text(values['유,무상']) === '무상' || values['무상출고'] === true || ['무상','Y','예'].includes(text(values['무상출고']));
  }
  function requireRow(row) {
    if (!row || typeof row.id !== 'string' || blank(row.id) || !row.values || typeof row.values !== 'object' || Array.isArray(row.values) || !row.meta || !STATES.includes(row.meta.status) || !Number.isSafeInteger(row.meta.revision) || row.meta.revision < 0 || !Array.isArray(row.meta.notes) || typeof row.meta.confirmed !== 'boolean') throw new Error('행 구조 또는 메타데이터가 올바르지 않습니다.');
    if (Object.keys(row.values).some(field => !COLUMNS.includes(field))) throw new Error('알 수 없는 원장 열이 포함되어 있습니다.');
    if (COLUMNS.some(field => !Object.prototype.hasOwnProperty.call(row.values, field))) throw new Error('원장 35개 열이 모두 필요합니다.');
  }
  function stamp(row) {
    if (row.meta.revision >= MAX) throw new Error('행 개정번호의 안전 범위를 초과했습니다.');
    row.meta.revision += 1;
    row.meta.updatedAt = new Date().toISOString();
  }
  function createRow({ id, values = {}, meta = {} } = {}) {
    if (Object.keys(values).some(field => !COLUMNS.includes(field))) throw new Error('알 수 없는 원장 열이 포함되어 있습니다.');
    const row = { id, values: {}, meta: { sourceFile: '', sourceSheet: '', sourceRow: null, status: '접수', revision: 0, updatedAt: null, confirmed: false, notes: [], ...clone(meta) } };
    for (const field of COLUMNS) row.values[field] = Object.prototype.hasOwnProperty.call(values, field) ? clone(values[field]) : '';
    requireRow(row);
    return row;
  }
  function fromOriginal({ id, cells, meta = {} } = {}) {
    if (!Array.isArray(cells) || cells.length !== COLUMNS.length) throw new Error('원본 행은 35개 셀이어야 합니다.');
    const values = {};
    COLUMNS.forEach((field, index) => { values[field] = index === 34 && typeof cells[index] !== 'string' ? '' : clone(cells[index]); });
    const row = createRow({ id, values, meta });
    if (typeof cells[34] !== 'string' && !blank(cells[34])) row.meta.notes.push({ kind: 'original-ai-note', column: 'AI', originalHeader: '비고', value: clone(cells[34]) });
    return row;
  }
  function validate(row, masters = []) {
    const errors = [], warnings = [];
    try { requireRow(row); } catch (_) { return { errors: ['행 구조 또는 메타데이터가 올바르지 않습니다.'], warnings }; }
    const values = row.values;
    for (const field of ['제품명','발주일','주문일','수취인명','수취인 연락처','주소']) {
      if (blank(values[field])) errors.push(field + '이(가) 누락되었습니다.');
      else if (typeof values[field] !== 'string' && !(field === '발주일' || field === '주문일')) errors.push(field + '은(는) 문자 값으로 입력하세요.');
    }
    for (const field of ['발주일','주문일']) if (!blank(values[field]) && !validDate(values[field])) errors.push(field + '은(는) 유효한 날짜여야 합니다.');
    if (!positive(values['수량'])) errors.push('수량은 안전한 양의 정수여야 합니다.');
    for (const field of NUMERIC) {
      const value = values[field];
      if (blank(value)) continue;
      if (formula(value)) errors.push(field + '에 수식 입력을 사용할 수 없습니다.');
      else if (!bounded(value)) errors.push(field + '은(는) 유한한 안전 범위의 숫자여야 합니다.');
    }
    for (const field of ['판매가','정산가']) {
      if (blank(values[field])) errors.push(field + '이(가) 누락되었습니다. 0과 빈 값은 다릅니다.');
      else if (minor(values[field]) === null) errors.push(field + '은(는) 0 이상, 소수 둘째 자리 이내의 안전한 금액이어야 합니다.');
    }
    if (!blank(values['배송비']) && minor(values['배송비']) === null) errors.push('배송비는 0 이상, 소수 둘째 자리 이내의 안전한 금액이어야 합니다.');
    if (blank(values['수량(알)'])) errors.push('수량(알)이 누락되었습니다. 승인 입수로 재계산한 뒤 검토하세요.');
    else if (!positive(values['수량(알)'])) errors.push('수량(알)은 안전한 양의 정수여야 합니다.');
    if (!blank(values['단가']) && (!bounded(values['단가']) || values['단가'] < 0)) errors.push('단가는 0 이상의 안전한 숫자여야 합니다.');
    const resolved = resolveMaster(values, masters);
    if (resolved.error) errors.push(resolved.error);
    else if (positive(values['수량'])) {
      const pieces = values['수량'] * resolved.master.pack;
      if (!Number.isSafeInteger(pieces)) errors.push('수량과 승인 입수의 곱이 안전한 정수 범위를 초과했습니다.');
      else if (!blank(values['수량(알)']) && values['수량(알)'] !== pieces) errors.push('수량(알)이 수량과 승인 입수의 곱과 일치하지 않습니다.');
    }
    if (isFree(values)) {
      if (minor(values['판매가']) !== null && values['판매가'] !== 0) errors.push('무상 행의 판매가는 0이어야 합니다. 유·무상 구분을 확인하세요.');
      if (minor(values['정산가']) !== null && values['정산가'] !== 0) errors.push('무상 행의 정산가는 0이어야 합니다. 유·무상 구분을 확인하세요.');
    } else if (values['판매가'] === 0) warnings.push('판매가가 0입니다. 무상 여부를 확인하세요.');
    if (minor(values['판매가']) !== null && minor(values['정산가']) !== null && values['정산가'] > values['판매가']) warnings.push('정산가가 판매가보다 큽니다. 계약 기준을 확인하세요.');
    if (blank(values['주문고유번호'])) warnings.push('주문고유번호가 없습니다. 주문 수 집계와 중복 자동 병합에서 제외됩니다.');
    else if (!text(values['주문고유번호'])) errors.push('주문고유번호는 문자 또는 유한한 숫자여야 합니다.');
    if (channel(values).every(blank)) warnings.push('판매 채널이 없습니다. 주문 수 집계와 중복 자동 병합에서 제외됩니다.');
    const stage = FLOW.indexOf(row.meta.status);
    if (stage >= 2 && row.meta.confirmed !== true) errors.push('출고마감 이후에는 출고 확인이 필요합니다.');
    if (stage >= 3 && !tracking(values['송장번호'])) errors.push('송장연결 이후에는 실제 입력한 송장번호가 필요합니다.');
    if (stage >= 3 && (typeof values['택배사'] !== 'string' || blank(values['택배사']) || formula(values['택배사']))) errors.push('송장연결 이후에는 택배사 입력이 필요합니다.');
    return { errors: [...new Set(errors)], warnings: [...new Set(warnings)] };
  }
  function recalc(row, masters = []) {
    requireRow(row);
    const next = clone(row), values = next.values, resolved = resolveMaster(values, masters);
    const quantity = values['수량'];
    const pieces = resolved.master && positive(quantity) ? quantity * resolved.master.pack : null;
    values['수량(알)'] = Number.isSafeInteger(pieces) && pieces > 0 ? pieces : null;
    const unit = positive(quantity) && minor(values['판매가']) !== null ? values['판매가'] / quantity : null;
    values['단가'] = bounded(unit) ? unit : null;
    const fee = values['판매가'] > 0 && minor(values['판매가']) !== null && minor(values['정산가']) !== null ? (values['판매가'] - values['정산가']) / values['판매가'] : null;
    values['수수료율'] = bounded(fee) ? fee : null;
    return next;
  }
  function tracking(value) {
    return !blank(value) && (typeof value === 'string' || (Number.isSafeInteger(value) && value > 0)) && !formula(value) && !['0','-','없음','미입력','미발급','대기','N/A','NA','NULL','UNDEFINED','TBD'].includes(text(value).toUpperCase());
  }
  function transition(row, target, masters = []) {
    requireRow(row);
    if (!STATES.includes(target)) throw new Error('알 수 없는 업무 상태입니다.');
    const current = row.meta.status;
    if (current === target) { if (active(row) && FLOW.indexOf(target) > 0 && validate(row, masters).errors.length) throw new Error('현재 상태의 검증 오류를 먼저 해결하세요.'); return clone(row); }
    const next = clone(row);
    if (target !== '보류' && target !== '취소') {
      if (current === '보류' || current === '취소') {
        if (target !== '접수') throw new Error('보류 또는 취소 행은 접수로만 재활성화할 수 있습니다.');
        next.meta.confirmed = false;
      } else if (FLOW.indexOf(target) !== FLOW.indexOf(current) + 1) throw new Error('업무 단계를 건너뛰거나 역행할 수 없습니다.');
      next.meta.status = target;
      if (target !== '접수') {
        const result = validate(next, masters);
        if (result.errors.length) throw new Error('상태 변경 불가: ' + result.errors.join(' '));
      }
    } else next.meta.status = target;
    stamp(next);
    next.meta.notes.push('상태 변경: ' + current + ' → ' + target);
    return next;
  }
  function update(row, patch, reason, masters = []) {
    requireRow(row);
    if (!patch || typeof patch !== 'object' || Array.isArray(patch)) throw new Error('변경 값은 원장 열 이름을 키로 하는 객체여야 합니다.');
    if (typeof reason !== 'string' || blank(reason)) throw new Error('변경 사유를 입력하세요.');
    const next = clone(row), changes = [], previousStatus = row.meta.status;
    for (const field of Object.keys(patch)) {
      if (!COLUMNS.includes(field)) throw new Error('알 수 없는 원장 열은 수정할 수 없습니다.');
      const after = patch[field], before = row.values[field];
      if (after !== null && after !== undefined && !['string','number','boolean'].includes(typeof after)) throw new Error('셀 값은 문자, 숫자, 불리언 또는 빈 값이어야 합니다.');
      if (NUMERIC.has(field) && !blank(after)) {
        if (formula(after)) throw new Error(field + '에 수식 입력을 사용할 수 없습니다.');
        if (!bounded(after) || (field !== '수수료율' && after < 0)) throw new Error(field + '은(는) 안전한 숫자로 입력하세요.');
        if (['수량','수량(알)'].includes(field) && !Number.isSafeInteger(after)) throw new Error(field + '은(는) 안전한 정수로 입력하세요.');
        if (MONEY.has(field) && minor(after) === null) throw new Error(field + '의 금액 정밀도 또는 안전 범위를 확인하세요.');
      }
      const sameValue = (blank(before) && blank(after)) || Object.is(before, after);
      if (!sameValue) { next.values[field] = after; changes.push({ field, before: clone(before), after: clone(after) }); }
    }
    if (!changes.length) return { row: next, changes, previousStatus };
    const important = changes.some(change => IMPORTANT.has(change.field) || (['송장번호','택배사'].includes(change.field) && FLOW.indexOf(previousStatus) >= 3));
    if (important) {
      next.meta.confirmed = false;
      if (FLOW.indexOf(previousStatus) > 0) {
        next.meta.status = '접수';
        next.meta.notes.push('중요 항목 변경으로 후속 상태를 접수로 되돌렸습니다. 재검토가 필요합니다.');
        if (!blank(next.values['송장번호'])) next.meta.notes.push('기존 송장번호를 보존했습니다. 변경된 주문과의 일치 여부를 확인하세요.');
      }
    }
    stamp(next);
    next.meta.notes.push('변경 사유: ' + reason.trim());
    // masters is intentionally not used to guess values or automatically recalculate.
    return { row: next, changes, previousStatus };
  }
  function channel(values) { return ['스토어','구매처','상세구매처'].map(field => text(values[field])); }
  function orderKey(row) {
    const values = row?.values || {}, order = text(values['주문고유번호']), scope = channel(values);
    return order && scope.some(value => value !== '') ? JSON.stringify([scope, order]) : null;
  }
  function duplicateKey(row) {
    const order = orderKey(row);
    if (!order) return null;
    return JSON.stringify([JSON.parse(order)[0], text(row.values['주문고유번호']), text(row.values['제품명']), text(row.values['상세상품'])]);
  }
  function summary(rows, masters = []) {
    if (!Array.isArray(rows)) throw new Error('집계 대상은 행 배열이어야 합니다.');
    const result = { rows: rows.length, active: 0, issues: 0, units: 0, sales: 0, settlement: 0, ordersKnown: 0, ordersUnknown: 0, ready: 0, excluded: 0, scope: 'validated-active' };
    const orders = new Set(); let sales = 0n, settlement = 0n;
    for (const row of rows) {
      const checked = validate(row, masters);
      if (checked.errors.length || checked.warnings.length) result.issues++;
      if (!active(row)) continue;
      result.active++;
      const key = orderKey(row);
      if (key) orders.add(key); else result.ordersUnknown++;
      if (checked.errors.length) { result.excluded++; continue; }
      const calculated = recalc(row, masters);
      const units = calculated.values['수량(알)'];
      if (!positive(units) || !Number.isSafeInteger(result.units + units)) throw new Error('낱알 수량 합계의 안전 범위를 초과했습니다.');
      result.units += units;
      sales += minor(row.values['판매가']); settlement += minor(row.values['정산가']);
      if (sales > MAX_MINOR || settlement > MAX_MINOR) throw new Error('금액 합계의 안전 범위를 초과했습니다.');
      result.ready++;
    }
    result.sales = Number(sales) / 100; result.settlement = Number(settlement) / 100; result.ordersKnown = orders.size;
    return result;
  }
  return Object.freeze({ COLUMNS, STATES, createRow, fromOriginal, validate, recalc, transition, update, duplicateKey, summary });
});
