/* 원본데이터 ⇄ 결과 엑셀 변환 + 순위 자동 산정 (브라우저/Node 공용)
 * 읽기: SheetJS(XLSX)  /  쓰기: ExcelJS
 */
(function (root) {
  const MAIN = '해외 공연 현황', SUB = '해외 공연 현황 (2)', LIST = '연도별 아티스트 목록', VARS = '변수 설명';
  const FIELDS = ['year', 'artist', 'type', 'start_date', 'end_date', 'days', 'country', 'city', 'location', 'capacity', 'attendance', 'attendance_label', 'ticket_price', 'remarks', 'gender', 'member', 'deyear', 'demon', 'deday', 'agency', 'agency2', 'agency3', 'attendance/days', 'attendance_rule', '집계묶음', '출처'];
  const MAIN_HDR = ['no', 'year', 'qt', 'artist', 'type', 'start_date', 'end_date', 'days', 'country', 'city', 'location', 'capacity', 'capacity*days', 'attendance', 'attendance/days', 'ticket_price', 'remarks', 'gender', 'member', 'deyear', 'demon', 'deday', 'agency', 'agency2', 'agency3'];
  // (2) 시트도 본 시트와 같은 25개 변수 (2026-10-21 박사님 요청: 두 시트 변수 동일)
  const SUB_HDR = MAIN_HDR.slice();
  const FILL_PLAN = 'FFFFF2CC', FILL_PART = 'FFFCE4D6';

  const norm = (s) => String(s == null ? '' : s).toLowerCase().replace(/[-–—_.·]/g, ' ').replace(/\bthe\b/g, '').replace(/\s+/g, ' ').trim();
  const isNum = (v) => typeof v === 'number' && isFinite(v);
  const serialToDate = (v) => {
    if (v instanceof Date) return new Date(Date.UTC(v.getUTCFullYear(), v.getUTCMonth(), v.getUTCDate()));
    if (!isNum(v)) return null;
    const p = root.XLSX.SSF.parse_date_code(v);
    return p ? new Date(Date.UTC(p.y, p.m - 1, p.d)) : null;
  };
  const fmt = (d) => d ? d.toISOString().slice(0, 10) : '';
  const stripTag = (s) => String(s || '').replace(/^\[(예정|일부 예정)[^\]]*\]\s*/, '');
  const labelFromFmt = (z) => { const m = /"\s*\(([^)]+)\)\s*"/.exec(z || ''); return m ? m[1] : ''; };

  function readTable(ws) {
    const XLSX = root.XLSX;
    const rg = XLSX.utils.decode_range(ws['!ref']);
    const hdr = [];
    for (let c = rg.s.c; c <= rg.e.c; c++) { const cell = ws[XLSX.utils.encode_cell({ r: rg.s.r, c })]; hdr.push(cell ? String(cell.v).trim() : ''); }
    const rows = [];
    for (let r = rg.s.r + 1; r <= rg.e.r; r++) {
      const o = {}; let any = false;
      hdr.forEach((h, i) => {
        if (!h) return;
        const cell = ws[XLSX.utils.encode_cell({ r, c: rg.s.c + i })];
        if (!cell || (cell.t === 'z' && !cell.f)) return;
        o[h] = cell.v; o['__z_' + h] = cell.z; if (cell.f) o['__f_' + h] = cell.f; any = true;
      });
      if (any) { Object.defineProperty(o, '__row', { value: r + 1, enumerable: false, writable: true }); rows.push(o); }
    }
    return rows;
  }
  // 표 영역만 읽기: 앞쪽의 빈 행·빈 열은 잘라내고 origin을 실제 시작 셀로 맞춤
  // (시트의 사용 범위가 A1로 기록돼 있어도 B2부터 시작하는 표를 똑같이 읽도록)
  const aoa = (ws) => {
    const s = root.XLSX.utils.decode_range(ws['!ref']).s;
    let rows = root.XLSX.utils.sheet_to_json(ws, { header: 1, defval: null, raw: true, blankrows: true });
    const filled = (v) => v !== null && v !== undefined && v !== '';
    let r0 = rows.findIndex(r => r.some(filled)); if (r0 < 0) r0 = 0;
    rows = rows.slice(r0);
    const c0 = Math.min(...rows.map(r => { const i = r.findIndex(filled); return i < 0 ? Infinity : i; }));
    const cut = isFinite(c0) ? c0 : 0;
    return { origin: { r: s.r + r0, c: s.c + cut }, rows: rows.map(r => r.slice(cut)) };
  };

  /* ---------- ① 기존 결과 엑셀 → 원본데이터 ---------- */
  function convertLegacy(wb) {
    const rows = readTable(wb.Sheets[MAIN]);
    let cut = rows.length;
    for (let i = 1; i < rows.length; i++) if (isNum(rows[i - 1].year) && isNum(rows[i].year) && rows[i].year < rows[i - 1].year) { cut = i; break; }
    const toSrc = (r) => {
      const o = {};
      FIELDS.forEach(f => o[f] = null);
      ['year', 'artist', 'type', 'days', 'country', 'city', 'location', 'capacity', 'attendance', 'ticket_price', 'gender', 'member', 'deyear', 'demon', 'deday', 'agency', 'agency2', 'agency3'].forEach(f => { if (r[f] !== undefined) o[f] = r[f]; });
      // 날짜가 아닌 텍스트(원자료 오기)는 그대로 보존 → 검증 탭에서 '날짜' 오류로 잡힘
      o.start_date = serialToDate(r.start_date) || (r.start_date ?? null); o.end_date = serialToDate(r.end_date) || (r.end_date ?? null);
      o.remarks = stripTag(r.remarks) || null;
      o.attendance_label = labelFromFmt(r['__z_attendance']) || null;
      o['attendance/days'] = isNum(r['attendance/days']) ? r['attendance/days'] : null;
      // 원자료에서 attendance가 '=수용인원×일수' 수식이던 행: 값 대신 규칙으로 보존
      if (r['__f_attendance']) { o.attendance = null; o.attendance_rule = '수용인원×일수(원자료 수식)'; }
      return o;
    };
    const venues = new Map();
    rows.forEach(r => {
      if (!r.location) return;
      const k = norm(r.location) + '|' + norm(r.city);
      if (!venues.has(k)) venues.set(k, { location: r.location, city: r.city, country: r.country, vals: new Set() });
      if (isNum(r.capacity)) venues.get(k).vals.add(r.capacity);
    });
    const master = [...venues.values()].map(v => ({ location: v.location, city: v.city, country: v.country, capacity: v.vals.size === 1 ? [...v.vals][0] : null, source: null, '기존 기입값': [...v.vals].sort((a, b) => a - b).join(' / ') }));
    return {
      main: rows.slice(0, cut).map(toSrc),
      appendix: rows.slice(cut).map(toSrc),
      master,
      official: [],
      list: wb.Sheets[LIST] ? aoa(wb.Sheets[LIST]) : null,
      vars: wb.Sheets[VARS] ? aoa(wb.Sheets[VARS]) : null,
      settings: { 기준일: fmt(new Date()), 자동순위연도: '2026', 상위N: 15 },
    };
  }

  async function writeSource(src) {
    const wb = new root.ExcelJS.Workbook();
    const addTable = (name, cols, data, widths) => {
      const ws = wb.addWorksheet(name);
      ws.addRow(cols); ws.getRow(1).font = { bold: true };
      data.forEach(d => ws.addRow(cols.map(c => d[c] === undefined ? null : d[c])));
      cols.forEach((c, i) => { ws.getColumn(i + 1).width = (widths && widths[c]) || 12; });
      ['start_date', 'end_date'].forEach(c => { const i = cols.indexOf(c); if (i >= 0) ws.getColumn(i + 1).numFmt = 'yyyy-mm-dd'; });
      ws.views = [{ state: 'frozen', ySplit: 1 }];
      return ws;
    };
    const W = { artist: 16, location: 32, remarks: 50, 출처: 40, city: 14, 집계묶음: 16 };
    addTable('공연목록', FIELDS, src.main, W);
    addTable('솔로부록', FIELDS, src.appendix, W);
    addTable('공연장마스터', ['location', 'city', 'country', 'capacity', 'source', '기존 기입값'], src.master, { location: 36, source: 40, '기존 기입값': 24 });
    addTable('공식누적', ['집계묶음', 'year', 'artist', '공식관객수', '출처', '설명'], src.official, { 출처: 50, 설명: 40, 집계묶음: 18 });
    const put = (name, a) => { const ws = wb.addWorksheet(name); if (!a) return; a.rows.forEach((row, i) => row.forEach((v, j) => { if (v !== null && v !== undefined) ws.getCell(a.origin.r + 1 + i, a.origin.c + 1 + j).value = v; })); };
    put('연도별목록', src.list); put('변수설명', src.vars);
    const st = wb.addWorksheet('설정');
    st.addRow(['항목', '값', '설명']); st.getRow(1).font = { bold: true };
    st.addRow(['기준일', src.settings.기준일, '이 날짜 이후 공연은 [예정], 걸치면 [일부 예정]으로 표시']);
    st.addRow(['자동순위연도', src.settings.자동순위연도, '쉼표로 여러 연도. 이 연도는 순위를 자동 산정하고 상위N만 상세 수록']);
    st.addRow(['상위N', src.settings.상위N, '자동 순위 연도에 상세 행으로 넣을 팀 수']);
    st.getColumn(1).width = 14; st.getColumn(2).width = 14; st.getColumn(3).width = 60;
    return wb.xlsx.writeBuffer();
  }

  /* ---------- ② 원본데이터 읽기 ---------- */
  function readSource(wb) {
    const need = ['공연목록', '공연장마스터'];
    const miss = need.filter(n => !wb.Sheets[n]);
    if (miss.length) throw new Error(`원본데이터 시트가 없어요: ${miss.join(', ')}`);
    const fix = (r) => { r.start_date = serialToDate(r.start_date) || (r.start_date ?? null); r.end_date = serialToDate(r.end_date) || (r.end_date ?? null); return r; };
    const settings = {};
    if (wb.Sheets['설정']) readTable(wb.Sheets['설정']).forEach(r => settings[r['항목']] = r['값']);
    return {
      main: readTable(wb.Sheets['공연목록']).map(fix),
      appendix: wb.Sheets['솔로부록'] ? readTable(wb.Sheets['솔로부록']).map(fix) : [],
      master: readTable(wb.Sheets['공연장마스터']),
      official: wb.Sheets['공식누적'] ? readTable(wb.Sheets['공식누적']) : [],
      list: wb.Sheets['연도별목록'] ? aoa(wb.Sheets['연도별목록']) : null,
      vars: wb.Sheets['변수설명'] ? aoa(wb.Sheets['변수설명']) : null,
      settings,
    };
  }

  /* ---------- ③ capacity 결정 + 순위 산정 ---------- */
  function plan(src, opts) {
    opts = opts || {};
    const s = src.settings || {};
    const refStr = opts.refDate || s['기준일'] || fmt(new Date());
    const refDate = serialToDate(isNum(refStr) ? refStr : null) || new Date(refStr + 'T00:00:00Z');
    const autoYears = String(opts.autoYears || s['자동순위연도'] || '').split(/[,\s]+/).map(Number).filter(Boolean);
    const topN = Number(opts.topN || s['상위N'] || 15);
    const notes = [];

    const mMap = new Map();
    // 출처(source)가 적힌 마스터 값만 사용 — 근거 없는 값이 빈 칸을 채우지 않도록
    src.master.forEach(m => { if (m.location && isNum(m.capacity) && String(m.source || '').trim()) mMap.set(norm(m.location) + '|' + norm(m.city), m); });
    const resolveCap = (r) => {
      const m = mMap.get(norm(r.location) + '|' + norm(r.city));
      if (r.capacity !== undefined && r.capacity !== null && r.capacity !== '') {
        if (m && isNum(r.capacity) && r.capacity !== m.capacity) notes.push({ kind: '마스터와 다른 capacity', year: r.year, artist: r.artist, detail: `${r.location}: 행 ${r.capacity.toLocaleString()} / 마스터 ${m.capacity.toLocaleString()} → 행 값 사용` });
        return { cap: r.capacity, from: '행' };
      }
      if (m) return { cap: m.capacity, from: '마스터' };
      return { cap: null, from: '' };
    };
    const all = src.main.map((r, i) => Object.assign({}, r, { __i: i }, (() => { const c = resolveCap(r); return { __cap: c.cap, __capFrom: c.from }; })()));

    // 공식누적
    const off = new Map();
    src.official.forEach(o => { if (o['집계묶음'] && isNum(o['공식관객수'])) off.set(String(o['집계묶음']).trim(), o); });

    // 연도별 순위
    const rankings = {};
    autoYears.forEach(y => {
      const per = new Map();
      all.filter(r => r.year === y).forEach(r => {
        if (!per.has(r.artist)) per.set(r.artist, { artist: r.artist, total: 0, rows: 0, groups: new Set(), usedOfficial: [], byOfficial: 0, byAtt: 0, byCap: 0, blank: 0 });
        const p = per.get(r.artist); p.rows++;
        const g = r['집계묶음'] ? String(r['집계묶음']).trim() : '';
        if (g && off.has(g)) { if (!p.groups.has(g)) { p.groups.add(g); p.total += off.get(g)['공식관객수']; p.byOfficial += off.get(g)['공식관객수']; p.usedOfficial.push(`${g} ${off.get(g)['공식관객수'].toLocaleString()}`); } return; }
        if (isNum(r.attendance)) { p.total += r.attendance; p.byAtt += r.attendance; }
        else if (isNum(r.__cap) && isNum(r.days)) { p.total += r.__cap * r.days; p.byCap += r.__cap * r.days; }
        else p.blank++;
      });
      const ranked = [...per.values()].sort((a, b) => b.total - a.total).map((p, i) => Object.assign(p, { rank: i + 1 }));
      rankings[y] = ranked;
    });

    // 기존 목록 순위
    const listRank = {};
    if (src.list) src.list.rows.forEach(row => { const y = row[0]; if (isNum(y)) listRank[y] = row.slice(1, 16).filter(Boolean).map(String); });

    // 출력 행 순서
    const out = [];
    const years = [...new Set(all.map(r => r.year))].sort((a, b) => a - b);
    const excluded = [];
    years.forEach(y => {
      const yr = all.filter(r => r.year === y);
      if (rankings[y]) {
        const top = rankings[y].slice(0, topN).map(p => p.artist);
        rankings[y].slice(topN).forEach(p => excluded.push(p));
        top.forEach(a => yr.filter(r => r.artist === a).sort((p, q) => ((p.start_date instanceof Date) ? p.start_date : 0) - ((q.start_date instanceof Date) ? q.start_date : 0)).forEach(r => out.push(r)));
      } else yr.forEach(r => out.push(r));
    });
    const appendix = src.appendix.map(r => Object.assign({}, r, (() => { const c = resolveCap(r); return { __cap: c.cap, __capFrom: c.from }; })()));

    // 예정 표시
    const tag = (r) => {
      const D = (d) => d instanceof Date ? d : null;
      const sd = D(r.start_date), ed = D(r.end_date) || sd;
      if (sd && sd > refDate) return '[예정]';
      if (sd && ed && sd <= refDate && ed > refDate) return '[일부 예정]';
      return '';
    };
    [...out, ...appendix].forEach(r => r.__tag = tag(r));

    // 자동 순위 연도: 15팀 안에 든 팀은 SUB에서 빼서 목록에 두 번 나오지 않게
    const subDrop = {};
    if (src.list) src.list.rows.forEach(row => {
      const y = Number(row[0]);
      if (!rankings[y]) return;
      const top = rankings[y].slice(0, topN).map(x => norm(x.artist));
      const drop = listSubs(row).filter(sname => top.includes(norm(sname)));
      if (drop.length) { subDrop[y] = drop.map(norm); drop.forEach(a => notes.push({ kind: 'SUB 정리', year: y, artist: a, detail: `자동 순위 ${top.indexOf(norm(a)) + 1}위로 상위 ${topN}팀 안에 들어 목록 SUB 칸에서 뺐어요` })); }
    });

    const capFromMaster = out.filter(r => r.__capFrom === '마스터').length;
    return { out, appendix, rankings, listRank, excluded, notes, refDate, autoYears, topN, capFromMaster, subDrop };
  }

  /* ---------- ④ 결과 엑셀 쓰기 ---------- */
  async function writeOutput(src, p) {
    const wb = new root.ExcelJS.Workbook();
    const thin = { style: 'thin', color: { argb: 'FF999999' } };
    const border = { top: thin, left: thin, bottom: thin, right: thin };
    const sub = wb.addWorksheet(SUB), lst = wb.addWorksheet(LIST), main = wb.addWorksheet(MAIN), vars = wb.addWorksheet(VARS);
    main.addRow(MAIN_HDR); sub.addRow(SUB_HDR);
    [main, sub].forEach(ws => { ws.getRow(1).font = { bold: true }; ws.views = [{ state: 'frozen', ySplit: 1 }]; });
    const rows = [...p.out, ...p.appendix];
    rows.forEach((r, i) => {
      const n = i + 2;
      const cap = r.__cap;
      const hasCap = cap !== null && cap !== undefined && cap !== '';
      const att = r.attendance;
      const stored = isNum(r['attendance/days']) ? r['attendance/days'] : null;
      // 원자료 값 우선 보존. 비어 있고 신규 범위(2025~)면 관객수÷일수 계산
      const apd = stored !== null ? stored : (isNum(att) && isNum(r.days) && r.days > 0 && r.year >= 2025 ? att / r.days : null);
      const ruleCD = !!r.attendance_rule && isNum(cap);   // 원자료처럼 attendance = 수용인원×일수 수식
      const cdRes = isNum(cap) && isNum(r.days) ? cap * r.days : undefined;
      const attM = ruleCD ? { formula: `$L${n}*$H${n}`, result: cdRes } : (att ?? null);
      const attS = ruleCD ? { formula: `$L${n}*$H${n}`, result: cdRes } : (att ?? null);
      const rem = r.__tag ? (r.__tag + ' ' + (r.remarks || '')).trim() : (r.remarks || null);
      const qtVal = r.start_date instanceof Date ? Math.floor(r.start_date.getUTCMonth() / 3) + 1 : null;
      const m = [i + 1, r.year, r.start_date instanceof Date ? { formula: `ROUNDUP(MONTH(F${n})/3,0)`, result: qtVal } : null, r.artist, r.type, r.start_date, r.end_date, r.days, r.country, r.city, r.location, hasCap ? cap : null,
        hasCap ? { formula: `$L${n}*$H${n}`, result: isNum(cap) && isNum(r.days) ? cap * r.days : undefined } : null, attM, apd, r.ticket_price ?? null, rem,
        r.gender ?? null, r.member ?? null, r.deyear ?? null, r.demon ?? null, r.deday ?? null, r.agency ?? null, r.agency2 ?? null, r.agency3 ?? null];
      // (2) 시트: 본 시트와 같은 열 구성. 원본처럼 qt는 값, capacity는 숫자만('-' 등 텍스트는 공란)
      const s = [i + 1, r.year, qtVal, r.artist, r.type, r.start_date, r.end_date, r.days, r.country, r.city, r.location, isNum(cap) ? cap : null,
        isNum(cap) ? { formula: `$L${n}*$H${n}`, result: isNum(cap) && isNum(r.days) ? cap * r.days : undefined } : null, attS, apd, r.ticket_price ?? null, rem,
        r.gender ?? null, r.member ?? null, r.deyear ?? null, r.demon ?? null, r.deday ?? null, r.agency ?? null, r.agency2 ?? null, r.agency3 ?? null];
      const mr = main.addRow(m), sr = sub.addRow(s);
      [mr, sr].forEach(row => { row.getCell(6).numFmt = row.getCell(7).numFmt = 'yyyy-mm-dd'; });
      const NUMF = '_-* #,##0_-;\\-* #,##0_-;_-* "-"_-;_-@_-';
      [mr, sr].forEach(row => [12, 13, 14, 15].forEach(c => row.getCell(c).numFmt = NUMF));
      if (r.attendance_label && isNum(att)) { const f = `#,##0" (${r.attendance_label})"`; mr.getCell(14).numFmt = f; sr.getCell(14).numFmt = f; }
      const fill = r.__tag === '[예정]' ? FILL_PLAN : r.__tag === '[일부 예정]' ? FILL_PART : null;
      [mr, sr].forEach(row => row.eachCell({ includeEmpty: true }, (c, col) => {
        if (col > 25) return;
        if (c.value !== null && c.value !== undefined && c.value !== '') c.border = border;
        if (fill) c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fill } };
      }));
    });
    [main, sub].forEach(ws => ws.getRow(1).eachCell(c => c.border = border));
    const widths = { artist: 16, location: 32, remarks: 50, city: 14, country: 12, start_date: 11, end_date: 11 };
    MAIN_HDR.forEach((h, i) => main.getColumn(i + 1).width = widths[h] || 9);
    SUB_HDR.forEach((h, i) => sub.getColumn(i + 1).width = widths[h] || 9);

    // 연도별 아티스트 목록
    if (src.list) {
      const a = src.list;
      a.rows.forEach((row, i) => {
        const y = row[0];
        let vals = row.slice();
        if (isNum(y) && p.rankings[y]) {
          const top = p.rankings[y].slice(0, p.topN).map(x => x.artist);
          for (let k = 0; k < 15; k++) vals[1 + k] = top[k] ?? null;
          const prevNote = String(row[16] || '').replace(/^\[자동 산정[^\]]*\]\s*/, '');
          vals[16] = `[자동 산정 ${fmt(p.refDate)}] ` + prevNote;
          const drop = (p.subDrop || {})[y] || [];
          if (drop.length) setSubs(vals, listSubs(vals).filter(sname => !drop.includes(norm(sname))));
        }
        vals.forEach((v, j) => { if (v !== null && v !== undefined) { const c = lst.getCell(a.origin.r + 1 + i, a.origin.c + 1 + j); c.value = v; c.border = border; } });
      });
      lst.getColumn(a.origin.c + 1).width = 8;
      for (let j = 1; j <= 15; j++) lst.getColumn(a.origin.c + 1 + j).width = 14;
      lst.getColumn(a.origin.c + 17).width = 60;
    }
    if (src.vars) src.vars.rows.forEach((row, i) => row.forEach((v, j) => { if (v !== null && v !== undefined) { const c = vars.getCell(src.vars.origin.r + 1 + i, src.vars.origin.c + 1 + j); c.value = v; c.border = border; } }));
    vars.getColumn(2).width = 18; vars.getColumn(3).width = 60;
    return wb.xlsx.writeBuffer();
  }


  /* ---------- ⑤ 업데이트 파일 (추가·수정·삭제) ---------- */
  const UPD_SHEET = '업데이트', LIST_SHEET = '목록';
  const LIST_COLS = ['작업', 'year', '순위', '찾을 문구', '값', '근거'];
  const UPD_COLS = ['작업', '대상_start_date', '대상_location', ...FIELDS, '근거', '대상시트'];
  const META = ['gender', 'member', 'deyear', 'demon', 'deday', 'agency', 'agency2', 'agency3'];
  const CLEAR = '(비움)';
  const dkey = (d) => d instanceof Date ? fmt(d) : String(d == null ? '' : d).trim();

  function readUpdate(wb) {
    if (![UPD_SHEET, LIST_SHEET, '공식누적', '공연장마스터'].some(n => wb.Sheets[n])) throw new Error(`'${UPD_SHEET}'·'${LIST_SHEET}'·'공식누적'·'공연장마스터' 시트가 하나도 없어요`);
    const toD = (v) => serialToDate(v) || (v ?? null);
    const withRow = (ws) => readTable(ws);   // 각 행에 엑셀 행 번호(__row)가 들어 있음
    const rows = wb.Sheets[UPD_SHEET] ? withRow(wb.Sheets[UPD_SHEET]).map(r => {
      ['start_date', 'end_date', '대상_start_date'].forEach(k => { if (r[k] !== undefined && r[k] !== CLEAR) r[k] = toD(r[k]); });
      return r;
    }) : [];
    const official = wb.Sheets['공식누적'] ? readTable(wb.Sheets['공식누적']) : [];
    const list = wb.Sheets[LIST_SHEET] ? withRow(wb.Sheets[LIST_SHEET]) : [];
    const master = wb.Sheets['공연장마스터'] ? withRow(wb.Sheets['공연장마스터']) : [];
    return { rows, official, list, master };
  }

  /* 연도별목록 행 구조: [연도, 1위…15위, 비고, SUB…] (SUB = 비고 오른쪽 칸, 2024 IU처럼) */
  const L_NOTE = 16, L_SUB = 17;
  const filled = (v) => v !== null && v !== undefined && String(v).trim() !== '';
  const listSubs = (row) => row.slice(L_SUB).filter(filled).map(v => String(v).trim());
  const setSubs = (row, subs) => { while (row.length < L_SUB) row.push(null); row.length = L_SUB; subs.forEach(s => row.push(s)); };
  const listTop = (row) => row.slice(1, 16).map(v => filled(v) ? String(v).trim() : '');

  function applyListOps(list, ops, autoYears, log) {
    const L = (u, 결과, 대상, 내용) => log.push({ 행: `목록 ${u.__row}`, 작업: u['작업'] || '', year: u.year ?? '', artist: filled(u['값']) && /SUB|순위/.test(u['작업'] || '') ? String(u['값']).trim() : '', 대상, 결과, 내용, 근거: u['근거'] || '' });
    ops.forEach(u => {
      const op = String(u['작업'] || '').replace(/\s/g, '');
      const year = Number(u.year);
      if (!op) return L(u, '오류', '', '작업 칸이 비어 있어요');
      if (!list) return L(u, '오류', '', "원본에 '연도별목록' 시트가 없어요");
      if (!year) return L(u, '오류', '', 'year가 비어 있어요');
      const row = list.rows.find(r => Number(r[0]) === year);
      if (!row) return L(u, '오류', '', `연도별목록에 ${year}년 행이 없어요`);
      const v = filled(u['값']) ? String(u['값']).trim() : '';
      const top = listTop(row), subs = listSubs(row);
      const at = (name) => top.findIndex(a => norm(a) === norm(name));
      if (op === 'SUB추가') {
        if (!v) return L(u, '오류', 'SUB', '값 칸에 아티스트 이름을 적어 주세요');
        if (at(v) >= 0) return L(u, '건너뜀', 'SUB', `${v}은(는) 이미 ${at(v) + 1}위에 있어요`);
        if (subs.some(s => norm(s) === norm(v))) return L(u, '건너뜀', 'SUB', `${v}은(는) 이미 SUB에 있어요`);
        subs.push(v); setSubs(row, subs);
        L(u, '목록', 'SUB', `SUB에 ${v} 추가 → SUB: ${subs.join(', ')}`);
      } else if (op === 'SUB삭제') {
        if (!v) return L(u, '오류', 'SUB', '값 칸에 아티스트 이름을 적어 주세요');
        const i = subs.findIndex(s => norm(s) === norm(v));
        if (i < 0) return L(u, '오류', 'SUB', `${year}년 SUB에 ${v}이(가) 없어요 (지금 SUB: ${subs.join(', ') || '없음'})`);
        subs.splice(i, 1); setSubs(row, subs);
        L(u, '목록', 'SUB', `SUB에서 ${v} 삭제 → SUB: ${subs.join(', ') || '없음'}`);
      } else if (op === '비고교체') {
        if (!v) return L(u, '오류', '비고', `값 칸이 비어 있어요 (비고를 지우려면 ${CLEAR})`);
        row[L_NOTE] = v === CLEAR ? null : v;
        L(u, '목록', '비고', v === CLEAR ? '비고를 비웠어요' : `비고를 새 글로 바꿨어요 (${v.length}자)`);
      } else if (op === '비고덧붙이기') {
        if (!v) return L(u, '오류', '비고', '값 칸에 덧붙일 글을 적어 주세요');
        row[L_NOTE] = [filled(row[L_NOTE]) ? String(row[L_NOTE]) : '', v.replace(/^\+\s*/, '')].filter(Boolean).join('\n');
        L(u, '목록', '비고', `비고 끝에 덧붙였어요: ${v.replace(/^\+\s*/, '')}`);
      } else if (op === '비고바꾸기') {
        const f = filled(u['찾을 문구']) ? String(u['찾을 문구']) : '';
        if (!f) return L(u, '오류', '비고', "'찾을 문구' 칸이 비어 있어요");
        const note = filled(row[L_NOTE]) ? String(row[L_NOTE]) : '';
        const n = note.split(f).length - 1;
        if (!n) return L(u, '오류', '비고', `비고에서 '${f}'을(를) 찾지 못했어요 (띄어쓰기·기호까지 같아야 해요)`);
        const to = v === CLEAR ? '' : v;
        row[L_NOTE] = note.split(f).join(to);
        L(u, '목록', '비고', `'${f}' → '${to || '(삭제)'}' ${n}곳`);
      } else if (op === '순위교체' || op === '순위넣기') {
        const k = Number(u['순위']);
        if (autoYears.includes(year)) return L(u, '오류', `${k || ''}위`, `${year}년은 자동 순위 연도라 순위가 공연 행으로 다시 계산돼요 — 원본 공연 행을 고쳐 주세요`);
        if (!(k >= 1 && k <= 15 && Number.isInteger(k))) return L(u, '오류', '순위', '순위 칸에 1~15를 적어 주세요');
        if (!v) return L(u, '오류', `${k}위`, '값 칸에 아티스트 이름을 적어 주세요');
        if (at(v) >= 0) return L(u, '오류', `${k}위`, `${v}은(는) 이미 ${at(v) + 1}위에 있어요`);
        const restSubs = subs.filter(s => norm(s) !== norm(v));
        const wasSub = restSubs.length !== subs.length;
        let out;
        if (op === '순위교체') {
          out = top[k - 1];
          top[k - 1] = v;
        } else {
          const nt = top.slice(0, k - 1).concat([v], top.slice(k - 1));
          out = nt[15]; nt.length = 15;
          top.splice(0, 15, ...nt);
        }
        if (out && !restSubs.some(s => norm(s) === norm(out))) restSubs.push(out);
        top.forEach((a, i) => row[1 + i] = a || null);
        setSubs(row, restSubs);
        L(u, '목록', `${k}위`, `${op === '순위교체' ? `${k}위 ${out || '(빈칸)'} → ${v}` : `${k}위에 ${v} 넣음, 아래 팀 한 칸씩 내려감`}${wasSub ? ' (SUB에서 올라옴)' : ''}${out ? ` / ${out}은(는) SUB로` : ''}`);
      } else L(u, '오류', '', `작업은 SUB추가·SUB삭제·비고교체·비고덧붙이기·비고바꾸기·순위교체·순위넣기 중 하나여야 해요 (지금: '${u['작업']}')`);
    });
  }

  /* 공연장마스터: 작업 추가/수정/삭제, location+city로 찾음 */
  const MASTER_COLS = ['작업', 'location', 'city', 'country', 'capacity', 'source', '근거'];
  function applyMasterOps(master, ops, log) {
    const L = (u, 결과, 내용) => log.push({ 행: `마스터 ${u.__row}`, 작업: u['작업'] || '', year: '', artist: '', 대상: `${u.location || ''}${u.city ? ' / ' + u.city : ''}`, 결과, 내용, 근거: u['근거'] || '' });
    const key = (x) => norm(x.location) + '|' + norm(x.city);
    ops.forEach(u => {
      const op = String(u['작업'] || '').trim();
      if (!u.location) return L(u, '오류', 'location이 비어 있어요');
      const i = master.findIndex(m => key(m) === key(u));
      const capWarn = (m) => isNum(m.capacity) && !String(m.source || '').trim() ? ' (source가 비어 있어 빈 capacity 채우기에는 쓰이지 않아요)' : '';
      if (u.capacity !== undefined && u.capacity !== CLEAR && !isNum(u.capacity)) return L(u, '오류', `capacity는 숫자여야 해요 (지금: '${u.capacity}')`);
      if (op === '추가') {
        if (i >= 0) return L(u, '건너뜀', '같은 공연장(location·city)이 이미 있어요 — 고치려면 수정으로 적어 주세요');
        const m = { location: u.location, city: u.city ?? null, country: u.country ?? null, capacity: isNum(u.capacity) ? u.capacity : null, source: u.source && u.source !== CLEAR ? u.source : null };
        master.push(m);
        L(u, '추가', `capacity ${isNum(m.capacity) ? m.capacity.toLocaleString() : '(빈칸)'}` + capWarn(m));
      } else if (op === '수정') {
        if (i < 0) return L(u, '오류', '고칠 공연장을 찾지 못했어요 (location·city를 마스터와 똑같이)');
        const m = master[i], ch = [];
        ['country', 'capacity', 'source'].forEach(f => {
          if (u[f] === undefined || u[f] === null || u[f] === '') return;
          const nv = u[f] === CLEAR ? null : u[f];
          if (dkey(m[f]) !== dkey(nv)) { ch.push(`${f}: ${dkey(m[f]) || '(빈칸)'} → ${dkey(nv) || '(빈칸)'}`); m[f] = nv; }
        });
        L(u, ch.length ? '수정' : '변경 없음', (ch.join(' / ') || '바뀐 값이 없어요') + capWarn(m));
      } else if (op === '삭제') {
        if (i < 0) return L(u, '오류', '지울 공연장을 찾지 못했어요');
        master.splice(i, 1);
        L(u, '삭제', '마스터에서 뺐어요 — 이 공연장 행의 빈 capacity는 더 이상 채워지지 않아요');
      } else L(u, '오류', `작업은 추가·수정·삭제 중 하나여야 해요 (지금: '${op}')`);
    });
  }

  function applyUpdate(src, upd, opts) {
    opts = opts || {};
    const autoYears = String(opts.autoYears || (src.settings || {})['자동순위연도'] || '').split(/[,\s]+/).map(Number).filter(Boolean);
    const out = Object.assign({}, src, {
      main: src.main.map(r => Object.assign({}, r)), official: src.official.map(o => Object.assign({}, o)),
      appendix: (src.appendix || []).map(r => Object.assign({}, r)), master: (src.master || []).map(m => Object.assign({}, m)),
      list: src.list ? { origin: src.list.origin, rows: src.list.rows.map(r => r.slice()) } : null,
    });
    const log = [];
    const isSolo = (u) => String(u['대상시트'] || '').replace(/\s/g, '') === '솔로부록';
    const L = (u, 결과, 내용) => log.push({ 행: u.__row, 작업: u['작업'] || '', year: u.year ?? '', artist: u.artist ?? '', 대상: `${isSolo(u) ? '[솔로부록] ' : ''}${dkey(u['대상_start_date'] || u.start_date)} ${u['대상_location'] || u.location || ''}`.trim(), 결과, 내용, 근거: u['근거'] || '' });
    // 대상시트: 비우면 공연목록, '솔로부록'이면 솔로부록
    const arrOf = (u) => isSolo(u) ? out.appendix : out.main;
    const find = (u) => {
      const d = dkey(u['대상_start_date'] || u.start_date), loc = norm(u['대상_location'] || u.location);
      return arrOf(u).findIndex(r => r.year === u.year && norm(r.artist) === norm(u.artist) && dkey(r.start_date) === d && (!loc || norm(r.location) === loc));
    };
    const addSrc = (r, u) => { if (u['출처']) r['출처'] = [r['출처'], u['출처']].filter(Boolean).join(' / '); };
    upd.rows.forEach(u => {
      const op = String(u['작업'] || '').trim();
      const ts = String(u['대상시트'] || '').replace(/\s/g, '');
      if (ts && ts !== '솔로부록' && ts !== '공연목록') return L(u, '오류', `대상시트는 비우거나(공연목록) '솔로부록'이어야 해요 (지금: '${u['대상시트']}')`);
      if (!isNum(u.year) || !u.artist) return L(u, '오류', 'year·artist가 비어 있어요');
      const arr = arrOf(u);
      if (op === '추가') {
        if (!(u.start_date instanceof Date) || !u.location) return L(u, '오류', 'start_date·location이 필요해요');
        if (find(u) >= 0) return L(u, '건너뜀', '같은 공연(연도·아티스트·시작일·공연장)이 이미 있어요');
        const r = {}; FIELDS.forEach(f => r[f] = (u[f] === undefined || u[f] === CLEAR) ? null : u[f]);
        const same = arr.find(x => norm(x.artist) === norm(u.artist)) || [...out.main, ...out.appendix].find(x => norm(x.artist) === norm(u.artist));
        const filled = [];
        if (same) META.forEach(f => { if (r[f] == null && same[f] != null) { r[f] = same[f]; filled.push(f); } });
        if (!r.type) r.type = 'concert';
        if (isSolo(u)) {
          // 솔로부록은 아티스트 블록 순서대로 출력되므로 같은 아티스트 블록 안 날짜 자리에 끼움
          const same = arr.map((x, i) => [x, i]).filter(([x]) => norm(x.artist) === norm(r.artist));
          if (!same.length) arr.push(r);
          else {
            const later = same.find(([x]) => x.start_date instanceof Date && x.start_date > r.start_date);
            arr.splice(later ? later[1] : same[same.length - 1][1] + 1, 0, r);
          }
        } else arr.push(r);
        L(u, '추가', `${r.city || ''} ${r.location} ${dkey(r.start_date)}~${dkey(r.end_date)} ${r.days ?? ''}일` + (filled.length ? ` (아티스트 정보 ${filled.length}칸은 기존 행에서 복사)` : ''));
      } else if (op === '수정') {
        const i = find(u);
        if (i < 0) return L(u, '오류', '고칠 공연을 찾지 못했어요 (대상_start_date·대상_location 확인)');
        const r = arr[i], ch = [];
        FIELDS.forEach(f => {
          if (f === '출처' || u[f] === undefined || u[f] === null || u[f] === '') return;
          let nv = u[f] === CLEAR ? null : u[f];
          if (typeof nv === 'string' && /^\+\s*/.test(nv)) nv = [r[f], nv.replace(/^\+\s*/, '')].filter(Boolean).join(' / ');   // '+ 내용' = 기존 값 뒤에 덧붙이기
          if (dkey(r[f]) !== dkey(nv)) { ch.push(`${f}: ${dkey(r[f]) || '(빈칸)'} → ${dkey(nv) || '(빈칸)'}`); r[f] = nv; }
        });
        addSrc(r, u);
        L(u, ch.length ? '수정' : '변경 없음', ch.join(' / ') || '바뀐 값이 없어요');
      } else if (op === '삭제') {
        const i = find(u);
        if (i < 0) return L(u, '오류', '지울 공연을 찾지 못했어요');
        const r = arr.splice(i, 1)[0];
        L(u, '삭제', `${r.city || ''} ${r.location} ${dkey(r.start_date)}`);
      } else L(u, '오류', `작업은 추가·수정·삭제 중 하나여야 해요 (지금: '${op}')`);
    });
    upd.official.forEach((o, k) => {
      const g = String(o['집계묶음'] || '').trim();
      if (!g || !isNum(o['공식관객수'])) return log.push({ 행: `공식누적 ${k + 2}`, 작업: '공식누적', 결과: '오류', 내용: '집계묶음·공식관객수가 필요해요' });
      const i = out.official.findIndex(x => String(x['집계묶음']).trim() === g);
      const v = { 집계묶음: g, year: o.year, artist: o.artist, 공식관객수: o['공식관객수'], 출처: o['출처'], 설명: o['설명'] };
      if (i >= 0) { const old = out.official[i]['공식관객수']; out.official[i] = v; log.push({ 행: `공식누적 ${k + 2}`, 작업: '공식누적', year: o.year, artist: o.artist, 대상: g, 결과: '수정', 내용: `${old} → ${v.공식관객수}`, 근거: o['출처'] || '' }); }
      else { out.official.push(v); log.push({ 행: `공식누적 ${k + 2}`, 작업: '공식누적', year: o.year, artist: o.artist, 대상: g, 결과: '추가', 내용: `${v.공식관객수.toLocaleString()}`, 근거: o['출처'] || '' }); }
    });
    applyListOps(out.list, upd.list || [], autoYears, log);
    applyMasterOps(out.master, upd.master || [], log);
    return { src: out, log };
  }

  async function writeUpdate(rows, official, list) {
    const wb = new root.ExcelJS.Workbook();
    const ws = wb.addWorksheet(UPD_SHEET);
    ws.addRow(UPD_COLS); ws.getRow(1).font = { bold: true };
    (rows || []).forEach(d => ws.addRow(UPD_COLS.map(c => d[c] === undefined ? null : d[c])));
    UPD_COLS.forEach((c, i) => {
      ws.getColumn(i + 1).width = ({ 작업: 7, 대상시트: 10, artist: 14, location: 28, 대상_location: 28, remarks: 40, 출처: 40, 근거: 50 })[c] || 12;
      if (/start_date|end_date/.test(c)) ws.getColumn(i + 1).numFmt = 'yyyy-mm-dd';
    });
    ws.views = [{ state: 'frozen', ySplit: 1, xSplit: 3 }];
    const o = wb.addWorksheet('공식누적');
    const oc = ['집계묶음', 'year', 'artist', '공식관객수', '출처', '설명'];
    o.addRow(oc); o.getRow(1).font = { bold: true };
    (official || []).forEach(d => o.addRow(oc.map(c => d[c] ?? null)));
    const ls = wb.addWorksheet(LIST_SHEET);
    ls.addRow(LIST_COLS); ls.getRow(1).font = { bold: true };
    (list || []).forEach(d => ls.addRow(LIST_COLS.map(c => d[c] ?? null)));
    LIST_COLS.forEach((c, i) => ls.getColumn(i + 1).width = ({ 작업: 12, year: 7, 순위: 6, '찾을 문구': 40, 값: 40, 근거: 50 })[c] || 12);
    ls.views = [{ state: 'frozen', ySplit: 1 }];
    const ms = wb.addWorksheet('공연장마스터');
    ms.addRow(MASTER_COLS); ms.getRow(1).font = { bold: true };
    MASTER_COLS.forEach((c, i) => ms.getColumn(i + 1).width = ({ 작업: 7, location: 36, city: 14, source: 40, 근거: 50 })[c] || 12);
    ms.views = [{ state: 'frozen', ySplit: 1 }];
    const h = wb.addWorksheet('작성법');
    [['작업', '추가 / 수정 / 삭제'],
     ['추가', '새 공연 1행. year·artist·start_date·location 필수. gender~agency3가 비면 같은 아티스트 기존 행에서 복사'],
     ['수정', '대상_start_date·대상_location으로 기존 공연을 찾고, 값을 적은 칸만 바꿈. 칸을 지우려면 (비움), 기존 글 뒤에 덧붙이려면 + 로 시작'],
     ['삭제', '대상_start_date·대상_location(또는 start_date·location)으로 찾은 공연을 삭제'],
     ['출처', '수정 시 기존 출처 뒤에 덧붙임'],
     ['공식누적 시트', '합계만 공식 발표된 경우. 같은 집계묶음이 있으면 덮어씀. 공연목록 행의 집계묶음 칸도 같이 적어야 반영됨'],
     [],
     ['목록 시트', "'연도별 아티스트 목록' 고치기. 한 행 = 작업 하나, 위에서부터 차례로 적용. year는 필수"],
     ['SUB추가', '값 = 아티스트 이름. 비고 오른쪽 SUB 칸(S열부터) 끝에 넣음. 이미 15팀이나 SUB에 있으면 건너뜀'],
     ['SUB삭제', '값 = 아티스트 이름. SUB 칸에서 빼고 뒤 칸을 앞으로 당김'],
     ['비고교체', '값 = 새 비고 전체. 비우려면 (비움)'],
     ['비고덧붙이기', '값 = 덧붙일 글. 기존 비고 다음 줄에 붙음'],
     ['비고바꾸기', "찾을 문구 = 지금 비고 속 글(띄어쓰기·기호까지 똑같이), 값 = 바꿀 글. 지우려면 값에 (비움). 예: 찾을 문구 'XG·&TEAM·KATSEYE 포함' → 값 '&TEAM·KATSEYE 포함'"],
     ['순위교체', '순위(1~15)의 팀을 값의 팀으로 바꿈. 밀려난 팀은 SUB로. 자동 순위 연도(설정의 자동순위연도)는 순위를 공연 행으로 다시 계산하므로 쓸 수 없음'],
     ['순위넣기', '순위(1~15) 자리에 값의 팀을 넣고 아래 팀을 한 칸씩 내림. 15위였던 팀은 SUB로. 넣는 팀이 SUB에 있었으면 SUB에서 빠짐'],
     ['자동 순위 연도', 'SUB추가·비고 작업은 가능. SUB 팀이 자동 순위로 15팀 안에 들면 결과 엑셀에서 SUB 칸에서 자동으로 빠짐'],
     [],
     ['대상시트 열', "'업데이트' 시트 맨 오른쪽. 비우면 공연목록, '솔로부록'이라고 적으면 솔로부록 행을 추가·수정·삭제. 솔로부록 추가 행은 같은 아티스트 블록 안 날짜 자리에 들어감"],
     ['공연장마스터 시트', '작업(추가/수정/삭제) + location·city로 공연장을 찾음. 수정은 적은 칸(country·capacity·source)만 바꾸고 (비움)은 칸 비우기'],
     ['', 'source(출처)를 적은 공연장만 공연 행의 빈 capacity를 채우는 데 쓰임. 배치별로 수용인원이 달라 공란으로 둘 공연장은 capacity를 (비움)으로']].forEach(r => h.addRow(r));
    h.getColumn(1).width = 14; h.getColumn(2).width = 100;
    return wb.xlsx.writeBuffer();
  }

  const api = { MASTER_COLS, convertLegacy, writeSource, readSource, plan, writeOutput, readUpdate, applyUpdate, writeUpdate, listSubs, FIELDS };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.KpopAssembler = api;
})(typeof window !== 'undefined' ? window : globalThis);
