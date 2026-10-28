/* K-pop 해외공연 데이터셋 검증 로직 (브라우저/Node 공용)
 * 입력: SheetJS workbook (cellDates:true, cellFormula:true 로 읽은 것)
 * 출력: { issues:[...], stats:{...}, venues:[...] }
 */
(function (root) {
  const MAIN = '해외 공연 현황';
  const SUB = '해외 공연 현황 (2)';
  const LIST = '연도별 아티스트 목록';
  const KOREA_COUNTRY = ['korea', 'south korea', 'republic of korea', '한국', '대한민국'];
  const KOREA_CITY = ['seoul', 'incheon', 'busan', 'goyang', 'daegu', 'gwangju', 'daejeon', 'ulsan', 'suwon', '서울', '인천', '부산', '고양'];

  const SEV = { ERROR: '오류', WARN: '확인필요', INFO: '참고' };

  const norm = (s) => String(s == null ? '' : s).toLowerCase().replace(/[-–—_.·]/g, ' ').replace(/\bthe\b/g, '').replace(/\s+/g, ' ').trim();
  const isNum = (v) => typeof v === 'number' && isFinite(v);
  // 엑셀 날짜(일련번호)를 시간대 영향 없이 UTC 자정 날짜로 변환
  const toDate = (v) => {
    if (typeof v === 'number' && isFinite(v) && v > 1000) {
      const p = root.XLSX.SSF.parse_date_code(v);
      return p ? new Date(Date.UTC(p.y, p.m - 1, p.d)) : null;
    }
    if (v instanceof Date && !isNaN(v)) return new Date(Date.UTC(v.getFullYear(), v.getMonth(), v.getDate()));
    return null;
  };
  const fmtDate = (d) => d ? `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}` : '';
  const normF = (f) => String(f || '').replace(/^=/, '').replace(/\$/g, '').replace(/\s/g, '').toUpperCase();

  function sheetRows(ws) {
    // 반환: [{r: 엑셀행번호, cells: {header: {v, f}}}]
    const XLSX = root.XLSX;
    const range = XLSX.utils.decode_range(ws['!ref']);
    const headers = [];
    for (let c = range.s.c; c <= range.e.c; c++) {
      const cell = ws[XLSX.utils.encode_cell({ r: range.s.r, c })];
      headers.push({ name: cell ? String(cell.v).trim() : '', col: c, letter: XLSX.utils.encode_col(c) });
    }
    const rows = [];
    for (let r = range.s.r + 1; r <= range.e.r; r++) {
      const cells = {};
      let any = false;
      for (const h of headers) {
        if (!h.name) continue;
        const cell = ws[XLSX.utils.encode_cell({ r, c: h.col })];
        if (cell) { const v = cell.t === 'z' ? undefined : cell.v; cells[h.name] = { v, f: cell.f, t: cell.t, z: cell.z }; if (v !== undefined && v !== '') any = true; }
        else cells[h.name] = { v: undefined };
      }
      if (any) rows.push({ r: r + 1, cells });
    }
    const col = {};
    headers.forEach(h => { if (h.name) col[h.name] = h.letter; });
    return { rows, col };
  }

  const val = (row, name) => row.cells[name] ? row.cells[name].v : undefined;
  // remarks에서 공연 회차 수 읽기 (없으면 days). 예: "총 3회", "3일 6회", "1일 2회"
  function showCount(rem, days) {
    const s = String(rem || '');
    let m = /총\s*(\d+)\s*회(?!차)/.exec(s);
    if (m) { const n = Number(m[1]); return n >= days ? n : days; }
    m = /(\d+)\s*일\s*(\d+)\s*회(?!차)/.exec(s);
    if (m) {
      const d = Number(m[1]), n = Number(m[2]);
      if (d === days && n >= days) return n;          // "3일 6회"
      if (d === 1 && n > 1) return n * days;          // "1일 2회" = 매일 n회
    }
    return days;
  }

  function run(wb, opts) {
    opts = opts || {};
    const rd = opts.refDate || new Date();
    const refDate = new Date(Date.UTC(rd.getFullYear(), rd.getMonth(), rd.getDate()));
    const strictFrom = opts.strictFrom || 2025;
    const venueMaster = opts.venueMaster || null; // Map key "loc|city" -> {capacity, source}
    const issues = [];
    const add = (sev, check, sheet, row, detail, extra) => issues.push(Object.assign({ sev, check, sheet, excelRow: row ? row.r : '', no: row ? val(row, 'no') : '', year: row ? val(row, 'year') : '', artist: row ? val(row, 'artist') : '', detail }, extra || {}));

    if (!wb.Sheets[MAIN]) {
      add(SEV.ERROR, '시트 구조', MAIN, null, `'${MAIN}' 시트를 찾을 수 없어요. 시트 이름을 확인해 주세요.`);
      return { issues, stats: {}, venues: [] };
    }
    const main = sheetRows(wb.Sheets[MAIN]);
    const sub = wb.Sheets[SUB] ? sheetRows(wb.Sheets[SUB]) : null;
    if (!sub) add(SEV.WARN, '시트 구조', SUB, null, `'${SUB}' 시트가 없어 두 시트 대조를 건너뛰었어요.`);
    const R = main.rows, C = main.col;
    const need = ['no', 'year', 'artist', 'start_date', 'end_date', 'days', 'country', 'city', 'location', 'capacity', 'capacity*days', 'attendance', 'remarks'];
    need.forEach(n => { if (!C[n]) add(SEV.ERROR, '시트 구조', MAIN, null, `'${n}' 열을 찾을 수 없어요.`); });

    // ---------- 섹션 구분 (본 데이터 / 솔로 부록) ----------
    let appendixStart = R.length;
    for (let i = 1; i < R.length; i++) {
      const y0 = val(R[i - 1], 'year'), y1 = val(R[i], 'year');
      if (isNum(y0) && isNum(y1) && y1 < y0) { appendixStart = i; break; }
    }
    if (appendixStart < R.length) add(SEV.INFO, '구조 안내', MAIN, R[appendixStart], `엑셀 ${R[appendixStart].r}행부터 연도가 다시 작아져서 '솔로 부록' 구간으로 보고 따로 검사했어요.`);

    // ---------- 행 단위 검사 ----------
    const venueMap = new Map();
    const dupKey = new Map();
    R.forEach((row, i) => {
      const year = val(row, 'year'), country = val(row, 'country'), city = val(row, 'city'), loc = val(row, 'location');
      const days = val(row, 'days'), cap = val(row, 'capacity'), att = val(row, 'attendance');
      const sd = toDate(val(row, 'start_date')), ed = toDate(val(row, 'end_date'));
      const rem = String(val(row, 'remarks') || '');

      // no 연번
      if (val(row, 'no') !== row.r - 1) add(SEV.ERROR, 'no 연번', MAIN, row, `no가 ${val(row, 'no')}인데 ${row.r - 1}이어야 해요.`);

      // 국내 공연
      if (KOREA_COUNTRY.includes(norm(country)) || (KOREA_CITY.includes(norm(city)) && !country)) add(SEV.ERROR, '국내 공연 포함', MAIN, row, `country='${country}', city='${city}' — 해외 공연 데이터셋에 들어갈 수 없어요.`);

      // 날짜·일수
      if (sd && ed && ed < sd) add(SEV.ERROR, '날짜', MAIN, row, `end_date(${fmtDate(ed)})가 start_date(${fmtDate(sd)})보다 빨라요.`);
      if (sd && ed && isNum(days)) {
        const span = Math.round((ed - sd) / 86400000) + 1;
        if (days > span) add(SEV.ERROR, '날짜', MAIN, row, `days=${days}인데 기간은 ${span}일뿐이에요.`);
      }
      if (isNum(year) && sd && sd.getUTCFullYear() !== year) add(SEV.ERROR, '날짜', MAIN, row, `year=${year}인데 start_date는 ${fmtDate(sd)}예요.`);

      // capacity / attendance
      if (cap !== undefined && cap !== null && cap !== '' && !isNum(cap)) add(SEV.WARN, 'capacity 형식', MAIN, row, `capacity가 숫자가 아니에요('${cap}'). 수식이 오류(#VALUE)가 나요.`);
      if (isNum(cap) && isNum(att) && isNum(days) && days > 0) {
        const perDay = att / days;
        // 하루 여러 회 공연: remarks의 "총 n회" / "n일 m회" / "1일 2회"로 회차 수를 읽어 회당 관객으로 판단
        const shows = showCount(rem, days);
        const perShow = att / shows;
        if (perShow > cap * 1.05) add(SEV.ERROR, '관객수 > 수용인원', MAIN, row, shows > days
          ? `총 ${shows}회 기준 회당 관객 ${Math.round(perShow).toLocaleString()}명이 capacity ${cap.toLocaleString()}보다 커요 — capacity가 틀렸을 가능성이 높아요.`
          : `하루 관객 ${Math.round(perDay).toLocaleString()}명이 capacity ${cap.toLocaleString()}보다 커요 — capacity가 틀렸을 가능성이 높아요. 하루 여러 회 공연이면 remarks에 "총 n회"를 적어 주세요.`);
        else if (perDay > cap * 1.05) add(SEV.INFO, '하루 여러 회 공연', MAIN, row, `하루 평균 관객 ${Math.round(perDay).toLocaleString()}명이 capacity보다 많지만 remarks의 총 ${shows}회 기준 회당 ${Math.round(perShow).toLocaleString()}명이라 수용인원 안이에요.`);
        if (days > 1 && cap >= att * 0.9 && perDay < cap * 0.6) add(SEV.WARN, 'capacity에 합계 의심', MAIN, row, `capacity ${cap.toLocaleString()}가 ${days}일 관객 합계(${att.toLocaleString()})와 비슷해요 — 1일 수용인원이 아니라 여러 날 합계를 넣었거나, 공연장 규모 대비 관객이 크게 적어요. 출처를 확인하세요.`);
        if (att === cap * days) {
          // attendance 셀 서식에 "(소속사 제공)" 같은 공식 출처 표시가 있으면 공식 전석 매진 발표로 보고 참고로 낮춤
          const lab = /"\s*\(([^)]+)\)\s*"/.exec((row.cells['attendance'] || {}).z || '');
          if (lab) add(SEV.INFO, '관객수=수용인원×일수', MAIN, row, `attendance(${att.toLocaleString()})가 capacity×days와 같지만 '${lab[1]}' 공식 수치라 참고로만 표시해요.`);
          else add(isNum(year) && year >= strictFrom ? SEV.WARN : SEV.INFO, '관객수=수용인원×일수', MAIN, row, `attendance(${att.toLocaleString()})가 capacity×days와 정확히 같아요 — 수용인원을 복사한 건 아닌지 출처를 확인하세요(공식 '전석 매진 ○만 명' 발표라면 괜찮음).`);
        }
      }
      // 수식 (본 시트)
      if (C['qt'] && C['start_date']) {
        const q = row.cells['qt'] || {};
        if (q.f && normF(q.f) !== `ROUNDUP(MONTH(${C['start_date']}${row.r})/3,0)`) add(SEV.ERROR, '수식 깨짐', MAIN, row, `qt 수식이 다른 행을 참조해요: =${q.f}`);
      }
      if (C['capacity*days']) {
        const f = (row.cells['capacity*days'] || {}).f;
        const expect = `${C['capacity']}${row.r}*${C['days']}${row.r}`;
        if (f && normF(f) !== expect) add(SEV.ERROR, '수식 깨짐', MAIN, row, `capacity*days 수식이 다른 행을 참조해요: =${f}`);
        if (!f && isNum(cap)) add(SEV.WARN, '수식 누락', MAIN, row, `capacity는 있는데 capacity*days 수식이 비어 있어요.`);
      }

      // [예정] 표시
      if (sd && sd > refDate && !rem.startsWith('[예정]') && !rem.startsWith('[일부 예정')) add(SEV.WARN, '[예정] 표시 누락', MAIN, row, `${fmtDate(sd)} 공연인데 기준일(${fmtDate(refDate)}) 이후라 remarks 앞에 [예정] 표시가 필요해요.`);
      if (sd && ed && sd <= refDate && ed > refDate && !rem.startsWith('[일부 예정')) add(SEV.WARN, '[예정] 표시 누락', MAIN, row, `기준일을 걸치는 공연이에요(${fmtDate(sd)}~${fmtDate(ed)}) — [일부 예정] 표시가 필요해요.`);
      if (rem.startsWith('[예정]') && ed && ed <= refDate) add(SEV.INFO, '개최 완료 확인', MAIN, row, `[예정]으로 표시돼 있지만 ${fmtDate(ed)}에 이미 끝났어요 — 표시·색을 지우고 공식 관객수를 확인하세요.`);

      // 중복
      const dk = [norm(val(row, 'artist')), fmtDate(sd), norm(loc)].join('|');
      if (sd && dupKey.has(dk)) add(SEV.ERROR, '중복 행', MAIN, row, `엑셀 ${dupKey.get(dk)}행과 아티스트·날짜·공연장이 같아요.`);
      else if (sd) dupKey.set(dk, row.r);

      // 공연장 수집
      if (loc) {
        const k = norm(loc) + '|' + norm(city);
        if (!venueMap.has(k)) venueMap.set(k, { location: loc, city, country, values: new Map(), rows: [] });
        const v = venueMap.get(k);
        v.rows.push(row);
        if (isNum(cap)) { if (!v.values.has(cap)) v.values.set(cap, []); v.values.get(cap).push(row); }
      }

      // 공연장 마스터 대조
      if (venueMaster && loc) {
        const m = venueMaster.get(norm(loc) + '|' + norm(city));
        if (m && isNum(m.capacity)) {
          if (isNum(cap) && cap !== m.capacity) add(SEV.WARN, '공연장 마스터와 불일치', MAIN, row, `capacity ${cap.toLocaleString()} ≠ 마스터 ${m.capacity.toLocaleString()} (${m.source || '출처 미기재'})`, { suggest: m.capacity });
          if (!isNum(cap)) add(SEV.INFO, '마스터로 채울 수 있는 칸', MAIN, row, `capacity가 비어 있어요 — 마스터 값 ${m.capacity.toLocaleString()} (${m.source || '출처 미기재'})`, { suggest: m.capacity });
        }
      }
    });

    // ---------- 같은 공연장 capacity 불일치 ----------
    const venues = [];
    venueMap.forEach((v) => {
      const vals = [...v.values.keys()].sort((a, b) => a - b);
      venues.push({ location: v.location, city: v.city, country: v.country, values: vals, count: v.rows.length, withCap: [...v.values.values()].reduce((s, a) => s + a.length, 0) });
      if (vals.length > 1) {
        const parts = vals.map(c => `${c.toLocaleString()}(${v.values.get(c).map(r => `${val(r, 'artist')} ${val(r, 'year')}`).slice(0, 3).join(', ')}${v.values.get(c).length > 3 ? ' 외' : ''})`);
        add(SEV.WARN, '같은 공연장 capacity 불일치', MAIN, v.values.get(vals[0])[0], `${v.location} / ${v.city}: ${parts.join(' vs ')}`, { artist: '(여러 아티스트)', no: '', year: '' });
      }
    });

    // ---------- 블록 정렬 ----------
    const checkBlocks = (rows, keyFn, label) => {
      const seen = new Map(); let prev = null;
      rows.forEach(row => {
        const k = keyFn(row);
        if (k !== prev) {
          if (seen.has(k)) add(SEV.ERROR, '그룹 묶음 깨짐', MAIN, row, `${label} '${val(row, 'artist')}'(${val(row, 'year')}) 블록이 엑셀 ${seen.get(k)}행에 이미 있었는데 다시 나타나요.`);
          else seen.set(k, row.r);
          prev = k;
        }
      });
    };
    const mainRows = R.slice(0, appendixStart), appRows = R.slice(appendixStart);
    checkBlocks(mainRows, r => val(r, 'year') + '|' + norm(val(r, 'artist')), '본 데이터');
    checkBlocks(appRows, r => norm(val(r, 'artist')), '솔로 부록');
    for (let i = 1; i < mainRows.length; i++) {
      const a = mainRows[i - 1], b = mainRows[i];
      if (val(a, 'year') === val(b, 'year') && norm(val(a, 'artist')) === norm(val(b, 'artist'))) {
        const da = toDate(val(a, 'start_date')), db = toDate(val(b, 'start_date'));
        if (da && db && db < da) add(SEV.WARN, '날짜순 정렬', MAIN, b, `같은 아티스트 블록 안에서 바로 위 행(${fmtDate(da)})보다 날짜가 빨라요(${fmtDate(db)}).`);
      }
    }

    // ---------- 연도별 아티스트 목록과 순서 대조 ----------
    if (wb.Sheets[LIST]) {
      const XLSX = root.XLSX;
      const ws = wb.Sheets[LIST];
      const rg = XLSX.utils.decode_range(ws['!ref']);
      for (let r = rg.s.r; r <= rg.e.r; r++) {
        let yc = null, year = null;
        for (let c = rg.s.c; c <= Math.min(rg.e.c, 3); c++) {
          const cell = ws[XLSX.utils.encode_cell({ r, c })];
          if (cell && isNum(cell.v) && cell.v > 2000 && cell.v < 2100) { yc = c; year = cell.v; break; }
        }
        if (year === null) continue;
        const listed = [];
        for (let c = yc + 1; c <= yc + 15; c++) {
          const cell = ws[XLSX.utils.encode_cell({ r, c })];
          if (cell && cell.v) listed.push(String(cell.v).trim());
        }
        const blocks = []; let prev = null;
        mainRows.filter(x => val(x, 'year') === year).forEach(x => { const a = String(val(x, 'artist')); if (a !== prev) { blocks.push({ a, row: x }); prev = a; } });
        if (!blocks.length) continue;
        const ln = listed.map(norm);
        const extra = blocks.filter(b => !ln.includes(norm(b.a)));
        if (year >= 2025) extra.forEach(b => add(SEV.WARN, '목록 밖 아티스트', MAIN, b.row, `${year}년 '연도별 아티스트 목록'(15팀)에 없는 '${b.a}' 상세 행이 있어요 — 16위 이하는 상세 행을 넣지 않는 규칙이에요.`));
        const ordered = blocks.filter(b => ln.includes(norm(b.a)));
        for (let i = 1; i < ordered.length; i++) {
          if (ln.indexOf(norm(ordered[i].a)) < ln.indexOf(norm(ordered[i - 1].a))) {
            add(year >= 2025 ? SEV.ERROR : SEV.INFO, '순위순 정렬', MAIN, ordered[i].row, `${year}년 블록 순서가 목록 순위와 달라요: '${ordered[i - 1].a}' 다음에 '${ordered[i].a}'가 나와요(목록상 ${ln.indexOf(norm(ordered[i].a)) + 1}위).`);
          }
        }
      }
    }

    // ---------- 두 시트 대조 ----------
    if (sub) {
      const S = sub.rows, SC = sub.col;
      if (S.length !== R.length) add(SEV.ERROR, '두 시트 불일치', SUB, null, `행 수가 달라요: '${MAIN}' ${R.length}행 / '${SUB}' ${S.length}행`);
      const n = Math.min(S.length, R.length);
      const fields = ['year', 'artist', 'type', 'start_date', 'end_date', 'days', 'country', 'city', 'location', 'capacity', 'attendance', 'ticket_price', 'remarks'];
      // 두 시트 변수 구성 동일 여부 (2026-10-21 박사님 요청)
      const missing = Object.keys(C).filter(h => h && !SC[h]);
      if (missing.length) add(SEV.ERROR, '두 시트 변수 불일치', SUB, null, `'${SUB}' 시트에 없는 변수: ${missing.join(', ')} — 두 시트 변수는 같아야 해요.`);
      let reported = 0;
      for (let i = 0; i < n; i++) {
        const a = R[i], b = S[i];
        const diff = fields.filter(f => SC[f] && C[f] && String(val(a, f) ?? '') !== String(val(b, f) ?? '') && !(val(a, f) === '-' && val(b, f) === undefined));
        if (diff.length && reported < 300) {
          reported++;
          add(SEV.ERROR, '두 시트 불일치', SUB, b, `같은 행인데 다른 값: ${diff.map(f => `${f} [${val(a, f) ?? '공란'} / ${val(b, f) ?? '공란'}]`).join(', ')}`);
        }
        if (SC['capacity*days']) {
          const f = (b.cells['capacity*days'] || {}).f;
          if (f && normF(f) !== `${SC['capacity']}${b.r}*${SC['days']}${b.r}`) add(SEV.ERROR, '수식 깨짐', SUB, b, `capacity*days 수식이 다른 행을 참조해요: =${f}`);
        }
        const sd = toDate(val(a, 'start_date'));
        const q = val(b, 'qt');
        if (sd && isNum(q) && q !== Math.floor(sd.getUTCMonth() / 3) + 1) add(SEV.ERROR, '분기(qt) 오류', SUB, b, `qt=${q}인데 start_date(${fmtDate(sd)}) 기준 ${Math.floor(sd.getUTCMonth() / 3) + 1}분기예요.`);
        if (val(b, 'no') !== b.r - 1) add(SEV.ERROR, 'no 연번', SUB, b, `no가 ${val(b, 'no')}인데 ${b.r - 1}이어야 해요.`);
      }
    }

    const stats = {
      rows: R.length,
      years: [...new Set(R.map(r => val(r, 'year')).filter(isNum))].sort(),
      error: issues.filter(i => i.sev === SEV.ERROR).length,
      warn: issues.filter(i => i.sev === SEV.WARN).length,
      info: issues.filter(i => i.sev === SEV.INFO).length,
    };
    return { issues, stats, venues };
  }

  function parseVenueMaster(wb) {
    const XLSX = root.XLSX;
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
    const m = new Map();
    rows.forEach(r => {
      const loc = r.location || r['공연장'], city = r.city || r['도시'];
      const cap = Number(r.capacity || r['capacity(확정)'] || r['수용인원']);
      if (loc && isFinite(cap) && cap > 0) m.set(norm(loc) + '|' + norm(city), { capacity: cap, source: r.source || r['출처'] || '' });
    });
    return m;
  }

  const api = { run, parseVenueMaster, SEV, norm, fmtDate };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  root.KpopChecker = api;
})(typeof window !== 'undefined' ? window : globalThis);
