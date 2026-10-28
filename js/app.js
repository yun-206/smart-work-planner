/* K-pop 해외공연 데이터 도구 — 화면 동작
   조립·검증·변환 로직은 assembler.js / checker.js에 있고, 여기서는 화면만 다룹니다. */
const $ = (id) => document.getElementById(id);
let mainWb = null, venueMaster = null, result = null;
const today = new Date();
$('refDate').value = today.toISOString().slice(0, 10);

function bindDrop(drop, input, onFile) {
  const run = async (file) => {
    if (!file) return;
    drop.classList.remove('has-file', 'is-error');
    drop.setAttribute('aria-busy', 'true');
    const ok = await onFile(file);
    drop.removeAttribute('aria-busy');
    drop.classList.add(ok === false ? 'is-error' : 'has-file');
  };
  drop.onclick = () => input.click();
  drop.onkeydown = (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); input.click(); } };
  input.onchange = () => { run(input.files[0]); input.value = ''; };
  drop.ondragover = (e) => { e.preventDefault(); drop.classList.add('over'); };
  drop.ondragleave = () => drop.classList.remove('over');
  drop.ondrop = (e) => { e.preventDefault(); drop.classList.remove('over'); run(e.dataTransfer.files[0]); };
}
const readWb = (file) => file.arrayBuffer().then(buf => XLSX.read(buf, { cellFormula: true, sheetStubs: true, cellNF: true }));

bindDrop($('dropMain'), $('fileMain'), async (f) => {
  $('status').textContent = '불러오는 중…';
  try { mainWb = await readWb(f); $('mainName').textContent = f.name; $('runBtn').disabled = false; $('status').textContent = `시트: ${mainWb.SheetNames.join(', ')}`; }
  catch (e) { $('status').textContent = '파일을 읽지 못했어요: ' + e.message; return false; }
});
bindDrop($('dropVenue'), $('fileVenue'), async (f) => {
  try { venueMaster = KpopChecker.parseVenueMaster(await readWb(f)); $('venueName').textContent = `${f.name}, 공연장 ${venueMaster.size}곳`; }
  catch (e) { $('venueName').textContent = '읽기 실패: ' + e.message; return false; }
});

const SEVS = [['오류', 'e'], ['확인필요', 'w'], ['참고', 'i']];
let sevOn = new Set(['오류', '확인필요']), checkOn = null;

$('runBtn').onclick = () => runCheck();
function runCheck() {
  const [y, m, d] = $('refDate').value.split('-').map(Number);
  const t0 = performance.now();
  result = KpopChecker.run(mainWb, { refDate: new Date(y, m - 1, d), venueMaster, strictFrom: Number($('strictFrom').value) || 2025 });
  $('status').textContent = `${result.stats.rows.toLocaleString()}행 검사 완료 (${Math.round(performance.now() - t0)}ms)`;
  const yrs = result.stats.years || [];
  $('yearFrom').value = Number($('strictFrom').value) || yrs[0] || '';
  $('yearTo').value = yrs[yrs.length - 1] || '';
  checkOn = null;
  $('resultCard').classList.remove('hide');
  render();
  return result;
}

function filtered() {
  const yf = Number($('yearFrom').value) || 0, yt = Number($('yearTo').value) || 9999, q = $('q').value.trim().toLowerCase();
  return result.issues.filter(i => {
    const y = Number(i.year);
    const yearOk = !i.year || (y >= yf && y <= yt) || i.artist === '(여러 아티스트)';
    const qOk = !q || (`${i.artist} ${i.detail} ${i.check}`).toLowerCase().includes(q);
    return yearOk && qOk;
  });
}

function render() {
  const base = filtered();
  const cnt = (s) => base.filter(i => i.sev === s).length;
  $('stats').innerHTML = SEVS.map(([s, c]) => `<div class="stat ${c}"><div>${s}</div><div class="n">${cnt(s).toLocaleString()}</div></div>`).join('')
    + `<div class="stat"><div>검사한 행</div><div class="n">${result.stats.rows.toLocaleString()}</div></div>`;
  $('sevChips').innerHTML = SEVS.map(([s]) => `<span class="chip ${sevOn.has(s) ? 'on' : ''}" data-sev="${s}">${s} ${cnt(s)}</span>`).join('');
  const checks = {};
  base.filter(i => sevOn.has(i.sev)).forEach(i => checks[i.check] = (checks[i.check] || 0) + 1);
  $('checkChips').innerHTML = `<span class="chip ${!checkOn ? 'on' : ''}" data-check="">전체</span>` +
    Object.entries(checks).sort((a, b) => b[1] - a[1]).map(([c, n]) => `<span class="chip ${checkOn === c ? 'on' : ''}" data-check="${c}">${c} ${n}</span>`).join('');
  const rows = base.filter(i => sevOn.has(i.sev) && (!checkOn || i.check === checkOn));
  const cls = { '오류': 'e', '확인필요': 'w', '참고': 'i' };
  const esc = (s) => String(s ?? '').replace(/[&<>]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c]));
  $('tbody').innerHTML = rows.slice(0, 2000).map(i => `<tr><td class="sev ${cls[i.sev]}">${i.sev}</td><td>${esc(i.check)}</td><td class="muted">${i.sheet === '해외 공연 현황' ? '본' : '(2)'}</td><td>${i.excelRow}</td><td>${esc(i.no)}</td><td>${esc(i.year)}</td><td>${esc(i.artist)}</td><td>${esc(i.detail)}</td></tr>`).join('');
  $('shown').textContent = `${rows.length.toLocaleString()}건 표시${rows.length > 2000 ? ' (화면엔 앞 2,000건, 다운로드엔 전부)' : ''}`;
  document.querySelectorAll('[data-sev]').forEach(el => el.onclick = () => { const s = el.dataset.sev; sevOn.has(s) ? sevOn.delete(s) : sevOn.add(s); render(); });
  document.querySelectorAll('[data-check]').forEach(el => el.onclick = () => { checkOn = el.dataset.check || null; render(); });
}
['yearFrom', 'yearTo', 'q'].forEach(id => $(id).oninput = () => result && render());

function saveWb(wb, name) {
  const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
  const url = URL.createObjectURL(new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  const a = document.createElement('a'); a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}
$('dlXlsx').onclick = () => {
  const rows = filtered().map(i => ({ 등급: i.sev, 검사항목: i.check, 시트: i.sheet, 엑셀행: i.excelRow, no: i.no, 연도: i.year, 아티스트: i.artist, 내용: i.detail, 제안값: i.suggest ?? '' }));
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows);
  ws['!cols'] = [{ wch: 8 }, { wch: 22 }, { wch: 16 }, { wch: 8 }, { wch: 7 }, { wch: 7 }, { wch: 16 }, { wch: 90 }, { wch: 10 }];
  XLSX.utils.book_append_sheet(wb, ws, '검증 결과');
  saveWb(wb, `검증리포트_${$('refDate').value}.xlsx`);
};
$('dlVenue').onclick = () => {
  const rows = result.venues.sort((a, b) => b.count - a.count).map(v => ({
    location: v.location, city: v.city, country: v.country,
    capacity: v.values.length === 1 ? v.values[0] : '',
    source: '',
    '기존 기입값': v.values.join(' / '),
    '상태': v.values.length > 1 ? '값 불일치 — 공식 수용인원 확인 필요' : (v.values.length ? '단일값(출처 확인 필요)' : 'capacity 없음'),
    '사용 행 수': v.count,
  }));
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows);
  ws['!cols'] = [{ wch: 40 }, { wch: 16 }, { wch: 12 }, { wch: 10 }, { wch: 40 }, { wch: 24 }, { wch: 30 }, { wch: 9 }];
  XLSX.utils.book_append_sheet(wb, ws, '공연장 마스터');
  saveWb(wb, '공연장_마스터_초안.xlsx');
};

/* ---------- 기능 이동 (한 페이지 스크롤) ---------- */
const reduceMotion = matchMedia('(prefers-reduced-motion: reduce)').matches;
function showTab(t) {
  const el = document.getElementById('sec-' + t);
  if (el) el.scrollIntoView({ behavior: reduceMotion ? 'auto' : 'smooth', block: 'start' });
}
// 상단 바: 스크롤하면 배경
const bar = $('bar');
const onBar = () => bar.classList.toggle('is-scrolled', window.scrollY > 8);
window.addEventListener('scroll', onBar, { passive: true }); onBar();
const esc2 = (s) => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const N = (v) => (typeof v === 'number' && v ? Math.round(v).toLocaleString() : '');
function saveBuf(buf, name) {
  const url = URL.createObjectURL(new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  const a = document.createElement('a'); a.href = url; a.download = name; document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
}
const ymd = () => { const d = new Date(); return String(d.getFullYear()).slice(2) + String(d.getMonth() + 1).padStart(2, '0') + String(d.getDate()).padStart(2, '0'); };

/* ---------- ③ 원본 변환 ---------- */
let legWb = null;
bindDrop($('dropLeg'), $('fileLeg'), async (f) => {
  try {
    legWb = await readWb(f);
    if (!legWb.Sheets['해외 공연 현황']) throw new Error("'해외 공연 현황' 시트가 없어요");
    $('legName').textContent = f.name; $('cnvBtn').disabled = false; $('cnvStatus').textContent = '';
  } catch (e) { $('legName').textContent = '읽기 실패: ' + e.message; return false; }
});
$('cnvBtn').onclick = async () => {
  try {
    const src = KpopAssembler.convertLegacy(legWb);
    const buf = await KpopAssembler.writeSource(src);
    saveBuf(buf, `원본데이터_${ymd()}.xlsx`);
    $('cnvStatus').textContent = `공연목록 ${src.main.length.toLocaleString()}행, 솔로부록 ${src.appendix.length}행, 공연장 ${src.master.length}곳으로 변환했어요.`;
  } catch (e) { $('cnvStatus').textContent = '변환 실패: ' + e.message; }
};

/* ---------- ① 조립·순위 ---------- */
let srcData = null, asmPlan = null, outBuf = null, updData = null, merged = null, updLog = [];
$('aRef').value = today.toISOString().slice(0, 10);
bindDrop($('dropSrc'), $('fileSrc'), async (f) => {
  try {
    const wb = await readWb(f);
    if (!wb.Sheets['공연목록'] && wb.Sheets['해외 공연 현황']) throw new Error('결과 엑셀이에요. ③ 원본 변환 탭에서 먼저 변환하세요');
    srcData = KpopAssembler.readSource(wb);
    const st = srcData.settings || {};
    if (st['자동순위연도']) $('aYears').value = st['자동순위연도'];
    if (st['상위N']) $('aTop').value = st['상위N'];
    $('srcName').textContent = `${f.name}, 공연 ${srcData.main.length.toLocaleString()}행, 공식누적 ${srcData.official.length}건`;
    $('asmBtn').disabled = false;
  } catch (e) { $('srcName').textContent = '읽기 실패: ' + e.message; return false; }
});

bindDrop($('dropUpd'), $('fileUpd'), async (f) => {
  try {
    updData = KpopAssembler.readUpdate(await readWb(f));
    $('updName').textContent = `${f.name}, 공연 ${updData.rows.length}행, 공식누적 ${updData.official.length}건, 목록 ${updData.list.length}건${updData.master.length ? `, 공연장마스터 ${updData.master.length}건` : ''}`;
  } catch (e) { updData = null; $('updName').textContent = '읽기 실패: ' + e.message; return false; }
});
$('dlTpl').onclick = async (e) => { e.preventDefault(); saveBuf(await KpopAssembler.writeUpdate([], []), 'update_template.xlsx'); };

$('asmBtn').onclick = async () => {
  $('asmStatus').textContent = '만드는 중…';
  try {
    merged = srcData; updLog = [];
    if (updData) { const r = KpopAssembler.applyUpdate(srcData, updData, { autoYears: $('aYears').value }); merged = r.src; updLog = r.log; }
    merged.settings = Object.assign({}, merged.settings, { 기준일: $('aRef').value, 자동순위연도: $('aYears').value || merged.settings['자동순위연도'], 상위N: Number($('aTop').value) || 15 });
    asmPlan = KpopAssembler.plan(merged, { refDate: $('aRef').value, autoYears: $('aYears').value, topN: Number($('aTop').value) });
    outBuf = await KpopAssembler.writeOutput(merged, asmPlan);
    // 만든 결과를 바로 검증
    mainWb = XLSX.read(outBuf, { type: 'array', cellFormula: true, sheetStubs: true, cellNF: true });
    $('refDate').value = $('aRef').value; $('mainName').textContent = '조립 결과(방금 생성)'; $('runBtn').disabled = false; $('dropMain').classList.add('has-file');
    const chk = runCheck();
    renderAsm(chk);
    $('asmStatus').textContent = '';
    $('asmResult').scrollIntoView({ behavior: reduceMotion ? 'auto' : 'smooth', block: 'start' });
  } catch (e) { console.error(e); $('asmStatus').textContent = '실패: ' + e.message; }
};

function renderAsm(chk) {
  const p = asmPlan;
  $('asmResult').classList.remove('hide');
  const c = (s) => chk.issues.filter(i => i.sev === s && Number(i.year) >= (Number($('strictFrom').value) || 2025)).length;
  const tagged = p.out.filter(r => r.__tag).length;
  $('asmStats').innerHTML = [
    ['결과 행', p.out.length + p.appendix.length, ''], ['[예정]·[일부 예정]', tagged, ''],
    ['검증 오류(엄격 연도)', c('오류'), 'e'], ['확인필요(엄격 연도)', c('확인필요'), 'w'],
  ].map(([t, n, k]) => `<div class="stat ${k}"><div>${t}</div><div class="n">${n.toLocaleString()}</div></div>`).join('');
  let h = '';
  p.autoYears.forEach(y => {
    const list = p.listRank[y] || [];
    const rk = p.rankings[y] || [];
    const changed = rk.slice(0, p.topN).some((x, i) => list[i] !== x.artist);
    const lrow = merged.list && merged.list.rows.find(r => Number(r[0]) === y);
    const drop = (p.subDrop || {})[y] || [];
    const subs = lrow ? KpopAssembler.listSubs(lrow).filter(a => !drop.includes(KpopChecker.norm(a))) : [];
    h += `<h3 style="margin:18px 0 4px;font-size:16px">${y}년 ${changed ? '<span class="up">기존 목록과 순위가 달라요</span>' : '<span class="same">기존 목록과 같음</span>'}</h3>
    ${subs.length ? `<p class="note" style="margin:0 0 6px">목록 SUB 칸: ${subs.map(esc2).join(', ')}</p>` : ''}
    <div class="tablewrap"><table><thead><tr><th class="num">순위</th><th>아티스트</th><th class="num">합계</th><th class="num">공식누적</th><th class="num">공식 관객수</th><th class="num">수용인원×일수</th><th class="num">행</th><th class="num">수치 없는 행</th><th>기존 목록</th></tr></thead><tbody>` +
      rk.map(x => {
        const was = list.indexOf(x.artist) + 1;
        const mv = !was ? '<span class="up">신규</span>' : was === x.rank ? `<span class="same">${was}위</span>` : was > x.rank ? `<span class="up">${was}위 → ▲${was - x.rank}</span>` : `<span class="down">${was}위 → ▼${x.rank - was}</span>`;
        const out = x.rank > p.topN ? ' style="opacity:.55"' : '';
        return `<tr${out}><td class="num">${x.rank}</td><td>${esc2(x.artist)}${x.rank > p.topN ? ' <span class="muted">(상세 제외)</span>' : ''}</td><td class="num"><b>${N(x.total)}</b></td><td class="num" title="${esc2(x.usedOfficial.join(', '))}">${N(x.byOfficial)}</td><td class="num">${N(x.byAtt)}</td><td class="num">${N(x.byCap)}</td><td class="num">${x.rows}</td><td class="num">${x.blank || ''}</td><td>${mv}</td></tr>`;
      }).join('') +
      list.filter(a => !rk.some(x => x.artist === a)).map(a => `<tr><td></td><td>${esc2(a)}</td><td colspan="6" class="muted">원본에 ${y}년 행이 없음 → 결과 목록에서 빠져요</td><td><span class="down">목록에서 제외</span></td></tr>`).join('') +
      `</tbody></table></div>`;
  });
  $('rankWrap').innerHTML = h || '<p class="note">자동 순위 연도가 없어요. 원본 순서 그대로 만들었어요.</p>';
  $('dlNewSrc').classList.toggle('hide', !updData);
  if (updData) {
    const cls = { 추가: 'r-add', 수정: 'r-mod', 목록: 'r-mod', 삭제: 'r-del', 오류: 'r-err', 건너뜀: 'r-err' };
    const cnt = (k) => updLog.filter(l => l.결과 === k).length;
    const isList = (l) => String(l.행).startsWith('목록'), isMaster = (l) => String(l.행).startsWith('마스터');
    const tbl = (logs, withLoc) => `<div class="tablewrap tablewrap--log"><table><thead><tr><th>행</th><th>결과</th><th>연도</th><th>아티스트</th><th>${withLoc ? '대상' : '칸'}</th><th>내용</th><th>근거</th></tr></thead><tbody>` +
      logs.map(l => `<tr><td>${esc2(l.행)}</td><td class="${cls[l.결과] || ''}">${esc2(l.결과)}</td><td>${esc2(l.year)}</td><td>${esc2(l.artist)}</td><td>${esc2(l.대상)}</td><td>${esc2(l.내용)}</td><td class="muted" style="max-width:280px;word-break:break-all">${esc2(l.근거)}</td></tr>`).join('') + `</tbody></table></div>`;
    const rowLog = updLog.filter(l => !isList(l) && !isMaster(l)), listLog = updLog.filter(isList), masterLog = updLog.filter(isMaster);
    const rc = (k) => rowLog.filter(l => l.결과 === k).length, mc = masterLog.filter(l => /추가|수정|삭제/.test(l.결과)).length;
    $('updWrap').innerHTML = `<h3 style="margin:18px 0 4px;font-size:16px">업데이트 반영 — 추가 ${rc('추가')} · 수정 ${rc('수정')} · 삭제 ${rc('삭제')}${cnt('목록') ? ` · 목록 ${cnt('목록')}` : ''}${mc ? ` · 공연장마스터 ${mc}` : ''}${cnt('오류') + cnt('건너뜀') ? ` · <span class="up">확인 필요 ${cnt('오류') + cnt('건너뜀')}</span>` : ''}</h3>` +
      (rowLog.length ? tbl(rowLog, true) : '') +
      (listLog.length ? `<h4 style="margin:16px 0 0;font-size:14px">연도별 아티스트 목록 변경</h4>` + tbl(listLog, false) : '') +
      (masterLog.length ? `<h4 style="margin:16px 0 0;font-size:14px">공연장마스터 변경</h4>` + tbl(masterLog, true) : '');
  } else $('updWrap').innerHTML = '';
  $('asmNotes').innerHTML = p.notes.length ? `<p class="note"><b>참고</b><br>${p.notes.map(n => esc2(`${n.year} ${n.artist} — ${n.kind}: ${n.detail}`)).join('<br>')}</p>` : '';
  $('asmTitle').textContent = `순위 산정 결과 (기준일 ${$('aRef').value})`;
}
$('dlOut').onclick = () => outBuf && saveBuf(outBuf, `k-pop_자료조사_데이터_${ymd()}.xlsx`);
$('toChk').onclick = () => $('resultCard').scrollIntoView({ behavior: reduceMotion ? 'auto' : 'smooth', block: 'start' });
$('dlNewSrc').onclick = async () => saveBuf(await KpopAssembler.writeSource(merged), `원본데이터_${$('aRef').value.slice(2).replace(/-/g, '')}.xlsx`);
$('dlBasis').onclick = () => {
  const rows = [];
  asmPlan.autoYears.forEach(y => (asmPlan.rankings[y] || []).forEach(x => rows.push({ 연도: y, 순위: x.rank, 아티스트: x.artist, 합계: x.total, 공식누적: x.byOfficial, '공식누적 내역': x.usedOfficial.join(', '), '공식 관객수': x.byAtt, '수용인원×일수': x.byCap, 행수: x.rows, '수치 없는 행': x.blank, 상세수록: x.rank <= asmPlan.topN ? 'O' : '제외', '기존 목록 순위': ((asmPlan.listRank[y] || []).indexOf(x.artist) + 1) || '' })));
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.json_to_sheet(rows); ws['!cols'] = [{ wch: 6 }, { wch: 5 }, { wch: 16 }, { wch: 11 }, { wch: 11 }, { wch: 36 }, { wch: 11 }, { wch: 13 }, { wch: 6 }, { wch: 10 }, { wch: 8 }, { wch: 12 }];
  XLSX.utils.book_append_sheet(wb, ws, '순위 산정 근거');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(merged.official.map(o => ({ 집계묶음: o['집계묶음'], 연도: o.year, 아티스트: o.artist, 공식관객수: o['공식관객수'], 출처: o['출처'], 설명: o['설명'] }))), '공식누적 출처');
  saveWb(wb, `순위산정근거_${$('aRef').value}.xlsx`);
};
