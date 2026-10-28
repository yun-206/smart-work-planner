/* 스크롤 스토리텔링: 히어로 패럴랙스, 문장 하이라이트, 기능 챕터 단계 전환, 상단 탭 위치 표시 */
(() => {
  const reduce = matchMedia('(prefers-reduced-motion: reduce)').matches;
  const clamp = (v, a = 0, b = 1) => Math.min(b, Math.max(a, v));

  /* ---------- 소개 문장을 어절 단위로 나눔 ---------- */
  const statement = document.getElementById('statement');
  let words = [];
  if (statement) {
    const text = statement.textContent.trim().replace(/\s+/g, ' ');
    statement.setAttribute('aria-label', text);
    statement.innerHTML = text.split(' ').map((w) => `<span class="w" aria-hidden="true">${w}</span>`).join(' ');
    words = [...statement.querySelectorAll('.w')];
  }

  /* ---------- 요소 ---------- */
  const heroWord = document.querySelector('.hero__word');
  const chapters = [...document.querySelectorAll('.chapter')].map((el) => ({
    el,
    steps: [...el.querySelectorAll('.steps li')],
    frames: [...el.querySelectorAll('.frame')],
    current: -1,
  }));
  const seg = document.querySelector('.seg');
  const thumb = document.getElementById('segThumb');
  const tabs = [...document.querySelectorAll('.seg .tab')];
  const zones = [...document.querySelectorAll('[data-tab]:not(.tab)')];
  let activeTab = null;

  function setTab(t) {
    if (t === activeTab) return;
    activeTab = t;
    tabs.forEach((a) => {
      const on = a.dataset.tab === t;
      a.classList.toggle('on', on);
      if (on) a.setAttribute('aria-current', 'true'); else a.removeAttribute('aria-current');
    });
    seg.classList.toggle('has-active', !!t);
    placeThumb();
  }
  function placeThumb() {
    const on = tabs.find((a) => a.dataset.tab === activeTab);
    if (!on) return;
    thumb.style.width = on.offsetWidth + 'px';
    thumb.style.transform = `translateX(${on.offsetLeft - 4}px)`;
  }

  function setStep(ch, i) {
    if (ch.current === i) return;
    ch.current = i;
    ch.steps.forEach((s, k) => {
      s.classList.toggle('is-active', k === i);
      const btn = s.querySelector('.steps__btn');
      if (btn) { if (k === i) btn.setAttribute('aria-current', 'step'); else btn.removeAttribute('aria-current'); }
    });
    ch.frames.forEach((f, k) => f.classList.toggle('is-active', k === i));
  }

  /* ---------- 매 프레임 ---------- */
  function frame() {
    const vh = window.innerHeight;

    // 히어로 단어: 스크롤할수록 살짝 올라가며 작아짐
    if (heroWord && !reduce) {
      const p = clamp(window.scrollY / vh);
      heroWord.style.transform = `translateY(${-60 * p}px) scale(${1 - 0.06 * p})`;
      heroWord.style.opacity = String(1 - 0.5 * p);
    }

    // 소개 문장: 화면 가운데를 지나는 만큼 어절이 밝아짐
    if (words.length) {
      const r = statement.getBoundingClientRect();
      const p = reduce ? 1 : clamp((vh * 0.8 - r.top) / (r.height + vh * 0.35));
      const lit = Math.round(p * words.length);
      words.forEach((w, i) => w.classList.toggle('lit', i < lit));
    }

    // 챕터: 고정 구간 진행률을 3단계로
    chapters.forEach((ch) => {
      const r = ch.el.getBoundingClientRect();
      const p = clamp(-r.top / Math.max(1, ch.el.offsetHeight - vh));
      setStep(ch, Math.min(ch.steps.length - 1, Math.floor(p * ch.steps.length * 0.999)));
    });

    // 상단 탭: 화면 40% 지점에 걸린 기능
    const probe = vh * 0.4;
    const z = zones.find((el) => { const r = el.getBoundingClientRect(); return r.top <= probe && r.bottom > probe; });
    setTab(z ? z.dataset.tab : null);

    updateDock(vh);
  }

  let ticking = false;
  const onScroll = () => {
    if (ticking) return;
    ticking = true;
    requestAnimationFrame(() => { frame(); ticking = false; });
  };
  window.addEventListener('scroll', onScroll, { passive: true });
  window.addEventListener('resize', () => { placeThumb(); onScroll(); });
  if (document.fonts) document.fonts.ready.then(placeThumb);

  /* ---------- 단계 번호 클릭: 그 단계에 해당하는 스크롤 위치로 이동 ---------- */
  chapters.forEach((ch) => ch.steps.forEach((s, k) => {
    const btn = s.querySelector('.steps__btn');
    if (!btn) return;
    btn.addEventListener('click', () => {
      const range = ch.el.offsetHeight - window.innerHeight;
      if (range <= 0) { setStep(ch, k); return; }   // 동작 줄이기 모드: 스크롤 구간이 없으니 바로 전환
      const top = ch.el.getBoundingClientRect().top + window.scrollY + range * (k + 0.5) / ch.steps.length;
      window.scrollTo({ top, behavior: reduce ? 'auto' : 'smooth' });
    });
  }));

  /* ---------- 기능으로 이동 + 도착 표시 ---------- */
  function go(hash) {
    const el = document.querySelector(hash);
    if (!el) return;
    el.scrollIntoView({ behavior: reduce ? 'auto' : 'smooth', block: 'start' });
    history.replaceState(null, '', hash);
    if (!el.classList.contains('tool')) return;
    let done = false;
    const arrive = () => {
      if (done) return; done = true;
      window.removeEventListener('scrollend', arrive);
      el.classList.remove('is-arrived'); void el.offsetWidth; el.classList.add('is-arrived');
      setTimeout(() => el.classList.remove('is-arrived'), 1900);
      const h = el.querySelector('.tool__head h2'); if (h) h.focus({ preventScroll: true });   // 키보드·스크린리더도 도착 위치에서 시작
    };
    if ('onscrollend' in window && !reduce) { window.addEventListener('scrollend', arrive); setTimeout(arrive, 1500); }
    else setTimeout(arrive, reduce ? 0 : 700);
  }
  // 탭, 상단 버튼, 기능 소개의 바로 쓰기, 따라다니는 바로가기 모두 같은 방식으로 이동
  document.addEventListener('click', (e) => {
    const a = e.target.closest('.seg .tab, .bar__cta, .cta[href^="#"]');
    if (!a) return;
    e.preventDefault();
    go(a.getAttribute('href'));
  });

  /* ---------- 따라다니는 바로가기: 기능 소개 구간에서만 보임 ---------- */
  const dock = document.getElementById('dock');
  const dockName = document.getElementById('dockName');
  const dockGo = document.getElementById('dockGo');
  const dots = dock ? [...dock.querySelectorAll('.dock__dots i')] : [];
  const NAMES = { asm: '조립·순위', chk: '검증', cnv: '원본 변환' };
  const ORDER = ['asm', 'chk', 'cnv'];
  let dockKey = null;
  function updateDock(vh) {
    if (!dock) return;
    const ch = chapters.find((c) => { const r = c.el.getBoundingClientRect(); return r.top <= vh * 0.6 && r.bottom > vh * 0.6; });
    const k = ch ? ch.el.dataset.tab : null;
    const on = !!k;
    dock.classList.toggle('is-on', on);
    dock.setAttribute('aria-hidden', String(!on));
    dockGo.tabIndex = on ? 0 : -1;
    if (!k || k === dockKey) return;
    dockKey = k;
    dockName.textContent = NAMES[k];
    dockName.classList.remove('swap'); void dockName.offsetWidth; dockName.classList.add('swap');
    dockGo.setAttribute('href', '#sec-' + k);
    const idx = ORDER.indexOf(k);
    dots.forEach((d, i) => { d.classList.toggle('on', i === idx); d.classList.toggle('done', i < idx); });
  }

  frame();
})();

/* 다크 / 라이트 전환 스위치 */
(() => {
  const btn = document.getElementById('themeToggle');
  if (!btn) return;
  const root = document.documentElement;
  const meta = document.querySelector('meta[name="theme-color"]');
  const reduce = matchMedia('(prefers-reduced-motion: reduce)').matches;
  const sync = () => {
    const light = root.getAttribute('data-theme') === 'light';
    btn.setAttribute('aria-checked', String(light));
    btn.setAttribute('aria-label', light ? '다크 모드로 바꾸기' : '라이트 모드로 바꾸기');
    if (meta) meta.setAttribute('content', light ? '#f5f6f8' : '#121212');
  };
  let t;
  btn.addEventListener('click', () => {
    const light = root.getAttribute('data-theme') !== 'light';
    if (!reduce) { root.classList.add('theme-anim'); clearTimeout(t); t = setTimeout(() => root.classList.remove('theme-anim'), 450); }
    if (light) root.setAttribute('data-theme', 'light'); else root.removeAttribute('data-theme');
    try { localStorage.setItem('tourdata-theme', light ? 'light' : 'dark'); } catch (e) {}
    sync();
  });
  sync();
})();
