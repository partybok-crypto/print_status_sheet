/*****************************************************
 * 세탁공정 배정 코어 (화면과 분리된 순수 계산 모듈)
 *
 *  - 수량 산정:  기사 등록(실수량) > AI 예상(같은 요일 최근 기록의 중앙값) > 수량 확인 필요
 *  - 배정:       4대 세탁기에 회차(배치)를 자동 배정. 안전 규칙은 하드 제약, 그 안에서 종료시간을 최소화
 *  - 검증:       중복·누락·용량·회차당 3곳·미분류 합세탁·수량 합계 등. 하나라도 어긋나면 plan.errors 에 담긴다
 *
 *  세탁기 용량·합세탁 규칙은 전부 DEFAULT_CONFIG(설정값)이며 계산 코드와 분리되어 있다.
 *  브라우저에서는 window.WashCore, Node 에서는 require() 로 쓴다.
 *****************************************************/
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.WashCore = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  // ---------------------------------------------------------------- 설정값
  var DEFAULT_CONFIG = {
    // 세탁기: 번호, 최대 수용량(포대), 1회 세탁 시간(분). 기계별 실제 시간을 알게 되면 cycleMin 만 바꾸면 된다.
    machines: [
      { no: 1, cap: 5, cycleMin: 20 },
      { no: 2, cap: 3, cycleMin: 20 },
      { no: 3, cap: 2, cycleMin: 20 },
      { no: 4, cap: 1, cycleMin: 20 }
    ],
    maxCompanies: 3,          // 한 회차 최대 거래처 수
    maxUnclassified: 1,       // 한 회차 최대 미분류 거래처 수
    historyWindow: 4,         // AI 예상에 쓰는 같은 요일 최근 기록 수
    medianMode: 'low',        // 짝수 개일 때 중앙값: low(작은 쪽) | high | round
    maxQtyValid: 30,          // 이보다 큰 값은 수거량이 아닌 값으로 보고 무시
    // 업종 → 같이 돌려도 되는 업종군. 목록에 없는 업종은 자기 자신이 한 군.
    industryFamily: { '헬스장': 'fit', '필라테스': 'fit', '미용실': 'hair', '피부관리': 'skin', '마사지': 'skin' },
    // 관리자가 승인한 예외 조합(업종이 달라도 같은 회차 허용). 실제 투입 전 식별·세탁조건 확인 필요.
    exceptionPairs: [
      { a: '142', b: '169', note: '국군수송+비주얼 — 수건 식별·세탁조건 확인 후 투입' }
    ],
    mixBlockedPairs: [],      // 절대 같이 돌리지 않을 거래처 번호 쌍 [{a,b}]
    soloCompanies: [],        // 항상 단독으로 돌릴 거래처 번호
    // 합세탁 칸(J열)에 이 값이 적혀 있으면 단독 세탁 대상
    soloWashValues: ['체크', '신규', '확인', '불가', '단독', 'X'],
    searchIterations: 14000,
    searchRestarts: 2
  };

  function mergeConfig(over) {
    var c = {};
    Object.keys(DEFAULT_CONFIG).forEach(function (k) { c[k] = DEFAULT_CONFIG[k]; });
    if (over) Object.keys(over).forEach(function (k) { c[k] = over[k]; });
    return c;
  }

  // ---------------------------------------------------------------- 그룹 표시
  var GROUP_LETTER = { A: '가', B: '나', C: '다' };
  var GROUP_ORDER = ['A', 'B', 'B1', 'B2', 'C', 'U'];

  function normalizeGroup(raw) {
    var g = String(raw === null || raw === undefined ? '' : raw).replace(/\s+/g, '').toUpperCase();
    return /^[ABC]\d{0,2}$/.test(g) ? g : 'U';
  }
  function groupLabel(g) {
    if (g === 'U') return '미분류';
    var n = g.slice(1);
    return (GROUP_LETTER[g.charAt(0)] || g.charAt(0)) + (n ? '-' + n : '') + ' 그룹';
  }
  function groupClass(g) {
    if (g === 'A') return 'a';
    if (g === 'B' || g === 'B1') return 'b';
    if (/^B\d/.test(g)) return 'b2';
    if (g.charAt(0) === 'C') return 'c';
    return 'u';
  }
  function groupRank(g) {
    var i = GROUP_ORDER.indexOf(g);
    return i === -1 ? 50 : i;
  }
  function isSoloRaw(raw, cfg) {
    var v = String(raw === null || raw === undefined ? '' : raw).replace(/\s+/g, '').toUpperCase();
    return cfg.soloWashValues.map(function (s) { return s.toUpperCase(); }).indexOf(v) !== -1;
  }
  function familyOf(industry, cfg) {
    var s = String(industry || '').replace(/\s+/g, '');
    return cfg.industryFamily[s] || s || '?';
  }
  function numKey(no) { // 번호 정렬용: "14_2" → 14.2
    var m = /^(\d+)(?:_(\d+))?/.exec(String(no));
    return m ? Number(m[1]) + (m[2] ? Number(m[2]) / 1000 : 0) : 1e9;
  }

  // ---------------------------------------------------------------- 수량 산정
  function medianOf(values, mode) {
    var s = values.slice().sort(function (a, b) { return a - b; });
    var n = s.length;
    if (!n) return null;
    if (n % 2) return s[(n - 1) / 2];
    var lo = s[n / 2 - 1], hi = s[n / 2];
    if (mode === 'high') return hi;
    if (mode === 'round') return Math.round((lo + hi) / 2 + 1e-9);
    return lo;
  }

  /**
   * entry: { actual: number|null, history: [{date, qty|null}, ...] }  (history는 최신순)
   * 반환: { qty, source: 'driver'|'ai'|'unknown', basis: { values, dates, median } }
   *  - actual 이 있으면(0 포함) 기사 등록이 AI 예상보다 우선
   *  - 이력이 하나도 없으면 unknown (임의로 1포대를 주지 않는다). 빈 칸과 실제 0은 다르다.
   */
  function estimateQty(entry, cfg) {
    cfg = cfg || DEFAULT_CONFIG;
    if (entry && entry.actual !== null && entry.actual !== undefined && entry.actual !== '') {
      return { qty: Number(entry.actual), source: 'driver', basis: { values: [Number(entry.actual)], dates: [], median: null, from: entry.actualFrom || '' } };
    }
    var values = [], dates = [];
    var hist = (entry && entry.history) || [];
    for (var i = 0; i < hist.length && values.length < cfg.historyWindow; i++) {
      var q = hist[i].qty;
      if (q === null || q === undefined || q === '') continue;
      values.push(Number(q)); dates.push(hist[i].date);
    }
    if (!values.length) return { qty: null, source: 'unknown', basis: { values: [], dates: [], median: null } };
    var med = medianOf(values, cfg.medianMode);
    return { qty: med, source: 'ai', basis: { values: values, dates: dates, median: med } };
  }

  /**
   * items: 화면의 거래처 목록(companyNumber, nickname, businessType, course, washGroup, combinedWash)
   * history: GAS 응답 { companies: { no: {actual, history} } }
   */
  function buildCompanies(items, history, config) {
    var cfg = mergeConfig(config);
    var comps = [], errors = [], seen = {};
    var map = (history && history.companies) || {};
    (items || []).forEach(function (it) {
      var no = String(it.companyNumber || '').trim();
      if (!no) { errors.push('거래처번호가 비어 있는 행이 있습니다(' + (it.nickname || '이름 없음') + ').'); return; }
      if (seen[no]) { errors.push('거래처 ' + no + ' 이(가) 오늘 목록에 중복되어 있습니다.'); return; }
      seen[no] = true;
      var est = estimateQty(map[no], cfg);
      var group = normalizeGroup(it.washGroup);
      var solo = cfg.soloCompanies.indexOf(no) !== -1 || (group === 'U' && isSoloRaw(it.combinedWash, cfg));
      comps.push({
        no: no, name: it.nickname || '', industry: it.businessType || '', family: familyOf(it.businessType, cfg),
        course: String(it.course || '').replace(/^-+/, ''), group: group, unclassified: group === 'U', solo: solo,
        qty: est.qty, source: est.source, basis: est.basis
      });
    });
    return { companies: comps, errors: errors };
  }

  // ---------------------------------------------------------------- 합세탁 가능 판정
  function pairKey(a, b) { return a < b ? a + '|' + b : b + '|' + a; }

  function makeRules(cfg) {
    var exc = {}, blocked = {};
    cfg.exceptionPairs.forEach(function (p) { exc[pairKey(String(p.a), String(p.b))] = p; });
    cfg.mixBlockedPairs.forEach(function (p) { blocked[pairKey(String(p.a), String(p.b))] = true; });
    return { exc: exc, blocked: blocked };
  }

  // 한 회차(chunk 목록)가 안전 규칙을 지키는가
  function cycleViolations(members, cfg, rules, maxCap) {
    var v = [];
    if (members.length > cfg.maxCompanies) v.push('회차당 최대 ' + cfg.maxCompanies + '곳 초과');
    var sum = 0, uncl = 0, grp = null, soloN = 0, nos = {};
    members.forEach(function (c) {
      sum += c.qty;
      if (nos[c.no]) v.push('같은 회차에 ' + c.no + ' 중복');
      nos[c.no] = true;
      if (c.solo) soloN++;
      if (c.unclassified) uncl++;
      else if (grp === null) grp = c.group;
      else if (grp !== c.group) v.push('서로 다른 그룹(' + grp + ',' + c.group + ') 합세탁');
    });
    if (sum > maxCap) v.push('수용량 초과(' + sum + '>' + maxCap + ')');
    if (uncl > cfg.maxUnclassified) v.push('미분류 ' + uncl + '곳 합세탁');
    if (soloN && members.length > 1) v.push('단독 세탁 거래처를 다른 거래처와 합세탁');
    for (var i = 0; i < members.length; i++) {
      for (var j = i + 1; j < members.length; j++) {
        var a = members[i], b = members[j], k = pairKey(a.no, b.no);
        if (rules.blocked[k]) v.push(a.no + '·' + b.no + ' 혼합 금지 조합');
        if (a.family !== b.family && !rules.exc[k]) v.push('업종이 다른 거래처 합세탁(' + a.industry + '·' + b.industry + ')');
      }
    }
    return v;
  }

  function cycleIsConditional(members, rules) {
    var uncl = 0, cls = 0, excUsed = false;
    members.forEach(function (c) { if (c.unclassified) uncl++; else cls++; });
    for (var i = 0; i < members.length; i++) {
      for (var j = i + 1; j < members.length; j++) {
        if (members[i].family !== members[j].family && rules.exc[pairKey(members[i].no, members[j].no)]) excUsed = true;
      }
    }
    return { conditional: (uncl > 0 && cls > 0) || excUsed, exception: excUsed };
  }

  // ---------------------------------------------------------------- 용량 가능성(홀 조건)
  // 세탁기를 용량 큰 순으로 놓고, 용량 j번째 이하로는 못 들어가는 회차 수가 그 기계들의 슬롯 합을 넘지 않아야 한다.
  function machineLevels(cfg) {
    return cfg.machines.slice().sort(function (a, b) { return b.cap - a.cap || a.no - b.no; });
  }
  function slotsFor(levels, budgetMin, lockedCount) {
    return levels.map(function (m) {
      return Math.max(0, Math.floor(budgetMin / m.cycleMin) - ((lockedCount && lockedCount[m.no]) || 0));
    });
  }
  function hallDeficit(sizes, levels, slots) {
    var deficit = 0, have = 0;
    for (var j = 0; j < levels.length; j++) {
      have += slots[j];
      var lower = j + 1 < levels.length ? levels[j + 1].cap : 0;
      var need = 0;
      for (var i = 0; i < sizes.length; i++) if (sizes[i] > lower) need++;
      if (need > have) deficit += need - have;
    }
    return deficit;
  }

  // ---------------------------------------------------------------- 난수(재현 가능)
  function makeRng(seed) {
    var a = seed >>> 0;
    return function () {
      a |= 0; a = (a + 0x6D2B79F5) | 0;
      var t = Math.imul(a ^ (a >>> 15), 1 | a);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }
  function hashStr(s) {
    var h = 2166136261;
    for (var i = 0; i < s.length; i++) { h ^= s.charCodeAt(i); h = Math.imul(h, 16777619); }
    return h >>> 0;
  }

  // ---------------------------------------------------------------- 배치 탐색(시뮬레이티드 어닐링)
  // chunks: [{no,qty,...}] → batches: [[chunkIndex,...]]
  function anneal(chunks, ctx, budgetMin, seed) {
    var cfg = ctx.cfg, rules = ctx.rules, levels = ctx.levels, maxCap = ctx.maxCap;
    var slots = slotsFor(levels, budgetMin, ctx.lockedCount);
    var rnd = makeRng(seed);
    var n = chunks.length;
    var okCache = {};

    function members(b) { return b.map(function (i) { return chunks[i]; }); }
    function okKey(b) { return b.slice().sort(function (x, y) { return x - y; }).join(','); }
    function batchOk(b) {
      var key = okKey(b), hit = okCache[key];
      if (hit !== undefined) return hit;
      var ms = members(b);
      var res = cycleViolations(ms, cfg, rules, maxCap).length === 0;
      okCache[key] = res;
      return res;
    }
    function batchCost(b) { // 조건부 합세탁 여부 + 코스 흩어짐
      var ms = members(b), cond = cycleIsConditional(ms, rules).conditional ? 1 : 0, cs = {}, k = 0;
      ms.forEach(function (c) { if (!cs[c.course]) { cs[c.course] = 1; k++; } });
      return { cond: cond, frag: k - 1 };
    }
    function energy(batches) {
      var sizes = [], cond = 0, frag = 0;
      for (var i = 0; i < batches.length; i++) {
        var s = 0;
        for (var j = 0; j < batches[i].length; j++) s += chunks[batches[i][j]].qty;
        sizes.push(s);
        var bc = batchCost(batches[i]);
        cond += bc.cond; frag += bc.frag;
      }
      var viol = hallDeficit(sizes, levels, slots);
      return { E: viol * 2000 + cond * 40 + batches.length * 10 + frag * 2, viol: viol };
    }

    var batches = [];
    for (var i = 0; i < n; i++) batches.push([i]);
    var cur = energy(batches), best = { batches: copy(batches), E: cur.E, viol: cur.viol };
    function copy(bs) { return bs.map(function (b) { return b.slice(); }); }

    var iters = cfg.searchIterations;
    var t0 = 30, t1 = 0.8;
    for (var it = 0; it < iters && n > 1; it++) {
      var temp = t0 * Math.pow(t1 / t0, it / iters);
      var next = null;
      var r = rnd();
      if (r < 0.65) {
        // 한 거래처를 다른 회차(또는 새 회차)로 옮긴다
        var si = Math.floor(rnd() * batches.length);
        var src = batches[si];
        var ci = Math.floor(rnd() * src.length);
        var dst = Math.floor(rnd() * (batches.length + 1));
        if (dst === si) continue;
        var moved = src[ci];
        if (dst < batches.length) {
          var cand = batches[dst].concat([moved]);
          if (!batchOk(cand)) continue;
          next = copy(batches);
          next[dst] = cand;
        } else {
          if (src.length === 1) continue;
          next = copy(batches);
          next.push([moved]);
        }
        next[si] = next[si].filter(function (x) { return x !== moved; });
        if (!next[si].length) next.splice(si, 1);
      } else {
        // 서로 다른 두 회차의 거래처 하나씩 맞바꾼다
        if (batches.length < 2) continue;
        var a = Math.floor(rnd() * batches.length), b = Math.floor(rnd() * batches.length);
        if (a === b) continue;
        var ia = Math.floor(rnd() * batches[a].length), ib = Math.floor(rnd() * batches[b].length);
        var ca = batches[a][ia], cb = batches[b][ib];
        var na = batches[a].map(function (x) { return x === ca ? cb : x; });
        var nb = batches[b].map(function (x) { return x === cb ? ca : x; });
        if (!batchOk(na) || !batchOk(nb)) continue;
        next = copy(batches);
        next[a] = na; next[b] = nb;
      }
      var e = energy(next);
      var d = e.E - cur.E;
      if (d <= 0 || rnd() < Math.exp(-d / temp)) {
        batches = next; cur = e;
        if (e.E < best.E) best = { batches: copy(batches), E: e.E, viol: e.viol };
      }
    }
    return best;
  }

  // ---------------------------------------------------------------- 배치 → 세탁기 배정
  function placeBatches(batches, chunks, ctx, budgetMin) {
    var levels = ctx.levels;
    var slots = slotsFor(levels, budgetMin, ctx.lockedCount);
    var used = levels.map(function () { return 0; });
    var list = batches.map(function (b, idx) {
      var ms = b.map(function (i) { return chunks[i]; });
      var size = ms.reduce(function (s, c) { return s + c.qty; }, 0);
      var classified = ms.filter(function (c) { return !c.unclassified; });
      var grp = classified.length ? classified[0].group : 'U';
      return { members: ms, size: size, group: grp, order: idx };
    });
    list.sort(function (a, b) {
      return (b.size - a.size) || (groupRank(a.group) - groupRank(b.group)) ||
        (numKey(a.members[0].no) - numKey(b.members[0].no)) || (a.order - b.order);
    });
    var perMachine = {};
    levels.forEach(function (m) { perMachine[m.no] = []; });
    var failed = [];
    list.forEach(function (b) {
      // 큰 물량부터 큰 세탁기에: 들어갈 수 있고 자리가 남은 세탁기 중 가장 큰 쪽(큰 세탁기에 1포대만 넣는 낭비를 막는다).
      // 큰 물량을 먼저 놓으므로(내림차순) 뒤에 오는 작은 물량은 남은 어느 세탁기에나 들어가 배정이 막히지 않는다.
      var pick = -1;
      for (var j = 0; j < levels.length; j++) {
        if (levels[j].cap < b.size || used[j] >= slots[j]) continue;
        if (pick === -1 || levels[j].cap > levels[pick].cap || (levels[j].cap === levels[pick].cap && used[j] < used[pick])) pick = j;
      }
      if (pick === -1) { failed.push(b); return; }
      used[pick]++;
      perMachine[levels[pick].no].push(b);
    });
    return { perMachine: perMachine, failed: failed };
  }

  function sortCycleItems(items) {
    return items.slice().sort(function (a, b) {
      return (a.course < b.course ? -1 : a.course > b.course ? 1 : 0) || (b.qty - a.qty) || (numKey(a.no) - numKey(b.no));
    });
  }
  function itemOf(c) {
    return {
      no: c.no, name: c.name, course: c.course, industry: c.industry, group: c.group,
      unclassified: c.unclassified, solo: !!c.solo, qty: c.qty, source: c.source,
      extra: !!c.extra, split: !!c.split
    };
  }

  // ---------------------------------------------------------------- 배정표 생성
  /**
   * companies: buildCompanies().companies
   * opts.locked: { "기계번호-회차": { items:[{no,name,course,industry,group,unclassified,solo,qty,source}] } }  완료(고정)된 회차
   * opts.config: DEFAULT_CONFIG 덮어쓰기
   */
  function buildPlan(companies, opts) {
    opts = opts || {};
    var cfg = mergeConfig(opts.config);
    var rules = makeRules(cfg);
    var levels = machineLevels(cfg);
    var maxCap = levels[0].cap;
    var doneKeys = opts.locked || {};   // 작업자가 완료 체크한 회차(고정)
    var locked = {};                    // 고정 대상 = 기준안(frozen, 데이터가 안 바뀌었을 때 그대로 유지) + 완료 회차
    Object.keys(opts.frozen || {}).forEach(function (k) { locked[k] = opts.frozen[k]; });
    Object.keys(doneKeys).forEach(function (k) { locked[k] = doneKeys[k]; });
    var warnings = [];

    var pending = [], zero = [], todo = [];
    companies.forEach(function (c) {
      if (c.source === 'unknown' || c.qty === null || c.qty === undefined) pending.push(c);
      else if (c.qty === 0) zero.push(c);
      else todo.push(c);
    });

    // 완료(고정) 회차가 가져간 수량을 먼저 뺀다
    var remaining = {}, byNo = {};
    todo.forEach(function (c) { remaining[c.no] = c.qty; byNo[c.no] = c; });
    var lockedCount = {};
    var lockedCycles = {};
    Object.keys(locked).forEach(function (key) {
      var m = /^(\d+)-(\d+)$/.exec(key);
      if (!m || !locked[key] || !locked[key].items) return;
      var mno = Number(m[1]), k = Number(m[2]);
      if (!levels.some(function (l) { return l.no === mno; })) return;
      lockedCount[mno] = (lockedCount[mno] || 0) + 1;
      (lockedCycles[mno] = lockedCycles[mno] || {})[k] = locked[key];
      locked[key].items.forEach(function (it) {
        if (remaining[it.no] === undefined) {
          var known = companies.filter(function (c) { return c.no === it.no; })[0];
          if (!known) warnings.push('완료한 ' + mno + '번 세탁기 ' + k + '차의 ' + it.no + ' ' + it.name + ' 은(는) 지금 목록에 없습니다.');
          return;
        }
        remaining[it.no] -= it.qty;
      });
    });
    Object.keys(remaining).forEach(function (no) {
      if (remaining[no] < 0) warnings.push(no + ' ' + byNo[no].name + ': 완료한 회차에 이미 ' + (byNo[no].qty - remaining[no]) + '포대가 들어가 현재 수량(' + byNo[no].qty + ')보다 많습니다.');
    });

    // 남은 물량을 회차 후보(chunk)로. 한 거래처가 세탁기 최대 용량을 넘으면 어쩔 수 없이 나눈다.
    var chunks = [];
    todo.slice().sort(function (a, b) { return numKey(a.no) - numKey(b.no); }).forEach(function (c) {
      var rem = remaining[c.no];
      if (rem <= 0) return;
      var wasLocked = rem !== c.qty;
      var parts = [];
      while (rem > maxCap) { parts.push(maxCap); rem -= maxCap; }
      parts.push(rem);
      parts.forEach(function (q) {
        var copy = {};
        Object.keys(c).forEach(function (k) { copy[k] = c[k]; });
        copy.qty = q; copy.split = parts.length > 1; copy.extra = wasLocked;
        chunks.push(copy);
      });
    });

    var ctx = { cfg: cfg, rules: rules, levels: levels, maxCap: maxCap, lockedCount: lockedCount };
    var seedBase = hashStr(chunks.map(function (c) { return c.no + ':' + c.qty + ':' + c.group; }).join('|'));

    var chosen = null, budget = 0;
    if (chunks.length) {
      var cycleSet = {}, cands = [];
      cfg.machines.forEach(function (m) {
        for (var k = 1; k <= chunks.length + 1 + (lockedCount[m.no] || 0); k++) cycleSet[k * m.cycleMin] = 1;
      });
      cands = Object.keys(cycleSet).map(Number).sort(function (a, b) { return a - b; });
      var totalBags = chunks.reduce(function (s, c) { return s + c.qty; }, 0);
      var perRound = levels.reduce(function (s, m) { return s + m.cap; }, 0);
      var minRounds = Math.max(1, Math.ceil(totalBags / perRound));
      var minMin = minRounds * Math.min.apply(null, cfg.machines.map(function (m) { return m.cycleMin; }));
      // 한 회차에는 미분류(또는 단독) 거래처가 하나만 들어가므로, 회차 수는 그 거래처 수 이상이다.
      var needSlots = chunks.filter(function (c) { return c.unclassified || c.solo; }).length;
      for (var ci = 0; ci < cands.length && !chosen; ci++) {
        if (cands[ci] < minMin) continue;
        var haveSlots = slotsFor(levels, cands[ci], lockedCount).reduce(function (s, x) { return s + x; }, 0);
        if (haveSlots < needSlots) continue;
        var bestHere = null;
        for (var r = 0; r < cfg.searchRestarts; r++) {
          var res = anneal(chunks, ctx, cands[ci], seedBase + r * 7919);
          if (res.viol === 0 && (!bestHere || res.E < bestHere.E)) bestHere = res;
        }
        if (bestHere) {
          var placed = placeBatches(bestHere.batches, chunks, ctx, cands[ci]);
          if (!placed.failed.length) { chosen = { batches: bestHere, placed: placed }; budget = cands[ci]; }
        }
      }
    } else {
      chosen = { batches: { batches: [] }, placed: { perMachine: {}, failed: [] } };
    }

    var machines = cfg.machines.map(function (m) {
      return { no: m.no, cap: m.cap, cycleMin: m.cycleMin, cycles: [] };
    });
    var errors = [];
    if (!chosen) {
      errors.push('세탁기 배정을 찾지 못했습니다(내부 오류). 수량·그룹 데이터를 확인하세요.');
    } else {
      machines.forEach(function (mach) {
        var newOnes = (chosen.placed.perMachine[mach.no] || []).slice();
        newOnes.sort(function (a, b) {
          return (groupRank(a.group) - groupRank(b.group)) || (b.size - a.size) || (numKey(a.members[0].no) - numKey(b.members[0].no));
        });
        var lockedHere = lockedCycles[mach.no] || {};
        var nextFree = 1, ni = 0, total = newOnes.length + Object.keys(lockedHere).length;
        var maxK = 0;
        Object.keys(lockedHere).forEach(function (k) { maxK = Math.max(maxK, Number(k)); });
        maxK = Math.max(maxK, total);
        for (var k = 1; k <= maxK; k++) {
          if (lockedHere[k]) {
            var li = lockedHere[k].items;
            mach.cycles.push(describeCycle(li.map(function (x) { return x; }), mach, k, !!doneKeys[mach.no + '-' + k], rules));
          } else if (ni < newOnes.length) {
            var b = newOnes[ni++];
            mach.cycles.push(describeCycle(sortCycleItems(b.members).map(itemOf), mach, k, false, rules, b.members));
          }
        }
      });
    }

    var plan = {
      machines: machines,
      pending: pending.map(pubCompany),
      zero: zero.map(pubCompany),
      warnings: warnings,
      errors: errors,
      budgetMin: budget,
      config: cfg
    };
    plan.stats = summarize(plan, companies);
    var v = validatePlan(plan, companies, { locked: doneKeys });
    plan.errors = plan.errors.concat(v);
    plan.valid = plan.errors.length === 0;
    return plan;
  }

  function pubCompany(c) {
    return { no: c.no, name: c.name, industry: c.industry, course: c.course, group: c.group, qty: c.qty, source: c.source, basis: c.basis };
  }

  function describeCycle(items, mach, k, isLocked, rules, memberObjs) {
    var qty = items.reduce(function (s, x) { return s + x.qty; }, 0);
    var classified = items.filter(function (x) { return !x.unclassified; });
    var group = classified.length ? classified[0].group : 'U';
    var inds = [];
    items.forEach(function (x) { if (x.industry && inds.indexOf(x.industry) === -1) inds.push(x.industry); });
    var courses = [];
    items.forEach(function (x) { if (courses.indexOf(x.course) === -1) courses.push(x.course); });
    courses.sort();
    // 저장된 내용(고정 회차)만으로도 같은 기준으로 다시 판정
    var ms = memberObjs || items.map(function (x) { return { no: x.no, unclassified: x.unclassified, family: familyOf(x.industry, DEFAULT_CONFIG) }; });
    var cond = cycleIsConditional(ms, rules);
    return {
      key: mach.no + '-' + k, machineNo: mach.no, k: k, items: items, qty: qty, cap: mach.cap, group: group,
      groupLabel: groupLabel(group), groupClass: groupClass(group), industries: inds, courses: courses,
      conditional: cond.conditional, exception: cond.exception, locked: !!isLocked
    };
  }

  function summarize(plan, companies) {
    var bags = 0, ai = 0, drv = 0, cycles = 0, cond = 0, maxMin = 0, assigned = {};
    plan.machines.forEach(function (m) {
      var last = 0;
      m.cycles.forEach(function (c) {
        cycles++; if (c.conditional) cond++;
        last = Math.max(last, c.k);
        c.items.forEach(function (it) {
          bags += it.qty; assigned[it.no] = true;
          if (it.source === 'driver') drv += it.qty; else ai += it.qty;
        });
      });
      maxMin = Math.max(maxMin, last * m.cycleMin);
    });
    var allBags = 0;
    companies.forEach(function (c) { if (c.qty) allBags += c.qty; });
    return {
      companies: companies.length, assignedCompanies: Object.keys(assigned).length, bags: bags, totalBags: allBags,
      aiBags: ai, driverBags: drv, cycles: cycles, conditionalCycles: cond, makespanMin: maxMin,
      pending: plan.pending.length, zero: plan.zero.length
    };
  }

  // ---------------------------------------------------------------- 검증
  /**
   * 어긋나는 점이 있으면 문장 배열로 돌려준다(빈 배열이면 통과).
   * 검증에 실패한 배정표는 화면이 정상 결과처럼 보여주면 안 된다.
   */
  function validatePlan(plan, companies, opts) {
    var errs = [];
    var cfg = plan.config || DEFAULT_CONFIG;
    var rules = makeRules(cfg);
    var lockedKeys = (opts && opts.locked) || {};
    var qtyAssigned = {}, cyclesOf = {}, seenInCycle;
    var maxCap = Math.max.apply(null, cfg.machines.map(function (m) { return m.cap; }));

    plan.machines.forEach(function (m) {
      var ks = {};
      m.cycles.forEach(function (c) {
        if (ks[c.k]) errs.push(m.no + '번 세탁기 ' + c.k + '차가 두 번 만들어졌습니다.');
        ks[c.k] = true;
        if (c.qty > m.cap) errs.push(m.no + '번 세탁기 ' + c.k + '차: ' + c.qty + '포대가 최대 ' + m.cap + '포대 용량을 초과합니다.');
        if (c.items.length > cfg.maxCompanies) errs.push(m.no + '번 세탁기 ' + c.k + '차: 거래처가 ' + c.items.length + '곳입니다(최대 ' + cfg.maxCompanies + ').');
        seenInCycle = {};
        c.items.forEach(function (it) {
          if (seenInCycle[it.no]) errs.push(m.no + '번 ' + c.k + '차에 거래처 ' + it.no + ' 이(가) 두 번 들어 있습니다.');
          seenInCycle[it.no] = true;
          qtyAssigned[it.no] = (qtyAssigned[it.no] || 0) + it.qty;
          (cyclesOf[it.no] = cyclesOf[it.no] || []).push(m.no + '-' + c.k);
        });
        if (!c.locked) {
          var ms = c.items.map(function (it) {
            return { no: it.no, qty: it.qty, group: it.group, unclassified: it.unclassified, solo: it.solo, industry: it.industry, family: familyOf(it.industry, cfg) };
          });
          cycleViolations(ms, cfg, rules, m.cap).forEach(function (v) { errs.push(m.no + '번 세탁기 ' + c.k + '차: ' + v); });
        }
      });
    });

    var pendNo = {}, zeroNo = {};
    plan.pending.forEach(function (c) { pendNo[c.no] = true; });
    plan.zero.forEach(function (c) { zeroNo[c.no] = true; });
    var sumAssignable = 0, sumAssigned = 0;
    companies.forEach(function (c) {
      var got = qtyAssigned[c.no] || 0;
      var isPending = c.source === 'unknown' || c.qty === null || c.qty === undefined;
      if (isPending) {
        if (got) errs.push('수량 미확인 거래처 ' + c.no + ' 이(가) 배정되었습니다.');
        if (!pendNo[c.no]) errs.push('수량 미확인 거래처 ' + c.no + ' 이(가) 확인 필요 목록에서 빠졌습니다.');
        return;
      }
      if (c.qty === 0) {
        if (got) errs.push('수량 0인 거래처 ' + c.no + ' 이(가) 배정되었습니다.');
        if (!zeroNo[c.no]) errs.push('수량 0 거래처 ' + c.no + ' 이(가) 목록에서 빠졌습니다.');
        return;
      }
      sumAssignable += c.qty; sumAssigned += got;
      if (!got) errs.push('거래처 ' + c.no + ' ' + c.name + ' 이(가) 배정에서 빠졌습니다(누락).');
      else if (got !== c.qty && !hasLockedMismatch(c.no, cyclesOf, lockedKeys)) errs.push('거래처 ' + c.no + ' ' + c.name + ': 배정 ' + got + '포대 ≠ 수량 ' + c.qty + '포대.');
      if (cyclesOf[c.no] && cyclesOf[c.no].length > 1 && c.qty <= maxCap && !hasLockedMismatch(c.no, cyclesOf, lockedKeys)) {
        errs.push('거래처 ' + c.no + ' ' + c.name + ' 이(가) 여러 회차(' + cyclesOf[c.no].join(', ') + ')에 나뉘어 배정되었습니다.');
      }
    });
    var known = {};
    companies.forEach(function (c) { known[c.no] = true; });
    Object.keys(qtyAssigned).forEach(function (no) {
      if (!known[no] && !hasLockedMismatch(no, cyclesOf, lockedKeys)) errs.push('목록에 없는 거래처 ' + no + ' 이(가) 배정되어 있습니다.');
    });
    if (sumAssigned !== sumAssignable && !Object.keys(lockedKeys).length) errs.push('배정 합계(' + sumAssigned + ')와 수량 합계(' + sumAssignable + ')가 다릅니다.');
    return errs;
  }
  // 완료(고정) 회차가 끼어 있으면 그 거래처의 수량은 완료 당시 값이라 현재 값과 다를 수 있다
  function hasLockedMismatch(no, cyclesOf, lockedKeys) {
    return (cyclesOf[no] || []).some(function (k) { return !!lockedKeys[k]; });
  }

  return {
    DEFAULT_CONFIG: DEFAULT_CONFIG,
    mergeConfig: mergeConfig,
    normalizeGroup: normalizeGroup,
    groupLabel: groupLabel,
    groupClass: groupClass,
    groupRank: groupRank,
    estimateQty: estimateQty,
    medianOf: medianOf,
    buildCompanies: buildCompanies,
    buildPlan: buildPlan,
    validatePlan: validatePlan,
    cycleViolations: cycleViolations,
    makeRules: makeRules,
    familyOf: familyOf,
    numKey: numKey
  };
});
