// 실행: node tests/wash-core.test.js   (외부 패키지 없음)
'use strict';
var assert = require('assert');
var W = require('../wash-core.js');

var passed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log('  ok  ' + name); }
  catch (e) { console.log('FAIL  ' + name + '\n      ' + e.message); process.exitCode = 1; }
}

// 거래처 만들기 도우미: [번호, 이름, 업종, 코스, 합세탁코드, 수량]
function comp(no, name, industry, course, group, qty, extra) {
  var c = {
    no: String(no), name: name, industry: industry, family: W.familyOf(industry, W.DEFAULT_CONFIG), course: course,
    group: W.normalizeGroup(group), unclassified: W.normalizeGroup(group) === 'U', solo: false,
    qty: qty, source: qty === null ? 'unknown' : 'ai', basis: { values: [], dates: [], median: qty }
  };
  if (extra) Object.keys(extra).forEach(function (k) { c[k] = extra[k]; });
  return c;
}
function allCycles(plan) { var l = []; plan.machines.forEach(function (m) { m.cycles.forEach(function (c) { l.push({ m: m, c: c }); }); }); return l; }

console.log('수량 산정');
test('기사 등록(실수량)이 AI 예상보다 우선, 0도 실수량', function () {
  var e = W.estimateQty({ actual: 0, history: [{ date: 'a', qty: 3 }] });
  assert.strictEqual(e.qty, 0); assert.strictEqual(e.source, 'driver');
  e = W.estimateQty({ actual: 4, history: [{ date: 'a', qty: 1 }] });
  assert.strictEqual(e.qty, 4); assert.strictEqual(e.source, 'driver');
});
test('이력이 없으면 확인 필요(임의 1포대 없음), 빈 칸과 0은 구분', function () {
  var e = W.estimateQty({ actual: null, history: [{ date: 'a', qty: null }, { date: 'b', qty: null }] });
  assert.strictEqual(e.source, 'unknown'); assert.strictEqual(e.qty, null);
  e = W.estimateQty({ actual: null, history: [{ date: 'a', qty: 0 }, { date: 'b', qty: 0 }] });
  assert.strictEqual(e.source, 'ai'); assert.strictEqual(e.qty, 0);
  assert.strictEqual(W.estimateQty(undefined).source, 'unknown');
});
test('최근 4회 기록(빈 칸 건너뜀)의 중앙값, 이상치에 강함', function () {
  var h = [{ date: '4', qty: 2 }, { date: '3', qty: null }, { date: '2', qty: 2 }, { date: '1', qty: 9 }, { date: '0', qty: 2 }, { date: 'old', qty: 7 }];
  var e = W.estimateQty({ actual: null, history: h });
  assert.deepStrictEqual(e.basis.values, [2, 2, 9, 2]); // 5번째 기록(7)은 4회 밖
  assert.strictEqual(e.qty, 2);
  assert.strictEqual(W.estimateQty({ history: [{ qty: 1 }, { qty: 1 }, { qty: 2 }, { qty: 2 }] }).qty, 1); // 짝수 개는 작은 쪽
});

console.log('배정 안전 규칙');
var base = [
  comp(27, '심', '헬스장', 'B', 'A', 2), comp(6, '올데이', '헬스장', 'B', 'A', 2), comp(24, '헬스코인', '헬스장', 'B', 'A', 1),
  comp(20, '박승철 범계', '미용실', 'B', 'B2', 2), comp(7, '리안 1호', '미용실', 'B', 'X', 2),
  comp(113, '휴이엠', '미용실', 'B', 'B', 1), comp(92, '글로스', '미용실', 'Z', 'B', 1),
  comp(120, '더부즈', '음식점', 'Z', 'X', 2), comp(15, '두빗 1호', '미용실', 'B', 'X', 1), comp(188, '라움', '미용실', 'Z', 'X', 1),
  comp(211, '버블테라피', '피부관리', 'B', 'X', 1), comp(142, '국군수송', '헬스장', 'D', 'B', 1), comp(169, '비주얼', '미용실', 'D', 'X', 1),
  comp(125, '위례나무', '병원', 'C', 'X', 2), comp(155, '유앤미', '병원', 'D', 'B', 1)
];
var plan = W.buildPlan(base);
test('검증 통과 + 모든 거래처가 정확히 한 번 배정', function () {
  assert.deepStrictEqual(plan.errors, []);
  var got = {};
  allCycles(plan).forEach(function (x) { x.c.items.forEach(function (it) { got[it.no] = (got[it.no] || 0) + it.qty; }); });
  base.forEach(function (c) { assert.strictEqual(got[c.no], c.qty, c.no); });
});
test('용량 초과·회차당 3곳 초과 없음', function () {
  allCycles(plan).forEach(function (x) {
    assert.ok(x.c.qty <= x.m.cap, x.m.no + '번 ' + x.c.qty + '>' + x.m.cap);
    assert.ok(x.c.items.length <= 3);
  });
});
test('미분류 2곳 이상 합세탁 없음, 서로 다른 그룹·업종 혼합 없음(예외 조합 제외)', function () {
  allCycles(plan).forEach(function (x) {
    assert.ok(x.c.items.filter(function (i) { return i.unclassified; }).length <= 1);
    var groups = {};
    x.c.items.forEach(function (i) { if (!i.unclassified) groups[i.group] = 1; });
    assert.ok(Object.keys(groups).length <= 1);
    var fams = {};
    x.c.items.forEach(function (i) { fams[W.familyOf(i.industry, W.DEFAULT_CONFIG)] = 1; });
    if (Object.keys(fams).length > 1) assert.ok(x.c.exception, '예외 조합이 아닌데 업종 혼합: ' + x.c.key);
  });
});
test('나-2 그룹은 나 그룹과 섞이지 않는다', function () {
  allCycles(plan).forEach(function (x) {
    var g = {};
    x.c.items.forEach(function (i) { if (!i.unclassified) g[i.group] = 1; });
    assert.ok(!(g.B && g.B2));
  });
});
test('재현성: 같은 입력은 같은 배정표', function () {
  var again = W.buildPlan(base);
  assert.strictEqual(JSON.stringify(again.machines), JSON.stringify(plan.machines));
});
test('수량 확인 필요·예상 0은 배정하지 않고 따로 표시', function () {
  var list = base.concat([comp(216, '차홍룸', '미용실', 'Z', 'B', null), comp(104, '다짐', '헬스장', 'B', 'A', 0)]);
  var p = W.buildPlan(list);
  assert.deepStrictEqual(p.errors, []);
  assert.deepStrictEqual(p.pending.map(function (c) { return c.no; }), ['216']);
  assert.deepStrictEqual(p.zero.map(function (c) { return c.no; }), ['104']);
  allCycles(p).forEach(function (x) { x.c.items.forEach(function (i) { assert.ok(i.no !== '216' && i.no !== '104'); }); });
});
test('승인 예외 조합(142+169)은 같이 돌릴 수 있고 조건부로 표시된다', function () {
  var p = W.buildPlan([comp(142, '국군수송', '헬스장', 'D', 'B', 1), comp(169, '비주얼', '미용실', 'D', 'X', 1)]);
  assert.deepStrictEqual(p.errors, []);
  var together = allCycles(p).filter(function (x) { return x.c.items.length === 2; });
  if (together.length) assert.ok(together[0].c.conditional && together[0].c.exception);
});
test('예외가 없으면 업종이 다른 거래처는 같은 회차에 안 들어간다', function () {
  var p = W.buildPlan([comp(1, 'a', '헬스장', 'B', 'A', 1), comp(2, 'b', '미용실', 'B', 'X', 1)]);
  allCycles(p).forEach(function (x) { assert.strictEqual(x.c.items.length, 1); });
});
test('단독(solo) 거래처는 혼자 돌린다', function () {
  var p = W.buildPlan([comp(1, 'a', '헬스장', 'B', 'A', 1), comp(2, 'b', '헬스장', 'B', 'X', 1, { solo: true }), comp(3, 'c', '헬스장', 'B', 'A', 1)]);
  assert.deepStrictEqual(p.errors, []);
  allCycles(p).forEach(function (x) { if (x.c.items.some(function (i) { return i.no === '2'; })) assert.strictEqual(x.c.items.length, 1); });
});
test('세탁기 최대 용량(5)보다 큰 거래처는 나눠서 배정하고 합계는 일치', function () {
  var p = W.buildPlan([comp(9, '대형', '헬스장', 'B', 'A', 7)]);
  assert.deepStrictEqual(p.errors, []);
  var sum = 0; allCycles(p).forEach(function (x) { x.c.items.forEach(function (i) { sum += i.qty; }); });
  assert.strictEqual(sum, 7);
});
test('종료시간이 이론 하한과 같다(회차 수 ≥ 미분류 수)', function () {
  var unclassified = base.filter(function (c) { return c.unclassified; }).length;
  var cycles = allCycles(plan).length;
  assert.ok(cycles >= unclassified);
  assert.strictEqual(plan.stats.makespanMin, Math.max.apply(null, plan.machines.map(function (m) { return m.cycles.length * 20; })));
});

console.log('완료 회차 고정과 재배정');
test('완료한 회차는 수량이 바뀌어도 그대로, 나머지만 다시 계산', function () {
  var first = allCycles(plan)[0];
  var locked = {}; locked[first.c.key] = { items: first.c.items };
  var list = base.map(function (c) { return comp(c.no, c.name, c.industry, c.course, c.group === 'U' ? 'X' : c.group, c.qty); });
  // 완료 회차에 없는 거래처 하나의 수량을 바꾼다
  var lockedNos = first.c.items.map(function (i) { return i.no; });
  var target = list.filter(function (c) { return lockedNos.indexOf(c.no) === -1 && c.qty > 0; })[0];
  target.qty += 1;
  var p = W.buildPlan(list, { locked: locked });
  assert.deepStrictEqual(p.errors, []);
  var keep = p.machines.filter(function (m) { return m.no === first.m.no; })[0].cycles.filter(function (c) { return c.k === first.c.k; })[0];
  assert.ok(keep.locked);
  assert.deepStrictEqual(keep.items.map(function (i) { return i.no + ':' + i.qty; }), first.c.items.map(function (i) { return i.no + ':' + i.qty; }));
  var sum = 0; allCycles(p).forEach(function (x) { x.c.items.forEach(function (i) { sum += i.qty; }); });
  var want = list.reduce(function (s, c) { return s + (c.qty || 0); }, 0);
  assert.strictEqual(sum, want);
});

console.log('검증기가 잘못된 배정표를 잡는다');
function clone(o) { return JSON.parse(JSON.stringify(o)); }
test('거래처 누락', function () {
  var p = clone(plan); p.machines[0].cycles[0].items.shift(); p.machines[0].cycles[0].qty = 0;
  assert.ok(W.validatePlan(p, base, {}).some(function (e) { return /빠졌|배정/.test(e); }));
});
test('거래처 중복 배정', function () {
  var p = clone(plan); var it = clone(p.machines[0].cycles[0].items[0]);
  p.machines[1].cycles[0].items.push(it);
  assert.ok(W.validatePlan(p, base, {}).length > 0);
});
test('용량 초과', function () {
  var p = clone(plan); p.machines[3].cycles[0].qty = 9;
  assert.ok(W.validatePlan(p, base, {}).some(function (e) { return /초과/.test(e); }));
});
test('미분류 2곳 합세탁', function () {
  var p = clone(plan);
  var un = allCycles(p).filter(function (x) { return x.c.group === 'U'; });
  var a = un[0].c, b = un[1].c;
  a.items = a.items.concat(b.items); a.qty += b.qty; b.items = [];
  assert.ok(W.validatePlan(p, base, {}).some(function (e) { return /미분류/.test(e); }));
});
test('수량 합계 불일치', function () {
  var p = clone(plan); p.machines[0].cycles[0].items[0].qty += 1;
  assert.ok(W.validatePlan(p, base, {}).some(function (e) { return /≠|합계/.test(e); }));
});

console.log('\n' + passed + '개 통과' + (process.exitCode ? ' (실패 있음)' : ''));
