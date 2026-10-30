/**
 * KK22 - Program Engine
 * Kalkulasi harga program/diskon untuk order.
 *
 * Version: 1.0 (F9.7B-1)
 *
 * RULES terimplementasi (5 simple):
 *  - item_fixed_price      {harga}
 *  - item_discount_pct     {persen}
 *  - item_discount_fixed   {rupiah}
 *  - order_discount_pct    {persen}
 *  - order_discount_fixed  {rupiah}
 *
 * RULES TODO (F9.7B-2..5):
 *  - item_tiered_qty, item_bxgy, item_bundle
 *  - order_tiered_total, order_tiered_qty
 *  - order_min_qty, order_min_value, order_free_shipping, gift_with_purchase
 *  - kalkulasi mode: best_for_customer
 *  - basis_harga modes: normal/paket/custom
 */

var PROGRAM_ENGINE_VERSION = '1.0';

var IMPLEMENTED_ITEM_RULES = ['item_fixed_price', 'item_discount_pct', 'item_discount_fixed'];
var IMPLEMENTED_ORDER_RULES = ['order_discount_pct', 'order_discount_fixed'];

/**
 * Main entry: hitung harga order dengan program.
 *
 * @param {Array} items - [{nama_produk, qty, harga_awal, subtotal_awal}]
 * @param {Array} programs - [{header, rules}] - program sudah di-resolve
 * @param {Object} produkMap - {nama_lowercase: {nama_produk, harga, satuan}}
 * @returns {Object} hasil kalkulasi
 */
function calculateProgramsForOrder_(items, programs, produkMap) {
  if (!Array.isArray(items) || items.length === 0) {
    return { ok: false, error: 'Items kosong.' };
  }

  if (!Array.isArray(programs) || programs.length === 0) {
    return _programPassthrough_(items);
  }

  // Sort program by priority DESC
  var sorted = programs.slice().sort(function(a, b) {
    var pa = (a.header && a.header.priority) || 10;
    var pb = (b.header && b.header.priority) || 10;
    return pb - pa;
  });

  // Cek mode kombinasi + kalkulasi
  var modeInfo = _getEffectiveMode_(sorted);
  if (!modeInfo.ok) return { ok: false, error: modeInfo.error };

  // Deep clone items
  var itemsFinal = [];
  for (var i = 0; i < items.length; i++) {
    var src = items[i];
    var harga = Number(src.harga_awal) || 0;
    var qty = Number(src.qty) || 0;
    var subAwal = Number(src.subtotal_awal) || (qty * harga);
    itemsFinal.push({
      nama_produk: String(src.nama_produk || ''),
      qty: qty,
      harga_awal: harga,
      subtotal_awal: subAwal,
      harga_akhir: harga,
      subtotal_akhir: subAwal,
      diskon_item: 0,
      bonus_qty: 0
    });
  }

  var snapshotLangkah = [];
  var programsApplied = [];
  var warnings = [];

  // PHASE 1: item-level rules
  for (var p = 0; p < sorted.length; p++) {
    var prog = sorted[p];
    var ph = prog.header || {};
    var pr = Array.isArray(prog.rules) ? prog.rules : [];

    for (var r = 0; r < pr.length; r++) {
      var rule = pr[r];
      var rt = String(rule.rule_type || '').toLowerCase();
      if (IMPLEMENTED_ITEM_RULES.indexOf(rt) === -1) continue;

      var param = _programParseParam_(rule.parameter);
      if (!param) {
        warnings.push('Program ' + (ph.id_program || '?') + ': parameter rule ' + rt + ' tidak valid.');
        continue;
      }

      var applied = 0;
      for (var it = 0; it < itemsFinal.length; it++) {
        var res = _applyItemRule_(rt, param, itemsFinal[it], produkMap);
        if (!res.ok) { warnings.push(res.error); continue; }
        if (res.applied) { itemsFinal[it] = res.item; applied++; }
      }
      if (applied > 0) {
        snapshotLangkah.push({ phase: 'item', program: ph.id_program, rule: rt, applied_to: applied + ' item' });
      }
    }
  }

  // Recalc setelah item rules
  var subtotalSetelahItem = 0;
  for (var s = 0; s < itemsFinal.length; s++) {
    itemsFinal[s].subtotal_akhir = itemsFinal[s].qty * itemsFinal[s].harga_akhir;
    itemsFinal[s].diskon_item = Math.max(0, itemsFinal[s].subtotal_awal - itemsFinal[s].subtotal_akhir);
    subtotalSetelahItem += itemsFinal[s].subtotal_akhir;
  }

  // PHASE 2: order-level rules
  var subtotalForOrder = subtotalSetelahItem;
  var totalDiskonOrder = 0;

  for (var p2 = 0; p2 < sorted.length; p2++) {
    var prog2 = sorted[p2];
    var ph2 = prog2.header || {};
    var pr2 = Array.isArray(prog2.rules) ? prog2.rules : [];
    var progDiskon = 0;
    var progBreakdown = [];

    for (var r2 = 0; r2 < pr2.length; r2++) {
      var rule2 = pr2[r2];
      var rt2 = String(rule2.rule_type || '').toLowerCase();
      if (IMPLEMENTED_ORDER_RULES.indexOf(rt2) === -1) continue;

      var param2 = _programParseParam_(rule2.parameter);
      if (!param2) {
        warnings.push('Program ' + (ph2.id_program || '?') + ': parameter rule ' + rt2 + ' tidak valid.');
        continue;
      }

      var basis = (modeInfo.kalkulasi === 'parallel') ? subtotalSetelahItem : subtotalForOrder;
      var res2 = _applyOrderRule_(rt2, param2, basis);
      if (!res2.ok) { warnings.push(res2.error); continue; }

      if (res2.diskon > 0) {
        progDiskon += res2.diskon;
        totalDiskonOrder += res2.diskon;
        progBreakdown.push({ rule: rt2, basis: basis, diskon: res2.diskon });
        if (modeInfo.kalkulasi === 'sequential') {
          subtotalForOrder -= res2.diskon;
          if (subtotalForOrder < 0) subtotalForOrder = 0;
        }
      }
    }

    if (progDiskon > 0) {
      programsApplied.push({
        id_program: ph2.id_program,
        nama_program: ph2.nama_program,
        diskon_order: progDiskon,
        breakdown: progBreakdown
      });
    }
  }

  // Cap maximum per program
  for (var pa = 0; pa < programsApplied.length; pa++) {
    var progA = programsApplied[pa];
    var matched = sorted.filter(function(x) { return (x.header||{}).id_program === progA.id_program; });
    if (matched.length > 0) {
      var maxD = Number((matched[0].header||{}).max_diskon_per_order) || 0;
      if (maxD > 0 && progA.diskon_order > maxD) {
        totalDiskonOrder -= (progA.diskon_order - maxD);
        progA.diskon_order = maxD;
      }
    }
  }

  var totalAkhir = subtotalSetelahItem - totalDiskonOrder;
  if (totalAkhir < 0) totalAkhir = 0;

  var subtotalAwal = 0;
  for (var sa = 0; sa < items.length; sa++) {
    var q = Number(items[sa].qty) || 0;
    var h = Number(items[sa].harga_awal) || 0;
    subtotalAwal += Number(items[sa].subtotal_awal) || (q * h);
  }
  var totalDiskonItem = subtotalAwal - subtotalSetelahItem;

  var snapshot = {
    engine_version: PROGRAM_ENGINE_VERSION,
    timestamp: nowJakarta_(),
    kalkulasi_mode: modeInfo.kalkulasi,
    kombinasi_mode: modeInfo.kombinasi,
    subtotal_awal: subtotalAwal,
    subtotal_setelah_item: subtotalSetelahItem,
    langkah: snapshotLangkah,
    programs_applied: programsApplied,
    total_diskon_item: totalDiskonItem,
    total_diskon_order: totalDiskonOrder,
    total_akhir: totalAkhir,
    warnings: warnings
  };

  return {
    ok: true,
    items_final: itemsFinal,
    subtotal_awal: subtotalAwal,
    total_diskon_item: totalDiskonItem,
    subtotal_setelah_item: subtotalSetelahItem,
    total_diskon_order: totalDiskonOrder,
    total_akhir: totalAkhir,
    programs_applied: programsApplied,
    warnings: warnings,
    snapshot: snapshot
  };
}

function _programPassthrough_(items) {
  var itemsFinal = [];
  var subtotalAwal = 0;
  for (var i = 0; i < items.length; i++) {
    var src = items[i];
    var h = Number(src.harga_awal) || 0;
    var q = Number(src.qty) || 0;
    var sub = Number(src.subtotal_awal) || (q * h);
    subtotalAwal += sub;
    itemsFinal.push({
      nama_produk: String(src.nama_produk || ''),
      qty: q,
      harga_awal: h,
      subtotal_awal: sub,
      harga_akhir: h,
      subtotal_akhir: sub,
      diskon_item: 0,
      bonus_qty: 0
    });
  }
  return {
    ok: true,
    items_final: itemsFinal,
    subtotal_awal: subtotalAwal,
    total_diskon_item: 0,
    subtotal_setelah_item: subtotalAwal,
    total_diskon_order: 0,
    total_akhir: subtotalAwal,
    programs_applied: [],
    warnings: [],
    snapshot: { engine_version: PROGRAM_ENGINE_VERSION, no_program: true }
  };
}

function _getEffectiveMode_(programs) {
  var exclusiveCount = 0;
  var kombinasiMode = 'stackable';
  var kalkulasiMode = 'sequential';
  for (var i = 0; i < programs.length; i++) {
    var h = programs[i].header || {};
    var km = String(h.kombinasi_mode || 'exclusive').toLowerCase();
    if (km === 'exclusive') exclusiveCount++;
    if (i === 0) {
      kombinasiMode = km;
      kalkulasiMode = String(h.kalkulasi_mode || 'sequential').toLowerCase();
    }
  }
  if (exclusiveCount > 1) {
    return { ok: false, error: 'Tidak boleh apply lebih dari 1 program exclusive sekaligus.' };
  }
  if (['sequential','parallel','best_for_customer'].indexOf(kalkulasiMode) === -1) {
    kalkulasiMode = 'sequential';
  }
  return { ok: true, kombinasi: kombinasiMode, kalkulasi: kalkulasiMode };
}

function _programParseParam_(param) {
  try {
    if (typeof param === 'string') return JSON.parse(param);
    if (typeof param === 'object' && param !== null) return param;
    return null;
  } catch (e) { return null; }
}

function _applyItemRule_(ruleType, param, item, produkMap) {
  var qty = Number(item.qty) || 0;
  var hargaAwal = Number(item.harga_awal) || 0;
  var newHarga = hargaAwal;

  if (ruleType === 'item_fixed_price') {
    var h = Number(param.harga);
    if (!h || h <= 0) return { ok: false, error: 'item_fixed_price: harga harus > 0' };
    newHarga = h;
  } else if (ruleType === 'item_discount_pct') {
    var persen = Number(param.persen);
    if (!persen || persen <= 0 || persen > 100) return { ok: false, error: 'item_discount_pct: persen harus 1-100' };
    newHarga = Math.round(hargaAwal * (1 - persen / 100));
  } else if (ruleType === 'item_discount_fixed') {
    var rupiah = Number(param.rupiah);
    if (!rupiah || rupiah <= 0) return { ok: false, error: 'item_discount_fixed: rupiah harus > 0' };
    newHarga = Math.max(0, hargaAwal - rupiah);
  } else {
    return { ok: false, error: 'Rule ' + ruleType + ' belum diimplementasi di versi ini' };
  }

  if (newHarga === hargaAwal) return { ok: true, applied: false };

  var updated = {};
  for (var k in item) updated[k] = item[k];
  updated.harga_akhir = newHarga;
  updated.subtotal_akhir = qty * newHarga;
  updated.diskon_item = Math.max(0, (qty * hargaAwal) - updated.subtotal_akhir);
  return { ok: true, applied: true, item: updated };
}

function _applyOrderRule_(ruleType, param, subtotal) {
  if (ruleType === 'order_discount_pct') {
    var persen = Number(param.persen);
    if (!persen || persen <= 0 || persen > 100) return { ok: false, error: 'order_discount_pct: persen harus 1-100' };
    return { ok: true, diskon: Math.round(subtotal * persen / 100) };
  }
  if (ruleType === 'order_discount_fixed') {
    var rupiah = Number(param.rupiah);
    if (!rupiah || rupiah <= 0) return { ok: false, error: 'order_discount_fixed: rupiah harus > 0' };
    return { ok: true, diskon: Math.min(rupiah, subtotal) };
  }
  return { ok: false, error: 'Rule ' + ruleType + ' belum diimplementasi di versi ini' };
}

// ============================================================
// TEST — jalankan dari Apps Script editor
// ============================================================
function testProgramEngine_() {
  Logger.log('==========================================');
  Logger.log('TEST PROGRAM ENGINE v' + PROGRAM_ENGINE_VERSION);
  Logger.log('==========================================');

  var produkMap = {
    'bolu a': { nama_produk: 'Bolu A', harga: 25000, satuan: 'pcs' },
    'bolu b': { nama_produk: 'Bolu B', harga: 30000, satuan: 'pcs' }
  };

  var items = [
    { nama_produk: 'Bolu A', qty: 10, harga_awal: 25000, subtotal_awal: 250000 },
    { nama_produk: 'Bolu B', qty: 5,  harga_awal: 30000, subtotal_awal: 150000 }
  ];

  Logger.log('');
  Logger.log('--- TEST 1: tanpa program ---');
  var r1 = calculateProgramsForOrder_(items, [], produkMap);
  Logger.log('ok=' + r1.ok + ', total_akhir=' + r1.total_akhir + ' (expected 400000)');

  Logger.log('');
  Logger.log('--- TEST 2: order_discount_pct 10% ---');
  var prog2 = {
    header: { id_program: 'P1', nama_program: 'Promo 10%', kombinasi_mode: 'exclusive', kalkulasi_mode: 'sequential', priority: 10, max_diskon_per_order: 0 },
    rules: [{ rule_type: 'order_discount_pct', parameter: '{"persen":10}' }]
  };
  var r2 = calculateProgramsForOrder_(items, [prog2], produkMap);
  Logger.log('ok=' + r2.ok + ', total_akhir=' + r2.total_akhir + ' (expected 360000), diskon=' + r2.total_diskon_order);

  Logger.log('');
  Logger.log('--- TEST 3: item_discount_pct 20% ---');
  var prog3 = {
    header: { id_program: 'P2', nama_program: 'Bolu A -20%', kombinasi_mode: 'exclusive', kalkulasi_mode: 'sequential', priority: 10, max_diskon_per_order: 0 },
    rules: [{ rule_type: 'item_discount_pct', parameter: '{"persen":20}' }]
  };
  var r3 = calculateProgramsForOrder_(items, [prog3], produkMap);
  Logger.log('ok=' + r3.ok + ', total_akhir=' + r3.total_akhir + ' (expected 350000)');
  r3.items_final.forEach(function(it) {
    Logger.log('  ' + it.nama_produk + ' qty=' + it.qty + ' harga_akhir=' + it.harga_akhir + ' subtotal=' + it.subtotal_akhir);
  });

  Logger.log('');
  Logger.log('--- TEST 4: item 20% + order 10% (sequential) ---');
  var prog4 = {
    header: { id_program: 'P3', nama_program: 'Bundling', kombinasi_mode: 'stackable', kalkulasi_mode: 'sequential', priority: 10, max_diskon_per_order: 0 },
    rules: [
      { rule_type: 'item_discount_pct', parameter: '{"persen":20}' },
      { rule_type: 'order_discount_pct', parameter: '{"persen":10}' }
    ]
  };
  var r4 = calculateProgramsForOrder_(items, [prog4], produkMap);
  Logger.log('ok=' + r4.ok + ', subtotal_setelah_item=' + r4.subtotal_setelah_item + ', diskon_order=' + r4.total_diskon_order + ', total_akhir=' + r4.total_akhir);
  Logger.log('  expected: subtotal_setelah_item=350000, diskon=35000, total_akhir=315000');

  Logger.log('');
  Logger.log('--- TEST 5: item_fixed_price Bolu A @ 18000 ---');
  var prog5 = {
    header: { id_program: 'P4', nama_program: 'Fixed 18rb', kombinasi_mode: 'exclusive', kalkulasi_mode: 'sequential', priority: 10, max_diskon_per_order: 0 },
    rules: [{ rule_type: 'item_fixed_price', parameter: '{"harga":18000}' }]
  };
  var r5 = calculateProgramsForOrder_(items, [prog5], produkMap);
  Logger.log('ok=' + r5.ok + ', total_akhir=' + r5.total_akhir + ' (expected 330000)');

  Logger.log('');
  Logger.log('--- TEST 6: order 50% cap 100rb ---');
  var prog6 = {
    header: { id_program: 'P5', nama_program: '50% cap 100rb', kombinasi_mode: 'exclusive', kalkulasi_mode: 'sequential', priority: 10, max_diskon_per_order: 100000 },
    rules: [{ rule_type: 'order_discount_pct', parameter: '{"persen":50}' }]
  };
  var r6 = calculateProgramsForOrder_(items, [prog6], produkMap);
  Logger.log('ok=' + r6.ok + ', diskon=' + r6.total_diskon_order + ' (expected 100000), total_akhir=' + r6.total_akhir + ' (expected 300000)');

  Logger.log('');
  Logger.log('==========================================');
  Logger.log('SEMUA TEST SELESAI');
  Logger.log('==========================================');
}