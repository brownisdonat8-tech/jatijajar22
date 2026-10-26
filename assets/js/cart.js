/**
 * KK22 - Cart Helper
 * Keranjang belanja disimpan di localStorage.
 * Kunci item = nama_produk (sesuai payload backend buatPesanan).
 */

window.KK22_CART = (function() {
  const cfg = window.KK22_CONFIG;

  function _key() { return cfg.STORAGE_KEY_CART; }

  function _read() {
    try {
      const raw = localStorage.getItem(_key());
      if (!raw) return { items: [], tgl_kirim: '' };
      const obj = JSON.parse(raw);
      if (!obj || !Array.isArray(obj.items)) return { items: [], tgl_kirim: '' };
      return { items: obj.items, tgl_kirim: obj.tgl_kirim || '' };
    } catch (e) { return { items: [], tgl_kirim: '' }; }
  }

  function _write(state) {
    try { localStorage.setItem(_key(), JSON.stringify(state)); } catch (e) {}
  }

  function getItems() { return _read().items; }

  function getTglKirim() { return _read().tgl_kirim; }

  function setTglKirim(tgl) {
    const st = _read();
    st.tgl_kirim = String(tgl || '');
    _write(st);
  }

  function count() {
    return _read().items.reduce(function(s, it) { return s + (Number(it.qty) || 0); }, 0);
  }

  function total() {
    return _read().items.reduce(function(s, it) {
      return s + ((Number(it.qty) || 0) * (Number(it.harga_jual) || 0));
    }, 0);
  }

  function findItem(nama) {
    const nl = String(nama || '').toLowerCase();
    return _read().items.find(function(it) {
      return String(it.nama_produk || '').toLowerCase() === nl;
    }) || null;
  }

  function add(produk, qty) {
    const q = Math.max(1, Number(qty) || 1);
    const st = _read();
    const nl = String(produk.nama_produk || '').toLowerCase();
    const existing = st.items.find(function(it) {
      return String(it.nama_produk || '').toLowerCase() === nl;
    });
    if (existing) {
      existing.qty = (Number(existing.qty) || 0) + q;
    } else {
      st.items.push({
        nama_produk: produk.nama_produk,
        id_produk:   produk.id_produk || '',
        satuan:      produk.satuan || '',
        harga_jual:  Number(produk.harga_jual) || 0,
        gambar_url:  produk.gambar_url || '',
        qty:         q
      });
    }
    _write(st);
    return st;
  }

  function setQty(nama, qty) {
    const st = _read();
    const q = Math.max(1, Number(qty) || 1);
    const nl = String(nama || '').toLowerCase();
    st.items.forEach(function(it) {
      if (String(it.nama_produk || '').toLowerCase() === nl) it.qty = q;
    });
    _write(st);
    return st;
  }

  function remove(nama) {
    const st = _read();
    const nl = String(nama || '').toLowerCase();
    st.items = st.items.filter(function(it) {
      return String(it.nama_produk || '').toLowerCase() !== nl;
    });
    _write(st);
    return st;
  }

  function clear() {
    _write({ items: [], tgl_kirim: '' });
  }

  // Payload siap kirim ke buatPesanan (backend minta {nama_produk, qty})
  function toPayload() {
    return _read().items.map(function(it) {
      return { nama_produk: it.nama_produk, qty: Number(it.qty) || 0 };
    });
  }

  return {
    getItems: getItems,
    getTglKirim: getTglKirim,
    setTglKirim: setTglKirim,
    count: count,
    total: total,
    findItem: findItem,
    add: add,
    setQty: setQty,
    remove: remove,
    clear: clear,
    toPayload: toPayload
  };
})();