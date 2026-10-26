// KK22 - Cache Layer
// CacheService helper untuk hemat execution time.

const CACHE = {
  TTL_PRODUK:   300,
  TTL_CONFIG:   600,
  TTL_SESSION:  120,
  TTL_RESELLER: 300
};

function cacheGet_(key) {
  try {
    var v = CacheService.getScriptCache().get(key);
    return v ? JSON.parse(v) : null;
  } catch (e) { return null; }
}

function cachePut_(key, obj, ttl) {
  try {
    var s = JSON.stringify(obj);
    if (s.length > 95000) return false;
    CacheService.getScriptCache().put(key, s, ttl || 300);
    return true;
  } catch (e) { return false; }
}

function cacheInvalidate_(keys) {
  try {
    if (typeof keys === 'string') keys = [keys];
    if (keys.length > 0) CacheService.getScriptCache().removeAll(keys);
  } catch (e) {}
}

function cacheBustProduk_() {
  cacheInvalidate_(['produk_list']);
}

function cacheBustConfig_() {
  cacheInvalidate_(['config_all']);
}

function cacheBustSession_(token) {
  if (token) cacheInvalidate_(['session_' + token]);
}