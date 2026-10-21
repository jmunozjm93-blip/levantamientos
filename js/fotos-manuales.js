/* Catálogo complementario: una carga de Excel nunca borra estas fotos. */
window.FotosManuales = (() => {
  const normalizar = value => String(value || '').trim().toUpperCase();
  const combinar = (base, extra) => {
    const out = Object.assign(Object.create(null), extra || {});
    for (const [k, v] of Object.entries(base || {})) if (v) out[k] = v;
    return out;
  };
  const url = (id, size = 240) => /^img\/manuales\/[a-f0-9-]+\.jpg$/.test(id) ? id : `https://lh3.googleusercontent.com/d/${id}=w${size}`;
  let pending;
  function leer() {
    if (!pending) pending = fetch('data/fotos-manuales.json?t=' + Date.now(), {cache:'no-store'})
      .then(r => { if (r.status === 404) return {img:{}}; if (!r.ok) throw new Error('No se pudo leer el catálogo de fotos'); return r.json(); })
      .then(j => j.img || {}).catch(error => { console.warn("Fotos complementarias no disponibles", error); return {}; });
    return pending;
  }
  return {normalizar, combinar, url, leer};
})();
