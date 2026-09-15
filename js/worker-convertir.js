/* Web Worker: lee el Excel y lo convierte sin congelar la pantalla */
importScripts('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js', 'convertir.js');

self.onmessage = function (e) {
  const { id, info, nombre, buffer } = e.data;
  const t0 = Date.now();
  try {
    self.postMessage({ id, etapa: 'leyendo' });
    const wb = XLSX.read(buffer, Convertir.opcionesLectura(info));
    const tLectura = Date.now() - t0;
    self.postMessage({ id, etapa: 'convirtiendo', hojas: wb.SheetNames });
    const json = Convertir.convertir(info, wb, nombre);
    const resumen = Convertir.resumen(info, json);
    const texto = JSON.stringify(json);
    self.postMessage({ id, etapa: 'listo', texto, resumen, ms: { lectura: tLectura, total: Date.now() - t0 } });
  } catch (err) {
    self.postMessage({ id, etapa: 'error', error: err.message || String(err) });
  }
};
