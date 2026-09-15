/* Conversión de los Excel de levantamiento (uno por cliente) al JSON que consume el dashboard.
 * Corre en el navegador sobre libros abiertos con SheetJS (XLSX.read).
 * Regla de oro: los números son los del Excel, tal cual. No se estima nada.
 *
 * Formato de cada Excel (fijo):
 *   - Una hoja por departamento: dinámica "Stk Final UN" con Modelo / Descripcion / Temporada / Marca
 *     y una columna por Nro Sucursal. Arriba trae el bloque de filtros (Linea, Departamento Interno,
 *     Periodo Semana (nombre), Nombre Cliente, ...).
 *   - Una hoja de supervisores (Cod · Tienda · Supervisor): es la lista de tiendas que se cuentan.
 *   - Si hay varias hojas de departamento, la que se llama igual que el cliente es una copia resumen
 *     antigua y se omite.
 */
(function (global) {
  'use strict';

  // ---------- nombres de archivo ----------
  const CLIENTES = [
    { clave: 'falabella', nombre: 'Falabella', excel: 'FALABELLA', re: /^Levantamiento\s+Fala(bella)?\.xlsx$/i, etiqueta: 'Levantamiento Fala.xlsx' },
    { clave: 'paris',     nombre: 'Paris',     excel: 'PARIS',     re: /^Levantamiento\s+Paris\.xlsx$/i,        etiqueta: 'Levantamiento Paris.xlsx' },
    { clave: 'ripley',    nombre: 'Ripley',    excel: 'RIPLEY',    re: /^Levantamiento\s+Ripley\.xlsx$/i,       etiqueta: 'Levantamiento Ripley.xlsx' },
    { clave: 'lapolar',   nombre: 'La Polar',  excel: 'LA POLAR',  re: /^Levantamiento\s+La\s*polar\.xlsx$/i,   etiqueta: 'Levantamiento La polar.xlsx' },
    { clave: 'hites',     nombre: 'Hites',     excel: 'HITES',     re: /^Levantamiento\s+Hites\.xlsx$/i,        etiqueta: 'Levantamiento Hites.xlsx' },
  ];
  const IMAGENES = { clave: 'imagenes', nombre: 'Imágenes', re: /^Excel_Macro\.xlsx$/i, etiqueta: 'Excel_Macro.xlsx' };
  const TIPOS = CLIENTES.map(c => ({ tipo: 'cliente', clave: c.clave, nombre: c.nombre, etiqueta: c.etiqueta }))
    .concat([{ tipo: 'imagenes', clave: 'imagenes', nombre: 'Imágenes (Excel_Macro)', etiqueta: IMAGENES.etiqueta }]);

  function identificar(nombre) {
    const n = nombre.trim();
    for (const c of CLIENTES) if (c.re.test(n)) return { tipo: 'cliente', clave: c.clave, nombre: c.nombre };
    if (IMAGENES.re.test(n)) return { tipo: 'imagenes', clave: 'imagenes', nombre: IMAGENES.nombre };
    return null;
  }

  // Departamento Interno del Excel → nombre que se muestra en el dashboard
  const DEPTOS = {
    'CALZADO MUJER': 'Calzado mujer', 'DEPORTES': 'Deporte', 'DEPORTES MUJER': 'Deporte mujer',
    'DEPORTES HOMBRE': 'Deporte hombre', 'JUVENIL HOMBRE': 'Juvenil', 'KIDS ZAPATILLAS': 'Kids',
  };

  // ---------- utilitarios ----------
  const vacio = v => v == null || v === '';
  const txt = v => vacio(v) ? '' : String(v).trim();
  const nombrePersona = v => txt(v).replace(/\s+/g, ' ').toUpperCase();
  const titulo = s => { s = txt(s).toLowerCase(); return s.charAt(0).toUpperCase() + s.slice(1); };
  // el código de sucursal puede venir como número o texto: siempre se compara como texto
  const cod = v => (typeof v === 'number') ? String(v) : txt(v);

  function filasDe(ws) {
    if (ws['!data']) return ws['!data'].map(row => row ? row.map(c => (c ? c.v : null)) : []);
    return XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
  }

  // ---------- reconocer hojas ----------
  // dinámica: la fila de títulos empieza con "Modelo"; arriba, "Stk Final UN" y "Nro Sucursal"
  function leerDinamica(nombre, filas) {
    const iHdr = filas.findIndex(r => txt(r[0]).toLowerCase() === 'modelo');
    if (iHdr < 0) return null;
    const hdr = filas[iHdr];
    const cabecera = [txt(hdr[1]), txt(hdr[2]), txt(hdr[3])].map(s => s.toLowerCase());
    if (cabecera[0] !== 'descripcion' || cabecera[1] !== 'temporada' || cabecera[2] !== 'marca')
      throw new Error(`Hoja "${nombre}": esperaba las columnas Modelo · Descripcion · Temporada · Marca y encontré "${[hdr[0], hdr[1], hdr[2], hdr[3]].map(txt).join(' · ')}"`);
    const arriba = filas.slice(0, iHdr).map(r => r.map(txt).join(' ')).join(' ').toLowerCase();
    if (!/stk final un/.test(arriba)) throw new Error(`Hoja "${nombre}": la dinámica no es de "Stk Final UN"`);
    if (!/nro sucursal/.test(arriba)) throw new Error(`Hoja "${nombre}": la dinámica no tiene "Nro Sucursal" en las columnas`);

    // bloque de filtros: etiqueta en la columna A, valor en la B
    const meta = {};
    for (let i = 0; i < iHdr; i++) { const k = txt(filas[i][0]); if (k && !vacio(filas[i][1])) meta[k.toLowerCase()] = txt(filas[i][1]); }

    // columnas de sucursal: desde la E hasta "Total general"
    const sucCols = [];
    let colTotal = -1;
    for (let c = 4; c < hdr.length; c++) {
      const h = txt(hdr[c]);
      if (!h) continue;
      if (h.toLowerCase() === 'total general') { colTotal = c; break; }
      sucCols.push({ c, cod: cod(hdr[c]) });
    }
    if (!sucCols.length) throw new Error(`Hoja "${nombre}": no encontré columnas de sucursal`);

    const datos = [];
    let totalGeneral = null;
    for (let i = iHdr + 1; i < filas.length; i++) {
      const r = filas[i];
      const m = txt(r[0]);
      if (!m) continue;
      if (m.toLowerCase() === 'total general') { if (colTotal >= 0 && typeof r[colTotal] === 'number') totalGeneral = r[colTotal]; break; }
      datos.push(r);
    }
    return { nombre, meta, hdr, sucCols, colTotal, datos, totalGeneral };
  }

  // supervisores: tres columnas Cod · Tienda · Supervisor
  function leerSupervisores(nombre, filas) {
    const iHdr = filas.findIndex(r => /^cod/i.test(txt(r[0])) && /tienda/i.test(txt(r[1])) && /supervisor/i.test(txt(r[2])));
    if (iHdr < 0) return null;
    const tiendas = [], vistos = {};
    for (let i = iHdr + 1; i < filas.length; i++) {
      const r = filas[i], c = cod(r[0]);
      if (!c) continue;
      if (vistos[c]) throw new Error(`Hoja "${nombre}": la tienda ${c} aparece dos veces`);
      vistos[c] = 1;
      const nom = txt(r[1]) || c, sup = nombrePersona(r[2]);
      if (!sup) throw new Error(`Hoja "${nombre}": la tienda ${nom} no tiene supervisor`);
      tiendas.push({ cod: c, nombre: nom, sup });
    }
    if (!tiendas.length) throw new Error(`Hoja "${nombre}": la lista de supervisores está vacía`);
    return tiendas;
  }

  // ---------- un cliente ----------
  function parseCliente(wb, cli, nombreArchivo) {
    const dinamicas = [], omitidas = [];
    let tiendas = null, hojaSup = null;
    for (const nombre of wb.SheetNames) {
      const filas = filasDe(wb.Sheets[nombre]);
      const d = leerDinamica(nombre, filas);
      if (d) { dinamicas.push(d); continue; }
      const s = leerSupervisores(nombre, filas);
      if (s) {
        if (tiendas) throw new Error(`Hay dos hojas de supervisores ("${hojaSup}" y "${nombre}")`);
        tiendas = s; hojaSup = nombre;
      }
    }
    if (!dinamicas.length) throw new Error('No encontré ninguna hoja con la dinámica de Stk Final UN (columna Modelo)');
    if (!tiendas) throw new Error('No encontré la hoja de supervisores (columnas Cod · Tienda · Supervisor)');

    // la hoja que se llama como el cliente es una copia resumen antigua: se omite si hay otras
    const esResumen = d => d.nombre.trim().toLowerCase() === cli.nombre.toLowerCase();
    let usadas = dinamicas.filter(d => !esResumen(d));
    if (!usadas.length) usadas = dinamicas;
    dinamicas.filter(d => !usadas.includes(d)).forEach(d => omitidas.push(d.nombre));

    const supDe = {};
    tiendas.forEach(t => { supDe[t.cod] = t.sup; });

    const filas = [], hojas = [], deptosVistos = {};
    for (const d of usadas) {
      const clienteExcel = txt(d.meta['nombre cliente']).toUpperCase();
      if (clienteExcel && clienteExcel !== cli.excel)
        throw new Error(`Hoja "${d.nombre}": el Excel dice "Nombre Cliente: ${clienteExcel}" y el archivo es de ${cli.nombre}`);
      const depRaw = txt(d.meta['departamento interno']);
      const linea = txt(d.meta['linea']);
      const depto = DEPTOS[depRaw.toUpperCase()] || (/^all$/i.test(depRaw) || !depRaw ? (linea && !/^all$/i.test(linea) ? titulo(linea) : 'Todos') : titulo(depRaw));
      if (deptosVistos[depto]) throw new Error(`Las hojas "${deptosVistos[depto]}" y "${d.nombre}" son del mismo departamento (${depto})`);
      deptosVistos[depto] = d.nombre;

      const conSup = d.sucCols.filter(s => supDe[s.cod]);
      const sinSup = d.sucCols.filter(s => !supDe[s.cod]).map(s => s.cod);
      let uds = 0, omitidos = 0;
      const porModelo = {};
      for (const r of d.datos) {
        const m = txt(r[0]);
        let fila = porModelo[m];
        if (!fila) { fila = porModelo[m] = { d: depto, m, n: txt(r[1]), t: txt(r[2]), b: txt(r[3]), q: 0, s: {} }; }
        for (const sc of conSup) {
          const v = r[sc.c];
          if (typeof v !== 'number' || !isFinite(v) || v === 0) continue;
          fila.s[sc.cod] = (fila.s[sc.cod] || 0) + v;
          fila.q += v;
        }
      }
      let modelos = 0;
      for (const m of Object.keys(porModelo)) {
        const f = porModelo[m];
        // solo modelos con unidades en tiendas con supervisor
        if (f.q > 0) { filas.push(f); uds += f.q; modelos++; } else omitidos++;
      }
      hojas.push({ hoja: d.nombre, depto, semana: txt(d.meta['periodo semana (nombre)']) || txt(d.meta['periodo semana']), anio: txt(d.meta['año']),
                   modelos, uds, filasExcel: d.datos.length, totalExcel: d.totalGeneral, sucursales: d.sucCols.length, sinSupervisor: sinSup, omitidos });
    }
    const sups = {};
    tiendas.forEach(t => { sups[t.sup] = 1; });
    return {
      cliente: cli.nombre, clave: cli.clave, archivo: nombreArchivo, generado: new Date().toISOString(),
      hojas, omitidas, hojaSupervisores: hojaSup, supervisores: Object.keys(sups).sort(), tiendas, filas,
    };
  }

  // ---------- imágenes (Excel_Macro.xlsx, hoja Imagenes) ----------
  // "MODELO.jpg" → ID de Google Drive. Si un modelo aparece varias veces, manda la última fila.
  function parseImagenes(wb) {
    const nombre = wb.SheetNames.find(n => n.trim().toLowerCase() === 'imagenes') || wb.SheetNames[0];
    const filas = filasDe(wb.Sheets[nombre]);
    const img = {};
    let leidas = 0;
    for (const r of filas) {
      const archivo = txt(r[0]), link = txt(r[1]);
      if (!archivo || /^nombre de/i.test(archivo)) continue;
      const m = link.match(/[?&]id=([A-Za-z0-9_-]+)/);
      if (!m) continue;
      leidas++;
      img[archivo.replace(/\.(jpe?g|png|webp)$/i, '').toUpperCase()] = m[1];
    }
    if (!leidas) throw new Error(`La hoja "${nombre}" no tiene links de Google Drive (columnas Nombre del Archivo · Link)`);
    return { generado: new Date().toISOString(), filasExcel: leidas, modelos: Object.keys(img).length, img };
  }

  // ---------- punto de entrada ----------
  function convertir(info, wb, nombreArchivo) {
    if (info.tipo === 'imagenes') return parseImagenes(wb);
    const cli = CLIENTES.find(c => c.clave === info.clave);
    if (!cli) throw new Error('Cliente desconocido: ' + info.clave);
    return parseCliente(wb, cli, nombreArchivo);
  }
  function opcionesLectura(info) {
    const o = { dense: true, cellText: false, cellHTML: false, cellStyles: false };
    if (info.tipo === 'imagenes') o.sheets = ['Imagenes'];
    return o;
  }
  function resumen(info, json) {
    if (info.tipo === 'imagenes') return { tipo: 'imagenes', filasExcel: json.filasExcel, modelos: json.modelos };
    return { tipo: 'cliente', cliente: json.cliente, hojas: json.hojas, omitidas: json.omitidas, hojaSupervisores: json.hojaSupervisores,
             tiendas: json.tiendas.length, supervisores: json.supervisores.length, skus: json.filas.length, uds: json.filas.reduce((a, f) => a + f.q, 0) };
  }

  global.Convertir = { TIPOS, CLIENTES, DEPTOS, identificar, convertir, opcionesLectura, resumen };
})(typeof window !== 'undefined' ? window : globalThis);
