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
  // El número de semana al final del nombre es opcional: "Levantamiento Hites.xlsx" y
  // "Levantamiento Hites W38.xlsx" valen igual. Si viene, se comprueba contra la semana del Excel.
  const SEM = '(?:[\\s_-]*[wW]\\s*(\\d{1,2}))?\\.*';
  const CLIENTES = [
    { clave: 'falabella', nombre: 'Falabella', excel: 'FALABELLA', re: new RegExp(`^Levantamiento\\s+Fala(?:bella)?${SEM}\\.xlsx$`, 'i'), etiqueta: 'Levantamiento Fala.xlsx' },
    { clave: 'paris',     nombre: 'Paris',     excel: 'PARIS',     re: new RegExp(`^Levantamiento\\s+Paris${SEM}\\.xlsx$`, 'i'),         etiqueta: 'Levantamiento Paris.xlsx' },
    { clave: 'ripley',    nombre: 'Ripley',    excel: 'RIPLEY',    re: new RegExp(`^Levantamiento\\s+Ripley${SEM}\\.xlsx$`, 'i'),        etiqueta: 'Levantamiento Ripley.xlsx' },
    { clave: 'lapolar',   nombre: 'La Polar',  excel: 'LA POLAR',  re: new RegExp(`^Levantamiento\\s+La\\s*polar${SEM}\\.xlsx$`, 'i'),   etiqueta: 'Levantamiento La polar.xlsx' },
    { clave: 'hites',     nombre: 'Hites',     excel: 'HITES',     re: new RegExp(`^Levantamiento\\s+Hites${SEM}\\.xlsx$`, 'i'),         etiqueta: 'Levantamiento Hites.xlsx' },
    // Steve Madden se levanta en Paris (Nombre Cliente: PARIS) pero va en su propia pestaña
    { clave: 'steve',     nombre: 'Steve Madden', excel: 'PARIS', re: new RegExp(`^Levantamiento\\s+Steve\\s*Madden${SEM}\\.xlsx$`, 'i'), etiqueta: 'Levantamiento Steve Madden.xlsx', cliente: 'Paris', grupo: 'steve' },
  ];
  const IMAGENES = { clave: 'imagenes', nombre: 'Imágenes', re: /^Excel_Macro\.xlsx$/i, etiqueta: 'Excel_Macro.xlsx' };
  const TIPOS = CLIENTES.map(c => ({ tipo: 'cliente', clave: c.clave, nombre: c.nombre, etiqueta: c.etiqueta }))
    .concat([{ tipo: 'imagenes', clave: 'imagenes', nombre: 'Imágenes (Excel_Macro)', etiqueta: IMAGENES.etiqueta }]);

  function identificar(nombre) {
    const n = nombre.trim();
    for (const c of CLIENTES) {
      const m = n.match(c.re);
      if (m) return { tipo: 'cliente', clave: c.clave, nombre: c.nombre, semana: m[1] ? +m[1] : null };
    }
    if (IMAGENES.re.test(n)) return { tipo: 'imagenes', clave: 'imagenes', nombre: IMAGENES.nombre };
    return null;
  }

  // Departamento Interno del Excel → nombre que se muestra en el dashboard
  const DEPTOS = {
    'CALZADO MUJER': 'Calzado mujer', 'DEPORTES': 'Deporte', 'DEPORTES MUJER': 'Deporte mujer',
    'DEPORTES HOMBRE': 'Deporte hombre', 'JUVENIL HOMBRE': 'Juvenil', 'KIDS ZAPATILLAS': 'Kids',
  };

  // Columnas de atributos de la dinámica, por título. Modelo · Descripcion · Temporada · Marca son
  // obligatorias; la de línea (d) es opcional y puede ir en cualquier posición entre ellas: Hites y
  // La Polar la traen en la B para separar CALZADO / ROPA / ACCESORIOS en una sola hoja.
  const COL_ATRIB = {
    'modelo': 'm', 'descripcion': 'n', 'descripción': 'n', 'temporada': 't', 'marca': 'b',
    'linea': 'd', 'línea': 'd', 'departamento': 'd', 'departamento interno': 'd',
    'categoria': 'd', 'categoría': 'd', 'tipo': 'd', 'clase': 'd',
  };

  // supervisores que cambiaron: el Excel puede seguir trayendo el nombre antiguo
  const RENOMBRAR_SUPERVISOR = { 'ABRAHAN ASCUCI': 'CATALINA BRAVO', 'SEBASTIAN PIZARRO': 'CATALINA BRAVO' };
  // valores de la columna Supervisor que significan "nadie": esas tiendas se descartan
  const SIN_SUPERVISOR = ['Z', '-', 'N/A', 'NA', 'SIN SUPERVISOR'];

  // ---------- utilitarios ----------
  const vacio = v => v == null || v === '';
  const txt = v => vacio(v) ? '' : String(v).trim();
  const nombrePersona = v => { const n = txt(v).replace(/\s+/g, ' ').toUpperCase(); return RENOMBRAR_SUPERVISOR[n] || n; };
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
    // columnas de atributos: se leen por título hasta que aparece la primera sucursal
    const col = {};
    let c0 = 0;
    while (c0 < hdr.length) {
      const k = COL_ATRIB[txt(hdr[c0]).toLowerCase()];
      if (!k || col[k] !== undefined) break;
      col[k] = c0; c0++;
    }
    if (['m', 'n', 't', 'b'].some(k => col[k] === undefined))
      throw new Error(`Hoja "${nombre}": esperaba las columnas Modelo · Descripcion · Temporada · Marca (la de Linea es opcional) y encontré "${hdr.slice(0, 6).map(txt).join(' · ')}"`);
    const arriba = filas.slice(0, iHdr).map(r => r.map(txt).join(' ')).join(' ').toLowerCase();
    if (!/stk final un/.test(arriba)) throw new Error(`Hoja "${nombre}": la dinámica no es de "Stk Final UN"`);
    if (!/nro sucursal/.test(arriba)) throw new Error(`Hoja "${nombre}": la dinámica no tiene "Nro Sucursal" en las columnas`);

    // bloque de filtros: etiqueta en la columna A, valor en la B
    const meta = {};
    for (let i = 0; i < iHdr; i++) { const k = txt(filas[i][0]); if (k && !vacio(filas[i][1])) meta[k.toLowerCase()] = txt(filas[i][1]); }

    // columnas de sucursal: hasta "Total general"
    const sucCols = [];
    let colTotal = -1;
    for (let c = c0; c < hdr.length; c++) {
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
      const m = txt(r[col.m]);
      if (!m) continue;
      if (m.toLowerCase() === 'total general') { if (colTotal >= 0 && typeof r[colTotal] === 'number') totalGeneral = r[colTotal]; break; }
      datos.push(r);
    }
    return { nombre, meta, hdr, col, sucCols, colTotal, datos, totalGeneral };
  }

  // departamento que se muestra: el "Departamento Interno" traducido; si viene "All" manda la "Linea"
  // (Hites: CALZADO → Calzado; una hoja de accesorios: ACCESORIOS → Accesorios)
  function deptoDe(d) {
    const depRaw = txt(d.meta['departamento interno']), linea = txt(d.meta['linea']);
    return DEPTOS[depRaw.toUpperCase()] || (/^all$/i.test(depRaw) || !depRaw ? (linea && !/^all$/i.test(linea) ? titulo(linea) : 'Todos') : titulo(depRaw));
  }
  // semana del bloque de filtros como número comparable (2026-W37 → 202637); 0 si no trae una sola semana
  function semanaDe(d) {
    const m = txt(d.meta['periodo semana (nombre)']).match(/(\d{4})\s*-\s*W\s*(\d{1,2})/i);
    return m ? (+m[1]) * 100 + (+m[2]) : 0;
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
      // "Z" es un relleno para tienda virtual / centro de distribución: cuenta como sin supervisor y no entra
      if (!sup || SIN_SUPERVISOR.includes(sup)) continue;
      tiendas.push({ cod: c, nombre: nom, sup });
    }
    if (!tiendas.length) throw new Error(`Hoja "${nombre}": la lista de supervisores está vacía`);
    return tiendas;
  }

  // ---------- un cliente ----------
  function parseCliente(wb, cli, nombreArchivo, semanaArchivo) {
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

    // La hoja que se llama como el cliente (Paris, Ripley, Falabella…) suele ser una copia resumen antigua.
    // Solo se descarta si además repite el departamento de otra hoja o trae una semana más vieja; así, al
    // agregarle a Hites o La Polar una hoja de accesorios, su hoja de calzado no se pierde.
    const seLlamaComoCliente = d => { const n = d.nombre.trim().toUpperCase(); return n === cli.excel || n === cli.nombre.toUpperCase() || n === txt(d.meta['nombre cliente']).toUpperCase(); };
    const ultima = Math.max(...dinamicas.map(semanaDe));
    const esResumen = d => seLlamaComoCliente(d) && (
      dinamicas.some(o => o !== d && deptoDe(o) === deptoDe(d)) ||
      (semanaDe(d) > 0 && semanaDe(d) < ultima)
    );
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
      const depHoja = deptoDe(d), col = d.col;
      const conSup = d.sucCols.filter(s => supDe[s.cod]);
      const sinSup = d.sucCols.filter(s => !supDe[s.cod]).map(s => s.cod);
      let uds = 0, omitidos = 0, ultimaLinea = '';
      const porModelo = {}, deptosHoja = {};
      for (const r of d.datos) {
        const m = txt(r[col.m]);
        // con columna de línea el departamento es el de la fila; si viene en blanco se arrastra el anterior
        let depto = depHoja;
        if (col.d !== undefined) {
          const v = txt(r[col.d]);
          if (v) ultimaLinea = v;
          if (ultimaLinea) depto = DEPTOS[ultimaLinea.toUpperCase()] || titulo(ultimaLinea);
        }
        deptosHoja[depto] = 1;
        const clave = depto + '\u0000' + m;
        let fila = porModelo[clave];
        if (!fila) { fila = porModelo[clave] = { d: depto, m, n: txt(r[col.n]), t: txt(r[col.t]), b: txt(r[col.b]), q: 0, s: {} }; }
        for (const sc of conSup) {
          const v = r[sc.c];
          if (typeof v !== 'number' || !isFinite(v) || v === 0) continue;
          fila.s[sc.cod] = (fila.s[sc.cod] || 0) + v;
          fila.q += v;
        }
      }
      for (const dep of Object.keys(deptosHoja)) {
        if (deptosVistos[dep] && deptosVistos[dep] !== d.nombre)
          throw new Error(`Las hojas "${deptosVistos[dep]}" y "${d.nombre}" traen el mismo departamento (${dep})`);
        deptosVistos[dep] = d.nombre;
      }
      const depto = Object.keys(deptosHoja).sort().join(', ') || depHoja;
      let modelos = 0;
      for (const m of Object.keys(porModelo)) {
        const f = porModelo[m];
        // solo modelos con unidades en tiendas con supervisor
        if (f.q > 0) { filas.push(f); uds += f.q; modelos++; } else omitidos++;
      }
      hojas.push({ hoja: d.nombre, depto, semana: txt(d.meta['periodo semana (nombre)']) || txt(d.meta['periodo semana']), anio: txt(d.meta['año']),
                   modelos, uds, filasExcel: d.datos.length, totalExcel: d.totalGeneral, sucursales: d.sucCols.length, sinSupervisor: sinSup, omitidos });
    }
    // si el nombre del archivo trae la semana, tiene que ser la del Excel (evita subir una semana vieja)
    if (semanaArchivo) {
      const enExcel = usadas.map(semanaDe).filter(Boolean).map(s => s % 100);
      if (enExcel.length && enExcel.indexOf(semanaArchivo) < 0)
        throw new Error(`El nombre dice semana ${semanaArchivo} pero el Excel es de la semana ${[...new Set(enExcel)].join(' y ')} (${txt(usadas[0].meta['periodo semana (nombre)'])}). Revisa el archivo o su nombre.`);
    }
    const sups = {};
    tiendas.forEach(t => { sups[t.sup] = 1; });
    return {
      cliente: cli.cliente || cli.nombre, clave: cli.clave, grupo: cli.grupo || 'marcas', archivo: nombreArchivo, generado: new Date().toISOString(),
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
    return parseCliente(wb, cli, nombreArchivo, info.semana);
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
