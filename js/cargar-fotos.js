/* Usa la conexión existente; publica archivos y catálogo juntos, sin forzar la rama. */
(() => {
  const box = document.createElement('section'); box.className = 'box';
  box.innerHTML = `<h2>Fotos faltantes · JPG / PNG</h2>
    <p>Agrega fotos sin volver a cargar las planillas. El nombre debe ser el código exacto del modelo (ejemplo: ABC123.jpg). Puedes corregirlo antes de subir. Las fotos que ya tienen registro se conservan.</p>
    <div class="drop" id="fotoDrop"><p>Arrastra aquí tus fotos o una carpeta</p>
      <button type="button" id="fotoElegir">Elegir fotos…</button> <button type="button" id="fotoCarpeta">Elegir carpeta…</button>
      <input hidden type="file" id="fotoFiles" accept="image/jpeg,image/png" multiple>
      <input hidden type="file" id="fotoFolder" accept="image/jpeg,image/png" multiple webkitdirectory>
    </div><p class="small">Hasta 100 fotos por carga, máximo 15 MB por foto. Se guardan optimizadas. Solo se completan códigos sin foto registrada; un enlace existente que no abre no se considera faltante.</p>
    <p id="fotoStatus" role="status" aria-live="polite"></p><div id="fotoLista"></div>
    <button type="button" id="fotoSubir" disabled>Subir fotos faltantes</button> <button type="button" id="fotoLimpiar">Limpiar selección</button>`;
  document.querySelector('main').append(box);
  const q = id => box.querySelector('#' + id);
  let rows = [], busy = false;
  const status = text => { q('fotoStatus').textContent = text; };
  const code = file => FotosManuales.normalizar(file.name.replace(/\.(jpe?g|png)$/i, ''));
  const valid = key => /^[A-Z0-9][A-Z0-9 ._()-]{0,99}$/.test(key);
  function lock(value) { busy = value; box.querySelectorAll('button,input').forEach(x => x.disabled = value); q('fotoSubir').disabled = value || !rows.length; }
  function clear() { rows.forEach(r => URL.revokeObjectURL(r.preview)); rows = []; q('fotoLista').replaceChildren(); q('fotoSubir').disabled = true; }
  function draw() {
    q('fotoLista').replaceChildren();
    rows.forEach(row => {
      const line = document.createElement('div'); line.style.cssText = 'display:flex;gap:12px;align-items:center;flex-wrap:wrap;margin:12px 0';
      const pic = document.createElement('img'); pic.src = row.preview; pic.alt = row.file.name; pic.width = 64; pic.height = 64; pic.style.objectFit = 'contain';
      const label = document.createElement('label'); label.textContent = row.file.name + ' · Código: ';
      const input = document.createElement('input'); input.value = row.key; input.maxLength = 100; input.style.cssText = "background:transparent;color:inherit;border:1px solid #687180;border-radius:6px;padding:8px;max-width:100%;box-sizing:border-box";
      input.oninput = () => { row.key = FotosManuales.normalizar(input.value); row.note.textContent = ''; };
      label.append(input); row.note = document.createElement('span'); line.append(pic, label, row.note); q('fotoLista').append(line);
    }); q('fotoSubir').disabled = !rows.length;
  }
  function select(files) {
    if (busy) return;
    const accepted = files.filter(f => /\.(jpe?g|png)$/i.test(f.name) && f.size <= 15 * 1024 * 1024 && f.size > 0);
    if (accepted.length + rows.length > 100) { status('Elige como máximo 100 fotos por carga.'); return; }
    rows.push(...accepted.map(file => ({file, key:code(file), preview:URL.createObjectURL(file)}))); draw();
    status(`${rows.length} fotos seleccionadas. ${files.length - accepted.length} archivos omitidos por formato o tamaño. Revisa los códigos antes de subir.`);
  }
  q('fotoElegir').onclick = () => q('fotoFiles').click(); q('fotoCarpeta').onclick = () => q('fotoFolder').click();
  ['fotoFiles','fotoFolder'].forEach(id => q(id).onchange = e => { select([...e.target.files]); e.target.value = ''; });
  q('fotoLimpiar').onclick = () => { clear(); status('Selección vacía.'); };
  async function walk(entry) {
    if (entry.isFile) return [await new Promise((ok,no) => entry.file(ok,no))];
    const reader = entry.createReader(); let result = [], batch;
    do { batch = await new Promise((ok,no) => reader.readEntries(ok,no)); for (const child of batch) result.push(...await walk(child)); } while(batch.length);
    return result;
  }
  q('fotoDrop').ondragover = e => { e.preventDefault(); };
  q('fotoDrop').ondrop = async e => {
    e.preventDefault(); if (busy) return;
    const items = [...e.dataTransfer.items]; const files = [...e.dataTransfer.files];
    const entries = items.map(i => i.webkitGetAsEntry?.()).filter(Boolean);
    lock(true);
    try { const found = entries.length ? (await Promise.all(entries.map(walk))).flat() : files; lock(false); select(found); }
    catch(err) { status('No se pudo leer la carpeta: ' + err.message); }
    finally { lock(false); }
  };
  async function api(path, method = 'GET', body) {
    const response = await fetch(API + path, {method, cache:'no-store', headers:{Authorization:'Bearer ' + token(), Accept:'application/vnd.github+json','Content-Type':'application/json'}, body:body === undefined ? undefined : JSON.stringify(body)});
    const data = await response.json();
    if (!response.ok) throw new Error(`GitHub ${response.status}: ${data.message || 'No se pudo guardar'}`);
    return data;
  }
  async function snapshot() {
    const ref = await api('/git/ref/heads/' + BRANCH), head = ref.object.sha;
    const commit = await api('/git/commits/' + head);
    const tree = await api('/git/trees/' + commit.tree.sha + '?recursive=1');
    if(tree.truncated) throw new Error('El repositorio es demasiado grande para revisar todas las fotos.');
    async function read(path) {
      const item = tree.tree.find(x => x.path === path); if (!item) return null;
      const blob = await api('/git/blobs/' + item.sha);
      return new TextDecoder().decode(Uint8Array.from(atob(blob.content.replace(/\s/g,'')), c => c.charCodeAt(0)));
    }
    let base = {};
    if (REPO === 'pagina-web') {
      for (const name of ['Dashboard sell out.html','Dashboard Steve Madden.html']) {
        const html = await read(name); if (!html) throw new Error('Falta ' + name);
        const doc = new DOMParser().parseFromString(html, 'text/html');
        const node = doc.getElementById('IMG');
        const match = html.match(/const IMG = (?:FotosManuales\.combinar\()?(\{[^\r\n]+\})(?:,await FotosManuales\.leer\(\)\))?;/);
        const text = node ? node.textContent : match?.[1];
        if (!text) throw new Error('No se pudo comprobar el catálogo de ' + name);
        base = FotosManuales.combinar(base, JSON.parse(text));
      }
    } else {
      const text = await read('data/imagenes.json'); base = text ? JSON.parse(text).img || {} : {};
    }
    const text = await read('data/fotos-manuales.json'); const extra = text ? JSON.parse(text) : {img:{}};
    return {head, tree:commit.tree.sha, extra, existing:FotosManuales.combinar(base, extra.img)};
  }
  async function jpeg(file) {
    const bitmap = await createImageBitmap(file);
    try {
      if (!bitmap.width || !bitmap.height) throw new Error('Imagen vacía');
      const ratio = Math.min(1, 1600 / Math.max(bitmap.width, bitmap.height));
      const canvas = document.createElement('canvas'); canvas.width = Math.max(1, Math.round(bitmap.width * ratio)); canvas.height = Math.max(1, Math.round(bitmap.height * ratio));
      const ctx = canvas.getContext('2d'); ctx.fillStyle = '#fff'; ctx.fillRect(0,0,canvas.width,canvas.height); ctx.drawImage(bitmap,0,0,canvas.width,canvas.height);
      return canvas.toDataURL('image/jpeg',.88).split(',')[1];
    } finally { bitmap.close(); }
  }
  q('fotoSubir').onclick = () => {
    if (busy) return;
    if (!token()) { status('Primero guarda tu conexión de GitHub en el paso 1.'); return; }
    if (rows.some(r => !valid(r.key))) { status('Corrige los códigos: usa letras, números, espacios, puntos, guiones o paréntesis.'); return; }
    lock(true);
    enCola(async () => {
      status('Comprobando las fotos que ya existen…');
      const snap = await snapshot(), additions = [], seen = new Set();
      for (const row of rows) {
        if (snap.existing[row.key]) { row.note.textContent = 'Ya tiene foto · se conserva'; continue; }
        if (seen.has(row.key)) { row.note.textContent = 'Código repetido · se omite'; continue; }
        seen.add(row.key); additions.push(row); row.note.textContent = 'Pendiente';
      }
      if (!additions.length) { status('No hay fotos faltantes para agregar. No se cambió ningún archivo.'); return; }
      const entries = [], catalog = {...snap.extra, img:{...snap.extra.img}};
      for (let i = 0; i < additions.length; i++) {
        const row = additions[i]; status(`Preparando foto ${i+1} de ${additions.length}…`);
        const content = await jpeg(row.file);
        const blob = await api('/git/blobs','POST',{content,encoding:'base64'});
        const path = 'img/manuales/' + crypto.randomUUID() + '.jpg';
        entries.push({path,mode:'100644',type:'blob',sha:blob.sha}); catalog.img[row.key] = path;
      }
      catalog.actualizado = new Date().toISOString();
      entries.push({path:'data/fotos-manuales.json',mode:'100644',type:'blob',content:JSON.stringify(catalog)});
      const tree = await api('/git/trees','POST',{base_tree:snap.tree,tree:entries});
      const commit = await api('/git/commits','POST',{message:`Agrega ${additions.length} fotos faltantes`,tree:tree.sha,parents:[snap.head]});
      await api('/git/refs/heads/' + BRANCH,'PATCH',{sha:commit.sha,force:false});
      additions.forEach(r => r.note.textContent = 'Guardada');
      status(`✓ ${additions.length} fotos guardadas. Las existentes se conservaron. Se verán al terminar la publicación de GitHub Pages; luego actualiza el panel y vuelve a descargar el Excel si lo necesitas.`);
    }).catch(err => status('No se confirmó la carga: ' + err.message + '. Puedes reintentar; se revisarán otra vez las fotos existentes. Si otra persona actualizó la página, vuelve a pulsar Subir.'))
      .finally(() => lock(false));
  };
})();
