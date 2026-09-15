# Levantamientos · Grupo Depor

Dashboard de levantamientos de stock por cliente, tienda y supervisor. Reemplaza al archivo único `levantamientos.html` (10 MB con todo embebido): ahora la página es liviana y los datos se publican por cliente desde el navegador, con el mismo flujo por token que el sell out.

**Sitio publicado:** https://jmunozjm93-blip.github.io/levantamientos/

| Archivo | Descripción |
|---|---|
| `index.html` | Dashboard: filtros por cliente, departamento, supervisor, tienda y búsqueda; foto, cantidad, paginación |
| `cargar.html` | App de carga: valida, convierte y sube cada Excel a GitHub |
| `js/convertir.js` | Conversión Excel → JSON |
| `js/worker-convertir.js` | Web Worker que lee el Excel sin congelar la pantalla |
| `data/manifiesto.json` | Qué clientes están publicados, con qué semana y cuándo se subieron |
| `data/clientes/<cliente>.json` | Levantamiento vigente de cada cliente (`falabella`, `paris`, `ripley`, `lapolar`, `hites`) |
| `data/imagenes.json` | Modelo → ID de foto en Google Drive (desde `Excel_Macro.xlsx`, hoja Imagenes) |

## Rutina de carga

1. Abrir **https://jmunozjm93-blip.github.io/levantamientos/cargar.html** (también hay un enlace "cargar levantamiento" arriba del dashboard).
2. Arrastrar los Excel, uno por cliente. El nombre debe ser exactamente uno de estos:
   `Levantamiento Fala.xlsx` · `Levantamiento Paris.xlsx` · `Levantamiento Ripley.xlsx` · `Levantamiento La polar.xlsx` · `Levantamiento Hites.xlsx`.
   Para actualizar las fotos se arrastra `Excel_Macro.xlsx`.
3. Revisar el resumen (modelos y unidades por hoja, tiendas sin supervisor) y pulsar **Subir a GitHub**.
4. En 1–2 minutos el dashboard muestra el levantamiento nuevo. Cada cliente reemplaza al anterior; el historial queda en los commits.

Los Excel no se suben: se quedan en el PC (están ignorados por git).

La app necesita un token de GitHub (fine-grained, permiso *Contents: Read and write* solo sobre este repositorio). Se pega una vez y queda guardado en el navegador. Si ya existe el token de `pagina-web`, basta con agregarle este repositorio en "Repository access".

## Formato de los Excel (fijo)

Cada archivo trae:

- **Una hoja por departamento** con la dinámica *Stk Final UN*: bloque de filtros arriba (Linea, Departamento Interno, Periodo Semana (nombre), Nombre Cliente…), fila de títulos `Modelo · Descripcion · Temporada · Marca` y una columna por `Nro Sucursal`, más `Total general`. Puede empezar en la fila 1 o en la 4.
- **Una hoja de supervisores** (`Supervisores` / `Hoja1`) con `Cod · Tienda · Supervisor`. Es la lista de tiendas que se cuentan.
- Si hay varias hojas de departamento, la que se llama igual que el cliente (`Paris`, `Ripley`, `Falabella`) es una copia resumen antigua y **se omite**. En La Polar y Hites es la única hoja, así que se usa.

Departamento que se muestra, según *Departamento Interno*: CALZADO MUJER → Calzado mujer · DEPORTES → Deporte · DEPORTES MUJER → Deporte mujer · DEPORTES HOMBRE → Deporte hombre · JUVENIL HOMBRE → Juvenil · KIDS ZAPATILLAS → Kids. Si viene `All` (Hites) se usa la *Linea* (Calzado).

## Reglas de conversión

Son las mismas del archivo original, verificadas contra sus datos embebidos:

- Solo cuentan las **tiendas que están en la hoja de supervisores** con un supervisor real: si la columna trae `Z` (relleno para tienda virtual o CD), `-` o `N/A`, la tienda se descarta. Las columnas de sucursales que no aparecen ahí (bodegas, .com, tiendas nuevas sin asignar) se descartan y se listan en el resumen de carga como "sin supervisor".
- La **cantidad** de un modelo es la suma de sus tiendas con supervisor, incluyendo valores negativos. Por eso puede diferir del "Total general" del Excel, que suma todas las columnas.
- Los modelos cuya suma queda en cero o negativa **no se incluyen**.
- Las fotos vienen de `Excel_Macro.xlsx`: `MODELO.jpg` → ID de Drive. Si un modelo aparece varias veces manda la última fila.
- El texto de Descripción, Temporada y Marca va tal cual viene en el Excel.

## Estructura de `data/clientes/<cliente>.json`

```json
{
  "cliente": "Paris", "clave": "paris", "archivo": "Levantamiento Paris.xlsx", "generado": "2026-09-15T21:21:27Z",
  "hojas": [{ "hoja": "CM", "depto": "Calzado mujer", "semana": "2026-W34 | 17-ago. al 23-ago.",
              "modelos": 99, "uds": 8656, "totalExcel": 8661, "sucursales": 47, "sinSupervisor": ["50"], "omitidos": 2 }],
  "omitidas": ["Paris"],
  "supervisores": ["ANGELICA GARCIA", "..."],
  "tiendas": [{ "cod": "104", "nombre": "104-PARIS BANDERA", "sup": "ANGELICA GARCIA" }],
  "filas": [{ "d": "Kids", "m": "149432C-001", "n": "CHUCK TAYLOR ALL STAR HIGH STRE", "t": "TODA TEMPORADA",
              "b": "CONVERSE", "q": 243, "s": { "10": 33, "16": 38 } }]
}
```

`s` es cantidad por código de tienda. El dashboard arma con esto la misma estructura interna del archivo original (`C/D/T/M/S/TI/RAW/SM`), así que la lógica de filtros es idéntica.
