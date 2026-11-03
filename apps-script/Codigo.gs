/**
 * Inventario Café Pretexto — base de datos en Google Sheets
 *
 * Este código va dentro de la hoja de cálculo (Extensiones → Apps Script).
 * Guarda todo lo que hace la app: productos, cada movimiento (lo que llega,
 * lo que se usa, lo que se tira y cada conteo) con su fecha y hora, las ventas
 * del día, proveedores y recordatorios. Cada movimiento es una fila, así queda
 * el historial día por día para ver cuánto se gasta de cada cosa.
 *
 * Instalación: ver README.md del repositorio.
 */

const ZONA = "America/Mexico_City";
// Lo que no se puede eliminar desde la app (solo directo en la hoja, por el dueño)
const NO_BORRAR = ["productos", "proveedores", "ventas", "config", "reportes"];

const HOJAS = {
  productos:     { nombre: "Productos",     cols: ["id","nombre","categoria","unidad","stock","minimo","precio","compraUnidad","contenido","proveedorId","diasPedido","diasSugeridos","contadoEn","actualizado"] },
  movimientos:   { nombre: "Movimientos",   cols: ["id","fecha","hora","dia_semana","productoId","nombre","tipo","delta","cantidad","unidad","costo","nota","autor","ts"] },
  ventas:        { nombre: "Ventas",        cols: ["id","fecha","dia_semana","monto","nota","autor","actualizado"] },
  proveedores:   { nombre: "Proveedores",   cols: ["id","nombre","contacto","telefono","dias","hora","entrega","nota","actualizado"] },
  recordatorios: { nombre: "Recordatorios", cols: ["id","texto","fecha","hora","repetir","hecho","autor","ts"] },
  config:        { nombre: "Config",        cols: ["id","datos"] },
  reportes:      { nombre: "Reportes",      cols: ["id","titulo","desde","hasta","url","autor","ts"] }
};
// Columnas que se guardan como texto (para que la hoja no las convierta en fechas u horas)
const TEXTO = ["desde","hasta","id","fecha","hora","telefono","dias","diasPedido","dia_semana","productoId","proveedorId","datos"];
const JSON_COLS = ["diasPedido","dias","datos"];
const NUM = ["stock","minimo","precio","contenido","contadoEn","actualizado","entrega","ts","delta","cantidad","costo","monto"];
const BOOL = ["diasSugeridos","hecho"];
const DIA_SEMANA = ["","lunes","martes","miércoles","jueves","viernes","sábado","domingo"];

/* ---------- menú en la hoja ---------- */
function onOpen() {
  SpreadsheetApp.getUi().createMenu("Inventario")
    .addItem("1. Configurar (primera vez)", "configurar")
    .addItem("Rehacer pestaña Análisis", "crearAnalisis")
    .addItem("Recuperar productos de la barra borrados", "restaurarProductos")
    .addItem("Cambiar clave del equipo", "cambiarClave")
    .addToUi();
}

function configurar() {
  const ui = SpreadsheetApp.getUi();
  Object.keys(HOJAS).forEach(hoja);
  // productos de la barra (del inventario del recetario)
  const shP = hoja("productos");
  if (shP.getLastRow() < 2) {
    const ahora = Date.now();
    PRODUCTOS_INICIALES.forEach(function (p) {
      guardar("productos", p[0], { nombre: p[1], categoria: p[2], unidad: p[3], stock: null, minimo: null, precio: null,
        compraUnidad: "", contenido: null, proveedorId: "", diasPedido: p[4], diasSugeridos: true, actualizado: ahora }, "set", "");
    });
  }
  crearAnalisis();
  ["Hoja 1", "Sheet1"].forEach(function (n) { const h = SpreadsheetApp.getActive().getSheetByName(n); if (h && h.getLastRow() === 0) SpreadsheetApp.getActive().deleteSheet(h); });
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty("CLAVE")) {
    const r = ui.prompt("Clave del equipo", "Escribe una clave que el equipo va a usar para entrar a la app (mínimo 4 caracteres):", ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK || r.getResponseText().trim().length < 4) {
      ui.alert("No se guardó la clave. Vuelve a correr Inventario → Configurar.");
      return;
    }
    props.setProperty("CLAVE", r.getResponseText().trim());
  }
  ui.alert("Listo. La hoja ya tiene sus pestañas, los productos de la barra y la pestaña Análisis.");
}

function restaurarProductos() {
  const hay = filas("productos").map(function (p) { return p.id; });
  const ahora = Date.now();
  const faltan = PRODUCTOS_INICIALES.filter(function (p) { return hay.indexOf(p[0]) < 0; });
  faltan.forEach(function (p) {
    guardar("productos", p[0], { nombre: p[1], categoria: p[2], unidad: p[3], stock: null, minimo: null, precio: null,
      compraUnidad: "", contenido: null, proveedorId: "", diasPedido: p[4], diasSugeridos: true, actualizado: ahora }, "set", "");
  });
  SpreadsheetApp.getUi().alert(faltan.length ? "Se recuperaron: " + faltan.map(function (p) { return p[1]; }).join(", ") : "No faltaba ningún producto de la barra.");
}

function cambiarClave() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt("Nueva clave del equipo", "Todos tendrán que escribir la nueva clave en su teléfono:", ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() === ui.Button.OK && r.getResponseText().trim().length >= 4) {
    PropertiesService.getScriptProperties().setProperty("CLAVE", r.getResponseText().trim());
    ui.alert("Clave cambiada.");
  }
}

/* Hojas de resumen que se calculan solas con fórmulas */
function crearResumenes() {
  const ss = SpreadsheetApp.getActive();
  const resumenes = [
    ["Consumo diario", '=QUERY(Movimientos!A:N,"select B, F, J, sum(I), sum(K) where G = \'salida\' or G = \'merma\' group by B, F, J order by B desc label B \'Fecha\', F \'Producto\', J \'Unidad\', sum(I) \'Cantidad usada\', sum(K) \'Costo ($)\'",1)'],
    ["Consumo por producto", '=QUERY(Movimientos!A:N,"select F, J, sum(I), sum(K), count(I) where G = \'salida\' or G = \'merma\' group by F, J order by sum(K) desc label F \'Producto\', J \'Unidad\', sum(I) \'Total usado\', sum(K) \'Costo ($)\', count(I) \'Registros\'",1)'],
    ["Compras por día", '=QUERY(Movimientos!A:N,"select B, sum(K), count(I) where G = \'entrada\' group by B order by B desc label B \'Fecha\', sum(K) \'Comprado ($)\', count(I) \'Entradas\'",1)'],
    ["Consumo por día de la semana", '=QUERY(Movimientos!A:N,"select D, sum(K), count(I) where G = \'salida\' or G = \'merma\' group by D order by sum(K) desc label D \'Día\', sum(K) \'Costo ($)\', count(I) \'Registros\'",1)']
  ];
  resumenes.forEach(function (r) {
    let sh = ss.getSheetByName(r[0]);
    if (!sh) sh = ss.insertSheet(r[0]);
    sh.getRange("A1").setFormula("=IFERROR(" + r[1].slice(1) + ",\"Sin datos todavía\")");
    sh.setFrozenRows(1);
  });
}

/* ---------- API para la app ---------- */
function doGet() {
  return salida({ ok: true, app: "inventario-pretexto" });
}

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); }
  catch (x) { return salida({ ok: false, code: "invalid_argument", error: "Petición inválida" }); }
  const clave = PropertiesService.getScriptProperties().getProperty("CLAVE");
  if (!clave) return salida({ ok: false, code: "sin_configurar", error: "Falta correr Inventario → Configurar en la hoja" });
  if (String(b.clave || "") !== clave) return salida({ ok: false, code: "clave", error: "Clave incorrecta" });
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    if (b.accion === "leer") return salida({ ok: true, datos: leerTodo(Number(b.desde) || 0), hoja: SpreadsheetApp.getActive().getUrl() });
    if (b.accion === "reporte") return salida({ ok: true, doc: crearReporte(String(b.desde), String(b.hasta), b.autor || "") });
    if (b.accion === "guardar") return salida({ ok: true, doc: guardar(b.col, b.id, b.data || {}, b.modo, b.autor || "") });
    if (b.accion === "borrar") {
      if (NO_BORRAR.indexOf(b.col) >= 0) return salida({ ok: false, code: "no_permitido", error: "Desde la app no se pueden eliminar productos ni proveedores. Se quitan directo en la hoja." });
      borrar(b.col, b.id); return salida({ ok: true });
    }
    return salida({ ok: false, code: "invalid_argument", error: "Acción desconocida" });
  } catch (err) {
    const msg = String(err && err.message || err);
    return salida({ ok: false, code: msg.indexOf("No existe") === 0 ? "invalid_argument" : "unavailable", error: msg });
  } finally {
    lock.releaseLock();
  }
}

function salida(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- lectura y escritura ---------- */
function hoja(col) {
  const def = HOJAS[col];
  if (!def) throw new Error("Colección desconocida: " + col);
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(def.nombre);
  if (!sh) {
    sh = ss.insertSheet(def.nombre);
    sh.getRange(1, 1, 1, def.cols.length).setValues([def.cols]).setFontWeight("bold");
    sh.setFrozenRows(1);
    def.cols.forEach(function (c, i) {
      if (TEXTO.indexOf(c) >= 0) sh.getRange(1, i + 1, sh.getMaxRows(), 1).setNumberFormat("@");
    });
  }
  return sh;
}

function encabezados(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
}

function leerValor(k, v) {
  if (v === "" || v === null || v === undefined) return null;
  if (v instanceof Date) {
    if (k === "hora") return Utilities.formatDate(v, ZONA, "HH:mm");
    return Utilities.formatDate(v, ZONA, "yyyy-MM-dd");
  }
  if (typeof v === "string" && (JSON_COLS.indexOf(k) >= 0 || /^[\[{]/.test(v))) { try { return JSON.parse(v); } catch (x) { return v; } }
  if (NUM.indexOf(k) >= 0) { const n = Number(v); return isNaN(n) ? null : n; }
  if (BOOL.indexOf(k) >= 0) return v === true || v === "TRUE" || v === "true";
  return k === "id" ? String(v) : v;
}

function aObjeto(h, fila) {
  const o = {};
  h.forEach(function (k, i) { if (k) o[k] = leerValor(k, fila[i]); });
  return o;
}

function filas(col) {
  const sh = hoja(col);
  if (sh.getLastRow() < 2) return [];
  const h = encabezados(sh);
  return sh.getRange(2, 1, sh.getLastRow() - 1, h.length).getValues()
    .filter(function (f) { return f[0] !== ""; })
    .map(function (f) { return aObjeto(h, f); });
}

function leerTodo(desde) {
  const r = {};
  Object.keys(HOJAS).forEach(function (col) {
    let lista = filas(col);
    if (col === "movimientos" && desde) lista = lista.filter(function (m) { return (m.ts || 0) >= desde; });
    r[col] = lista;
  });
  return r;
}

function guardar(col, id, data, modo, autor) {
  if (!id) throw new Error("Falta el id");
  const sh = hoja(col);
  const h = encabezados(sh);
  Object.keys(data).forEach(function (k) {
    if (h.indexOf(k) < 0) { sh.getRange(1, h.length + 1).setValue(k).setFontWeight("bold"); h.push(k); }
  });
  const n = sh.getLastRow() - 1;
  const ids = n > 0 ? sh.getRange(2, 1, n, 1).getValues().map(function (r) { return String(r[0]); }) : [];
  const fila = ids.indexOf(String(id));
  let obj;
  if (fila >= 0 && modo === "update") {
    obj = aObjeto(h, sh.getRange(fila + 2, 1, 1, h.length).getValues()[0]);
    Object.keys(data).forEach(function (k) { obj[k] = data[k]; });
  } else if (modo === "update") {
    throw new Error("No existe " + col + "/" + id);
  } else {
    obj = JSON.parse(JSON.stringify(data));
  }
  obj.id = String(id);
  if (col === "movimientos") completarMovimiento(obj, autor);
  if (col === "ventas") { obj.fecha = obj.fecha || obj.id; obj.dia_semana = diaSemana(new Date(obj.fecha + "T12:00:00")); }
  const valores = h.map(function (k) {
    const v = obj[k];
    if (v === null || v === undefined) return "";
    if (JSON_COLS.indexOf(k) >= 0 || typeof v === "object") return JSON.stringify(v);
    return v;
  });
  if (fila >= 0) sh.getRange(fila + 2, 1, 1, h.length).setValues([valores]);
  else sh.getRange(sh.getLastRow() + 1, 1, 1, h.length).setValues([valores]);
  return obj;
}

function borrar(col, id) {
  const sh = hoja(col);
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const ids = sh.getRange(2, 1, n, 1).getValues().map(function (r) { return String(r[0]); });
  const i = ids.indexOf(String(id));
  if (i >= 0) sh.deleteRow(i + 2);
}

function diaSemana(d) {
  return DIA_SEMANA[Number(Utilities.formatDate(d, ZONA, "u"))];
}

/* Cada movimiento guarda su fecha, hora, día de la semana y costo, para poder analizar después */
function completarMovimiento(m, autor) {
  const d = new Date(Number(m.ts) || Date.now());
  m.ts = d.getTime();
  m.fecha = Utilities.formatDate(d, ZONA, "yyyy-MM-dd");
  m.hora = Utilities.formatDate(d, ZONA, "HH:mm");
  m.dia_semana = diaSemana(d);
  m.cantidad = Math.abs(Number(m.delta) || 0);
  if (!m.autor) m.autor = autor || "";
  if (m.costo === undefined || m.costo === null || m.costo === "") {
    const p = filas("productos").filter(function (x) { return x.id === m.productoId; })[0];
    if (p && Number(p.precio) > 0) {
      const porUnidad = Number(p.contenido) > 0 && p.compraUnidad ? Number(p.precio) / Number(p.contenido) : Number(p.precio);
      m.costo = Math.round(m.cantidad * porUnidad * 100) / 100;
    }
  }
}

/* ---------- pestaña Análisis: tablas y gráficas que se actualizan solas ---------- */
function crearAnalisis() {
  crearResumenes();
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName("Análisis");
  if (!sh) sh = ss.insertSheet("Análisis", 0);
  sh.getCharts().forEach(function (c) { sh.removeChart(c); });
  sh.clear();
  const desde = '"&TEXT(TODAY()-29,"yyyy-mm-dd")&"';
  const dinero = "$#,##0.00";

  sh.getRange("A1").setValue("Análisis de inventario · Café Pretexto").setFontSize(18).setFontWeight("bold");
  sh.getRange("A2").setValue("Se actualiza solo con lo que el equipo anota en la app. Dinero de los últimos 30 días; el lunes no se abre.").setFontStyle("italic");

  // Resumen de dinero (últimos 30 días)
  sh.getRange("A4:B4").setValues([["Últimos 30 días", "Monto"]]).setFontWeight("bold");
  sh.getRange("A5:A10").setValues([["Vendido (caja)"], ["Compras (lo que llegó)"], ["Insumos usados"], ["Merma (se tiró)"], ["Vendido − insumos"], ["Insumos / venta"]]);
  sh.getRange("B5").setFormula('=IFERROR(QUERY(Ventas!A:G,"select sum(D) where B >= \'' + desde + '\' label sum(D) \'\'",1),0)');
  sh.getRange("B6").setFormula('=IFERROR(QUERY(Movimientos!A:N,"select sum(K) where G = \'entrada\' and B >= \'' + desde + '\' label sum(K) \'\'",1),0)');
  sh.getRange("B7").setFormula('=IFERROR(QUERY(Movimientos!A:N,"select sum(K) where (G = \'salida\' or G = \'merma\') and B >= \'' + desde + '\' label sum(K) \'\'",1),0)');
  sh.getRange("B8").setFormula('=IFERROR(QUERY(Movimientos!A:N,"select sum(K) where G = \'merma\' and B >= \'' + desde + '\' label sum(K) \'\'",1),0)');
  sh.getRange("B9").setFormula("=B5-B7");
  sh.getRange("B10").setFormula('=IF(B5>0,B7/B5,"—")');
  sh.getRange("B5:B9").setNumberFormat(dinero);
  sh.getRange("B10").setNumberFormat("0.0%");

  // Dinero por día (30 días)
  sh.getRange("D4:G4").setValues([["Fecha", "Vendido", "Compras", "Insumos usados"]]).setFontWeight("bold");
  const filasDia = [];
  for (let i = 0; i < 30; i++) {
    const r = 5 + i;
    filasDia.push([
      '=TEXT(TODAY()-' + (29 - i) + ',"yyyy-mm-dd")',
      "=SUMIF(Ventas!$B:$B,$D" + r + ",Ventas!$D:$D)",
      '=SUMIFS(Movimientos!$K:$K,Movimientos!$B:$B,$D' + r + ',Movimientos!$G:$G,"entrada")',
      '=SUMIFS(Movimientos!$K:$K,Movimientos!$B:$B,$D' + r + ',Movimientos!$G:$G,"salida")+SUMIFS(Movimientos!$K:$K,Movimientos!$B:$B,$D' + r + ',Movimientos!$G:$G,"merma")'
    ]);
  }
  sh.getRange(5, 4, 30, 4).setFormulas(filasDia);
  sh.getRange("E5:G34").setNumberFormat(dinero);

  // Consumo por día de la semana (todo el historial)
  sh.getRange("I4:K4").setValues([["Día", "Insumos usados", "Vendido"]]).setFontWeight("bold");
  const dias = ["martes", "miércoles", "jueves", "viernes", "sábado", "domingo"];
  sh.getRange(5, 9, dias.length, 3).setFormulas(dias.map(function (d, i) {
    const r = 5 + i;
    return ['="' + d + '"',
      '=SUMIFS(Movimientos!$K:$K,Movimientos!$D:$D,$I' + r + ',Movimientos!$G:$G,"salida")+SUMIFS(Movimientos!$K:$K,Movimientos!$D:$D,$I' + r + ',Movimientos!$G:$G,"merma")',
      "=SUMIF(Ventas!$C:$C,$I" + r + ",Ventas!$D:$D)"];
  }));
  sh.getRange("J5:K10").setNumberFormat(dinero);

  // Lo que más se gasta (todo el historial)
  sh.getRange("M4").setFormula('=IFERROR(QUERY(Movimientos!A:N,"select F, sum(K), sum(I), J where (G = \'salida\' or G = \'merma\') group by F, J order by sum(K) desc limit 10 label F \'Producto\', sum(K) \'Costo ($)\', sum(I) \'Cantidad usada\', J \'Unidad\'",1),"Sin consumo anotado")');
  sh.getRange("N5:N14").setNumberFormat(dinero);

  // Por surtir ahora
  sh.getRange("R4").setFormula('=IFERROR(QUERY(Productos!A:N,"select B, E, F, D where E < F order by B label B \'Por surtir\', E \'Hay\', F \'Mínimo\', D \'Unidad\'",1),"Todo surtido")');

  // Gráficas
  const graf = function (tipo, rango, fila, col, titulo) {
    sh.insertChart(sh.newChart().setChartType(tipo).addRange(sh.getRange(rango)).setPosition(fila, col, 0, 0)
      .setOption("title", titulo).setOption("legend", { position: "bottom" }).setOption("width", 620).setOption("height", 320).build());
  };
  graf(Charts.ChartType.COLUMN, "D4:G34", 37, 1, "Dinero por día (30 días)");
  graf(Charts.ChartType.COLUMN, "I4:K10", 37, 9, "Por día de la semana");
  graf(Charts.ChartType.BAR, "M4:N14", 55, 1, "Lo que más se gasta (costo)");

  sh.setFrozenRows(2);
  sh.setColumnWidth(1, 190);
  sh.setColumnWidth(13, 220);
  sh.setColumnWidth(18, 200);
}

/* ---------- reportes en Google Docs ---------- */
function crearReporte(desde, hasta, autor) {
  const movs = filas("movimientos").filter(function (m) { return m.fecha >= desde && m.fecha <= hasta; });
  const ventas = filas("ventas").filter(function (v) { return v.fecha >= desde && v.fecha <= hasta; });
  const suma = function (lista, f) { return lista.reduce(function (a, x) { return a + (Number(f(x)) || 0); }, 0); };
  const dinero = function (v) { return "$" + (Math.round(v * 100) / 100).toLocaleString("es-MX"); };
  const usados = movs.filter(function (m) { return m.tipo === "salida" || m.tipo === "merma"; });
  const compras = movs.filter(function (m) { return m.tipo === "entrada"; });
  const vendido = suma(ventas, function (v) { return v.monto; });
  const insumos = suma(usados, function (m) { return m.costo; });
  const comprado = suma(compras, function (m) { return m.costo; });
  const merma = suma(usados.filter(function (m) { return m.tipo === "merma"; }), function (m) { return m.costo; });

  const titulo = "Reporte de inventario Café Pretexto · " + desde + " a " + hasta;
  const doc = DocumentApp.create(titulo);
  const body = doc.getBody();
  body.appendParagraph(titulo).setHeading(DocumentApp.ParagraphHeading.TITLE);
  body.appendParagraph("Generado el " + Utilities.formatDate(new Date(), ZONA, "yyyy-MM-dd HH:mm") + (autor ? " por " + autor : "") + ". El lunes no se abre.").setItalic(true);

  body.appendParagraph("Dinero").setHeading(DocumentApp.ParagraphHeading.HEADING1);
  body.appendTable([
    ["Concepto", "Monto"],
    ["Vendido (caja)", dinero(vendido)],
    ["Compras (lo que llegó)", dinero(comprado)],
    ["Insumos usados", dinero(insumos)],
    ["Merma (se tiró)", dinero(merma)],
    ["Vendido − insumos", dinero(vendido - insumos)],
    ["Insumos / venta", vendido > 0 ? Math.round(insumos / vendido * 1000) / 10 + "%" : "—"]
  ]);

  body.appendParagraph("Consumo por producto").setHeading(DocumentApp.ParagraphHeading.HEADING1);
  const porProd = {};
  usados.forEach(function (m) { const k = m.nombre + "|" + m.unidad; porProd[k] = porProd[k] || { n: m.nombre, u: m.unidad, c: 0, d: 0 }; porProd[k].c += Number(m.cantidad) || 0; porProd[k].d += Number(m.costo) || 0; });
  const dias = {};
  movs.concat(ventas).forEach(function (x) { if (x.dia_semana !== "lunes") dias[x.fecha] = 1; });
  const nd = Math.max(1, Object.keys(dias).length);
  const filasProd = Object.keys(porProd).map(function (k) { return porProd[k]; }).sort(function (a, b) { return b.d - a.d || b.c - a.c; });
  if (filasProd.length) {
    body.appendTable([["Producto", "Usado", "Por día", "Por semana (mar–dom)", "Costo"]].concat(filasProd.map(function (p) {
      return [p.n, red(p.c) + " " + p.u, red(p.c / nd) + " " + p.u, red(p.c / nd * 6) + " " + p.u, p.d ? dinero(p.d) : "—"];
    })));
  } else body.appendParagraph("No se anotó consumo en este periodo.");

  body.appendParagraph("Por día").setHeading(DocumentApp.ParagraphHeading.HEADING1);
  const porDia = {};
  ventas.forEach(function (v) { porDia[v.fecha] = porDia[v.fecha] || { v: 0, i: 0, c: 0, ds: v.dia_semana }; porDia[v.fecha].v += Number(v.monto) || 0; });
  movs.forEach(function (m) { porDia[m.fecha] = porDia[m.fecha] || { v: 0, i: 0, c: 0, ds: m.dia_semana };
    if (m.tipo === "entrada") porDia[m.fecha].c += Number(m.costo) || 0;
    if (m.tipo === "salida" || m.tipo === "merma") porDia[m.fecha].i += Number(m.costo) || 0; });
  const fechas = Object.keys(porDia).sort();
  if (fechas.length) {
    body.appendTable([["Fecha", "Día", "Vendido", "Insumos usados", "Compras"]].concat(fechas.map(function (f) {
      const d = porDia[f]; return [f, d.ds || "", dinero(d.v), dinero(d.i), dinero(d.c)];
    })));
  } else body.appendParagraph("Sin registros.");

  body.appendParagraph("Lo que llegó").setHeading(DocumentApp.ParagraphHeading.HEADING1);
  if (compras.length) body.appendTable([["Fecha", "Producto", "Cantidad", "Costo", "Anotó"]].concat(compras.map(function (m) {
    return [m.fecha, m.nombre, red(m.cantidad) + " " + m.unidad, m.costo ? dinero(m.costo) : "—", m.autor || ""];
  })));
  else body.appendParagraph("No se anotaron entradas.");
  doc.saveAndClose();

  const id = "r" + Date.now();
  const r = { titulo: titulo, desde: desde, hasta: hasta, url: doc.getUrl(), autor: autor || "", ts: Date.now() };
  guardar("reportes", id, r, "set", autor);
  r.id = id;
  return r;
}
function red(v) { return Math.round((Number(v) || 0) * 100) / 100; }

/* Productos de la barra (de inventario.md del recetario). Solo se cargan la primera vez. */
const PRODUCTOS_INICIALES = [
  ["agua-mineral", "Agua mineral", "Otros", "botella", ["TH"]],
  ["base-cristal", "Base cristal (smoothies)", "Tés y polvos", "kg", ["TU"]],
  ["base-neutra", "Base neutra (frappés y malteadas)", "Tés y polvos", "kg", ["TU"]],
  ["cafe-cold-brew", "Café para cold brew", "Café", "kg", ["TU"]],
  ["cafe-grano-espresso", "Café en grano (espresso)", "Café", "kg", ["TU"]],
  ["cafe-premolido", "Café premolido (café del día)", "Café", "kg", ["TU"]],
  ["canela-polvo", "Canela en polvo", "Tés y polvos", "kg", ["TU"]],
  ["chocolate-barra", "Chocolate en barra", "Tés y polvos", "pzas", ["TU"]],
  ["crema-batida", "Crema batida", "Lácteos y cremas", "pzas", ["TU", "FR"]],
  ["crema-batir-dulce", "Crema para batir dulce", "Lácteos y cremas", "L", ["TU", "FR"]],
  ["helado-vainilla", "Helado de vainilla", "Lácteos y cremas", "L", ["TH"]],
  ["hielo", "Hielo", "Otros", "bolsa", ["WE", "FR"]],
  ["jarabe-chocolate", "Jarabe de chocolate", "Jarabes y pulpas", "botella", ["TU"]],
  ["jarabe-frutos-rojos", "Jarabe de frutos rojos", "Jarabes y pulpas", "botella", ["TU"]],
  ["jarabe-kiwi", "Jarabe de kiwi", "Jarabes y pulpas", "botella", ["TU"]],
  ["jarabe-natural", "Jarabe natural", "Jarabes y pulpas", "L", ["TU"]],
  ["jarabes-sabor", "Jarabes de sabor (frappé, soda, latte)", "Jarabes y pulpas", "botella", ["TU"]],
  ["leche-entera", "Leche entera", "Lácteos y cremas", "L", ["TU", "FR"]],
  ["matcha", "Polvo para matcha", "Tés y polvos", "kg", ["TU"]],
  ["pulpa-limon", "Pulpa de limón", "Jarabes y pulpas", "L", ["WE"]],
  ["pulpa-mango", "Pulpa de mango", "Jarabes y pulpas", "L", ["WE"]],
  ["pulpa-naranja", "Pulpa de naranja", "Jarabes y pulpas", "L", ["WE"]],
  ["pulpas-fruta-smoothie", "Pulpas de fruta (smoothie)", "Jarabes y pulpas", "L", ["WE"]],
  ["te", "Té", "Tés y polvos", "caja", ["TU"]],
  ["tisanas", "Tisanas (Amore y las del menú)", "Tés y polvos", "bolsa", ["TU"]]
];
