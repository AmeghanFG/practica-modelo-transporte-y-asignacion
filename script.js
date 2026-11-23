/* =====================================================================
  MODELOS DE TRANSPORTE Y ASIGNACIÓN
  Flujo: leerEntrada → validar → balancear → construirModelo → resolver → mostrarResultados
  El cálculo lo hace javascript-lp-solver (simplex, como Solver de Excel).

  Documentación: https://github.com/JWally/jsLPSolver
   ===================================================================== */

/* Tolerancia numérica (igual que en el método gráfico): dos números son "iguales" si su diferencia es menor que EPS. Evita problemas como 0.1 + 0.2 = 0.30000000000000004 */
const EPS = 1e-7;
/* Tamaño máximo de la matriz 100 variables */
const MAX_TAM = 10;

//-----------------------------------------------------------------
// * Formato
// -----------------------------------------------------------------

/* toNumber(texto): string → number. Acepta coma decimal. "" o texto → NaN */
function toNumber(texto) {
    const t = String(texto).trim().replace(",", ".");
    if (t === "") return NaN;
    return Number(t);
}

/* fmt(n): número visual con separador de miles y máximo 4 decimales: 41450 → "41,450" */
function fmt(n) {
    return n.toLocaleString("es-MX", { maximumFractionDigits: 4 });
}

/* escapar(texto): función por seguridad, evita que algo escrito por el usuario (ej. "<b>") rompa el HTML */
function escapar(texto) {
    return String(texto)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;");
}

/* nombreVar(i, j): recibe índices desde 0 y devuelve el nombre con índices desde 1. (0,0) → "x11";  si algún índice es 10 o más usa coma para no confundir: (0,9) → "x1,10" */
function nombreVar(i, j) {
    const a = i + 1, b = j + 1;
    return a < 10 && b < 10 ? "x" + a + b : "x" + a + "," + b;
}

/* Textos que cambian según el modelo. fila / col = nombre por defecto de cada grupo */
const TEXTOS = {
    transporte: { fila: "Origen", col: "Destino", filas: "Orígenes (filas)", cols: "Destinos (columnas)" },
    asignacion: { fila: "Trabajador", col: "Tarea", filas: "Trabajadores (filas)", cols: "Tareas (columnas)" },
};

/* grupos(): nombres de los grupos escritos arriba de la matriz (o los de por defecto). Devuelve {fila: "Planta", col: "Almacén"} */
function grupos() {
    const tipo = document.getElementById("tipo").value;
    const gfila = document.getElementById("grupoFilas");
    const gc = document.getElementById("grupoCols");
    return {
        fila: (gfila && gfila.value.trim()) || TEXTOS[tipo].fila,
        col: (gc && gc.value.trim()) || TEXTOS[tipo].col,
    };
}

//-----------------------------------------------------------------
// * Matriz de entrada (dibujar y leer)
// -----------------------------------------------------------------

/* lee la tabla y devuelve {m, n, grupoFilas, grupoCols, costos[m][n], oferta[m], demanda[n]}. Se usa para dos cosas: conservar lo escrito al redibujar y como parte de leerEntrada. */
function leerMatriz() {
    // VARIABLES DE LA FUNCIÓN
    const tabla = document.getElementById("tablaMatriz");
    const ofertas = tabla.querySelectorAll(".oferta");
    const demandas = tabla.querySelectorAll(".demanda");
    const m = ofertas.length, n = demandas.length; // hay un input de oferta al final de cada fila y uno de demanda al final de cada columna
    const gfila = document.getElementById("grupoFilas");
    const gcol = document.getElementById("grupoCols");

    // OBJETO
    const datos = {
        m: m, n: n,
        grupoFilas: gfila ? gfila.value : "", grupoCols: gcol ? gcol.value : "", // "" = usar el de por defecto
        costos: [], oferta: [], demanda: [],
    };

    for (let i = 0; i < m; i++) {
        datos.costos.push([]);
        for (let j = 0; j < n; j++) {
            datos.costos[i].push(tabla.querySelector('.costo[data-i="' + i + '"][data-j="' + j + '"]').value);
        }
    }
    // En asignación la oferta/demanda están deshabilitadas mostrando "1"
    // el valor de transporte que había antes se guarda en data-previo para no perderlo.
    ofertas.forEach(function (inp) { datos.oferta.push(inp.disabled ? inp.dataset.previo : inp.value); });
    demandas.forEach(function (inp) { datos.demanda.push(inp.disabled ? inp.dataset.previo : inp.value); });
    
    return datos;
}

/* Crea la tabla con el tamaño de la configuración, rellenando con lo que había en prev (lo que no existía queda vacío). Forma de la tabla (como en Excel)  */
function dibujarMatriz(prev) {
    const tipo = document.getElementById("tipo").value;
    const m = parseInt(document.getElementById("numFilas").value, 10);
    const n = parseInt(document.getElementById("numCols").value, 10);
    const txt = TEXTOS[tipo];
    const esAsig = tipo === "asignacion";

    // valor(arreglo, k): devuelve arreglo[k] si existe, si no "" (para celdas nuevas)
    function valor(arr, k) { return arr && arr[k] !== undefined ? arr[k] : ""; }

    // Fila 1 del encabezado: esquina | nombre del grupo de columnas | Oferta
    let h = '<thead><tr><th colspan="2" rowspan="2"></th>' +
        '<th colspan="' + n + '" class="grupo"><input type="text" id="grupoCols" class="form-control form-control-sm" ' +
        'placeholder="' + txt.col + '" value="' + escapar(prev.grupoCols || "") + '" aria-label="Nombre de las columnas (opcional)"></th>' +
        '<th rowspan="2" class="zona-oferta text-center align-middle">Oferta</th></tr>';
    // Fila 2 del encabezado: números de columna 1, 2, 3...
    h += "<tr>";
    for (let j = 0; j < n; j++) h += '<th class="text-center">' + (j + 1) + "</th>";
    h += "</tr></thead><tbody>";

    // Filas: (nombre del grupo, solo en la primera) + número + celdas de costo + oferta
    for (let i = 0; i < m; i++) {
        h += "<tr>";
        if (i === 0) {
            // rowspan = m: una sola celda alta que abarca todas las filas
            h += '<th rowspan="' + m + '" class="grupo align-middle" style="min-width:130px">' +
                '<input type="text" id="grupoFilas" class="form-control form-control-sm" placeholder="' + txt.fila +
                '" value="' + escapar(prev.grupoFilas || "") + '" aria-label="Nombre de las filas (opcional)"></th>';
        }
        h += '<th class="text-center" style="min-width:45px">' + (i + 1) + "</th>";
        for (let j = 0; j < n; j++) {
            const v = nombreVar(i, j);
            const costo = prev.costos && prev.costos[i] ? valor(prev.costos[i], j) : "";
            h += '<td><div class="celda"><span class="etiqueta-var">' + v + "</span>" +
                '<input type="text" inputmode="decimal" class="form-control form-control-sm costo" data-i="' + i +
                '" data-j="' + j + '" value="' + escapar(costo) + '" aria-label="Costo de ' + v + '"></div></td>';
        }
        h += '<td class="zona-oferta">' + campoTotal("oferta", i, valor(prev.oferta, i), esAsig) + "</td></tr>";
    }

    // Última fila: demanda
    h += '<tr><th colspan="2" class="zona-demanda text-center">Demanda</th>';
    for (let j = 0; j < n; j++) h += '<td class="zona-demanda">' + campoTotal("demanda", j, valor(prev.demanda, j), esAsig) + "</td>";
    h += "<td></td></tr></tbody>";

    document.getElementById("tablaMatriz").innerHTML = h;
    actualizarTextos();
    actualizarSelectsVariables();
}

/* campoTotal(clase, k, valorPrevio, esAsig): input de oferta o demanda.
   En asignación se muestra "1" deshabilitado y se guarda el valor anterior en data-previo. */
function campoTotal(clase, k, valorPrevio, esAsig) {
    const etiqueta = (clase === "oferta" ? "Oferta de la fila " : "Demanda de la columna ") + (k + 1);
    return '<input type="text" inputmode="decimal" class="form-control form-control-sm text-end ' + clase +
        '" data-k="' + k + '" aria-label="' + etiqueta + '"' +
        (esAsig ? ' value="1" disabled data-previo="' + escapar(valorPrevio) + '"' : ' value="' + escapar(valorPrevio) + '"') + ">";
}

/* actualizarTextos(): etiquetas y notas que dependen del modelo elegido */
function actualizarTextos() {
    const tipo = document.getElementById("tipo").value;
    document.getElementById("etqFilas").textContent = TEXTOS[tipo].filas;
    document.getElementById("etqCols").textContent = TEXTOS[tipo].cols;
    document.getElementById("notaMatriz").textContent = (tipo === "transporte"
        ? "Escribe el costo (o beneficio) por unidad de cada ruta. Si la oferta total no es igual a la demanda total, se agrega una fila o columna ficticia con costo 0 al calcular."
        : "Escribe el costo (o beneficio) de que cada trabajador haga cada tarea. Oferta y demanda valen 1. Si hay distinto número de filas y columnas, se agregan ficticias con costo 0 al calcular.") +
        " Los nombres de arriba y de la izquierda son opcionales.";
    document.getElementById("notaRestricciones").textContent = (tipo === "transporte"
        ? "Las restricciones de oferta (Σ fila = oferta), demanda (Σ columna = demanda) y no negatividad (xij ≥ 0) se agregan automáticamente. Aquí solo agrega restricciones extra, como x33 = 950 o x41 ≤ 100."
        : "Las restricciones de asignación (Σ fila = 1, Σ columna = 1) y xij ∈ {0, 1} se agregan automáticamente. Aquí solo agrega restricciones extra, como x21 = 1 (asignación obligatoria).") +
        " Para una ruta o combinación no disponible, usa xij = 0.";
}

/* redibujar(forzar): se llama cuando cambia el modelo (forzar = true) o el tamaño. Valida el tamaño primero. */
function redibujar(forzar) {
    const m = Number(document.getElementById("numFilas").value);
    const n = Number(document.getElementById("numCols").value);
    const error = document.getElementById("errorTam");
    if (!Number.isInteger(m) || !Number.isInteger(n) || m < 1 || n < 1 || m > MAX_TAM || n > MAX_TAM) {
        error.textContent = "Filas y columnas deben ser números enteros entre 1 y " + MAX_TAM + ".";
        return; // no se redibuja: se conserva la tabla anterior
    }
    error.textContent = "";
    const actual = leerMatriz();
    // Si el tamaño no cambió (y no es cambio de modelo), no se redibuja: así no se pierde
    // el clic o lo que el usuario estaba escribiendo en la tabla
    if (forzar !== true && actual.m === m && actual.n === n) return;
    dibujarMatriz(actual); // ← aquí se conserva lo que ya estaba escrito
}

//-----------------------------------------------------------------
// * Restricciones adicionales (interfaz)
// -----------------------------------------------------------------

/* opcionesVariables(): lista de variables para los menús: [{valor:"i,j", texto}] */
function opcionesVariables() {
    const tabla = document.getElementById("tablaMatriz");
    const m = tabla.querySelectorAll(".oferta").length;
    const n = tabla.querySelectorAll(".demanda").length;
    const g = grupos();
    const lista = [];
    for (let i = 0; i < m; i++) {
        for (let j = 0; j < n; j++) {
            lista.push({ valor: i + "," + j, texto: nombreVar(i, j) + " (" + g.fila + " " + (i + 1) + " → " + g.col + " " + (j + 1) + ")" });
        }
    }
    return lista;
}

/* Rellena un menú conservando lo que estaba elegido */
function llenarSelect(select, opciones) {
    const elegido = select.value;
    let h = '<option value="">— variable —</option>';
    opciones.forEach(function (o) { h += '<option value="' + o.valor + '">' + escapar(o.texto) + "</option>"; });
    select.innerHTML = h;
    // Si lo elegido ya no existe (ej. se redujo la matriz) queda en "— variable —" y validar() avisa
    const sigue = opciones.some(function (o) { return o.valor === elegido; });
    select.value = sigue ? elegido : "";
}

/* Refresca todos los menús de variables */
function actualizarSelectsVariables() {
    const opciones = opcionesVariables();
    document.querySelectorAll("#restriccionesExtra .var").forEach(function (s) { llenarSelect(s, opciones); });
}

/* Añade "coef · xij" a una restricción */
function agregarTermino(contenedor) {
    const span = document.createElement("span");
    span.className = "termino me-2 mb-1";
    span.innerHTML =
        '<span class="signo-mas text-muted">+</span>' +
        '<input type="text" inputmode="decimal" class="form-control form-control-sm coef" style="width:64px" value="1">' +
        '<select class="form-select form-select-sm var" style="width:auto"></select>' +
        '<button type="button" class="btn btn-sm btn-outline-secondary btnQuitarTermino" title="Quitar variable">×</button>';
    llenarSelect(span.querySelector(".var"), opcionesVariables());
    span.querySelector(".btnQuitarTermino").addEventListener("click", function () {
        // Siempre se deja al menos una variable en la restricción
        if (contenedor.querySelectorAll(".termino").length > 1) {
            span.remove();
            renumerarRestricciones();
        }
    });
    contenedor.appendChild(span);
    renumerarRestricciones();
}

/* Fila completa: [términos] [+ variable] [signo] [valor] [Quitar] */
function agregarRestriccion() {
    const div = document.createElement("div");
    div.className = "restriccion border rounded bg-white p-2 mb-2 d-flex flex-wrap align-items-center gap-2";
    div.setAttribute("role", "group");
    div.innerHTML =
        '<div class="terminos d-flex flex-wrap align-items-center"></div>' +
        '<button type="button" class="btn btn-outline-primary btn-sm btnMasTermino">+ variable</button>' +
        '<select class="form-select form-select-sm signo" style="width:auto">' +
        '<option value="<=">≤</option><option value=">=">≥</option><option value="=">=</option></select>' +
        '<input type="text" inputmode="decimal" class="form-control form-control-sm valor" style="width:90px" placeholder="valor">' +
        '<button type="button" class="btn btn-outline-danger btn-sm btnQuitar">Quitar</button>';
    const cont = div.querySelector(".terminos");
    div.querySelector(".btnMasTermino").addEventListener("click", function () { agregarTermino(cont); });
    div.querySelector(".btnQuitar").addEventListener("click", function () {
        div.remove(); // aquí sí se pueden quitar todas: son opcionales
        renumerarRestricciones();
    });
    document.getElementById("restriccionesExtra").appendChild(div);
    agregarTermino(cont);
}

/* Etiquetas de accesibilidad y oculta el "+" del primer término */
function renumerarRestricciones() {
    document.querySelectorAll("#restriccionesExtra .restriccion").forEach(function (r, k) {
        const n = k + 1;
        r.setAttribute("aria-label", "Restricción adicional " + n);
        r.querySelectorAll(".termino").forEach(function (t, q) {
            t.querySelector(".signo-mas").style.display = q === 0 ? "none" : "";
            t.querySelector(".coef").setAttribute("aria-label", "Coeficiente " + (q + 1) + ", restricción " + n);
            t.querySelector(".var").setAttribute("aria-label", "Variable " + (q + 1) + ", restricción " + n);
        });
        r.querySelector(".signo").setAttribute("aria-label", "Signo, restricción " + n);
        r.querySelector(".valor").setAttribute("aria-label", "Valor, restricción " + n);
        r.querySelector(".btnQuitar").setAttribute("aria-label", "Quitar restricción " + n);
    });
}

//-----------------------------------------------------------------
// * Paso 1:lectura
// -----------------------------------------------------------------

/* Todo el formulario como TEXTO en un solo objeto */
function leerEntrada() {
    const extras = [];
    document.querySelectorAll("#restriccionesExtra .restriccion").forEach(function (r) {
        const terminos = [];
        r.querySelectorAll(".termino").forEach(function (t) {
            terminos.push({ coef: t.querySelector(".coef").value, var: t.querySelector(".var").value });
        });
        extras.push({ terminos: terminos, signo: r.querySelector(".signo").value, valor: r.querySelector(".valor").value });
    });
    return {
        tipo: document.getElementById("tipo").value,
        objetivo: document.getElementById("objetivo").value,
        grupos: grupos(),
        matriz: leerMatriz(),
        extras: extras,
    };
}

//-----------------------------------------------------------------
// * Paso 2: validar y convertir a números
// -----------------------------------------------------------------

/* validar(entrada): revisa todo y convierte a números. Devuelve {errores:[...], datos:{tipo, objetivo, m, n, nombresFilas, nombresCols, etiquetasFilas, etiquetasCols, costos[m][n], oferta[m], demanda[n], extras}} */
function validar(entrada) {
    const errores = [];
    const mat = entrada.matriz;
    const g = entrada.grupos;
    const esAsig = entrada.tipo === "asignacion";

    const nombresFilas = [], nombresCols = [], etiquetasFilas = [], etiquetasCols = [];
    for (let i = 0; i < mat.m; i++) { nombresFilas.push(g.fila + " " + (i + 1)); etiquetasFilas.push(String(i + 1)); }
    for (let j = 0; j < mat.n; j++) { nombresCols.push(g.col + " " + (j + 1)); etiquetasCols.push(String(j + 1)); }

    // Costos: todos deben ser números
    const costos = mat.costos.map(function (fila, i) {
        return fila.map(function (texto, j) {
            const c = toNumber(texto);
            if (!isFinite(c)) errores.push("Costo de " + nombreVar(i, j) + " (" + nombresFilas[i] + " → " + nombresCols[j] + "): vacío o no numérico.");
            return c;
        });
    });

    // Oferta y demanda: en asignación siempre 1; en transporte números ≥ 0
    let oferta, demanda;
    if (esAsig) {
        oferta = nombresFilas.map(function () { return 1; });
        demanda = nombresCols.map(function () { return 1; });
    } else {
        oferta = mat.oferta.map(function (v, i) {
            const x = toNumber(v);
            if (!isFinite(x) || x < 0) errores.push("Oferta de " + nombresFilas[i] + ": debe ser un número mayor o igual a 0.");
            return x;
        });
        demanda = mat.demanda.map(function (v, j) {
            const x = toNumber(v);
            if (!isFinite(x) || x < 0) errores.push("Demanda de " + nombresCols[j] + ": debe ser un número mayor o igual a 0.");
            return x;
        });
    }

    // Restricciones adicionales
    const extras = [];
    entrada.extras.forEach(function (r, k) {
        const n = k + 1;
        const terminos = [];
        r.terminos.forEach(function (t) {
            const coef = toNumber(t.coef);
            if (!isFinite(coef)) errores.push("Restricción adicional " + n + ": hay un coeficiente vacío o no numérico.");
            if (t.var === "") { errores.push("Restricción adicional " + n + ": falta elegir una variable."); return; }
            const partes = t.var.split(",").map(Number); // "2,2" → [2, 2]
            terminos.push({ i: partes[0], j: partes[1], coef: coef });
        });
        const valor = toNumber(r.valor);
        if (!isFinite(valor)) errores.push("Restricción adicional " + n + ": el valor está vacío o no es numérico.");
        extras.push({ terminos: terminos, signo: r.signo, valor: valor });
    });

    return {
        errores: errores,
        datos: {
            tipo: entrada.tipo, objetivo: entrada.objetivo, m: mat.m, n: mat.n, grupos: g,
            nombresFilas: nombresFilas, nombresCols: nombresCols, etiquetasFilas: etiquetasFilas, etiquetasCols: etiquetasCols,
            costos: costos, oferta: oferta, demanda: demanda, extras: extras,
        },
    };
}

//-----------------------------------------------------------------
// * Paso 3: balancear (filas / columnas ficticias)
// -----------------------------------------------------------------

/* balancear(d): si el problema no está balanceado, agrega ficticias con costo 0. Devuelve una COPIA de d con m, n, nombres, etiquetas, costos, oferta y demanda ampliados, filaFicticia[] y colFicticia[] (true/false por fila/columna) y avisos[] (texto). */
function balancear(d) {
    // Copias para no modificar los datos originales
    const b = {
        tipo: d.tipo, objetivo: d.objetivo, m: d.m, n: d.n, grupos: d.grupos, extras: d.extras,
        nombresFilas: d.nombresFilas.slice(), nombresCols: d.nombresCols.slice(),
        etiquetasFilas: d.etiquetasFilas.slice(), etiquetasCols: d.etiquetasCols.slice(),
        costos: d.costos.map(function (fila) { return fila.slice(); }),
        oferta: d.oferta.slice(), demanda: d.demanda.slice(),
        filaFicticia: d.nombresFilas.map(function () { return false; }),
        colFicticia: d.nombresCols.map(function () { return false; }),
        holgura: false, // true = transporte con oferta > demanda: filas con "≤" en vez de columna ficticia
        avisos: [],
    };

    // agregarFila / agregarCol: añaden una fila o columna de costo 0. La etiqueta corta es "F", "F2"...
    function agregarFila(oferta) {
        const k = b.filaFicticia.filter(Boolean).length + 1; // cuántas ficticias van + 1
        b.costos.push(new Array(b.n).fill(0));
        b.nombresFilas.push(b.grupos.fila + " ficticio" + (k > 1 ? " " + k : ""));
        b.etiquetasFilas.push("F" + (k > 1 ? k : ""));
        b.oferta.push(oferta); b.filaFicticia.push(true); b.m++;
    }
    function agregarCol(demanda) {
        const k = b.colFicticia.filter(Boolean).length + 1;
        b.costos.forEach(function (fila) { fila.push(0); });
        b.nombresCols.push(b.grupos.col + " ficticio" + (k > 1 ? " " + k : ""));
        b.etiquetasCols.push("F" + (k > 1 ? k : ""));
        b.demanda.push(demanda); b.colFicticia.push(true); b.n++;
    }

    if (d.tipo === "transporte") {
        const sumaO = d.oferta.reduce(function (s, x) { return s + x; }, 0);
        const sumaD = d.demanda.reduce(function (s, x) { return s + x; }, 0);
        const dif = sumaO - sumaD;
        if (dif > EPS) {
            // No se agrega columna: las restricciones de oferta se vuelven "≤" y lo que no se envía es holgura
            b.holgura = true;
            b.avisos.push("Oferta total (" + fmt(sumaO) + ") > demanda total (" + fmt(sumaD) + "): las restricciones de oferta se usan como <strong>≤</strong>. Lo que no se envía queda como <strong>holgura</strong> (" + fmt(dif) + " en total).");
        } else if (dif < -EPS) {
            agregarFila(-dif);
            b.avisos.push("Demanda total (" + fmt(sumaD) + ") > oferta total (" + fmt(sumaO) + "): se agregó una <strong>fila ficticia (F)</strong> con oferta " + fmt(-dif) + " y costo 0. Lo que envíe es demanda que no se cubre.");
        }
    } else {
        // Asignación: se completan filas o columnas hasta tener n × n
        while (b.m < b.n) agregarFila(1);
        while (b.n < b.m) agregarCol(1);
        if (b.m !== d.m) b.avisos.push("Hay más columnas que filas: se agregaron " + (b.m - d.m) + " <strong>fila(s) ficticia(s) (F)</strong> con costo 0. La columna que les toque queda sin asignar.");
        if (b.n !== d.n) b.avisos.push("Hay más filas que columnas: se agregaron " + (b.n - d.n) + " <strong>columna(s) ficticia(s) (F)</strong> con costo 0. La fila que la reciba queda sin asignar.");
    }
    return b;
}

//-----------------------------------------------------------------
// * Paso 4: construir el modelo (texto para mostrar + objeto para la librería)
// -----------------------------------------------------------------

/* [{coef, nombre}] -> "20x11 + 12x12 − 3x13" (omite el coeficiente 1) */
function textoSuma(terminos) {
    if (terminos.length === 0) return "0";
    return terminos.map(function (t, k) {
        const abs = Math.abs(t.coef);
        const cuerpo = (abs === 1 ? "" : fmt(abs)) + t.nombre;
        if (k === 0) return (t.coef < 0 ? "−" : "") + cuerpo;
        return (t.coef < 0 ? " − " : " + ") + cuerpo;
    }).join("");
}

/* variables que una restricción adicional obliga a valer 0 (una sola variable con "= 0" o "≤ 0", como xij ≥ 0, en ambos casos xij = 0). Devuelve un Set con "i,j". Solo se usa para el TEXTO, el cálculo lo hace la propia restricción. */
function fijadasEnCero(extras) {
    const set = new Set();
    extras.forEach(function (r) {
        if (r.terminos.length !== 1 || r.valor !== 0) return;   // solo restricciones de una variable con valor 0
        const t = r.terminos[0];
        if (t.coef === 0) return;                               // 0·xij = 0 no dice nada de xij
        if (r.signo === "=" || (r.signo === "<=" && t.coef > 0)) set.add(t.i + "," + t.j);
    });
    return set;
}

/* construirModelo(b): recibe los datos ya balanceados/sin errores.
   Devuelve {variables:[{i,j,clave,nombre,costo}], texto:{variables, fo, auto[], manual[]}, lp} */
function construirModelo(b) {
    const esAsig = b.tipo === "asignacion";
    const SIGNO = { "<=": "≤", ">=": "≥", "=": "=" };
    // Cómo se llama cada tipo de restricción en el formato de la librería
    const LIMITE = { "<=": "max", ">=": "min", "=": "equal" };

    // Modelo para javascript-lp-solver:
    //  optimize: nombre del atributo que se optimiza ("z")
    //  constraints: cada restricción con su límite;  variables: cada xij con su aporte a cada restricción
    const lp = { optimize: "z", opType: b.objetivo, constraints: {}, variables: {} };
    if (esAsig) lp.binaries = {}; // en asignación las variables solo pueden valer 0 o 1

    // 1) Variables: una por celda (incluidas las ficticias)
    const variables = [];
    for (let i = 0; i < b.m; i++) {
        for (let j = 0; j < b.n; j++) {
            variables.push({ i: i, j: j, clave: "x" + i + "_" + j, nombre: nombreVar(i, j), costo: b.costos[i][j] });
        }
    }

    // Cada variable aporta: su costo a "z", 1 a su fila ("f"+i) y 1 a su columna ("c"+j)
    variables.forEach(function (v) {
        lp.variables[v.clave] = { z: v.costo };
        lp.variables[v.clave]["f" + v.i] = 1;
        lp.variables[v.clave]["c" + v.j] = 1;
        if (esAsig) lp.binaries[v.clave] = 1;
    });

    const auto = [];   // restricciones generadas automáticamente (texto)
    const manual = []; // restricciones adicionales del usuario (texto)

    // 2) Por fila: Σ fila = oferta (en asignación, = 1)
    for (let i = 0; i < b.m; i++) {
        // Con holgura: Σ fila ≤ oferta ("max" en la librería), sin holgura: Σ fila = oferta
        lp.constraints["f" + i] = b.holgura ? { max: b.oferta[i] } : { equal: b.oferta[i] };
        const t = variables.filter(function (v) { return v.i === i; }).map(function (v) { return { coef: 1, nombre: v.nombre }; });
        auto.push(textoSuma(t) + (b.holgura ? " ≤ " : " = ") + fmt(b.oferta[i]));
    }
    // 3) Por columna: Σ columna = demanda (en asignación, = 1)
    for (let j = 0; j < b.n; j++) {
        lp.constraints["c" + j] = { equal: b.demanda[j] };
        const t = variables.filter(function (v) { return v.j === j; }).map(function (v) { return { coef: 1, nombre: v.nombre }; });
        auto.push(textoSuma(t) + " = " + fmt(b.demanda[j]));
    }
    // 4) No negatividad / binarias (solo texto: la librería ya asume xij ≥ 0)
    auto.push(esAsig ? "xij ∈ {0, 1}" : "xij ≥ 0");

    // 5) Restricciones adicionales del usuario ("e"+k)
    b.extras.forEach(function (r, k) {
        const clave = "e" + k;
        lp.constraints[clave] = {};
        lp.constraints[clave][LIMITE[r.signo]] = r.valor;
        r.terminos.forEach(function (t) {
            const v = lp.variables["x" + t.i + "_" + t.j];
            // "+=" por si el usuario eligió la misma variable dos veces en la restricción
            v[clave] = (v[clave] || 0) + t.coef;
        });
        const t = r.terminos.map(function (t) { return { coef: t.coef, nombre: nombreVar(t.i, t.j) }; });
        manual.push(textoSuma(t) + " " + SIGNO[r.signo] + " " + fmt(r.valor));
    });

    // Función objetivo: sin las variables con costo 0 ni las fijadas en 0 (no aportan nada a Z)
    const cero = fijadasEnCero(b.extras);
    const enFO = variables.filter(function (v) { return v.costo !== 0 && !cero.has(v.i + "," + v.j); });
    const fo = (b.objetivo === "min" ? "Min" : "Max") + " Z = " +
        textoSuma(enFO.map(function (v) { return { coef: v.costo, nombre: v.nombre }; }));

    // Texto de las variables: todas, más notas de las que no aparecen en Z
    const txtVar = esAsig
        ? "x<sub>ij</sub> = 1 si " + escapar(b.grupos.fila.toLowerCase()) + " i se asigna a " + escapar(b.grupos.col.toLowerCase()) + " j; 0 si no."
        : "x<sub>ij</sub> = unidades que se envían de " + escapar(b.grupos.fila.toLowerCase()) + " i a " + escapar(b.grupos.col.toLowerCase()) + " j.";
    let textoVariables = "<p class='mb-1'>" + txtVar + " i = 1…" + b.m + ", j = 1…" + b.n +
        ": <strong>" + variables.length + " variables</strong></p>" +
        "<p class='font-monospace small text-break bg-white border rounded p-2 mb-1'>" +
        variables.map(function (v) { return v.nombre; }).join(", ") + "</p>";
    const notas = [];
    if (cero.size) notas.push("Fijadas en 0 por restricción (no disponibles): " + Array.from(cero).map(function (k) {
        const p = k.split(",").map(Number); return nombreVar(p[0], p[1]);
    }).join(", ") + ".");
    if (variables.some(function (v) { return v.costo === 0; })) notas.push("Las variables con costo 0 (como las ficticias) no aparecen en la función objetivo porque no aportan a Z.");
    if (notas.length) textoVariables += "<p class='small text-secondary mb-0'>Nota: " + notas.join(" ") + "</p>";

    return { variables: variables, texto: { variables: textoVariables, fo: fo, auto: auto, manual: manual }, lp: lp };
}

//-----------------------------------------------------------------
// * Paso 5: resolver con javascript-lp-solver
// -----------------------------------------------------------------

/* resolver(modelo): función que usa la librería y devuelve {estado:'optimo'|'infactible'|'no_acotado'|'error', z, valores:[{...variable, valor}], mensaje} */
function resolver(modelo) {
    // typeof evita un error si la librería no cargó 
    if (typeof solver === "undefined") {
        return { estado: "error", mensaje: "No se pudo cargar la librería del solver. Revisa tu conexión a internet y recarga la página." };
    }
    const r = solver.Solve(modelo.lp);
    // Si no es factible, la librería igual devuelve números "basura"; por eso se revisa primero
    if (!r.feasible) return { estado: "infactible" };
    if (r.bounded === false) return { estado: "no_acotado" };

    // La librería solo devuelve las variables distintas de 0, las demás valen 0
    const valores = modelo.variables.map(function (v) {
        let x = r[v.clave] || 0;
        if (Math.abs(x) < EPS) x = 0; // quita "basura" de redondeo tipo 1e-12
        x = Math.round(x * 10000) / 10000;    // máximo 4 decimales
        return Object.assign({}, v, { valor: x }); // copia de la variable + su valor
    });
    // Z se recalcula con los valores (Σ costo · valor), así coincide con lo que se ve en pantalla
    const z = valores.reduce(function (s, v) { return s + v.costo * v.valor; }, 0);
    return { estado: "optimo", z: z, valores: valores };
}

//-----------------------------------------------------------------
// * Paso 6: mostrar resultados
// -----------------------------------------------------------------

/* "Escapa" el texto y, si es ficticio, lo pone en gris cursiva */
function etiqueta(texto, ficticia) {
    return ficticia ? '<span class="ficticia">' + escapar(texto) + "</span>" : escapar(texto);
}

function mostrarResultados(b, modelo, sol) {
    const esAsig = b.tipo === "asignacion";
    const zTxt = "Z " + (b.objetivo === "min" ? "mínima" : "máxima") + " = " + fmt(sol.z);

    // Alerta principal con Z
    document.getElementById("resultadoAlerta").innerHTML =
        '<div class="alert alert-success"><strong>Óptimo:</strong> ' + zTxt +
        '<br><span class="small">Puede existir otra solución con el mismo valor de Z, el solver muestra una de ellas.</span></div>';

    // Avisos de resultantes en el balanceo
    document.getElementById("avisos").innerHTML = b.avisos.length
        ? '<div class="alert alert-info small">' + b.avisos.join("<br>") + "</div>" : "";

    // 1, 2, 3: el modelo. Restricciones: primero las automáticas, luego (separadas por una línea) las manuales
    document.getElementById("modeloVariables").innerHTML = modelo.texto.variables;
    document.getElementById("modeloFO").textContent = modelo.texto.fo;
    let h = modelo.texto.auto.map(function (t) { return "<div>" + escapar(t) + "</div>"; }).join("");
    if (modelo.texto.manual.length) {
        h += '<hr class="my-2">' + modelo.texto.manual.map(function (t) { return "<div>" + escapar(t) + "</div>"; }).join("");
    }
    document.getElementById("modeloRestricciones").innerHTML = h;

    // Matriz de valores: valor[i][j]
    const valor = [];
    for (let i = 0; i < b.m; i++) valor.push(new Array(b.n).fill(0));
    sol.valores.forEach(function (v) { valor[v.i][v.j] = v.valor; });

    // 4: matriz de solución (misma forma que la de entrada) con totales
    h = '<thead><tr><th colspan="2" rowspan="2"></th><th colspan="' + b.n + '" class="grupo">' + escapar(b.grupos.col) + "</th>" +
        '<th rowspan="2" class="zona-oferta align-middle">Total / ' + (esAsig ? "1" : "Oferta") + "</th>" +
        (b.holgura ? '<th rowspan="2" class="zona-oferta align-middle">Holgura</th>' : "") + "</tr><tr>";
    for (let j = 0; j < b.n; j++) h += "<th>" + etiqueta(b.etiquetasCols[j], b.colFicticia[j]) + "</th>";
    h += "</tr></thead><tbody>";
    for (let i = 0; i < b.m; i++) {
        h += "<tr>";
        if (i === 0) h += '<th rowspan="' + b.m + '" class="grupo align-middle">' + escapar(b.grupos.fila) + "</th>";
        h += "<th>" + etiqueta(b.etiquetasFilas[i], b.filaFicticia[i]) + "</th>";
        let total = 0;
        for (let j = 0; j < b.n; j++) {
            const x = valor[i][j];
            total += x;
            h += "<td" + (x > 0 ? ' class="fila-optima"' : "") + '><span class="d-block small text-muted">' + nombreVar(i, j) + "</span>" + fmt(x) + "</td>";
        }
        h += '<td class="zona-oferta">' + fmt(total) + " / " + fmt(b.oferta[i]) + "</td>";
        // Holgura = oferta − lo que se envió (redondeado para no mostrar basura tipo 1e-12)
        if (b.holgura) h += '<td class="zona-oferta">' + fmt(Math.round((b.oferta[i] - total) * 10000) / 10000) + "</td>";
        h += "</tr>";
    }
    h += '<tr><th colspan="2" class="zona-demanda">Total / ' + (esAsig ? "1" : "Demanda") + "</th>";
    for (let j = 0; j < b.n; j++) {
        let total = 0;
        for (let i = 0; i < b.m; i++) total += valor[i][j];
        h += '<td class="zona-demanda">' + fmt(total) + " / " + fmt(b.demanda[j]) + "</td>";
    }
    document.getElementById("tablaSolucion").innerHTML = h + "<td></td>" + (b.holgura ? "<td></td>" : "") + "</tr></tbody>";

    // 5: tabla de variables (todas; el checkbox oculta las que valen 0)
    h = "<thead><tr><th>Variable</th><th>" + (esAsig ? "Asignación" : "Ruta") + "</th><th class='text-end'>Valor</th><th class='text-end'>" +
        (b.objetivo === "min" ? "Costo" : "Beneficio") + "</th><th class='text-end'>Aporte a Z</th></tr></thead><tbody>";
    sol.valores.forEach(function (v) {
        h += '<tr class="' + (v.valor > 0 ? "fila-optima" : "fila-cero") + '"><td>' + v.nombre + "</td><td>" +
            etiqueta(b.nombresFilas[v.i], b.filaFicticia[v.i]) + " → " + etiqueta(b.nombresCols[v.j], b.colFicticia[v.j]) +
            "</td><td class='text-end'>" + fmt(v.valor) + "</td><td class='text-end'>" + fmt(v.costo) +
            "</td><td class='text-end'>" + fmt(v.costo * v.valor) + "</td></tr>";
    });
    h += "<tr><th colspan='4' class='text-end'>Z</th><th class='text-end'>" + fmt(sol.z) + "</th></tr>";
    document.getElementById("tablaVariables").innerHTML = h + "</tbody>";
    filtrarVariables();

    // 6: solución final: cada variable distinta de 0 en una caja (.solucion) + su interpretación
    const palabra = b.objetivo === "min" ? "costo" : "beneficio"; // en Max, el "costo" es en realidad ganancia
    const lineas = [];
    sol.valores.forEach(function (v) {
        if (v.valor <= 0) return;
        const caja = '<span class="solucion">' + v.nombre + " = " + fmt(v.valor) + "</span> ";
        const fila = escapar(b.nombresFilas[v.i]), col = escapar(b.nombresCols[v.j]);
        let texto;
        if (b.filaFicticia[v.i]) {
            texto = esAsig ? "<strong>" + col + "</strong> queda sin asignar (le tocó una fila ficticia)."
                : "<strong>" + col + "</strong>: " + fmt(v.valor) + " unidades de demanda no se cubren con las filas reales.";
        } else if (b.colFicticia[v.j]) {
            texto = esAsig ? "<strong>" + fila + "</strong> queda sin asignar (le tocó una columna ficticia)."
                : "<strong>" + fila + "</strong>: sobran " + fmt(v.valor) + " unidades (no se envían).";
        } else {
            texto = esAsig ? fila + " → " + col + ", " + palabra + " " + fmt(v.costo) + "."
                : "Enviar " + fmt(v.valor) + " de " + fila + " a " + col + ", " + palabra + " " + fmt(v.costo * v.valor) + ".";
        }
        lineas.push('<li class="mb-2">' + caja + texto + "</li>");
    });
    // Con holgura: capacidad sin usar de cada fila (oferta − enviado)
    if (b.holgura) {
        for (let i = 0; i < b.m; i++) {
            const enviado = valor[i].reduce(function (s, x) { return s + x; }, 0);
            const holgura = Math.round((b.oferta[i] - enviado) * 10000) / 10000;
            if (holgura > 0) lineas.push('<li class="mb-2"><strong>' + escapar(b.nombresFilas[i]) + "</strong>: holgura de " + fmt(holgura) + " (capacidad sin usar).</li>");
        }
    }
    document.getElementById("interpretacion").innerHTML =
        '<ul class="list-unstyled">' + lineas.join("") + "</ul>" +
        '<p class="mt-3"><span class="solucion solucion-z">' + zTxt + "</span></p>" +
        '<p class="small text-secondary">Todas las variables que no aparecen aquí valen 0.</p>';
}

/* Oculta (d-none) las filas con valor 0 si el checkbox está marcado */
function filtrarVariables() {
    const solo = document.getElementById("soloNoCero").checked;
    document.querySelectorAll("#tablaVariables .fila-cero").forEach(function (tr) { tr.classList.toggle("d-none", solo); });
}

//-----------------------------------------------------------------
// * Calcular
// -----------------------------------------------------------------

function mostrarErrores(lista) {
    document.getElementById("alertas").innerHTML = '<div class="alert alert-danger"><ul class="mb-0">' +
        lista.map(function (e) { return "<li>" + escapar(e) + "</li>"; }).join("") + "</ul></div>";
    document.getElementById("resultados").style.display = "none";
}

function calcular() {
    document.getElementById("alertas").innerHTML = "";
    const v = validar(leerEntrada());  // pasos 1 y 2
    if (v.errores.length > 0) return mostrarErrores(v.errores);
    const b = balancear(v.datos);   // paso 3
    const modelo = construirModelo(b);  // paso 4
    const sol = resolver(modelo);  // paso 5
    if (sol.estado === "error") return mostrarErrores([sol.mensaje]);
    if (sol.estado === "infactible") return mostrarErrores(["INFACTIBLE: no existe una solución que cumpla todas las restricciones. Revisa las restricciones adicionales (por ejemplo, una fila o columna con todas sus variables en 0)."]);
    if (sol.estado === "no_acotado") return mostrarErrores(["NO ACOTADO: Z puede mejorar sin límite. Revisa los datos."]);
    document.getElementById("resultados").style.display = "block";
    mostrarResultados(b, modelo, sol);  // paso 6
}

//-----------------------------------------------------------------
// * Inicio (eventos)
// -----------------------------------------------------------------

if (typeof document !== "undefined") {
    dibujarMatriz({}); // matriz por defecto de 3 × 3
    document.getElementById("tipo").addEventListener("change", function () { redibujar(true); });
    document.getElementById("numFilas").addEventListener("change", function () { redibujar(false); });
    document.getElementById("numCols").addEventListener("change", function () { redibujar(false); });
    document.getElementById("btnAgregarRestriccion").addEventListener("click", agregarRestriccion);
    document.getElementById("soloNoCero").addEventListener("change", filtrarVariables);

    // La tabla se redibuja, así que se escucha en la tabla (que no cambia)
    // Al cambiar el nombre de un grupo, se actualizan los textos de los menús de variables
    document.getElementById("tablaMatriz").addEventListener("change", function (ev) {
        if (ev.target.id === "grupoFilas" || ev.target.id === "grupoCols") actualizarSelectsVariables();
    });

    document.getElementById("formulario").addEventListener("submit", function (evento) {
        evento.preventDefault(); // evita que el navegador recargue la página
        calcular();
    });
}