/**
 * ============================================================
 *  MAMUSHKA – Sistema de gestión
 * ============================================================
 *  Módulos:
 *   - Menú: ventas de hoy (clickeable) + créditos pendientes
 *     (clickeable) + accesos a los 4 formularios
 *   - Nuevo crédito: prendas dinámicas (hasta 5, tipo+descripción)
 *   - Registrar pago: total o parcial, con recargo/extensión
 *   - Cargar venta: prendas dinámicas con precio, guarda directo
 *     y opcionalmente manda WhatsApp después
 *     sin guardar nada
 *   - Registrar gasto: solo gastos FIJOS del negocio, con
 *     consolidado del mes actual clickeable
 *   - Registrar compra: control de inversión en mercadería
 *   - Historial: libro diario de movimientos (ver Historial.gs)
 *
 *  Hojas (se crean solas):
 *   - "Creditos", "Clientes", "Ventas", "Gastos", "Pagos", "Compras"
 *
 *  NOTA: el módulo Stock fue retirado por el momento. Si tu planilla
 *  todavía tiene una hoja "Stock", no se borra ni se modifica: el
 *  sistema simplemente ya no la usa.
 *
 *  REGLA DE RECARGO / EXTENSIÓN EN CRÉDITOS (sin cambios):
 *   Un solo ajuste por crédito: o 7 días de gracia sin recargo
 *   (pagó a tiempo pero quedó saldo), o 10% de recargo una vez
 *   si pagó tarde. Nunca los dos, nunca más de una vez.
 * ============================================================
 */

var NOMBRE_HOJA_CREDITOS = "Creditos";
var NOMBRE_HOJA_CLIENTES = "Clientes";
var NOMBRE_HOJA_VENTAS = "Ventas";
var NOMBRE_HOJA_GASTOS = "Gastos";
var NOMBRE_HOJA_PAGOS = "Pagos";
var NOMBRE_HOJA_COMPRAS = "Compras";
var ALIAS_TRANSFERENCIA = "Mamushka.10";

var MESES = ["enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"];
var CATEGORIAS_GASTO = ["🏠 Alquiler","💡 Luz","🔥 Gas","📋 Monotributo","📄 Habilitación","🌐 Internet","🧑‍💼 Ayudante","🧮 Contador","🔧 Otros"];
var TIPOS_PRENDA = ["Jean","Remera","Campera","Buzo/Sweater","Vestido","Pantalón","Camisa","Tapado/Blazer","Calzado","Lencería","Pijama","Packs","Pantuflas","Chaleco","Corpiño","Body","Top","Pollera/Falda","Accesorios","Otro"];

// Columnas de Creditos (1-indexed). Prenda4/5 van al final para no
// romper el orden de datos ya cargados con versiones anteriores.
// ID_CREDITO también va al final por el mismo motivo: es la clave
// única para identificar CADA crédito puntual (nunca se busca ni se
// actualiza un crédito por nombre o por teléfono, siempre por este ID).
var COL = {
  FECHA_CARGA: 1, NOMBRE: 2, APELLIDO: 3, CELULAR: 4,
  PRENDA1: 5, PRENDA2: 6, PRENDA3: 7,
  PRECIO_TOTAL: 8, ENTREGA: 9, SALDO: 10, VENCIMIENTO: 11,
  AJUSTE: 12, ESTADO: 13, FECHA_PAGO_TOTAL: 14, METODO_PAGO: 15, LINK: 16,
  PRENDA4: 17, PRENDA5: 18, CUOTAS: 19, TARJETA: 20, ID_CREDITO: 21,
  // Prendas 6 en adelante (texto, separadas por " | "). Las primeras 5
  // siguen en sus columnas de siempre; esto evita el límite de columnas.
  EXTRAS: 22
};

// Columnas de Ventas (1-indexed)
var COL_VENTA = {
  FECHA: 1, NOMBRE: 2, APELLIDO: 3, CELULAR: 4,
  TIPO1: 5, DESC1: 6, PRECIO1: 7,
  TIPO2: 8, DESC2: 9, PRECIO2: 10,
  TIPO3: 11, DESC3: 12, PRECIO3: 13,
  TIPO4: 14, DESC4: 15, PRECIO4: 16,
  TOTAL: 17, METODO_PAGO: 18, LINK: 19, CUOTAS: 20, TARJETA: 21,
  // Prendas 5 en adelante (texto, separadas por " | "): detalle y precio.
  // Las primeras 4 siguen en sus columnas de siempre.
  EXTRAS: 22
};

// Columnas de Gastos (1-indexed)
var COL_GASTO = { FECHA: 1, CATEGORIA: 2, MONTO: 3, OBSERVACIONES: 4 };

// Columnas de Pagos (1-indexed) — registro de dinero que entra: entregas de crédito y pagos
var COL_PAGO = { FECHA: 1, CLIENTE: 2, TIPO: 3, MONTO: 4, METODO: 5, CUOTAS: 6, TARJETA: 7 };

// Columnas de Compras (1-indexed) — registro histórico de mercadería
// comprada. No calcula ganancias: es solo control de cuánto dinero
// se invirtió (valor de la compra + transporte).
var COL_COMPRA = {
  ID: 1, FECHA: 2, PROVEEDOR: 3, VALOR: 4, TRANSPORTE: 5,
  INVERSION_TOTAL: 6, OBSERVACIONES: 7
};

function doGet(e) {
  asegurarTodasLasHojas();
  return HtmlService.createHtmlOutputFromFile('Formulario')
    .setTitle('Mamushka')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Se asegura de que las 6 hojas del sistema existan (con sus encabezados)
 * apenas se abre la app o la planilla, en vez de crearse recién la primera
 * vez que se usa cada sección.
 * También migra cualquier crédito viejo que todavía no tenga ID único.
 */
function asegurarTodasLasHojas() {
  obtenerHojaCreditos();
  obtenerHojaClientes();
  obtenerHojaVentas();
  obtenerHojaGastos();
  obtenerHojaPagos();
  obtenerHojaCompras();
  asegurarIdsCreditos();
}

/**
 * Se dispara solo al abrir la PLANILLA (no la app web). Además de
 * preparar las hojas, agrega un menú con herramientas de mantenimiento
 * para revisar/fusionar clientes duplicados directamente desde Sheets.
 */
function onOpen() {
  asegurarTodasLasHojas();
  SpreadsheetApp.getUi()
    .createMenu('🧹 Mantenimiento')
    .addItem('Revisar clientes duplicados', 'revisarClientesDuplicados')
    .addItem('Fusionar clientes duplicados', 'limpiarClientesDuplicados')
    .addToUi();
}

/* =========================================================
 *  MENÚ
 * ========================================================= */

function obtenerDatosMenu() {
  var hoja = obtenerHojaCreditos();
  var datos = hoja.getDataRange().getValues();
  var hoy = new Date();

  var total = 0;
  var clientes = [];

  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    if (fila[COL.ESTADO - 1] === "Pagado" || fila[COL.ESTADO - 1] === "") continue;

    var saldo = Number(fila[COL.SALDO - 1]) || 0;
    var vencimiento = new Date(fila[COL.VENCIMIENTO - 1]);
    var ajuste = fila[COL.AJUSTE - 1];
    var vencido = estaVencido(vencimiento, hoy);

    var saldoMostrar = saldo;
    if (vencido && ajuste === "No") {
      saldoMostrar = redondear2(saldo * 1.10);
    }

    var color = "verde";
    var diasParaVencer = Math.ceil((vencimiento - hoy) / (1000 * 60 * 60 * 24));
    if (vencido) color = "rojo";
    else if (diasParaVencer <= 3) color = "amarillo";

    total += saldoMostrar;

    var prendas = prendasDeCredito(fila);

    clientes.push({
      idCredito: fila[COL.ID_CREDITO - 1],
      nombreCompleto: fila[COL.NOMBRE - 1] + " " + fila[COL.APELLIDO - 1],
      saldo: formatearMoneda(saldoMostrar),
      color: color,
      vencido: vencido,
      prendas: prendas.join(", "),
      fechaVencimiento: Utilities.formatDate(vencimiento, Session.getScriptTimeZone(), "dd/MM/yyyy"),
      precioTotal: formatearMoneda(Number(fila[COL.PRECIO_TOTAL - 1]) || 0),
      entregaInicial: formatearMoneda(Number(fila[COL.ENTREGA - 1]) || 0)
    });
  }

  clientes.sort(function (a, b) {
    var orden = { rojo: 0, amarillo: 1, verde: 2 };
    return orden[a.color] - orden[b.color];
  });

  var ventasHoy = calcularVentasHoy();
  var ventasMes = calcularVentasMes();

  return {
    totalPendiente: formatearMoneda(total),
    cantidad: clientes.length,
    clientes: clientes,
    ventasHoy: ventasHoy,
    ventasMes: ventasMes
  };
}

function calcularVentasHoy() {
  var resultadoVentas = sumarPeriodo(obtenerHojaVentas(), 'venta', true);
  var resultadoPagos = sumarPeriodo(obtenerHojaPagos(), 'pago', true);

  return {
    total: formatearMoneda(resultadoVentas.total + resultadoPagos.total),
    cantidad: resultadoVentas.detalle.length + resultadoPagos.detalle.length,
    detalle: resultadoVentas.detalle.concat(resultadoPagos.detalle)
  };
}

function calcularVentasMes() {
  var resultadoVentas = sumarPeriodo(obtenerHojaVentas(), 'venta', false);
  var resultadoPagos = sumarPeriodo(obtenerHojaPagos(), 'pago', false);

  return {
    total: formatearMoneda(resultadoVentas.total + resultadoPagos.total),
    cantidad: resultadoVentas.detalle.length + resultadoPagos.detalle.length,
    detalle: resultadoVentas.detalle.concat(resultadoPagos.detalle)
  };
}

/**
 * Suma los movimientos de una hoja (Ventas o Pagos) dentro del período
 * pedido (hoy o el mes actual), devolviendo total + detalle para mostrar.
 */
function sumarPeriodo(hoja, origen, soloHoy) {
  var datos = hoja.getDataRange().getValues();
  var hoy = new Date();
  var total = 0;
  var detalle = [];

  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    var fecha, monto, cliente, descripcion;

    if (origen === 'venta') {
      fecha = new Date(fila[COL_VENTA.FECHA - 1]);
      if (isNaN(fecha.getTime())) continue;

      // Una venta con "Crédito personal" NO representa dinero que entró
      // hoy: el total de esa venta es el valor de la mercadería, no lo
      // cobrado. Lo único que efectivamente ingresó (la entrega, y más
      // adelante cada pago de saldo) ya se contabiliza por separado a
      // través de la hoja Pagos. Sumar también esta fila duplicaría esa
      // plata (contaría la venta completa Y la entrega/pago por su lado).
      if (fila[COL_VENTA.METODO_PAGO - 1] === "Crédito personal") continue;

      monto = Number(fila[COL_VENTA.TOTAL - 1]) || 0;
      var nombre = (fila[COL_VENTA.NOMBRE - 1] + " " + fila[COL_VENTA.APELLIDO - 1]).trim();
      cliente = nombre || "Sin nombre";
      var tipos = tiposDeVenta(fila);
      descripcion = tipos.join(", ") || "Venta";
    } else {
      fecha = new Date(fila[COL_PAGO.FECHA - 1]);
      if (isNaN(fecha.getTime())) continue;
      monto = Number(fila[COL_PAGO.MONTO - 1]) || 0;
      cliente = fila[COL_PAGO.CLIENTE - 1];
      descripcion = fila[COL_PAGO.TIPO - 1];
    }

    var mismoMesYAnio = fecha.getMonth() === hoy.getMonth() && fecha.getFullYear() === hoy.getFullYear();
    if (!mismoMesYAnio) continue;
    if (soloHoy && fecha.getDate() !== hoy.getDate()) continue;

    total += monto;
    detalle.push({
      cliente: cliente,
      productos: descripcion,
      monto: formatearMoneda(monto),
      fecha: Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM")
    });
  }

  return { total: total, detalle: detalle };
}

/* =========================================================
 *  CRÉDITO PERSONAL — se carga desde "Cargar venta" eligiendo ese
 *  método de pago (ya no existe como formulario/botón aparte).
 *  Reutiliza las prendas ya cargadas en la venta y además guarda
 *  fila en Ventas (para el consolidado) y en Creditos (para el
 *  seguimiento de saldo/vencimiento/recargo, sin tocar esa lógica).
 * ========================================================= */

function guardarVentaCreditoPersonal(datos) {
  requerido(datos.nombre, "Nombre");
  requerido(datos.apellido, "Apellido");
  requerido(datos.celular, "Celular");
  validarCelular(datos.celular);
  requerido(datos.fechaVencimiento, "Fecha de vencimiento");

  var prendas = datos.prendas || [];
  if (prendas.length === 0) {
    throw new Error("Cargá al menos una prenda con precio.");
  }

  var total = 0;
  for (var i = 0; i < prendas.length; i++) total += parsearMonto(prendas[i].precio);
  total = redondear2(total);

  var montoEntrega = parsearMonto(datos.montoEntrega);
  noNegativo(montoEntrega, "El monto de entrega");
  if (montoEntrega > total) {
    throw new Error("El monto de entrega no puede ser mayor al total de la venta.");
  }

  var saldoPendiente = redondear2(total - montoEntrega);

  var fechaVencimiento = new Date(datos.fechaVencimiento + "T00:00:00");
  var fechaFormateada = Utilities.formatDate(fechaVencimiento, Session.getScriptTimeZone(), "dd/MM/yyyy");

  var textosPrendas = prendas.map(function (p) {
    return p.descripcion ? (p.tipo + " (" + p.descripcion + ")") : p.tipo;
  });
  var textoPrendas = textosPrendas.length > 1
    ? textosPrendas.slice(0, -1).join(", ") + " y " + textosPrendas[textosPrendas.length - 1]
    : textosPrendas[0];

  var mensaje =
    "Hola " + datos.nombre + "! 🌸 Muchas gracias por tu compra en Mamushka.\n\n" +
    "Quedó registrada tu " + (textosPrendas.length > 1 ? "compra de " : "compra: ") + textoPrendas +
    ", por un total de $" + formatearMoneda(total) + ". " +
    "Hoy dejaste $" + formatearMoneda(montoEntrega) + ", así que te queda un saldo de $" + formatearMoneda(saldoPendiente) + " 💕\n\n" +
    "La fecha límite para abonarlo es el " + fechaFormateada + ". Pasada esa fecha se suma un recargo del 10% sobre el saldo, así que ¡no te olvides! 😉\n\n" +
    "Podés transferir a nuestro alias: *" + ALIAS_TRANSFERENCIA + "* 💳\n\n" +
    "Cualquier duda escribime por acá. ¡Que la disfrutes! 🌸";

  var celularFormateado = limpiarCelularArgentina(datos.celular);
  var urlWhatsApp = "https://wa.me/" + celularFormateado + "?text=" + encodeURIComponent(mensaje);

  // ---- Fila en Creditos (saldo, vencimiento, recargo — igual que antes) ----
  var hojaCreditos = obtenerHojaCreditos();
  var filaCredito = [];
  filaCredito[COL.FECHA_CARGA - 1] = new Date();
  filaCredito[COL.NOMBRE - 1] = datos.nombre;
  filaCredito[COL.APELLIDO - 1] = datos.apellido;
  filaCredito[COL.CELULAR - 1] = datos.celular;

  var slots = [COL.PRENDA1, COL.PRENDA2, COL.PRENDA3, COL.PRENDA4, COL.PRENDA5];
  for (var s = 0; s < 5; s++) {
    filaCredito[slots[s] - 1] = prendas[s] ? (prendas[s].descripcion ? prendas[s].tipo + " - " + prendas[s].descripcion : prendas[s].tipo) : "";
  }

  filaCredito[COL.EXTRAS - 1] = textoExtrasPrendas(prendas, 5, false);
  filaCredito[COL.PRECIO_TOTAL - 1] = total;
  filaCredito[COL.ENTREGA - 1] = montoEntrega;
  filaCredito[COL.SALDO - 1] = saldoPendiente;
  filaCredito[COL.VENCIMIENTO - 1] = fechaVencimiento;
  filaCredito[COL.AJUSTE - 1] = "No";
  filaCredito[COL.ESTADO - 1] = "Pendiente";
  filaCredito[COL.FECHA_PAGO_TOTAL - 1] = "";
  filaCredito[COL.METODO_PAGO - 1] = datos.metodoPagoEntrega || "";
  filaCredito[COL.LINK - 1] = urlWhatsApp;
  filaCredito[COL.CUOTAS - 1] = datos.cuotasEntrega || "";
  filaCredito[COL.TARJETA - 1] = datos.tarjetaEntrega || "";

  // Generar el ID único y agregar la fila es una sola operación
  // "atómica" bajo lock, para que dos créditos cargados al mismo
  // tiempo nunca puedan terminar con el mismo ID.
  var idCredito = conLockDeEscritura(function () {
    var id = generarNuevoIdCredito(hojaCreditos);
    filaCredito[COL.ID_CREDITO - 1] = id;
    hojaCreditos.appendRow(filaCredito);
    return id;
  });

  // ---- Fila en Ventas (para que sume en el consolidado de ventas) ----
  var hojaVentas = obtenerHojaVentas();
  var filaVenta = [];
  filaVenta[COL_VENTA.FECHA - 1] = new Date();
  filaVenta[COL_VENTA.NOMBRE - 1] = datos.nombre;
  filaVenta[COL_VENTA.APELLIDO - 1] = datos.apellido;
  filaVenta[COL_VENTA.CELULAR - 1] = datos.celular;

  var slotsTipo = [COL_VENTA.TIPO1, COL_VENTA.TIPO2, COL_VENTA.TIPO3, COL_VENTA.TIPO4];
  var slotsDesc = [COL_VENTA.DESC1, COL_VENTA.DESC2, COL_VENTA.DESC3, COL_VENTA.DESC4];
  var slotsPrecio = [COL_VENTA.PRECIO1, COL_VENTA.PRECIO2, COL_VENTA.PRECIO3, COL_VENTA.PRECIO4];
  for (var j = 0; j < 4; j++) {
    filaVenta[slotsTipo[j] - 1] = prendas[j] ? (prendas[j].tipo || "") : "";
    filaVenta[slotsDesc[j] - 1] = prendas[j] ? (prendas[j].descripcion || "") : "";
    filaVenta[slotsPrecio[j] - 1] = prendas[j] ? parsearMonto(prendas[j].precio) : "";
  }
  filaVenta[COL_VENTA.EXTRAS - 1] = textoExtrasPrendas(prendas, 4, true);
  filaVenta[COL_VENTA.TOTAL - 1] = total;
  filaVenta[COL_VENTA.METODO_PAGO - 1] = "Crédito personal";
  filaVenta[COL_VENTA.LINK - 1] = urlWhatsApp;
  filaVenta[COL_VENTA.CUOTAS - 1] = "";
  filaVenta[COL_VENTA.TARJETA - 1] = "";

  hojaVentas.appendRow(filaVenta);

  actualizarClienteUnico(datos.celular, datos.nombre, datos.apellido);
  registrarMovimientoDinero(datos.nombre + " " + datos.apellido, "Entrega crédito", montoEntrega, datos.metodoPagoEntrega, datos.cuotasEntrega, datos.tarjetaEntrega);

  return {
    urlWhatsApp: urlWhatsApp,
    idCredito: idCredito,
    nombre: datos.nombre,
    total: formatearMoneda(total),
    saldoPendiente: formatearMoneda(saldoPendiente),
    fechaFormateada: fechaFormateada
  };
}

/* =========================================================
 *  FORM 2 — REGISTRAR PAGO (sin cambios de lógica)
 * ========================================================= */

function obtenerClientesConSaldo() {
  var hoja = obtenerHojaCreditos();
  var datos = hoja.getDataRange().getValues();
  var hoy = new Date();
  var tz = Session.getScriptTimeZone();
  var resultado = [];

  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    if (fila[COL.ESTADO - 1] === "Pagado" || fila[COL.ESTADO - 1] === "") continue;

    var saldo = Number(fila[COL.SALDO - 1]) || 0;
    var vencimiento = new Date(fila[COL.VENCIMIENTO - 1]);
    var ajuste = fila[COL.AJUSTE - 1];
    var vencido = estaVencido(vencimiento, hoy);
    var saldoMostrar = (vencido && ajuste === "No") ? redondear2(saldo * 1.10) : saldo;
    var fechaCarga = new Date(fila[COL.FECHA_CARGA - 1]);

    resultado.push({
      // Identificador único del crédito puntual — es lo único que se
      // usa después para aplicar el pago. Nunca celular, nunca nombre:
      // un mismo cliente puede tener más de un crédito pendiente a la
      // vez, y dos clientes distintos podrían compartir/errar un
      // número de celular.
      idCredito: fila[COL.ID_CREDITO - 1],
      celular: fila[COL.CELULAR - 1],
      nombreCompleto: fila[COL.NOMBRE - 1] + " " + fila[COL.APELLIDO - 1],
      saldoMostrar: formatearMoneda(saldoMostrar),
      vencido: vencido,
      // Datos para desglosar valor base + interés en pantalla y
      // recalcularlo si se elige otra fecha de pago.
      saldoBase: saldo,
      ajuste: ajuste,
      vencimiento: Utilities.formatDate(vencimiento, tz, "yyyy-MM-dd"),
      fechaCarga: isNaN(fechaCarga.getTime()) ? "" : Utilities.formatDate(fechaCarga, tz, "yyyy-MM-dd")
    });
  }
  return resultado;
}

function procesarPago(datos) {
  var montoPagado = parsearMonto(datos.montoPagado);
  if (montoPagado <= 0) {
    throw new Error("Ingresá un monto mayor a cero.");
  }

  var idCredito = String(datos.idCredito || "").trim();
  if (!idCredito) {
    throw new Error("No se pudo identificar el crédito a pagar. Volvé a seleccionar el cliente e intentá de nuevo.");
  }

  var ahora = new Date();
  // Fecha en la que realmente se cobró (por defecto hoy). Permite cargar
  // pagos retroactivos: el ingreso queda en el día/mes correcto.
  var fechaPago = parsearFechaPago(datos.fechaPago, ahora);
  var tz = Session.getScriptTimeZone();

  return conLockDeEscritura(function () {
    var hoja = obtenerHojaCreditos();
    var rango = hoja.getDataRange();
    var valores = rango.getValues();
    var filaIndex = -1;

    // Búsqueda EXCLUSIVAMENTE por ID de crédito (clave única de la
    // fila). Nunca por celular ni por nombre: son datos que se pueden
    // repetir entre clientes o entre créditos del mismo cliente, y
    // aplicar un pago a la fila equivocada corrompe el saldo de otra
    // persona.
    for (var i = 1; i < valores.length; i++) {
      if (String(valores[i][COL.ID_CREDITO - 1] || "").trim() === idCredito) {
        filaIndex = i;
        break;
      }
    }

    if (filaIndex === -1) {
      throw new Error("No se encontró ese crédito (puede que ya se haya saldado). Actualizá la lista de clientes e intentá de nuevo.");
    }

    var fila = valores[filaIndex];
    if (fila[COL.ESTADO - 1] === "Pagado") {
      throw new Error("Ese crédito ya está saldado.");
    }

    var fechaCarga = new Date(fila[COL.FECHA_CARGA - 1]);
    if (!isNaN(fechaCarga.getTime()) && claveDiaFecha(fechaPago) < claveDiaFecha(fechaCarga)) {
      throw new Error("La fecha de pago no puede ser anterior a la fecha en que se cargó el crédito (" +
        Utilities.formatDate(fechaCarga, tz, "dd/MM/yyyy") + ").");
    }

    var nombre = fila[COL.NOMBRE - 1];
    var saldoActual = Number(fila[COL.SALDO - 1]) || 0;   // valor base
    var fechaVencimiento = new Date(fila[COL.VENCIMIENTO - 1]);
    var ajuste = fila[COL.AJUSTE - 1];
    // El vencimiento se evalúa contra la FECHA DE PAGO (no contra hoy):
    // un pago cargado con fecha anterior al vencimiento no lleva recargo.
    var vencido = estaVencido(fechaVencimiento, fechaPago);

    var interes = 0;
    var nuevaFechaVencimiento = fechaVencimiento;
    var unaSemana = 7 * 24 * 60 * 60 * 1000;

    if (ajuste === "No") {
      if (vencido) {
        interes = redondear2(saldoActual * 0.10);
        ajuste = "Recargo";
        nuevaFechaVencimiento = new Date(fechaPago.getTime() + unaSemana);
      } else if (redondear2(saldoActual - montoPagado) > 0) {
        ajuste = "Extensión";
        nuevaFechaVencimiento = new Date(fechaPago.getTime() + unaSemana);
      }
    }

    var totalAdeudado = redondear2(saldoActual + interes);

    if (montoPagado > totalAdeudado + 0.01) {
      throw new Error("El monto pagado ($" + formatearMoneda(montoPagado) + ") no puede ser mayor al saldo pendiente ($" + formatearMoneda(totalAdeudado) + ").");
    }

    var nuevoSaldo = redondear2(totalAdeudado - montoPagado);
    var estadoNuevo, fechaPagoTotal;

    if (nuevoSaldo <= 0.01) {
      nuevoSaldo = 0;
      estadoNuevo = "Pagado";
      fechaPagoTotal = Utilities.formatDate(fechaPago, tz, "dd/MM/yyyy");
    } else {
      estadoNuevo = "Parcial";
      fechaPagoTotal = "";
    }

    var fechaVencFormateada = Utilities.formatDate(nuevaFechaVencimiento, tz, "dd/MM/yyyy");

    var mensaje;
    if (estadoNuevo === "Pagado") {
      mensaje =
        "Hola " + nombre + "! 🌸 Recibimos tu pago de $" + formatearMoneda(montoPagado) + ". " +
        "¡Tu crédito quedó saldado por completo, gracias por confiar en Mamushka! 💕";
    } else {
      mensaje =
        "Hola " + nombre + "! 🌸 Recibimos tu pago de $" + formatearMoneda(montoPagado) + ". " +
        "Te queda un saldo de $" + formatearMoneda(nuevoSaldo) + ", con fecha límite " + fechaVencFormateada + ".\n\n" +
        "Podés transferir a nuestro alias: *" + ALIAS_TRANSFERENCIA + "* 💳\n\n" +
        "¡Gracias! 🌸";
    }

    var celularFormateado = limpiarCelularArgentina(fila[COL.CELULAR - 1]);
    var urlWhatsApp = "https://wa.me/" + celularFormateado + "?text=" + encodeURIComponent(mensaje);

    var numeroFila = filaIndex + 1;
    // COL.SALDO..COL.LINK son 7 columnas contiguas (10 a 16): las actualizamos
    // todas juntas en una sola llamada en vez de 7 llamadas separadas — mismo
    // resultado, pero mucho más rápido porque es un solo viaje a Sheets.
    hoja.getRange(numeroFila, COL.SALDO, 1, 7).setValues([[
      nuevoSaldo, nuevaFechaVencimiento, ajuste, estadoNuevo, fechaPagoTotal, datos.metodoPago || "", urlWhatsApp
    ]]);
    registrarMovimientoDinero(nombre, "Pago crédito", montoPagado, datos.metodoPago, datos.cuotas, datos.tarjeta, fechaPago);

    return {
      urlWhatsApp: urlWhatsApp,
      nombre: nombre,
      estado: estadoNuevo,
      valorBase: formatearMoneda(saldoActual),
      interes: formatearMoneda(interes),
      hayInteres: interes > 0,
      totalAdeudado: formatearMoneda(totalAdeudado),
      montoPagado: formatearMoneda(montoPagado),
      fechaPago: Utilities.formatDate(fechaPago, tz, "dd/MM/yyyy"),
      saldoRestante: formatearMoneda(nuevoSaldo)
    };
  });
}

/* =========================================================
 *  FORM 3 — CARGAR VENTA
 * ========================================================= */

function guardarVenta(datos) {
  var prendas = datos.prendas || [];
  if (prendas.length === 0) {
    throw new Error("Cargá al menos una prenda con precio.");
  }

  var hoja = obtenerHojaVentas();

  var total = 0;
  for (var i = 0; i < prendas.length; i++) total += parsearMonto(prendas[i].precio);
  total = redondear2(total);

  var fila = [];
  fila[COL_VENTA.FECHA - 1] = new Date();
  fila[COL_VENTA.NOMBRE - 1] = "";
  fila[COL_VENTA.APELLIDO - 1] = "";
  fila[COL_VENTA.CELULAR - 1] = "";

  var slotsTipo = [COL_VENTA.TIPO1, COL_VENTA.TIPO2, COL_VENTA.TIPO3, COL_VENTA.TIPO4];
  var slotsDesc = [COL_VENTA.DESC1, COL_VENTA.DESC2, COL_VENTA.DESC3, COL_VENTA.DESC4];
  var slotsPrecio = [COL_VENTA.PRECIO1, COL_VENTA.PRECIO2, COL_VENTA.PRECIO3, COL_VENTA.PRECIO4];

  for (var j = 0; j < 4; j++) {
    fila[slotsTipo[j] - 1] = prendas[j] ? (prendas[j].tipo || "") : "";
    fila[slotsDesc[j] - 1] = prendas[j] ? (prendas[j].descripcion || "") : "";
    fila[slotsPrecio[j] - 1] = prendas[j] ? parsearMonto(prendas[j].precio) : "";
  }

  fila[COL_VENTA.EXTRAS - 1] = textoExtrasPrendas(prendas, 4, true);
  fila[COL_VENTA.TOTAL - 1] = total;
  fila[COL_VENTA.METODO_PAGO - 1] = datos.metodoPago || "";
  fila[COL_VENTA.LINK - 1] = "";
  fila[COL_VENTA.CUOTAS - 1] = datos.cuotas || "";
  fila[COL_VENTA.TARJETA - 1] = datos.tarjeta || "";

  hoja.appendRow(fila);
  var filaIndex = hoja.getLastRow();

  return { filaIndex: filaIndex, total: formatearMoneda(total) };
}

function enviarComprobanteVenta(datos) {
  var hoja = obtenerHojaVentas();
  var fila = hoja.getRange(datos.filaIndex, 1, 1, hoja.getLastColumn()).getValues()[0];

  var prendas = [];
  var slotsTipo = [COL_VENTA.TIPO1, COL_VENTA.TIPO2, COL_VENTA.TIPO3, COL_VENTA.TIPO4];
  var slotsDesc = [COL_VENTA.DESC1, COL_VENTA.DESC2, COL_VENTA.DESC3, COL_VENTA.DESC4];
  var slotsPrecio = [COL_VENTA.PRECIO1, COL_VENTA.PRECIO2, COL_VENTA.PRECIO3, COL_VENTA.PRECIO4];
  for (var i = 0; i < 4; i++) {
    if (fila[slotsTipo[i] - 1]) {
      prendas.push({ tipo: fila[slotsTipo[i] - 1], descripcion: fila[slotsDesc[i] - 1], precio: fila[slotsPrecio[i] - 1] });
    }
  }
  var total = Number(fila[COL_VENTA.TOTAL - 1]) || 0;
  var metodoPago = fila[COL_VENTA.METODO_PAGO - 1];
  var cuotas = fila[COL_VENTA.CUOTAS - 1];
  var tarjeta = fila[COL_VENTA.TARJETA - 1];

  var mensaje = armarMensajeVenta(datos, prendas, total, metodoPago, cuotas, tarjeta);
  var celularFormateado = limpiarCelularArgentina(datos.celular);
  var urlWhatsApp = "https://wa.me/" + celularFormateado + "?text=" + encodeURIComponent(mensaje);

  hoja.getRange(datos.filaIndex, COL_VENTA.NOMBRE).setValue(datos.nombre);
  hoja.getRange(datos.filaIndex, COL_VENTA.APELLIDO).setValue(datos.apellido || "");
  hoja.getRange(datos.filaIndex, COL_VENTA.CELULAR).setValue(datos.celular);
  hoja.getRange(datos.filaIndex, COL_VENTA.LINK).setValue(urlWhatsApp);

  actualizarClienteUnico(datos.celular, datos.nombre, datos.apellido || "");

  return { urlWhatsApp: urlWhatsApp };
}

function armarMensajeVenta(datos, prendas, total, metodoPago, cuotas, tarjeta) {
  var lineas = prendas.map(function (p) {
    return "🔹 " + p.tipo + (p.descripcion ? " — " + p.descripcion : "") + ": $" + formatearMoneda(parsearMonto(p.precio));
  }).join("\n");

  var lineaPago = "";
  if (metodoPago) {
    lineaPago = "\n💳 Pagaste con: " + metodoPago;
    if (metodoPago === "Tarjeta de crédito" && cuotas) {
      lineaPago += " (" + cuotas + (Number(cuotas) === 1 ? " cuota" : " cuotas") + (tarjeta ? " — " + tarjeta : "") + ")";
    }
  }

  return "Hola " + datos.nombre + "! 🌸 Gracias por tu compra en Mamushka.\n\n" +
    lineas + "\n\n💰 Total: $" + formatearMoneda(total) +
    lineaPago + "\n\n¡Que lo disfrutes! 🌸";
}

/* =========================================================
 *  FORM 4 — REGISTRAR GASTO (fijos, con consolidado del mes)
 * ========================================================= */

function obtenerResumenGastosMes() {
  var hoja = obtenerHojaGastos();
  var datos = hoja.getDataRange().getValues();
  var hoy = new Date();

  var total = 0;
  var categorias = {};

  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    var fecha = new Date(fila[COL_GASTO.FECHA - 1]);
    if (isNaN(fecha.getTime())) continue;
    if (fecha.getMonth() !== hoy.getMonth() || fecha.getFullYear() !== hoy.getFullYear()) continue;

    var monto = Number(fila[COL_GASTO.MONTO - 1]) || 0;
    var categoria = fila[COL_GASTO.CATEGORIA - 1];
    total += monto;
    categorias[categoria] = (categorias[categoria] || 0) + monto;
  }

  var detalle = Object.keys(categorias).map(function (cat) {
    return { categoria: cat, monto: formatearMoneda(categorias[cat]) };
  });

  return {
    mesNombre: MESES[hoy.getMonth()],
    total: formatearMoneda(total),
    detalle: detalle
  };
}

function registrarGasto(datos) {
  requerido(datos.categoria, "Categoría");
  var monto = parsearMonto(datos.monto);
  if (monto <= 0) {
    throw new Error("Ingresá un monto mayor a cero.");
  }

  var hoja = obtenerHojaGastos();
  hoja.appendRow([new Date(), datos.categoria, monto, datos.observaciones || ""]);
  ordenarGastosPorFechaDesc(hoja);
  actualizarResumenGastos();

  return obtenerResumenGastosMes();
}

function ordenarGastosPorFechaDesc(hoja) {
  var ultimaFila = hoja.getLastRow();
  if (ultimaFila <= 2) return; // nada que ordenar con 0 o 1 gasto
  var rango = hoja.getRange(2, 1, ultimaFila - 1, hoja.getLastColumn());
  rango.sort({ column: COL_GASTO.FECHA, ascending: false });
}

/**
 * Reescribe la hoja "ResumenGastos" con el total de cada categoría
 * por mes del año actual, para poder ver de un vistazo cómo viene
 * cada mes y compararlos a fin de año.
 */
function actualizarResumenGastos() {
  var planilla = SpreadsheetApp.getActiveSpreadsheet();
  var hojaResumen = planilla.getSheetByName("ResumenGastos");
  if (!hojaResumen) {
    hojaResumen = planilla.insertSheet("ResumenGastos");
  } else {
    hojaResumen.clear();
  }

  var hojaGastos = obtenerHojaGastos();
  var datos = hojaGastos.getDataRange().getValues();
  var anioActual = new Date().getFullYear();

  // matriz [mes][categoria] = total
  var totalesPorMes = [];
  for (var m = 0; m < 12; m++) {
    totalesPorMes.push({});
  }

  for (var i = 1; i < datos.length; i++) {
    var fila = datos[i];
    var fecha = new Date(fila[COL_GASTO.FECHA - 1]);
    if (isNaN(fecha.getTime()) || fecha.getFullYear() !== anioActual) continue;

    var mes = fecha.getMonth();
    var categoria = fila[COL_GASTO.CATEGORIA - 1];
    var monto = Number(fila[COL_GASTO.MONTO - 1]) || 0;

    totalesPorMes[mes][categoria] = (totalesPorMes[mes][categoria] || 0) + monto;
  }

  var encabezado = ["Mes"].concat(CATEGORIAS_GASTO, ["Total"]);
  var filas = [encabezado];

  for (var mesIdx = 0; mesIdx < 12; mesIdx++) {
    var totalMes = 0;
    var fila = [MESES[mesIdx].charAt(0).toUpperCase() + MESES[mesIdx].slice(1) + " " + anioActual];
    CATEGORIAS_GASTO.forEach(function (cat) {
      var monto = totalesPorMes[mesIdx][cat] || 0;
      totalMes += monto;
      fila.push(monto);
    });
    fila.push(totalMes);
    filas.push(fila);
  }

  hojaResumen.getRange(1, 1, filas.length, encabezado.length).setValues(filas);
  hojaResumen.getRange(1, 1, 1, encabezado.length).setFontWeight("bold");
}

/* =========================================================
 *  UTILIDADES
 * ========================================================= */

function obtenerHojaPagos() {
  var planilla = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = planilla.getSheetByName(NOMBRE_HOJA_PAGOS);

  if (!hoja) {
    hoja = planilla.insertSheet(NOMBRE_HOJA_PAGOS);
    hoja.appendRow(["Fecha", "Cliente", "Tipo", "Monto", "Método de pago", "Cuotas", "Tarjeta"]);
  }
  return hoja;
}

function registrarMovimientoDinero(cliente, tipo, monto, metodoPago, cuotas, tarjeta, fecha) {
  if (!monto || monto <= 0) return;
  var hoja = obtenerHojaPagos();
  // "fecha" permite cargar pagos retroactivos: el ingreso cuenta en el
  // día/mes en que realmente se cobró, no en el que se registró.
  hoja.appendRow([fecha || new Date(), cliente, tipo, monto, metodoPago || "", cuotas || "", tarjeta || ""]);
}

/* =========================================================
 *  REGISTRAR COMPRA — control histórico de inversión en
 *  mercadería (NO calcula ganancias: es solo control de cuánto
 *  se invirtió en cada compra).
 * ========================================================= */

function obtenerHojaCompras() {
  var planilla = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = planilla.getSheetByName(NOMBRE_HOJA_COMPRAS);

  if (!hoja) {
    hoja = planilla.insertSheet(NOMBRE_HOJA_COMPRAS);
    hoja.appendRow([
      "ID Compra", "Fecha", "Proveedor", "Valor de la compra",
      "Transporte", "Inversión total", "Observaciones"
    ]);
  }
  return hoja;
}

function generarIdCompra(hoja) {
  // El próximo número es la cantidad de filas ya usadas (el
  // encabezado ocupa la fila 1, así que con solo encabezado
  // getLastRow() = 1 y la primera compra es la Nº 1 → "C-001").
  var siguienteNumero = hoja.getLastRow();
  var numeroTexto = String(siguienteNumero);
  while (numeroTexto.length < 3) numeroTexto = "0" + numeroTexto;
  return "C-" + numeroTexto;
}

function registrarCompra(datos) {
  requerido(datos.fecha, "Fecha");
  requerido(datos.proveedor, "Proveedor");

  var valorCompra = parsearMonto(datos.valorCompra);
  if (valorCompra <= 0) {
    throw new Error("Ingresá el valor de la compra.");
  }

  var transporte = parsearMonto(datos.transporte);
  noNegativo(transporte, "El transporte");

  var inversionTotal = redondear2(valorCompra + transporte);
  var fecha = new Date(datos.fecha + "T00:00:00");

  var hoja = obtenerHojaCompras();
  var idCompra = generarIdCompra(hoja);

  var fila = [];
  fila[COL_COMPRA.ID - 1] = idCompra;
  fila[COL_COMPRA.FECHA - 1] = fecha;
  fila[COL_COMPRA.PROVEEDOR - 1] = datos.proveedor;
  fila[COL_COMPRA.VALOR - 1] = valorCompra;
  fila[COL_COMPRA.TRANSPORTE - 1] = transporte;
  fila[COL_COMPRA.INVERSION_TOTAL - 1] = inversionTotal;
  fila[COL_COMPRA.OBSERVACIONES - 1] = datos.observaciones || "";

  hoja.appendRow(fila);

  return {
    idCompra: idCompra,
    fechaFormateada: Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy"),
    proveedor: datos.proveedor,
    valorCompra: formatearMoneda(valorCompra),
    transporte: formatearMoneda(transporte),
    inversionTotal: formatearMoneda(inversionTotal)
  };
}

function obtenerHojaCreditos() {
  var planilla = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = planilla.getSheetByName(NOMBRE_HOJA_CREDITOS);

  if (!hoja) {
    hoja = planilla.insertSheet(NOMBRE_HOJA_CREDITOS);
    hoja.appendRow([
      "Fecha carga", "Nombre", "Apellido", "Celular",
      "Prenda 1", "Prenda 2", "Prenda 3",
      "Precio total", "Entrega inicial", "Saldo pendiente", "Fecha vencimiento",
      "Ajuste aplicado", "Estado", "Fecha pago total", "Método de pago", "Link WhatsApp",
      "Prenda 4", "Prenda 5", "Cuotas", "Tarjeta", "ID Crédito", "Prendas adicionales"
    ]);
  } else if (!hoja.getRange(1, COL.ID_CREDITO).getValue()) {
    // La hoja ya existía de antes de que existiera esta columna (o de
    // antes de que existiera este sistema de identificador único):
    // agregamos el encabezado sin tocar ninguna fila ni columna existente.
    hoja.getRange(1, COL.ID_CREDITO).setValue("ID Crédito");
  }
  asegurarEncabezado(hoja, COL.EXTRAS, "Prendas adicionales");
  return hoja;
}

/**
 * Ejecuta fn() con un lock exclusivo a nivel de todo el script. Sirve
 * para que dos ejecuciones al mismo tiempo (dos personas usando la
 * app a la vez, o un doble-toque accidental) nunca lean y escriban la
 * misma hoja en simultáneo, evitando IDs repetidos, clientes
 * duplicados o pagos aplicados dos veces.
 */
function conLockDeEscritura(fn) {
  var lock = LockService.getScriptLock();
  var conseguido = lock.tryLock(10000);
  if (!conseguido) {
    throw new Error("El sistema está procesando otro cambio, probá de nuevo en unos segundos.");
  }
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function formatearIdCredito(numero) {
  var s = String(numero);
  while (s.length < 3) s = "0" + s;
  return "CR-" + s;
}

function obtenerMaxNumeroCredito(hoja) {
  var ultimaFila = hoja.getLastRow();
  if (ultimaFila < 2) return 0;

  var ids = hoja.getRange(2, COL.ID_CREDITO, ultimaFila - 1, 1).getValues();
  var max = 0;
  for (var i = 0; i < ids.length; i++) {
    var match = String(ids[i][0] || "").match(/CR-(\d+)/);
    if (match) {
      var n = parseInt(match[1], 10);
      if (n > max) max = n;
    }
  }
  return max;
}

function generarNuevoIdCredito(hoja) {
  return formatearIdCredito(obtenerMaxNumeroCredito(hoja) + 1);
}

/**
 * Recorre TODA la hoja Creditos y le asigna un ID único a cualquier
 * fila que todavía no lo tenga (créditos cargados antes de que este
 * sistema de identificador único existiera). Se puede llamar todas
 * las veces que haga falta: las filas que ya tienen ID quedan intactas,
 * y solo se completan los huecos, siguiendo la numeración existente.
 */
function asegurarIdsCreditos() {
  conLockDeEscritura(function () {
    var hoja = obtenerHojaCreditos();
    var ultimaFila = hoja.getLastRow();
    if (ultimaFila < 2) return;

    var rango = hoja.getRange(2, COL.ID_CREDITO, ultimaFila - 1, 1);
    var ids = rango.getValues();
    var maxNumero = 0;
    var huecos = [];

    for (var i = 0; i < ids.length; i++) {
      var idActual = String(ids[i][0] || "").trim();
      var match = idActual.match(/CR-(\d+)/);
      if (match) {
        var n = parseInt(match[1], 10);
        if (n > maxNumero) maxNumero = n;
      } else if (!idActual) {
        huecos.push(i);
      }
    }

    if (huecos.length === 0) return;

    huecos.forEach(function (idx) {
      maxNumero++;
      ids[idx][0] = formatearIdCredito(maxNumero);
    });

    rango.setValues(ids);
  });
}

function obtenerHojaVentas() {
  var planilla = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = planilla.getSheetByName(NOMBRE_HOJA_VENTAS);

  if (!hoja) {
    hoja = planilla.insertSheet(NOMBRE_HOJA_VENTAS);
    hoja.appendRow([
      "Fecha", "Nombre", "Apellido", "Celular",
      "Tipo 1", "Descripción 1", "Precio 1",
      "Tipo 2", "Descripción 2", "Precio 2",
      "Tipo 3", "Descripción 3", "Precio 3",
      "Tipo 4", "Descripción 4", "Precio 4",
      "Total", "Método de pago", "Link WhatsApp", "Cuotas", "Tarjeta", "Prendas adicionales"
    ]);
  }
  asegurarEncabezado(hoja, COL_VENTA.EXTRAS, "Prendas adicionales");
  return hoja;
}

function obtenerHojaGastos() {
  var planilla = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = planilla.getSheetByName(NOMBRE_HOJA_GASTOS);

  if (!hoja) {
    hoja = planilla.insertSheet(NOMBRE_HOJA_GASTOS);
    hoja.appendRow(["Fecha", "Categoría", "Monto", "Observaciones"]);
  }
  return hoja;
}

function obtenerHojaClientes() {
  var planilla = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = planilla.getSheetByName(NOMBRE_HOJA_CLIENTES);

  if (!hoja) {
    hoja = planilla.insertSheet(NOMBRE_HOJA_CLIENTES);
    hoja.appendRow(["Celular", "Nombre", "Apellido", "Primer registro", "Último registro"]);
  }
  return hoja;
}

/**
 * Punto ÚNICO de creación/actualización de clientes. El celular es la
 * clave: primero busca si ya existe (comparando de forma robusta, ver
 * normalizarCelularParaComparar), y si existe reutiliza esa fila en
 * vez de crear una nueva. Todo el ciclo lectura+decisión+escritura va
 * bajo un lock exclusivo: sin el lock, dos ventas cargadas casi al
 * mismo tiempo para un cliente nuevo podrían las dos "no encontrar"
 * al cliente todavía inexistente y terminar creando dos filas para el
 * mismo teléfono.
 */
function actualizarClienteUnico(celular, nombre, apellido) {
  conLockDeEscritura(function () {
    var hojaClientes = obtenerHojaClientes();
    var datos = hojaClientes.getDataRange().getValues();
    var celularNormalizado = normalizarCelularParaComparar(celular);

    for (var i = 1; i < datos.length; i++) {
      var celularFila = normalizarCelularParaComparar(datos[i][0]);
      if (celularFila === celularNormalizado) {
        hojaClientes.getRange(i + 1, 2).setValue(nombre);
        hojaClientes.getRange(i + 1, 3).setValue(apellido);
        hojaClientes.getRange(i + 1, 5).setValue(new Date());
        return;
      }
    }

    hojaClientes.appendRow([celular, nombre, apellido, new Date(), new Date()]);
  });
}

/**
 * Agrupa todas las filas de Clientes por celular normalizado y
 * devuelve solo los grupos con más de una fila (duplicados reales).
 * No modifica nada, es solo lectura.
 */
function agruparClientesDuplicados() {
  var hoja = obtenerHojaClientes();
  var datos = hoja.getDataRange().getValues();
  var grupos = {};

  for (var i = 1; i < datos.length; i++) {
    var celular = datos[i][0];
    var clave = normalizarCelularParaComparar(celular);
    if (!clave) continue; // fila sin celular utilizable: no se agrupa

    if (!grupos[clave]) grupos[clave] = [];
    grupos[clave].push({
      fila: i + 1,
      celular: celular,
      nombre: datos[i][1],
      apellido: datos[i][2],
      primerRegistro: datos[i][3],
      ultimoRegistro: datos[i][4]
    });
  }

  var duplicados = [];
  Object.keys(grupos).forEach(function (clave) {
    if (grupos[clave].length > 1) duplicados.push(grupos[clave]);
  });
  return duplicados;
}

/**
 * Menú "🧹 Mantenimiento" en Sheets: solo INFORMA qué celulares tienen
 * más de una fila en Clientes. No modifica nada — pensada para correr
 * antes de "Fusionar clientes duplicados".
 */
function revisarClientesDuplicados() {
  var ui = SpreadsheetApp.getUi();
  var duplicados = agruparClientesDuplicados();

  if (duplicados.length === 0) {
    ui.alert("Sin duplicados", "No se encontraron clientes duplicados por celular. ✅", ui.ButtonSet.OK);
    return;
  }

  var lineas = duplicados.map(function (grupo) {
    var nombres = grupo.map(function (c) {
      return (c.nombre + " " + c.apellido).trim() + " (fila " + c.fila + ")";
    }).join(" / ");
    return grupo[0].celular + " → " + nombres;
  });

  ui.alert(
    "Clientes duplicados encontrados",
    "Se encontraron " + duplicados.length + " celular(es) con más de una fila:\n\n" +
    lineas.join("\n\n") +
    "\n\nUsá '🧹 Mantenimiento → Fusionar clientes duplicados' para unificarlos.",
    ui.ButtonSet.OK
  );
}

/**
 * Menú "🧹 Mantenimiento" en Sheets: fusiona cada grupo de duplicados
 * en UNA sola fila (conserva la de "Último registro" más reciente, con
 * el rango completo de fechas) y borra las filas sobrantes. Pide
 * confirmación antes, porque borrar filas es irreversible.
 */
function limpiarClientesDuplicados() {
  var ui = SpreadsheetApp.getUi();
  var duplicados = agruparClientesDuplicados();

  if (duplicados.length === 0) {
    ui.alert("Sin duplicados", "No se encontraron clientes duplicados por celular. Nada para fusionar. ✅", ui.ButtonSet.OK);
    return;
  }

  var totalFilasABorrar = duplicados.reduce(function (acc, g) { return acc + g.length - 1; }, 0);
  var respuesta = ui.alert(
    "Confirmar fusión",
    "Se encontraron " + duplicados.length + " celular(es) duplicados (" + totalFilasABorrar + " fila(s) de más). " +
    "Se va a conservar una sola fila por celular (con el nombre más reciente) y se van a BORRAR las filas sobrantes.\n\n" +
    "Esta acción no se puede deshacer. ¿Continuar?",
    ui.ButtonSet.YES_NO
  );

  if (respuesta !== ui.Button.YES) return;

  conLockDeEscritura(function () {
    var hoja = obtenerHojaClientes();
    var filasABorrar = [];

    duplicados.forEach(function (grupo) {
      var masReciente = grupo.reduce(function (a, b) {
        return new Date(b.ultimoRegistro) > new Date(a.ultimoRegistro) ? b : a;
      });
      var primerRegistroMinimo = grupo.reduce(function (a, b) {
        return new Date(b.primerRegistro) < new Date(a.primerRegistro) ? b : a;
      }).primerRegistro;

      hoja.getRange(masReciente.fila, 4).setValue(primerRegistroMinimo);

      grupo.forEach(function (c) {
        if (c.fila !== masReciente.fila) filasABorrar.push(c.fila);
      });
    });

    // De abajo hacia arriba para que borrar una fila no corra los
    // índices de las que todavía faltan borrar.
    filasABorrar.sort(function (a, b) { return b - a; });
    filasABorrar.forEach(function (fila) { hoja.deleteRow(fila); });

    ui.alert("Listo ✅", "Se fusionaron " + duplicados.length + " cliente(s) duplicado(s). Se borraron " + filasABorrar.length + " fila(s) sobrante(s).", ui.ButtonSet.OK);
  });
}

/**
 * Convierte el texto de un campo de monto al número real.
 * Formato argentino: "." separa miles y "," separa decimales
 * (ej: "1.250,50" -> 1250.5). Máximo 2 decimales.
 */
function parsearMonto(valor) {
  if (!valor) return 0;
  var limpio = String(valor).replace(/\./g, "").replace(",", ".");
  return redondear2(parseFloat(limpio) || 0);
}

/** Redondea a 2 decimales (centavos), evitando errores de coma flotante. */
function redondear2(numero) {
  var n = Number(numero);
  if (!isFinite(n)) return 0;
  return Math.round(n * 100 + (n >= 0 ? 1e-6 : -1e-6)) / 100;
}

/** Día como "yyyyMMdd" (para comparar fechas sin importar la hora). */
function claveDiaFecha(fecha) {
  return Utilities.formatDate(new Date(fecha), Session.getScriptTimeZone(), "yyyyMMdd");
}

/**
 * Un crédito está vencido recién DESPUÉS de su fecha límite (el mismo
 * día de vencimiento todavía se puede pagar sin recargo, como dice el
 * mensaje de WhatsApp: "Pasada esa fecha se suma un recargo").
 * Se compara por día, no por hora, contra la fecha de referencia.
 */
function estaVencido(vencimiento, fechaReferencia) {
  return claveDiaFecha(fechaReferencia) > claveDiaFecha(vencimiento);
}

/** Interpreta "YYYY-MM-DD" como fecha de pago (con la hora actual). */
function parsearFechaPago(iso, ahora) {
  if (!iso) return ahora;
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso));
  if (!m) throw new Error("La fecha de pago no es válida.");
  var f = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), ahora.getHours(), ahora.getMinutes(), ahora.getSeconds());
  if (f.getMonth() !== Number(m[2]) - 1) throw new Error("La fecha de pago no es válida.");
  if (claveDiaFecha(f) > claveDiaFecha(ahora)) throw new Error("La fecha de pago no puede ser futura.");
  return f;
}

/** Asegura que la hoja tenga la columna (y su encabezado) sin tocar datos. */
function asegurarEncabezado(hoja, columna, texto) {
  if (hoja.getLastColumn() >= columna) return;
  if (hoja.getMaxColumns() < columna) {
    hoja.insertColumnsAfter(hoja.getMaxColumns(), columna - hoja.getMaxColumns());
  }
  hoja.getRange(1, columna).setValue(texto);
}

/** Texto de las prendas que no entran en las columnas fijas. */
function textoExtrasPrendas(prendas, desde, conPrecio) {
  var items = [];
  for (var i = desde; i < prendas.length; i++) {
    var p = prendas[i];
    var t = p.tipo + (p.descripcion ? " - " + p.descripcion : "");
    if (conPrecio) t += " ($" + formatearMoneda(parsearMonto(p.precio)) + ")";
    items.push(t);
  }
  return items.join(" | ");
}

function extrasDeFila(fila, columna) {
  var t = String(fila[columna - 1] || "").trim();
  return t ? t.split(" | ") : [];
}

/** Tipos de prenda de una fila de Ventas (columnas fijas + adicionales). */
function tiposDeVenta(fila) {
  var base = [fila[COL_VENTA.TIPO1 - 1], fila[COL_VENTA.TIPO2 - 1], fila[COL_VENTA.TIPO3 - 1], fila[COL_VENTA.TIPO4 - 1]].filter(Boolean);
  var extras = extrasDeFila(fila, COL_VENTA.EXTRAS).map(function (e) {
    return e.split(" - ")[0].split(" ($")[0];
  });
  return base.concat(extras);
}

/** Prendas de una fila de Creditos (columnas fijas + adicionales). */
function prendasDeCredito(fila) {
  var base = [fila[COL.PRENDA1 - 1], fila[COL.PRENDA2 - 1], fila[COL.PRENDA3 - 1], fila[COL.PRENDA4 - 1], fila[COL.PRENDA5 - 1]].filter(Boolean);
  return base.concat(extrasDeFila(fila, COL.EXTRAS));
}

/**
 * Validaciones básicas del lado del servidor — no reemplazan las del
 * formulario, son una segunda barrera por si algo raro llega igual.
 */
function requerido(valor, nombreCampo) {
  if (!valor || String(valor).trim() === "") {
    throw new Error("Falta completar: " + nombreCampo + ".");
  }
}

function noNegativo(valor, nombreCampo) {
  if (Number(valor) < 0) {
    throw new Error(nombreCampo + " no puede ser negativo.");
  }
}

function validarCelular(celular) {
  var digitos = String(celular || "").replace(/\D/g, "");
  if (digitos.length < 6) {
    throw new Error("El número de celular parece incompleto.");
  }
}

/**
 * Formato argentino con MÁXIMO 2 decimales: miles con ".", decimales
 * con ",". Si el monto es entero no muestra decimales (1.250);
 * si tiene centavos muestra siempre 2 (1.250,50).
 */
function formatearMoneda(numero) {
  var n = redondear2(numero);
  var negativo = n < 0;
  var partes = Math.abs(n).toFixed(2).split(".");
  var entero = partes[0].replace(/\B(?=(\d{3})+(?!\d))/g, ".");
  var texto = partes[1] === "00" ? entero : entero + "," + partes[1];
  return (negativo ? "-" : "") + texto;
}

function limpiarCelularArgentina(celular) {
  var soloNumeros = String(celular).replace(/\D/g, "");

  if (soloNumeros.indexOf("0") === 0) {
    soloNumeros = soloNumeros.substring(1);
  }
  soloNumeros = soloNumeros.replace(/^(\d{2,4})15(\d+)$/, "$1$2");

  if (soloNumeros.indexOf("54") === 0) {
    if (soloNumeros.indexOf("549") !== 0) {
      soloNumeros = "549" + soloNumeros.substring(2);
    }
    return soloNumeros;
  }

  return "549" + soloNumeros;
}

/**
 * Normaliza un celular para COMPARAR si dos son "el mismo", no para
 * mostrarlo ni para armar un link. Reutiliza limpiarCelularArgentina
 * (que ya saca el 0 inicial, el 15, y homogeneiza el 549) para que
 * "011 15-1234-5678", "11 1234-5678" y "+54 9 11 1234-5678" —tres
 * formas típicas de escribir el MISMO número— den igual. Si el valor
 * es demasiado corto para ser un celular real, no se "canoniza" (para
 * no hacer que dos valores vacíos/basura parezcan el mismo cliente):
 * se compara tal cual, solo con los dígitos.
 */
function normalizarCelularParaComparar(celular) {
  var digitos = String(celular || "").replace(/\D/g, "");
  if (digitos.length < 6) return digitos;
  return limpiarCelularArgentina(celular);
}

/* =========================================================
 *  HISTORIAL DE MOVIMIENTOS / LIBRO DIARIO
 *  Devuelve el mes en curso (abierto) y los meses anteriores
 *  (para desplegables), cada uno agrupado por día, con lo último
 *  registrado primero.
 *   - 'venta'   -> ventas al contado (nombre opcional)
 *   - 'credito' -> crédito nuevo: seña + saldo a financiar
 *   - 'pago'    -> cobro de un crédito
 *   - 'gasto'   -> gasto fijo del negocio (sale plata)
 *   - 'compra'  -> compra de mercadería (sale plata)
 *  REGLAS PARA NO DUPLICAR PLATA:
 *   - Las ventas con método "Crédito personal" NO se listan como
 *     venta: ya aparecen como 'credito'.
 *   - Las filas "Entrega crédito" de Pagos van dentro de la tarjeta
 *     del crédito (como seña) pero SÍ suman a lo facturado.
 *   - "Entró" = plata que efectivamente entró: ventas al contado
 *     + cobros de crédito + señas (igual que el menú).
 *   - "Ganancia" del mes = Entró - Gastos fijos - Compras.
 *     Es una cuenta de CAJA: no incluye el saldo que los clientes
 *     todavía deben de créditos, ni el valor de la mercadería
 *     comprada que aún no se vendió.
 * ========================================================= */

var DIAS_SEMANA_ES = ["", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado", "Domingo"];

function obtenerHistorial() {
  var tz = Session.getScriptTimeZone();
  var ahora = new Date();

  function claveMes(f) { return Utilities.formatDate(f, tz, "yyyy-MM"); }
  function claveDia(f) { return Utilities.formatDate(f, tz, "yyyy-MM-dd"); }

  var movs = [];
  var ingresosPorDia = {};
  var ingresosPorMes = {};
  var creditosPorMes = {};
  var cantCreditosPorMes = {};
  var gastosPorDia = {};
  var gastosPorMes = {};
  var comprasPorMes = {};

  // Plata que sale (gastos y compras) por día y por mes.
  function sumarEgreso(f, monto, tipo) {
    var kd = claveDia(f), km = claveMes(f);
    gastosPorDia[kd] = (gastosPorDia[kd] || 0) + monto;
    if (tipo === "gasto") gastosPorMes[km] = (gastosPorMes[km] || 0) + monto;
    else comprasPorMes[km] = (comprasPorMes[km] || 0) + monto;
  }

  function sumarIngreso(f, monto) {
    var kd = claveDia(f), km = claveMes(f);
    ingresosPorDia[kd] = (ingresosPorDia[kd] || 0) + monto;
    ingresosPorMes[km] = (ingresosPorMes[km] || 0) + monto;
  }

  // ---------- 1) Ventas al contado ----------
  var ventas = obtenerHojaVentas().getDataRange().getValues();
  for (var i = 1; i < ventas.length; i++) {
    var fv = ventas[i];
    var fecha = new Date(fv[COL_VENTA.FECHA - 1]);
    if (isNaN(fecha.getTime())) continue;
    if (fv[COL_VENTA.METODO_PAGO - 1] === "Crédito personal") continue;

    var monto = Number(fv[COL_VENTA.TOTAL - 1]) || 0;
    var nombre = (fv[COL_VENTA.NOMBRE - 1] + " " + fv[COL_VENTA.APELLIDO - 1]).trim();
    var tipos = tiposDeVenta(fv);

    sumarIngreso(fecha, monto);
    movs.push({
      ts: fecha.getTime(),
      tipo: "venta",
      // Con cliente registrado: nombre arriba y prendas abajo.
      // Sin cliente: directamente las prendas como título.
      titulo: nombre || tipos.join(", ") || "Venta",
      detalle: nombre ? tipos.join(", ") : "",
      metodo: fv[COL_VENTA.METODO_PAGO - 1] || "",
      monto: formatearMoneda(monto)
    });
  }

  // ---------- 2) Pagos (cobros de crédito + señas) ----------
  var pagos = obtenerHojaPagos().getDataRange().getValues();
  for (var p = 1; p < pagos.length; p++) {
    var fp = pagos[p];
    var fechaP = new Date(fp[COL_PAGO.FECHA - 1]);
    if (isNaN(fechaP.getTime())) continue;

    var montoP = Number(fp[COL_PAGO.MONTO - 1]) || 0;
    sumarIngreso(fechaP, montoP);

    // La seña ("Entrega crédito") se muestra dentro del crédito, no sola.
    if (fp[COL_PAGO.TIPO - 1] === "Entrega crédito") continue;

    movs.push({
      ts: fechaP.getTime(),
      tipo: "pago",
      titulo: fp[COL_PAGO.CLIENTE - 1] || "Cliente",
      detalle: "Pago de crédito",
      metodo: fp[COL_PAGO.METODO - 1] || "",
      monto: formatearMoneda(montoP)
    });
  }

  // ---------- 3) Créditos otorgados ----------
  var creditos = obtenerHojaCreditos().getDataRange().getValues();
  for (var c = 1; c < creditos.length; c++) {
    var fc = creditos[c];
    var fechaC = new Date(fc[COL.FECHA_CARGA - 1]);
    if (isNaN(fechaC.getTime())) continue;

    var precioTotal = Number(fc[COL.PRECIO_TOTAL - 1]) || 0;
    var entrega = Number(fc[COL.ENTREGA - 1]) || 0;
    var saldoInicial = Math.max(precioTotal - entrega, 0);

    var kmC = claveMes(fechaC);
    creditosPorMes[kmC] = (creditosPorMes[kmC] || 0) + saldoInicial;
    cantCreditosPorMes[kmC] = (cantCreditosPorMes[kmC] || 0) + 1;

    var prendas = prendasDeCredito(fc);
    var venc = new Date(fc[COL.VENCIMIENTO - 1]);

    movs.push({
      ts: fechaC.getTime(),
      tipo: "credito",
      titulo: (fc[COL.NOMBRE - 1] + " " + fc[COL.APELLIDO - 1]).trim() || "Crédito",
      detalle: prendas.join(", "),
      entrega: formatearMoneda(entrega),
      saldo: formatearMoneda(saldoInicial),
      vence: isNaN(venc.getTime()) ? "" : Utilities.formatDate(venc, tz, "dd/MM/yyyy")
    });
  }

  // ---------- 4) Gastos fijos (sale plata) ----------
  var gastos = obtenerHojaGastos().getDataRange().getValues();
  for (var gi = 1; gi < gastos.length; gi++) {
    var fg = gastos[gi];
    var fechaG = new Date(fg[COL_GASTO.FECHA - 1]);
    if (isNaN(fechaG.getTime())) continue;

    var montoG = Number(fg[COL_GASTO.MONTO - 1]) || 0;
    sumarEgreso(fechaG, montoG, "gasto");
    movs.push({
      ts: fechaG.getTime(),
      tipo: "gasto",
      titulo: fg[COL_GASTO.CATEGORIA - 1] || "Gasto",
      detalle: fg[COL_GASTO.OBSERVACIONES - 1] || "",
      monto: formatearMoneda(montoG)
    });
  }

  // ---------- 5) Compras de mercadería (sale plata) ----------
  var compras = obtenerHojaCompras().getDataRange().getValues();
  for (var ci = 1; ci < compras.length; ci++) {
    var fcm = compras[ci];
    var fechaCm = new Date(fcm[COL_COMPRA.FECHA - 1]);
    if (isNaN(fechaCm.getTime())) continue;

    var montoCm = Number(fcm[COL_COMPRA.INVERSION_TOTAL - 1]) || 0;
    sumarEgreso(fechaCm, montoCm, "compra");
    movs.push({
      ts: fechaCm.getTime(),
      tipo: "compra",
      titulo: fcm[COL_COMPRA.PROVEEDOR - 1] || "Compra",
      detalle: fcm[COL_COMPRA.OBSERVACIONES - 1] || "",
      monto: formatearMoneda(montoCm)
    });
  }

  // ---------- Orden: lo último registrado primero ----------
  movs.sort(function (a, b) { return b.ts - a.ts; });

  // Arma los totales de un mes: entró, gastos, compras y ganancia.
  function totalesMes(km) {
    var entro = ingresosPorMes[km] || 0;
    var gas = gastosPorMes[km] || 0;
    var comp = comprasPorMes[km] || 0;
    var neto = redondear2(entro - gas - comp);
    return {
      entro: formatearMoneda(entro),
      gastos: formatearMoneda(gas),
      compras: formatearMoneda(comp),
      ganancia: formatearMoneda(Math.abs(neto)),
      gananciaNegativa: neto < 0,
      creditos: formatearMoneda(creditosPorMes[km] || 0),
      cantCreditos: cantCreditosPorMes[km] || 0
    };
  }

  var hoyKey = claveDia(ahora);
  var ayerKey = claveDia(new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate() - 1));

  // ---------- Agrupar por mes y, dentro, por día ----------
  var meses = [];
  var mesPorClave = {};
  var diaPorClave = {};

  movs.forEach(function (m) {
    var d = new Date(m.ts);
    var km = claveMes(d);
    var kd = claveDia(d);

    if (!mesPorClave[km]) {
      mesPorClave[km] = totalesMes(km);
      mesPorClave[km].clave = km;
      mesPorClave[km].nombre = MESES[Number(Utilities.formatDate(d, tz, "M")) - 1];
      mesPorClave[km].anio = Number(Utilities.formatDate(d, tz, "yyyy"));
      mesPorClave[km].dias = [];
      meses.push(mesPorClave[km]);
    }

    if (!diaPorClave[kd]) {
      var etiqueta;
      if (kd === hoyKey) etiqueta = "Hoy";
      else if (kd === ayerKey) etiqueta = "Ayer";
      else etiqueta = DIAS_SEMANA_ES[Number(Utilities.formatDate(d, tz, "u"))] + " " + Utilities.formatDate(d, tz, "dd/MM");

      diaPorClave[kd] = {
        etiqueta: etiqueta,
        ingresos: formatearMoneda(ingresosPorDia[kd] || 0),
        egresos: formatearMoneda(gastosPorDia[kd] || 0),
        hayEgresos: (gastosPorDia[kd] || 0) > 0,
        movs: []
      };
      mesPorClave[km].dias.push(diaPorClave[kd]);
    }

    delete m.ts; // no hace falta en el navegador
    diaPorClave[kd].movs.push(m);
  });

  // El mes en curso siempre se devuelve (aunque todavía no tenga
  // movimientos) para poder mostrar los totales en $0.
  var claveActual = claveMes(ahora);
  var actual = mesPorClave[claveActual];
  if (!actual) {
    actual = totalesMes(claveActual);
    actual.clave = claveActual;
    actual.nombre = MESES[ahora.getMonth()];
    actual.anio = ahora.getFullYear();
    actual.dias = [];
  }

  var anteriores = meses.filter(function (m) { return m.clave !== claveActual; });

  return { actual: actual, anteriores: anteriores };
}
