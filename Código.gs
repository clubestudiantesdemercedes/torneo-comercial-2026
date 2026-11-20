/**
 * Torneo Comercial 2026
 * Club Atlético Estudiantes de Mercedes
 * Sistema de gestión del torneo
 */

// ====================== CONFIGURACIÓN ======================
const SPREADSHEET_ID = '1FWF_e0RxK4i7Jj5IG2GxhC0nIoCnb-YrlQvLTz1tTf0';

const HOJAS = {
  JUGADORES: 'Jugadores',
  EQUIPOS: 'Equipos',
  PARTIDOS: 'Partidos',
  ESTADISTICAS: 'Estadisticas_Partido',
  LOG: 'Log_Auditoria',
  CONFIG: 'Configuracion'
};

// ====================== MENÚ ======================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Torneo Comercial 2026')
    .addItem('1. Inicializar estructura de hojas', 'inicializarEstructura')
    .addItem('2. Cargar / Actualizar equipos', 'cargarEquipos')
    .addItem('3. Normalizar nombres', 'normalizarNombres')
    .addSeparator()
    .addItem('Validar jugadores', 'validarJugadores')
    .addToUi();
}

// ====================== APP WEB ======================
function doGet(e) {
  const page = (e && e.parameter && e.parameter.page) ? e.parameter.page : 'home';
  const id = (e && e.parameter && e.parameter.id) ? e.parameter.id : null;

  let template;
  let titulo = 'Torneo Comercial 2026';

  if (page === 'equipo' && id) {
    template = HtmlService.createTemplateFromFile('Equipo');
    template.datos = obtenerJugadoresPorEquipo(id);
    titulo = template.datos.nombreEquipo + ' · Torneo Comercial 2026';
  } else {
    template = HtmlService.createTemplateFromFile('Index');
    template.datos = obtenerDatosDashboard();
  }

  return template.evaluate()
    .setTitle(titulo)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setFaviconUrl('https://cdn-icons-png.flaticon.com/512/861/861512.png');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ====================== DATOS PÚBLICOS ======================
function obtenerDatosDashboard() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  const hojaEquipos = ss.getSheetByName(HOJAS.EQUIPOS);
  const hojaJugadores = ss.getSheetByName(HOJAS.JUGADORES);
  
  let equipos = [];
  let totalJugadores = 0;
  
  if (hojaEquipos && hojaEquipos.getLastRow() > 1) {
    const datosEquipos = hojaEquipos.getRange(2, 1, hojaEquipos.getLastRow() - 1, 3).getValues();
    
    equipos = datosEquipos.map(fila => ({
      codigo: fila[0],
      nombre: fila[1],
      cantidad: fila[2]
    }));
    
    totalJugadores = equipos.reduce((sum, eq) => sum + Number(eq.cantidad || 0), 0);
  }
  
  return {
    nombreTorneo: 'Torneo Comercial 2026',
    organizador: 'Club Atlético Estudiantes de Mercedes',
    totalEquipos: equipos.length,
    totalJugadores: totalJugadores,
    equipos: equipos
  };
}

function obtenerJugadoresPorEquipo(codEquipo) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hojaJugadores = ss.getSheetByName(HOJAS.JUGADORES);
  const hojaEquipos = ss.getSheetByName(HOJAS.EQUIPOS);
  const hojaStats = ss.getSheetByName(HOJAS.ESTADISTICAS);
  const hojaPartidos = ss.getSheetByName(HOJAS.PARTIDOS);

  let nombreEquipo = codEquipo;

  if (hojaEquipos && hojaEquipos.getLastRow() > 1) {
    const datosEquipos = hojaEquipos.getDataRange().getValues();
    for (let i = 1; i < datosEquipos.length; i++) {
      if (String(datosEquipos[i][0]) === String(codEquipo)) {
        nombreEquipo = datosEquipos[i][1];
        break;
      }
    }
  }

  const partidosConfirmados = new Set();
  if (hojaPartidos && hojaPartidos.getLastRow() > 1) {
    const datosPartidos = hojaPartidos.getDataRange().getValues();
    const headersP = datosPartidos[0];
    const colId = headersP.indexOf('Id_Partido');
    const colEstado = headersP.indexOf('Estado');

    for (let i = 1; i < datosPartidos.length; i++) {
      if (String(datosPartidos[i][colEstado]).toLowerCase() === 'confirmado') {
        partidosConfirmados.add(String(datosPartidos[i][colId]));
      }
    }
  }

  const statsPorJugador = {};

  if (hojaStats && hojaStats.getLastRow() > 1) {
    const datosStats = hojaStats.getDataRange().getValues();
    const headersS = datosStats[0];
    const colIdPartido = headersS.indexOf('Id_Partido');
    const colCodJ = headersS.indexOf('Cod. J.');
    const colPuntos = headersS.indexOf('Puntos');
    const colTriples = headersS.indexOf('Triples');

    for (let i = 1; i < datosStats.length; i++) {
      const idPartido = String(datosStats[i][colIdPartido]);
      const codJ = String(datosStats[i][colCodJ]);

      if (!partidosConfirmados.has(idPartido)) continue;

      if (!statsPorJugador[codJ]) {
        statsPorJugador[codJ] = { pj: 0, puntos: 0, triples: 0 };
      }

      statsPorJugador[codJ].pj += 1;
      statsPorJugador[codJ].puntos += Number(datosStats[i][colPuntos] || 0);
      statsPorJugador[codJ].triples += Number(datosStats[i][colTriples] || 0);
    }
  }

  const jugadores = [];

  if (hojaJugadores && hojaJugadores.getLastRow() > 1) {
    const datos = hojaJugadores.getDataRange().getValues();
    const headers = datos[0];

    const colCod = headers.indexOf('Cod. J.');
    const colApellido = headers.indexOf('Apellido');
    const colNombre = headers.indexOf('Nombre');
    const colEquipo = headers.indexOf('Cod. Equipo');

    for (let i = 1; i < datos.length; i++) {
      if (String(datos[i][colEquipo]) === String(codEquipo)) {
        const codJ = String(datos[i][colCod]);
        const stats = statsPorJugador[codJ] || { pj: 0, puntos: 0, triples: 0 };

        jugadores.push({
          codigo: codJ,
          apellido: datos[i][colApellido],
          nombre: datos[i][colNombre],
          pj: stats.pj,
          puntos: stats.puntos,
          triples: stats.triples
        });
      }
    }
  }

  jugadores.sort((a, b) => a.apellido.localeCompare(b.apellido));

  return {
    codEquipo: codEquipo,
    nombreEquipo: nombreEquipo,
    jugadores: jugadores,
    cantidad: jugadores.length
  };
}

// ====================== UTILIDADES DE PLANILLA ======================
function inicializarEstructura() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  
  let hojaEquipos = ss.getSheetByName(HOJAS.EQUIPOS);
  if (!hojaEquipos) hojaEquipos = ss.insertSheet(HOJAS.EQUIPOS);
  hojaEquipos.clear();
  hojaEquipos.getRange(1, 1, 1, 5).setValues([['Cod. Equipo', 'Nombre', 'Cantidad Jugadores', 'Activo', 'Observaciones']]);
  hojaEquipos.setFrozenRows(1);
  hojaEquipos.getRange(1, 1, 1, 5).setFontWeight('bold');

  let hojaPartidos = ss.getSheetByName(HOJAS.PARTIDOS);
  if (!hojaPartidos) hojaPartidos = ss.insertSheet(HOJAS.PARTIDOS);
  hojaPartidos.clear();
  hojaPartidos.getRange(1, 1, 1, 14).setValues([[
    'Id_Partido', 'Fecha', 'Jornada', 'Equipo_Local', 'Equipo_Visitante',
    'Puntos_Local', 'Puntos_Visitante', 'Estado', 'Es_No_Presentacion',
    'Ganador', 'Observaciones', 'Fecha_Creacion', 'Fecha_Confirmacion', 'Usuario_Confirmacion'
  ]]);
  hojaPartidos.setFrozenRows(1);
  hojaPartidos.getRange(1, 1, 1, 14).setFontWeight('bold');

  let hojaStats = ss.getSheetByName(HOJAS.ESTADISTICAS);
  if (!hojaStats) hojaStats = ss.insertSheet(HOJAS.ESTADISTICAS);
  hojaStats.clear();
  hojaStats.getRange(1, 1, 1, 10).setValues([[
    'Id_Partido', 'Cod. J.', 'Equipo', 'TL', 'Dobles', 'Triples',
    'Puntos', 'Es_Convocado', 'Observaciones', 'Timestamp'
  ]]);
  hojaStats.setFrozenRows(1);
  hojaStats.getRange(1, 1, 1, 10).setFontWeight('bold');

  let hojaLog = ss.getSheetByName(HOJAS.LOG);
  if (!hojaLog) hojaLog = ss.insertSheet(HOJAS.LOG);
  hojaLog.clear();
  hojaLog.getRange(1, 1, 1, 7).setValues([[
    'Timestamp', 'Usuario', 'Accion', 'Entidad', 'Id_Entidad', 'Detalle', 'Valores_Anteriores'
  ]]);
  hojaLog.setFrozenRows(1);
  hojaLog.getRange(1, 1, 1, 7).setFontWeight('bold');

  let hojaConfig = ss.getSheetByName(HOJAS.CONFIG);
  if (!hojaConfig) hojaConfig = ss.insertSheet(HOJAS.CONFIG);
  hojaConfig.clear();
  hojaConfig.getRange(1, 1, 1, 2).setValues([['Clave', 'Valor']]);
  hojaConfig.setFrozenRows(1);
  hojaConfig.getRange(1, 1, 1, 2).setFontWeight('bold');

  const configInicial = [
    ['Nombre_Torneo', 'Torneo Comercial 2026'],
    ['Organizador', 'Club Atlético Estudiantes de Mercedes'],
    ['Anio', '2026'],
    ['Puntos_Victoria', '2'],
    ['Puntos_Derrota', '1'],
    ['Puntos_No_Presentacion', '0'],
    ['Permitir_Empates', 'false'],
    ['Sistema_Desempate', 'FIBA'],
    ['Version_Reglamento', 'FIBA vigente 2026']
  ];
  hojaConfig.getRange(2, 1, configInicial.length, 2).setValues(configInicial);

  mostrarMensaje('Estructura de hojas creada correctamente.');
}

function cargarEquipos() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hojaJugadores = ss.getSheetByName(HOJAS.JUGADORES);
  const hojaEquipos = ss.getSheetByName(HOJAS.EQUIPOS);

  if (!hojaJugadores || !hojaEquipos) {
    mostrarMensaje('Faltan las hojas Jugadores o Equipos. Ejecutá primero "Inicializar estructura".');
    return;
  }

  const datos = hojaJugadores.getDataRange().getValues();
  const headers = datos[0];
  
  const colCodEquipo = headers.indexOf('Cod. Equipo');
  const colEquipo = headers.indexOf('Equipo');

  if (colCodEquipo === -1 || colEquipo === -1) {
    mostrarMensaje('No se encontraron las columnas "Cod. Equipo" o "Equipo".');
    return;
  }

  const equiposMap = {};

  for (let i = 1; i < datos.length; i++) {
    const cod = datos[i][colCodEquipo];
    const nombre = datos[i][colEquipo];
    
    if (cod && nombre) {
      if (!equiposMap[cod]) {
        equiposMap[cod] = { nombre: nombre, cantidad: 0 };
      }
      equiposMap[cod].cantidad++;
    }
  }

  if (hojaEquipos.getLastRow() > 1) {
    hojaEquipos.getRange(2, 1, hojaEquipos.getLastRow() - 1, 5).clearContent();
  }

  const filas = Object.keys(equiposMap).sort().map(cod => [
    cod,
    equiposMap[cod].nombre,
    equiposMap[cod].cantidad,
    'SI',
    ''
  ]);

  if (filas.length > 0) {
    hojaEquipos.getRange(2, 1, filas.length, 5).setValues(filas);
  }

  mostrarMensaje(`Se cargaron/actualizaron ${filas.length} equipos correctamente.`);
}

function normalizarNombres() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hoja = ss.getSheetByName(HOJAS.JUGADORES);
  
  if (!hoja) {
    mostrarMensaje('No se encontró la hoja Jugadores.');
    return;
  }

  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  
  const colApellido = headers.indexOf('Apellido');
  const colNombre = headers.indexOf('Nombre');

  if (colApellido === -1 || colNombre === -1) {
    mostrarMensaje('No se encontraron las columnas Apellido o Nombre.');
    return;
  }

  let cambios = 0;

  for (let i = 1; i < datos.length; i++) {
    const apellidoOriginal = datos[i][colApellido];
    const nombreOriginal = datos[i][colNombre];

    const apellidoNuevo = aFormatoTitulo(apellidoOriginal);
    const nombreNuevo = aFormatoTitulo(nombreOriginal);

    if (apellidoNuevo !== apellidoOriginal || nombreNuevo !== nombreOriginal) {
      hoja.getRange(i + 1, colApellido + 1).setValue(apellidoNuevo);
      hoja.getRange(i + 1, colNombre + 1).setValue(nombreNuevo);
      cambios++;
    }
  }

  mostrarMensaje(`Normalización terminada.\nSe actualizaron ${cambios} jugadores.`);
}

function aFormatoTitulo(texto) {
  if (!texto || typeof texto !== 'string') return texto;
  
  return texto
    .toLowerCase()
    .split(' ')
    .map(palabra => {
      if (palabra.length === 0) return palabra;
      return palabra.charAt(0).toUpperCase() + palabra.slice(1);
    })
    .join(' ');
}

function validarJugadores() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hoja = ss.getSheetByName(HOJAS.JUGADORES);
  
  if (!hoja) {
    mostrarMensaje('No se encontró la hoja Jugadores.');
    return;
  }

  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  
  const colCod = headers.indexOf('Cod. J.');
  const colDNI = headers.indexOf('DNI');
  const colFecha = headers.indexOf('Fecha de nacimiento');
  const colApellido = headers.indexOf('Apellido');
  const colNombre = headers.indexOf('Nombre');

  const problemas = [];
  const codigos = new Set();

  for (let i = 1; i < datos.length; i++) {
    const fila = i + 1;
    const cod = datos[i][colCod];
    const dni = datos[i][colDNI];
    const fecha = datos[i][colFecha];
    const apellido = datos[i][colApellido];
    const nombre = datos[i][colNombre];

    if (!cod) {
      problemas.push(`Fila ${fila}: Falta Código de Jugador`);
    } else if (codigos.has(cod)) {
      problemas.push(`Fila ${fila}: Código duplicado (${cod})`);
    } else {
      codigos.add(cod);
    }

    if (!apellido || !nombre) {
      problemas.push(`Fila ${fila} (${cod || 'sin código'}): Falta Apellido o Nombre`);
    }

    if (!dni) {
      problemas.push(`Fila ${fila} (${cod}): Falta DNI`);
    }

    if (!fecha) {
      problemas.push(`Fila ${fila} (${cod}): Falta Fecha de nacimiento`);
    }
  }

  if (problemas.length === 0) {
    mostrarMensaje('Validación OK. No se encontraron problemas graves.');
  } else {
    const mensaje = `Se encontraron ${problemas.length} problemas:\n\n` + problemas.slice(0, 12).join('\n');
    const extra = problemas.length > 12 ? `\n\n... y ${problemas.length - 12} más.` : '';
    mostrarMensaje(mensaje + extra);
  }
}

function mostrarMensaje(mensaje) {
  try {
    SpreadsheetApp.getUi().alert(mensaje);
  } catch (e) {
    Logger.log(mensaje);
    console.log(mensaje);
  }
}

function registrarLog(accion, entidad, idEntidad, detalle, valoresAnteriores) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const hojaLog = ss.getSheetByName(HOJAS.LOG);
  if (!hojaLog) return;

  const usuario = Session.getActiveUser().getEmail() || 'Desconocido';
  const timestamp = new Date();

  hojaLog.appendRow([
    timestamp,
    usuario,
    accion,
    entidad,
    idEntidad || '',
    detalle || '',
    valoresAnteriores || ''
  ]);
}
