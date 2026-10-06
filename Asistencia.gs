function obtenerSesiones() {
  const ss = obtenerLibro_();
  const hoja = ss.getSheetByName(HOJAS.SESIONES);

  if (!hoja || hoja.getLastRow() < 2) {
    return [];
  }

  const datos = hoja
    .getRange(
      2,
      1,
      hoja.getLastRow() - 1,
      5
    )
    .getValues();

  const zona = Session.getScriptTimeZone();

  return datos.map(fila => ({
    id: fila[0],
    numero: fila[1],

    fecha:
      fila[2] instanceof Date
        ? Utilities.formatDate(
            fila[2],
            zona,
            'yyyy-MM-dd'
          )
        : fila[2],

    tema: fila[3],
    estado: fila[4]
  }));
}


function obtenerParticipantes(idSesion) {
  if (!idSesion) {
    throw new Error(
      'Debes seleccionar una sesión.'
    );
  }

  const ss = obtenerLibro_();

  const hojaParticipantes =
    ss.getSheetByName(HOJAS.PARTICIPANTES);

  const hojaAsistencias =
    ss.getSheetByName(HOJAS.ASISTENCIAS);

  if (!hojaParticipantes || !hojaAsistencias) {
    throw new Error(
      'No se encontraron las hojas necesarias. Ejecuta prepararBaseDatos().'
    );
  }

  const participantes =
    hojaParticipantes.getLastRow() < 2
      ? []
      : hojaParticipantes
          .getRange(
            2,
            1,
            hojaParticipantes.getLastRow() - 1,
            4
          )
          .getValues();

  const asistencias =
    hojaAsistencias.getLastRow() < 2
      ? []
      : hojaAsistencias
          .getRange(
            2,
            1,
            hojaAsistencias.getLastRow() - 1,
            5
          )
          .getValues();

  const asistenciaPorParticipante = new Map();

  asistencias.forEach(fila => {
    if (fila[0] === idSesion) {
      asistenciaPorParticipante.set(
        fila[1],
        {
          estado: fila[2],
          descripcion: fila[3] || ''
        }
      );
    }
  });

  return participantes
    .filter(fila => fila[3] === true)
    .map(fila => {

      const asistencia =
        asistenciaPorParticipante.get(fila[0]);

      return {
        id: fila[0],
        nombre: fila[1],
        correo: fila[2],

        presente:
          asistencia
            ? asistencia.estado === 'PRESENTE'
            : false,

        estado:
          asistencia
            ? asistencia.estado
            : '',

        descripcion:
          asistencia
            ? asistencia.descripcion
            : ''
      };

    });
}


function guardarAsistencia(idSesion, participantes) {
  if (!idSesion) {
    throw new Error(
      'Debes seleccionar una sesión.'
    );
  }

  if (
    !Array.isArray(participantes) ||
    participantes.length === 0
  ) {
    throw new Error(
      'No se recibieron participantes.'
    );
  }

  const ss = obtenerLibro_();

  const hojaSesiones =
    ss.getSheetByName(HOJAS.SESIONES);

  const hojaAsistencias =
    ss.getSheetByName(HOJAS.ASISTENCIAS);

  if (!hojaSesiones || !hojaAsistencias) {
    throw new Error(
      'No se encontraron las hojas necesarias. Ejecuta prepararBaseDatos().'
    );
  }

  // ==============================
  // VALIDAR SESIÓN
  // ==============================

  if (hojaSesiones.getLastRow() < 2) {
    throw new Error(
      'No hay sesiones registradas.'
    );
  }

  const datosSesiones = hojaSesiones
    .getRange(
      2,
      1,
      hojaSesiones.getLastRow() - 1,
      5
    )
    .getValues();

  const sesionActual = datosSesiones.find(
    fila => fila[0] === idSesion
  );

  if (!sesionActual) {
    throw new Error(
      'La sesión seleccionada no existe.'
    );
  }

  // Evitar sobrescribir una sesión registrada
  if (sesionActual[4] === 'REGISTRADA') {
    throw new Error(
      'Esta sesión ya fue registrada y no puede modificarse.'
    );
  }

  // Buscar la primera sesión pendiente
  const primeraPendiente = datosSesiones.find(
    fila => fila[4] !== 'REGISTRADA'
  );

  if (!primeraPendiente) {
    throw new Error(
      'Todas las sesiones ya fueron registradas.'
    );
  }

  // Evitar saltarse sesiones
  if (primeraPendiente[0] !== idSesion) {
    throw new Error(
      `Primero debes registrar la Sesión ${primeraPendiente[1]}.`
    );
  }

  // ==============================
  // GUARDAR ASISTENCIA
  // ==============================

  const existentes =
    hojaAsistencias.getLastRow() < 2
      ? []
      : hojaAsistencias
          .getRange(
            2,
            1,
            hojaAsistencias.getLastRow() - 1,
            5
          )
          .getValues();

  const indice = new Map();

  existentes.forEach((fila, i) => {
    indice.set(
      `${fila[0]}|${fila[1]}`,
      i
    );
  });

  const ahora = new Date();

  participantes.forEach(p => {

    const clave =
      `${idSesion}|${p.id}`;

    // En una sesión nueva la descripción
    // inicia vacía.
    const registro = [
      idSesion,
      p.id,
      p.presente
        ? 'PRESENTE'
        : 'AUSENTE',
      '',
      ahora
    ];

    if (indice.has(clave)) {

      existentes[
        indice.get(clave)
      ] = registro;

    } else {

      indice.set(
        clave,
        existentes.length
      );

      existentes.push(registro);

    }

  });

  // Limpiar los datos anteriores
  if (hojaAsistencias.getLastRow() > 1) {
    hojaAsistencias
      .getRange(
        2,
        1,
        hojaAsistencias.getLastRow() - 1,
        5
      )
      .clearContent();
  }

  // Escribir nuevamente los datos
  if (existentes.length > 0) {

    hojaAsistencias
      .getRange(
        2,
        1,
        existentes.length,
        5
      )
      .setValues(existentes);

    hojaAsistencias
      .getRange(
        2,
        5,
        existentes.length,
        1
      )
      .setNumberFormat(
        'yyyy-mm-dd hh:mm:ss'
      );

  }

  // Marcar la sesión como registrada
  marcarSesionRegistrada_(idSesion);

  const presentes =
    participantes.filter(
      p => p.presente
    ).length;

  // Buscar la siguiente sesión
  const posicionActual =
    datosSesiones.findIndex(
      fila => fila[0] === idSesion
    );

  const siguienteSesion =
    posicionActual >= 0 &&
    posicionActual <
      datosSesiones.length - 1
      ? datosSesiones[
          posicionActual + 1
        ][0]
      : null;

  return {
    ok: true,
    idSesion,
    presentes,
    ausentes:
      participantes.length -
      presentes,
    total:
      participantes.length,
    siguienteSesion
  };
}


function guardarDescripciones(idSesion, participantes) {
  if (!idSesion) {
    throw new Error(
      'Debes seleccionar una sesión.'
    );
  }

  if (!Array.isArray(participantes)) {
    throw new Error(
      'No se recibieron participantes.'
    );
  }

  const ss = obtenerLibro_();

  const hojaSesiones =
    ss.getSheetByName(HOJAS.SESIONES);

  const hojaAsistencias =
    ss.getSheetByName(HOJAS.ASISTENCIAS);

  if (!hojaSesiones || !hojaAsistencias) {
    throw new Error(
      'No se encontraron las hojas necesarias.'
    );
  }

  // ======================================
  // VALIDAR QUE LA SESIÓN YA ESTÉ GUARDADA
  // ======================================

  const datosSesiones = hojaSesiones
    .getRange(
      2,
      1,
      hojaSesiones.getLastRow() - 1,
      5
    )
    .getValues();

  const sesion = datosSesiones.find(
    fila => fila[0] === idSesion
  );

  if (!sesion) {
    throw new Error(
      'La sesión seleccionada no existe.'
    );
  }

  if (sesion[4] !== 'REGISTRADA') {
    throw new Error(
      'Solo se pueden agregar descripciones a sesiones ya registradas.'
    );
  }

  if (hojaAsistencias.getLastRow() < 2) {
    throw new Error(
      'No existen registros de asistencia para esta sesión.'
    );
  }

  const asistencias = hojaAsistencias
    .getRange(
      2,
      1,
      hojaAsistencias.getLastRow() - 1,
      5
    )
    .getValues();

  const participantesPorId = new Map();

  participantes.forEach(p => {
    participantesPorId.set(
      p.id,
      p.descripcion || ''
    );
  });

  let cambios = 0;

  asistencias.forEach((fila, indice) => {

    const idSesionFila = fila[0];
    const idParticipante = fila[1];
    const estado = fila[2];

    if (
      idSesionFila === idSesion &&
      estado === 'AUSENTE' &&
      participantesPorId.has(idParticipante)
    ) {

      const descripcion =
        participantesPorId.get(idParticipante);

      hojaAsistencias
        .getRange(
          indice + 2,
          4
        )
        .setValue(descripcion);

      cambios++;

    }

  });

  return {
    ok: true,
    idSesion,
    cambios
  };
}


function marcarSesionRegistrada_(idSesion) {
  const ss = obtenerLibro_();

  const hoja =
    ss.getSheetByName(HOJAS.SESIONES);

  if (
    !hoja ||
    hoja.getLastRow() < 2
  ) {
    return;
  }

  const datos = hoja
    .getRange(
      2,
      1,
      hoja.getLastRow() - 1,
      5
    )
    .getValues();

  const posicion =
    datos.findIndex(
      fila => fila[0] === idSesion
    );

  if (posicion >= 0) {

    hoja
      .getRange(
        posicion + 2,
        5
      )
      .setValue('REGISTRADA');

  }
}