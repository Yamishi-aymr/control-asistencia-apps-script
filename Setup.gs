function prepararBaseDatos() {
  const ss = obtenerLibro_();

  const participantes = obtenerOCrearHoja_(
    ss,
    HOJAS.PARTICIPANTES,
    ['id_participante', 'nombre', 'correo', 'activo']
  );

  const sesiones = obtenerOCrearHoja_(
    ss,
    HOJAS.SESIONES,
    ['id_sesion', 'numero', 'fecha', 'tema', 'estado']
  );

  const asistencias = obtenerOCrearHoja_(
    ss,
    HOJAS.ASISTENCIAS,
    [
      'id_sesion',
      'id_participante',
      'estado',
      'descripcion',
      'hora_registro'
    ]
  );

  // =========================================
  // CREAR PARTICIPANTES DE PRUEBA
  // =========================================

  if (participantes.getLastRow() === 1) {
    const datosParticipantes = Array.from(
      { length: 30 },
      (_, i) => {

        const numero =
          String(i + 1).padStart(2, '0');

        return [
          `P${numero}`,
          `Participante ${numero}`,
          `participante${numero}@ejemplo.com`,
          true
        ];
      }
    );

    participantes
      .getRange(
        2,
        1,
        datosParticipantes.length,
        datosParticipantes[0].length
      )
      .setValues(datosParticipantes);
  }

  // =========================================
  // CREAR LAS 10 SESIONES
  // =========================================

  if (sesiones.getLastRow() === 1) {
    const inicio = new Date();

    inicio.setHours(0, 0, 0, 0);

    const datosSesiones = Array.from(
      { length: 10 },
      (_, i) => {

        const numero =
          String(i + 1).padStart(2, '0');

        const fecha =
          new Date(inicio);

        fecha.setDate(
          inicio.getDate() + (i * 7)
        );

        return [
          `S${numero}`,
          i + 1,
          fecha,
          `Sesión ${i + 1}`,
          'PENDIENTE'
        ];
      }
    );

    sesiones
      .getRange(
        2,
        1,
        datosSesiones.length,
        datosSesiones[0].length
      )
      .setValues(datosSesiones);

    sesiones
      .getRange(
        2,
        3,
        datosSesiones.length,
        1
      )
      .setNumberFormat('yyyy-mm-dd');
  }

  // =========================================
  // FORMATO DE LAS HOJAS
  // =========================================

  formatearHoja_(participantes);
  formatearHoja_(sesiones);
  formatearHoja_(asistencias);

  return 'Base de datos preparada correctamente.';
}


function obtenerOCrearHoja_(
  ss,
  nombre,
  encabezados
) {
  let hoja =
    ss.getSheetByName(nombre);

  // Si no existe, crearla
  if (!hoja) {
    hoja = ss.insertSheet(nombre);
  }

  // Si está completamente vacía,
  // agregar los encabezados
  if (hoja.getLastRow() === 0) {
    hoja
      .getRange(
        1,
        1,
        1,
        encabezados.length
      )
      .setValues([encabezados]);

    return hoja;
  }

  // =========================================
  // ACTUALIZAR ENCABEZADOS SIN BORRAR DATOS
  // =========================================

  const ultimaColumna =
    Math.max(
      hoja.getLastColumn(),
      encabezados.length
    );

  const encabezadosActuales =
    hoja
      .getRange(
        1,
        1,
        1,
        ultimaColumna
      )
      .getValues()[0];

  encabezados.forEach(
    (encabezado, indice) => {

      if (
        encabezadosActuales[indice] !==
        encabezado
      ) {

        hoja
          .getRange(
            1,
            indice + 1
          )
          .setValue(encabezado);
      }

    }
  );

  return hoja;
}


function formatearHoja_(hoja) {
  if (!hoja) return;

  const ultimaColumna =
    hoja.getLastColumn();

  if (ultimaColumna === 0) return;

  hoja.setFrozenRows(1);

  hoja
    .getRange(
      1,
      1,
      1,
      ultimaColumna
    )
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  hoja.autoResizeColumns(
    1,
    ultimaColumna
  );
}