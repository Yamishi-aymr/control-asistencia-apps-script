const SPREADSHEET_ID = '1C59uZp1LwCHNau0jfOUJF3hTHuN9Jwc33XdzeiJqYOo';

const HOJAS = Object.freeze({
  PARTICIPANTES: 'Participantes',
  SESIONES: 'Sesiones',
  ASISTENCIAS: 'Asistencias'
});

function obtenerLibro_() {
  if (!SPREADSHEET_ID || SPREADSHEET_ID.includes('PEGA_AQUI')) {
    throw new Error(
      'Configura SPREADSHEET_ID en Config.gs antes de ejecutar la aplicación.'
    );
  }

  return SpreadsheetApp.openById(SPREADSHEET_ID);
}