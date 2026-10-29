// Pegar este código en la hoja de Google: Extensiones > Apps Script.
// Luego: Implementar > Nueva implementación > Aplicación web
//   Ejecutar como: Yo   |   Quién tiene acceso: Cualquier usuario

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(["Fecha", "Nombre", "¿Asiste?", "Mensaje"]);
  }
  const p = e.parameter;
  // Evita que un texto que empiece con = + - @ se interprete como fórmula
  const limpio = (v, max) => {
    const s = String(v || "").slice(0, max);
    return /^[=+\-@]/.test(s) ? "'" + s : s;
  };
  sheet.appendRow([
    new Date(),
    limpio(p.nombre, 80),
    limpio(p.asiste, 5),
    limpio(p.mensaje, 300),
  ]);
  return ContentService.createTextOutput("ok");
}
