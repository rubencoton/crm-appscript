const fs = require('fs');
const path = 'C:/Users/elrub/Desktop/CARPETA CODEX/01_PROYECTOS/festivales-github/Código.js';
let text = fs.readFileSync(path, 'utf8').replace(/\r\n/g, '\n');

text = text.replace(/SpreadsheetApp\.getUi\(\)\.alert\([\s\S]*?'Nombres de contacto ajustados: ' \+ fixedContacts\n  \);/,
`SpreadsheetApp.getUi().alert(
    '✅ Depuracion completada.\\n' +
    'Pestanas tocadas: ' + touchedSheets + '\\n' +
    'Emails ajustados: ' + fixedEmails + '\\n' +
    'Telefonos ajustados: ' + fixedPhones + '\\n' +
    'Nombres de contacto ajustados: ' + fixedContacts
  );`
);

fs.writeFileSync(path, text, 'utf8');
console.log('OK');
