const fs = require('fs');
const path = 'C:/Users/elrub/Desktop/CARPETA CODEX/01_PROYECTOS/festivales-github/Código.js';
let text = fs.readFileSync(path, 'utf8').replace(/\r\n/g, '\n');

text = text.replace(/SpreadsheetApp\.getUi\(\)\.alert\('🛰️ Auditoria de estructura[\s\S]*?lines\.join\('[\s\S]*?'\)\);/,
  "SpreadsheetApp.getUi().alert('🛰️ Auditoria de estructura\\n\\n' + lines.join('\\n'));"
);

text = text.replace(/examples\.length \? 'Muestras:[\s\S]*?\) : 'No se detectaron desajustes en la muestra analizada\.'/,
  "examples.length ? 'Muestras:\\n' + examples.join('\\n') : 'No se detectaron desajustes en la muestra analizada.'"
);

fs.writeFileSync(path, text, 'utf8');
console.log('OK');
