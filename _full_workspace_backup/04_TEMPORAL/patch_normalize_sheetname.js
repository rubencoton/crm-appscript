const fs = require('fs');
const path = 'C:/Users/elrub/Desktop/CARPETA CODEX/01_PROYECTOS/festivales-github/Código.js';
let text = fs.readFileSync(path, 'utf8').replace(/\r\n/g, '\n');

text = text.replace('const normalizedRows = normalizeSheetRows_(data);', 'const normalizedRows = normalizeSheetRows_(data, sheet.getName());');
text = text.replace('function normalizeSheetRows_(data) {', "function normalizeSheetRows_(data, sheetName) {");
text = text.replace('  const out = [];\n', "  const out = [];\n  const tax = parseSheetTaxonomy_(sheetName || '');\n");
text = text.replace("      genero: normalizeGenreCode_(valueAt_(row, map.genero)) || cleanText_(valueAt_(row, map.genero)) || parseSheetTaxonomy_((data[0] && data[0][0]) || '').genre,", "      genero: normalizeGenreCode_(valueAt_(row, map.genero)) || tax.genre || cleanText_(valueAt_(row, map.genero)),");

fs.writeFileSync(path, text, 'utf8');
console.log('OK');
