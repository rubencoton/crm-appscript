function auditarEmailsCRMFestivales() {
  if (!validarSesionSegura_('Auditar emails + contraste web + duplicados')) return;
  const summary = ejecutarAuditoriaEmailsFestivales_({ source: 'manual', showUi: true, forceRun: true });
  SpreadsheetApp.getUi().alert(summary);
}

function botonAuditarContactosWebCRMFestivales() {
  abrirPanelAuditoriaContactosCRMFestivales_();
}

function abrirPanelAuditoriaContactosCRMFestivales_() {
  const html = HtmlService.createHtmlOutput(buildHtmlPanelAuditoriaContactos_())
    .setWidth(560)
    .setHeight(460);
  SpreadsheetApp.getUi().showModelessDialog(html, 'AUDITORIA IA | ARTES BUHO');
}

function iniciarAuditoriaContactosDesdeDialogoCRMFestivales_() {
  return ejecutarAuditoriaEmailsFestivales_({ source: 'manual_button', showUi: false, forceRun: true });
}

function obtenerEstadoAuditoriaContactosCRMFestivales_() {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty('FEST_AUDIT_PROGRESS_JSON');
    if (!raw) return { pct: 0, label: 'Esperando inicio...', status: 'IDLE' };
    return JSON.parse(raw);
  } catch (err) {
    return { pct: 0, label: 'Estado no disponible', status: 'ERROR' };
  }
}

function buildHtmlPanelAuditoriaContactos_() {
  return [
    '<!doctype html><html><head><meta charset="utf-8">',
    '<style>',
    'body{font-family:Arial,sans-serif;margin:0;padding:0;background:radial-gradient(circle at 20% 20%,#1f2a44,#070b14 60%);color:#fff;}',
    '.wrap{padding:18px;}',
    '.card{background:rgba(255,255,255,0.08);border:1px solid rgba(255,255,255,0.2);border-radius:14px;padding:16px;}',
    '.hero{height:110px;border-radius:10px;background-image:url(\"https://images.unsplash.com/photo-1451187580459-43490279c0fa?auto=format&fit=crop&w=1200&q=60\");background-size:cover;background-position:center;margin-bottom:12px;box-shadow:inset 0 0 0 9999px rgba(0,0,0,0.35);} ',
    '.title{font-size:20px;font-weight:800;letter-spacing:1px;margin-bottom:10px;}',
    '.sub{font-size:12px;opacity:.85;margin-bottom:14px;}',
    '.bar{height:18px;background:#222;border-radius:20px;overflow:hidden;border:1px solid #555;}',
    '.fill{height:100%;width:0%;background:linear-gradient(90deg,#00e5ff,#00ff9c);} ',
    '.pct{margin-top:8px;font-size:26px;font-weight:900;color:#ffd54f;}',
    '.log{margin-top:10px;background:#0b1020;border:1px solid #2f3d67;border-radius:10px;padding:10px;height:120px;overflow:auto;font-size:12px;white-space:pre-wrap;}',
    '.meta{margin-top:10px;font-size:12px;opacity:.9;}',
    '</style></head><body>',
    '<div class="wrap"><div class="card">',
    '<div class="hero"></div>',
    '<div class="title">AUDITORIA IA DE CONTACTOS</div>',
    '<div class="sub">Desarrollador: RUBEN COTON | Empresa: ARTES BUHO</div>',
    '<div class="bar"><div id="fill" class="fill"></div></div>',
    '<div id="pct" class="pct">0%</div>',
    '<div id="log" class="log">Iniciando...</div>',
    '<div class="meta" id="meta">Estado: preparando...</div>',
    '</div></div>',
    '<script>',
    'const logEl=document.getElementById("log"); const pctEl=document.getElementById("pct"); const fillEl=document.getElementById("fill"); const metaEl=document.getElementById("meta");',
    'function add(msg){logEl.textContent=(logEl.textContent+"\\n"+msg).trim(); logEl.scrollTop=logEl.scrollHeight;}',
    'function paint(st){const p=Math.max(0,Math.min(100,Number(st.pct||0))); pctEl.textContent=p+"%"; fillEl.style.width=p+"%"; metaEl.textContent="Estado: "+(st.status||"RUNNING")+" | "+(st.updatedAt||""); if(st.label) add("["+new Date().toLocaleTimeString()+"] "+st.label);}',
    'function poll(){google.script.run.withSuccessHandler(function(st){paint(st||{}); if((st&&st.status)==="DONE"||(st&&st.status)==="ERROR"){return;} setTimeout(poll,900);}).obtenerEstadoAuditoriaContactosCRMFestivales_();}',
    'google.script.run.withSuccessHandler(function(msg){add("Finalizado: "+(msg||"OK")); poll();}).withFailureHandler(function(err){add("Error: "+(err&&err.message?err.message:err)); poll();}).iniciarAuditoriaContactosDesdeDialogoCRMFestivales_();',
    'setTimeout(poll,600);',
    '</script></body></html>'
  ].join('');
}

function auditarEmailsAutomaticaCRMFestivales_() {
  try {
    ejecutarAuditoriaEmailsFestivales_({ source: 'trigger', showUi: false, forceRun: true });
  } catch (err) {
    Logger.log('auditarEmailsAutomaticaCRMFestivales_ ERROR: ' + (err && err.message ? err.message : err));
  }
}

function auditarEmailsAlAbrirCRMFestivales_() {
  // Desactivado por requerimiento: auditoria de contactos solo por boton manual.
  return;
}

function ejecutarAuditoriaAlAbrirSiCorrespondeCRMFestivales_() {
  // Desactivado por requerimiento: auditoria de contactos solo por boton manual.
  return;
}

function debeEjecutarseRevisionAlAbrir_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(500)) return false;
  try {
    const props = PropertiesService.getScriptProperties();
    const now = Date.now();
    const cooldownMs = Math.max(5, Number(FEST_EMAIL_OPEN_COOLDOWN_MINUTES) || 45) * 60 * 1000;
    const last = Number(props.getProperty('FEST_EMAIL_LAST_OPEN_RUN_TS') || '0');
    if (last && now - last < cooldownMs) return false;
    props.setProperty('FEST_EMAIL_LAST_OPEN_RUN_TS', String(now));
    return true;
  } catch (err) {
    return false;
  } finally {
    lock.releaseLock();
  }
}

function asegurarTriggersAutoRevisionEmailsCRMFestivales_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) return;
  const triggers = ScriptApp.getProjectTriggers();
  let hasClock = false;
  let hasOpen = false;
  for (let i = 0; i < triggers.length; i++) {
    const tr = triggers[i];
    const fn = tr.getHandlerFunction();
    const ev = tr.getEventType();
    if (fn === FEST_EMAIL_TRIGGER_HANDLER && ev === ScriptApp.EventType.CLOCK) hasClock = true;
    if (fn === FEST_EMAIL_OPEN_TRIGGER_HANDLER && ev === ScriptApp.EventType.ON_OPEN) hasOpen = true;
  }
  if (!hasClock) {
    ScriptApp.newTrigger(FEST_EMAIL_TRIGGER_HANDLER)
      .timeBased()
      .everyHours(FEST_EMAIL_TRIGGER_INTERVAL_HOURS)
      .create();
  }
  if (!hasOpen) {
    ScriptApp.newTrigger(FEST_EMAIL_OPEN_TRIGGER_HANDLER)
      .forSpreadsheet(ss)
      .onOpen()
      .create();
  }
}

function instalarTriggerRevisionEmailsCRMFestivales() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const existing = ScriptApp.getProjectTriggers();
  let deleted = 0;
  for (let i = 0; i < existing.length; i++) {
    const fn = existing[i].getHandlerFunction();
    if (fn === FEST_EMAIL_TRIGGER_HANDLER || fn === FEST_EMAIL_OPEN_TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(existing[i]);
      deleted++;
    }
  }
  ScriptApp.newTrigger(FEST_EMAIL_TRIGGER_HANDLER)
    .timeBased()
    .everyHours(FEST_EMAIL_TRIGGER_INTERVAL_HOURS)
    .create();
  ScriptApp.newTrigger(FEST_EMAIL_OPEN_TRIGGER_HANDLER)
    .forSpreadsheet(ss)
    .onOpen()
    .create();
  SpreadsheetApp.getUi().alert(
    'Modo automatico de revision de emails instalado.\n' +
    'Trigger horario: cada ' + FEST_EMAIL_TRIGGER_INTERVAL_HOURS + ' horas.\n' +
    'Trigger al abrir: activo.\n' +
    'Triggers antiguos eliminados: ' + deleted
  );
}

function limpiarTriggerRevisionEmailsCRMFestivales() {
  const deleted = desactivarTriggersAutoRevisionEmailsSilencioso_();
  SpreadsheetApp.getUi().alert('Triggers de revision emails eliminados: ' + deleted);
}

function desactivarTriggersAutoRevisionEmailsSilencioso_() {
  const existing = ScriptApp.getProjectTriggers();
  let deleted = 0;
  for (let i = 0; i < existing.length; i++) {
    const fn = existing[i].getHandlerFunction();
    if (fn === FEST_EMAIL_TRIGGER_HANDLER || fn === FEST_EMAIL_OPEN_TRIGGER_HANDLER) {
      try {
        ScriptApp.deleteTrigger(existing[i]);
        deleted++;
      } catch (err) {
        // no-op
      }
    }
  }
  return deleted;
}

function botonAutocompletadoCeldasIA_CRMFestivales() {
  const summary = autocompletarCeldasIA_CRMFestivales_();
  SpreadsheetApp.getUi().alert(summary);
}

function autocompletarCeldasIA_CRMFestivales_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return 'Hay otro proceso en curso. Intenta de nuevo en unos segundos.';
  let progressCtx = null;
  let ok = false;
  let finishMsg = '';
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = getFestivalSheets_(ss);
    if (!sheets.length) return 'No hay pestañas de festivales para autocompletar.';

    progressCtx = iniciarProcesoRevisionConBloqueo_(ss);
    actualizarProgresoRevision_(progressCtx, 0, 'Iniciando autocompletado IA...');

    const webBudget = {
      webRemaining: Math.max(0, Number(FEST_EMAIL_MAX_WEB_FETCHES_PER_RUN) || 0),
      searchRemaining: Math.max(0, Number(FEST_EMAIL_MAX_SEARCH_FETCHES_PER_RUN) || 0),
      webFetched: 0,
      webCached: 0,
      searchFetched: 0,
      searchCached: 0
    };

    let totalRows = 0;
    for (let s = 0; s < sheets.length; s++) totalRows += Math.max(0, sheets[s].getLastRow() - 1);
    totalRows = Math.max(1, totalRows);

    let touched = 0;
    let placeholders = 0;
    let scanned = 0;

    for (let s = 0; s < sheets.length; s++) {
      const sheet = sheets[s];
      const values = sheet.getDataRange().getValues();
      if (values.length < 2) continue;
      const map = buildHeaderMap_(values[0]);

      for (let r = 1; r < values.length; r++) {
        const row = values[r];
        const festival = cleanText_(valueAt_(row, map.nombre));
        if (!festival) continue;
        scanned++;

        const rec = {
          festival: festival,
          tokens: extraerEmailsValidos_(normalizeEmailCell_(valueAt_(row, map.email))),
          domains: extraerDominiosDeEmails_(extraerEmailsValidos_(normalizeEmailCell_(valueAt_(row, map.email)))),
          seedUrls: extraerUrlsDesdeTexto_(cleanText_(valueAt_(row, map.observaciones)))
        };

        if (!cleanText_(valueAt_(row, map.email))) {
          const web = buscarEvidenciaWebCorreo_(rec, webBudget);
          const email = cleanText_(web.bestEmail || '');
          const v = email || 'IA NO ENCUENTRA';
          sheet.getRange(r + 1, map.email + 1).setValue(v);
          if (email) touched++;
          else placeholders++;
        }

        if (!cleanText_(valueAt_(row, map.telefono))) {
          const tel = buscarTelefonoWebFila_(rec, webBudget);
          const v = tel || 'IA NO ENCUENTRA';
          sheet.getRange(r + 1, map.telefono + 1).setValue(v);
          if (tel) touched++;
          else placeholders++;
        }

        if (!cleanText_(valueAt_(row, map.contacto))) {
          sheet.getRange(r + 1, map.contacto + 1).setValue('IA NO ENCUENTRA');
          placeholders++;
        }
        if (!cleanText_(valueAt_(row, map.ubicacion))) {
          sheet.getRange(r + 1, map.ubicacion + 1).setValue('IA NO ENCUENTRA');
          placeholders++;
        }
        if (!cleanText_(valueAt_(row, map.provincia))) {
          sheet.getRange(r + 1, map.provincia + 1).setValue('IA NO ENCUENTRA');
          placeholders++;
        }
        if (!cleanText_(valueAt_(row, map.ccaa))) {
          sheet.getRange(r + 1, map.ccaa + 1).setValue('IA NO ENCUENTRA');
          placeholders++;
        }

        if (scanned === 1 || scanned % 20 === 0 || scanned === totalRows) {
          const pct = Math.round((scanned / totalRows) * 100);
          actualizarProgresoRevision_(progressCtx, pct, 'Autocompletando fila ' + scanned + ' de ' + totalRows + '...');
        }
      }
    }

    const summary = [
      'AUTOCOMPLETADO IA COMPLETADO',
      'Filas escaneadas: ' + scanned,
      'Celdas completadas con dato: ' + touched,
      'Celdas marcadas IA NO ENCUENTRA: ' + placeholders
    ].join('\n');
    actualizarProgresoRevision_(progressCtx, 100, 'Autocompletado finalizado');
    ok = true;
    finishMsg = summary;
    return summary;
  } finally {
    finalizarProcesoRevisionConBloqueo_(progressCtx, ok, finishMsg || 'Autocompletado IA finalizado');
    lock.releaseLock();
  }
}

function ejecutarAuditoriaEmailsFestivales_(opts) {
  const options = opts || {};
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return 'Otro proceso esta en curso. Intenta en unos segundos.';
  let progressCtx = null;
  let completed = false;
  let completionSummary = '';

  try {
    if (!options.forceRun && !adquirirVentanaEjecucionRevisionEmails_(FEST_EMAIL_RUN_COOLDOWN_MINUTES)) {
      return 'Revision omitida por ventana anti-saturacion.';
    }

    const started = Date.now();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = getFestivalSheets_(ss);
    if (!sheets.length) return 'No encontre pestanas de festivales para revisar.';
    progressCtx = iniciarProcesoRevisionConBloqueo_(ss);
    actualizarProgresoRevision_(progressCtx, 0, 'Iniciando revision...');

    const contexts = [];
    const records = [];
    const tokenCounter = {};
    const uniqueDomains = {};
    let reviewColumnCreated = 0;
    let geoNormalizedCount = 0;
    const webBudget = {
      webRemaining: Math.max(0, Number(FEST_EMAIL_MAX_WEB_FETCHES_PER_RUN) || 0),
      searchRemaining: Math.max(0, Number(FEST_EMAIL_MAX_SEARCH_FETCHES_PER_RUN) || 0),
      webFetched: 0,
      webCached: 0,
      searchFetched: 0,
      searchCached: 0
    };

    for (let s = 0; s < sheets.length; s++) {
      const sheet = sheets[s];
      actualizarProgresoRevision_(progressCtx, Math.round(((s + 1) / Math.max(1, sheets.length)) * 25), 'Preparando: ' + sheet.getName());
      const init = asegurarColumnaRevisionEmailEnSheet_(sheet);
      if (init.created) reviewColumnCreated++;

      const values = sheet.getDataRange().getValues();
      const map = values.length ? buildHeaderMap_(values[0]) : init.map;
      const rowsCount = Math.max(0, values.length - 1);
      const reviewCol = map.reviewEmail + 1;
      const emailCol = map.email + 1;
      const lastCol = Math.max(sheet.getLastColumn(), FEST_HOMO.HEADER.length, reviewCol);
      const currentEmail = rowsCount ? sheet.getRange(2, emailCol, rowsCount, 1).getValues() : [];
      const reviewOut = [];
      const emailOut = [];
      const rowColors = [];
      const ubicacionOut = [];
      const provinciaOut = [];
      const ccaaOut = [];

      for (let r = 0; r < rowsCount; r++) {
        const row = values[r + 1];
        const normUbicacion = normalizeMunicipioName_(valueAt_(row, map.ubicacion));
        const normProvincia = normalizeProvinciaName_(valueAt_(row, map.provincia));
        const normCcaa = normalizeCcaaName_(valueAt_(row, map.ccaa));
        if (cleanText_(valueAt_(row, map.ubicacion)) !== cleanText_(normUbicacion)) geoNormalizedCount++;
        if (cleanText_(valueAt_(row, map.provincia)) !== cleanText_(normProvincia)) geoNormalizedCount++;
        if (cleanText_(valueAt_(row, map.ccaa)) !== cleanText_(normCcaa)) geoNormalizedCount++;
        if (map.ubicacion >= 0 && map.ubicacion < row.length) row[map.ubicacion] = normUbicacion;
        if (map.provincia >= 0 && map.provincia < row.length) row[map.provincia] = normProvincia;
        if (map.ccaa >= 0 && map.ccaa < row.length) row[map.ccaa] = normCcaa;

        reviewOut.push(['MAL']);
        emailOut.push([normalizeEmailCell_(currentEmail[r][0]) || '']);
        rowColors.push(FEST_HOMO.COLORS.BODY_BG);
        ubicacionOut.push([normUbicacion]);
        provinciaOut.push([normProvincia]);
        ccaaOut.push([normCcaa]);
      }

      contexts.push({
        sheet: sheet,
        map: map,
        rowsCount: rowsCount,
        lastCol: lastCol,
        reviewOut: reviewOut,
        emailOut: emailOut,
        rowColors: rowColors,
        ubicacionOut: ubicacionOut,
        provinciaOut: provinciaOut,
        ccaaOut: ccaaOut
      });
      aplicarPoliticaColumnasRevision_(sheet, map, rowsCount);

      if (!rowsCount) continue;
      const taxonomy = parseSheetTaxonomy_(sheet.getName());
      for (let r = 1; r < values.length; r++) {
        const row = values[r];
        const festival = cleanText_(valueAt_(row, map.nombre));
        if (!festival) continue;

        const email = normalizeEmailCell_(valueAt_(row, map.email));
        const tokens = extraerEmailsValidos_(email);
        const domains = extraerDominiosDeEmails_(tokens);
        const obs = cleanText_(valueAt_(row, map.observaciones));
        const seedUrls = extraerUrlsDesdeTexto_(obs);
        for (let d = 0; d < domains.length; d++) uniqueDomains[domains[d]] = true;
        for (let t = 0; t < tokens.length; t++) tokenCounter[tokens[t]] = (tokenCounter[tokens[t]] || 0) + 1;

        const rowGenre = normalizeGenreCode_(valueAt_(row, map.genero));
        const rowSize = sizeCodeFromAforo_(valueAt_(row, map.aforo));
        const genreMismatch = taxonomy.genre && rowGenre && taxonomy.genre !== rowGenre;
        const sizeMismatch = taxonomy.size && rowSize && taxonomy.size !== rowSize;

        records.push({
          sheetName: sheet.getName(),
          rowIndex: r + 1,
          rowOffset: r - 1,
          festival: festival,
          emailBefore: email,
          emailAfter: email,
          tokens: tokens,
          domains: domains,
          seedUrls: seedUrls,
          duplicate: false,
          taxonomyMismatch: !!(genreMismatch || sizeMismatch),
          taxonomyReason: (genreMismatch ? 'GENERO' : '') + (genreMismatch && sizeMismatch ? '+' : '') + (sizeMismatch ? 'AFORO' : ''),
          status: normalizeRevisionStatus_(valueAt_(row, map.reviewEmail)),
          dnsSummary: '',
          webEvidence: '',
          webEmails: '',
          action: ''
        });
      }
    }

    for (let i = 0; i < records.length; i++) {
      const rec = records[i];
      let dup = false;
      for (let j = 0; j < rec.tokens.length; j++) {
        if ((tokenCounter[rec.tokens[j]] || 0) > 1) {
          dup = true;
          break;
        }
      }
      rec.duplicate = dup;
    }

    const domainBatch = obtenerEstadoDominiosEnLote_(Object.keys(uniqueDomains), FEST_EMAIL_MAX_DOMAIN_CHECKS_PER_RUN);
    const bySheet = {};
    for (let c = 0; c < contexts.length; c++) bySheet[contexts[c].sheet.getName()] = contexts[c];

    let bienCount = 0;
    let corregidoCount = 0;
    let malCount = 0;
    let changedEmailCount = 0;
    let duplicateCount = 0;
    let taxonomyMismatchCount = 0;

    for (let i = 0; i < records.length; i++) {
      if (i === 0 || i === records.length - 1 || i % 40 === 0) {
        const pct = records.length ? 25 + Math.round(((i + 1) / records.length) * 55) : 80;
        actualizarProgresoRevision_(progressCtx, pct, 'Chequeando emails (' + (i + 1) + '/' + Math.max(1, records.length) + ')');
      }
      const rec = records[i];
      const result = obtenerResultadoRevisionEmail_(rec, tokenCounter, domainBatch.checks, webBudget);
      rec.status = result.status;
      rec.emailAfter = result.finalEmail;
      rec.dnsSummary = result.dnsSummary;
      rec.webEvidence = result.webEvidence;
      rec.webEmails = result.webEmails;
      rec.action = result.action;

      if (rec.taxonomyMismatch) {
        rec.action = appendReason_(rec.action, 'Revisar clasificacion genero/aforo: ' + rec.taxonomyReason + '.');
        taxonomyMismatchCount++;
      }

      const ctx = bySheet[rec.sheetName];
      if (!ctx || rec.rowOffset < 0 || rec.rowOffset >= ctx.reviewOut.length) continue;
      const previousEmail = normalizeEmailCell_(ctx.emailOut[rec.rowOffset][0]);
      const nextEmail = normalizeEmailCell_(rec.emailAfter);
      if (nextEmail && nextEmail !== previousEmail) {
        ctx.emailOut[rec.rowOffset][0] = nextEmail;
        rec.emailAfter = nextEmail;
        changedEmailCount++;
        actualizarContadorEmailsTrasCambio_(tokenCounter, rec.tokens, nextEmail);
      } else {
        rec.emailAfter = previousEmail;
      }

      ctx.reviewOut[rec.rowOffset][0] = rec.status;
      ctx.rowColors[rec.rowOffset] = colorPorEstadoRevisionEmail_(rec.status);
      if (rec.status === 'BIEN') bienCount++;
      else if (rec.status === 'CORREGIDO') corregidoCount++;
      else malCount++;
      if (rec.duplicate) duplicateCount++;
    }

    for (let c = 0; c < contexts.length; c++) {
      const ctx = contexts[c];
      if (!ctx.rowsCount) continue;
      actualizarProgresoRevision_(progressCtx, 82 + Math.round(((c + 1) / Math.max(1, contexts.length)) * 14), 'Escribiendo: ' + ctx.sheet.getName());
      const emailCol = ctx.map.email + 1;
      const reviewCol = ctx.map.reviewEmail + 1;
      const ubicacionCol = ctx.map.ubicacion + 1;
      const provinciaCol = ctx.map.provincia + 1;
      const ccaaCol = ctx.map.ccaa + 1;
      ctx.sheet.getRange(2, emailCol, ctx.rowsCount, 1).setValues(ctx.emailOut);
      ctx.sheet.getRange(2, reviewCol, ctx.rowsCount, 1).setValues(ctx.reviewOut);
      ctx.sheet.getRange(2, ubicacionCol, ctx.rowsCount, 1).setValues(ctx.ubicacionOut);
      ctx.sheet.getRange(2, provinciaCol, ctx.rowsCount, 1).setValues(ctx.provinciaOut);
      ctx.sheet.getRange(2, ccaaCol, ctx.rowsCount, 1).setValues(ctx.ccaaOut);
      aplicarPoliticaColumnasRevision_(ctx.sheet, ctx.map, ctx.rowsCount);
      const backgrounds = construirMatrizColorFilas_(ctx.rowColors, ctx.lastCol);
      ctx.sheet.getRange(2, 1, ctx.rowsCount, ctx.lastCol).setBackgrounds(backgrounds);
    }

    actualizarProgresoRevision_(progressCtx, 97, 'Generando reporte...');
    escribirPestanaRevisionEmails_(ss, records, domainBatch, webBudget, options.source || 'manual');

    const elapsedSec = Math.round((Date.now() - started) / 1000);
    const summary = [
      'Revision emails completada (' + FEST_ARCHITECT + ')',
      '',
      'Pestanas revisadas: ' + sheets.length,
      'Filas analizadas: ' + records.length,
      'Columna \"REVISION EMAIL\" creada: ' + reviewColumnCreated,
      'Estado BIEN: ' + bienCount,
      'Estado CORREGIDO: ' + corregidoCount,
      'Estado MAL: ' + malCount,
      'Correos actualizados: ' + changedEmailCount,
      'Duplicados detectados: ' + duplicateCount,
      'Filas con conflicto genero/aforo: ' + taxonomyMismatchCount,
      'Campos geo homogeneizados: ' + geoNormalizedCount,
      'DNS consultados: ' + domainBatch.fetched + ' | cache DNS: ' + domainBatch.cached + ' | DNS omitidos por limite: ' + domainBatch.skipped,
      'Web consultada: ' + webBudget.webFetched + ' | cache web: ' + webBudget.webCached,
      'Busquedas web: ' + webBudget.searchFetched + ' | cache busquedas: ' + webBudget.searchCached,
      'Tiempo total: ' + elapsedSec + 's'
    ].join('\n');

    completionSummary = summary;
    completed = true;
    actualizarProgresoRevision_(progressCtx, 100, 'Completado');
    if (!options.showUi) Logger.log(summary);
    return summary;
  } finally {
    finalizarProcesoRevisionConBloqueo_(progressCtx, completed, completionSummary);
    lock.releaseLock();
  }
}

function adquirirVentanaEjecucionRevisionEmails_(cooldownMinutes) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(500)) return false;
  try {
    const props = PropertiesService.getScriptProperties();
    const now = Date.now();
    const last = Number(props.getProperty('FEST_EMAIL_LAST_RUN_TS') || '0');
    const cooldownMs = Math.max(1, Number(cooldownMinutes) || 10) * 60 * 1000;
    if (last && now - last < cooldownMs) return false;
    props.setProperty('FEST_EMAIL_LAST_RUN_TS', String(now));
    return true;
  } catch (err) {
    return true;
  } finally {
    lock.releaseLock();
  }
}

function iniciarProcesoRevisionConBloqueo_(ss) {
  const ctx = {
    ss: ss,
    progressSheet: asegurarPestanaProgresoRevision_(ss),
    protection: null,
    lastPct: -1,
    lastLabel: ''
  };
  ctx.protection = crearBloqueoTemporalRevision_(ss);
  return ctx;
}

function finalizarProcesoRevisionConBloqueo_(ctx, ok, summary) {
  if (!ctx) return;
  try {
    const estado = ok ? 'Completado' : 'Finalizado con incidencias';
    actualizarProgresoRevision_(ctx, 100, estado);
    const sh = ctx.progressSheet || asegurarPestanaProgresoRevision_(ctx.ss);
    sh.getRange('A6').setValue('RESUMEN');
    sh.getRange('B6').setValue(cleanText_(summary || '').substring(0, 500));
  } catch (err) {
    // no-op
  }
  try {
    PropertiesService.getScriptProperties().setProperty(
      'FEST_AUDIT_PROGRESS_JSON',
      JSON.stringify({
        pct: 100,
        label: ok ? 'Auditoria completada' : 'Auditoria con incidencias',
        status: ok ? 'DONE' : 'ERROR',
        updatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
      })
    );
  } catch (err) {
    // no-op
  }
  try {
    if (ctx.protection) ctx.protection.remove();
  } catch (err) {
    // no-op
  }
}

function crearBloqueoTemporalRevision_(ss) {
  try {
    const p = ss.protect();
    p.setDescription('CRM FESTIVALES - BLOQUEO TEMPORAL REVISION EMAIL');
    p.setWarningOnly(false);
    const me = cleanText_(Session.getEffectiveUser().getEmail()).toLowerCase();
    if (!me || !p.canEdit()) {
      p.setWarningOnly(true);
      return p;
    }
    const editors = p.getEditors();
    for (let i = 0; i < editors.length; i++) {
      const email = cleanText_(editors[i].getEmail()).toLowerCase();
      if (email && email !== me) {
        try { p.removeEditor(editors[i]); } catch (err) {}
      }
    }
    try { if (p.canDomainEdit()) p.setDomainEdit(false); } catch (err) {}
    return p;
  } catch (err) {
    return null;
  }
}

function asegurarPestanaProgresoRevision_(ss) {
  let sh = ss.getSheetByName('PROGRESO_CRM');
  if (!sh) sh = ss.insertSheet('PROGRESO_CRM');
  sh.getRange('A1:C1').setValues([['METRICA', 'VALOR', 'BARRA']]);
  sh.getRange('A2:C5').clearContent();
  sh.getRange('A2:A5').setValues([
    ['ESTADO'],
    ['PROGRESO'],
    ['ULTIMA_ACTUALIZACION'],
    ['BLOQUEO']
  ]);
  sh.getRange('A1:C1')
    .setBackground('#8B0000')
    .setFontColor('#FFFFFF')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sh.setColumnWidth(1, 190);
  sh.setColumnWidth(2, 380);
  sh.setColumnWidth(3, 360);
  sh.setFrozenRows(1);
  return sh;
}

function actualizarProgresoRevision_(ctx, pct, label) {
  if (!ctx || !ctx.ss) return;
  const safePct = Math.max(0, Math.min(100, Math.round(Number(pct) || 0)));
  const safeLabel = cleanText_(label);
  if (safePct === ctx.lastPct && safeLabel === ctx.lastLabel) return;
  ctx.lastPct = safePct;
  ctx.lastLabel = safeLabel;

  const sh = ctx.progressSheet || asegurarPestanaProgresoRevision_(ctx.ss);
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  sh.getRange('B2').setValue(safeLabel || 'En curso');
  sh.getRange('B3').setValue(safePct + '%');
  sh.getRange('C3').setValue(renderBarraProgreso_(safePct));
  sh.getRange('B4').setValue(now);
  sh.getRange('B5').setValue(ctx.protection ? 'SI' : 'NO');
  try {
    PropertiesService.getScriptProperties().setProperty(
      'FEST_AUDIT_PROGRESS_JSON',
      JSON.stringify({
        pct: safePct,
        label: safeLabel || 'En curso',
        status: safePct >= 100 ? 'DONE' : 'RUNNING',
        updatedAt: now
      })
    );
  } catch (err) {
    // no-op
  }
  try {
    ctx.ss.toast('Revision emails: ' + safePct + '% - ' + (safeLabel || 'En curso'), 'CRM FESTIVALES', 5);
  } catch (err) {}
}

function renderBarraProgreso_(pct) {
  const total = 20;
  const fill = Math.max(0, Math.min(total, Math.round((pct / 100) * total)));
  return '[' + Array(fill + 1).join('#') + Array(total - fill + 1).join('-') + '] ' + pct + '%';
}

function aplicarPoliticaColumnasRevision_(sheet, map, rowsCount) {
  const reviewCol = map.reviewEmail + 1;
  const mergeCol = map.mergeStatus + 1;

  sheet.getRange(1, reviewCol, 1, 1)
    .setValue('REVISION EMAIL')
    .setBackground('#FBC02D')
    .setFontColor('#8B0000')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  sheet.getRange(1, mergeCol, 1, 1)
    .setValue(FEST_MERGE_STATUS_HEADER)
    .setHorizontalAlignment('center');

  if (rowsCount <= 0) return;
  sheet.getRange(2, reviewCol, rowsCount, 1)
    .setDataValidation(buildEmailReviewValidationRule_())
    .setHorizontalAlignment('center')
    .setFontWeight('bold');
  const mergeRows = Math.max(0, sheet.getMaxRows() - 1);
  if (mergeRows > 0) {
    sheet.getRange(2, mergeCol, mergeRows, 1)
      .clearDataValidations()
      .setHorizontalAlignment('left')
      .setFontWeight('normal');
  }
}

function asegurarColumnaRevisionEmailEnSheet_(sheet) {
  const currentMaxCol = Math.max(sheet.getLastColumn(), FEST_HOMO.HEADER.length);
  const firstRow = sheet.getRange(1, 1, 1, currentMaxCol).getValues()[0];
  const map = buildHeaderMap_(firstRow);
  let created = false;
  const reviewHeader = normalizeHeader_(valueAt_(firstRow, map.reviewEmail));
  const mergeHeader = cleanText_(valueAt_(firstRow, map.mergeStatus));
  const inOrder = isHeaderInCanonicalOrder_(firstRow);
  const hasReviewHeader = reviewHeader === 'REVISION EMAIL';
  const hasMergeHeader = mergeHeader === FEST_MERGE_STATUS_HEADER;

  if (!hasReviewHeader || !hasMergeHeader || !inOrder) {
    const data = sheet.getDataRange().getValues();
    const normalizedRows = normalizeSheetRows_(data, sheet.getName());
    rewriteSheet_(sheet, normalizedRows);
    applyVisualDesignToSheet_(sheet);
    created = true;
  }

  const finalMaxCol = Math.max(sheet.getLastColumn(), FEST_HOMO.HEADER.length);
  const finalHeader = sheet.getRange(1, 1, 1, finalMaxCol).getValues()[0];
  const finalMap = buildHeaderMap_(finalHeader);
  const reviewCol = finalMap.reviewEmail + 1;
  const mergeCol = finalMap.mergeStatus + 1;

  if (normalizeHeader_(sheet.getRange(1, reviewCol).getValue()) !== 'REVISION EMAIL') {
    sheet.getRange(1, reviewCol).setValue('REVISION EMAIL');
    created = true;
  }
  if (cleanText_(sheet.getRange(1, mergeCol).getValue()) !== FEST_MERGE_STATUS_HEADER) {
    sheet.getRange(1, mergeCol).setValue(FEST_MERGE_STATUS_HEADER);
    created = true;
  }

  const lastRow = sheet.getLastRow();
  aplicarPoliticaColumnasRevision_(sheet, finalMap, Math.max(0, lastRow - 1));

  return { map: finalMap, created: created };
}

function extraerEmailsValidos_(rawEmail) {
  const txt = cleanText_(rawEmail).toLowerCase();
  if (!txt) return [];
  const tokens = txt.split(/[;,|\s]+/).map((x) => x.trim()).filter((x) => x);
  const out = [];
  const seen = {};
  for (let i = 0; i < tokens.length; i++) {
    const email = tokens[i];
    if (!isValidEmail_(email)) continue;
    if (seen[email]) continue;
    seen[email] = true;
    out.push(email);
  }
  return out;
}

function extraerDominiosDeEmails_(emails) {
  const out = [];
  const seen = {};
  for (let i = 0; i < emails.length; i++) {
    const email = emails[i];
    const at = email.indexOf('@');
    if (at === -1) continue;
    const domain = cleanText_(email.substring(at + 1)).toLowerCase();
    if (!domain || seen[domain]) continue;
    seen[domain] = true;
    out.push(domain);
  }
  return out;
}

function obtenerResultadoRevisionEmail_(record, counter, checks, webBudget) {
  const out = {
    status: 'MAL',
    finalEmail: record.emailBefore,
    dnsSummary: resumirDnsFila_(record, checks),
    webEvidence: '',
    webEmails: '',
    action: ''
  };

  const web = buscarEvidenciaWebCorreo_(record, webBudget);
  out.webEvidence = web.evidenceUrls.length ? web.evidenceUrls[0] : '';
  out.webEmails = web.emails.join('; ');

  const current = record.tokens.length ? record.tokens[0] : '';
  const best = web.bestEmail || '';
  const dnsStatusCurrent = estadoDnsPorEmail_(current, checks);
  const currentLooksValid = !!current && isValidEmail_(current);

  if (current && web.matchCurrent) {
    out.status = 'BIEN';
    out.finalEmail = current;
    out.action = 'Email confirmado en web.';
  } else if (best) {
    out.finalEmail = best;
    out.status = (current && best === current) ? 'BIEN' : 'CORREGIDO';
    out.action = out.status === 'CORREGIDO'
      ? 'Email actualizado segun evidencia web.'
      : 'Email confirmado en web.';
  } else if (currentLooksValid && (dnsStatusCurrent === 'MX_OK' || dnsStatusCurrent === 'A_OK')) {
    out.status = 'BIEN';
    out.finalEmail = current;
    out.action = 'Email valido por sintaxis y DNS.';
  } else if (currentLooksValid && (dnsStatusCurrent === 'BUDGET_LIMIT' || dnsStatusCurrent === 'ERROR' || dnsStatusCurrent === 'N/D')) {
    out.status = 'BIEN';
    out.finalEmail = current;
    out.action = 'Email valido por sintaxis (pendiente de validacion DNS/web completa).';
  } else {
    out.status = 'MAL';
    out.finalEmail = current || '';
    out.action = current
      ? 'No se pudo validar ese email en la web.'
      : 'No hay email y no se encontro uno fiable en web.';
  }

  if (record.duplicate) {
    if (out.status === 'CORREGIDO') {
      const cnt = counter[out.finalEmail] || 0;
      if (cnt > 1) {
        out.status = 'MAL';
        out.action = appendReason_(out.action, 'Sigue duplicado tras el cambio.');
      } else {
        out.action = appendReason_(out.action, 'Cambio aplicado para resolver duplicado.');
      }
    } else {
      out.action = appendReason_(out.action, 'Email duplicado en CRM (revisar manualmente).');
    }
  }

  if (!out.finalEmail) out.status = 'MAL';
  return out;
}

function estadoDnsPorEmail_(email, checks) {
  const e = cleanText_(email).toLowerCase();
  if (!e || e.indexOf('@') === -1) return 'N/D';
  const domain = e.split('@')[1] || '';
  if (!domain) return 'N/D';
  const chk = checks[domain];
  if (!chk) return 'N/D';
  if (chk.hasMx) return 'MX_OK';
  if (chk.hasA) return 'A_OK';
  return chk.status || 'N/D';
}

function resumirDnsFila_(record, checks) {
  if (!record.domains.length) return 'N/D';
  const parts = [];
  for (let i = 0; i < record.domains.length; i++) {
    const domain = record.domains[i];
    const chk = checks[domain];
    if (!chk || !chk.checked) {
      parts.push(domain + ':N/D');
    } else if (chk.hasMx) {
      parts.push(domain + ':MX_OK');
    } else if (chk.hasA) {
      parts.push(domain + ':A_OK');
    } else {
      parts.push(domain + ':SIN_DNS');
    }
  }
  return parts.join(' | ');
}

function buscarEvidenciaWebCorreo_(record, webBudget) {
  const emailsSet = {};
  const evidence = [];
  let matchCurrent = false;
  const currentSet = {};
  for (let i = 0; i < record.tokens.length; i++) currentSet[record.tokens[i]] = true;

  const urls = construirUrlsCandidatasFila_(record);
  const maxUrlsPerRow = 8;
  for (let i = 0; i < urls.length && i < maxUrlsPerRow; i++) {
    const hit = buscarEmailEnSitio_(urls[i], webBudget);
    if (!hit || !hit.emails.length) continue;
    evidence.push(hit.url);
    for (let j = 0; j < hit.emails.length; j++) {
      const email = hit.emails[j];
      emailsSet[email] = true;
      if (currentSet[email]) matchCurrent = true;
    }
    if (matchCurrent) break;
  }

  if (!matchCurrent && !Object.keys(emailsSet).length) {
    const searchUrls = buscarUrlsWebPorFestival_(record.festival, webBudget);
    for (let i = 0; i < searchUrls.length && i < 4; i++) {
      const hit = buscarEmailEnSitio_(searchUrls[i], webBudget);
      if (!hit || !hit.emails.length) continue;
      evidence.push(hit.url);
      for (let j = 0; j < hit.emails.length; j++) {
        const email = hit.emails[j];
        emailsSet[email] = true;
        if (currentSet[email]) matchCurrent = true;
      }
    }
  }

  const allEmails = Object.keys(emailsSet);
  const bestEmail = seleccionarMejorEmailCandidato_(allEmails, record);
  return {
    emails: allEmails,
    bestEmail: bestEmail,
    evidenceUrls: dedupeStrings_(evidence).slice(0, 5),
    matchCurrent: matchCurrent
  };
}

function buscarTelefonoWebFila_(record, webBudget) {
  const urls = construirUrlsCandidatasFila_(record).slice(0, 8);
  for (let i = 0; i < urls.length; i++) {
    const html = obtenerHtmlCacheado_(urls[i], webBudget, 'web');
    if (!html) continue;
    const phones = extraerTelefonosDesdeHtml_(html);
    if (phones.length) return phones[0];
  }
  const searchUrls = buscarUrlsWebPorFestival_(record.festival, webBudget).slice(0, 4);
  for (let i = 0; i < searchUrls.length; i++) {
    const html = obtenerHtmlCacheado_(searchUrls[i], webBudget, 'search');
    if (!html) continue;
    const phones = extraerTelefonosDesdeHtml_(html);
    if (phones.length) return phones[0];
  }
  return '';
}

function construirUrlsCandidatasFila_(record) {
  const out = [];
  const seen = {};
  function add(url) {
    const u = normalizarUrlSegura_(url);
    if (!u || seen[u]) return;
    seen[u] = true;
    out.push(u);
  }
  for (let i = 0; i < record.seedUrls.length; i++) add(record.seedUrls[i]);
  for (let i = 0; i < record.domains.length; i++) {
    const domain = record.domains[i];
    if (!domain) continue;
    add('https://' + domain);
    if (domain.indexOf('www.') !== 0) add('https://www.' + domain);
    add('https://' + domain + '/contacto');
    add('https://' + domain + '/contact');
    add('https://' + domain + '/contact-us');
    add('https://' + domain + '/about');
  }
  return out;
}

function buscarEmailEnSitio_(url, webBudget) {
  const normalized = normalizarUrlSegura_(url);
  if (!normalized) return null;
  const html = obtenerHtmlCacheado_(normalized, webBudget, 'web');
  if (!html) return null;
  const emails = extraerEmailsDesdeHtml_(html);
  if (!emails.length) return null;
  return { url: normalized, emails: emails };
}

function obtenerHtmlCacheado_(url, webBudget, mode) {
  const cache = CacheService.getScriptCache();
  const key = 'FEST_WEB_HTML_' + hashText_(url);
  const cached = cache.get(key);
  if (cached) {
    if (mode === 'search') webBudget.searchCached++;
    else webBudget.webCached++;
    return cached;
  }

  if (mode === 'search') {
    if (webBudget.searchRemaining <= 0) return '';
    webBudget.searchRemaining--;
    webBudget.searchFetched++;
  } else {
    if (webBudget.webRemaining <= 0) return '';
    webBudget.webRemaining--;
    webBudget.webFetched++;
  }

  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'get',
      muteHttpExceptions: true,
      followRedirects: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (compatible; CRM-FESTIVALES/1.0)' }
    });
    const code = res.getResponseCode();
    if (code < 200 || code >= 400) return '';
    const html = res.getContentText() || '';
    if (html) {
      try {
        // CacheService limita el tamano por entrada. Si no cabe, no rompemos la auditoria.
        if (html.length <= 90000) cache.put(key, html, FEST_EMAIL_WEB_CACHE_TTL_SEC);
      } catch (cacheErr) {
        // no-op
      }
    }
    return html;
  } catch (err) {
    return '';
  }
}

function buscarUrlsWebPorFestival_(festivalName, webBudget) {
  const query = cleanText_(festivalName);
  if (!query) return [];
  const url = 'https://duckduckgo.com/html/?q=' + encodeURIComponent(query + ' festival email contacto');
  const html = obtenerHtmlCacheado_(url, webBudget, 'search');
  if (!html) return [];
  return extraerUrlsDeDuckDuckGo_(html).slice(0, 8);
}

function extraerUrlsDeDuckDuckGo_(html) {
  const urls = [];
  const seen = {};
  const reUddg = /uddg=([^\"&]+)/ig;
  let m;
  while ((m = reUddg.exec(html)) !== null) {
    const decoded = decodeURIComponent(m[1] || '');
    const url = normalizarUrlSegura_(decoded);
    if (!url || seen[url]) continue;
    if (!esUrlCandidataParaRevision_(url)) continue;
    seen[url] = true;
    urls.push(url);
  }
  return urls;
}

function esUrlCandidataParaRevision_(url) {
  const t = cleanText_(url).toLowerCase();
  if (!t) return false;
  if (t.indexOf('duckduckgo.com') > -1) return false;
  if (t.indexOf('google.com/search') > -1) return false;
  if (t.indexOf('instagram.com') > -1) return false;
  if (t.indexOf('facebook.com') > -1) return false;
  if (t.indexOf('x.com/') > -1 || t.indexOf('twitter.com/') > -1) return false;
  return /^https?:\/\//i.test(t);
}

function extraerEmailsDesdeHtml_(html) {
  const raw = String(html || '');
  const matches = raw.match(/[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}/ig) || [];
  const out = [];
  const seen = {};
  for (let i = 0; i < matches.length; i++) {
    let email = cleanText_(matches[i]).toLowerCase();
    email = email.replace(/[),.;:]+$/g, '');
    if (!isValidEmail_(email)) continue;
    if (seen[email]) continue;
    seen[email] = true;
    out.push(email);
  }
  return out;
}

function extraerTelefonosDesdeHtml_(html) {
  const raw = String(html || '');
  const matches = raw.match(/(?:\+34|0034)?[\s().-]*(?:\d[\s().-]*){9,12}/g) || [];
  const out = [];
  const seen = {};
  for (let i = 0; i < matches.length; i++) {
    const normalized = formatSpanishPhone_(matches[i]);
    if (!normalized || seen[normalized]) continue;
    seen[normalized] = true;
    out.push(normalized);
  }
  return out;
}

function seleccionarMejorEmailCandidato_(emails, record) {
  if (!emails || !emails.length) return '';
  const currentSet = {};
  const domainSet = {};
  for (let i = 0; i < record.tokens.length; i++) currentSet[record.tokens[i]] = true;
  for (let i = 0; i < record.domains.length; i++) domainSet[record.domains[i]] = true;

  const keys = normalizeHeader_(record.festival).toLowerCase().split(/[^a-z0-9]+/).filter((x) => x && x.length >= 4);
  let best = '';
  let bestScore = -9999;
  for (let i = 0; i < emails.length; i++) {
    const email = emails[i];
    const at = email.indexOf('@');
    const local = at > 0 ? email.substring(0, at) : '';
    const domain = at > 0 ? email.substring(at + 1) : '';
    let score = 0;
    if (currentSet[email]) score += 30;
    if (domainSet[domain]) score += 12;
    if (/info|contact|booking|prensa|press|admin|oficina/i.test(local)) score += 4;
    if (/noreply|no-reply|do-not-reply/i.test(local)) score -= 8;
    for (let k = 0; k < keys.length; k++) {
      if (domain.indexOf(keys[k]) > -1) score += 2;
      if (local.indexOf(keys[k]) > -1) score += 2;
    }
    if (score > bestScore || (score === bestScore && email < best)) {
      bestScore = score;
      best = email;
    }
  }
  return best;
}

function extraerUrlsDesdeTexto_(text) {
  const raw = cleanText_(text);
  if (!raw) return [];
  const out = [];
  const seen = {};
  const re = /((https?:\/\/|www\.)[^\s,;]+)/ig;
  let m;
  while ((m = re.exec(raw)) !== null) {
    const u = normalizarUrlSegura_(m[1]);
    if (!u || seen[u]) continue;
    seen[u] = true;
    out.push(u);
  }
  return out;
}

function normalizarUrlSegura_(url) {
  let t = cleanText_(url);
  if (!t) return '';
  if (!/^https?:\/\//i.test(t)) t = 'https://' + t;
  t = t.replace(/[#].*$/g, '');
  if (t.length > 400) return '';
  if (!/^https?:\/\/[a-z0-9.\-]+/i.test(t)) return '';
  return t;
}

function appendReason_(base, extra) {
  const a = cleanText_(base);
  const b = cleanText_(extra);
  if (!a) return b;
  if (!b) return a;
  return a + ' ' + b;
}

function dedupeStrings_(arr) {
  const out = [];
  const seen = {};
  for (let i = 0; i < arr.length; i++) {
    const t = cleanText_(arr[i]);
    if (!t || seen[t]) continue;
    seen[t] = true;
    out.push(t);
  }
  return out;
}

function hashText_(input) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    cleanText_(input),
    Utilities.Charset.UTF_8
  );
  const chars = [];
  for (let i = 0; i < bytes.length; i++) {
    const v = (bytes[i] + 256) % 256;
    chars.push((v < 16 ? '0' : '') + v.toString(16));
  }
  return chars.join('');
}

function actualizarContadorEmailsTrasCambio_(counter, oldTokens, newEmail) {
  for (let i = 0; i < oldTokens.length; i++) {
    const old = oldTokens[i];
    if (!counter[old]) continue;
    counter[old] = Math.max(0, counter[old] - 1);
  }
  const newTokens = extraerEmailsValidos_(newEmail);
  for (let i = 0; i < newTokens.length; i++) {
    const e = newTokens[i];
    counter[e] = (counter[e] || 0) + 1;
  }
}

function colorPorEstadoRevisionEmail_(status) {
  const s = cleanText_(status).toUpperCase();
  if (s === 'BIEN') return '#D9EAD3';
  if (s === 'CORREGIDO') return '#D9E8FB';
  return '#F4CCCC';
}

function construirMatrizColorFilas_(rowColors, totalCols) {
  const cols = Math.max(1, Number(totalCols) || 1);
  const out = [];
  for (let r = 0; r < rowColors.length; r++) {
    const color = rowColors[r] || FEST_HOMO.COLORS.BODY_BG;
    const row = [];
    for (let c = 0; c < cols; c++) row.push(color);
    out.push(row);
  }
  return out;
}

function obtenerEstadoDominiosEnLote_(domains, maxChecks) {
  const checks = {};
  let fetched = 0;
  let cached = 0;
  let skipped = 0;
  const budget = Math.max(0, Number(maxChecks) || 0);

  for (let i = 0; i < domains.length; i++) {
    const domain = cleanText_(domains[i]).toLowerCase();
    if (!domain) continue;

    const cachedObj = leerCacheDominio_(domain);
    if (cachedObj) {
      checks[domain] = cachedObj;
      cached++;
      continue;
    }

    if (fetched >= budget) {
      checks[domain] = { checked: false, hasMx: false, hasA: false, status: 'BUDGET_LIMIT', checkedAt: Date.now() };
      skipped++;
      continue;
    }

    const fresh = resolverDominioPorDns_(domain);
    checks[domain] = fresh;
    guardarCacheDominio_(domain, fresh);
    fetched++;
  }

  return { checks: checks, fetched: fetched, cached: cached, skipped: skipped };
}

function claveCacheDominio_(domain) {
  const safe = cleanText_(domain).toLowerCase().replace(/[^a-z0-9]/g, '_');
  return 'FEST_EMAIL_DNS_' + safe.substring(0, 120);
}

function leerCacheDominio_(domain) {
  try {
    const key = claveCacheDominio_(domain);
    const raw = PropertiesService.getScriptProperties().getProperty(key);
    if (!raw) return null;
    const obj = JSON.parse(raw);
    if (!obj || !obj.checkedAt) return null;
    const ageMs = Date.now() - Number(obj.checkedAt);
    const ttlMs = FEST_EMAIL_DOMAIN_CACHE_TTL_HOURS * 60 * 60 * 1000;
    if (ageMs > ttlMs) return null;
    return obj;
  } catch (err) {
    return null;
  }
}

function guardarCacheDominio_(domain, obj) {
  try {
    const key = claveCacheDominio_(domain);
    PropertiesService.getScriptProperties().setProperty(key, JSON.stringify(obj));
  } catch (err) {
    // no-op
  }
}

function resolverDominioPorDns_(domain) {
  const out = {
    checked: true,
    hasMx: false,
    hasA: false,
    status: 'ERROR',
    checkedAt: Date.now()
  };

  try {
    out.hasMx = dnsTieneRespuesta_(domain, 'MX');
    if (!out.hasMx) out.hasA = dnsTieneRespuesta_(domain, 'A');

    if (out.hasMx) out.status = 'MX_OK';
    else if (out.hasA) out.status = 'A_OK';
    else out.status = 'NO_DNS';
  } catch (err) {
    out.status = 'ERROR';
  }

  return out;
}

function dnsTieneRespuesta_(domain, type) {
  const url = 'https://dns.google/resolve?name=' + encodeURIComponent(domain) + '&type=' + encodeURIComponent(type);
  const res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true, followRedirects: true });
  if (res.getResponseCode() !== 200) return false;
  const body = JSON.parse(res.getContentText() || '{}');
  return Array.isArray(body.Answer) && body.Answer.length > 0;
}

function escribirPestanaRevisionEmails_(ss, records, domainBatch, webBudget, source) {
  const sh = asegurarPestanaRevisionEmails_(ss);
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  records.sort((a, b) => {
    if (a.sheetName === b.sheetName) return a.rowIndex - b.rowIndex;
    return a.sheetName < b.sheetName ? -1 : 1;
  });

  const rows = [];
  for (let i = 0; i < records.length; i++) {
    const rec = records[i];
    rows.push([
      now,
      rec.sheetName,
      rec.rowIndex,
      rec.festival,
      rec.emailBefore,
      rec.emailAfter,
      rec.status,
      rec.duplicate ? 'SI' : 'NO',
      rec.taxonomyMismatch ? 'NO' : 'SI',
      rec.dnsSummary,
      rec.webEvidence,
      rec.webEmails,
      rec.action
    ]);
  }

  sh.getRange(1, 1, 1, 14).setValues([[
    'FECHA_REVISION',
    'PESTANA',
    'FILA',
    'FESTIVAL',
    'EMAIL_ANTERIOR',
    'EMAIL_FINAL',
    'ESTADO_EMAIL',
    'DUPLICADO',
    'GENERO_AFORO_OK',
    'DNS',
    'WEB_EVIDENCIA',
    'EMAILS_WEB',
    'ACCION',
    'FUENTE'
  ]]);

  if (sh.getMaxRows() > 1) {
    sh.getRange(2, 1, sh.getMaxRows() - 1, 14).clearContent().clearFormat();
  }

  if (rows.length) {
    // Columna FUENTE
    for (let i = 0; i < rows.length; i++) rows[i][13] = source || 'manual';
    sh.getRange(2, 1, rows.length, 14).setValues(rows);
    sh.getRange(2, 7, rows.length, 1).setDataValidation(buildEmailReviewValidationRule_());

    const colors = [];
    for (let i = 0; i < rows.length; i++) {
      const color = colorPorEstadoRevisionEmail_(rows[i][6]);
      const line = [];
      for (let c = 0; c < 14; c++) line.push(color);
      colors.push(line);
    }
    sh.getRange(2, 1, rows.length, 14).setBackgrounds(colors);
  }

  sh.setFrozenRows(1);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.getRange(1, 1, Math.max(2, rows.length + 1), 14).createFilter();
  sh.setColumnWidth(1, 170);
  sh.setColumnWidth(2, 140);
  sh.setColumnWidth(3, 70);
  sh.setColumnWidth(4, 250);
  sh.setColumnWidth(5, 250);
  sh.setColumnWidth(6, 250);
  sh.setColumnWidth(7, 130);
  sh.setColumnWidth(8, 90);
  sh.setColumnWidth(9, 130);
  sh.setColumnWidth(10, 210);
  sh.setColumnWidth(11, 270);
  sh.setColumnWidth(12, 300);
  sh.setColumnWidth(13, 350);
  sh.setColumnWidth(14, 120);

  sh.getRange(1, 16, 7, 2).setValues([
    ['METRICA', 'VALOR'],
    ['DNS_FETCHED', domainBatch.fetched],
    ['DNS_CACHED', domainBatch.cached],
    ['DNS_SKIPPED', domainBatch.skipped],
    ['WEB_FETCHED', webBudget.webFetched],
    ['WEB_CACHED', webBudget.webCached],
    ['SEARCH_FETCHED', webBudget.searchFetched]
  ]);
}

function asegurarPestanaRevisionEmails_(ss) {
  let sh = ss.getSheetByName(FEST_EMAIL_REVIEW_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(FEST_EMAIL_REVIEW_SHEET_NAME);
  }
  return sh;
}
