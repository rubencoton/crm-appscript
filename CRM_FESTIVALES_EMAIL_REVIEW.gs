function auditarEmailsCRMFestivales() {
  if (!validarSesionSegura_('Auditar emails + duplicados')) return;
  const summary = ejecutarAuditoriaEmailsFestivales_({ source: 'manual', showUi: true });
  SpreadsheetApp.getUi().alert(summary);
}

function auditarEmailsAutomaticaCRMFestivales_() {
  try {
    ejecutarAuditoriaEmailsFestivales_({ source: 'trigger', showUi: false });
  } catch (err) {
    Logger.log('auditarEmailsAutomaticaCRMFestivales_ ERROR: ' + (err && err.message ? err.message : err));
  }
}

function instalarTriggerRevisionEmailsCRMFestivales() {
  const existing = ScriptApp.getProjectTriggers();
  let deleted = 0;
  for (let i = 0; i < existing.length; i++) {
    if (existing[i].getHandlerFunction() === FEST_EMAIL_TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(existing[i]);
      deleted++;
    }
  }

  ScriptApp.newTrigger(FEST_EMAIL_TRIGGER_HANDLER)
    .timeBased()
    .everyHours(FEST_EMAIL_TRIGGER_INTERVAL_HOURS)
    .create();

  SpreadsheetApp.getUi().alert(
    'Trigger de revision de emails instalado.\n' +
    'Frecuencia: cada ' + FEST_EMAIL_TRIGGER_INTERVAL_HOURS + ' horas.\n' +
    'Triggers antiguos eliminados: ' + deleted
  );
}

function limpiarTriggerRevisionEmailsCRMFestivales() {
  const existing = ScriptApp.getProjectTriggers();
  let deleted = 0;
  for (let i = 0; i < existing.length; i++) {
    if (existing[i].getHandlerFunction() === FEST_EMAIL_TRIGGER_HANDLER) {
      ScriptApp.deleteTrigger(existing[i]);
      deleted++;
    }
  }

  SpreadsheetApp.getUi().alert('Trigger revision emails eliminados: ' + deleted);
}

function ejecutarAuditoriaEmailsFestivales_(opts) {
  const options = opts || {};
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return 'Otro proceso esta en curso. Intenta en unos segundos.';
  }

  try {
    const started = Date.now();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = getFestivalSheets_(ss);
    if (!sheets.length) return 'No encontre pestanas de festivales para revisar.';

    const contexts = [];
    const records = [];
    const tokenCounter = {};
    const uniqueDomains = {};
    let reviewColumnCreated = 0;

    for (let s = 0; s < sheets.length; s++) {
      const sheet = sheets[s];
      const init = asegurarColumnaRevisionEmailEnSheet_(sheet);
      if (init.created) reviewColumnCreated++;

      const values = sheet.getDataRange().getValues();
      if (values.length < 2) {
        contexts.push({
          sheet: sheet,
          map: init.map,
          rowsCount: 0,
          reviewOut: []
        });
        continue;
      }

      const map = buildHeaderMap_(values[0]);
      const rowsCount = values.length - 1;
      const reviewCol = map.reviewEmail + 1;
      const currentReview = sheet.getRange(2, reviewCol, rowsCount, 1).getValues();
      const reviewOut = [];

      for (let r = 0; r < rowsCount; r++) {
        reviewOut.push([normalizeRevisionStatus_(currentReview[r][0])]);
      }

      contexts.push({
        sheet: sheet,
        map: map,
        rowsCount: rowsCount,
        reviewOut: reviewOut
      });

      for (let r = 1; r < values.length; r++) {
        const row = values[r];
        const festival = cleanText_(valueAt_(row, map.nombre));
        if (!festival) continue;

        const email = normalizeEmailCell_(valueAt_(row, map.email));
        const tokens = extraerEmailsValidos_(email);
        const domains = extraerDominiosDeEmails_(tokens);
        for (let d = 0; d < domains.length; d++) uniqueDomains[domains[d]] = true;
        for (let t = 0; t < tokens.length; t++) {
          tokenCounter[tokens[t]] = (tokenCounter[tokens[t]] || 0) + 1;
        }

        records.push({
          sheetName: sheet.getName(),
          rowIndex: r + 1,
          rowOffset: r - 1,
          festival: festival,
          email: email,
          tokens: tokens,
          domains: domains,
          currentReview: normalizeRevisionStatus_(valueAt_(row, map.reviewEmail)),
          autoStatus: 'PENDIENTE_REVISION',
          finalReview: 'PENDIENTE_REVISION',
          duplicate: false,
          dnsSummary: '',
          recommendedAction: ''
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

    const domainBatch = obtenerEstadoDominiosEnLote_(
      Object.keys(uniqueDomains),
      FEST_EMAIL_MAX_DOMAIN_CHECKS_PER_RUN
    );

    const bySheet = {};
    for (let c = 0; c < contexts.length; c++) {
      bySheet[contexts[c].sheet.getName()] = contexts[c];
    }

    let duplicatesCount = 0;
    let noEmailCount = 0;
    let fixCount = 0;
    let reviewCount = 0;

    for (let i = 0; i < records.length; i++) {
      const rec = records[i];
      const result = calcularEstadoAutoRevisionEmail_(rec, domainBatch.checks);
      rec.autoStatus = result.status;
      rec.dnsSummary = result.dnsSummary;
      rec.recommendedAction = result.action;

      if (rec.duplicate) duplicatesCount++;
      if (rec.autoStatus === 'SIN_EMAIL') noEmailCount++;
      if (rec.autoStatus === 'CORREGIR_EMAIL') fixCount++;
      if (rec.autoStatus === 'REVISAR_WEB') reviewCount++;

      const ctx = bySheet[rec.sheetName];
      if (!ctx || rec.rowOffset < 0 || rec.rowOffset >= ctx.reviewOut.length) continue;
      const current = normalizeRevisionStatus_(ctx.reviewOut[rec.rowOffset][0]);

      // Si ya esta verificado manualmente, no sobreescribir.
      if (current === 'OK_VERIFICADO_WEB') {
        rec.finalReview = current;
      } else {
        ctx.reviewOut[rec.rowOffset][0] = rec.autoStatus;
        rec.finalReview = rec.autoStatus;
      }
    }

    for (let c = 0; c < contexts.length; c++) {
      const ctx = contexts[c];
      if (!ctx.rowsCount) continue;
      const reviewCol = ctx.map.reviewEmail + 1;
      ctx.sheet.getRange(2, reviewCol, ctx.rowsCount, 1).setValues(ctx.reviewOut);
      ctx.sheet.getRange(2, reviewCol, ctx.rowsCount, 1).setDataValidation(buildEmailReviewValidationRule_());
    }

    escribirPestanaRevisionEmails_(ss, records, domainBatch, options.source || 'manual');

    const elapsedSec = Math.round((Date.now() - started) / 1000);
    const summary = [
      'Revision emails completada (' + FEST_ARCHITECT + ')',
      '',
      'Pestanas de festivales: ' + sheets.length,
      'Filas analizadas: ' + records.length,
      'Columna "REVISION EMAIL" creada: ' + reviewColumnCreated,
      'Duplicados detectados: ' + duplicatesCount,
      'Sin email: ' + noEmailCount,
      'Para corregir: ' + fixCount,
      'Para revisar web: ' + reviewCount,
      'Dominios consultados en esta ejecucion: ' + domainBatch.fetched,
      'Dominios desde cache: ' + domainBatch.cached,
      'Dominios pendientes por limite: ' + domainBatch.skipped,
      'Tiempo total: ' + elapsedSec + 's',
      '',
      'Nota: no se usa Gemini aqui (control de coste API).'
    ].join('\n');

    if (!options.showUi) Logger.log(summary);
    return summary;
  } finally {
    lock.releaseLock();
  }
}

function asegurarColumnaRevisionEmailEnSheet_(sheet) {
  const currentMaxCol = Math.max(sheet.getLastColumn(), FEST_HOMO.HEADER.length);
  const firstRow = sheet.getRange(1, 1, 1, currentMaxCol).getValues()[0];
  const map = buildHeaderMap_(firstRow);
  const reviewCol = map.reviewEmail + 1;
  let created = false;

  if (sheet.getMaxColumns() < reviewCol) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), reviewCol - sheet.getMaxColumns());
    created = true;
  }

  const headerCell = cleanText_(sheet.getRange(1, reviewCol).getValue()).toUpperCase();
  if (headerCell !== 'REVISION EMAIL') {
    sheet.getRange(1, reviewCol).setValue('REVISION EMAIL');
    created = true;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, reviewCol, lastRow - 1, 1).setDataValidation(buildEmailReviewValidationRule_());
  }

  if (created) {
    try {
      applyVisualDesignToSheet_(sheet);
    } catch (err) {
      // no-op
    }
  }

  return { map: buildHeaderMap_(sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), FEST_HOMO.HEADER.length)).getValues()[0]), created: created };
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

  if (current && web.matchCurrent) {
    out.status = 'BIEN';
    out.finalEmail = current;
    out.action = 'Email confirmado en web.';
  } else if (best) {
    out.finalEmail = best;
    out.status = (current && best === current) ? 'BIEN' : 'CAMBIADO';
    out.action = out.status === 'CAMBIADO'
      ? 'Email actualizado segun evidencia web.'
      : 'Email confirmado en web.';
  } else {
    out.status = 'MAL';
    out.finalEmail = current || '';
    out.action = current
      ? 'No se pudo validar ese email en la web.'
      : 'No hay email y no se encontro uno fiable en web.';
  }

  if (record.duplicate) {
    if (out.status === 'CAMBIADO') {
      const cnt = counter[out.finalEmail] || 0;
      if (cnt > 1) {
        out.status = 'MAL';
        out.action = appendReason_(out.action, 'Sigue duplicado tras el cambio.');
      } else {
        out.action = appendReason_(out.action, 'Cambio aplicado para resolver duplicado.');
      }
    } else {
      out.status = 'MAL';
      out.action = appendReason_(out.action, 'Email duplicado en CRM.');
    }
  }

  if (!out.finalEmail) out.status = 'MAL';
  return out;
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
    if (html) cache.put(key, html, FEST_EMAIL_WEB_CACHE_TTL_SEC);
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
  if (s === 'CAMBIADO') return '#D9E8FB';
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

function escribirPestanaRevisionEmails_(ss, records, domainBatch, source) {
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
      rec.email,
      rec.autoStatus,
      rec.finalReview,
      rec.duplicate ? 'SI' : 'NO',
      rec.domains.join('; '),
      rec.dnsSummary,
      rec.recommendedAction
    ]);
  }

  sh.getRange(1, 1, 1, 12).setValues([[
    'FECHA_REVISION',
    'PESTANA',
    'FILA',
    'FESTIVAL',
    'EMAIL',
    'ESTADO_AUTO',
    'ESTADO_REVISION',
    'DUPLICADO',
    'DOMINIOS',
    'DNS',
    'ACCION',
    'FUENTE'
  ]]);

  if (sh.getMaxRows() > 1) {
    sh.getRange(2, 1, sh.getMaxRows() - 1, 12).clearContent();
  }

  if (rows.length) {
    // Columna FUENTE
    for (let i = 0; i < rows.length; i++) rows[i][11] = source || 'manual';
    sh.getRange(2, 1, rows.length, 12).setValues(rows);
    sh.getRange(2, 7, rows.length, 1).setDataValidation(buildEmailReviewValidationRule_());
  }

  sh.setFrozenRows(1);
  if (sh.getFilter()) sh.getFilter().remove();
  sh.getRange(1, 1, Math.max(2, rows.length + 1), 12).createFilter();
  sh.setColumnWidth(1, 170);
  sh.setColumnWidth(2, 130);
  sh.setColumnWidth(3, 70);
  sh.setColumnWidth(4, 260);
  sh.setColumnWidth(5, 260);
  sh.setColumnWidth(6, 170);
  sh.setColumnWidth(7, 190);
  sh.setColumnWidth(8, 90);
  sh.setColumnWidth(9, 220);
  sh.setColumnWidth(10, 260);
  sh.setColumnWidth(11, 320);
  sh.setColumnWidth(12, 120);
}

function asegurarPestanaRevisionEmails_(ss) {
  let sh = ss.getSheetByName(FEST_EMAIL_REVIEW_SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(FEST_EMAIL_REVIEW_SHEET_NAME);
  }
  return sh;
}
