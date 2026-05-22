// JK Grader Apps Script backend.
const CONFIG = {
  SPREADSHEET_ID: '1_xOjYXh_PQzXERVyp-k_lyge9jAxiWTPAedS6lOdl0c',
  SHEETS: {
    SHOW_NIGHT: 'Show Night Resposnes',
    PEER_GRADES: 'peer-grade-responses',
    ROSTER: 'Roster',
    GROUP_GRADES: 'Group Grades',
    INDIVIDUAL: 'Individual Grades',
    REFLECTIONS: 'Reflection Links',
    ROLLUP: 'Grade Rollup',
    RUBRIC_CONFIG: 'Rubric Config'
  },
  SHOW_NIGHT_HEADERS: {
    TIMESTAMP: 'Timestamp',
    EMAIL: 'Email Address',
    PERIOD: 'Period',
    COUNTRY_P5: 'Period 5 Country Name',
    COUNTRY_P6: 'Period 6 Country Name',
    PHOTO_1: 'Upload 1: Your Wow factor pose',
    PHOTO_2: 'Upload 2: Front photo',
    PHOTO_3: 'Upload 3: Side photo',
    PHOTO_4: 'Upload 4: Back Photo',
    BLURB: 'Program Blurb Paragraph'
  },
  PEER_HEADERS: {
    SUBMITTER: 'SubmitterName',
    PERIOD: 'Period',
    COUNTRY: 'Country'
  },
  ROSTER_HEADERS: {
    LAST_NAME: 'Last Name',
    PREF_FIRST: 'Preferred First',
    LEGAL: 'Legal First',
    PERIOD: 'Period',
    COUNTRY: 'Country',
    EMAIL: 'Email'
  },
  RUBRIC_HEADERS: {
    CATEGORY: 'Category',
    GRADE_TYPE: 'Grade Type',
    MAX_POINTS: 'Max Points',
    SCORING_TYPE: 'Scoring Type',
    CRITERION_LABEL: 'Criterion Label',
    CRITERION_TEXT: 'Criterion Text'
  },
  GROUP_GRADE_HEADERS: [
    'GroupKey', 'Period', 'Country', 'GroupWorkScore', 'GroupWorkComment',
    'OutfitQualityScore', 'OutfitQualityComment', 'MetaphorScore', 'MetaphorComment',
    'GroupStatus', 'LastSavedAt'
  ],
  INDIVIDUAL_HEADERS: [
    'Email', 'StudentName', 'Period', 'Country',
    'EffortScore', 'EffortComment',
    'ProfessionalismScore', 'ProfessionalismComment', 'ShowNightRole',
    'ExtraCreditScore', 'ExtraCreditNote',
    'GroupWorkOverride', 'GroupWorkOverrideNote',
    'OutfitQualityOverride', 'OutfitQualityOverrideNote',
    'MetaphorOverride', 'MetaphorOverrideNote',
    'ManualScoreOverride', 'ManualScoreOverrideNote',
    'IndividualStatus', 'LastSavedAt',
    'EmailSent', 'EmailSentAt'
  ],
  REFLECTION_HEADERS: [
    'Email', 'StudentName', 'Period', 'Country', 'ReflectionLink', 'SubmittedAt'
  ],
  ROLLUP_HEADERS: [
    'Email', 'StudentName', 'Period', 'Country',
    'GroupWorkFinal', 'OutfitQualityFinal', 'MetaphorFinal',
    'GroupTotal', 'EffortScore', 'ProfessionalismScore', 'ExtraCreditScore',
    'IndividualTotal', 'FinalScore', 'FinalPercent',
    'EmailSent', 'EmailSentAt'
  ],
  TEST_EMAIL_RECIPIENT: 'ktinsley@pausd.org',
  DEFAULT_SIGNOFF: 'Mr. Tinsley',
  STATUS: {
    UNGRADED: 'Ungraded',
    IN_PROGRESS: 'In Progress',
    COMPLETE: 'Complete'
  },
  CACHE: {
    RUBRIC_CONFIG: 'jkRubricConfig',
    ROSTER: 'jkRoster',
    TTL_SECONDS: 3600
  }
};

function doGet(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || '').trim();
    let payload;
    if (action === 'getAppData') payload = getAppData_(e.parameter || {});
    else if (action === 'getGroupDetails') payload = getGroupDetails_(e.parameter || {});
    else if (action === 'getRubricConfig') payload = { ok: true, rubric: getRubricConfig_() };
    else if (action === 'getGradeRollup') payload = getGradeRollup_();
    else if (action === 'getRubricEmailPreview') payload = getRubricEmailPreview_(e.parameter || {});
    else if (action === 'setupSheets') payload = setupSheets_();
    else payload = { ok: true, message: 'JK Show and Dress Grader API', actions: ['getAppData', 'getGroupDetails', 'getRubricConfig', 'getGradeRollup', 'getRubricEmailPreview', 'setupSheets'] };
    return jsonResponse_(payload);
  } catch (err) {
    logError_('doGet', err, e && e.parameter);
    return jsonResponse_(errorResponse_(err));
  }
}

function doPost(e) {
  try {
    const payload = parsePostPayload_(e);
    const action = String(payload.action || '').trim();
    let response;
    if (action === 'saveGroupGrades') response = saveGroupGrades_(payload);
    else if (action === 'saveIndividualGrades') response = saveIndividualGrades_(payload);
    else if (action === 'sendRubricEmail') response = sendRubricEmail_(payload);
    else if (action === 'batchSendEmails') response = batchSendEmails_(payload);
    else if (action === 'updateGradeRollup') response = updateGradeRollup_();
    else throw new Error('Unknown POST action: ' + action);
    return jsonResponse_(response);
  } catch (err) {
    logError_('doPost', err, e && e.postData && e.postData.contents);
    return jsonResponse_(errorResponse_(err));
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('JK Grader')
    .addItem('Setup grader sheets', 'setupSheets_')
    .addItem('Update grade rollup', 'updateGradeRollup_')
    .addToUi();
}

function getAppData_(params) {
  const data = readWorkbook_();
  const rubric = getRubricConfig_();
  const period = normalizePeriodFilter_(params.period);
  const groups = buildGroups_(data, period);
  return {
    ok: true,
    rubric,
    groups,
    summary: summarizeGroups_(groups),
    spreadsheetUrl: data.ss.getUrl(),
    scriptUrl: 'https://script.google.com/home/projects/1Wj32Xppq-IlyxQ2p--guCnTnxc3eva7ckX8DYT6NsvtVMYX4i62KWWcF/edit',
    generatedAt: new Date().toISOString()
  };
}

function getGroupDetails_(params) {
  const groupKey = String(params.groupKey || '').trim();
  if (!groupKey) throw new Error('Missing groupKey.');
  const data = readWorkbook_();
  const group = buildGroups_(data, 'All').find(item => item.groupKey === groupKey);
  if (!group) throw new Error('Group not found: ' + groupKey);
  return { ok: true, group, rubric: getRubricConfig_() };
}

function getGradeRollup_() {
  const rows = computeRollupRows_();
  return { ok: true, rows, generatedAt: new Date().toISOString() };
}

function getRubricEmailPreview_(params) {
  const email = String(params.email || '').trim();
  if (!email) throw new Error('Missing email.');
  const context = getStudentEmailContext_(email);
  const html = buildEmailHtml_(context.student, context.groupGrades, context.rubric, params.signoff || CONFIG.DEFAULT_SIGNOFF);
  return {
    ok: true,
    to: context.student.email,
    subject: buildEmailSubject_(context.student),
    html,
    htmlContent: html,
    emailAlreadySent: asBool_(context.student.EmailSent),
    emailSentAt: context.student.EmailSentAt || ''
  };
}

function saveGroupGrades_(payload) {
  const lock = getWriteLock_();
  try {
    lock.waitLock(30000);
    const groupKey = String(payload.groupKey || '').trim();
    if (!groupKey) throw new Error('Missing groupKey.');
    const ss = getSpreadsheet_();
    const sheet = getRequiredSheet_(ss, CONFIG.SHEETS.GROUP_GRADES);
    requireHeaders_(sheet, CONFIG.GROUP_GRADE_HEADERS);
    const rowNumber = findOrCreateRow_(sheet, 'GroupKey', groupKey);
    const headerMap = getHeaderMap_(sheet);
    const headers = Object.keys(headerMap).sort((a, b) => headerMap[a] - headerMap[b]);
    const rowValues = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
    const period = payload.period || groupKey.split('|')[0];
    const country = payload.country || groupKey.split('|').slice(1).join('|');
    mergeRowValue_(rowValues, headerMap, 'GroupKey', groupKey);
    mergeRowValue_(rowValues, headerMap, 'Period', period);
    mergeRowValue_(rowValues, headerMap, 'Country', country);
    ['GroupWorkScore', 'GroupWorkComment', 'OutfitQualityScore', 'OutfitQualityComment', 'MetaphorScore', 'MetaphorComment'].forEach(header => {
      if (Object.prototype.hasOwnProperty.call(payload, header)) mergeRowValue_(rowValues, headerMap, header, payload[header]);
    });
    const savedAt = new Date().toISOString();
    mergeRowValue_(rowValues, headerMap, 'LastSavedAt', savedAt);
    sheet.getRange(rowNumber, 1, 1, headers.length).setValues([rowValues]);
    return { ok: true, savedAt, group: buildTargetedGroupResponse_(ss, groupKey, { light: true }).group };
  } catch (err) {
    logError_('saveGroupGrades_', err, payload);
    return errorResponse_(err);
  } finally {
    try { lock.releaseLock(); } catch (releaseErr) {}
  }
}

function saveIndividualGrades_(payload) {
  const lock = getWriteLock_();
  try {
    lock.waitLock(30000);
    const email = String(payload.email || '').trim();
    if (!email) throw new Error('Missing email.');
    const ss = getSpreadsheet_();
    const sheet = getRequiredSheet_(ss, CONFIG.SHEETS.INDIVIDUAL);
    requireHeaders_(sheet, CONFIG.INDIVIDUAL_HEADERS);
    const headerMap = getHeaderMap_(sheet);
    const rowNumber = findSingleExactRow_(sheet, headerMap, 'Email', email);
    const headers = Object.keys(headerMap).sort((a, b) => headerMap[a] - headerMap[b]);
    const rowValues = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
    const rowEmail = String(rowValues[headerMap.Email - 1] || '').trim();
    const overrideFields = ['GroupWorkOverride', 'GroupWorkOverrideNote', 'OutfitQualityOverride', 'OutfitQualityOverrideNote', 'MetaphorOverride', 'MetaphorOverrideNote']
      .filter(header => Object.prototype.hasOwnProperty.call(payload, header));
    Logger.log('saveIndividualGrades_ incoming payload.email=%s StudentName=%s Period=%s Country=%s overrideFieldsPresent=%s matchedRowIndex=%s rowEmailBeingWritten=%s',
      payload.email,
      payload.StudentName,
      payload.Period,
      payload.Country,
      JSON.stringify(overrideFields),
      rowNumber,
      rowEmail);
    mergeRowValue_(rowValues, headerMap, 'Email', email);
    ['StudentName', 'Period', 'Country', 'EffortScore', 'EffortComment', 'ProfessionalismScore',
     'ProfessionalismComment', 'ShowNightRole', 'ExtraCreditScore', 'ExtraCreditNote',
     'GroupWorkOverride', 'GroupWorkOverrideNote', 'OutfitQualityOverride', 'OutfitQualityOverrideNote',
     'MetaphorOverride', 'MetaphorOverrideNote',
     'ManualScoreOverride', 'ManualScoreOverrideNote',
     'EmailSent', 'EmailSentAt'].forEach(header => {
      if (Object.prototype.hasOwnProperty.call(payload, header)) mergeRowValue_(rowValues, headerMap, header, payload[header]);
    });
    const updated = rowArrayToObject_(headers, rowValues);
    const savedAt = new Date().toISOString();
    mergeRowValue_(rowValues, headerMap, 'IndividualStatus', computeIndividualStatus_(updated));
    mergeRowValue_(rowValues, headerMap, 'LastSavedAt', savedAt);
    sheet.getRange(rowNumber, 1, 1, headers.length).setValues([rowValues]);
    const period = String(payload.Period || updated.Period || '').trim();
    const country = String(payload.Country || updated.Country || '').trim();
    return { ok: true, savedAt, group: buildTargetedGroupResponse_(ss, `${period}|${country}`, { light: true }).group };
  } catch (err) {
    logError_('saveIndividualGrades_', err, payload);
    return errorResponse_(err);
  } finally {
    try { lock.releaseLock(); } catch (releaseErr) {}
  }
}

function sendRubricEmail_(payload) {
  const email = String(payload.email || '').trim();
  if (!email) throw new Error('Missing email.');
  if (!payload.confirmed) throw new Error('Email send requires confirmed=true.');
  const context = getStudentEmailContext_(email);
  if (asBool_(context.student.EmailSent) && !payload.force) {
    return { ok: false, alreadySent: true, emailSentAt: context.student.EmailSentAt || '', message: 'Email already sent. Confirm resend to continue.' };
  }
  const html = buildEmailHtml_(context.student, context.groupGrades, context.rubric, payload.signoff || CONFIG.DEFAULT_SIGNOFF);
  const to = payload.testMode ? CONFIG.TEST_EMAIL_RECIPIENT : context.student.email;
  if (!payload.dryRun) {
    GmailApp.sendEmail(to, buildEmailSubject_(context.student), 'Please see the HTML version of this email for your JK grade breakdown.', { htmlBody: html });
    if (!payload.testMode) markEmailSent_(context.student.email);
  }
  Logger.log('sendRubricEmail dryRun=%s testMode=%s to=%s student=%s', !!payload.dryRun, !!payload.testMode, to, context.student.name);
  return { ok: true, dryRun: !!payload.dryRun, testMode: !!payload.testMode, to, sentAt: new Date().toISOString() };
}

function batchSendEmails_(payload) {
  if (!payload.confirmed) throw new Error('Batch send requires confirmed=true.');
  const emails = Array.isArray(payload.emails) ? payload.emails : [];
  if (!emails.length) throw new Error('No emails selected.');
  const results = [];
  emails.forEach(email => {
    try {
      const result = sendRubricEmail_({
        email,
        confirmed: true,
        force: !!payload.force,
        dryRun: !!payload.dryRun,
        testMode: !!payload.testMode,
        signoff: payload.signoff || CONFIG.DEFAULT_SIGNOFF
      });
      results.push({ email, ok: result.ok, message: result.message || '', to: result.to || '' });
    } catch (err) {
      results.push({ email, ok: false, message: err.message });
    }
  });
  return { ok: true, results, sentCount: results.filter(item => item.ok).length };
}

function updateGradeRollup_() {
  const lock = getWriteLock_();
  try {
    lock.waitLock(30000);
    const ss = getSpreadsheet_();
    const sheet = getRequiredSheet_(ss, CONFIG.SHEETS.ROLLUP);
    requireHeaders_(sheet, CONFIG.ROLLUP_HEADERS);
    const rows = computeRollupRows_();
    sheet.clearContents();
    sheet.getRange(1, 1, 1, CONFIG.ROLLUP_HEADERS.length).setValues([CONFIG.ROLLUP_HEADERS]);
    if (rows.length) {
      const values = rows.map(row => CONFIG.ROLLUP_HEADERS.map(header => row[header] === undefined ? '' : row[header]));
      sheet.getRange(2, 1, values.length, CONFIG.ROLLUP_HEADERS.length).setValues(values);
    }
    return { ok: true, rowsWritten: rows.length, updatedAt: new Date().toISOString() };
  } catch (err) {
    logError_('updateGradeRollup_', err);
    return errorResponse_(err);
  } finally {
    try { lock.releaseLock(); } catch (releaseErr) {}
  }
}

function setupSheets_() {
  const ss = getSpreadsheet_();
  const specs = [
    { name: CONFIG.SHEETS.GROUP_GRADES, headers: CONFIG.GROUP_GRADE_HEADERS },
    { name: CONFIG.SHEETS.INDIVIDUAL, headers: CONFIG.INDIVIDUAL_HEADERS },
    { name: CONFIG.SHEETS.REFLECTIONS, headers: CONFIG.REFLECTION_HEADERS },
    { name: CONFIG.SHEETS.ROLLUP, headers: CONFIG.ROLLUP_HEADERS }
  ];
  const results = specs.map(spec => {
    const sheet = ss.getSheetByName(spec.name) || ss.insertSheet(spec.name);
    const firstRow = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0].join('').trim();
    if (!firstRow) {
      sheet.clearContents();
      sheet.getRange(1, 1, 1, spec.headers.length).setValues([spec.headers]);
      return { sheet: spec.name, action: 'created headers' };
    }
    return { sheet: spec.name, action: 'left existing headers' };
  });
  const cache = CacheService.getScriptCache();
  cache.remove(CONFIG.CACHE.RUBRIC_CONFIG);
  cache.remove(CONFIG.CACHE.ROSTER);
  return { ok: true, results };
}

function readWorkbook_() {
  const ss = getSpreadsheet_();
  const roster = getCachedRoster_(ss);
  const groupGrades = tabToObjects_(getRequiredSheet_(ss, CONFIG.SHEETS.GROUP_GRADES));
  const individualGrades = tabToObjects_(getRequiredSheet_(ss, CONFIG.SHEETS.INDIVIDUAL));
  const showNight = tabToObjects_(getRequiredSheet_(ss, CONFIG.SHEETS.SHOW_NIGHT));
  const peerResponses = tabToObjects_(getRequiredSheet_(ss, CONFIG.SHEETS.PEER_GRADES));
  const reflections = tabToObjects_(getRequiredSheet_(ss, CONFIG.SHEETS.REFLECTIONS));
  return {
    ss,
    roster,
    groupGrades,
    individualGrades,
    showNight,
    peerResponses,
    reflections,
    groupByKey: indexBy_(groupGrades, 'GroupKey'),
    individualByEmail: indexBy_(individualGrades, 'Email')
  };
}

function buildGroups_(data, periodFilter) {
  const buckets = {};
  data.roster.forEach(row => {
    const period = cleanPeriod_(row[CONFIG.ROSTER_HEADERS.PERIOD]);
    const country = stringOrBlank_(row[CONFIG.ROSTER_HEADERS.COUNTRY]);
    const email = stringOrBlank_(row[CONFIG.ROSTER_HEADERS.EMAIL]);
    if (!period || !country || !email) return;
    if (periodFilter !== 'All' && String(period) !== String(periodFilter)) return;
    const groupKey = `${period}|${country}`;
    if (!buckets[groupKey]) buckets[groupKey] = { groupKey, period, country, students: [] };
    const studentName = `${stringOrBlank_(row[CONFIG.ROSTER_HEADERS.PREF_FIRST]) || stringOrBlank_(row[CONFIG.ROSTER_HEADERS.LEGAL])} ${stringOrBlank_(row[CONFIG.ROSTER_HEADERS.LAST_NAME])}`.trim();
    const individual = data.individualByEmail[email] || {};
    buckets[groupKey].students.push({
      email,
      preferredFirst: stringOrBlank_(row[CONFIG.ROSTER_HEADERS.PREF_FIRST]),
      lastName: stringOrBlank_(row[CONFIG.ROSTER_HEADERS.LAST_NAME]),
      legalFirst: stringOrBlank_(row[CONFIG.ROSTER_HEADERS.LEGAL]),
      name: studentName,
      individualGrades: individual,
      reflection: data.reflections.find(item => String(item.Email || '').toLowerCase() === email.toLowerCase()) || null
    });
  });

  return Object.keys(buckets).sort(sortGroupKeys_).map(groupKey => {
    const group = buckets[groupKey];
    group.students.sort((a, b) => String(a.lastName).localeCompare(String(b.lastName)) || String(a.preferredFirst).localeCompare(String(b.preferredFirst)));
    group.groupGrades = data.groupByKey[groupKey] || null;
    group.showNightData = findShowNight_(data.showNight, group.period, group.country);
    group.peerResponses = data.peerResponses.filter(row => cleanPeriod_(row[CONFIG.PEER_HEADERS.PERIOD]) === String(group.period) && sameText_(row[CONFIG.PEER_HEADERS.COUNTRY], group.country));
    group.status = computeGroupStatus_(groupKey, group.groupGrades, data.individualByEmail, group.students);
    group.studentPeerData = {};
    group.students.forEach(student => {
      group.studentPeerData[student.email] = findPeerForStudent_(group.peerResponses, student);
    });
    return group;
  });
}

function studentsForGroup_(rosterRows, period, country) {
  return rosterRows.filter(row => {
    return String(cleanPeriod_(row[CONFIG.ROSTER_HEADERS.PERIOD])) === String(period) &&
      sameText_(row[CONFIG.ROSTER_HEADERS.COUNTRY], country);
  }).map(row => ({ email: stringOrBlank_(row[CONFIG.ROSTER_HEADERS.EMAIL]) })).filter(student => student.email);
}

function findShowNight_(rows, period, country) {
  const match = rows.find(row => {
    const rowPeriod = cleanPeriod_(row[CONFIG.SHOW_NIGHT_HEADERS.PERIOD]);
    const rowCountry = activeShowCountry_(row);
    return String(rowPeriod) === String(period) && sameText_(rowCountry, country);
  });
  if (!match) return null;
  return {
    timestamp: match[CONFIG.SHOW_NIGHT_HEADERS.TIMESTAMP] || '',
    email: match[CONFIG.SHOW_NIGHT_HEADERS.EMAIL] || '',
    period,
    country: activeShowCountry_(match),
    photos: [
      { label: 'Wow pose', url: match[CONFIG.SHOW_NIGHT_HEADERS.PHOTO_1] || '' },
      { label: 'Front', url: match[CONFIG.SHOW_NIGHT_HEADERS.PHOTO_2] || '' },
      { label: 'Side', url: match[CONFIG.SHOW_NIGHT_HEADERS.PHOTO_3] || '' },
      { label: 'Back', url: match[CONFIG.SHOW_NIGHT_HEADERS.PHOTO_4] || '' }
    ].filter(photo => String(photo.url || '').trim()),
    blurb: match[CONFIG.SHOW_NIGHT_HEADERS.BLURB] || ''
  };
}

function findPeerForStudent_(responses, student) {
  const first = normalizeName_(student.preferredFirst);
  const last = normalizeName_(student.lastName);
  const full = normalizeName_(`${student.preferredFirst} ${student.lastName}`);

  function isStudentMatch(name) {
    const norm = normalizeName_(name);
    if (!norm) return false;
    return norm === first || norm === last || norm === full ||
      (first && norm.indexOf(first) >= 0) ||
      (last && norm.indexOf(last) >= 0);
  }

  let submissionRow = null;
  for (let i = 0; i < responses.length; i++) {
    if (isStudentMatch(responses[i][CONFIG.PEER_HEADERS.SUBMITTER] || '')) {
      submissionRow = responses[i];
      break;
    }
  }

  const allocations = [];
  if (submissionRow) {
    for (let n = 1; n <= 4; n++) {
      const name = String(submissionRow[`Peer${n}Name`] || '').trim();
      if (!name) continue;
      allocations.push({
        name,
        percent: submissionRow[`Peer${n}Percent`] || '',
        justification: submissionRow[`Peer${n}Justification`] || ''
      });
    }
  }

  const receivedGrades = [];
  for (let i = 0; i < responses.length; i++) {
    const row = responses[i];
    for (let n = 1; n <= 4; n++) {
      if (isStudentMatch(row[`Peer${n}Name`] || '')) {
        receivedGrades.push({
          submitter: row[CONFIG.PEER_HEADERS.SUBMITTER] || '',
          percent: row[`Peer${n}Percent`] || '',
          justification: row[`Peer${n}Justification`] || ''
        });
        break;
      }
    }
  }

  if (!submissionRow && !receivedGrades.length) return null;
  return {
    submitter: submissionRow ? (submissionRow[CONFIG.PEER_HEADERS.SUBMITTER] || '') : '',
    allocations,
    receivedGrades
  };
}

function getRubricConfig_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CONFIG.CACHE.RUBRIC_CONFIG);
  if (cached) return JSON.parse(cached);
  const rows = tabToObjects_(getRequiredSheet_(getSpreadsheet_(), CONFIG.SHEETS.RUBRIC_CONFIG));
  const rubric = {};
  let lastCategory = '';
  let lastMax = '';
  let lastScoring = '';
  rows.forEach(row => {
    const category = stringOrBlank_(row[CONFIG.RUBRIC_HEADERS.CATEGORY]) || lastCategory;
    if (!category) return;
    const maxCell = row[CONFIG.RUBRIC_HEADERS.MAX_POINTS];
    const scoringCell = row[CONFIG.RUBRIC_HEADERS.SCORING_TYPE];
    if (maxCell !== '' && maxCell !== null && maxCell !== undefined) lastMax = maxCell;
    if (scoringCell) lastScoring = scoringCell;
    lastCategory = category;
    if (!rubric[category]) rubric[category] = { maxPoints: Number(lastMax || 0), scoringType: stringOrBlank_(lastScoring), criteria: [] };
    if (lastMax !== '') rubric[category].maxPoints = Number(lastMax);
    if (lastScoring) rubric[category].scoringType = stringOrBlank_(lastScoring);
    const label = stringOrBlank_(row[CONFIG.RUBRIC_HEADERS.CRITERION_LABEL]);
    const text = stringOrBlank_(row[CONFIG.RUBRIC_HEADERS.CRITERION_TEXT]);
    if (label || text) rubric[category].criteria.push({ label, text });
  });
  try {
    cache.put(CONFIG.CACHE.RUBRIC_CONFIG, JSON.stringify(rubric), CONFIG.CACHE.TTL_SECONDS);
  } catch (err) {
    Logger.log('CacheService put failed for %s: %s', CONFIG.CACHE.RUBRIC_CONFIG, err && err.message ? err.message : err);
  }
  return rubric;
}

function getCachedRoster_(ss) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(CONFIG.CACHE.ROSTER);
  if (cached) return JSON.parse(cached);
  const rows = tabToObjects_(getRequiredSheet_(ss, CONFIG.SHEETS.ROSTER));
  try {
    cache.put(CONFIG.CACHE.ROSTER, JSON.stringify(rows), CONFIG.CACHE.TTL_SECONDS);
  } catch (err) {
    Logger.log('CacheService put failed for %s: %s', CONFIG.CACHE.ROSTER, err && err.message ? err.message : err);
  }
  return rows;
}

function computeRollupRows_() {
  const data = readWorkbook_();
  const groups = buildGroups_(data, 'All');
  const rows = [];
  groups.forEach(group => {
    const gg = group.groupGrades || {};
    group.students.forEach(student => {
      const ind = data.individualByEmail[student.email] || {};
      const groupWork = finalGroupScore_(ind.GroupWorkOverride, gg.GroupWorkScore);
      const outfit = finalGroupScore_(ind.OutfitQualityOverride, gg.OutfitQualityScore);
      const metaphor = finalGroupScore_(ind.MetaphorOverride, gg.MetaphorScore);
      const groupTotal = sumNums_([groupWork, outfit, metaphor]);
      const effort = numOrBlank_(ind.EffortScore);
      const prof = numOrBlank_(ind.ProfessionalismScore);
      const extra = numOrZero_(ind.ExtraCreditScore);
      const individualTotal = sumNums_([effort, prof]);
      const calculatedScore = blankIfIncomplete_([groupWork, outfit, metaphor, effort, prof], (groupTotal * 0.75) + (individualTotal * 0.25) + extra);
      const manualOverride = numOrBlank_(ind.ManualScoreOverride);
      const finalScore = manualOverride !== '' ? manualOverride : calculatedScore;
      rows.push({
        Email: student.email,
        StudentName: student.name,
        Period: group.period,
        Country: group.country,
        GroupWorkFinal: groupWork,
        OutfitQualityFinal: outfit,
        MetaphorFinal: metaphor,
        GroupTotal: groupTotal,
        EffortScore: effort,
        ProfessionalismScore: prof,
        ExtraCreditScore: extra,
        IndividualTotal: individualTotal,
        FinalScore: finalScore,
        FinalPercent: finalScore === '' ? '' : finalScore / 100,
        EmailSent: ind.EmailSent || '',
        EmailSentAt: ind.EmailSentAt || ''
      });
    });
  });
  return rows;
}

function getStudentEmailContext_(email) {
  const data = readWorkbook_();
  const groups = buildGroups_(data, 'All');
  const rubric = getRubricConfig_();
  for (let i = 0; i < groups.length; i++) {
    const group = groups[i];
    const student = group.students.find(item => String(item.email).toLowerCase() === email.toLowerCase());
    if (student) {
      const rollup = computeRollupRows_().find(row => String(row.Email).toLowerCase() === email.toLowerCase()) || {};
      return {
        rubric,
        groupGrades: group.groupGrades || {},
        student: Object.assign({}, student.individualGrades || {}, rollup, {
          email: student.email,
          name: student.name,
          country: group.country,
          period: group.period
        })
      };
    }
  }
  throw new Error('Student not found: ' + email);
}

function buildEmailHtml_(student, groupGrades, rubric, signoff) {
  const e = escapeHtmlEmail_;
  const pct = student.FinalPercent === '' ? '' : Math.round(Number(student.FinalPercent) * 100);
  function scoreLine(label, score, max, note) {
    return `<tr><td style="padding:12px 0;border-bottom:1px solid #e5e7eb;"><strong>${e(label)}</strong>${note ? `<p style="margin:6px 0 0;color:#92400e;">Override note: ${e(note)}</p>` : ''}</td><td align="right" style="padding:12px 0;border-bottom:1px solid #e5e7eb;font-weight:800;">${e(displayScore_(score))}/${e(max)}</td></tr>`;
  }
  const teacherNotes = [
    groupGrades.GroupWorkComment,
    groupGrades.OutfitQualityComment,
    groupGrades.MetaphorComment,
    student.EffortComment,
    student.ProfessionalismComment,
    student.ExtraCreditNote
  ].filter(Boolean).map(escapeHtmlEmail_).join('<br><br>');
  return `<!doctype html><html><body style="margin:0;background:#f4f6ff;font-family:Arial,Helvetica,sans-serif;color:#203044;">
  <table width="100%" cellpadding="0" cellspacing="0" style="background:#f4f6ff;padding:28px 12px;"><tr><td align="center">
  <table width="640" cellpadding="0" cellspacing="0" style="max-width:640px;width:100%;background:#ffffff;border-radius:18px;overflow:hidden;">
  <tr><td style="background:#4a40e0;color:#ffffff;padding:28px;"><p style="margin:0 0 6px;font-size:12px;font-weight:800;text-transform:uppercase;letter-spacing:.08em;">Junk Kouture Feedback</p><h1 style="margin:0;font-size:26px;">${e(student.name)}</h1><p style="margin:8px 0 0;">${e(student.country)} · Period ${e(student.period)}</p></td></tr>
  <tr><td style="padding:24px 28px;"><table width="100%" cellpadding="0" cellspacing="0">
  ${scoreLine('Group Work', student.GroupWorkFinal, rubric['Group Work'] ? rubric['Group Work'].maxPoints : 30, student.GroupWorkOverrideNote)}
  ${scoreLine('Outfit Quality', student.OutfitQualityFinal, rubric['Outfit Quality'] ? rubric['Outfit Quality'].maxPoints : 30, student.OutfitQualityOverrideNote)}
  ${scoreLine('Metaphor', student.MetaphorFinal, rubric.Metaphor ? rubric.Metaphor.maxPoints : 15, student.MetaphorOverrideNote)}
  ${scoreLine('Group subtotal', student.GroupTotal, 75, '')}
  ${scoreLine('Effort and Initiative', student.EffortScore, rubric['Effort and Initiative'] ? rubric['Effort and Initiative'].maxPoints : 10, '')}
  ${scoreLine('Professionalism', student.ProfessionalismScore, rubric.Professionalism ? rubric.Professionalism.maxPoints : 15, '')}
  ${scoreLine('Extra Credit', student.ExtraCreditScore, rubric['Extra Credit'] ? rubric['Extra Credit'].maxPoints : 5, '')}
  ${scoreLine('Individual subtotal', student.IndividualTotal, 25, '')}
  </table>
  <div style="margin-top:22px;background:#f4f6ff;border-radius:12px;padding:18px;"><p style="margin:0;color:#4d5d73;font-size:12px;font-weight:800;text-transform:uppercase;">Final Score</p><p style="margin:6px 0 0;font-size:34px;font-weight:800;color:#4a40e0;">${e(displayScore_(student.FinalScore))}/100 ${pct !== '' ? `<span style="font-size:18px;color:#4d5d73;">(${e(pct)}%)</span>` : ''}</p></div>
  ${teacherNotes ? `<div style="margin-top:20px;border-left:4px solid #f59e0b;background:#fffbeb;padding:14px 18px;"><strong>Teacher Comment</strong><p style="white-space:pre-line;line-height:1.55;">${teacherNotes}</p></div>` : ''}
  <p style="margin:24px 0 0;white-space:pre-line;">${e(signoff || CONFIG.DEFAULT_SIGNOFF)}</p>
  </td></tr></table></td></tr></table></body></html>`;
}

function buildEmailSubject_(student) {
  return `Junk Kouture Grade Feedback - ${student.name}`;
}

function markEmailSent_(email) {
  const ss = getSpreadsheet_();
  const sheet = getRequiredSheet_(ss, CONFIG.SHEETS.INDIVIDUAL);
  requireHeaders_(sheet, CONFIG.INDIVIDUAL_HEADERS);
  const rowNumber = findOrCreateRow_(sheet, 'Email', email);
  const headerMap = getHeaderMap_(sheet);
  setCellByHeader_(sheet, headerMap, rowNumber, 'EmailSent', 'TRUE');
  setCellByHeader_(sheet, headerMap, rowNumber, 'EmailSentAt', new Date().toISOString());
}

function parsePostPayload_(e) {
  const contents = e && e.postData && e.postData.contents ? e.postData.contents : '{}';
  return JSON.parse(contents);
}

function jsonResponse_(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload)).setMimeType(ContentService.MimeType.JSON);
}

function getSpreadsheet_() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function getRequiredSheet_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Missing sheet: ' + name);
  return sheet;
}

function tabToObjects_(sheet) {
  const values = sheet.getDataRange().getValues();
  if (!values.length) return [];
  const headers = values[0].map(String);
  return values.slice(1).filter(row => row.some(cell => cell !== '')).map(row => {
    const obj = {};
    headers.forEach((header, i) => obj[header] = row[i]);
    return obj;
  });
}

function getHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const map = {};
  headers.forEach((header, i) => map[header] = i + 1);
  return map;
}

function requireHeaders_(sheet, required) {
  const map = getHeaderMap_(sheet);
  const missing = required.filter(header => !map[header]);
  if (missing.length) throw new Error(`Sheet "${sheet.getName()}" is missing headers: ${missing.join(', ')}`);
}

function findOrCreateRow_(sheet, keyCol, keyVal) {
  const map = getHeaderMap_(sheet);
  const col = map[keyCol];
  if (!col) throw new Error(`Missing key column ${keyCol} in ${sheet.getName()}.`);
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const values = sheet.getRange(2, col, lastRow - 1, 1).getValues();
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0]).trim().toLowerCase() === String(keyVal).trim().toLowerCase()) return i + 2;
    }
  }
  const row = lastRow + 1;
  sheet.getRange(row, col).setValue(keyVal);
  return row;
}

function findSingleExactRow_(sheet, headerMap, keyCol, keyVal) {
  const col = headerMap[keyCol];
  if (!col) throw new Error(`Missing key column ${keyCol} in ${sheet.getName()}.`);
  const exactValue = String(keyVal || '').trim();
  const matches = [];
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const values = sheet.getRange(2, col, lastRow - 1, 1).getValues();
    for (let i = 0; i < values.length; i++) {
      if (String(values[i][0] || '').trim() === exactValue) matches.push(i + 2);
    }
  }
  if (matches.length !== 1) {
    throw new Error(`Expected exactly one ${sheet.getName()} row with ${keyCol}="${exactValue}", found ${matches.length}.`);
  }
  return matches[0];
}

function setCellByHeader_(sheet, headerMap, rowNumber, header, value) {
  if (!headerMap[header]) throw new Error(`Missing column ${header} in ${sheet.getName()}.`);
  sheet.getRange(rowNumber, headerMap[header]).setValue(value === undefined || value === null ? '' : value);
}

function objectFromRow_(sheet, rowNumber) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const values = sheet.getRange(rowNumber, 1, 1, headers.length).getValues()[0];
  const obj = {};
  headers.forEach((header, i) => obj[header] = values[i]);
  return obj;
}

function rowArrayToObject_(headers, values) {
  const obj = {};
  headers.forEach((header, i) => obj[header] = values[i]);
  return obj;
}

function mergeRowValue_(rowValues, headerMap, header, value) {
  if (!headerMap[header]) throw new Error('Missing column ' + header + '.');
  rowValues[headerMap[header] - 1] = value === undefined || value === null ? '' : value;
}

function readIndividualRowsForStudents_(ss, students) {
  const emails = {};
  (students || []).forEach(student => {
    const email = String(student.email || '').trim().toLowerCase();
    if (email) emails[email] = true;
  });
  if (!Object.keys(emails).length) return [];
  const sheet = getRequiredSheet_(ss, CONFIG.SHEETS.INDIVIDUAL);
  requireHeaders_(sheet, CONFIG.INDIVIDUAL_HEADERS);
  const headerMap = getHeaderMap_(sheet);
  const headers = Object.keys(headerMap).sort((a, b) => headerMap[a] - headerMap[b]);
  const emailCol = headerMap.Email;
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  return values.reduce((rows, row) => {
    if (emails[String(row[emailCol - 1] || '').trim().toLowerCase()]) rows.push(rowArrayToObject_(headers, row));
    return rows;
  }, []);
}

function findExistingRowByHeader_(sheet, headerMap, header, value) {
  const col = headerMap[header];
  if (!col) throw new Error('Missing key column ' + header + ' in ' + sheet.getName() + '.');
  const target = String(value || '').trim().toLowerCase();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return null;
  const values = sheet.getRange(2, col, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim().toLowerCase() === target) return i + 2;
  }
  return null;
}

function buildTargetedGroupResponse_(ss, groupKey, options) {
  options = options || {};
  const light = options.light === true;
  const parts = String(groupKey || '').split('|');
  const period = cleanPeriod_(parts[0]);
  const country = parts.slice(1).join('|');
  if (!period || !country) throw new Error('Missing groupKey.');

  const rosterRows = getCachedRoster_(ss).filter(row => {
    return String(cleanPeriod_(row[CONFIG.ROSTER_HEADERS.PERIOD])) === String(period) &&
      sameText_(row[CONFIG.ROSTER_HEADERS.COUNTRY], country);
  });
  if (!rosterRows.length) throw new Error('Group not found: ' + groupKey);

  const groupSheet = getRequiredSheet_(ss, CONFIG.SHEETS.GROUP_GRADES);
  requireHeaders_(groupSheet, CONFIG.GROUP_GRADE_HEADERS);
  const groupHeaderMap = getHeaderMap_(groupSheet);
  const groupHeaders = Object.keys(groupHeaderMap).sort((a, b) => groupHeaderMap[a] - groupHeaderMap[b]);
  const groupRowNumber = findExistingRowByHeader_(groupSheet, groupHeaderMap, 'GroupKey', groupKey);
  const groupGrades = groupRowNumber
    ? rowArrayToObject_(groupHeaders, groupSheet.getRange(groupRowNumber, 1, 1, groupHeaders.length).getValues()[0])
    : null;

  const students = rosterRows.map(row => {
    const email = stringOrBlank_(row[CONFIG.ROSTER_HEADERS.EMAIL]);
    const studentName = `${stringOrBlank_(row[CONFIG.ROSTER_HEADERS.PREF_FIRST]) || stringOrBlank_(row[CONFIG.ROSTER_HEADERS.LEGAL])} ${stringOrBlank_(row[CONFIG.ROSTER_HEADERS.LAST_NAME])}`.trim();
    return {
      email,
      preferredFirst: stringOrBlank_(row[CONFIG.ROSTER_HEADERS.PREF_FIRST]),
      lastName: stringOrBlank_(row[CONFIG.ROSTER_HEADERS.LAST_NAME]),
      legalFirst: stringOrBlank_(row[CONFIG.ROSTER_HEADERS.LEGAL]),
      name: studentName,
      individualGrades: {},
      reflection: null
    };
  }).filter(student => student.email);
  const individualByEmail = indexBy_(readIndividualRowsForStudents_(ss, students), 'Email');
  students.forEach(student => student.individualGrades = individualByEmail[student.email] || {});
  students.sort((a, b) => String(a.lastName).localeCompare(String(b.lastName)) || String(a.preferredFirst).localeCompare(String(b.preferredFirst)));
  const status = computeGroupStatus_(groupKey, groupGrades, individualByEmail, students);
  if (groupRowNumber && groupGrades && groupGrades.GroupStatus !== status) {
    const rowValues = groupSheet.getRange(groupRowNumber, 1, 1, groupHeaders.length).getValues()[0];
    mergeRowValue_(rowValues, groupHeaderMap, 'GroupStatus', status);
    groupSheet.getRange(groupRowNumber, 1, 1, groupHeaders.length).setValues([rowValues]);
    groupGrades.GroupStatus = status;
  }

  let peerResponses = [];
  let showNightData = null;
  if (!light) {
    const reflections = tabToObjects_(getRequiredSheet_(ss, CONFIG.SHEETS.REFLECTIONS));
    students.forEach(student => {
      student.reflection = reflections.find(item => String(item.Email || '').toLowerCase() === student.email.toLowerCase()) || null;
    });
    peerResponses = tabToObjects_(getRequiredSheet_(ss, CONFIG.SHEETS.PEER_GRADES)).filter(row => {
      return cleanPeriod_(row[CONFIG.PEER_HEADERS.PERIOD]) === String(period) && sameText_(row[CONFIG.PEER_HEADERS.COUNTRY], country);
    });
    showNightData = findShowNight_(tabToObjects_(getRequiredSheet_(ss, CONFIG.SHEETS.SHOW_NIGHT)), period, country);
  }
  const group = {
    groupKey,
    period,
    country,
    students,
    groupGrades,
    reflections: light ? null : undefined,
    showNightData,
    peerResponses,
    status,
    studentPeerData: {}
  };
  if (!light) {
    students.forEach(student => {
      group.studentPeerData[student.email] = findPeerForStudent_(peerResponses, student);
    });
  }
  return { ok: true, group };
}

function getWriteLock_() {
  return LockService.getScriptLock();
}

function logError_(fn, err, context) {
  Logger.log('%s failed: %s\nContext: %s\nStack: %s', fn, err && err.message, JSON.stringify(context || {}), err && err.stack);
}

function errorResponse_(err) {
  return { ok: false, message: err && err.message ? err.message : String(err) };
}

function indexBy_(rows, key) {
  const out = {};
  rows.forEach(row => {
    const val = String(row[key] || '').trim();
    if (val) out[val] = row;
  });
  return out;
}

function cleanPeriod_(value) {
  return String(value || '').replace(/^Period\s*/i, '').replace(/^P/i, '').trim();
}

function normalizePeriodFilter_(value) {
  const period = cleanPeriod_(value);
  return period && period !== 'All' ? period : 'All';
}

function stringOrBlank_(value) {
  return value === null || value === undefined ? '' : String(value).trim();
}

function sameText_(a, b) {
  return stringOrBlank_(a).toLowerCase() === stringOrBlank_(b).toLowerCase();
}

function activeShowCountry_(row) {
  return stringOrBlank_(row[CONFIG.SHOW_NIGHT_HEADERS.COUNTRY_P5]) || stringOrBlank_(row[CONFIG.SHOW_NIGHT_HEADERS.COUNTRY_P6]);
}

function normalizeName_(value) {
  return stringOrBlank_(value).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}

function sortGroupKeys_(a, b) {
  const ap = a.split('|')[0], bp = b.split('|')[0];
  if (Number(ap) !== Number(bp)) return Number(ap) - Number(bp);
  return a.split('|').slice(1).join('|').localeCompare(b.split('|').slice(1).join('|'));
}

function computeIndividualStatus_(row) {
  const effort = row.EffortScore !== '' && row.EffortScore !== null && row.EffortScore !== undefined;
  const prof = row.ProfessionalismScore !== '' && row.ProfessionalismScore !== null && row.ProfessionalismScore !== undefined;
  if (effort && prof) return CONFIG.STATUS.COMPLETE;
  if (effort || prof || row.ExtraCreditScore !== '' || row.ShowNightRole) return CONFIG.STATUS.IN_PROGRESS;
  return CONFIG.STATUS.UNGRADED;
}

function computeGroupStatus_(groupKey, groupGrades, individualByEmail, students) {
  groupGrades = groupGrades || {};
  const groupDone = ['GroupWorkScore', 'OutfitQualityScore', 'MetaphorScore'].every(key => groupGrades[key] !== '' && groupGrades[key] !== null && groupGrades[key] !== undefined);
  const rosterStudents = students || [];
  const individualsDone = rosterStudents.length ? rosterStudents.every(student => {
    const row = individualByEmail[student.email] || {};
    return row.EffortScore !== '' && row.EffortScore !== null && row.EffortScore !== undefined &&
           row.ProfessionalismScore !== '' && row.ProfessionalismScore !== null && row.ProfessionalismScore !== undefined;
  }) : false;
  if (groupDone && individualsDone) return CONFIG.STATUS.COMPLETE;
  const hasGroupWork = ['GroupWorkScore', 'OutfitQualityScore', 'MetaphorScore'].some(key => groupGrades[key] !== '' && groupGrades[key] !== null && groupGrades[key] !== undefined);
  return hasGroupWork ? CONFIG.STATUS.IN_PROGRESS : CONFIG.STATUS.UNGRADED;
}

function summarizeGroups_(groups) {
  const complete = groups.filter(group => group.status === CONFIG.STATUS.COMPLETE).length;
  return { total: groups.length, complete, percent: groups.length ? Math.round((complete / groups.length) * 100) : 0 };
}

function numOrBlank_(value) {
  if (value === '' || value === null || value === undefined) return '';
  const num = Number(value);
  return Number.isFinite(num) ? num : '';
}

function numOrZero_(value) {
  const num = numOrBlank_(value);
  return num === '' ? 0 : num;
}

function finalGroupScore_(override, inherited) {
  const over = numOrBlank_(override);
  return over === '' ? numOrBlank_(inherited) : over;
}

function sumNums_(values) {
  return values.reduce((sum, value) => sum + (numOrBlank_(value) === '' ? 0 : Number(value)), 0);
}

function blankIfIncomplete_(required, value) {
  return required.some(item => numOrBlank_(item) === '') ? '' : Math.round(Number(value) * 100) / 100;
}

function displayScore_(value) {
  return value === '' || value === null || value === undefined ? '?' : String(Math.round(Number(value) * 100) / 100);
}

function asBool_(value) {
  return String(value || '').toLowerCase() === 'true' || String(value || '').toLowerCase() === 'yes';
}

function escapeHtmlEmail_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
