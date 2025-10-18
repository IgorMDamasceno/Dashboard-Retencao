const DATA_SHEET = 'BD – GAM';
const OPERATIONS_SHEET = 'Config - Operações';
const TRAFFIC_SHEET = 'Config - Trafego';
const REVSHARE_SHEET = 'Config - RevShare';
const TRAFFIC_TYPES = ['Automação', 'Broadcast', 'Push'];

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Dashboard de Retenção')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getConfigData() {
  const snapshot = getConfigSnapshot_();
  return {
    operations: snapshot.operationRows,
    traffics: snapshot.trafficRows,
    revshare: snapshot.revshareRows,
    trafficTypes: TRAFFIC_TYPES.slice()
  };
}

function saveOperationConfig(payload) {
  if (!payload) throw new Error('Dados inválidos.');
  const sheet = ensureOperationsSheet_();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = getHeaderIndexMap_(headers);
  if (!payload.operation || !payload.url) {
    throw new Error('Informe operação e URL.');
  }
  var id = payload.id || Utilities.getUuid();
  var data = {
    id: id,
    operation: String(payload.operation).trim(),
    url: String(payload.url).trim(),
    note: payload.note ? String(payload.note).trim() : ''
  };
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) lastRow = 1;
  var range = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  var rowIndex = -1;
  for (var i = 0; i < range.length; i++) {
    if (String(range[i][idx.id]) === id) {
      rowIndex = i + 2;
      break;
    }
  }
  var rowValues = [
    data.id,
    data.operation,
    data.url,
    data.note
  ];
  if (rowIndex > -1) {
    sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
  refreshRevShareSites_();
  return getConfigData();
}

function deleteOperationConfig(id) {
  if (!id) return getConfigData();
  const sheet = ensureOperationsSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return getConfigData();
  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = getHeaderIndexMap_(headers);
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][idx.id]) === String(id)) {
      sheet.deleteRow(i + 2);
      break;
    }
  }
  refreshRevShareSites_();
  return getConfigData();
}

function saveTrafficConfig(payload) {
  if (!payload) throw new Error('Dados inválidos.');
  const sheet = ensureTrafficSheet_();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = getHeaderIndexMap_(headers);
  if (!payload.type || TRAFFIC_TYPES.indexOf(payload.type) === -1) {
    throw new Error('Selecione um tipo de tráfego válido.');
  }
  if (!payload.source) {
    throw new Error('Informe a utm_source.');
  }
  var id = payload.id || Utilities.getUuid();
  var rowValues = [
    id,
    String(payload.type).trim(),
    String(payload.source).trim()
  ];
  var lastRow = sheet.getLastRow();
  var range = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues() : [];
  var rowIndex = -1;
  for (var i = 0; i < range.length; i++) {
    if (String(range[i][idx.id]) === id) {
      rowIndex = i + 2;
      break;
    }
  }
  if (rowIndex > -1) {
    sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
  return getConfigData();
}

function deleteTrafficConfig(id) {
  if (!id) return getConfigData();
  const sheet = ensureTrafficSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return getConfigData();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = getHeaderIndexMap_(headers);
  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][idx.id]) === String(id)) {
      sheet.deleteRow(i + 2);
      break;
    }
  }
  return getConfigData();
}

function updateRevShare(payload) {
  if (!payload || !payload.site) return getConfigData();
  const sheet = ensureRevShareSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return getConfigData();
  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
  var updated = false;
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(payload.site)) {
      sheet.getRange(i + 2, 2).setValue(payload.revshare != null && payload.revshare !== '' ? Number(payload.revshare) : '');
      updated = true;
      break;
    }
  }
  if (!updated) {
    sheet.appendRow([payload.site, payload.revshare != null ? Number(payload.revshare) : '']);
  }
  return getConfigData();
}

function getDashboardData(params) {
  if (!params) throw new Error('Parâmetros inválidos.');
  const snapshot = getConfigSnapshot_();
  const operationName = String(params.operation || '').trim();
  const trafficType = String(params.traffic || '').trim();
  if (!operationName) throw new Error('Selecione a operação.');
  if (!trafficType) throw new Error('Selecione o tráfego.');
  const operationUrls = snapshot.operationsByName[operationName];
  if (!operationUrls || !operationUrls.length) {
    throw new Error('Operação sem URLs configuradas.');
  }
  const trafficSources = snapshot.trafficByType[trafficType] || [];
  if (!trafficSources.length) {
    throw new Error('Tráfego sem utm_source configurada.');
  }
  const startDate = params.startDate ? parseDate_(params.startDate) : null;
  const endDate = params.endDate ? parseDate_(params.endDate) : null;
  const window = params.urlWindow || 'day';
  if (startDate && endDate && startDate.getTime() > endDate.getTime()) {
    throw new Error('A data inicial não pode ser maior que a final.');
  }
  const baseRows = loadBaseData_(operationUrls, trafficSources);
  if (!baseRows.length) {
    return buildEmptyDashboardResponse_();
  }
  const filteredRows = filterRowsByDate_(baseRows, startDate, endDate);
  if (!filteredRows.length) {
    return buildEmptyDashboardResponse_();
  }
  const hoursInfo = getHourTimeline_(filteredRows);
  if (!hoursInfo.allHours.length) {
    return buildEmptyDashboardResponse_();
  }
  const siteSummary = buildSiteSummary_(filteredRows, hoursInfo, snapshot.revshareMap);
  const urlRows = filterRowsForWindow_(filteredRows, hoursInfo, window);
  const urlSummary = buildUrlSummary_(urlRows, snapshot.revshareMap);
  return {
    filters: {
      operation: operationName,
      traffic: trafficType,
      startDate: startDate ? formatDateISO_(startDate) : '',
      endDate: endDate ? formatDateISO_(endDate) : '',
      window: window,
      latestHour: hoursInfo.latestHourLabel,
      hours: hoursInfo.windowHours.map(function (h) { return h.label; })
    },
    site: siteSummary,
    url: urlSummary
  };
}

function buildEmptyDashboardResponse_() {
  return {
    filters: {},
    site: {
      hours: [],
      rows: [],
      totals: null
    },
    url: {
      rows: [],
      totals: null,
      windowLabel: ''
    }
  };
}

function loadBaseData_(operationUrls, trafficSources) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(DATA_SHEET);
  if (!sheet) throw new Error('Aba "' + DATA_SHEET + '" não encontrada.');
  const range = sheet.getDataRange();
  const values = range.getValues();
  if (!values.length) return [];
  const headers = values[0].map(function (h) { return String(h).trim(); });
  const idx = getHeaderIndexMap_(headers);
  const dataRows = [];
  const urlMatchers = buildUrlMatchers_(operationUrls);
  const sourceSet = buildLowerSet_(trafficSources);
  const idxData = findIndex_(idx, ['data']);
  const idxHour = findIndex_(idx, ['hora']);
  const idxSite = findIndex_(idx, ['site']);
  const idxChannel = findIndex_(idx, ['canal']);
  const idxUrl = findIndex_(idx, ['url']);
  const idxBlock = findIndex_(idx, ['bloco', 'blocodeanuncio']);
  const idxRequests = findIndex_(idx, ['solicitacoes', 'solicitacaoad']);
  const idxRevenue = findIndex_(idx, ['receitausd', 'receita']);
  const idxCoverage = findIndex_(idx, ['cobertura']);
  const idxEcpm = findIndex_(idx, ['ecpm']);
  if (idxData == null || idxHour == null || idxSite == null || idxUrl == null) {
    throw new Error('Planilha base está sem colunas obrigatórias.');
  }
  if (idxRequests == null || idxRevenue == null) {
    throw new Error('Planilha base está sem colunas de solicitações ou receita.');
  }
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var rawUrl = idxUrl != null ? row[idxUrl] : '';
    if (!matchUrl_(rawUrl, urlMatchers)) continue;
    var canal = idxChannel != null ? String(row[idxChannel] || '').trim() : '';
    if (sourceSet.size && !sourceSet.has(canal.toLowerCase())) continue;
    var date = idxData != null ? parseDate_(row[idxData]) : null;
    if (!date) continue;
    var hour = idxHour != null ? parseHour_(row[idxHour]) : null;
    if (hour == null) continue;
    var site = idxSite != null ? String(row[idxSite] || '').trim() : '';
    if (!site) continue;
    var block = idxBlock != null ? String(row[idxBlock] || '').trim() : '';
    var requests = idxRequests != null ? toNumber_(row[idxRequests]) : 0;
    var revenue = idxRevenue != null ? toNumber_(row[idxRevenue]) : 0;
    var coverage = idxCoverage != null ? toNumber_(row[idxCoverage]) : 0;
    var ecpm = idxEcpm != null ? toNumber_(row[idxEcpm]) : 0;
    var timestamp = new Date(date.getFullYear(), date.getMonth(), date.getDate(), hour).getTime();
    dataRows.push({
      date: date,
      hour: hour,
      hourLabel: formatHour_(hour),
      timestamp: timestamp,
      site: site,
      url: String(rawUrl || '').trim(),
      normUrl: normUrlStrict_(rawUrl),
      block: block,
      isInterstitial: /interstitial/i.test(block),
      requests: requests,
      revenue: revenue,
      coverage: coverage,
      ecpm: ecpm
    });
  }
  return dataRows;
}

function filterRowsByDate_(rows, startDate, endDate) {
  if (!rows.length) return [];
  var filtered = rows.slice();
  if (startDate) {
    var startValue = startDate.getTime();
    filtered = filtered.filter(function (row) {
      return startOfDay_(row.date).getTime() >= startValue;
    });
  }
  if (endDate) {
    var endValue = endDate.getTime();
    filtered = filtered.filter(function (row) {
      return startOfDay_(row.date).getTime() <= endValue;
    });
  }
  if (!startDate && !endDate && filtered.length) {
    var latest = filtered.reduce(function (acc, row) {
      var dayValue = startOfDay_(row.date).getTime();
      return dayValue > acc ? dayValue : acc;
    }, 0);
    filtered = filtered.filter(function (row) {
      return startOfDay_(row.date).getTime() === latest;
    });
  }
  return filtered;
}

function getHourTimeline_(rows) {
  var unique = {};
  var list = [];
  rows.forEach(function (row) {
    if (!unique[row.timestamp]) {
      unique[row.timestamp] = true;
      list.push(row.timestamp);
    }
  });
  list.sort(function (a, b) { return a - b; });
  if (!list.length) {
    return {
      allHours: [],
      windowHours: [],
      latestHour: null,
      latestHourLabel: ''
    };
  }
  var latestHour = list[list.length - 1];
  var windowCandidates = list.slice(-4);
  if (windowCandidates.length && windowCandidates[windowCandidates.length - 1] === latestHour) {
    windowCandidates.pop();
  }
  var windowHours = windowCandidates.map(function (stamp) {
    return {
      key: stamp,
      label: formatHourFromTimestamp_(stamp)
    };
  });
  return {
    allHours: list,
    windowHours: windowHours,
    latestHour: latestHour,
    latestHourLabel: formatHourFromTimestamp_(latestHour)
  };
}

function buildSiteSummary_(rows, hoursInfo, revshareMap) {
  var siteMap = {};
  var hourKeys = hoursInfo.windowHours.map(function (h) { return h.key; });
  rows.forEach(function (row) {
    var siteKey = row.site;
    if (!siteMap[siteKey]) {
      siteMap[siteKey] = {
        site: siteKey,
        totals: { sessions: 0, revenue: 0 },
        hours: {}
      };
    }
    var siteData = siteMap[siteKey];
    var revshare = toNumber_(revshareMap[siteKey] || 0) / 100;
    var adjustedRevenue = row.revenue * (1 - revshare);
    if (row.isInterstitial) {
      siteData.totals.sessions += row.requests;
    }
    siteData.totals.revenue += adjustedRevenue;
    if (!siteData.hours[row.timestamp]) {
      siteData.hours[row.timestamp] = { sessions: 0, revenue: 0 };
    }
    if (row.isInterstitial) {
      siteData.hours[row.timestamp].sessions += row.requests;
    }
    siteData.hours[row.timestamp].revenue += adjustedRevenue;
  });

  var rowsOut = [];
  var totals = {
    sessions: {},
    revenue: {},
    totalSessions: 0,
    totalRevenue: 0
  };
  hourKeys.forEach(function (key) {
    totals.sessions[key] = 0;
    totals.revenue[key] = 0;
  });

  Object.keys(siteMap).sort().forEach(function (siteKey) {
    var siteData = siteMap[siteKey];
    var rowOut = {
      site: siteKey,
      sessions: {},
      revenue: {},
      rps: {},
      totalSessions: siteData.totals.sessions,
      totalRevenue: siteData.totals.revenue,
      totalRps: computeRps_(siteData.totals.revenue, siteData.totals.sessions)
    };
    hourKeys.forEach(function (key) {
      var hourData = siteData.hours[key] || { sessions: 0, revenue: 0 };
      rowOut.sessions[key] = hourData.sessions;
      rowOut.revenue[key] = hourData.revenue;
      rowOut.rps[key] = computeRps_(hourData.revenue, hourData.sessions);
      totals.sessions[key] += hourData.sessions;
      totals.revenue[key] += hourData.revenue;
    });
    totals.totalSessions += siteData.totals.sessions;
    totals.totalRevenue += siteData.totals.revenue;
    rowsOut.push(rowOut);
  });

  var totalsRow = {
    sessions: {},
    revenue: {},
    rps: {},
    totalSessions: totals.totalSessions,
    totalRevenue: totals.totalRevenue,
    totalRps: computeRps_(totals.totalRevenue, totals.totalSessions)
  };
  hourKeys.forEach(function (key) {
    totalsRow.sessions[key] = totals.sessions[key];
    totalsRow.revenue[key] = totals.revenue[key];
    totalsRow.rps[key] = computeRps_(totals.revenue[key], totals.sessions[key]);
  });

  return {
    hours: hoursInfo.windowHours,
    rows: rowsOut,
    totals: rowsOut.length ? totalsRow : null
  };
}

function filterRowsForWindow_(rows, hoursInfo, window) {
  if (!rows.length) return [];
  var latestHour = hoursInfo.latestHour;
  if (!latestHour) return [];
  var filtered = [];
  var lowerBound;
  var includeDayOnly = false;
  if (window === 'last6h') {
    lowerBound = latestHour - 6 * 3600000;
  } else if (window === 'last3h') {
    lowerBound = latestHour - 3 * 3600000;
  } else {
    includeDayOnly = true;
  }
  rows.forEach(function (row) {
    if (row.timestamp >= latestHour) return;
    if (includeDayOnly) {
      var latestDate = new Date(latestHour);
      var rowDate = row.date;
      if (rowDate.getFullYear() === latestDate.getFullYear() &&
          rowDate.getMonth() === latestDate.getMonth() &&
          rowDate.getDate() === latestDate.getDate()) {
        filtered.push(row);
      }
      return;
    }
    if (row.timestamp > lowerBound && row.timestamp < latestHour) {
      filtered.push(row);
    }
  });
  return filtered;
}

function buildUrlSummary_(rows, revshareMap) {
  var map = {};
  rows.forEach(function (row) {
    var key = row.normUrl || row.url;
    if (!map[key]) {
      map[key] = {
        url: row.url,
        site: row.site,
        sessions: 0,
        revenue: 0,
        weightedRequests: 0,
        coverageWeighted: 0,
        ecpmWeighted: 0
      };
    }
    var data = map[key];
    var revshare = toNumber_(revshareMap[row.site] || 0) / 100;
    var adjustedRevenue = row.revenue * (1 - revshare);
    data.revenue += adjustedRevenue;
    var rowCoverageRatio = row.coverage > 1 ? row.coverage / 100 : row.coverage;
    if (row.isInterstitial) {
      data.sessions += row.requests;
      data.coverageWeighted += rowCoverageRatio * row.requests;
      data.ecpmWeighted += (row.ecpm * (1 - revshare)) * row.requests;
      data.weightedRequests += row.requests;
    }
  });
  var rowsOut = [];
  var totals = {
    sessions: 0,
    revenue: 0,
    weightedRequests: 0,
    coverageWeighted: 0,
    ecpmWeighted: 0
  };
  Object.keys(map).map(function (key) {
    return map[key];
  }).sort(function (a, b) {
    return b.revenue - a.revenue;
  }).forEach(function (data) {
    var avgCoverageRatio = data.weightedRequests ? data.coverageWeighted / data.weightedRequests : 0;
    var coverage = avgCoverageRatio * 100;
    var ecpm = data.weightedRequests ? data.ecpmWeighted / data.weightedRequests : 0;
    rowsOut.push({
      url: data.url,
      site: data.site,
      sessions: data.sessions,
      revenue: data.revenue,
      rps: computeRps_(data.revenue, data.sessions),
      coverage: coverage,
      ecpm: ecpm,
      effectiveEcpm: ecpm * avgCoverageRatio
    });
    totals.sessions += data.sessions;
    totals.revenue += data.revenue;
    totals.weightedRequests += data.weightedRequests;
    totals.coverageWeighted += data.coverageWeighted;
    totals.ecpmWeighted += data.ecpmWeighted;
  });
  var totalsRow = null;
  if (rowsOut.length) {
    var totalsCoverageRatio = totals.weightedRequests ? totals.coverageWeighted / totals.weightedRequests : 0;
    var totalsCoverage = totalsCoverageRatio * 100;
    var totalsEcpm = totals.weightedRequests ? totals.ecpmWeighted / totals.weightedRequests : 0;
    totalsRow = {
      sessions: totals.sessions,
      revenue: totals.revenue,
      rps: computeRps_(totals.revenue, totals.sessions),
      coverage: totalsCoverage,
      ecpm: totalsEcpm,
      effectiveEcpm: totalsEcpm * totalsCoverageRatio
    };
  }
  return {
    rows: rowsOut,
    totals: totalsRow
  };
}

function getConfigSnapshot_() {
  const operationSheet = ensureOperationsSheet_();
  const trafficSheet = ensureTrafficSheet_();
  const revSheet = ensureRevShareSheet_();

  refreshRevShareSites_();

  const operationValues = getSheetValues_(operationSheet);
  const trafficValues = getSheetValues_(trafficSheet);
  const revValues = getSheetValues_(revSheet);

  const operationRows = operationValues.rows.map(function (row) {
    return {
      id: row.id,
      operation: row.operacao,
      url: row.url,
      note: row.observacao
    };
  });

  const operationsByName = {};
  operationRows.forEach(function (row) {
    var name = row.operation;
    if (!name) return;
    if (!operationsByName[name]) operationsByName[name] = [];
    operationsByName[name].push(row.url);
  });

  const trafficRows = trafficValues.rows.map(function (row) {
    return {
      id: row.id,
      type: row.tipo,
      source: row.utm_source
    };
  });
  const trafficByType = {};
  trafficRows.forEach(function (row) {
    if (!row.type || !row.source) return;
    if (!trafficByType[row.type]) trafficByType[row.type] = [];
    trafficByType[row.type].push(row.source);
  });

  const revshareRows = revValues.rows.map(function (row) {
    return {
      site: row.site,
      revshare: row.revshare
    };
  });
  const revshareMap = {};
  revshareRows.forEach(function (row) {
    if (row.site) revshareMap[row.site] = row.revshare;
  });

  return {
    operationRows: operationRows,
    operationsByName: operationsByName,
    trafficRows: trafficRows,
    trafficByType: trafficByType,
    revshareRows: revshareRows,
    revshareMap: revshareMap
  };
}

function getSheetValues_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return { headers: [], rows: [] };
  }
  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = values[0].map(function (h) { return String(h).trim(); });
  var idx = getHeaderIndexMap_(headers);
  var rows = [];
  for (var i = 1; i < values.length; i++) {
    var rowValues = values[i];
    if (rowValues.join('').trim() === '') continue;
    rows.push({
      id: pickValue_(rowValues, idx, 'id'),
      operacao: pickValue_(rowValues, idx, 'operacao'),
      url: pickValue_(rowValues, idx, 'url'),
      observacao: pickValue_(rowValues, idx, 'observacao'),
      tipo: pickValue_(rowValues, idx, 'tipo'),
      utm_source: pickValue_(rowValues, idx, 'utmsource'),
      site: pickValue_(rowValues, idx, 'site'),
      revshare: pickValue_(rowValues, idx, 'revshare')
    });
  }
  return { headers: headers, rows: rows };
}

function ensureOperationsSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(OPERATIONS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(OPERATIONS_SHEET);
  }
  var headers = ['ID', 'Operação', 'URL', 'Observação'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  return sheet;
}

function ensureTrafficSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(TRAFFIC_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(TRAFFIC_SHEET);
  }
  var headers = ['ID', 'Tipo', 'utm_source'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  return sheet;
}

function ensureRevShareSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(REVSHARE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(REVSHARE_SHEET);
  }
  var headers = ['Site', 'RevShare (%)'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  return sheet;
}

function refreshRevShareSites_() {
  var sheet = ensureRevShareSheet_();
  var config = getSheetValues_(ensureOperationsSheet_());
  var existing = getSheetValues_(sheet);
  var revshareMap = {};
  existing.rows.forEach(function (row) {
    if (row.site) revshareMap[row.site] = row.revshare;
  });
  var sites = {};
  config.rows.forEach(function (row) {
    if (row.url) {
      var site = extractDomain_(row.url);
      if (site) sites[site] = true;
    }
  });
  var siteList = Object.keys(sites).sort();
  var values = [['Site', 'RevShare (%)']];
  siteList.forEach(function (site) {
    values.push([site, revshareMap[site] != null ? revshareMap[site] : '']);
  });
  sheet.clearContents();
  if (values.length) {
    sheet.getRange(1, 1, values.length, 2).setValues(values);
  }
  sheet.setFrozenRows(1);
}

function buildUrlMatchers_(urls) {
  return urls.map(function (url) {
    var raw = String(url || '').trim();
    var wildcard = /\*$/.test(raw);
    var normalized = normUrlStrict_(raw.replace(/\*+$/, ''));
    return {
      raw: raw,
      wildcard: wildcard,
      normalized: normalized
    };
  }).filter(function (item) { return !!item.normalized; });
}

function matchUrl_(url, matchers) {
  var normalized = normUrlStrict_(url);
  if (!normalized) return false;
  for (var i = 0; i < matchers.length; i++) {
    var matcher = matchers[i];
    if (matcher.wildcard) {
      if (normalized.indexOf(matcher.normalized) === 0) return true;
    } else if (normalized === matcher.normalized) {
      return true;
    }
  }
  return false;
}

function buildLowerSet_(items) {
  var set = new Set();
  items.forEach(function (item) {
    if (item != null && item !== '') {
      set.add(String(item).trim().toLowerCase());
    }
  });
  return set;
}

function extractDomain_(url) {
  var normalized = normUrlStrict_(url);
  if (!normalized) return '';
  var slash = normalized.indexOf('/');
  return slash === -1 ? normalized : normalized.substring(0, slash);
}

function getHeaderIndexMap_(headers) {
  var map = {};
  headers.forEach(function (header, index) {
    var key = normalizeKey_(header);
    if (key) map[key] = index;
  });
  return map;
}

function normalizeKey_(value) {
  if (value == null) return '';
  return String(value)
    .toLowerCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/g, '');
}

function pickValue_(row, idxMap, key) {
  if (!idxMap) return '';
  var index = idxMap[key];
  if (index == null) return '';
  return row[index];
}

function findIndex_(idxMap, keys) {
  for (var i = 0; i < keys.length; i++) {
    var key = keys[i];
    if (idxMap.hasOwnProperty(key)) return idxMap[key];
  }
  return null;
}

function toNumber_(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val;
  var str = String(val).trim();
  if (!str) return 0;
  str = str.replace(/\s+/g, '').replace(/%/g, '');
  str = str.replace(/\./g, '').replace(/,/g, '.');
  var num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}

function parseDate_(value) {
  if (value instanceof Date && !isNaN(value)) return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  var str = String(value || '').trim();
  if (!str) return null;
  var parts;
  parts = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (parts) {
    return new Date(+parts[1], +parts[2] - 1, +parts[3]);
  }
  parts = str.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (parts) {
    return new Date(+parts[3], +parts[2] - 1, +parts[1]);
  }
  var parsed = new Date(str);
  return isNaN(parsed) ? null : new Date(parsed.getFullYear(), parsed.getMonth(), parsed.getDate());
}

function parseHour_(value) {
  if (value === null || value === undefined || value === '') return null;
  if (typeof value === 'number') {
    var num = Math.floor(value);
    return isNaN(num) ? null : Math.max(0, Math.min(23, num));
  }
  var str = String(value).trim().toLowerCase();
  if (!str) return null;
  str = str.replace(/[h]/g, '');
  if (str.indexOf(':') !== -1) {
    str = str.split(':')[0];
  }
  var num = parseInt(str, 10);
  if (isNaN(num)) return null;
  return Math.max(0, Math.min(23, num));
}

function normUrlStrict_(url) {
  if (!url) return '';
  var str = String(url).trim().toLowerCase();
  str = str.replace(/^https?:\/\//, '');
  str = str.replace(/^www\./, '');
  str = str.split('#')[0];
  str = str.split('?')[0];
  str = str.replace(/\/+$/, '');
  return str;
}

function computeRps_(revenue, sessions) {
  if (!sessions) return 0;
  return revenue / sessions * 1000;
}

function startOfDay_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function formatHour_(hour) {
  var h = ('0' + hour).slice(-2);
  return h + 'h';
}

function formatHourFromTimestamp_(timestamp) {
  var d = new Date(timestamp);
  return ('0' + d.getHours()).slice(-2) + 'h';
}

function formatDateISO_(date) {
  return date.toISOString().slice(0, 10);
}
