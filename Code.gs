const DATA_SHEET = 'BD – GAM';
const OPERATIONS_SHEET = 'Config - Operações';
const OPERATION_INTEGRATION_SHEET = 'Config - Operação API';
const TRAFFIC_SHEET = 'Config - Trafego';
const REVSHARE_SHEET = 'Config - RevShare';
const LINK_ROUTER_HISTORY_SHEET = 'Histórico - Link Router';
const LINK_ROUTER_BASE_URL = 'https://rotasv2.spun.com.br';
const LINK_ROUTER_TOKEN = '1|crewZvM3pCrHsz9gwWVfzyzQ5IrTk0T7gltREA4d67dc7105';
const TRAFFIC_TYPES = ['Automação', 'Broadcast', 'Push'];
const DISTRIBUTION_STATE_SHEET = 'Controle - Distribuição';
const DISTRIBUTION_CONFIG_SHEET = 'Config - Distribuição';
const DISTRIBUTION_SCORE_SMOOTHING = 800;
const DISTRIBUTION_MIN_SESSION_SHARE = 0.01;
const DISTRIBUTION_MOMENTUM_WEIGHT = 0.15;
const DISTRIBUTION_MOMENTUM_CLAMP = 0.6;
const DISTRIBUTION_COVERAGE_TARGET = 0.75;
const DISTRIBUTION_COVERAGE_FLOOR = 0.5;
const DISTRIBUTION_COVERAGE_CAP = 0.97;
const DISTRIBUTION_COVERAGE_BONUS = 0.25;

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
    operationIntegrations: snapshot.operationIntegrationRows,
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

function saveOperationIntegrationConfig(payload) {
  if (!payload) throw new Error('Dados inválidos.');
  var operation = String(payload.operation || '').trim();
  if (!operation) {
    throw new Error('Selecione a operação.');
  }
  var trafficType = String(payload.trafficType || '').trim();
  if (!trafficType || TRAFFIC_TYPES.indexOf(trafficType) === -1) {
    throw new Error('Selecione o tipo de tráfego.');
  }
  var companyId = String(payload.companyId || '').trim();
  var domainId = String(payload.domainId || '').trim();
  var routeSlug = String(payload.routeSlug || '').trim();
  var active = !!payload.active;
  if (active && (!companyId || !domainId || !routeSlug)) {
    throw new Error('Informe Empresa ID, Domínio ID e Slug da rota para ativar.');
  }
  const sheet = ensureOperationIntegrationSheet_();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = getHeaderIndexMap_(headers);
  const slugIndex = idx.slugrota;
  var id = String(payload.id || '').trim();
  var lastRow = sheet.getLastRow();
  var range = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues() : [];
  var targetRow = -1;
  if (id) {
    for (var i = 0; i < range.length; i++) {
      if (String(range[i][idx.id] || '').trim() === id) {
        targetRow = i + 2;
        break;
      }
    }
  }
  var normalizedOp = operation.toLowerCase();
  var normalizedTraffic = trafficType.toLowerCase();
  var normalizedSlug = routeSlug.toLowerCase();
  for (var j = 0; j < range.length; j++) {
    var existingId = String(range[j][idx.id] || '').trim();
    var existingOp = String(range[j][idx.operacao] || '').trim().toLowerCase();
    var existingTraffic = String(range[j][idx.tipotrafego] || '').trim().toLowerCase();
    var existingSlug = slugIndex == null ? '' : String(range[j][slugIndex] || '').trim().toLowerCase();
    var isSameRecord = existingId && id && existingId === id;
    if (!isSameRecord && existingOp === normalizedOp && existingTraffic === normalizedTraffic && existingSlug === normalizedSlug) {
      throw new Error('Já existe uma integração para esta operação, tráfego e slug de rota.');
    }
  }
  if (!id) {
    id = Utilities.getUuid();
  }
  var statusValue = active ? 'Ativo' : 'Inativo';
  var rowValues = [id, operation, trafficType, companyId, domainId, routeSlug, statusValue];
  if (targetRow > -1) {
    sheet.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
  updateLinkRouterAutoSyncTrigger_();
  return getConfigData();
}

function deleteOperationIntegrationConfig(id) {
  if (!id) return getConfigData();
  const sheet = ensureOperationIntegrationSheet_();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return getConfigData();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = getHeaderIndexMap_(headers);
  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][idx.id] || '').trim() === String(id)) {
      sheet.deleteRow(i + 2);
      break;
    }
  }
  updateLinkRouterAutoSyncTrigger_();
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
  const requestedRouteSlug = String(params.routeSlug || '').trim();
  const opTrafficKey = operationName + '\u0000' + trafficType;
  const integrationCandidates = (snapshot.operationIntegrationsByOperationTraffic && snapshot.operationIntegrationsByOperationTraffic[opTrafficKey]) ? snapshot.operationIntegrationsByOperationTraffic[opTrafficKey] : [];
  var integrationConfig = null;
  var selectedRouteSlug = requestedRouteSlug;
  var normalizedRequestedRoute = requestedRouteSlug.toLowerCase();
  if (normalizedRequestedRoute) {
    integrationConfig = integrationCandidates.filter(function (row) {
      return String(row.routeSlug || '').trim().toLowerCase() === normalizedRequestedRoute;
    })[0] || null;
    if (!integrationConfig) {
      throw new Error('Integração Link Router não encontrada para o slug informado.');
    }
    selectedRouteSlug = integrationConfig.routeSlug || '';
  } else if (integrationCandidates.length === 1) {
    integrationConfig = integrationCandidates[0];
    selectedRouteSlug = integrationConfig.routeSlug || '';
  } else if (integrationCandidates.length > 1) {
    throw new Error('Mais de uma rota configurada para esta operação e tráfego. Selecione o slug desejado.');
  }
  const operationUrls = snapshot.operationsByName[operationName];
  if (!operationUrls || !operationUrls.length) {
    throw new Error('Operação sem URLs configuradas.');
  }
  const trafficSources = snapshot.trafficByType[trafficType] || [];
  if (!trafficSources.length) {
    throw new Error('Tráfego sem utm_source configurada.');
  }
  var startDate = params.startDate ? parseDate_(params.startDate) : null;
  var endDate = params.endDate ? parseDate_(params.endDate) : null;
  const window = params.urlWindow || 'day';
  if (startDate && endDate && startDate.getTime() > endDate.getTime()) {
    throw new Error('A data inicial não pode ser maior que a final.');
  }
  if (!startDate && !endDate) {
    var today = startOfDay_(new Date());
    startDate = today;
    endDate = today;
  }
  var routeLinkInfo = null;
  if (integrationConfig && integrationConfig.companyId && integrationConfig.domainId && integrationConfig.routeSlug) {
    var routeResponse = fetchLinkRouterRoute_(integrationConfig.companyId, integrationConfig.domainId, integrationConfig.routeSlug);
    if (routeResponse && routeResponse.ok && routeResponse.body && routeResponse.body.status === 'success') {
      var routeData = routeResponse.body.data || {};
      routeLinkInfo = buildLinkRouterUrlIndex_(routeData.links || []);
    } else if (routeResponse) {
      Logger.log('[Link Router] Falha ao carregar rota "' + integrationConfig.routeSlug + '": ' + formatLinkRouterError_(routeResponse));
    }
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
  var distributionRows = filteredRows;
  if (routeLinkInfo && routeLinkInfo.urls && routeLinkInfo.urls.length) {
    var routeKeySet = routeLinkInfo.keySet || {};
    distributionRows = filteredRows.filter(function (row) {
      var normKey = row.normUrl || '';
      if (normKey && routeKeySet[normKey]) {
        return true;
      }
      var rawKey = String(row.url || '').trim().toLowerCase();
      if (rawKey && routeKeySet[rawKey]) {
        return true;
      }
      return false;
    });
  }
  const urlRows = filterRowsForWindow_(filteredRows, hoursInfo, window);
  const siteSummary = buildSiteSummary_(
    filteredRows,
    hoursInfo,
    snapshot.revshareMap,
    urlRows
  );
  const urlSummary = buildUrlSummary_(urlRows, snapshot.revshareMap);
  const detailed = buildDetailedAnalysis_(filteredRows, hoursInfo, snapshot.revshareMap);

  var savedControls = getDistributionControlsConfig_(operationName, trafficType, selectedRouteSlug);
  var distributionParams = {};
  if (savedControls) {
    Object.keys(savedControls).forEach(function (key) {
      var value = savedControls[key];
      if (value != null && value !== '') {
        distributionParams[key] = value;
      }
    });
  }
  if (params.distribution) {
    Object.keys(params.distribution).forEach(function (key) {
      var value = params.distribution[key];
      if (value != null && value !== '') {
        distributionParams[key] = value;
      }
    });
  }

  const distribution = buildDistributionPlan_(
    distributionRows,
    hoursInfo,
    snapshot.revshareMap,
    distributionParams
  );

  if (distribution && distribution.controls) {
    saveDistributionControlsConfig_(operationName, trafficType, selectedRouteSlug, distribution.controls);
  }
  return {
    filters: {
      operation: operationName,
      traffic: trafficType,
      startDate: startDate ? formatDateISO_(startDate) : '',
      endDate: endDate ? formatDateISO_(endDate) : '',
      window: window,
      route: selectedRouteSlug,
      latestHour: hoursInfo.latestHourLabel,
      hours: hoursInfo.windowHours.map(function (h) { return h.label; })
    },
    site: siteSummary,
    url: urlSummary,
    detailed: detailed,
    distribution: distribution,
    integration: integrationConfig ? {
      id: integrationConfig.id || '',
      operation: integrationConfig.operation || '',
      trafficType: integrationConfig.trafficType || '',
      companyId: integrationConfig.companyId || '',
      domainId: integrationConfig.domainId || '',
      routeSlug: integrationConfig.routeSlug || '',
      active: !!integrationConfig.active
    } : null
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
    },
    detailed: {
      hours: [],
      urls: [],
      sites: []
    },
    distribution: buildEmptyDistribution_(),
    integration: null
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
  var hourMs = 60 * 60 * 1000;
  var windowHours = [];
  for (var offset = 3; offset >= 1; offset--) {
    var stamp = latestHour - offset * hourMs;
    windowHours.push({
      key: stamp,
      label: formatHourFromTimestamp_(stamp),
      hasData: !!unique[stamp]
    });
  }
  return {
    allHours: list,
    windowHours: windowHours,
    latestHour: latestHour,
    latestHourLabel: formatHourFromTimestamp_(latestHour)
  };
}

function buildSiteSummary_(rows, hoursInfo, revshareMap, totalRows) {
  var siteMap = {};
  var hourKeys = hoursInfo.windowHours.map(function (h) { return h.key; });
  rows.forEach(function (row) {
    var siteKey = row.site;
    if (!siteMap[siteKey]) {
      siteMap[siteKey] = {
        site: siteKey,
        hours: {}
      };
    }
    var siteData = siteMap[siteKey];
    var revshare = toNumber_(revshareMap[siteKey] || 0) / 100;
    var adjustedRevenue = row.revenue * (1 - revshare);
    if (!siteData.hours[row.timestamp]) {
      siteData.hours[row.timestamp] = { sessions: 0, revenue: 0 };
    }
    if (row.isInterstitial) {
      siteData.hours[row.timestamp].sessions += row.requests;
    }
    siteData.hours[row.timestamp].revenue += adjustedRevenue;
  });

  var totalInput = Array.isArray(totalRows) ? totalRows : rows;
  var siteTotalsMap = {};
  totalInput.forEach(function (row) {
    var siteKey = row.site;
    if (!siteTotalsMap[siteKey]) {
      siteTotalsMap[siteKey] = { sessions: 0, revenue: 0 };
    }
    var totalsData = siteTotalsMap[siteKey];
    var revshare = toNumber_(revshareMap[siteKey] || 0) / 100;
    var adjustedRevenue = row.revenue * (1 - revshare);
    if (row.isInterstitial) {
      totalsData.sessions += row.requests;
    }
    totalsData.revenue += adjustedRevenue;
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
    var siteTotals = siteTotalsMap[siteKey] || { sessions: 0, revenue: 0 };
    var rowOut = {
      site: siteKey,
      sessions: {},
      revenue: {},
      rps: {},
      totalSessions: siteTotals.sessions,
      totalRevenue: siteTotals.revenue,
      totalRps: computeRps_(siteTotals.revenue, siteTotals.sessions)
    };
    hourKeys.forEach(function (key) {
      var hourData = siteData.hours[key] || { sessions: 0, revenue: 0 };
      rowOut.sessions[key] = hourData.sessions;
      rowOut.revenue[key] = hourData.revenue;
      rowOut.rps[key] = computeRps_(hourData.revenue, hourData.sessions);
      totals.sessions[key] += hourData.sessions;
      totals.revenue[key] += hourData.revenue;
    });
    totals.totalSessions += siteTotals.sessions;
    totals.totalRevenue += siteTotals.revenue;
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
        mobTopBlocks: {}
      };
    }
    var data = map[key];
    var revshare = toNumber_(revshareMap[row.site] || 0) / 100;
    var adjustedRevenue = row.revenue * (1 - revshare);
    data.revenue += adjustedRevenue;
    if (row.isInterstitial) {
      data.sessions += row.requests;
    }
    if (row.block && /mob_top/i.test(row.block)) {
      var blockKey = String(row.block || '').trim();
      if (!data.mobTopBlocks[blockKey]) {
        data.mobTopBlocks[blockKey] = {
          requests: 0,
          coverageWeighted: 0,
          adjustedEcpmWeighted: 0
        };
      }
      var blockData = data.mobTopBlocks[blockKey];
      var rowCoverageRatio = row.coverage > 1 ? row.coverage / 100 : row.coverage;
      blockData.requests += row.requests;
      blockData.coverageWeighted += rowCoverageRatio * row.requests;
      blockData.adjustedEcpmWeighted += (row.ecpm * (1 - revshare)) * row.requests;
    }
  });

  var rowsOut = [];
  var totals = {
    sessions: 0,
    revenue: 0,
    mobTopRequests: 0,
    mobTopCoverageWeighted: 0,
    mobTopEcpmWeighted: 0
  };

  Object.keys(map).map(function (key) {
    return map[key];
  }).sort(function (a, b) {
    return b.revenue - a.revenue;
  }).forEach(function (data) {
    var selectedBlock = selectMobTopBlock_(data.mobTopBlocks);
    var avgCoverageRatio = selectedBlock.requests ? selectedBlock.coverageWeighted / selectedBlock.requests : 0;
    var coverage = avgCoverageRatio * 100;
    var ecpm = selectedBlock.requests ? selectedBlock.adjustedEcpmWeighted / selectedBlock.requests : 0;
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
    totals.mobTopRequests += selectedBlock.requests;
    totals.mobTopCoverageWeighted += selectedBlock.coverageWeighted;
    totals.mobTopEcpmWeighted += selectedBlock.adjustedEcpmWeighted;
  });

  var totalsRow = null;
  if (rowsOut.length) {
    var totalsCoverageRatio = totals.mobTopRequests ? totals.mobTopCoverageWeighted / totals.mobTopRequests : 0;
    var totalsCoverage = totalsCoverageRatio * 100;
    var totalsEcpm = totals.mobTopRequests ? totals.mobTopEcpmWeighted / totals.mobTopRequests : 0;
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
    totals: totalsRow,
    windowLabel: ''
  };
}

function buildDetailedAnalysis_(rows, hoursInfo, revshareMap) {
  if (!rows || !rows.length) {
    return {
      hours: [],
      urls: [],
      sites: []
    };
  }
  var allHours = hoursInfo && hoursInfo.allHours ? hoursInfo.allHours.slice() : [];
  var latestHour = hoursInfo && hoursInfo.latestHour != null ? hoursInfo.latestHour : null;
  var hourList = [];
  var hourSet = {};
  allHours.forEach(function (timestamp) {
    if (latestHour != null && timestamp >= latestHour) {
      return;
    }
    hourList.push(timestamp);
    hourSet[timestamp] = true;
  });
  if (!hourList.length && allHours.length) {
    hourList = allHours.slice();
    hourSet = {};
    hourList.forEach(function (timestamp) {
      hourSet[timestamp] = true;
    });
  }
  if (!hourList.length) {
    return {
      hours: [],
      urls: [],
      sites: []
    };
  }

  var siteMap = {};
  var urlMap = {};

  rows.forEach(function (row) {
    if (latestHour != null && row.timestamp >= latestHour) return;
    if (!hourSet[row.timestamp]) return;
    var revshare = toNumber_(revshareMap[row.site] || 0) / 100;
    var adjustedRevenue = row.revenue * (1 - revshare);
    var hourKey = row.timestamp;

    if (!siteMap[row.site]) {
      siteMap[row.site] = {
        site: row.site,
        totals: {
          sessions: 0,
          revenue: 0,
          mobTopRequests: 0,
          mobTopCoverageWeighted: 0,
          mobTopAdjustedEcpmWeighted: 0
        },
        hours: {}
      };
    }
    var siteData = siteMap[row.site];
    if (!siteData.hours[hourKey]) {
      siteData.hours[hourKey] = {
        sessions: 0,
        revenue: 0,
        mobTopRequests: 0,
        mobTopCoverageWeighted: 0,
        mobTopAdjustedEcpmWeighted: 0
      };
    }
    var siteHour = siteData.hours[hourKey];
    siteData.totals.revenue += adjustedRevenue;
    siteHour.revenue += adjustedRevenue;
    if (row.isInterstitial) {
      siteData.totals.sessions += row.requests;
      siteHour.sessions += row.requests;
    }
    if (row.block && /mob_top/i.test(row.block)) {
      var coverageRatio = row.coverage > 1 ? row.coverage / 100 : row.coverage;
      var adjustedEcpm = row.ecpm * (1 - revshare);
      siteData.totals.mobTopRequests += row.requests;
      siteData.totals.mobTopCoverageWeighted += coverageRatio * row.requests;
      siteData.totals.mobTopAdjustedEcpmWeighted += adjustedEcpm * row.requests;
      siteHour.mobTopRequests += row.requests;
      siteHour.mobTopCoverageWeighted += coverageRatio * row.requests;
      siteHour.mobTopAdjustedEcpmWeighted += adjustedEcpm * row.requests;
    }

    var urlKey = row.normUrl || row.url;
    if (!urlMap[urlKey]) {
      urlMap[urlKey] = {
        key: urlKey,
        url: String(row.url || '').trim(),
        site: row.site,
        totals: {
          sessions: 0,
          revenue: 0
        },
        hours: {},
        hourBlocks: {},
        mobTopTotals: {}
      };
    }
    var urlData = urlMap[urlKey];
    if (!urlData.hours[hourKey]) {
      urlData.hours[hourKey] = {
        sessions: 0,
        revenue: 0
      };
    }
    urlData.totals.revenue += adjustedRevenue;
    urlData.hours[hourKey].revenue += adjustedRevenue;
    if (row.isInterstitial) {
      urlData.totals.sessions += row.requests;
      urlData.hours[hourKey].sessions += row.requests;
    }
    if (row.block && /mob_top/i.test(row.block)) {
      var blockKey = String(row.block || '').trim();
      if (!urlData.hourBlocks[hourKey]) {
        urlData.hourBlocks[hourKey] = {};
      }
      if (!urlData.hourBlocks[hourKey][blockKey]) {
        urlData.hourBlocks[hourKey][blockKey] = {
          requests: 0,
          coverageWeighted: 0,
          adjustedEcpmWeighted: 0
        };
      }
      var hourBlock = urlData.hourBlocks[hourKey][blockKey];
      var coverageRatioUrl = row.coverage > 1 ? row.coverage / 100 : row.coverage;
      var adjustedEcpmUrl = row.ecpm * (1 - revshare);
      hourBlock.requests += row.requests;
      hourBlock.coverageWeighted += coverageRatioUrl * row.requests;
      hourBlock.adjustedEcpmWeighted += adjustedEcpmUrl * row.requests;

      if (!urlData.mobTopTotals[blockKey]) {
        urlData.mobTopTotals[blockKey] = {
          requests: 0,
          coverageWeighted: 0,
          adjustedEcpmWeighted: 0
        };
      }
      var totalBlock = urlData.mobTopTotals[blockKey];
      totalBlock.requests += row.requests;
      totalBlock.coverageWeighted += coverageRatioUrl * row.requests;
      totalBlock.adjustedEcpmWeighted += adjustedEcpmUrl * row.requests;
    }
  });

  var hourDescriptors = hourList.map(function (timestamp) {
    return {
      key: timestamp,
      label: formatHourFromTimestamp_(timestamp)
    };
  });

  var siteRows = Object.keys(siteMap).map(function (siteKey) {
    var siteData = siteMap[siteKey];
    var totalsCoverageRatio = siteData.totals.mobTopRequests ? siteData.totals.mobTopCoverageWeighted / siteData.totals.mobTopRequests : 0;
    var totalsEcpm = siteData.totals.mobTopRequests ? siteData.totals.mobTopAdjustedEcpmWeighted / siteData.totals.mobTopRequests : 0;
    var totals = {
      sessions: siteData.totals.sessions,
      revenue: siteData.totals.revenue,
      rps: computeRps_(siteData.totals.revenue, siteData.totals.sessions),
      coverage: totalsCoverageRatio * 100,
      ecpm: totalsEcpm,
      effectiveEcpm: totalsEcpm * totalsCoverageRatio,
      mobTopRequests: siteData.totals.mobTopRequests,
      mobTopCoverageWeighted: siteData.totals.mobTopCoverageWeighted,
      mobTopAdjustedEcpmWeighted: siteData.totals.mobTopAdjustedEcpmWeighted
    };
    var hourly = {};
    hourList.forEach(function (timestamp) {
      var hourData = siteData.hours[timestamp] || {
        sessions: 0,
        revenue: 0,
        mobTopRequests: 0,
        mobTopCoverageWeighted: 0,
        mobTopAdjustedEcpmWeighted: 0
      };
      var coverageRatio = hourData.mobTopRequests ? hourData.mobTopCoverageWeighted / hourData.mobTopRequests : 0;
      var ecpm = hourData.mobTopRequests ? hourData.mobTopAdjustedEcpmWeighted / hourData.mobTopRequests : 0;
      hourly[timestamp] = {
        sessions: hourData.sessions,
        revenue: hourData.revenue,
        rps: computeRps_(hourData.revenue, hourData.sessions),
        coverage: coverageRatio * 100,
        ecpm: ecpm,
        effectiveEcpm: ecpm * coverageRatio,
        mobTopRequests: hourData.mobTopRequests,
        mobTopCoverageWeighted: hourData.mobTopCoverageWeighted,
        mobTopAdjustedEcpmWeighted: hourData.mobTopAdjustedEcpmWeighted
      };
    });
    return {
      key: siteKey,
      site: siteKey,
      totals: totals,
      hourly: hourly
    };
  }).sort(function (a, b) {
    return b.totals.revenue - a.totals.revenue;
  });

  var urlRows = Object.keys(urlMap).map(function (urlKey) {
    var urlData = urlMap[urlKey];
    var selectedBlock = selectMobTopBlock_(urlData.mobTopTotals);
    var totalsCoverageRatio = selectedBlock.requests ? selectedBlock.coverageWeighted / selectedBlock.requests : 0;
    var totalsEcpm = selectedBlock.requests ? selectedBlock.adjustedEcpmWeighted / selectedBlock.requests : 0;
    var totals = {
      sessions: urlData.totals.sessions,
      revenue: urlData.totals.revenue,
      rps: computeRps_(urlData.totals.revenue, urlData.totals.sessions),
      coverage: totalsCoverageRatio * 100,
      ecpm: totalsEcpm,
      effectiveEcpm: totalsEcpm * totalsCoverageRatio,
      mobTopRequests: selectedBlock.requests,
      mobTopCoverageWeighted: selectedBlock.coverageWeighted,
      mobTopAdjustedEcpmWeighted: selectedBlock.adjustedEcpmWeighted
    };
    var hourly = {};
    hourList.forEach(function (timestamp) {
      var base = urlData.hours[timestamp] || { sessions: 0, revenue: 0 };
      var block = selectMobTopBlock_(urlData.hourBlocks[timestamp] || {});
      var coverageRatio = block.requests ? block.coverageWeighted / block.requests : 0;
      var ecpm = block.requests ? block.adjustedEcpmWeighted / block.requests : 0;
      hourly[timestamp] = {
        sessions: base.sessions,
        revenue: base.revenue,
        rps: computeRps_(base.revenue, base.sessions),
        coverage: coverageRatio * 100,
        ecpm: ecpm,
        effectiveEcpm: ecpm * coverageRatio,
        mobTopRequests: block.requests,
        mobTopCoverageWeighted: block.coverageWeighted,
        mobTopAdjustedEcpmWeighted: block.adjustedEcpmWeighted
      };
    });
    return {
      key: urlKey,
      url: urlData.url,
      site: urlData.site,
      totals: totals,
      hourly: hourly
    };
  }).sort(function (a, b) {
    return b.totals.revenue - a.totals.revenue;
  });

  return {
    hours: hourDescriptors,
    urls: urlRows,
    sites: siteRows
  };
}

function buildEmptyDistribution_() {
  var controls = getDefaultDistributionControls_();
  return {
    controls: controls,
    sites: [],
    urlsBySite: {},
    selectedSite: '',
    latestHour: '',
    totalSessions: controls.totalSessions,
    hours: [],
    globalUrls: {
      entries: [],
      all: [],
      averageShare: 0,
      activeCount: 0,
      totalCount: 0,
      targetRecipients: 0,
      limitApplied: false
    }
  };
}

function getDefaultDistributionControls_() {
  return {
    totalSessions: 0,
    urlSeedPercent: 0.02,
    urlMinRecipients: 1,
    urlPriorityShare: 0.6,
    urlTargetRecipients: 0,
    urlRequireAll: 0,
    modeThresholdRps: 250,
    controlReserve: 0.4,
    betMinShare: 0.05
  };
}

function buildDistributionControls_(params) {
  var defaults = getDefaultDistributionControls_();
  var controls = {};
  params = params || {};
  Object.keys(defaults).forEach(function (key) {
    var value = params.hasOwnProperty(key) ? params[key] : defaults[key];
    if (key === 'urlRequireAll') {
      controls[key] = coerceBoolean_(value) ? 1 : 0;
      return;
    }
    var parsed = coerceNumber_(value, defaults[key]);
    if (key === 'totalSessions') {
      controls[key] = Math.max(0, Math.round(parsed));
    } else if (key === 'urlMinRecipients' || key === 'urlTargetRecipients') {
      controls[key] = Math.max(0, Math.round(parsed));
    } else if (key === 'urlSeedPercent' || key === 'urlPriorityShare' || key === 'controlReserve' || key === 'betMinShare') {
      controls[key] = clamp_(parsed, 0, 1);
    } else if (key === 'modeThresholdRps') {
      controls[key] = Math.max(0, parsed);
    } else {
      controls[key] = parsed;
    }
  });
  return controls;
}

function buildDistributionPlan_(rows, hoursInfo, revshareMap, params) {
  var controls = buildDistributionControls_(params);
  var windowHours = hoursInfo && hoursInfo.windowHours ? hoursInfo.windowHours : [];
  var hourKeys = windowHours.map(function (hour) { return hour.key; });
  if (!rows || !rows.length || !hourKeys.length) {
    return {
      controls: controls,
      sites: [],
      urlsBySite: {},
      selectedSite: '',
      latestHour: hoursInfo && hoursInfo.latestHourLabel ? hoursInfo.latestHourLabel : '',
      totalSessions: controls.totalSessions,
      hours: windowHours.map(function (hour) { return hour.label; }),
      globalUrls: {
        entries: [],
        all: [],
        averageShare: 0,
        activeCount: 0,
        totalCount: 0,
        targetRecipients: 0,
        limitApplied: false
      }
    };
  }

  var hourSet = {};
  hourKeys.forEach(function (key) { hourSet[key] = true; });
  var state = getDistributionState_();
  var siteStateMap = state.sites || {};
  var urlStateMap = state.urls || {};

  var siteMap = {};
  rows.forEach(function (row) {
    if (!hourSet[row.timestamp]) return;
    var siteKey = row.site;
    if (!siteMap[siteKey]) {
      siteMap[siteKey] = {
        site: siteKey,
        hours: {},
        urls: {}
      };
    }
    var siteData = siteMap[siteKey];
    if (!siteData.hours[row.timestamp]) {
      siteData.hours[row.timestamp] = { sessions: 0, revenue: 0 };
    }
    var revshare = toNumber_(revshareMap[row.site] || 0) / 100;
    var adjustedRevenue = row.revenue * (1 - revshare);
    siteData.hours[row.timestamp].revenue += adjustedRevenue;
    if (row.isInterstitial) {
      siteData.hours[row.timestamp].sessions += row.requests;
    }

    var urlKey = row.normUrl || row.url;
    if (!siteData.urls[urlKey]) {
      siteData.urls[urlKey] = {
        url: row.url,
        key: urlKey,
        hours: {},
        mobTopBlocks: {}
      };
    }
    var urlData = siteData.urls[urlKey];
    if (!urlData.hours[row.timestamp]) {
      urlData.hours[row.timestamp] = { sessions: 0, revenue: 0 };
    }
    urlData.hours[row.timestamp].revenue += adjustedRevenue;
    if (row.isInterstitial) {
      urlData.hours[row.timestamp].sessions += row.requests;
    }
    if (row.block && /mob_top/i.test(row.block)) {
      var blockKey = String(row.block || '').trim();
      if (!urlData.mobTopBlocks[blockKey]) {
        urlData.mobTopBlocks[blockKey] = {
          requests: 0,
          coverageWeighted: 0,
          adjustedEcpmWeighted: 0
        };
      }
      var blockData = urlData.mobTopBlocks[blockKey];
      var coverageRatio = row.coverage > 1 ? row.coverage / 100 : row.coverage;
      blockData.requests += row.requests;
      blockData.coverageWeighted += coverageRatio * row.requests;
      blockData.adjustedEcpmWeighted += (row.ecpm * (1 - revshare)) * row.requests;
    }
  });

  var siteRows = [];
  var urlsBySite = {};

  Object.keys(siteMap).forEach(function (siteKey) {
    var siteData = siteMap[siteKey];
    var urlStates = urlStateMap[siteKey] || {};
    var sessionsWindow = 0;
    var revenueWindow = 0;
    var hourRps = [];
    hourKeys.forEach(function (key) {
      var hourData = siteData.hours[key] || { sessions: 0, revenue: 0 };
      sessionsWindow += hourData.sessions;
      revenueWindow += hourData.revenue;
      hourRps.push(computeRps_(hourData.revenue, hourData.sessions));
    });
    var rps = computeRps_(revenueWindow, sessionsWindow);

    var urlEntries = [];
    var coverageRequests = 0;
    var coverageWeighted = 0;
    var ecpmWeighted = 0;

    Object.keys(siteData.urls).forEach(function (urlKey) {
      var urlInfo = siteData.urls[urlKey];
      var urlHours = urlInfo.hours;
      var urlSessions = 0;
      var urlRevenue = 0;
      var urlHourRps = [];
      hourKeys.forEach(function (key) {
        var urlHour = urlHours[key] || { sessions: 0, revenue: 0 };
        urlSessions += urlHour.sessions;
        urlRevenue += urlHour.revenue;
        urlHourRps.push(computeRps_(urlHour.revenue, urlHour.sessions));
      });
      var urlRps = computeRps_(urlRevenue, urlSessions);
      var block = selectMobTopBlock_(urlInfo.mobTopBlocks);
      coverageRequests += block.requests;
      coverageWeighted += block.coverageWeighted;
      ecpmWeighted += block.adjustedEcpmWeighted;
      var coverageRatio = block.requests ? block.coverageWeighted / block.requests : 0;
      var coverage = coverageRatio * 100;
      var ecpm = block.requests ? block.adjustedEcpmWeighted / block.requests : 0;
      var urlState = urlStates[urlKey] || {};
      var rounds = Math.max(0, Math.round(coerceNumber_(urlState.rounds, 0)));
      var weight = urlSessions / (urlSessions + DISTRIBUTION_SCORE_SMOOTHING);
      var momentum = computeMomentum_(urlHourRps);
      var coverageFactor = computeCoverageFactor_(coverageRatio);
      var ecpmEff = ecpm * coverageRatio;
      var ecpmScore = ecpmEff * coverageFactor;
      urlEntries.push({
        site: siteKey,
        url: urlInfo.url,
        key: urlKey,
        sessions: urlSessions,
        revenue: urlRevenue,
        rps: urlRps,
        coverage: coverage,
        coverageRatio: coverageRatio,
        ecpm: ecpm,
        ecpmEff: ecpmEff,
        ecpmScore: ecpmScore,
        weight: weight,
        momentum: momentum,
        rounds: rounds,
        state: urlState,
        block: block,
        hourRps: urlHourRps
      });
    });

    var siteCoverageRatio = coverageRequests ? coverageWeighted / coverageRequests : 0;
    var siteCoverage = siteCoverageRatio * 100;
    var siteCoverageFactor = computeCoverageFactor_(siteCoverageRatio);
    var siteEcpm = coverageRequests ? ecpmWeighted / coverageRequests : 0;
    var siteEcpmEff = siteEcpm * siteCoverageRatio;
    var siteEcpmScore = siteEcpmEff * siteCoverageFactor;
    var siteMomentum = computeMomentum_(hourRps);
    var siteConfidence = sessionsWindow / (sessionsWindow + DISTRIBUTION_SCORE_SMOOTHING);
    var siteScore = siteEcpmScore;

    siteRows.push({
      site: siteKey,
      sessions: sessionsWindow,
      revenue: revenueWindow,
      rps: rps,
      ecpm: siteEcpm,
      coverage: siteCoverage,
      coverageRatio: siteCoverageRatio,
      ecpmEff: siteEcpmEff,
      ecpmScore: siteEcpmScore,
      momentum: siteMomentum,
      confidence: siteConfidence,
      score: siteScore,
      urls: urlEntries
    });
  });

  if (!siteRows.length) {
    controls.totalSessions = 0;
    return {
      controls: controls,
      sites: [],
      urlsBySite: {},
      selectedSite: '',
      latestHour: hoursInfo && hoursInfo.latestHourLabel ? hoursInfo.latestHourLabel : '',
      totalSessions: controls.totalSessions,
      hours: windowHours.map(function (hour) { return hour.label; }),
      globalUrls: {
        entries: [],
        all: [],
        averageShare: 0,
        activeCount: 0,
        totalCount: 0,
        targetRecipients: 0,
        limitApplied: false,
        priorityShareTarget: 0,
        priorityCount: 0,
        priorityActive: 0,
        priorityShareActual: 0,
        betMode: false,
        betShareActual: 0,
        betRecipients: 0
      }
    };
  }

  var totalSessionsWindow = 0;
  var totalRevenueWindow = 0;
  var bestSiteRps = 0;
  siteRows.forEach(function (row) {
    totalSessionsWindow += row.sessions || 0;
    totalRevenueWindow += row.revenue || 0;
    if (row.rps > bestSiteRps) bestSiteRps = row.rps;
  });
  controls.totalSessions = totalSessionsWindow;

  var globalMaxScore = 0;
  var globalMaxEcpmEff = 0;
  var globalMaxRps = 0;
  siteRows.forEach(function (siteRow) {
    var urls = siteRow.urls || [];
    urls.forEach(function (entry) {
      var reliability = entry.sessions / (entry.sessions + DISTRIBUTION_SCORE_SMOOTHING);
      var coverageFactor = computeCoverageFactor_(entry.coverageRatio);
      var blended = reliability * (entry.ecpmEff || 0) + (1 - reliability) * (siteRow.ecpmEff || 0);
      var momentumClamp = clamp_(entry.momentum, -DISTRIBUTION_MOMENTUM_CLAMP, DISTRIBUTION_MOMENTUM_CLAMP);
      var momentumFactor = 1 + DISTRIBUTION_MOMENTUM_WEIGHT * momentumClamp;
      var score = blended * coverageFactor * Math.max(0.5, momentumFactor);
      entry.score = score;
      entry.baseScore = blended;
      entry.coverageFactor = coverageFactor;
      entry.momentumFactor = momentumFactor;
      if (score > globalMaxScore) globalMaxScore = score;
      if (entry.ecpmEff > globalMaxEcpmEff) globalMaxEcpmEff = entry.ecpmEff;
      if (entry.rps > globalMaxRps) globalMaxRps = entry.rps;
    });
  });

  siteRows.forEach(function (siteRow) {
    var urls = siteRow.urls || [];
    urls.forEach(function (entry) {
      entry.normalizedScore = globalMaxScore > 0 ? clamp_(entry.score / globalMaxScore, 0, 1) : 0;
      entry.normalizedEcpmEff = globalMaxEcpmEff > 0 ? clamp_(entry.ecpmEff / globalMaxEcpmEff, 0, 1) : 0;
      entry.normalizedRpsEff = globalMaxRps > 0 ? clamp_(entry.rps / globalMaxRps, 0, 1) : 0;
      entry.performanceScore = average_([entry.normalizedScore, entry.normalizedEcpmEff, entry.normalizedRpsEff]);
      entry.sessionShare = totalSessionsWindow > 0 ? entry.sessions / totalSessionsWindow : 0;
    });
    var siteScores = urls.map(function (entry) { return entry.score || 0; }).filter(function (value) { return value > 0; });
    siteRow.score = siteScores.length ? average_(siteScores) : 0;
  });

  var allUrls = [];
  var siteShareMap = {};
  siteRows.forEach(function (siteRow) {
    var siteState = siteStateMap[siteRow.site] || {};
    var siteCurrentShare = normalizeShareValue_(siteState.currentShare);
    var urlStates = urlStateMap[siteRow.site] || {};
    var siteUrls = siteRow.urls || [];
    siteRow.currentShare = siteCurrentShare || 0;
    siteRow.currentSessions = siteRow.currentShare * totalSessionsWindow;
    siteRow.suggestedShare = 0;
    siteRow.suggestedSessions = 0;
    siteRow.deltaShare = 0;
    siteRow.nextEvaluation = siteState.nextEvalHour || '';
    siteRow.lastChange = siteState.lastChangeHour || '';
    siteUrls.forEach(function (entry) {
      var stateEntry = urlStates[entry.key] || {};
      var urlCurrentShare = normalizeShareValue_(stateEntry.currentShare);
      var urlRow = {
        site: siteRow.site,
        url: entry.url,
        key: entry.key,
        sessions: entry.sessions,
        revenue: entry.revenue,
        rps: entry.rps,
        ecpm: entry.ecpm,
        coverage: entry.coverage,
        ecpmEff: entry.ecpmEff,
        momentum: entry.momentum,
        score: entry.score,
        performanceScore: entry.performanceScore,
        performanceNormalizedScore: entry.normalizedScore,
        performanceNormalizedEcpm: entry.normalizedEcpmEff,
        performanceNormalizedRps: entry.normalizedRpsEff,
        sessionShare: entry.sessionShare,
        baseScore: entry.baseScore,
        coverageFactor: entry.coverageFactor,
        momentumFactor: entry.momentumFactor,
        status: stateEntry.status || (entry.sessions > 0 ? 'ok' : 'seed'),
        rounds: stateEntry.rounds || 0,
        lastChange: stateEntry.lastChangeHour || '',
        nextEval: stateEntry.nextEvalHour || '',
        currentShare: urlCurrentShare || 0,
        currentGlobalShare: (urlCurrentShare || 0) * (siteRow.currentShare || 0),
        siteCurrentShare: siteRow.currentShare || 0,
        siteEcpmEff: siteRow.ecpmEff,
        siteCoverage: siteRow.coverage,
        siteMomentum: siteRow.momentum,
        siteScore: siteRow.score,
        bet: false,
        betShare: 0,
        priority: false,
        active: false,
        guaranteed: false,
        limited: false
      };
      allUrls.push(urlRow);
    });
  });

  if (!allUrls.length) {
    return {
      controls: controls,
      sites: [],
      urlsBySite: {},
      selectedSite: '',
      latestHour: hoursInfo && hoursInfo.latestHourLabel ? hoursInfo.latestHourLabel : '',
      totalSessions: controls.totalSessions,
      hours: windowHours.map(function (hour) { return hour.label; }),
      globalUrls: {
        entries: [],
        all: [],
        averageShare: 0,
        activeCount: 0,
        totalCount: 0,
        targetRecipients: 0,
        limitApplied: false,
        priorityShareTarget: 0,
        priorityCount: 0,
        priorityActive: 0,
        priorityShareActual: 0,
        betMode: false,
        betShareActual: 0,
        betRecipients: 0
      }
    };
  }

  var referenceRows = allUrls.filter(function (row) { return (row.sessionShare || 0) >= DISTRIBUTION_MIN_SESSION_SHARE; });
  if (!referenceRows.length) {
    referenceRows = allUrls.slice();
  }
  var referencePerformance = average_(referenceRows.map(function (row) { return row.performanceScore || 0; }));

  function compareByPerformance(a, b) {
    var perfDiff = (b.performanceScore || 0) - (a.performanceScore || 0);
    if (Math.abs(perfDiff) > 1e-6) return perfDiff;
    var rpsDiff = (b.rps || 0) - (a.rps || 0);
    if (Math.abs(rpsDiff) > 1e-6) return rpsDiff;
    var ecpmDiff = (b.ecpmEff || 0) - (a.ecpmEff || 0);
    if (Math.abs(ecpmDiff) > 1e-6) return ecpmDiff;
    return (b.sessions || 0) - (a.sessions || 0);
  }

  var sortedUrls = allUrls.slice().sort(compareByPerformance);

  var guaranteeAll = !!controls.urlRequireAll;
  var controlEligible = guaranteeAll ? sortedUrls.slice() : sortedUrls.filter(function (row) {
    return (row.sessionShare || 0) >= DISTRIBUTION_MIN_SESSION_SHARE;
  });
  var totalEligibleCount = guaranteeAll ? sortedUrls.length : controlEligible.length;

  var priorityConfigured = Math.max(0, Math.round(controls.urlMinRecipients || 0));
  var prioritySummaryCount = Math.min(priorityConfigured, sortedUrls.length);
  var priorityTargetCount = guaranteeAll ? Math.min(priorityConfigured, sortedUrls.length) : Math.min(priorityConfigured, controlEligible.length);

  var betModeActive = bestSiteRps < controls.modeThresholdRps;
  var controlReserve = clamp_(controls.controlReserve != null ? controls.controlReserve : 0, 0, 1);
  var controlPortion = betModeActive ? controlReserve : 1;
  var betShare = betModeActive ? Math.max(0, 1 - controlPortion) : 0;
  var betMinShare = clamp_(controls.betMinShare != null ? controls.betMinShare : 0, 0, 1);
  var betCandidates = [];
  if (betModeActive) {
    betCandidates = allUrls.filter(function (row) {
      return (row.sessionShare || 0) < DISTRIBUTION_MIN_SESSION_SHARE && (row.performanceScore || 0) > referencePerformance;
    }).sort(compareByPerformance);
    if (!betCandidates.length) {
      controlPortion = 1;
      betShare = 0;
      betModeActive = false;
    }
  }

  var betTargets = [];
  if (betShare > 0 && betCandidates.length) {
    var maxRecipients = betMinShare > 0 ? Math.floor(betShare / betMinShare) : betCandidates.length;
    if (maxRecipients === 0 && betShare > 0) {
      maxRecipients = 1;
    }
    if (maxRecipients > betCandidates.length) {
      maxRecipients = betCandidates.length;
    }
    betTargets = betCandidates.slice(0, maxRecipients);
  } else {
    betShare = 0;
  }

  var betTargetKeys = {};
  betTargets.forEach(function (row) { betTargetKeys[row.key] = true; });

  var minShare = clamp_(controls.urlSeedPercent || 0, 0, 1);
  var targetPriorityTotal = priorityTargetCount > 0 ? clamp_(controls.urlPriorityShare || 0, 0, 1) : 0;

  if (priorityTargetCount > 0 && targetPriorityTotal > controlPortion && betShare > 0) {
    var borrowForPriority = Math.min(targetPriorityTotal - controlPortion, betShare);
    if (borrowForPriority > 0) {
      controlPortion += borrowForPriority;
      betShare -= borrowForPriority;
    }
  }
  var priorityKeys = {};
  for (var p = 0; p < priorityTargetCount; p++) {
    priorityKeys[controlEligible[p].key] = true;
  }

  var limit = Math.max(0, Math.round(controls.urlTargetRecipients || 0));
  if (priorityTargetCount > 0 && limit > 0 && limit < priorityTargetCount) {
    limit = priorityTargetCount;
  }
  var activeMap = {};
  var activeSet = [];
  if (guaranteeAll) {
    sortedUrls.forEach(function (row) {
      if (!activeMap[row.key]) {
        activeMap[row.key] = true;
        activeSet.push(row);
      }
    });
  } else if (limit === 0) {
    controlEligible.forEach(function (row) {
      if (!activeMap[row.key]) {
        activeMap[row.key] = true;
        activeSet.push(row);
      }
    });
  } else {
    controlEligible.forEach(function (row) {
      if (priorityKeys[row.key] && !activeMap[row.key]) {
        activeMap[row.key] = true;
        activeSet.push(row);
      }
    });
    controlEligible.forEach(function (row) {
      if (activeSet.length >= limit) return;
      if (!activeMap[row.key]) {
        activeMap[row.key] = true;
        activeSet.push(row);
      }
    });
  }
  var limitApplied = !guaranteeAll && limit > 0 && activeSet.length < totalEligibleCount;
  var activeKeys = {};
  activeSet.forEach(function (row) { activeKeys[row.key] = true; });
  if (!activeSet.length) {
    activeSet = controlEligible.length ? controlEligible.slice() : sortedUrls.slice();
    activeSet.forEach(function (row) { activeKeys[row.key] = true; });
  }

  var baseCount = activeSet.length;
  var maxAssignable = controlPortion + betShare;
  if (baseCount > 0 && minShare * baseCount > maxAssignable) {
    minShare = maxAssignable / baseCount;
  }
  var nonPriorityRows = activeSet.filter(function (row) { return !priorityKeys[row.key]; });
  var nonPriorityCount = nonPriorityRows.length;
  var minPriorityTotal = priorityTargetCount * minShare;
  var priorityAllocation = 0;
  var remainingControl = 0;
  function refreshPriorityAllocation() {
    priorityAllocation = priorityTargetCount > 0 ? Math.min(controlPortion, Math.max(targetPriorityTotal, minPriorityTotal)) : 0;
    remainingControl = Math.max(0, controlPortion - priorityAllocation);
    if (nonPriorityCount === 0) {
      priorityAllocation = controlPortion;
      remainingControl = 0;
    }
  }
  refreshPriorityAllocation();
  if (priorityTargetCount > 0 && priorityAllocation < minPriorityTotal && betShare > 0) {
    var borrowForMinimum = Math.min(minPriorityTotal - priorityAllocation, betShare);
    if (borrowForMinimum > 0) {
      controlPortion += borrowForMinimum;
      betShare -= borrowForMinimum;
      refreshPriorityAllocation();
    }
  }
  var requiredNonPriority = nonPriorityCount * minShare;
  function enforceNonPriorityFloor() {
    if (requiredNonPriority > remainingControl) {
      if (priorityTargetCount > 0) {
        var maxPriorityAllocation = Math.max(0, controlPortion - requiredNonPriority);
        if (priorityAllocation > maxPriorityAllocation) {
          priorityAllocation = Math.max(minPriorityTotal, maxPriorityAllocation);
          remainingControl = Math.max(0, controlPortion - priorityAllocation);
        }
      } else {
        remainingControl = controlPortion;
      }
    }
  }
  while (requiredNonPriority > remainingControl && betShare > 0) {
    var borrow = Math.min(requiredNonPriority - remainingControl, betShare);
    betShare -= borrow;
    controlPortion += borrow;
    refreshPriorityAllocation();
    requiredNonPriority = nonPriorityCount * minShare;
  }
  enforceNonPriorityFloor();

  if (betShare > 0 && betMinShare > 0 && betShare < betMinShare) {
    controlPortion = Math.min(1, controlPortion + betShare);
    betShare = 0;
    betTargets = [];
    betModeActive = false;
    refreshPriorityAllocation();
    requiredNonPriority = nonPriorityCount * minShare;
  }
  enforceNonPriorityFloor();

  if (betShare <= 0) {
    betShare = 0;
    betTargets = [];
    betModeActive = false;
  } else if (betTargets.length) {
    var betLimit = betMinShare > 0 ? Math.floor(betShare / betMinShare) : betTargets.length;
    if (betLimit === 0 && betShare > 0) {
      betLimit = 1;
    }
    if (betLimit < betTargets.length) {
      betTargets = betTargets.slice(0, betLimit);
    }
  }

  var allocationMap = {};
  var betAllocationMap = {};

  function allocatePortion(rows, totalShare, minEach, map, tracker) {
    totalShare = Math.max(0, totalShare);
    if (!rows || !rows.length || totalShare <= 0) return;
    var perItem = Math.max(0, minEach || 0);
    var baseline = perItem * rows.length;
    if (baseline > totalShare) {
      perItem = totalShare / rows.length;
      baseline = perItem * rows.length;
    }
    var leftover = totalShare - baseline;
    var weights = rows.map(function (row) { return Math.max(row.performanceScore || 0, 0); });
    var weightSum = weights.reduce(function (acc, val) { return acc + val; }, 0);
    rows.forEach(function (row, index) {
      var portion = perItem;
      if (leftover > 0) {
        if (weightSum > 0) {
          portion += leftover * (weights[index] / weightSum);
        } else {
          portion += leftover / rows.length;
        }
      }
      map[row.key] = (map[row.key] || 0) + portion;
      if (tracker) {
        tracker[row.key] = (tracker[row.key] || 0) + portion;
      }
    });
  }

  var priorityRows = activeSet.filter(function (row) { return priorityKeys[row.key]; });
  if (priorityAllocation > 0 && priorityRows.length) {
    allocatePortion(priorityRows, priorityAllocation, minShare, allocationMap);
  } else if (priorityAllocation > 0) {
    remainingControl += priorityAllocation;
  }
  if (remainingControl > 0 && nonPriorityRows.length) {
    allocatePortion(nonPriorityRows, remainingControl, minShare, allocationMap);
  }

  if (betShare > 0 && betTargets.length) {
    allocatePortion(betTargets, betShare, betMinShare, allocationMap, betAllocationMap);
  }

  var totalAllocated = 0;
  allUrls.forEach(function (row) {
    totalAllocated += allocationMap[row.key] || 0;
  });
  if (totalAllocated <= 0 && allUrls.length) {
    var fallbackCandidates = [];
    if (betShare > 0 && betTargets.length) {
      fallbackCandidates = betTargets.slice();
    } else if (controlEligible.length) {
      fallbackCandidates = controlEligible.slice();
    } else {
      fallbackCandidates = sortedUrls.slice();
    }
    if (fallbackCandidates.length) {
      var equal = 1 / fallbackCandidates.length;
      fallbackCandidates.forEach(function (row) {
        allocationMap[row.key] = equal;
        if (betTargetKeys[row.key]) {
          betAllocationMap[row.key] = (betAllocationMap[row.key] || 0) + equal;
        }
      });
      totalAllocated = 1;
    }
  }
  if (totalAllocated > 0 && Math.abs(totalAllocated - 1) > 1e-6) {
    var scale = 1 / totalAllocated;
    Object.keys(allocationMap).forEach(function (key) {
      allocationMap[key] = (allocationMap[key] || 0) * scale;
    });
    Object.keys(betAllocationMap).forEach(function (key) {
      betAllocationMap[key] = (betAllocationMap[key] || 0) * scale;
    });
  }

  var urlsBySite = {};
  var globalAggregates = [];
  allUrls.forEach(function (row) {
    var totalShare = allocationMap[row.key] || 0;
    var betPortion = betAllocationMap[row.key] || 0;
    var sessionEligible = guaranteeAll || (row.sessionShare || 0) >= DISTRIBUTION_MIN_SESSION_SHARE;
    row.betShare = betPortion;
    row.bet = betPortion > 0;
    row.priority = !!priorityKeys[row.key];
    row.active = !!activeKeys[row.key] || row.bet;
    row.guaranteed = row.priority || guaranteeAll;
    row.sessionEligible = sessionEligible;
    row.ineligibleBySessions = !sessionEligible;
    row.limited = !row.active && !row.bet && sessionEligible && !guaranteeAll && limit > 0;
    row.suggestedGlobalShare = totalShare;
    globalAggregates.push(row);
    siteShareMap[row.site] = (siteShareMap[row.site] || 0) + totalShare;
  });

  var bestShare = -1;
  var defaultSite = '';
  var siteOutput = [];
  siteRows.forEach(function (siteRow) {
    var siteShare = siteShareMap[siteRow.site] || 0;
    siteRow.suggestedShare = siteShare;
    siteRow.suggestedSessions = siteShare * totalSessionsWindow;
    siteRow.deltaShare = siteShare - siteRow.currentShare;
    siteRow.mode = betModeActive ? 'Apostas' : 'Normal';
    if (siteShare > bestShare) {
      bestShare = siteShare;
      defaultSite = siteRow.site;
    }
    var siteCurrentSessions = siteRow.currentShare * totalSessionsWindow;
    siteRow.currentSessions = siteCurrentSessions;
    siteOutput.push({
      site: siteRow.site,
      rps: siteRow.rps,
      ecpm: siteRow.ecpm,
      coverage: siteRow.coverage,
      ecpmEff: siteRow.ecpmEff,
      momentum: siteRow.momentum,
      confidence: siteRow.sessions > 0 ? siteRow.sessions / (siteRow.sessions + DISTRIBUTION_SCORE_SMOOTHING) : 0,
      score: siteRow.score,
      currentShare: siteRow.currentShare,
      suggestedShare: siteShare,
      deltaShare: siteRow.deltaShare,
      mode: siteRow.mode,
      nextEvaluation: siteRow.nextEvaluation,
      lastChange: siteRow.lastChange,
      currentSessions: siteCurrentSessions,
      suggestedSessions: siteRow.suggestedSessions
    });
  });

  globalAggregates.forEach(function (row) {
    var siteShare = siteShareMap[row.site] || 0;
    var siteCurrentShare = (siteStateMap[row.site] && normalizeShareValue_(siteStateMap[row.site].currentShare)) || 0;
    var siteCurrentSessions = siteCurrentShare * totalSessionsWindow;
    var siteSuggestedSessions = siteShare * totalSessionsWindow;
    row.suggestedShare = siteShare > 0 ? row.suggestedGlobalShare / siteShare : 0;
    row.deltaShare = row.suggestedShare - (row.currentShare || 0);
    row.deltaGlobalShare = row.suggestedGlobalShare - (row.currentGlobalShare || 0);
    row.currentSessions = (row.currentShare || 0) * siteCurrentSessions;
    row.suggestedSessions = row.suggestedShare * siteSuggestedSessions;
    if (!urlsBySite[row.site]) {
      urlsBySite[row.site] = [];
    }
    urlsBySite[row.site].push(row);
  });

  var sortedGlobal = globalAggregates.slice().sort(function (a, b) {
    return (b.suggestedGlobalShare || 0) - (a.suggestedGlobalShare || 0);
  });

  var priorityShareAccum = 0;
  var priorityActiveCount = 0;
  var betShareAccum = 0;
  var betCount = 0;
  sortedGlobal.forEach(function (row, index) {
    if (row.priority && (row.suggestedGlobalShare || 0) > 0) {
      priorityActiveCount++;
      priorityShareAccum += row.suggestedGlobalShare || 0;
    }
    if (row.betShare > 0) {
      betCount++;
      betShareAccum += row.betShare;
    }
    row.reason = describeUrlAllocationReason_(row, siteStateMap[row.site] || {}, controls, index, sortedGlobal.length, limitApplied, betModeActive);
  });

  var activeCount = sortedGlobal.filter(function (row) { return (row.suggestedGlobalShare || 0) > 0; }).length;
  var totalCount = sortedGlobal.length;
  var averageGlobalShare = activeCount ? 1 / activeCount : (totalCount ? 1 / totalCount : 0);
  var displayEntries = sortedGlobal.filter(function (row) { return (row.suggestedGlobalShare || 0) > 0; });
  if (!displayEntries.length) {
    displayEntries = sortedGlobal.slice();
  }

  var globalSummary = {
    entries: displayEntries,
    all: sortedGlobal,
    averageShare: averageGlobalShare,
    activeCount: activeCount,
    totalCount: totalCount,
    targetRecipients: guaranteeAll ? 0 : limit,
    limitApplied: limitApplied,
    priorityShareTarget: targetPriorityTotal,
    priorityCount: prioritySummaryCount,
    priorityActive: priorityActiveCount,
    priorityShareActual: priorityShareAccum,
    betMode: betModeActive,
    betShareActual: betShareAccum,
    betRecipients: betCount
  };

  return {
    controls: controls,
    sites: siteOutput,
    urlsBySite: urlsBySite,
    selectedSite: defaultSite,
    latestHour: hoursInfo && hoursInfo.latestHourLabel ? hoursInfo.latestHourLabel : '',
    totalSessions: totalSessionsWindow,
    hours: windowHours.map(function (hour) { return hour.label; }),
    globalUrls: globalSummary
  };
}

function selectMobTopBlock_(blocks) {
  var best = null;
  Object.keys(blocks || {}).forEach(function (name) {
    var block = blocks[name];
    if (!best || block.requests > best.requests) {
      best = block;
    }
  });
  if (!best) {
    return {
      requests: 0,
      coverageWeighted: 0,
      adjustedEcpmWeighted: 0
    };
  }
  return best;
}

function describeUrlAllocationReason_(row, siteCtx, controls, index, totalCount, limitApplied, betModeActive) {
  var share = toNumber_(row && row.suggestedGlobalShare != null ? row.suggestedGlobalShare : 0);
  var delta = toNumber_(row && row.deltaGlobalShare != null ? row.deltaGlobalShare : 0);
  var parts = [];

  if (share <= 0) {
    if (row && row.bet) {
      parts.push('URL ficou elegível para apostas, mas não recebeu tráfego nesta rodada.');
    } else if (row && row.limited) {
      parts.push('Sem tráfego nesta rodada por causa do limite de URLs ativas.');
    } else if (limitApplied) {
      parts.push('Sem tráfego nesta rodada por causa do limite global de URLs ativas.');
    } else {
      parts.push('Sem tráfego sugerido nesta rodada.');
    }
    if (row && row.ineligibleBySessions && !(controls && controls.urlRequireAll)) {
      parts.push('Aguardando atingir pelo menos ' + formatShareValue_(DISTRIBUTION_MIN_SESSION_SHARE, 1) + ' das sessões no período para receber tráfego fora do modo apostas.');
    }
    if (row && row.guaranteed) {
      parts.push('Mantida como garantida (' + (row.status || 'seed') + ').');
    } else if (row && row.status) {
      parts.push('Status atual: ' + row.status + '.');
    }
    return parts.join(' ');
  }

  var shareText = formatShareValue_(share);
  var shareChange = '';
  if (Math.abs(delta) >= 0.0005) {
    shareChange = (delta >= 0 ? '+' : '-') + formatShareDeltaValue_(delta);
  }
  var siteShare = toNumber_(row && row.suggestedShare != null ? row.suggestedShare : 0);
  var siteShareText = formatShareValue_(siteShare);

  var intro = 'Share sugerido ' + shareText;
  if (shareChange) {
    intro += ' (' + shareChange + ')';
  }
  parts.push(intro + '.');
  parts.push('Dentro do site recebe ' + siteShareText + ' do tráfego sugerido.');
  if (row && row.priority) {
    var targetPriorityShare = controls && controls.urlPriorityShare != null ? controls.urlPriorityShare : 0;
    if (targetPriorityShare > 0) {
      parts.push('No grupo prioritário (' + formatShareValue_(targetPriorityShare, 1) + ' do tráfego total).');
    } else {
      parts.push('No grupo prioritário definido manualmente.');
    }
  } else if (controls && controls.urlMinRecipients > 0 && controls.urlPriorityShare > 0 && !controls.urlRequireAll) {
    var residualShare = Math.max(0, 1 - controls.urlPriorityShare);
    if (residualShare > 0) {
      parts.push('Distribuição dentro do grupo complementar (' + formatShareValue_(residualShare, 1) + ' restante).');
    }
  }

  if (row && row.bet) {
    parts.push('Recebe ' + formatShareValue_(row.betShare || 0) + ' do bloco de apostas para acelerar ganho de sinal.');
  } else if (betModeActive) {
    parts.push('Mantém posição na reserva de controle durante o modo de apostas.');
  }

  if (row && row.performanceScore != null) {
    parts.push('Desempenho composto ' + formatShareValue_(row.performanceScore, 1) + ' (score ' + formatShareValue_(row.performanceNormalizedScore || 0, 1) + ', eCPM ef. ' + formatShareValue_(row.performanceNormalizedEcpm || 0, 1) + ', RPS ' + formatShareValue_(row.performanceNormalizedRps || 0, 1) + ').');
  }
  parts.push('Métricas base: score ' + formatScoreValue_(row && row.score != null ? row.score : 0, 2) + ', eCPM efetivo ' + formatCurrencyValue_(row && row.ecpmEff != null ? row.ecpmEff : 0, 2) + ', RPS ' + formatCurrencyValue_(row && row.rps != null ? row.rps : 0, 2) + '.');

  var rankText = 'Prioridade global #' + (index + 1) + ' de ' + totalCount + '.';
  parts.push(rankText);

  if (row && row.status) {
    parts.push('Status atual: ' + row.status + '.');
  }

  return parts.join(' ');
}


function formatShareValue_(share, decimals) {
  var dec = decimals != null ? decimals : 1;
  var factor = Math.pow(10, dec);
  var percent = toNumber_(share) * 100;
  var rounded = Math.round(percent * factor) / factor;
  return rounded.toFixed(dec) + '%';
}

function formatShareDeltaValue_(delta, decimals) {
  var dec = decimals != null ? decimals : 1;
  var factor = Math.pow(10, dec);
  var percentPoints = Math.abs(toNumber_(delta) * 100);
  var rounded = Math.round(percentPoints * factor) / factor;
  return rounded.toFixed(dec) + 'pp';
}

function formatCurrencyValue_(value, decimals) {
  var dec = decimals != null ? decimals : 2;
  var factor = Math.pow(10, dec);
  var amount = Math.round(toNumber_(value) * factor) / factor;
  return '$' + amount.toFixed(dec);
}

function formatScoreValue_(value, decimals) {
  var dec = decimals != null ? decimals : 2;
  var factor = Math.pow(10, dec);
  var amount = Math.round(toNumber_(value) * factor) / factor;
  return amount.toFixed(dec);
}

function formatPercentDisplay_(value, decimals) {
  var dec = decimals != null ? decimals : 1;
  var factor = Math.pow(10, dec);
  var percent = Math.round(toNumber_(value) * factor) / factor;
  return percent.toFixed(dec) + '%';
}

function getConfigSnapshot_() {
  const operationSheet = ensureOperationsSheet_();
  const operationIntegrationSheet = ensureOperationIntegrationSheet_();
  const trafficSheet = ensureTrafficSheet_();
  const revSheet = ensureRevShareSheet_();

  refreshRevShareSites_();

  const operationValues = getSheetValues_(operationSheet);
  const operationIntegrationValues = getSheetValues_(operationIntegrationSheet);
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

  const operationIntegrationRows = operationIntegrationValues.rows.map(function (row) {
    var rawActive = row.ativo;
    return {
      id: row.id,
      operation: row.operacao,
      trafficType: row.tipotrafego,
      companyId: row.empresaid,
      domainId: row.dominioid,
      routeSlug: row.slugrota,
      active: coerceBoolean_(rawActive)
    };
  });
  const operationIntegrationsByOperation = {};
  const operationIntegrationsByOperationTraffic = {};
  operationIntegrationRows.forEach(function (row) {
    if (!row.operation) return;
    if (!operationIntegrationsByOperation[row.operation]) {
      operationIntegrationsByOperation[row.operation] = [];
    }
    operationIntegrationsByOperation[row.operation].push(row);
    var trafficKey = row.trafficType || '';
    var opTrafficKey = row.operation + '\u0000' + trafficKey;
    if (!operationIntegrationsByOperationTraffic[opTrafficKey]) {
      operationIntegrationsByOperationTraffic[opTrafficKey] = [];
    }
    operationIntegrationsByOperationTraffic[opTrafficKey].push(row);
  });
  Object.keys(operationIntegrationsByOperation).forEach(function (op) {
    operationIntegrationsByOperation[op].sort(function (a, b) {
      var trafficCompare = String(a.trafficType || '').localeCompare(String(b.trafficType || ''), 'pt-BR');
      if (trafficCompare !== 0) return trafficCompare;
      return String(a.routeSlug || '').localeCompare(String(b.routeSlug || ''), 'pt-BR');
    });
  });
  Object.keys(operationIntegrationsByOperationTraffic).forEach(function (key) {
    operationIntegrationsByOperationTraffic[key].sort(function (a, b) {
      return String(a.routeSlug || '').localeCompare(String(b.routeSlug || ''), 'pt-BR');
    });
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
    operationIntegrationRows: operationIntegrationRows,
    operationIntegrationsByOperation: operationIntegrationsByOperation,
    operationIntegrationsByOperationTraffic: operationIntegrationsByOperationTraffic,
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
      tipotrafego: pickValue_(rowValues, idx, 'tipotrafego'),
      tipo: pickValue_(rowValues, idx, 'tipo'),
      utm_source: pickValue_(rowValues, idx, 'utmsource'),
      site: pickValue_(rowValues, idx, 'site'),
      revshare: pickValue_(rowValues, idx, 'revshare'),
      empresaid: pickValue_(rowValues, idx, 'empresaid'),
      dominioid: pickValue_(rowValues, idx, 'dominioid'),
      slugrota: pickValue_(rowValues, idx, 'slugrota'),
      ativo: pickValue_(rowValues, idx, 'ativo')
    });
  }
  return { headers: headers, rows: rows };
}

function getDistributionState_() {
  var sheet = ensureDistributionSheet_();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var state = { sites: {}, urls: {} };
  if (lastRow < 2 || lastCol < 1) {
    return state;
  }
  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  if (!values.length) {
    return state;
  }
  var headers = values[0].map(function (h) { return String(h).trim(); });
  var idx = getHeaderIndexMap_(headers);
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (!row || String(row.join('')).trim() === '') continue;
    var type = String(pickValue_(row, idx, 'tipo') || '').trim().toLowerCase();
    var site = String(pickValue_(row, idx, 'site') || '').trim();
    if (!site) continue;
    var shareAtual = pickValue_(row, idx, 'shareatual');
    var shareAnterior = pickValue_(row, idx, 'shareanterior');
    var lastChange = pickValue_(row, idx, 'ultimaatualizacao');
    var nextEval = pickValue_(row, idx, 'proximaavaliacao');
    var rounds = coerceNumber_(pickValue_(row, idx, 'rodadas'), 0);
    var status = pickValue_(row, idx, 'status');
    var mode = pickValue_(row, idx, 'modo');
    if (type === 'site') {
      state.sites[site] = {
        currentShare: shareAtual,
        lastShare: shareAnterior,
        lastChangeHour: lastChange,
        nextEvalHour: nextEval,
        rounds: Math.max(0, Math.round(rounds)),
        status: status,
        mode: mode
      };
    } else if (type === 'url') {
      var urlValue = pickValue_(row, idx, 'url');
      var key = normUrlStrict_(urlValue) || String(urlValue || '').trim();
      if (!state.urls[site]) state.urls[site] = {};
      state.urls[site][key] = {
        currentShare: shareAtual,
        lastShare: shareAnterior,
        lastChangeHour: lastChange,
        nextEvalHour: nextEval,
        rounds: Math.max(0, Math.round(rounds)),
        status: status,
        mode: mode,
        rawUrl: urlValue
      };
    }
  }
  return state;
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

function ensureOperationIntegrationSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(OPERATION_INTEGRATION_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(OPERATION_INTEGRATION_SHEET);
  }
  var headers = ['ID', 'Operação', 'Tipo Tráfego', 'Empresa ID', 'Domínio ID', 'Slug Rota', 'Ativo?'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    var dataRange = sheet.getRange(2, 1, lastRow - 1, headers.length);
    var values = dataRange.getValues();
    var needsMigration = values.some(function (row) {
      var status = String(row[5] || '').toLowerCase();
      var hasStatus = status === 'ativo' || status === 'inativo';
      return hasStatus && (row[6] === '' || row[6] == null);
    });
    if (needsMigration) {
      var migrated = values.map(function (row) {
        var status = row[5];
        return [
          row[0],
          row[1],
          '',
          row[2],
          row[3],
          row[4],
          status
        ];
      });
      dataRange.setValues(migrated);
    }
  }
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

function ensureLinkRouterHistorySheet_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(LINK_ROUTER_HISTORY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LINK_ROUTER_HISTORY_SHEET);
  }
  var headers = ['Timestamp', 'Operação', 'Tráfego', 'Empresa ID', 'Domínio ID', 'Slug', 'URL', 'Percentual Enviado (%)', 'Origem do Percentual', 'Status', 'Mensagem'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  return sheet;
}

function ensureDistributionSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(DISTRIBUTION_STATE_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(DISTRIBUTION_STATE_SHEET);
  }
  var headers = ['Tipo', 'Site', 'URL', 'Share Atual', 'Share Anterior', 'Última Atualização', 'Próxima Avaliação', 'Rodadas', 'Status', 'Modo'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  return sheet;
}

function ensureDistributionConfigSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(DISTRIBUTION_CONFIG_SHEET);
  var defaults = getDefaultDistributionControls_();
  var requiredHeaders = ['Operação', 'Tráfego', 'Slug Rota'];
  Object.keys(defaults).forEach(function (key) {
    requiredHeaders.push(key);
  });

  if (!sheet) {
    sheet = ss.insertSheet(DISTRIBUTION_CONFIG_SHEET);
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  var maxColumns = sheet.getMaxColumns();
  if (maxColumns < requiredHeaders.length) {
    sheet.insertColumnsAfter(maxColumns, requiredHeaders.length - maxColumns);
  }

  var lastColumn = Math.max(sheet.getLastColumn(), requiredHeaders.length);
  var headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  var hasMeaningfulHeaders = headers.some(function (value) {
    return !!normalizeKey_(value);
  });

  if (!hasMeaningfulHeaders) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    sheet.setFrozenRows(1);
    return sheet;
  }

  var normalizedExisting = {};
  headers.forEach(function (header, index) {
    var key = normalizeKey_(header);
    if (key && normalizedExisting[key] == null) {
      normalizedExisting[key] = index;
    }
  });

  requiredHeaders.forEach(function (header) {
    var normalized = normalizeKey_(header);
    if (normalizedExisting.hasOwnProperty(normalized)) {
      return;
    }
    var targetIndex = -1;
    for (var i = 0; i < headers.length; i++) {
      if (!normalizeKey_(headers[i])) {
        targetIndex = i;
        break;
      }
    }
    if (targetIndex === -1) {
      targetIndex = headers.length;
      headers.push('');
      if (sheet.getMaxColumns() < headers.length) {
        sheet.insertColumnsAfter(sheet.getMaxColumns(), headers.length - sheet.getMaxColumns());
      }
    }
    headers[targetIndex] = header;
    sheet.getRange(1, targetIndex + 1).setValue(header);
    normalizedExisting[normalized] = targetIndex;
  });

  sheet.setFrozenRows(1);
  return sheet;
}

function getDistributionControlsConfig_(operation, traffic, route) {
  if (!operation || !traffic) {
    return {};
  }
  var sheet = ensureDistributionConfigSheet_();
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return {};
  }
  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = values[0];
  var idx = getHeaderIndexMap_(headers);
  var opIndex = idx.operacao;
  var trafficIndex = idx.trafego;
  var slugIndex = idx.slugrota;
  if (opIndex == null || trafficIndex == null) {
    return {};
  }
  var targetRow = null;
  var fallbackRow = null;
  var normalizedRoute = String(route || '').trim().toLowerCase();
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (!row || row.join('').trim() === '') continue;
    var rowOp = String(row[opIndex] || '').trim();
    var rowTraffic = String(row[trafficIndex] || '').trim();
    if (rowOp && rowTraffic && rowOp.toLowerCase() === String(operation).trim().toLowerCase() && rowTraffic.toLowerCase() === String(traffic).trim().toLowerCase()) {
      var rowSlug = slugIndex == null ? '' : String(row[slugIndex] || '').trim();
      var normalizedRowSlug = rowSlug.toLowerCase();
      var matchesSlug;
      if (slugIndex == null) {
        matchesSlug = !normalizedRoute;
      } else if (normalizedRoute) {
        matchesSlug = normalizedRowSlug === normalizedRoute;
      } else {
        matchesSlug = !normalizedRowSlug;
      }
      if (matchesSlug) {
        targetRow = row;
        break;
      }
      if (!normalizedRoute && !fallbackRow) {
        fallbackRow = row;
      }
    }
  }
  if (!targetRow && fallbackRow) {
    targetRow = fallbackRow;
  }
  if (!targetRow) {
    return {};
  }
  var defaults = getDefaultDistributionControls_();
  var controls = {};
  Object.keys(defaults).forEach(function (key) {
    var normalized = normalizeKey_(key);
    var columnIndex = idx[normalized];
    if (columnIndex == null) return;
    controls[key] = toNumber_(targetRow[columnIndex]);
  });
  return controls;
}

function saveDistributionControlsConfig_(operation, traffic, route, controls) {
  if (!operation || !traffic || !controls) {
    return;
  }
  var sheet = ensureDistributionConfigSheet_();
  var defaults = getDefaultDistributionControls_();
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var idx = getHeaderIndexMap_(headers);
  var opIndex = idx.operacao;
  var trafficIndex = idx.trafego;
  var slugIndex = idx.slugrota;
  if (opIndex == null || trafficIndex == null) {
    return;
  }
  var lastRow = sheet.getLastRow();
  var rangeValues = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  var rowNumber = -1;
  var normalizedRoute = String(route || '').trim().toLowerCase();
  for (var i = 0; i < rangeValues.length; i++) {
    var row = rangeValues[i];
    var rowOp = String(row[opIndex] || '').trim();
    var rowTraffic = String(row[trafficIndex] || '').trim();
    if (rowOp && rowTraffic && rowOp.toLowerCase() === String(operation).trim().toLowerCase() && rowTraffic.toLowerCase() === String(traffic).trim().toLowerCase()) {
      var rowSlug = slugIndex == null ? '' : String(row[slugIndex] || '').trim().toLowerCase();
      var matchesSlug;
      if (slugIndex == null) {
        matchesSlug = !normalizedRoute;
      } else if (normalizedRoute) {
        matchesSlug = rowSlug === normalizedRoute;
      } else {
        matchesSlug = !rowSlug;
      }
      if (matchesSlug) {
        rowNumber = i + 2;
        break;
      }
    }
  }
  var rowValues = new Array(headers.length);
  for (var h = 0; h < headers.length; h++) {
    rowValues[h] = '';
  }
  rowValues[opIndex] = String(operation).trim();
  rowValues[trafficIndex] = String(traffic).trim();
  if (slugIndex != null) {
    rowValues[slugIndex] = String(route || '').trim();
  }
  Object.keys(defaults).forEach(function (key) {
    var normalized = normalizeKey_(key);
    var columnIndex = idx[normalized];
    if (columnIndex == null) return;
    var value = controls.hasOwnProperty(key) ? controls[key] : defaults[key];
    rowValues[columnIndex] = value != null && value !== '' ? Number(value) : defaults[key];
  });
  if (rowNumber > -1) {
    sheet.getRange(rowNumber, 1, 1, rowValues.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
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

function coerceNumber_(value, fallback) {
  if (value === null || value === undefined || value === '') {
    return fallback != null ? fallback : 0;
  }
  if (typeof value === 'number') {
    return isNaN(value) ? (fallback != null ? fallback : 0) : value;
  }
  var num = toNumber_(value);
  if (isNaN(num)) {
    return fallback != null ? fallback : 0;
  }
  return num;
}

function coerceBoolean_(value) {
  if (value === true) return true;
  if (value === false) return false;
  if (typeof value === 'number') {
    if (isNaN(value)) return false;
    return value !== 0;
  }
  var str = String(value || '').trim().toLowerCase();
  if (!str) return false;
  if (str === '0' || str === 'false' || str === 'nao' || str === 'não' || str === 'inativo' || str === 'off') {
    return false;
  }
  return ['1', 'true', 'sim', 'yes', 'ativo', 'on', 'y'].indexOf(str) !== -1;
}

function normalizeShareValue_(value) {
  var num = coerceNumber_(value, 0);
  if (num > 1) {
    return num / 100;
  }
  return num;
}

function computeCoverageFactor_(ratio) {
  var target = DISTRIBUTION_COVERAGE_TARGET;
  var penaltyFloor = DISTRIBUTION_COVERAGE_FLOOR;
  var cap = DISTRIBUTION_COVERAGE_CAP;
  var bonus = DISTRIBUTION_COVERAGE_BONUS;
  var value = Math.max(0, ratio || 0);
  if (value < target) {
    var factor = target ? value / target : 0;
    return Math.max(penaltyFloor, factor);
  }
  var capped = Math.min(value, cap);
  if (capped <= target) {
    return 1;
  }
  var range = cap - target;
  if (range <= 0) {
    return 1;
  }
  return 1 + bonus * ((capped - target) / range);
}

function clamp_(value, min, max) {
  if (value == null || isNaN(value)) return min;
  return Math.max(min, Math.min(max, value));
}

function computeSoftmaxShares_(items, tau, scoreKey) {
  if (!items || !items.length) {
    return {};
  }
  var t = Math.max(coerceNumber_(tau, 1), 0.0001);
  var scores = items.map(function (item) {
    var value = scoreKey ? item[scoreKey] : item.score;
    return coerceNumber_(value, 0);
  });
  var maxScore = Math.max.apply(Math, scores);
  if (!isFinite(maxScore)) {
    maxScore = 0;
  }
  var exps = scores.map(function (score) {
    return Math.exp((score - maxScore) / t);
  });
  var sum = exps.reduce(function (acc, val) { return acc + val; }, 0);
  var result = {};
  if (!sum) {
    var equal = 1 / items.length;
    items.forEach(function (item) { result[item.key] = equal; });
    return result;
  }
  items.forEach(function (item, index) {
    result[item.key] = exps[index] / sum;
  });
  return result;
}

function computeAdaptiveSoftmaxShares_(items, tau, scoreKey) {
  if (!items || !items.length) {
    return {};
  }
  if (items.length === 1) {
    var single = {};
    single[items[0].key] = 1;
    return single;
  }
  var baseTau = Math.max(coerceNumber_(tau, 1), 0.0001);
  var scores = items.map(function (item) {
    var value = scoreKey ? item[scoreKey] : item.score;
    return coerceNumber_(value, 0);
  });
  var sortedScores = scores.slice().sort(function (a, b) { return b - a; });
  var topScore = sortedScores[0];
  var runnerUp = sortedScores.length > 1 ? sortedScores[1] : topScore;
  var medianScore = sortedScores[Math.floor(sortedScores.length / 2)];
  var meanScore = average_(scores);
  var stdScore = standardDeviation_(scores);
  var scale = Math.max(
    Math.abs(topScore),
    Math.abs(runnerUp),
    Math.abs(medianScore),
    Math.abs(meanScore),
    1
  );
  var dominance = Math.max(0, topScore - runnerUp, topScore - medianScore) / scale;
  var dispersion = stdScore > 0 ? stdScore / scale : 0;
  var intensity = Math.min(3, Math.max(dominance, dispersion));
  var effectiveTau = baseTau / Math.max(1, 1 + 3 * intensity);
  if (effectiveTau < 0.35) {
    effectiveTau = 0.35;
  }
  return computeSoftmaxShares_(items, effectiveTau, scoreKey);
}

function normalizeSharesWithBounds_(items) {
  if (!items || !items.length) return;
  items.forEach(function (item) {
    item.minBound = clamp_(item.minBound != null ? item.minBound : 0, 0, 1);
    item.maxBound = clamp_(item.maxBound != null ? item.maxBound : 1, item.minBound, 1);
    item.value = clamp_(item.value != null ? item.value : 0, item.minBound, item.maxBound);
  });
  var minTotal = items.reduce(function (acc, item) { return acc + item.minBound; }, 0);
  if (minTotal > 1) {
    items.forEach(function (item) {
      var proportion = item.minBound / minTotal;
      item.value = proportion;
      item.minBound = proportion;
      item.maxBound = Math.max(item.maxBound, proportion);
    });
    return;
  }
  var maxTotal = items.reduce(function (acc, item) { return acc + item.maxBound; }, 0);
  if (maxTotal < 1) {
    var divisor = Math.max(maxTotal, 1e-6);
    items.forEach(function (item) {
      var proportion = item.maxBound / divisor;
      item.value = proportion;
      item.minBound = Math.min(item.minBound, proportion);
      item.maxBound = proportion;
    });
    return;
  }
  for (var iteration = 0; iteration < 12; iteration++) {
    var sum = items.reduce(function (acc, item) { return acc + item.value; }, 0);
    var diff = sum - 1;
    if (Math.abs(diff) < 1e-6) break;
    if (diff > 0) {
      var reducible = items.filter(function (item) { return item.value > item.minBound + 1e-9; });
      if (!reducible.length) break;
      var flex = reducible.reduce(function (acc, item) { return acc + (item.value - item.minBound); }, 0);
      if (flex <= 0) break;
      reducible.forEach(function (item) {
        var share = (item.value - item.minBound) / flex;
        item.value = Math.max(item.minBound, item.value - diff * share);
      });
    } else {
      var increasable = items.filter(function (item) { return item.value < item.maxBound - 1e-9; });
      if (!increasable.length) break;
      var flexGrow = increasable.reduce(function (acc, item) { return acc + (item.maxBound - item.value); }, 0);
      if (flexGrow <= 0) break;
      increasable.forEach(function (item) {
        var share = (item.maxBound - item.value) / flexGrow;
        item.value = Math.min(item.maxBound, item.value - diff * share);
      });
    }
  }
  var finalSum = items.reduce(function (acc, item) { return acc + item.value; }, 0);
  if (Math.abs(finalSum - 1) > 1e-6) {
    var residual = 1 - finalSum;
    for (var i = 0; i < items.length; i++) {
      var item = items[i];
      var capacity = residual > 0 ? item.maxBound - item.value : item.value - item.minBound;
      if (capacity <= 0) continue;
      var adjustment = Math.max(Math.min(residual, capacity), -capacity);
      item.value += adjustment;
      residual -= adjustment;
      if (Math.abs(residual) < 1e-6) break;
    }
  }
}

function computeMomentum_(values) {
  if (!values || !values.length) return 0;
  var latest = coerceNumber_(values[values.length - 1], 0);
  if (values.length < 2) return 0;
  var sum = 0;
  var count = 0;
  for (var i = 0; i < values.length - 1; i++) {
    var num = coerceNumber_(values[i], 0);
    if (!isNaN(num)) {
      sum += num;
      count++;
    }
  }
  if (!count) return 0;
  var avg = sum / count;
  if (!avg) return 0;
  return (latest - avg) / Math.max(avg, 1);
}

function average_(values) {
  if (!values || !values.length) return 0;
  var sum = 0;
  var count = 0;
  values.forEach(function (value) {
    var num = coerceNumber_(value, null);
    if (num !== null && !isNaN(num)) {
      sum += num;
      count++;
    }
  });
  return count ? sum / count : 0;
}

function standardDeviation_(values) {
  if (!values || !values.length) return 0;
  var avg = average_(values);
  if (!avg && avg !== 0) return 0;
  var sum = 0;
  var count = 0;
  values.forEach(function (value) {
    var num = coerceNumber_(value, null);
    if (num !== null && !isNaN(num)) {
      sum += Math.pow(num - avg, 2);
      count++;
    }
  });
  if (!count) return 0;
  return Math.sqrt(sum / count);
}

function syncLinkRouterDistribution(request) {
  if (!request) {
    throw new Error('Parâmetros inválidos.');
  }
  var operation = String(request.operation || '').trim();
  var traffic = String(request.traffic || '').trim();
  var routeSlug = String(request.routeSlug || '').trim();
  if (!operation) {
    throw new Error('Selecione a operação.');
  }
  if (!traffic) {
    throw new Error('Selecione o tráfego.');
  }
  var params = {
    operation: operation,
    traffic: traffic,
    routeSlug: routeSlug,
    startDate: request.startDate || '',
    endDate: request.endDate || '',
    urlWindow: request.window || 'day'
  };
  if (request.distribution) {
    params.distribution = request.distribution;
  }

  var dashboard = getDashboardData(params);
  var integration = dashboard && dashboard.integration ? dashboard.integration : null;
  if (!integration || !integration.companyId || !integration.domainId || !integration.routeSlug) {
    throw new Error('Integração Link Router não configurada para a operação selecionada.');
  }
  if (!integration.trafficType) {
    throw new Error('Integração sem tipo de tráfego vinculado. Atualize a configuração.');
  }
  if (integration.trafficType && integration.trafficType !== traffic) {
    throw new Error('Integração configurada para o tráfego "' + integration.trafficType + '", selecione o tipo correspondente.');
  }

  var distribution = dashboard && dashboard.distribution ? dashboard.distribution : null;
  if (!distribution || !distribution.globalUrls || !distribution.globalUrls.all || !distribution.globalUrls.all.length) {
    throw new Error('Plano de distribuição indisponível para sincronização.');
  }

  var executionMode = String(request.mode || '').toLowerCase() === 'auto' ? 'auto' : 'manual';
  var historyOriginSuffix = executionMode === 'auto' ? ' (Execução automática)' : '';

  var shareState = buildLinkRouterShareState_(distribution);
  var routeResponse = fetchLinkRouterRoute_(integration.companyId, integration.domainId, integration.routeSlug);
  if (!routeResponse.ok || !routeResponse.body || routeResponse.body.status !== 'success') {
    throw new Error('Falha ao consultar a rota na API: ' + formatLinkRouterError_(routeResponse));
  }
  var routeData = routeResponse.body.data || {};
  var links = routeData.links || [];
  if (!links.length) {
    throw new Error('Nenhum link encontrado na API para o slug informado.');
  }

  var linkMap = {};
  var pctMap = {};
  var historyRows = [];
  var unmatchedLinks = [];
  var timestamp = new Date();
  var totalPercent = 0;

  links.forEach(function (link) {
    var id = String(link && (link.id != null ? link.id : link.link_id != null ? link.link_id : '')).trim();
    if (!id) {
      return;
    }
    var linkUrl = String(link.link || link.url || '').trim();
    linkMap[id] = linkUrl;
    var shareInfo = matchShareInfoForLink_(linkUrl, shareState);
    var percent;
    if (shareInfo) {
      percent = shareInfo.share * 100;
    } else {
      unmatchedLinks.push(linkUrl);
      percent = link && link.percentage != null && link.percentage !== '' ? toNumber_(link.percentage) : 0;
    }
    if (isNaN(percent)) {
      percent = 0;
    }
    percent = Math.max(0, percent);
    percent = Math.round(percent * 100) / 100;
    pctMap[id] = percent;
    totalPercent += percent;
    historyRows.push([
      timestamp,
      operation,
      traffic,
      integration.companyId,
      integration.domainId,
      integration.routeSlug,
      linkUrl,
      percent,
      (shareInfo ? 'Plano sugerido' : 'Mantido existente') + historyOriginSuffix,
      '',
      ''
    ]);
  });

  var linkIds = Object.keys(pctMap);
  if (!linkIds.length) {
    throw new Error('Nenhum link válido retornado pela API para sincronização.');
  }

  if (totalPercent > 0 && Math.abs(totalPercent - 100) > 0.05) {
    var scale = 100 / totalPercent;
    totalPercent = 0;
    linkIds.forEach(function (id) {
      var scaled = Math.round(pctMap[id] * scale * 100) / 100;
      pctMap[id] = scaled;
      totalPercent += scaled;
    });
  }
  var diff = totalPercent > 0 ? Math.round((totalPercent - 100) * 100) / 100 : 0;
  if (totalPercent > 0 && Math.abs(diff) >= 0.01 && linkIds.length) {
    var adjustId = linkIds[linkIds.length - 1];
    pctMap[adjustId] = Math.max(0, Math.round((pctMap[adjustId] - diff) * 100) / 100);
    totalPercent = Math.round((totalPercent - diff) * 100) / 100;
  }

  var payload = {
    domain: parseLinkRouterId_(integration.domainId),
    company_id: parseLinkRouterId_(integration.companyId),
    slug: integration.routeSlug,
    link: linkMap,
    percentage: pctMap
  };

  var updateResponse = callLinkRouterApi_('/api/link_router/links_and_percentage', 'put', null, payload);
  var success = updateResponse.ok && updateResponse.body && updateResponse.body.status === 'success';
  var message = success ? 'Percentuais sincronizados com sucesso.' : formatLinkRouterError_(updateResponse);

  historyRows = historyRows.map(function (row) {
    row[9] = success ? 'OK' : 'ERRO';
    row[10] = message;
    return row;
  });
  appendLinkRouterHistory_(historyRows);

  if (!success) {
    throw new Error(message);
  }

  var leftovers = collectUnmatchedShares_(shareState);

  return {
    success: true,
    message: message,
    dashboard: dashboard,
    linksUpdated: linkIds.length,
    totalPercentage: totalPercent,
    unmatchedLinks: unmatchedLinks,
    leftoverShares: leftovers
  };
}

function parseLinkRouterId_(value) {
  var num = Number(value);
  return isNaN(num) ? value : num;
}

function runLinkRouterAutoSync() {
  var snapshot = getConfigSnapshot_();
  var integrations = (snapshot.operationIntegrationRows || []).filter(function (row) {
    return row && row.active && row.operation && row.companyId && row.domainId && row.routeSlug && row.trafficType;
  });
  var trafficByType = snapshot.trafficByType || {};
  var successes = [];
  var failures = [];

  integrations.forEach(function (integration) {
    var trafficType = integration.trafficType;
    var sources = trafficByType[trafficType] || [];
    if (!sources.length) {
      var error = new Error('Nenhuma utm_source configurada para o tráfego vinculado.');
      failures.push({
        operation: integration.operation,
        traffic: trafficType,
        message: error.message
      });
      logLinkRouterAutoSyncFailure_(integration, trafficType, error);
      return;
    }
    try {
      var result = syncLinkRouterDistribution({
        operation: integration.operation,
        traffic: trafficType,
        routeSlug: integration.routeSlug,
        window: 'day',
        mode: 'auto'
      });
      successes.push({
        operation: integration.operation,
        traffic: trafficType,
        linksUpdated: result && result.linksUpdated ? result.linksUpdated : 0,
        unmatchedLinks: result && result.unmatchedLinks ? result.unmatchedLinks.length : 0,
        leftoverShares: result && result.leftoverShares ? result.leftoverShares.length : 0
      });
    } catch (err) {
      failures.push({
        operation: integration.operation,
        traffic: trafficType,
        message: err && err.message ? err.message : String(err)
      });
      logLinkRouterAutoSyncFailure_(integration, trafficType, err);
    }
  });

  updateLinkRouterAutoSyncTrigger_(snapshot);

  return {
    timestamp: new Date(),
    successes: successes,
    failures: failures
  };
}

function updateLinkRouterAutoSyncTrigger_(snapshot) {
  snapshot = snapshot || getConfigSnapshot_();
  var hasActiveIntegration = (snapshot.operationIntegrationRows || []).some(function (row) {
    return row && row.active && row.operation && row.companyId && row.domainId && row.routeSlug && row.trafficType;
  });
  var triggers = ScriptApp.getProjectTriggers().filter(function (trigger) {
    return trigger.getHandlerFunction && trigger.getHandlerFunction() === 'runLinkRouterAutoSync';
  });

  if (hasActiveIntegration) {
    if (triggers.length) {
      triggers.forEach(function (trigger) {
        ScriptApp.deleteTrigger(trigger);
      });
    }
    ScriptApp.newTrigger('runLinkRouterAutoSync')
      .timeBased()
      .everyHours(1)
      .create();
  } else if (triggers.length) {
    triggers.forEach(function (trigger) {
      ScriptApp.deleteTrigger(trigger);
    });
  }
}

function logLinkRouterAutoSyncFailure_(integration, trafficType, error) {
  var message = error && error.message ? error.message : String(error || 'Erro desconhecido');
  var row = [
    new Date(),
    integration && integration.operation ? integration.operation : '',
    trafficType || '',
    integration && integration.companyId ? integration.companyId : '',
    integration && integration.domainId ? integration.domainId : '',
    integration && integration.routeSlug ? integration.routeSlug : '',
    '',
    '',
    'Execução automática',
    'ERRO',
    message
  ];
  appendLinkRouterHistory_([row]);
}

function appendLinkRouterHistory_(rows) {
  if (!rows || !rows.length) {
    return;
  }
  var sheet = ensureLinkRouterHistorySheet_();
  var startRow = Math.max(sheet.getLastRow(), 1) + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
}

function buildLinkRouterShareState_(distribution) {
  var entries = (distribution && distribution.globalUrls && distribution.globalUrls.all) ? distribution.globalUrls.all : [];
  var exact = {};
  var queue = {};
  var infos = [];
  entries.forEach(function (entry) {
    var info = {
      url: entry.url || '',
      share: toNumber_(entry.suggestedGlobalShare || 0),
      siteShare: toNumber_(entry.suggestedShare || 0),
      site: entry.site || '',
      matched: false,
      original: entry
    };
    infos.push(info);
    var fullKey = String(info.url || '').trim().toLowerCase();
    if (fullKey) {
      if (!exact[fullKey]) {
        exact[fullKey] = [];
      }
      exact[fullKey].push(info);
    }
    var normKey = normUrlStrict_(info.url || '');
    if (normKey) {
      if (!queue[normKey]) {
        queue[normKey] = [];
      }
      queue[normKey].push(info);
    }
  });
  Object.keys(exact).forEach(function (key) {
    exact[key].sort(function (a, b) { return (b.share || 0) - (a.share || 0); });
  });
  Object.keys(queue).forEach(function (key) {
    queue[key].sort(function (a, b) { return (b.share || 0) - (a.share || 0); });
  });
  return { exact: exact, queue: queue, infos: infos };
}

function matchShareInfoForLink_(url, state) {
  if (!url || !state) {
    return null;
  }
  var fullKey = String(url || '').trim().toLowerCase();
  var exactList = fullKey ? state.exact[fullKey] : null;
  if (exactList && exactList.length) {
    while (exactList.length) {
      var exactCandidate = exactList.shift();
      if (!exactCandidate.matched) {
        exactCandidate.matched = true;
        return exactCandidate;
      }
    }
  }
  var normKey = normUrlStrict_(url || '');
  var list = normKey ? state.queue[normKey] : null;
  if (list && list.length) {
    while (list.length) {
      var candidate = list.shift();
      if (!candidate.matched) {
        candidate.matched = true;
        return candidate;
      }
    }
  }
  return null;
}

function collectUnmatchedShares_(state) {
  if (!state || !state.infos) {
    return [];
  }
  var leftovers = [];
  state.infos.forEach(function (info) {
    if (!info.matched && toNumber_(info.share || 0) > 0.0005) {
      leftovers.push({ url: info.url || '', share: toNumber_(info.share || 0) });
    }
  });
  return leftovers;
}

function buildLinkRouterUrlIndex_(links) {
  var entries = [];
  var seenNormalized = {};
  var seenRaw = {};
  (links || []).forEach(function (item) {
    if (!item) return;
    var raw = '';
    if (item.link != null && item.link !== '') {
      raw = item.link;
    } else if (item.url != null && item.url !== '') {
      raw = item.url;
    }
    raw = String(raw || '').trim();
    if (!raw) return;
    var normalized = normUrlStrict_(raw);
    var rawKey = raw.toLowerCase();
    if (normalized) {
      if (seenNormalized[normalized]) {
        return;
      }
      seenNormalized[normalized] = true;
    } else if (rawKey) {
      if (seenRaw[rawKey]) {
        return;
      }
      seenRaw[rawKey] = true;
    }
    entries.push({
      url: raw,
      normalized: normalized,
      rawKey: rawKey,
      key: normalized || rawKey
    });
  });
  var keySet = {};
  entries.forEach(function (entry) {
    if (entry.normalized) {
      keySet[entry.normalized] = true;
    }
    if (entry.rawKey) {
      keySet[entry.rawKey] = true;
    }
  });
  return {
    urls: entries,
    keySet: keySet
  };
}

function fetchLinkRouterRoute_(companyId, domainId, slug) {
  var params = {
    company_id: companyId,
    domain: domainId,
    slug: slug
  };
  return callLinkRouterApi_('/api/link_router/route', 'get', params, null);
}

function callLinkRouterApi_(path, method, qs, payload) {
  var baseUrl = String(LINK_ROUTER_BASE_URL || '').replace(/\/+$/, '');
  var url = baseUrl + path;
  if (qs && Object.keys(qs).length) {
    var query = Object.keys(qs).filter(function (key) {
      var value = qs[key];
      return value !== null && value !== undefined && value !== '';
    }).map(function (key) {
      return encodeURIComponent(key) + '=' + encodeURIComponent(String(qs[key]));
    }).join('&');
    if (query) {
      url += (url.indexOf('?') >= 0 ? '&' : '?') + query;
    }
  }
  var options = {
    method: (method || 'get').toUpperCase(),
    muteHttpExceptions: true,
    headers: buildLinkRouterHeaders_()
  };
  if (payload) {
    options.payload = JSON.stringify(payload);
    options.contentType = 'application/json';
  }
  var response;
  try {
    response = UrlFetchApp.fetch(url, options);
  } catch (err) {
    return { ok: false, statusCode: 0, body: { status: 'error', data: err && err.message ? err.message : String(err) } };
  }
  var statusCode = response.getResponseCode();
  var text = response.getContentText();
  var body;
  try {
    body = text ? JSON.parse(text) : {};
  } catch (err) {
    body = { status: 'error', data: text || 'Resposta inválida da API.' };
  }
  var ok = statusCode >= 200 && statusCode < 300;
  if (!ok && body && typeof body === 'object') {
    body.status = body.status || 'error';
  }
  return { ok: ok, statusCode: statusCode, body: body };
}

function buildLinkRouterHeaders_() {
  return {
    'Authorization': 'Bearer ' + getLinkRouterToken_(),
    'Content-Type': 'application/json'
  };
}

function getLinkRouterToken_() {
  var token = String(LINK_ROUTER_TOKEN || '').trim();
  if (/^Bearer\s+/i.test(token)) {
    token = token.replace(/^Bearer\s+/i, '');
  }
  return token;
}

function formatLinkRouterError_(response) {
  if (!response) {
    return 'Erro desconhecido.';
  }
  if (response.body) {
    var body = response.body;
    if (body.message) {
      return body.message;
    }
    if (typeof body.data === 'string' && body.data) {
      return body.data;
    }
    if (body.data && typeof body.data === 'object') {
      if (Array.isArray(body.data)) {
        return body.data.join('; ');
      }
      var parts = [];
      Object.keys(body.data).forEach(function (key) {
        parts.push(key + ': ' + body.data[key]);
      });
      if (parts.length) {
        return parts.join('; ');
      }
    }
    if (body.status && body.status !== 'success') {
      return 'Erro na API (' + body.status + ').';
    }
  }
  if (response.statusCode) {
    return 'Erro ' + response.statusCode + ' ao chamar a API.';
  }
  return 'Erro desconhecido na chamada da API.';
}
