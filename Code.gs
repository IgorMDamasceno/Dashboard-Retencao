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
  if (targetRow === -1) {
    for (var j = 0; j < range.length; j++) {
      if (String(range[j][idx.operacao] || '').trim() === operation) {
        targetRow = j + 2;
        id = String(range[j][idx.id] || '').trim() || Utilities.getUuid();
        break;
      }
    }
  }
  if (!id) {
    id = Utilities.getUuid();
  }
  var statusValue = active ? 'Ativo' : 'Inativo';
  var rowValues = [id, operation, companyId, domainId, routeSlug, statusValue];
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
  const integrationConfig = snapshot.operationIntegrationsByOperation[operationName] || null;
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
  const urlRows = filterRowsForWindow_(filteredRows, hoursInfo, window);
  const siteSummary = buildSiteSummary_(
    filteredRows,
    hoursInfo,
    snapshot.revshareMap,
    urlRows
  );
  const urlSummary = buildUrlSummary_(urlRows, snapshot.revshareMap);
  const detailed = buildDetailedAnalysis_(filteredRows, hoursInfo, snapshot.revshareMap);

  var savedControls = getDistributionControlsConfig_(operationName, trafficType);
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
    filteredRows,
    hoursInfo,
    snapshot.revshareMap,
    distributionParams
  );

  if (params.distribution && Object.keys(params.distribution).length) {
    saveDistributionControlsConfig_(operationName, trafficType, distribution.controls);
  }
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
    url: urlSummary,
    detailed: detailed,
    distribution: distribution,
    integration: integrationConfig ? {
      id: integrationConfig.id || '',
      operation: integrationConfig.operation || '',
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
    totalSessions: 10000,
    tau: 60,
    tauLocal: 50,
    reliabilityK: 100,
    ucbZ: 10,
    siteMinShare: 0.05,
    siteMaxShare: 0.35,
    siteStep: 0.4,
    siteMinSessions: 300,
    urlSeedPercent: 0.02,
    urlSeedSessions: 50,
    urlMinRecipients: 3,
    urlTargetRecipients: 0,
    urlPriorityShare: 0.7,
    urlRequireAll: 0,
    urlMaxShare: 0.4,
    urlStep: 0.35,
    coverageTarget: 0.75,
    coverageCap: 0.95,
    coverageBonusFactor: 0.2,
    coveragePenaltyFloor: 0.5,
    modeThresholdRps: 250,
    modeThresholdCv: 0.15,
    exploitPortion: 0.7,
    explorePortion: 0.3,
    momentumExploreThreshold: 0.1,
    coverageExploreThreshold: 0.7,
    controlReserve: 0.1
  };
}

function buildDistributionControls_(params) {
  var defaults = getDefaultDistributionControls_();
  var controls = {};
  params = params || {};
  Object.keys(defaults).forEach(function (key) {
    var value = params.hasOwnProperty(key) ? params[key] : defaults[key];
    var parsed = coerceNumber_(value, defaults[key]);
    if (key === 'totalSessions') {
      controls[key] = Math.max(1, Math.round(parsed));
    } else if (key === 'tau' || key === 'tauLocal') {
      controls[key] = Math.max(1, parsed);
    } else if (key === 'reliabilityK' || key === 'siteMinSessions' || key === 'urlSeedSessions' || key === 'urlMinRecipients' || key === 'urlTargetRecipients') {
      controls[key] = Math.max(0, Math.round(parsed));
    } else if (key === 'ucbZ') {
      controls[key] = Math.max(0, parsed);
    } else if (key === 'urlRequireAll') {
      controls[key] = parsed ? 1 : 0;
    } else if (key === 'urlPriorityShare') {
      controls[key] = Math.max(0, Math.min(1, parsed));
    } else {
      controls[key] = Math.max(0, parsed);
    }
  });
  if (controls.exploitPortion + controls.explorePortion > 1) {
    var total = controls.exploitPortion + controls.explorePortion;
    controls.exploitPortion = controls.exploitPortion / total;
    controls.explorePortion = controls.explorePortion / total;
  }
  return controls;
}

function buildDistributionPlan_(rows, hoursInfo, revshareMap, params) {
  var controls = buildDistributionControls_(params);
  var globalLimitEnabled = !controls.urlRequireAll && Math.max(0, Math.round(controls.urlTargetRecipients || 0)) > 0;
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
      var weight = urlSessions / (urlSessions + controls.reliabilityK);
      var momentum = computeMomentum_(urlHourRps);
      var coverageFactor = computeCoverageFactor_(coverageRatio, controls);
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
    var siteCoverageFactor = computeCoverageFactor_(siteCoverageRatio, controls);
    var siteEcpm = coverageRequests ? ecpmWeighted / coverageRequests : 0;
    var siteEcpmEff = siteEcpm * siteCoverageRatio;
    var siteEcpmScore = siteEcpmEff * siteCoverageFactor;
    var siteMomentum = computeMomentum_(hourRps);
    var siteConfidence = sessionsWindow / (sessionsWindow + controls.reliabilityK);
    var siteScore = siteEcpmScore;
    if (siteCoverageRatio >= controls.coverageTarget) {
      siteScore *= 1.02;
    }

    var roundsTotal = 0;
    var roundsCount = 0;
    urlEntries.forEach(function (entry) {
      var scoreBase = entry.weight * entry.ecpmScore + (1 - entry.weight) * siteEcpmScore;
      var momentumFactor = 1 + 0.15 * clamp_(entry.momentum, -0.5, 0.5);
      entry.baseScore = scoreBase;
      entry.momentumFactor = momentumFactor;
      entry.siteEcpmScore = siteEcpmScore;
      entry.score = scoreBase * momentumFactor;
      roundsTotal += entry.rounds;
      roundsCount++;
    });
    var R = Math.max(roundsTotal, roundsCount, 1);
    var lnR = Math.log(R + 1);
    urlEntries.forEach(function (entry) {
      entry.explorationBonus = controls.ucbZ * Math.sqrt(lnR / (entry.rounds + 1));
    });

    var maxRawScore = 0;
    var maxEcpmEff = 0;
    var maxRpsEff = 0;
    urlEntries.forEach(function (entry) {
      var scoreValue = Math.max(0, coerceNumber_(entry.score, 0));
      var ecpmEffValue = Math.max(0, coerceNumber_(entry.ecpmEff, 0));
      var rpsValue = Math.max(0, coerceNumber_(entry.rps, 0));
      if (scoreValue > maxRawScore) maxRawScore = scoreValue;
      if (ecpmEffValue > maxEcpmEff) maxEcpmEff = ecpmEffValue;
      if (rpsValue > maxRpsEff) maxRpsEff = rpsValue;
    });

    urlEntries.forEach(function (entry) {
      var normalizedScore = maxRawScore > 0 ? Math.max(0, coerceNumber_(entry.score, 0)) / maxRawScore : 0;
      var normalizedEcpm = maxEcpmEff > 0 ? Math.max(0, coerceNumber_(entry.ecpmEff, 0)) / maxEcpmEff : 0;
      var normalizedRps = maxRpsEff > 0 ? Math.max(0, coerceNumber_(entry.rps, 0)) / maxRpsEff : 0;
      var components = [normalizedScore, normalizedEcpm, normalizedRps];
      entry.performanceScore = components.length ? average_(components) : 0;
      entry.normalizedScore = normalizedScore;
      entry.normalizedEcpmEff = normalizedEcpm;
      entry.normalizedRpsEff = normalizedRps;
      entry.ucbScore = entry.performanceScore + entry.explorationBonus;
    });

    var apostas = false;
    if (urlEntries.length) {
      var sortedByRps = urlEntries.slice().sort(function (a, b) { return b.rps - a.rps; });
      var quartileCount = Math.max(1, Math.ceil(sortedByRps.length / 4));
      var topQuartil = sortedByRps.slice(0, quartileCount).filter(function (entry) { return entry.sessions > 0; });
      if (topQuartil.length) {
        var mean = average_(topQuartil.map(function (entry) { return entry.rps; }));
        var std = standardDeviation_(topQuartil.map(function (entry) { return entry.rps; }));
        var cv = mean ? std / mean : 0;
        apostas = mean < controls.modeThresholdRps && cv < controls.modeThresholdCv;
      }
    }

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
      urls: urlEntries,
      apostas: apostas
    });
  });

  if (!siteRows.length) {
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

  var siteItems = siteRows.map(function (row) {
    var stateRow = siteStateMap[row.site] || {};
    var prevShare = stateRow.currentShare != null && stateRow.currentShare !== '' ? normalizeShareValue_(stateRow.currentShare) : null;
    if (prevShare == null && stateRow.lastShare != null && stateRow.lastShare !== '') {
      prevShare = normalizeShareValue_(stateRow.lastShare);
    }
    return {
      key: row.site,
      score: row.score,
      state: stateRow,
      prevShare: prevShare
    };
  });

  var siteSoftmax = computeSoftmaxShares_(siteItems, controls.tau, 'score');
  var minSiteShareBase = controls.totalSessions > 0 ? (controls.siteMinSessions / controls.totalSessions) : 0;
  var minSiteShare = Math.max(controls.siteMinShare, minSiteShareBase);
  var siteBounds = siteItems.map(function (item) {
    var minBound = clamp_(minSiteShare, 0, 1);
    var maxBound = clamp_(controls.siteMaxShare, minBound, 1);
    if (item.prevShare != null && item.prevShare > 0) {
      minBound = Math.max(minBound, item.prevShare * (1 - controls.siteStep));
      maxBound = Math.min(maxBound, item.prevShare * (1 + controls.siteStep));
    }
    minBound = clamp_(minBound, 0, 1);
    maxBound = clamp_(maxBound, minBound, 1);
    return {
      key: item.key,
      value: siteSoftmax[item.key] != null ? siteSoftmax[item.key] : 0,
      minBound: minBound,
      maxBound: maxBound,
      prevShare: item.prevShare,
      state: item.state
    };
  });
  normalizeSharesWithBounds_(siteBounds);

  var siteOutput = [];
  var globalAggregates = [];
  var defaultSite = '';
  var bestShare = -1;

  siteBounds.forEach(function (item) {
    var siteRow = siteRows.find(function (row) { return row.site === item.key; });
    if (!siteRow) return;
    var siteState = item.state || {};
    var siteCurrentShare = siteState.currentShare != null && siteState.currentShare !== '' ? normalizeShareValue_(siteState.currentShare) : 0;
    var siteSuggestedShare = item.value;
    var siteDeltaShare = siteSuggestedShare - siteCurrentShare;
    var totalSessions = controls.totalSessions;
    var siteCurrentSessions = siteCurrentShare * totalSessions;
    var siteSuggestedSessions = siteSuggestedShare * totalSessions;
    if (siteSuggestedShare > bestShare) {
      bestShare = siteSuggestedShare;
      defaultSite = siteRow.site;
    }

    siteRow.currentShare = siteCurrentShare;
    siteRow.suggestedShare = siteSuggestedShare;
    siteRow.deltaShare = siteDeltaShare;
    siteRow.currentSessions = siteCurrentSessions;
    siteRow.suggestedSessions = siteSuggestedSessions;
    siteRow.nextEvaluation = siteState.nextEvalHour || '';
    siteRow.lastChange = siteState.lastChangeHour || '';

    var siteUrls = siteRow.urls || [];
    var siteSessionsReference = Math.max(siteSuggestedSessions, siteCurrentSessions);
    var minSeedShare = siteSessionsReference > 0 ? Math.max(controls.urlSeedPercent, controls.urlSeedSessions / siteSessionsReference) : controls.urlSeedPercent;
    minSeedShare = clamp_(minSeedShare, 0, 1);

    var promising = siteUrls.filter(function (entry) {
      return entry.momentum > controls.momentumExploreThreshold && entry.coverageRatio >= controls.coverageExploreThreshold;
    });
    var exploit = siteUrls.slice().sort(function (a, b) { return (b.performanceScore || 0) - (a.performanceScore || 0); }).slice(0, Math.min(3, siteUrls.length));
    var explore = siteRow.apostas ? promising : [];
    if (siteRow.apostas && !explore.length) {
      explore = siteUrls.slice().sort(function (a, b) { return a.rounds - b.rounds; }).slice(0, Math.min(3, siteUrls.length));
    }

    var baseShareMap = {};
    if (!siteUrls.length) {
      baseShareMap = {};
    } else if (siteRow.apostas) {
      var exploitMap = computeAdaptiveSoftmaxShares_(exploit, controls.tauLocal, 'ucbScore');
      var exploreMap = explore.length ? computeAdaptiveSoftmaxShares_(explore, controls.tauLocal, 'ucbScore') : {};
      siteUrls.forEach(function (entry) { baseShareMap[entry.key] = 0; });
      if (exploit.length) {
        exploit.forEach(function (entry) {
          baseShareMap[entry.key] += controls.exploitPortion * (exploitMap[entry.key] || 0);
        });
      }
      if (explore.length) {
        explore.forEach(function (entry) {
          baseShareMap[entry.key] += controls.explorePortion * (exploreMap[entry.key] || 0);
        });
      }
      var baseTotal = 0;
      siteUrls.forEach(function (entry) { baseTotal += baseShareMap[entry.key] || 0; });
      if (baseTotal < 1 && exploit.length) {
        var remaining = 1 - baseTotal;
        exploit.forEach(function (entry) {
          baseShareMap[entry.key] += remaining * (exploitMap[entry.key] || 0);
        });
      }
      if (!exploit.length) {
        baseShareMap = computeAdaptiveSoftmaxShares_(siteUrls, controls.tauLocal, 'ucbScore');
      }
    } else {
      baseShareMap = computeAdaptiveSoftmaxShares_(siteUrls, controls.tauLocal, 'ucbScore');
    }

    var requiredKeys = {};
    var priorityKeys = {};
    var requiredCount = 0;
    if (controls.urlRequireAll && siteUrls.length) {
      requiredCount = siteUrls.length;
    } else if (!globalLimitEnabled) {
      requiredCount = Math.min(siteUrls.length, Math.max(0, controls.urlMinRecipients));
    }
    if (requiredCount > 0 && siteUrls.length) {
      var priority = siteUrls.slice().sort(function (a, b) {
        var perfDiff = (b.performanceScore || 0) - (a.performanceScore || 0);
        if (Math.abs(perfDiff) > 1e-6) {
          return perfDiff;
        }
        var bBase = baseShareMap[b.key] != null ? baseShareMap[b.key] : 0;
        var aBase = baseShareMap[a.key] != null ? baseShareMap[a.key] : 0;
        if (Math.abs(bBase - aBase) > 1e-6) {
          return bBase - aBase;
        }
        return (b.score || 0) - (a.score || 0);
      });
      for (var r = 0; r < requiredCount && r < priority.length; r++) {
        priorityKeys[priority[r].key] = true;
        requiredKeys[priority[r].key] = true;
      }
    }
    siteUrls.forEach(function (entry) {
      var statusValue = String((entry.state && entry.state.status) || '').toLowerCase();
      if (statusValue === 'controle') {
        requiredKeys[entry.key] = true;
      }
    });

    var urlItems = siteUrls.map(function (entry) {
      var stateEntry = entry.state || {};
      var prevShare = stateEntry.currentShare != null && stateEntry.currentShare !== '' ? normalizeShareValue_(stateEntry.currentShare) : null;
      if (prevShare == null && stateEntry.lastShare != null && stateEntry.lastShare !== '') {
        prevShare = normalizeShareValue_(stateEntry.lastShare);
      }
      var isRequired = !!requiredKeys[entry.key];
      var isPriority = !!priorityKeys[entry.key];
      entry.isPriority = isPriority;
      var minBound = isRequired ? minSeedShare : 0;
      if (String(stateEntry.status || '').toLowerCase() === 'controle') {
        minBound = Math.max(minBound, controls.controlReserve);
      }
      var maxBound = controls.urlMaxShare;
      if (prevShare != null && prevShare > 0) {
        minBound = Math.max(minBound, prevShare * (1 - controls.urlStep));
        maxBound = Math.min(maxBound, prevShare * (1 + controls.urlStep));
      }
      minBound = clamp_(minBound, 0, 1);
      maxBound = clamp_(maxBound, minBound, 1);
      return {
        key: entry.key,
        value: baseShareMap[entry.key] != null ? baseShareMap[entry.key] : (1 / Math.max(1, siteUrls.length)),
        minBound: minBound,
        maxBound: maxBound,
        prevShare: prevShare,
        entry: entry,
        state: stateEntry,
        required: isRequired,
        priority: isPriority
      };
    });

    var desiredPriorityShare = clamp_(controls.urlPriorityShare != null ? controls.urlPriorityShare : 0, 0, 1);
    var priorityItems = urlItems.filter(function (item) { return item.priority; });
    var complementaryItems = urlItems.filter(function (item) { return !item.priority; });
    if (priorityItems.length && (desiredPriorityShare > 0 || !complementaryItems.length)) {
      var allocationPriorityShare = complementaryItems.length ? desiredPriorityShare : 1;
      var priorityWeight = priorityItems.reduce(function (acc, item) {
        return acc + Math.max(0, item.entry.performanceScore || 0);
      }, 0);
      if (priorityWeight <= 0) {
        priorityWeight = priorityItems.length;
      }
      priorityItems.forEach(function (item) {
        var weight = Math.max(0, item.entry.performanceScore || 0);
        if (priorityWeight <= 0) {
          item.value = allocationPriorityShare / priorityItems.length;
        } else {
          item.value = allocationPriorityShare * (weight > 0 ? weight / priorityWeight : 1 / priorityItems.length);
        }
      });
      var remainingShare = Math.max(0, 1 - allocationPriorityShare);
      if (complementaryItems.length) {
        if (remainingShare <= 0) {
          complementaryItems.forEach(function (item) { item.value = 0; });
        } else {
          var complementaryWeight = complementaryItems.reduce(function (acc, item) {
            return acc + Math.max(0, item.entry.performanceScore || 0);
          }, 0);
          if (complementaryWeight <= 0) {
            complementaryWeight = complementaryItems.length;
          }
          complementaryItems.forEach(function (item) {
            var weight = Math.max(0, item.entry.performanceScore || 0);
            if (complementaryWeight <= 0) {
              item.value = remainingShare / complementaryItems.length;
            } else {
              item.value = remainingShare * (weight > 0 ? weight / complementaryWeight : 1 / complementaryItems.length);
            }
          });
        }
      }
    }

    normalizeSharesWithBounds_(urlItems);

    var siteUrlRows = urlItems.map(function (item) {
      var entry = item.entry;
      var stateEntry = item.state || {};
      var urlCurrentShare = stateEntry.currentShare != null && stateEntry.currentShare !== '' ? normalizeShareValue_(stateEntry.currentShare) : 0;
      var urlSuggestedShare = item.value;
      var urlDeltaShare = urlSuggestedShare - urlCurrentShare;
      var currentSessions = urlCurrentShare * siteCurrentSessions;
      var suggestedSessions = urlSuggestedShare * siteSuggestedSessions;
      var currentGlobalShare = urlCurrentShare * siteCurrentShare;
      var suggestedGlobalShare = urlSuggestedShare * siteSuggestedShare;
      var deltaGlobalShare = suggestedGlobalShare - currentGlobalShare;
      var status = stateEntry.status || (entry.rounds <= 1 ? 'seed' : 'ok');
      return {
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
        rounds: entry.rounds,
        score: entry.score,
        ucb: entry.ucbScore,
        performanceScore: entry.performanceScore,
        performanceNormalizedScore: entry.normalizedScore,
        performanceNormalizedEcpm: entry.normalizedEcpmEff,
        performanceNormalizedRps: entry.normalizedRpsEff,
        explorationBonus: entry.explorationBonus,
        weight: entry.weight,
        ecpmScore: entry.ecpmScore,
        siteEcpmScore: entry.siteEcpmScore,
        baseScore: entry.baseScore,
        momentumFactor: entry.momentumFactor,
        currentShare: urlCurrentShare,
        suggestedShare: urlSuggestedShare,
        deltaShare: urlDeltaShare,
        status: status,
        lastChange: stateEntry.lastChangeHour || '',
        nextEval: stateEntry.nextEvalHour || '',
        currentSessions: currentSessions,
        suggestedSessions: suggestedSessions,
        currentGlobalShare: currentGlobalShare,
        suggestedGlobalShare: suggestedGlobalShare,
        deltaGlobalShare: deltaGlobalShare,
        guaranteed: !!item.required,
        limited: false,
        priority: !!item.priority,
        siteEcpmEff: siteRow.ecpmEff,
        siteCoverage: siteRow.coverage,
        siteMode: siteRow.apostas ? 'Apostas' : 'Normal',
        siteMomentum: siteRow.momentum,
        siteScore: siteRow.score
      };
    });

    urlsBySite[siteRow.site] = siteUrlRows;
    Array.prototype.push.apply(globalAggregates, siteUrlRows);
  });

  var requestedRecipients = Math.max(0, Math.round(controls.urlTargetRecipients || 0));
  var targetRecipients = requestedRecipients;
  var limitApplied = false;
  if (!controls.urlRequireAll && targetRecipients > 0 && globalAggregates.length > targetRecipients) {
    limitApplied = true;
    var selectedKeys = {};
    globalAggregates.forEach(function (row) {
      if (row.guaranteed) {
        selectedKeys[row.key] = true;
      }
    });
    var sortedForSelection = globalAggregates.slice().sort(function (a, b) {
      var shareDiff = (b.suggestedGlobalShare || 0) - (a.suggestedGlobalShare || 0);
      if (Math.abs(shareDiff) > 1e-6) return shareDiff;
      var priorityDiff = (b.priority ? 1 : 0) - (a.priority ? 1 : 0);
      if (priorityDiff !== 0) return priorityDiff;
      var perfDiff = (b.performanceScore || 0) - (a.performanceScore || 0);
      if (Math.abs(perfDiff) > 1e-6) return perfDiff;
      return (b.score || 0) - (a.score || 0);
    });
    var selectedCount = Object.keys(selectedKeys).length;
    if (targetRecipients > 0 && selectedCount > targetRecipients) {
      targetRecipients = selectedCount;
    }
    for (var s = 0; s < sortedForSelection.length && selectedCount < targetRecipients; s++) {
      var candidate = sortedForSelection[s];
      if (!selectedKeys[candidate.key]) {
        selectedKeys[candidate.key] = true;
        selectedCount++;
      }
    }
    Object.keys(urlsBySite).forEach(function (siteKey) {
      var siteRowsList = urlsBySite[siteKey] || [];
      if (!siteRowsList.length) return;
      var hasSelected = siteRowsList.some(function (row) { return selectedKeys[row.key]; });
      if (!hasSelected) {
        var bestRow = siteRowsList.slice().sort(function (a, b) {
          var shareDiff = (b.suggestedGlobalShare || 0) - (a.suggestedGlobalShare || 0);
          if (Math.abs(shareDiff) > 1e-6) return shareDiff;
          var priorityDiff = (b.priority ? 1 : 0) - (a.priority ? 1 : 0);
          if (priorityDiff !== 0) return priorityDiff;
          var perfDiff = (b.performanceScore || 0) - (a.performanceScore || 0);
          if (Math.abs(perfDiff) > 1e-6) return perfDiff;
          return (b.score || 0) - (a.score || 0);
        })[0];
        if (bestRow && selectedCount < targetRecipients) {
          selectedKeys[bestRow.key] = true;
          selectedCount++;
        }
      }
    });

    siteRows.forEach(function (siteRow) {
      var siteKey = siteRow.site;
      var siteUrlsList = urlsBySite[siteKey] || [];
      if (!siteUrlsList.length) return;
      var selectedRows = [];
      var sumSelected = 0;
      siteUrlsList.forEach(function (row) {
        if (selectedKeys[row.key]) {
          selectedRows.push(row);
          sumSelected += row.suggestedShare;
        }
      });
      if (!selectedRows.length) {
        siteUrlsList.forEach(function (row) {
          row.limited = true;
          row.suggestedShare = 0;
          row.suggestedSessions = 0;
          row.suggestedGlobalShare = 0;
          row.deltaShare = 0 - (row.currentShare || 0);
          row.deltaGlobalShare = 0 - (row.currentGlobalShare || 0);
        });
        siteRow.suggestedShare = 0;
        siteRow.suggestedSessions = 0;
        siteRow.deltaShare = 0 - (siteRow.currentShare || 0);
        return;
      }
      if (sumSelected <= 0 && selectedRows.length) {
        var equalShare = 1 / selectedRows.length;
        selectedRows.forEach(function (row) {
          row.suggestedShare = equalShare;
        });
        sumSelected = 1;
      }
      var siteSuggestedShare = siteRow.suggestedShare || 0;
      var siteSuggestedSessions = siteRow.suggestedSessions || 0;
      siteUrlsList.forEach(function (row) {
        if (!selectedKeys[row.key]) {
          row.limited = true;
          row.suggestedShare = 0;
          row.suggestedSessions = 0;
          row.suggestedGlobalShare = 0;
          row.deltaShare = 0 - row.currentShare;
          row.deltaGlobalShare = 0 - row.currentGlobalShare;
        } else {
          var normalizedShare = row.suggestedShare / sumSelected;
          row.limited = false;
          row.suggestedShare = normalizedShare;
          row.suggestedSessions = normalizedShare * siteSuggestedSessions;
          row.suggestedGlobalShare = normalizedShare * siteSuggestedShare;
          var currentShare = row.currentShare || 0;
          var currentGlobal = row.currentGlobalShare || 0;
          row.deltaShare = normalizedShare - currentShare;
          row.deltaGlobalShare = row.suggestedGlobalShare - currentGlobal;
        }
      });
    });
  }

  var activeSiteShare = 0;
  siteRows.forEach(function (siteRow) {
    activeSiteShare += siteRow.suggestedShare || 0;
  });
  if (activeSiteShare > 0 && Math.abs(activeSiteShare - 1) > 1e-6) {
    var siteScale = 1 / activeSiteShare;
    siteRows.forEach(function (siteRow) {
      var originalSuggested = siteRow.suggestedShare || 0;
      if (originalSuggested <= 0) {
        siteRow.suggestedShare = 0;
        siteRow.suggestedSessions = 0;
        siteRow.deltaShare = 0 - (siteRow.currentShare || 0);
        return;
      }
      var scaledShare = clamp_(originalSuggested * siteScale, 0, 1);
      siteRow.suggestedShare = scaledShare;
      siteRow.suggestedSessions = scaledShare * controls.totalSessions;
      siteRow.deltaShare = scaledShare - (siteRow.currentShare || 0);
      var siteUrlsList = urlsBySite[siteRow.site] || [];
      siteUrlsList.forEach(function (row) {
        var normalizedShare = row.suggestedShare || 0;
        var newSuggestedSessions = normalizedShare * siteRow.suggestedSessions;
        var newGlobalShare = normalizedShare * siteRow.suggestedShare;
        row.suggestedSessions = newSuggestedSessions;
        row.deltaGlobalShare = newGlobalShare - (row.currentGlobalShare || 0);
        row.suggestedGlobalShare = newGlobalShare;
      });
    });
  }

  var siteOutput = siteBounds.map(function (item) {
    var row = siteRows.find(function (siteRow) { return siteRow.site === item.key; });
    if (!row) return null;
    return {
      site: row.site,
      rps: row.rps,
      ecpm: row.ecpm,
      coverage: row.coverage,
      ecpmEff: row.ecpmEff,
      momentum: row.momentum,
      confidence: row.confidence,
      score: row.score,
      currentShare: row.currentShare,
      suggestedShare: row.suggestedShare,
      deltaShare: row.deltaShare,
      mode: row.apostas ? 'Apostas' : 'Normal',
      nextEvaluation: row.nextEvaluation || '',
      lastChange: row.lastChange || '',
      currentSessions: row.currentSessions,
      suggestedSessions: row.suggestedSessions
    };
  }).filter(function (row) { return !!row; });

  siteOutput.sort(function (a, b) { return b.suggestedShare - a.suggestedShare; });
  if (!defaultSite && siteOutput.length) {
    defaultSite = siteOutput[0].site;
  }

  var siteContextMap = {};
  siteRows.forEach(function (row) {
    siteContextMap[row.site] = {
      ecpmEff: row.ecpmEff,
      coverage: row.coverage,
      momentum: row.momentum,
      mode: row.apostas ? 'Apostas' : 'Normal',
      score: row.score
    };
  });

  var sortedGlobal = globalAggregates.slice().sort(function (a, b) {
    return (b.suggestedGlobalShare || 0) - (a.suggestedGlobalShare || 0);
  });
  var priorityShareAccum = 0;
  var priorityActiveCount = 0;
  sortedGlobal.forEach(function (row, index) {
    var siteCtx = siteContextMap[row.site] || {};
    row.reason = describeUrlAllocationReason_(row, siteCtx, controls, index, sortedGlobal.length, limitApplied);
    if (row.priority && (row.suggestedGlobalShare || 0) > 0) {
      priorityActiveCount++;
      priorityShareAccum += row.suggestedGlobalShare || 0;
    }
  });
  var activeCount = sortedGlobal.filter(function (row) { return (row.suggestedGlobalShare || 0) > 0; }).length;
  var totalCount = sortedGlobal.length;
  var averageGlobalShare = activeCount ? 1 / activeCount : (totalCount ? 1 / totalCount : 0);
  var displayEntries = sortedGlobal.filter(function (row) {
    return (row.suggestedGlobalShare || 0) > 0;
  });
  if (!displayEntries.length) {
    displayEntries = sortedGlobal.slice();
  }
  var globalSummary = {
    entries: displayEntries,
    all: sortedGlobal,
    averageShare: averageGlobalShare,
    activeCount: activeCount,
    totalCount: totalCount,
    targetRecipients: requestedRecipients,
    limitApplied: limitApplied,
    priorityShareTarget: controls.urlPriorityShare || 0,
    priorityCount: Math.max(0, Math.round(controls.urlMinRecipients || 0)),
    priorityActive: priorityActiveCount,
    priorityShareActual: priorityShareAccum
  };

  return {
    controls: controls,
    sites: siteOutput,
    urlsBySite: urlsBySite,
    selectedSite: defaultSite,
    latestHour: hoursInfo && hoursInfo.latestHourLabel ? hoursInfo.latestHourLabel : '',
    totalSessions: controls.totalSessions,
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

function describeUrlAllocationReason_(row, siteCtx, controls, index, totalCount, limitApplied) {
  var share = toNumber_(row && row.suggestedGlobalShare != null ? row.suggestedGlobalShare : 0);
  var delta = toNumber_(row && row.deltaGlobalShare != null ? row.deltaGlobalShare : 0);
  var parts = [];

  if (share <= 0) {
    if ((row && row.limited) || limitApplied) {
      parts.push('Sem tráfego nesta rodada por causa do limite de URLs ativas.');
    } else {
      parts.push('Sem tráfego sugerido nesta rodada.');
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
      parts.push('No grupo prioritário (' + formatShareValue_(targetPriorityShare, 1) + ' do tráfego reservado).');
    } else {
      parts.push('No grupo prioritário definido manualmente.');
    }
  } else if (controls && controls.urlMinRecipients > 0 && controls.urlPriorityShare > 0 && !controls.urlRequireAll) {
    var residualShare = Math.max(0, 1 - controls.urlPriorityShare);
    if (residualShare > 0) {
      parts.push('Distribuição dentro do grupo complementar (' + formatShareValue_(residualShare, 1) + ' restante).');
    }
  }

  var scoreValue = toNumber_(row && row.score != null ? row.score : 0);
  var baseScore = toNumber_(row && row.baseScore != null ? row.baseScore : 0);
  var momentumFactor = toNumber_(row && row.momentumFactor != null ? row.momentumFactor : 1);
  var weightValue = clamp_(toNumber_(row && row.weight != null ? row.weight : 0), 0, 1);
  var complementWeight = clamp_(1 - weightValue, 0, 1);
  var urlScoreComponent = toNumber_(row && row.ecpmScore != null ? row.ecpmScore : 0);
  var siteScoreComponent = toNumber_(row && row.siteEcpmScore != null ? row.siteEcpmScore : 0);
  if ((row && row.score != null) || (row && row.baseScore != null)) {
    parts.push('Score ' + formatScoreValue_(scoreValue, 2) + ' (base ' + formatScoreValue_(baseScore, 2) + ' × momentum ' + formatScoreValue_(momentumFactor, 3) + ').');
    parts.push('Base = ' + formatPercentDisplay_(weightValue * 100, 1) + ' da URL (' + formatScoreValue_(urlScoreComponent, 2) + ') + ' + formatPercentDisplay_(complementWeight * 100, 1) + ' do site (' + formatScoreValue_(siteScoreComponent, 2) + ').');
  }
  if (row && row.ucb != null) {
    parts.push('UCB ' + formatScoreValue_(row.ucb, 2) + ' para exploração.');
  }

  if (row && row.performanceScore != null) {
    parts.push('Desempenho composto ' + formatShareValue_(row.performanceScore, 1) + ' (score ' + formatShareValue_(row.performanceNormalizedScore || 0, 1) + ', eCPM ef. ' + formatShareValue_(row.performanceNormalizedEcpm || 0, 1) + ', RPS ' + formatShareValue_(row.performanceNormalizedRps || 0, 1) + ').');
    parts.push('Métricas base: score ' + formatScoreValue_(row.score, 2) + ', eCPM efetivo ' + formatCurrencyValue_(row.ecpmEff, 2) + ', RPS ' + formatCurrencyValue_(row.rps, 2) + '.');
  }

  var rankText = 'Prioridade global #' + (index + 1) + ' de ' + totalCount + '.';
  parts.push(rankText);

  if (siteCtx && siteCtx.mode === 'Apostas') {
    parts.push('Site em modo Apostas (exploração controlada).');
  }

  var ecpmEff = toNumber_(row && row.ecpmEff != null ? row.ecpmEff : 0);
  if (ecpmEff > 0) {
    var siteEcpmEff = toNumber_(siteCtx && siteCtx.ecpmEff != null ? siteCtx.ecpmEff : 0);
    var ecpmText = formatCurrencyValue_(ecpmEff, 2);
    if (siteEcpmEff > 0) {
      var diffPct = ((ecpmEff - siteEcpmEff) / siteEcpmEff) * 100;
      if (diffPct > 5) {
        parts.push('eCPM efetivo ' + ecpmText + ' (' + formatPercentDisplay_(diffPct, 0) + ' acima da média do site).');
      } else if (diffPct < -5) {
        parts.push('eCPM efetivo ' + ecpmText + ' (' + formatPercentDisplay_(Math.abs(diffPct), 0) + ' abaixo da média do site).');
      } else {
        parts.push('eCPM efetivo ' + ecpmText + ' alinhado com o site.');
      }
    } else {
      parts.push('eCPM efetivo ' + ecpmText + '.');
    }
  }

  var coverage = toNumber_(row && row.coverage != null ? row.coverage : 0);
  if (coverage > 0) {
    var target = toNumber_(controls && controls.coverageTarget != null ? controls.coverageTarget : 0) * 100;
    if (target && coverage >= target) {
      parts.push('Cobertura ' + formatPercentDisplay_(coverage, 1) + ' acima da meta.');
    } else if (coverage >= 60) {
      parts.push('Cobertura ' + formatPercentDisplay_(coverage, 1) + ' levemente abaixo da meta.');
    } else {
      parts.push('Cobertura baixa ' + formatPercentDisplay_(coverage, 1) + ', monitorar fill.');
    }
  }

  var momentum = toNumber_(row && row.momentum != null ? row.momentum : 0);
  if (momentum >= 0.1) {
    parts.push('Momentum positivo ' + formatPercentDisplay_(momentum * 100, 0) + ' vs média das 2h anteriores.');
  } else if (momentum <= -0.1) {
    parts.push('Momentum negativo ' + formatPercentDisplay_(Math.abs(momentum) * 100, 0) + ' vs média das 2h anteriores.');
  }

  var rounds = Math.max(0, Math.round(toNumber_(row && row.rounds != null ? row.rounds : 0)));
  if (rounds > 0) {
    parts.push('Já avaliada em ' + rounds + ' ' + (rounds === 1 ? 'rodada' : 'rodadas') + '.');
  } else {
    parts.push('Em fase de seed para ganhar sinal.');
  }

  if (row && row.guaranteed) {
    parts.push('Mantém tráfego mínimo garantido (' + (row.status || 'controle') + ').');
  } else if (row && row.status) {
    parts.push('Status: ' + row.status + '.');
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
      companyId: row.empresaid,
      domainId: row.dominioid,
      routeSlug: row.slugrota,
      active: coerceBoolean_(rawActive)
    };
  });
  const operationIntegrationsByOperation = {};
  operationIntegrationRows.forEach(function (row) {
    if (!row.operation) return;
    operationIntegrationsByOperation[row.operation] = row;
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
  var headers = ['ID', 'Operação', 'Empresa ID', 'Domínio ID', 'Slug Rota', 'Ativo?'];
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
  if (!sheet) {
    sheet = ss.insertSheet(DISTRIBUTION_CONFIG_SHEET);
  }
  var defaults = getDefaultDistributionControls_();
  var headers = ['Operação', 'Tráfego'];
  Object.keys(defaults).forEach(function (key) {
    headers.push(key);
  });
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  return sheet;
}

function getDistributionControlsConfig_(operation, traffic) {
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
  if (opIndex == null || trafficIndex == null) {
    return {};
  }
  var targetRow = null;
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (!row || row.join('').trim() === '') continue;
    var rowOp = String(row[opIndex] || '').trim();
    var rowTraffic = String(row[trafficIndex] || '').trim();
    if (rowOp && rowTraffic && rowOp.toLowerCase() === String(operation).trim().toLowerCase() && rowTraffic.toLowerCase() === String(traffic).trim().toLowerCase()) {
      targetRow = row;
      break;
    }
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

function saveDistributionControlsConfig_(operation, traffic, controls) {
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
  if (opIndex == null || trafficIndex == null) {
    return;
  }
  var lastRow = sheet.getLastRow();
  var rangeValues = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  var rowNumber = -1;
  for (var i = 0; i < rangeValues.length; i++) {
    var row = rangeValues[i];
    var rowOp = String(row[opIndex] || '').trim();
    var rowTraffic = String(row[trafficIndex] || '').trim();
    if (rowOp && rowTraffic && rowOp.toLowerCase() === String(operation).trim().toLowerCase() && rowTraffic.toLowerCase() === String(traffic).trim().toLowerCase()) {
      rowNumber = i + 2;
      break;
    }
  }
  var rowValues = new Array(headers.length);
  for (var h = 0; h < headers.length; h++) {
    rowValues[h] = '';
  }
  rowValues[opIndex] = String(operation).trim();
  rowValues[trafficIndex] = String(traffic).trim();
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

function computeCoverageFactor_(ratio, controls) {
  controls = controls || getDefaultDistributionControls_();
  var target = controls.coverageTarget != null ? controls.coverageTarget : 0.75;
  var penaltyFloor = controls.coveragePenaltyFloor != null ? controls.coveragePenaltyFloor : 0.5;
  var cap = controls.coverageCap != null ? controls.coverageCap : 0.95;
  var bonus = controls.coverageBonusFactor != null ? controls.coverageBonusFactor : 0.2;
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
  if (!operation) {
    throw new Error('Selecione a operação.');
  }
  if (!traffic) {
    throw new Error('Selecione o tráfego.');
  }
  var params = {
    operation: operation,
    traffic: traffic,
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
    return row && row.active && row.operation && row.companyId && row.domainId && row.routeSlug;
  });
  var trafficByType = snapshot.trafficByType || {};
  var successes = [];
  var failures = [];

  integrations.forEach(function (integration) {
    TRAFFIC_TYPES.forEach(function (trafficType) {
      var sources = trafficByType[trafficType] || [];
      if (!sources.length) {
        return;
      }
      try {
        var result = syncLinkRouterDistribution({
          operation: integration.operation,
          traffic: trafficType,
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
    return row && row.active && row.operation && row.companyId && row.domainId && row.routeSlug;
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
      .everyHours(4)
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
