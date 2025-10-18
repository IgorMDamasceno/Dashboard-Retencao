const DATA_SHEET = 'BD – GAM';
const OPERATIONS_SHEET = 'Config - Operações';
const TRAFFIC_SHEET = 'Config - Trafego';
const REVSHARE_SHEET = 'Config - RevShare';
const TRAFFIC_TYPES = ['Automação', 'Broadcast', 'Push'];
const DISTRIBUTION_STATE_SHEET = 'Controle - Distribuição';

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
  const distribution = buildDistributionPlan_(
    filteredRows,
    hoursInfo,
    snapshot.revshareMap,
    params.distribution
  );
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
    distribution: distribution
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
    distribution: buildEmptyDistribution_()
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

function buildEmptyDistribution_() {
  var controls = getDefaultDistributionControls_();
  return {
    controls: controls,
    sites: [],
    urlsBySite: {},
    selectedSite: '',
    latestHour: '',
    totalSessions: controls.totalSessions,
    hours: []
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
      controls[key] = Math.max(1000, Math.round(parsed));
    } else if (key === 'tau' || key === 'tauLocal') {
      controls[key] = Math.max(1, parsed);
    } else if (key === 'reliabilityK' || key === 'siteMinSessions' || key === 'urlSeedSessions' || key === 'urlMinRecipients') {
      controls[key] = Math.max(0, Math.round(parsed));
    } else if (key === 'ucbZ') {
      controls[key] = Math.max(0, parsed);
    } else if (key === 'urlRequireAll') {
      controls[key] = parsed ? 1 : 0;
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
      globalUrls: { more: [], less: [], all: [], averageShare: 0 }
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
      var rpsEff = urlRps * coverageFactor;
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
        rpsEff: rpsEff,
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
    var rpsEffSite = rps * siteCoverageFactor;
    var siteMomentum = computeMomentum_(hourRps);
    var siteConfidence = sessionsWindow / (sessionsWindow + controls.reliabilityK);
    var siteScore = rpsEffSite;
    if (siteCoverageRatio >= controls.coverageTarget) {
      siteScore *= 1.02;
    }

    var roundsTotal = 0;
    var roundsCount = 0;
    urlEntries.forEach(function (entry) {
      var scoreBase = entry.weight * entry.rpsEff + (1 - entry.weight) * rpsEffSite;
      var momentumFactor = 1 + 0.15 * clamp_(entry.momentum, -0.5, 0.5);
      entry.score = scoreBase * momentumFactor;
      roundsTotal += entry.rounds;
      roundsCount++;
    });
    var R = Math.max(roundsTotal, roundsCount, 1);
    var lnR = Math.log(R + 1);
    urlEntries.forEach(function (entry) {
      entry.ucbScore = entry.score + controls.ucbZ * Math.sqrt(lnR / (entry.rounds + 1));
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
      coverage: siteCoverage,
      coverageRatio: siteCoverageRatio,
      rpsEff: rpsEffSite,
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
      globalUrls: { more: [], less: [], all: [], averageShare: 0 }
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
  var minSiteShare = Math.max(controls.siteMinShare, controls.siteMinSessions / controls.totalSessions);
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
    var exploit = siteUrls.slice().sort(function (a, b) { return b.score - a.score; }).slice(0, Math.min(3, siteUrls.length));
    var explore = siteRow.apostas ? promising : [];
    if (siteRow.apostas && !explore.length) {
      explore = siteUrls.slice().sort(function (a, b) { return a.rounds - b.rounds; }).slice(0, Math.min(3, siteUrls.length));
    }

    var baseShareMap = {};
    if (!siteUrls.length) {
      baseShareMap = {};
    } else if (siteRow.apostas) {
      var exploitMap = computeSoftmaxShares_(exploit, controls.tauLocal, 'ucbScore');
      var exploreMap = explore.length ? computeSoftmaxShares_(explore, controls.tauLocal, 'ucbScore') : {};
      siteUrls.forEach(function (entry) { baseShareMap[entry.key] = 0; });
      exploit.forEach(function (entry) {
        baseShareMap[entry.key] += controls.exploitPortion * (exploitMap[entry.key] || 0);
      });
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
        baseShareMap = computeSoftmaxShares_(siteUrls, controls.tauLocal, 'ucbScore');
      }
    } else {
      baseShareMap = computeSoftmaxShares_(siteUrls, controls.tauLocal, 'ucbScore');
    }

    var requiredKeys = {};
    var requiredCount = controls.urlRequireAll ? siteUrls.length : Math.min(siteUrls.length, Math.max(0, controls.urlMinRecipients));
    if (controls.urlRequireAll && siteUrls.length) {
      requiredCount = siteUrls.length;
    }
    if (requiredCount > 0 && siteUrls.length) {
      var priority = siteUrls.slice().sort(function (a, b) {
        var aBase = baseShareMap[a.key] != null ? baseShareMap[a.key] : 0;
        var bBase = baseShareMap[b.key] != null ? baseShareMap[b.key] : 0;
        if (bBase === aBase) {
          return b.score - a.score;
        }
        return bBase - aBase;
      });
      for (var r = 0; r < requiredCount && r < priority.length; r++) {
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
        required: isRequired
      };
    });

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
        coverage: entry.coverage,
        rpsEff: entry.rpsEff,
        momentum: entry.momentum,
        rounds: entry.rounds,
        score: entry.score,
        ucb: entry.ucbScore,
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
        guaranteed: !!item.required
      };
    });

    urlsBySite[siteRow.site] = siteUrlRows;
    Array.prototype.push.apply(globalAggregates, siteUrlRows);
  });

  var siteOutput = siteBounds.map(function (item) {
    var row = siteRows.find(function (siteRow) { return siteRow.site === item.key; });
    if (!row) return null;
    return {
      site: row.site,
      rps: row.rps,
      coverage: row.coverage,
      rpsEff: row.rpsEff,
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

  var sortedGlobal = globalAggregates.slice().sort(function (a, b) {
    return (b.suggestedGlobalShare || 0) - (a.suggestedGlobalShare || 0);
  });
  var averageGlobalShare = sortedGlobal.length ? 1 / sortedGlobal.length : 0;
  var requiredTop = controls.urlRequireAll ? sortedGlobal.length : Math.min(sortedGlobal.length, Math.max(0, controls.urlMinRecipients));
  var globalMore = [];
  var globalLess = [];
  sortedGlobal.forEach(function (row, index) {
    var qualifies = index < requiredTop || (row.suggestedGlobalShare || 0) >= averageGlobalShare || row.guaranteed;
    if (qualifies) {
      globalMore.push(row);
    } else {
      globalLess.push(row);
    }
  });
  if (!globalMore.length && sortedGlobal.length) {
    globalMore.push(sortedGlobal[0]);
  }
  if (!globalLess.length && globalMore.length < sortedGlobal.length) {
    sortedGlobal.forEach(function (row) {
      if (globalMore.indexOf(row) === -1) {
        globalLess.push(row);
      }
    });
  }
  var globalSummary = {
    more: globalMore,
    less: globalLess,
    all: sortedGlobal,
    averageShare: averageGlobalShare
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
