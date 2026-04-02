// Utilities for Pitching/Campaigns sync + archive + Median Views enrichment.


// 1) Push approved -> Campaigns
function pushConfirmedCreatorsToCampaigns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pitchingSheet = ss.getSheetByName("Pitching");
  const campaignsSheet = ss.getSheetByName("Campaigns");
  if (!pitchingSheet || !campaignsSheet) return Logger.log("❌ Missing 'Pitching' or 'Campaigns' sheet.");

  const pitchData = pitchingSheet.getDataRange().getValues();
  const campData = campaignsSheet.getDataRange().getValues();
  if (pitchData.length < 2 || campData.length < 1) return Logger.log("ℹ️ One of the sheets has no data.");

  const pitchHeader = (pitchData[0] || []).map(v => String(v || "").trim());
  const campHeader = (campData[0] || []).map(v => String(v || "").trim());
  const campNumCols = campHeader.length;

  const pStatusCol = pitchHeader.indexOf("Status");
  const pAvailabilityCol = pitchHeader.indexOf("Availability");
  const pChannelCol = pitchHeader.indexOf("Channel Name");
  const cChannelCol = campHeader.indexOf("Channel Name");
  if (pStatusCol === -1 || pAvailabilityCol === -1 || pChannelCol === -1) {
    return Logger.log("❌ Pitching missing required columns: Status, Availability, Channel Name.");
  }
  if (cChannelCol === -1) return Logger.log("❌ Campaigns missing required column: Channel Name.");

  const archivedStart0 = findSectionRowByLabel_(pitchData, "Archived");
  if (archivedStart0 === -1) return Logger.log("❌ Pitching section 'Archived' not found.");

  const commonCols = getCommonColumnsByHeader_(pitchHeader, campHeader);
  if (commonCols.length === 0) return Logger.log("ℹ️ No matching columns between Pitching and Campaigns.");
  const pitchRowsByKey = getPitchRowsByKey_(pitchData, pitchHeader, archivedStart0);

  const existingBySection = getExistingSignaturesBySection_(campData, commonCols);
  const candidates = [];
  for (let r = 1; r < archivedStart0; r++) {
    const row = pitchData[r];
    if (isBlankRow_(row) || isSectionLabelRow_(row)) continue;
    if (String(row[pStatusCol] || "").trim() !== "Approved") continue;

    const availability = String(row[pAvailabilityCol] || "").trim();
    const channelName = String(row[pChannelCol] || "").trim();
    if (!channelName) continue;
    if (!availability) {
      Logger.log(`⚠️ Skipping "${channelName}" (row ${r + 1}): Availability is empty.`);
      continue;
    }

    candidates.push({
      pitchRow1: r + 1,
      values: row,
      availability,
      channelName
    });
  }

  if (candidates.length === 0) return Logger.log("ℹ️ No approved creators to push above Archived.");

  let inserted = 0;
  let skippedDup = 0;
  let skippedNoSection = 0;

  for (const item of candidates) {
    const signature = buildSignatureFromPitchRow_(item.values, pitchHeader, commonCols);
    if (!existingBySection[item.availability]) existingBySection[item.availability] = new Set();
    if (existingBySection[item.availability].has(signature)) {
      skippedDup++;
      continue;
    }

    const campLive = campaignsSheet.getDataRange().getValues();
    const sectionStart0 = findSectionRowByLabel_(campLive, item.availability);
    if (sectionStart0 === -1) {
      skippedNoSection++;
      Logger.log(`⚠️ Campaigns section "${item.availability}" not found. Skipping "${item.channelName}".`);
      continue;
    }

    const firstEmptyRow1 = findFirstEmptyRowInSection1_(campLive, sectionStart0);
    if (firstEmptyRow1 === -1) {
      Logger.log(`⚠️ No empty row found in section "${item.availability}" for "${item.channelName}".`);
      continue;
    }

    campaignsSheet.insertRowBefore(firstEmptyRow1);
    const targetRow1 = firstEmptyRow1;
    const formatSourceRow1 = firstEmptyRow1 + 1; // original empty row shifted down

    const out = new Array(campNumCols).fill("");
    for (const m of commonCols) out[m.cIdx] = getPitchValueForCampaignColumn_(item.values, pitchHeader, m);
    applyPitchSpecialCampaignMappings_(out, campHeader, item.values, pitchHeader, pitchRowsByKey);

    const targetRange = campaignsSheet.getRange(targetRow1, 1, 1, campNumCols);
    const templateRange = getSafeFormatRow_(campaignsSheet, formatSourceRow1, campNumCols);
    templateRange.copyTo(targetRange, { formatOnly: true });
    targetRange.setValues([out]);

    existingBySection[item.availability].add(signature);
    inserted++;
  }

  Logger.log(`✅ Push done. Inserted: ${inserted}. Skipped duplicates: ${skippedDup}. Missing section: ${skippedNoSection}.`);
}


// 2) Archive pitches (Status = Approved or Rejected)
function archivePitches() {
  try {
    Logger.log("▶️ Running Pitching → Campaigns before archive.");
    pushConfirmedCreatorsToCampaigns();
    SpreadsheetApp.flush();
  } catch (e) {
    Logger.log(`⚠️ Pitching → Campaigns step failed; continuing archive: ${e && e.stack ? e.stack : e}`);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Pitching");
  if (!sheet) return Logger.log("❌ 'Pitching' sheet not found.");

  const data = sheet.getDataRange().getValues();
  const header = (data[0] || []).map(v => String(v || "").trim());

  const statusCol = header.indexOf("Status");
  if (statusCol === -1) return Logger.log("❌ Pitching missing required column: Status.");

  const activeStart0 = findSectionRowByLabel_(data, "Active Pitches");
  const archivedStart0 = findSectionRowByLabel_(data, "Archived");
  if (activeStart0 === -1) return Logger.log("❌ Could not find 'Active Pitches' section in Pitching.");
  if (archivedStart0 === -1) return Logger.log("❌ Could not find 'Archived' section in Pitching.");

  const rowsToArchive = [];
  for (let r = activeStart0 + 1; r < archivedStart0; r++) {
    const row = data[r];
    if (isBlankRow_(row) || isSectionLabelRow_(row)) continue;
    const status = String(row[statusCol] || "").trim();
    if (status !== "Approved" && status !== "Rejected") continue;
    rowsToArchive.push({ sourceRow1: r + 1, values: row.slice() });
  }

  if (rowsToArchive.length === 0) return Logger.log("ℹ️ No rows to archive.");

  // Delete from bottom up to preserve row numbers
  for (const entry of rowsToArchive.slice().sort((a, b) => b.sourceRow1 - a.sourceRow1)) {
    sheet.deleteRow(entry.sourceRow1);
  }

  // Re-read after deletions
  const refreshed = sheet.getDataRange().getValues();
  const refreshedActiveStart0 = findSectionRowByLabel_(refreshed, "Active Pitches");
  const refreshedArchivedStart0 = findSectionRowByLabel_(refreshed, "Archived");
  if (refreshedArchivedStart0 === -1) return Logger.log("❌ Archived section disappeared after deletion.");

  // Restore blank rows at the top of Active Pitches (right below the section label)
  if (refreshedActiveStart0 !== -1) {
    const restoreAt1 = refreshedActiveStart0 + 2; // 1-indexed row right below "Active Pitches" label
    if (rowsToArchive.length > 1) {
      sheet.insertRowsBefore(restoreAt1, rowsToArchive.length);
    } else {
      sheet.insertRowBefore(restoreAt1);
    }
    // Format/validation source: the original first data row, now shifted down by the inserted rows
    const formatSourceRow1 = restoreAt1 + rowsToArchive.length;
    const numColsActive = sheet.getLastColumn();
    const templateActiveRange = getSafeFormatRow_(sheet, formatSourceRow1, numColsActive);
    const restoreRange = sheet.getRange(restoreAt1, 1, rowsToArchive.length, numColsActive);
    templateActiveRange.copyTo(restoreRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    templateActiveRange.copyTo(restoreRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    sheet.setRowHeights(restoreAt1, rowsToArchive.length, sheet.getRowHeight(formatSourceRow1));
  }

  // Re-read again after blank rows were inserted
  const refreshed2 = sheet.getDataRange().getValues();
  const archivedStart0final = findSectionRowByLabel_(refreshed2, "Archived");
  if (archivedStart0final === -1) return Logger.log("❌ Archived section not found after row restore.");

  let insertAt1 = findFirstEmptyRowInSection1_(refreshed2, archivedStart0final);
  if (insertAt1 === -1) insertAt1 = archivedStart0final + 2;

  const numCols = sheet.getLastColumn();
  const arrayFormulaCols0 = getArrayFormulaColumns_(sheet, 2, numCols);
  const templateRange = getSafeFormatRow_(sheet, archivedStart0final + 2, numCols);

  if (rowsToArchive.length > 1) {
    sheet.insertRowsBefore(insertAt1, rowsToArchive.length);
  } else {
    sheet.insertRowBefore(insertAt1);
  }

  const writeRange = sheet.getRange(insertAt1, 1, rowsToArchive.length, numCols);
  templateRange.copyTo(writeRange, { formatOnly: true });
  writeRange.setValues(rowsToArchive.map(r => r.values));
  for (const col0 of arrayFormulaCols0) {
    sheet.getRange(insertAt1, col0 + 1, rowsToArchive.length, 1).clearContent();
  }

  Logger.log(`✅ Archived ${rowsToArchive.length} row(s).`);
}

// 4) Push published Campaigns rows -> Performance
function pushPublishedCampaignsToPerformance() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const campaignsSheet = ss.getSheetByName("Campaigns");
  const performanceSheet = ss.getSheetByName("Performance");
  const pitchingSheet = ss.getSheetByName("Pitching");
  if (!campaignsSheet || !performanceSheet) {
    return Logger.log("❌ Missing 'Campaigns' or 'Performance' sheet.");
  }

  const campaignsData = campaignsSheet.getDataRange().getValues();
  const performanceData = performanceSheet.getDataRange().getValues();
  if (campaignsData.length < 2 || performanceData.length < 1) {
    return Logger.log("ℹ️ One of the sheets has no data.");
  }

  const campaignsHeader = (campaignsData[0] || []).map(v => String(v || "").trim());
  const performanceHeader = (performanceData[0] || []).map(v => String(v || "").trim());
  const statusCol = campaignsHeader.indexOf("Status");
  if (statusCol === -1) return Logger.log("❌ Campaigns missing required column: Status.");

  const commonCols = getCommonColumnsByHeader_(campaignsHeader, performanceHeader);
  if (commonCols.length === 0) return Logger.log("ℹ️ No matching columns between Campaigns and Performance.");

  // Build Pitching lookup by composite key (HubSpot Record ID + Deal Type + Activation Type)
  let pitchRowsByKey = new Map();
  let pitchHeader = [];
  if (pitchingSheet) {
    const pitchData = pitchingSheet.getDataRange().getValues();
    if (pitchData.length >= 2) {
      pitchHeader = (pitchData[0] || []).map(v => String(v || "").trim());
      const archivedStart0 = findSectionRowByLabel_(pitchData, "Archived");
      const endRow = archivedStart0 === -1 ? pitchData.length : archivedStart0;
      pitchRowsByKey = getPitchRowsByKey_(pitchData, pitchHeader, endRow);
    }
  }

  const rowsToAppend = [];
  for (let r = 1; r < campaignsData.length; r++) {
    const row = campaignsData[r];
    if (isBlankRow_(row) || isSectionLabelRow_(row)) continue;
    if (String(row[statusCol] || "").trim() !== "Published") continue;

    const out = new Array(performanceHeader.length).fill("");
    for (const m of commonCols) out[m.cIdx] = row[m.pIdx];

    // Look up matching Pitching row and copy Median Views -> Expected Views, CPM -> Expected CPM
    if (pitchRowsByKey.size > 0) {
      const key = buildPitchCompositeKey_(row, campaignsHeader);
      const pitchRow = key ? pitchRowsByKey.get(key) : null;
      if (pitchRow) {
        setValueIfTargetColumnExists_(out, performanceHeader, "Expected Views", getPitchValueByHeader_(pitchRow, pitchHeader, "Median Views"));
        setValueIfTargetColumnExists_(out, performanceHeader, "Expected CPM", getPitchValueByHeader_(pitchRow, pitchHeader, "CPM"));
      }
    }

    rowsToAppend.push(out);
  }

  if (rowsToAppend.length === 0) return Logger.log("ℹ️ No published Campaigns rows to push.");

  const startRow1 = findFirstFreeRowFrom1_(performanceSheet, 3, performanceHeader.length);
  const writeRange = performanceSheet.getRange(startRow1, 1, rowsToAppend.length, performanceHeader.length);
  writeRange.setValues(rowsToAppend);

  Logger.log(`✅ Pushed ${rowsToAppend.length} published row(s) to Performance starting at row ${startRow1}.`);
}


// Menu
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🧰 Scripts")
    .addItem("HubSpot → Pitching", "importHubSpotDealsListToPitching")
    .addSeparator()
    .addItem("Pitching → Campaigns", "pushConfirmedCreatorsToCampaigns")
    .addItem("Archive pitches", "archivePitches")
    .addSeparator()
    .addItem("Campaigns → Performance", "pushPublishedCampaignsToPerformance")
    .addToUi();
}


// Shared helper
function getSafeFormatRow_(sheet, preferredRow, numCols) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const row = (preferredRow >= 1 && preferredRow <= lastRow) ? preferredRow : lastRow;
  return sheet.getRange(row, 1, 1, numCols);
}

function asBool_(value) {
  return value === true || String(value).toLowerCase() === "true";
}

function isBlankRow_(row) {
  return !row || row.join("").trim() === "";
}

function isSectionLabelRow_(row) {
  if (!row) return false;
  const first = String(row[0] || "").trim();
  if (!first) return false;
  for (let i = 1; i < row.length; i++) {
    if (String(row[i] || "").trim() !== "") return false;
  }
  return true;
}

function findSectionRowByLabel_(data, label) {
  const target = String(label || "").trim();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0] || "").trim() === target) return i;
  }
  return -1;
}

function findNextSectionStart0_(data, sectionStart0) {
  for (let r = sectionStart0 + 1; r < data.length; r++) {
    if (isSectionLabelRow_(data[r])) return r;
  }
  return -1;
}

function findFirstEmptyRowInSection1_(data, sectionStart0) {
  const nextSection0 = findNextSectionStart0_(data, sectionStart0);
  const end0 = nextSection0 === -1 ? data.length : nextSection0;
  for (let r = sectionStart0 + 1; r < end0; r++) {
    if (isBlankRow_(data[r])) return r + 1;
  }
  return -1;
}

function findFirstFreeRowFrom1_(sheet, startRow1, numCols) {
  const firstRow1 = Math.max(Number(startRow1) || 1, 1);
  const width = Math.max(Number(numCols) || 1, 1);
  const lastRow = Math.max(sheet.getLastRow(), firstRow1);

  if (lastRow < firstRow1) return firstRow1;

  const values = sheet.getRange(firstRow1, 1, lastRow - firstRow1 + 1, width).getValues();
  for (let i = 0; i < values.length; i++) {
    if (isBlankRow_(values[i])) return firstRow1 + i;
  }

  return lastRow + 1;
}

function getCommonColumnsByHeader_(sourceHeader, targetHeader) {
  const sourceMap = new Map();
  for (let i = 0; i < sourceHeader.length; i++) {
    const name = String(sourceHeader[i] || "").trim();
    if (name) sourceMap.set(name, i);
  }

  const common = [];
  for (let j = 0; j < targetHeader.length; j++) {
    const name = String(targetHeader[j] || "").trim();
    if (!name) continue;
    if (sourceMap.has(name)) common.push({ name, pIdx: sourceMap.get(name), cIdx: j });
  }
  return common;
}

function getArrayFormulaColumns_(sheet, probeRow1, numCols) {
  if (!sheet || probeRow1 < 1 || numCols < 1) return [];
  if (sheet.getMaxRows() < probeRow1) return [];

  const formulas = sheet.getRange(probeRow1, 1, 1, numCols).getFormulas()[0] || [];
  const cols = [];

  for (let c = 0; c < formulas.length; c++) {
    if (/ARRAYFORMULA\s*\(/i.test(String(formulas[c] || ""))) cols.push(c);
  }
  return cols;
}

function getExistingSignaturesBySection_(campaignData, commonCols) {
  const out = {};
  let currentSection = "";

  for (let r = 1; r < campaignData.length; r++) {
    const row = campaignData[r];
    if (isSectionLabelRow_(row)) {
      currentSection = String(row[0] || "").trim();
      if (!out[currentSection]) out[currentSection] = new Set();
      continue;
    }
    if (!currentSection || isBlankRow_(row)) continue;

    const signature = buildSignatureFromCampaignRow_(row, commonCols);
    if (signature) out[currentSection].add(signature);
  }

  return out;
}

function buildSignatureFromPitchRow_(row, pitchHeader, commonCols) {
  return commonCols
    .map(m => String(getPitchValueForCampaignColumn_(row, pitchHeader, m) || "").trim())
    .join("||");
}

function buildSignatureFromCampaignRow_(row, commonCols) {
  return commonCols.map(m => String(row[m.cIdx] || "").trim()).join("||");
}

function getPitchValueForCampaignColumn_(row, pitchHeader, mapping) {
  if (!mapping || !mapping.name) return "";
  if (mapping.name === "Rate") return getNegotiatedRateValue_(row, pitchHeader);
  if (mapping.name === "Status") return "Active";
  return row[mapping.pIdx];
}

function getNegotiatedRateValue_(row, pitchHeader) {
  const rateStart0 = pitchHeader.indexOf("Rate");
  if (rateStart0 === -1) return "";

  let rateEnd0 = pitchHeader.length - 1;
  for (let c = rateStart0 + 1; c < pitchHeader.length; c++) {
    if (String(pitchHeader[c] || "").trim() !== "") {
      rateEnd0 = c - 1;
      break;
    }
  }

  for (let c = rateEnd0; c >= rateStart0; c--) {
    const value = row[c];
    if (String(value || "").trim() !== "") return value;
  }

  return "";
}

function getPitchValueByHeader_(row, pitchHeader, headerName) {
  const idx = pitchHeader.indexOf(String(headerName || "").trim());
  return idx === -1 ? "" : row[idx];
}

function applyPitchSpecialCampaignMappings_(outRow, campHeader, pitchRow, pitchHeader, pitchRowsByKey) {
  const matchedPitchRow =
    getPitchRowByCompositeKey_(pitchRow, pitchHeader, pitchRowsByKey) || pitchRow;

  setValueIfTargetColumnExists_(outRow, campHeader, "Expected Views", getPitchValueByHeader_(matchedPitchRow, pitchHeader, "Median Views"));
  setValueIfTargetColumnExists_(outRow, campHeader, "Expected CPM", getPitchValueByHeader_(matchedPitchRow, pitchHeader, "CPM"));
}

function setValueIfTargetColumnExists_(row, header, columnName, value) {
  const idx = header.indexOf(String(columnName || "").trim());
  if (idx !== -1) row[idx] = value;
}

function getPitchRowsByKey_(pitchData, pitchHeader, archivedStart0) {
  const out = new Map();
  for (let r = 1; r < archivedStart0; r++) {
    const row = pitchData[r];
    if (isBlankRow_(row) || isSectionLabelRow_(row)) continue;

    const key = buildPitchCompositeKey_(row, pitchHeader);
    if (!key || out.has(key)) continue;
    out.set(key, row);
  }
  return out;
}

function getPitchRowByCompositeKey_(row, pitchHeader, pitchRowsByKey) {
  const key = buildPitchCompositeKey_(row, pitchHeader);
  if (!key || !pitchRowsByKey) return null;
  return pitchRowsByKey.get(key) || null;
}

function buildPitchCompositeKey_(row, pitchHeader) {
  const keys = [
    getPitchValueByHeader_(row, pitchHeader, "HubSpot Record ID"),
    getPitchValueByHeader_(row, pitchHeader, "Deal Type"),
    getPitchValueByHeader_(row, pitchHeader, "Activation Type")
  ].map(v => String(v || "").trim());

  if (!keys[0]) return "";
  return keys.join("||");
}


const YT_KEY_PROP_ = "YOUTUBE_API_KEY";
const YT_AVG_CACHE_PREFIX_ = "AVG_VIEWS::";
const YT_KEY_VALID_CACHE_PREFIX_ = "YT_API_KEY_VALID::";
const YT_CACHE_TTL_SECONDS_ = 21600;
const YT_KEY_VALID_TTL_SECONDS_ = 1800;
const YT_KEY_INVALID_TTL_SECONDS_ = 600;
const YT_MAX_PLAYLIST_ITEMS_ = 50;
const YT_MAX_VIDEOS_FOR_AVG_ = 15;
const YT_MIN_DURATION_SECONDS_ = 180;
const YT_MONTHS_BACK_ = 6;

const YT_SKIP_PATHS_ = ["watch", "shorts", "playlist", "results", "feed"];
const YT_CHANNEL_ID_PATTERNS_ = [
  /"channelId":"(UC[a-zA-Z0-9_-]{22})"/,
  /"externalId":"(UC[a-zA-Z0-9_-]{22})"/,
  /"browseId":"(UC[a-zA-Z0-9_-]{22})"/,
  /<meta\s+itemprop="channelId"\s+content="(UC[a-zA-Z0-9_-]{22})"/i,
  /youtube\.com\/channel\/(UC[a-zA-Z0-9_-]{22})/
];

var YT_API_KEY_INVALID_ = false;
var YT_API_KEY_INVALID_LOGGED_ = false;

function computeMedianViewsForChannel_(channelUrl) {
  if (!channelUrl || typeof channelUrl !== "string") return null;

  const apiKey = getYouTubeApiKey_();
  if (!apiKey) {
    Logger.log("❌ Missing YouTube API key. Set Script Property: YOUTUBE_API_KEY");
    return null;
  }
  if (!validateYouTubeApiKey_(apiKey)) return null;

  const normalized = normalizeChannelUrl_(channelUrl);
  const cacheKey = YT_AVG_CACHE_PREFIX_ + normalized;
  const cache = CacheService.getDocumentCache();
  const props = PropertiesService.getDocumentProperties();

  const shortCached = cache.get(cacheKey);
  if (shortCached != null) return Number(shortCached);

  const persisted = props.getProperty(cacheKey);
  if (persisted != null) {
    cache.put(cacheKey, persisted, YT_CACHE_TTL_SECONDS_);
    return Number(persisted);
  }

  const median = computeMedianViewsFromYouTube_(normalized, apiKey);
  if (median == null) return null;

  const stored = String(median);
  props.setProperty(cacheKey, stored);
  cache.put(cacheKey, stored, YT_CACHE_TTL_SECONDS_);
  return median;
}

function normalizeChannelUrl_(url) {
  return String(url).trim().replace(/\s+/g, "");
}

function getYouTubeApiKey_() {
  const value = PropertiesService.getScriptProperties().getProperty(YT_KEY_PROP_);
  return String(value || "").trim();
}

function validateYouTubeApiKey_(apiKey) {
  if (!apiKey || YT_API_KEY_INVALID_) return false;

  const cache = CacheService.getDocumentCache();
  const cacheKey = YT_KEY_VALID_CACHE_PREFIX_ + apiKey.slice(-8);
  const cached = cache.get(cacheKey);
  if (cached === "1") return true;
  if (cached === "0") return false;

  try {
    const pingUrl =
      "https://www.googleapis.com/youtube/v3/search?part=id&type=channel&maxResults=1&q=youtube" +
      "&key=" + encodeURIComponent(apiKey);
    const resp = UrlFetchApp.fetch(pingUrl, { muteHttpExceptions: true });
    const code = resp.getResponseCode();
    const text = String(resp.getContentText() || "");

    if (code >= 400) {
      if (isInvalidApiKeyResponse_(text)) {
        markApiKeyInvalid_();
      } else {
        Logger.log("❌ YouTube API key validation failed: HTTP " + code);
      }
      cache.put(cacheKey, "0", YT_KEY_INVALID_TTL_SECONDS_);
      return false;
    }

    cache.put(cacheKey, "1", YT_KEY_VALID_TTL_SECONDS_);
    return true;
  } catch (e) {
    Logger.log("❌ YouTube API key validation error: " + e);
    return false;
  }
}

function resetYouTubeKeyValidationCache_() {
  const key = getYouTubeApiKey_();
  if (key) {
    const cacheKey = YT_KEY_VALID_CACHE_PREFIX_ + key.slice(-8);
    CacheService.getDocumentCache().remove(cacheKey);
  }
  YT_API_KEY_INVALID_ = false;
  YT_API_KEY_INVALID_LOGGED_ = false;
  Logger.log("✅ YouTube API key validation cache reset.");
}

function computeMedianViewsFromYouTube_(channelUrl, apiKey) {
  try {
    if (YT_API_KEY_INVALID_) return null;

    const channelId = getChannelIdFromUrl(channelUrl, apiKey);
    if (!channelId) {
      Logger.log(`resolve failed for URL: ${channelUrl}`);
      return null;
    }

    const channelData = ytFetchJson_(
      `https://www.googleapis.com/youtube/v3/channels?part=contentDetails&id=${encodeURIComponent(channelId)}&key=${encodeURIComponent(apiKey)}`
    );
    if (!channelData.items || channelData.items.length === 0) {
      Logger.log(`channel lookup returned no items for channelId=${channelId} url=${channelUrl}`);
      return null;
    }

    const uploadsPlaylistId = channelData.items[0].contentDetails.relatedPlaylists.uploads;
    if (!uploadsPlaylistId) {
      Logger.log(`uploads playlist missing for channelId=${channelId} url=${channelUrl}`);
      return null;
    }

    const playlistData = ytFetchJson_(
      `https://www.googleapis.com/youtube/v3/playlistItems?part=snippet&maxResults=${YT_MAX_PLAYLIST_ITEMS_}&playlistId=${encodeURIComponent(uploadsPlaylistId)}&key=${encodeURIComponent(apiKey)}`
    );
    if (!playlistData.items || playlistData.items.length === 0) {
      Logger.log(`uploads playlist returned no items for channelId=${channelId} url=${channelUrl}`);
      return null;
    }

    const cutoff = new Date();
    cutoff.setMonth(cutoff.getMonth() - YT_MONTHS_BACK_);
    const candidateIds = [];

    for (const item of playlistData.items) {
      const snippet = item.snippet;
      if (!snippet) continue;

      const publishedAt = snippet.publishedAt ? new Date(snippet.publishedAt) : null;
      if (!publishedAt || publishedAt < cutoff) continue;

      const videoId = snippet.resourceId && snippet.resourceId.videoId ? snippet.resourceId.videoId : null;
      if (videoId) candidateIds.push(videoId);
    }

    if (candidateIds.length === 0) {
      Logger.log(`no uploads in last 6 months for channelId=${channelId} url=${channelUrl}`);
      return null;
    }

    const ids = candidateIds.slice(0, YT_MAX_PLAYLIST_ITEMS_);
    const videosData = ytFetchJson_(
      `https://www.googleapis.com/youtube/v3/videos?part=statistics,contentDetails&id=${encodeURIComponent(ids.join(","))}&key=${encodeURIComponent(apiKey)}`
    );
    if (!videosData.items || videosData.items.length === 0) {
      Logger.log(`videos lookup returned no items for channelId=${channelId} url=${channelUrl}`);
      return null;
    }

    const pickedViews = [];
    for (const video of videosData.items) {
      const duration = video.contentDetails && video.contentDetails.duration ? video.contentDetails.duration : null;
      if (iso8601DurationToSeconds_(duration) <= YT_MIN_DURATION_SECONDS_) continue;

      const rawViews = video.statistics && video.statistics.viewCount ? Number(video.statistics.viewCount) : null;
      if (rawViews == null || isNaN(rawViews)) continue;

      pickedViews.push(rawViews);
      if (pickedViews.length >= YT_MAX_VIDEOS_FOR_AVG_) break; // reused for median sample size
    }

    if (pickedViews.length === 0) {
      Logger.log(`no videos >3 minutes with viewCount in last 6 months for channelId=${channelId} url=${channelUrl}`);
      return null;
    }

    const sorted = pickedViews.slice().sort((a, b) => a - b);
    const mid = Math.floor(sorted.length / 2);
    const median = sorted.length % 2 !== 0
      ? sorted[mid]
      : Math.round((sorted[mid - 1] + sorted[mid]) / 2);
    return median;
  } catch (e) {
    Logger.log("computeMedianViewsFromYouTube_ error: " + (e && e.stack ? e.stack : e));
    return null;
  }
}

function getChannelIdFromUrl(channelUrl, apiKey, depth) {
  if (!channelUrl || typeof channelUrl !== "string") return null;
  const level = Number(depth || 0);
  if (level > 1) return null;

  const raw = String(channelUrl).trim();
  const directIdMatch = raw.match(/(UC[a-zA-Z0-9_-]{22})/);
  if (directIdMatch && directIdMatch[1]) return directIdMatch[1];

  const normalizedUrl = /^https?:\/\//i.test(raw) ? raw : ("https://" + raw);
  if (!/youtu\.?be|youtube\.com/i.test(normalizedUrl)) return null;

  const pathMatch = normalizedUrl.match(/^https?:\/\/(?:www\.)?(?:m\.)?(?:youtube\.com|youtu\.be)\/([^?#]*)/i);
  const pathParts = ((pathMatch && pathMatch[1]) ? pathMatch[1] : "")
    .split("/")
    .filter(Boolean)
    .map(function (part) {
      try { return decodeURIComponent(part); } catch (e) { return part; }
    });

  if (pathParts.length === 0) return null;

  const first = pathParts[0];
  const second = pathParts[1] || "";

  if (first.charAt(0) === "@") {
    return (
      pickChannelIdFromHandle_(first, apiKey) ||
      channelIdFromOEmbed_(normalizedUrl, apiKey, level) ||
      channelIdFromPageHtml_(normalizedUrl)
    );
  }

  if (first === "channel" && second) return second;

  if (first === "user" && second) {
    return resolveByNameOrFallback_(second, normalizedUrl, apiKey, level);
  }

  if (first === "c" && second) {
    return resolveByNameOrFallback_(second, normalizedUrl, apiKey, level);
  }

  if (first && YT_SKIP_PATHS_.indexOf(first) === -1) {
    return resolveByNameOrFallback_(first, normalizedUrl, apiKey, level);
  }

  return channelIdFromOEmbed_(normalizedUrl, apiKey, level) || channelIdFromPageHtml_(normalizedUrl);
}

function resolveByNameOrFallback_(name, channelUrl, apiKey, depth) {
  return (
    pickChannelIdFromUsername_(name, apiKey) ||
    pickChannelIdFromSearch_(name, apiKey) ||
    channelIdFromOEmbed_(channelUrl, apiKey, depth) ||
    channelIdFromPageHtml_(channelUrl)
  );
}

function pickChannelIdFromSearch_(query, apiKey) {
  const q = String(query || "").trim();
  if (!q) return null;
  const data = ytFetchJson_(
    "https://www.googleapis.com/youtube/v3/search?part=snippet&type=channel&maxResults=1" +
    "&q=" + encodeURIComponent(q) +
    "&key=" + encodeURIComponent(apiKey)
  );
  const item = data.items && data.items[0];
  if (!item) return null;
  return (item.id && item.id.channelId) ? item.id.channelId : ((item.snippet && item.snippet.channelId) || null);
}

function pickChannelIdFromHandle_(handle, apiKey) {
  const clean = String(handle || "").replace(/^@+/, "").trim();
  if (!clean) return null;

  const candidates = [clean, clean.toLowerCase()];
  for (const candidate of candidates) {
    const data = ytFetchJson_(
      "https://www.googleapis.com/youtube/v3/channels?part=id" +
      "&forHandle=" + encodeURIComponent(candidate) +
      "&key=" + encodeURIComponent(apiKey)
    );
    if (data.items && data.items.length > 0 && data.items[0].id) return data.items[0].id;
  }

  return pickChannelIdFromSearch_("@" + clean, apiKey) || pickChannelIdFromSearch_(clean, apiKey);
}

function pickChannelIdFromUsername_(username, apiKey) {
  const clean = String(username || "").trim();
  if (!clean) return null;
  const data = ytFetchJson_(
    "https://www.googleapis.com/youtube/v3/channels?part=id" +
    "&forUsername=" + encodeURIComponent(clean) +
    "&key=" + encodeURIComponent(apiKey)
  );
  if (data.items && data.items.length > 0 && data.items[0].id) return data.items[0].id;
  return null;
}

function channelIdFromOEmbed_(channelUrl, apiKey, depth) {
  if (!channelUrl) return null;
  const level = Number(depth || 0);
  if (level > 1) return null;

  const data = ytFetchJson_(
    "https://www.youtube.com/oembed?url=" + encodeURIComponent(channelUrl) + "&format=json"
  );
  const authorUrl = data && data.author_url ? String(data.author_url).trim() : "";
  if (!authorUrl || authorUrl === channelUrl) return null;
  return getChannelIdFromUrl(authorUrl, apiKey, level + 1);
}

function channelIdFromPageHtml_(channelUrl) {
  if (!channelUrl) return null;

  const base = String(channelUrl).replace(/\/+$/, "");
  const urls = [base];
  if (!/\/about(\?|$)/i.test(base) && !/\/(watch|shorts|playlist|results|feed)(\/|$)/i.test(base)) {
    urls.push(base + "/about");
  }

  for (const url of urls) {
    try {
      const resp = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        followRedirects: true,
        headers: { "User-Agent": "Mozilla/5.0 (compatible; GoogleAppsScript)" }
      });
      if (resp.getResponseCode() >= 400) continue;

      const html = String(resp.getContentText() || "");
      if (!html) continue;

      for (const pattern of YT_CHANNEL_ID_PATTERNS_) {
        const match = html.match(pattern);
        if (match && match[1]) return match[1];
      }
    } catch (e) {}
  }

  return null;
}

function isInvalidApiKeyResponse_(text) {
  return text.indexOf("API_KEY_INVALID") !== -1 || text.indexOf("API key not valid") !== -1;
}

function markApiKeyInvalid_() {
  YT_API_KEY_INVALID_ = true;
  if (!YT_API_KEY_INVALID_LOGGED_) {
    Logger.log("❌ YouTube API key invalid. Update Script Property YOUTUBE_API_KEY and rerun.");
    YT_API_KEY_INVALID_LOGGED_ = true;
  }
}

function ytFetchJson_(url) {
  const strUrl = String(url || "");
  if (YT_API_KEY_INVALID_ && /googleapis\.com\/youtube\/v3/i.test(strUrl)) return {};

  const resp = UrlFetchApp.fetch(strUrl, { muteHttpExceptions: true });
  const code = resp.getResponseCode();
  const text = String(resp.getContentText() || "");

  if (code >= 400) {
    if (isInvalidApiKeyResponse_(text)) {
      markApiKeyInvalid_();
      return {};
    }
    Logger.log(`YouTube API error ${code}: ${text.slice(0, 500)}`);
    return {};
  }

  try {
    return JSON.parse(text);
  } catch (e) {
    Logger.log("YouTube API parse error: " + e);
    return {};
  }
}

function iso8601DurationToSeconds_(duration) {
  const match = String(duration || "").match(/PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?/);
  if (!match) return 0;
  const hours = Number(match[1] || 0);
  const minutes = Number(match[2] || 0);
  const seconds = Number(match[3] || 0);
  return (hours * 3600) + (minutes * 60) + seconds;
}

/***************************************
 * HUBSPOT -> PITCHING IMPORT
 ***************************************/

const HUBSPOT_TOKEN_PROP_ = "HUBSPOT_PRIVATE_APP_TOKEN";
const HUBSPOT_API_BASE_ = "https://api.hubapi.com";

// Internal name of the Activations custom object type in HubSpot.
// Update this if the object type identifier differs in your portal.
const HUBSPOT_ACTIVATION_OBJECT_TYPE_ = "activations";

// Keep all internal property names here so they are easy to fix later.
const HUBSPOT_PROP_MAP_ = {
  deal: {
    dealType: "dealtype",
    pitchingStatus: "pitching_status"
  },
  activation: {
    activationType: "activation_type",
    amount: "ext_amount"
  },
  contact: {
    firstName: "firstname",
    lastName: "lastname",
    influencerUrl: "influencer_url",
    countryRegion: "country",
    influencerVertical: "influencer_vertical"
  }
};

function importHubSpotDealsListToPitching() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Pitching");
  if (!sheet) return Logger.log("❌ 'Pitching' sheet not found.");

  const segmentId = promptForHubSpotSegmentId_();
  if (!segmentId) {
    return Logger.log("ℹ️ Import cancelled: no HubSpot segment ID provided.");
  }

  const token = getHubSpotToken_();
  if (!token) {
    Logger.log("❌ Missing HubSpot token. Set Script Property: HUBSPOT_PRIVATE_APP_TOKEN");
    return;
  }

  const data = sheet.getDataRange().getValues();
  if (!data.length) return Logger.log("❌ Pitching sheet is empty.");

  const header = (data[0] || []).map(v => String(v || "").trim());

  const requiredSheetCols = [
    "Channel Name",
    "Deal Type",
    "Activation Type",
    "Channel URL",
    "Country/Region",
    "Influencer Vertical",
    "Rate",
    "Median Views",
    "HubSpot Record ID"
  ];

  const missingSheetCols = requiredSheetCols.filter(name => header.indexOf(name) === -1);
  if (missingSheetCols.length) {
    return Logger.log("❌ Pitching missing required columns: " + missingSheetCols.join(", "));
  }

  const activeSection0 = findSectionRowByLabel_(data, "Active Pitches");
  if (activeSection0 === -1) {
    return Logger.log("❌ Could not find 'Active Pitches' section in Pitching.");
  }

  const templateRow1 = activeSection0 + 2; // row below the section label
  if (templateRow1 > sheet.getLastRow()) {
    return Logger.log("❌ Template row below 'Active Pitches' not found.");
  }

  Logger.log(`▶️ Fetching deal IDs from HubSpot segment ${segmentId}...`);
  const dealIds = fetchHubSpotListDealIds_(segmentId, token);
  if (!dealIds.length) {
    return Logger.log(`ℹ️ No deal IDs found in HubSpot segment ${segmentId}.`);
  }
  Logger.log(`✅ Got ${dealIds.length} deal ID(s) from HubSpot segment ${segmentId}.`);

  const deals = fetchHubSpotDealsByIds_(dealIds, token);
  if (!deals.length) {
    return Logger.log("ℹ️ No deals returned from HubSpot batch read.");
  }

  const dealTypeLookup = fetchHubSpotPropertyLabelLookup_("deals", HUBSPOT_PROP_MAP_.deal.dealType, token);

  const dealToContactId = fetchDealToPrimaryContactMap_(dealIds, token);
  const contactIds = Object.values(dealToContactId).filter(Boolean);

  let contactsById = {};
  if (contactIds.length) {
    contactsById = fetchHubSpotContactsByIds_(contactIds, token);
  }

  let activationsObjectTypeId;
  try {
    const activationsInfo = loadHubSpotCustomObjectInfo_({ key: HUBSPOT_ACTIVATION_OBJECT_TYPE_, aliases: ["Activation", "Activations", HUBSPOT_ACTIVATION_OBJECT_TYPE_] }, token);
    activationsObjectTypeId = activationsInfo.objectTypeId;
    Logger.log(`✅ Resolved Activations object type ID: ${activationsObjectTypeId}`);
  } catch (e) {
    Logger.log(`❌ Could not resolve Activations object type: ${e}`);
    return;
  }

  const dealToActivationIds = fetchDealToActivationIdsMap_(dealIds, activationsObjectTypeId, token);
  const allActivationIds = uniqueStrings_(
    Object.values(dealToActivationIds).reduce((acc, ids) => acc.concat(ids), [])
  );
  const activationsById = allActivationIds.length
    ? fetchHubSpotActivationsByIds_(allActivationIds, activationsObjectTypeId, token)
    : {};

  const rowsToInsert = [];
  const insertedDealIds = new Set();
  for (const deal of deals) {
    const dealId = String(deal.id || "").trim();
    if (!dealId) continue;

    const contactId = dealToContactId[dealId];
    const contact = contactId ? contactsById[String(contactId)] : null;
    const activationIds = dealToActivationIds[dealId] || [];

    if (activationIds.length === 0) {
      Logger.log(`⚠️ Deal ${dealId} has no activations, skipping.`);
      continue;
    }

    for (const activationId of activationIds) {
      const activation = activationsById[activationId] || null;
      const row = mapHubSpotDealAndContactToPitchingRow_(deal, contact, activation, header, dealTypeLookup);
      rowsToInsert.push(row);
    }
    insertedDealIds.add(dealId);
  }

  if (!rowsToInsert.length) {
    return Logger.log("ℹ️ No rows to insert.");
  }

  insertRowsUnderActivePitches_(sheet, rowsToInsert, header, templateRow1);
  Logger.log(`✅ Imported ${rowsToInsert.length} row(s) into Pitching.`);

  updateHubSpotDealsPitchingStatus_(Array.from(insertedDealIds), "Pitched", token);
}


/***************************************
 * HUBSPOT FETCHERS
 ***************************************/

function getHubSpotToken_() {
  return String(
    PropertiesService.getScriptProperties().getProperty(HUBSPOT_TOKEN_PROP_) || ""
  ).trim();
}

function promptForHubSpotSegmentId_() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "Import HubSpot segment to Pitching",
    "Paste the HubSpot segment ID to import.",
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return "";

  const segmentId = String(response.getResponseText() || "").trim();
  if (!segmentId) {
    ui.alert("HubSpot segment ID is required.");
    return "";
  }

  return segmentId;
}

function hubspotHeaders_(token) {
  return {
    Authorization: "Bearer " + token,
    "Content-Type": "application/json"
  };
}

function loadHubSpotCustomObjectInfo_(spec, token) {
  const schemasResponse = hubspotFetchJson_(`${HUBSPOT_API_BASE_}/crm-object-schemas/v3/schemas`, token);
  const schemas = Array.isArray(schemasResponse && schemasResponse.results)
    ? schemasResponse.results
    : [];
  const wanted = (spec.aliases || [spec.key]).map(normalizeHubSpotOptionToken_).filter(Boolean);
  const match = schemas.find(schema => {
    const candidates = [
      schema && schema.name,
      schema && schema.labels && schema.labels.singular,
      schema && schema.labels && schema.labels.plural
    ].map(normalizeHubSpotOptionToken_).filter(Boolean);
    return wanted.some(alias => candidates.indexOf(alias) !== -1);
  });
  if (!match) throw new Error(`Could not find HubSpot custom object schema for "${spec.key}".`);
  return { objectTypeId: match.objectTypeId, label: (match.labels && (match.labels.plural || match.labels.singular)) || spec.key };
}

function normalizeHubSpotOptionToken_(s) {
  return s ? String(s).toLowerCase().replace(/[^a-z0-9]/g, "") : "";
}

function hubspotFetchJson_(url, token, options) {
  const opts = Object.assign(
    {
      method: "get",
      muteHttpExceptions: true,
      headers: hubspotHeaders_(token)
    },
    options || {}
  );

  const resp = UrlFetchApp.fetch(url, opts);
  const code = resp.getResponseCode();
  const text = String(resp.getContentText() || "");

  if (code >= 400) {
    Logger.log(`❌ HubSpot API error ${code}: ${text.slice(0, 1000)}`);
    return null;
  }

  try {
    return JSON.parse(text);
  } catch (e) {
    Logger.log("❌ HubSpot API parse error: " + e);
    return null;
  }
}

function fetchHubSpotListDealIds_(listId, token) {
  const ids = [];
  let after = null;

  do {
    let url = `${HUBSPOT_API_BASE_}/crm/v3/lists/${encodeURIComponent(listId)}/memberships`;
    if (after != null) {
      url += `?after=${encodeURIComponent(after)}`;
    }

    const data = hubspotFetchJson_(url, token);
    if (!data) break;

    const results = Array.isArray(data.results) ? data.results : [];
    for (const item of results) {
      const recordId = item && item.recordId != null ? String(item.recordId).trim() : "";
      if (recordId) ids.push(recordId);
    }

    after = data.paging && data.paging.next && data.paging.next.after
      ? data.paging.next.after
      : null;
  } while (after);

  return ids;
}

function fetchHubSpotDealsByIds_(dealIds, token) {
  const chunks = chunkArray_(dealIds, 100);
  const out = [];

  for (const chunk of chunks) {
    const body = {
      inputs: chunk.map(id => ({ id: String(id) })),
      properties: [
        HUBSPOT_PROP_MAP_.deal.dealType
      ]
    };

    const data = hubspotFetchJson_(
      `${HUBSPOT_API_BASE_}/crm/v3/objects/deals/batch/read`,
      token,
      {
        method: "post",
        payload: JSON.stringify(body)
      }
    );

    if (!data || !Array.isArray(data.results)) continue;
    out.push.apply(out, data.results);
  }

  return out;
}

function fetchDealToPrimaryContactMap_(dealIds, token) {
  const chunks = chunkArray_(dealIds, 1000);
  const out = {};

  for (const chunk of chunks) {
    const body = {
      inputs: chunk.map(id => ({ id: String(id) }))
    };

    const data = hubspotFetchJson_(
      `${HUBSPOT_API_BASE_}/crm/v4/associations/deals/contacts/batch/read`,
      token,
      {
        method: "post",
        payload: JSON.stringify(body)
      }
    );

    if (!data || !Array.isArray(data.results)) continue;

    for (const item of data.results) {
      const fromId = item && item.from && item.from.id != null ? String(item.from.id) : "";
      const to = Array.isArray(item.to) ? item.to : [];
      const firstContactId = to.length && to[0] && to[0].toObjectId != null
        ? String(to[0].toObjectId)
        : "";
      if (fromId && firstContactId) out[fromId] = firstContactId;
    }
  }

  return out;
}

function fetchHubSpotContactsByIds_(contactIds, token) {
  const chunks = chunkArray_(uniqueStrings_(contactIds), 100);
  const out = {};

  for (const chunk of chunks) {
    const body = {
      inputs: chunk.map(id => ({ id: String(id) })),
      properties: [
        HUBSPOT_PROP_MAP_.contact.firstName,
        HUBSPOT_PROP_MAP_.contact.lastName,
        HUBSPOT_PROP_MAP_.contact.influencerUrl,
        HUBSPOT_PROP_MAP_.contact.countryRegion,
        HUBSPOT_PROP_MAP_.contact.influencerVertical
      ]
    };

    const data = hubspotFetchJson_(
      `${HUBSPOT_API_BASE_}/crm/v3/objects/contacts/batch/read`,
      token,
      {
        method: "post",
        payload: JSON.stringify(body)
      }
    );

    if (!data || !Array.isArray(data.results)) continue;

    for (const item of data.results) {
      if (item && item.id != null) {
        out[String(item.id)] = item;
      }
    }
  }

  return out;
}

function fetchHubSpotPropertyLabelLookup_(objectType, propertyName, token) {
  const data = hubspotFetchJson_(
    `${HUBSPOT_API_BASE_}/crm/v3/properties/${encodeURIComponent(objectType)}/${encodeURIComponent(propertyName)}`,
    token
  );
  const options = data && Array.isArray(data.options) ? data.options : [];
  const lookup = {};
  for (const option of options) {
    if (option && option.value != null) {
      lookup[String(option.value)] = option.label != null ? String(option.label) : String(option.value);
    }
  }
  return lookup;
}

function updateHubSpotDealsPitchingStatus_(dealIds, status, token) {
  const ids = (dealIds || []).map(id => String(id).trim()).filter(Boolean);
  if (!ids.length) return;

  const chunks = chunkArray_(ids, 100);
  let updated = 0;

  for (const chunk of chunks) {
    const body = {
      inputs: chunk.map(id => ({
        id,
        properties: { [HUBSPOT_PROP_MAP_.deal.pitchingStatus]: status }
      }))
    };

    const data = hubspotFetchJson_(
      `${HUBSPOT_API_BASE_}/crm/v3/objects/deals/batch/update`,
      token,
      { method: "post", payload: JSON.stringify(body) }
    );

    if (data && Array.isArray(data.results)) updated += data.results.length;
  }

  Logger.log(`✅ Updated Pitching Status to "${status}" for ${updated} deal(s) in HubSpot.`);
}


function fetchDealToActivationIdsMap_(dealIds, activationsObjectTypeId, token) {
  const out = {};
  const chunks = chunkArray_(dealIds, 100);

  for (const chunk of chunks) {
    const body = { inputs: chunk.map(id => ({ id: String(id) })) };
    const data = hubspotFetchJson_(
      `${HUBSPOT_API_BASE_}/crm/v4/associations/deals/${encodeURIComponent(activationsObjectTypeId)}/batch/read`,
      token,
      { method: "post", payload: JSON.stringify(body) }
    );
    if (!data || !Array.isArray(data.results)) continue;
    for (const item of data.results) {
      const fromId = item && item.from && item.from.id != null ? String(item.from.id) : "";
      const to = Array.isArray(item.to) ? item.to : [];
      if (fromId) {
        out[fromId] = to
          .map(t => t && t.toObjectId != null ? String(t.toObjectId) : "")
          .filter(Boolean);
      }
    }
  }

  return out;
}

function fetchHubSpotActivationsByIds_(activationIds, activationsObjectTypeId, token) {
  const chunks = chunkArray_(activationIds, 100);
  const out = {};

  for (const chunk of chunks) {
    const body = {
      inputs: chunk.map(id => ({ id: String(id) })),
      properties: [
        HUBSPOT_PROP_MAP_.activation.activationType,
        HUBSPOT_PROP_MAP_.activation.amount
      ]
    };

    const data = hubspotFetchJson_(
      `${HUBSPOT_API_BASE_}/crm/v3/objects/${encodeURIComponent(activationsObjectTypeId)}/batch/read`,
      token,
      { method: "post", payload: JSON.stringify(body) }
    );

    if (!data || !Array.isArray(data.results)) continue;

    for (const item of data.results) {
      if (item && item.id != null) out[String(item.id)] = item;
    }
  }

  return out;
}


/***************************************
 * MAPPING + INSERT
 ***************************************/

function mapHubSpotDealAndContactToPitchingRow_(deal, contact, activation, header, dealTypeLookup) {
  const row = new Array(header.length).fill("");

  const dealProps = (deal && deal.properties) ? deal.properties : {};
  const contactProps = (contact && contact.properties) ? contact.properties : {};
  const activationProps = (activation && activation.properties) ? activation.properties : {};

  const firstName = safeHubSpotValue_(contactProps[HUBSPOT_PROP_MAP_.contact.firstName]);
  const lastName = safeHubSpotValue_(contactProps[HUBSPOT_PROP_MAP_.contact.lastName]);
  const fullName = [firstName, lastName].filter(Boolean).join(" ").trim();

  const rawDealType = safeHubSpotValue_(dealProps[HUBSPOT_PROP_MAP_.deal.dealType]);
  const dealTypeLabel = (dealTypeLookup && rawDealType && dealTypeLookup[rawDealType]) ? dealTypeLookup[rawDealType] : rawDealType;

  setIfColumnExists_(row, header, "Channel Name", fullName);
  setIfColumnExists_(row, header, "Status", "Pitched");
  setIfColumnExists_(row, header, "Deal Type", dealTypeLabel);
  setIfColumnExists_(row, header, "Activation Type", safeHubSpotValue_(activationProps[HUBSPOT_PROP_MAP_.activation.activationType]));
  setIfColumnExists_(row, header, "Channel URL", safeHubSpotValue_(contactProps[HUBSPOT_PROP_MAP_.contact.influencerUrl]));
  setIfColumnExists_(row, header, "Country/Region", safeHubSpotValue_(contactProps[HUBSPOT_PROP_MAP_.contact.countryRegion]));
  setIfColumnExists_(row, header, "Influencer Vertical", safeHubSpotValue_(contactProps[HUBSPOT_PROP_MAP_.contact.influencerVertical]));
  setIfColumnExists_(row, header, "Rate", safeHubSpotValue_(activationProps[HUBSPOT_PROP_MAP_.activation.amount]));
  // Median Views is not imported from HubSpot
  setIfColumnExists_(row, header, "HubSpot Record ID", String(deal.id || "").trim());

  // Leave these blank on purpose:
  // Availability, CPM, ARCH. Comment, Approved, Rejected, Still Checking, Counteroffer

  return row;
}

function insertRowsUnderActivePitches_(sheet, rowsToInsert, header, templateRow1) {
  if (!rowsToInsert.length) return;

  const numCols = header.length;
  const insertAt1 = templateRow1; // insert before the template row
  const formatSourceRow1 = templateRow1 + rowsToInsert.length; // original template row shifts down

  sheet.insertRowsBefore(insertAt1, rowsToInsert.length);

  const writeRange = sheet.getRange(insertAt1, 1, rowsToInsert.length, numCols);
  const templateRange = getSafeFormatRow_(sheet, formatSourceRow1, numCols);

  templateRange.copyTo(writeRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  templateRange.copyTo(writeRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  sheet.setRowHeights(insertAt1, rowsToInsert.length, sheet.getRowHeight(formatSourceRow1));
  writeRange.setValues(rowsToInsert);
}


/***************************************
 * SMALL HELPERS
 ***************************************/

function safeHubSpotValue_(value) {
  return value == null ? "" : value;
}

function setIfColumnExists_(row, header, columnName, value) {
  const idx = header.indexOf(columnName);
  if (idx !== -1) row[idx] = value;
}

function chunkArray_(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) {
    out.push(arr.slice(i, i + size));
  }
  return out;
}

function uniqueStrings_(arr) {
  return Array.from(new Set((arr || []).map(v => String(v || "").trim()).filter(Boolean)));
}
