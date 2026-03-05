// Utilities for Pitching/Campaigns sync + archive + Average Views enrichment.


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

  const pApprovedCol = pitchHeader.indexOf("Approved");
  const pAvailabilityCol = pitchHeader.indexOf("Availability");
  const pChannelCol = pitchHeader.indexOf("Channel Name");
  const cChannelCol = campHeader.indexOf("Channel Name");
  if (pApprovedCol === -1 || pAvailabilityCol === -1 || pChannelCol === -1) {
    return Logger.log("❌ Pitching missing required columns: Approved, Availability, Channel Name.");
  }
  if (cChannelCol === -1) return Logger.log("❌ Campaigns missing required column: Channel Name.");

  const archivedStart0 = findSectionRowByLabel_(pitchData, "Archived");
  if (archivedStart0 === -1) return Logger.log("❌ Pitching section 'Archived' not found.");

  const commonCols = getCommonColumnsByHeader_(pitchHeader, campHeader);
  if (commonCols.length === 0) return Logger.log("ℹ️ No matching columns between Pitching and Campaigns.");

  const existingBySection = getExistingSignaturesBySection_(campData, commonCols);
  const candidates = [];
  for (let r = 1; r < archivedStart0; r++) {
    const row = pitchData[r];
    if (isBlankRow_(row) || isSectionLabelRow_(row)) continue;
    if (!asBool_(row[pApprovedCol])) continue;

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
    const signature = buildSignatureFromPitchRow_(item.values, commonCols);
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
    for (const m of commonCols) out[m.cIdx] = item.values[m.pIdx];

    const targetRange = campaignsSheet.getRange(targetRow1, 1, 1, campNumCols);
    const templateRange = getSafeFormatRow_(campaignsSheet, formatSourceRow1, campNumCols);
    templateRange.copyTo(targetRange, { formatOnly: true });
    targetRange.setValues([out]);

    existingBySection[item.availability].add(signature);
    inserted++;
  }

  Logger.log(`✅ Push done. Inserted: ${inserted}. Skipped duplicates: ${skippedDup}. Missing section: ${skippedNoSection}.`);
}


// 2) Archive pitches (Approved + Rejected)
function archivePitches() {
  try {
    Logger.log("▶️ Running pushConfirmedCreatorsToCampaigns() before archive.");
    pushConfirmedCreatorsToCampaigns();
    SpreadsheetApp.flush();
  } catch (e) {
    Logger.log(`⚠️ Push step failed; continuing archive: ${e && e.stack ? e.stack : e}`);
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Pitching");
  if (!sheet) return Logger.log("❌ 'Pitching' sheet not found.");

  const data = sheet.getDataRange().getValues();
  const archivedStart0 = findSectionRowByLabel_(data, "Archived");
  if (archivedStart0 === -1) return Logger.log("❌ Could not find 'Archived' section in Pitching.");

  const header = (data[0] || []).map(v => String(v || "").trim());
  const approvedCol = header.indexOf("Approved");
  const rejectedCol = header.indexOf("Rejected");
  const activationCol = header.indexOf("Activation");
  if (approvedCol === -1 || rejectedCol === -1) {
    return Logger.log("❌ Pitching missing required columns: Approved and/or Rejected.");
  }

  const rowsToArchive = [];
  for (let r = 1; r < archivedStart0; r++) {
    const row = data[r];
    if (isBlankRow_(row) || isSectionLabelRow_(row)) continue;
    if (!(asBool_(row[approvedCol]) || asBool_(row[rejectedCol]))) continue;
    rowsToArchive.push({ sourceRow1: r + 1, values: row.slice() });
  }

  if (rowsToArchive.length === 0) return Logger.log("ℹ️ No rows to archive.");

  for (const entry of rowsToArchive.slice().sort((a, b) => b.sourceRow1 - a.sourceRow1)) {
    sheet.deleteRow(entry.sourceRow1);
  }

  const refreshed = sheet.getDataRange().getValues();
  const refreshedArchivedStart0 = findSectionRowByLabel_(refreshed, "Archived");
  if (refreshedArchivedStart0 === -1) return Logger.log("❌ Archived section disappeared after deletion.");

  let insertAt1 = findFirstEmptyRowInSection1_(refreshed, refreshedArchivedStart0);
  if (insertAt1 === -1) insertAt1 = refreshedArchivedStart0 + 2;

  const numCols = sheet.getLastColumn();
  const arrayFormulaCols0 = getArrayFormulaColumns_(sheet, 2, numCols);
  const templateRange = getSafeFormatRow_(sheet, refreshedArchivedStart0 + 2, numCols);

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
  if (activationCol !== -1) {
    sheet.getRange(insertAt1, activationCol + 1, rowsToArchive.length, 1).setShowHyperlink(false);
  }

  Logger.log(`✅ Archived ${rowsToArchive.length} row(s).`);
}


// 3) Fill missing Average Views (menu action)
function fillMissingAverageViews_Pitching() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Pitching");
  if (!sh) return Logger.log("❌ 'Pitching' sheet not found.");

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return Logger.log("ℹ️ Pitching has no data.");

  const values = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();

  const header = (values[0] || []).map(h => String(h || "").trim());
  const urlCol = header.indexOf("Channel URL");
  const avgCol = header.indexOf("Average Views");
  const nameCol = header.indexOf("Channel Name");

  Logger.log(`🔎 Header lookup: Channel URL col=${urlCol} | Average Views col=${avgCol}`);
  if (urlCol === -1 || avgCol === -1) {
    Logger.log(`❌ Missing header(s). Found headers: ${header.join(" | ")}`);
    return;
  }

  const norm = v => String(v == null ? "" : v).trim();
  const isEmptyCell = v => norm(v) === "";
  const isBlankRow = rowArr => rowArr.join("").trim() === "";

  function isSectionLabelRow(rowArr) {
    const a = norm(rowArr[0]);
    if (a === "Active Pitches" || a === "Archived") return true;
    if (!a) return false;
    for (let c = 1; c < rowArr.length; c++) {
      if (!isEmptyCell(rowArr[c])) return false;
    }
    return true;
  }

  // stop at Archived
  let stopAt = lastRow + 1;
  for (let r = 2; r <= lastRow; r++) {
    if (norm(values[r - 1][0]) === "Archived") { stopAt = r; break; }
  }
  Logger.log(`🧭 Scanning rows 2..${stopAt - 1} (stop at Archived row ${stopAt <= lastRow ? stopAt : "not found"})`);

  const updates = [];
  let checked = 0;
  let emptyCount = 0;
  let nullCount = 0;

  for (let r = 2; r < stopAt; r++) {
    const row = values[r - 1];
    if (isSectionLabelRow(row)) continue;
    if (isBlankRow(row)) continue;

    const url = norm(row[urlCol]);
    const cur = row[avgCol];
    const chName = (nameCol !== -1) ? norm(row[nameCol]) : "";

    checked++;

    if (!url) continue;
    if (!isEmptyCell(cur)) continue;

    emptyCount++;

    const avg = getAverageViewsCached_(url);

    if (avg == null) {
      nullCount++;
      Logger.log(`⚠️ Row ${r} (${chName || "no name"}): empty Average Views but could not compute. URL=${url}`);
      continue;
    }

    updates.push({ row1: r, value: avg });
  }

  Logger.log(`📊 Checked rows: ${checked} | Empty Average Views found: ${emptyCount} | Could not compute: ${nullCount} | Will update: ${updates.length}`);

  if (updates.length === 0) {
    Logger.log("ℹ️ No missing Average Views to fill.");
    return;
  }

  // (Optional) simple throttle to avoid quota spikes on big fills
  // Utilities.sleep(150);

  updates.forEach(u => sh.getRange(u.row1, avgCol + 1).setValue(u.value));
  Logger.log(`✅ Filled Average Views for ${updates.length} row(s).`);
}


// Menu
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🧰 Scripts")
    .addItem("Push confirmed creators to Campaigns", "pushConfirmedCreatorsToCampaigns")
    .addItem("Archive pitches", "archivePitches")
    .addSeparator()
    .addItem("Fill missing Average Views", "fillMissingAverageViews_Pitching")
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

function buildSignatureFromPitchRow_(row, commonCols) {
  return commonCols.map(m => String(row[m.pIdx] || "").trim()).join("||");
}

function buildSignatureFromCampaignRow_(row, commonCols) {
  return commonCols.map(m => String(row[m.cIdx] || "").trim()).join("||");
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

function getAverageViewsCached_(channelUrl) {
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

  const avg = computeAverageViewsForChannel_(normalized, apiKey);
  if (avg == null) return null;

  const stored = String(avg);
  props.setProperty(cacheKey, stored);
  cache.put(cacheKey, stored, YT_CACHE_TTL_SECONDS_);
  return avg;
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

function computeAverageViewsForChannel_(channelUrl, apiKey) {
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
      if (pickedViews.length >= YT_MAX_VIDEOS_FOR_AVG_) break;
    }

    if (pickedViews.length === 0) {
      Logger.log(`no videos >3 minutes with viewCount in last 6 months for channelId=${channelId} url=${channelUrl}`);
      return null;
    }

    return Math.round(pickedViews.reduce((sum, value) => sum + value, 0) / pickedViews.length);
  } catch (e) {
    Logger.log("computeAverageViewsForChannel_ error: " + (e && e.stack ? e.stack : e));
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
