function fillMissingAverageViews_PitchList_Safe() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Pitch List");
  if (!sh) return Logger.log("❌ Sheet 'Pitch List' not found.");

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return Logger.log("ℹ️ Pitch List has no data.");

  const displayValues = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const rawValues = sh.getRange(1, 1, lastRow, lastCol).getValues();

  const headerRow1 = 2;
  const header = (displayValues[headerRow1 - 1] || []).map(v => String(v || "").trim());
  const urlCol = header.indexOf("Channel URL");
  const avgCol = header.indexOf("Average Views");
  const rateCol = header.indexOf("Rate");
  const cpmCol = header.indexOf("CPM");
  const nameCol = header.indexOf("Channel Name");

  if (urlCol === -1 || avgCol === -1) {
    return Logger.log("❌ Missing required headers on row 2: Channel URL and/or Average Views.");
  }

  const canWriteCpm = rateCol !== -1 && cpmCol !== -1;
  if (!canWriteCpm) Logger.log("⚠️ CPM skipped: missing Rate and/or CPM header on row 2.");

  const startDataRow1 = headerRow1 + 1;
  if (startDataRow1 > lastRow) return Logger.log("ℹ️ No data rows below header row 2.");

  const updates = [];
  let checked = 0;
  let avgMissing = 0;
  let avgComputed = 0;
  let cpmMissing = 0;
  let cpmComputed = 0;
  let failed = 0;

  for (let r = startDataRow1; r <= lastRow; r++) {
    const rowDisplay = displayValues[r - 1];
    const rowRaw = rawValues[r - 1];
    if (!rowDisplay || rowDisplay.join("").trim() === "") continue;

    const url = String(rowDisplay[urlCol] || "").trim();
    const avgCell = String(rowDisplay[avgCol] || "").trim();
    const cpmCell = canWriteCpm ? String(rowDisplay[cpmCol] || "").trim() : "";
    const name = nameCol !== -1 ? String(rowDisplay[nameCol] || "").trim() : "";

    checked++;
    if (!url) continue;

    let avg = plParseNumber_(rowRaw[avgCol]);
    let avgNeedsWrite = avgCell === "";

    if (avgNeedsWrite) {
      avgMissing++;
      avg = plGetAverageViewsCached_(url);
      if (!plIsValidAverage_(avg)) {
        failed++;
        Logger.log(`⚠️ Row ${r} (${name || "no name"}): Average Views not written. URL=${url}`);
        continue;
      }
      avgComputed++;
    }

    let cpm = null;
    let cpmNeedsWrite = false;
    if (canWriteCpm && cpmCell === "") {
      cpmMissing++;
      cpm = plCalculateCpm_(rowRaw[rateCol], avg);
      if (cpm == null) {
        Logger.log(`⚠️ Row ${r} (${name || "no name"}): CPM not written. Missing or invalid Rate / Average Views.`);
      } else {
        cpmNeedsWrite = true;
        cpmComputed++;
      }
    }

    if (!avgNeedsWrite && !cpmNeedsWrite) continue;
    updates.push({
      row1: r,
      url,
      avg: avgNeedsWrite ? avg : null,
      cpm: cpmNeedsWrite ? cpm : null
    });
  }

  Logger.log(
    `📊 Checked: ${checked} | Avg Missing: ${avgMissing} | Avg Computed: ${avgComputed}` +
    ` | CPM Missing: ${cpmMissing} | CPM Computed: ${cpmComputed} | Failed: ${failed}`
  );
  if (updates.length === 0) return Logger.log("ℹ️ Nothing to write.");

  // Final guard: write only if the row still points at the same channel.
  let avgWritten = 0;
  let cpmWritten = 0;
  for (const u of updates) {
    const currentUrl = String(sh.getRange(u.row1, urlCol + 1).getDisplayValue() || "").trim();
    if (currentUrl !== u.url) continue;

    const avgRange = sh.getRange(u.row1, avgCol + 1);
    const currentAvg = String(avgRange.getDisplayValue() || "").trim();
    if (u.avg != null && currentAvg === "") {
      avgRange.setValue(u.avg);
      avgWritten++;
    }

    if (canWriteCpm && u.cpm != null) {
      const cpmRange = sh.getRange(u.row1, cpmCol + 1);
      const currentCpm = String(cpmRange.getDisplayValue() || "").trim();
      if (currentCpm === "") {
        const liveRate = sh.getRange(u.row1, rateCol + 1).getValue();
        const liveAvg = avgRange.getValue();
        const liveCpm = plCalculateCpm_(liveRate, liveAvg);
        if (liveCpm != null) {
          cpmRange.setValue(liveCpm);
          cpmWritten++;
        }
      }
    }
  }

  Logger.log(`✅ Wrote Average Views for ${avgWritten} row(s) and CPM for ${cpmWritten} row(s).`);
}

const PL_KEY_PROP_ = "YOUTUBE_API_KEY";
const PL_CACHE_PREFIX_ = "PL_AVG_VIEWS::";
const PL_KEY_VALID_PREFIX_ = "PL_YT_KEY_VALID::";
const PL_CACHE_TTL_SECONDS_ = 21600;
const PL_KEY_VALID_TTL_SECONDS_ = 1800;
const PL_KEY_INVALID_TTL_SECONDS_ = 600;

const PL_MAX_PLAYLIST_ITEMS_ = 50;
const PL_MAX_VIDEOS_FOR_AVG_ = 15;
const PL_MIN_DURATION_SECONDS_ = 180;
const PL_MONTHS_BACK_ = 6;
const PL_SKIP_PATHS_ = ["watch", "shorts", "playlist", "results", "feed"];

const PL_CHANNEL_ID_PATTERNS_ = [
  /"channelId":"(UC[a-zA-Z0-9_-]{22})"/,
  /"externalId":"(UC[a-zA-Z0-9_-]{22})"/,
  /"browseId":"(UC[a-zA-Z0-9_-]{22})"/,
  /<meta\s+itemprop="channelId"\s+content="(UC[a-zA-Z0-9_-]{22})"/i,
  /youtube\.com\/channel\/(UC[a-zA-Z0-9_-]{22})/
];

var PL_API_KEY_INVALID_ = false;
var PL_API_KEY_INVALID_LOGGED_ = false;

function plIsValidAverage_(avg) {
  return typeof avg === "number" && isFinite(avg) && !isNaN(avg) && avg > 0;
}

function plParseNumber_(value) {
  if (typeof value === "number") return isFinite(value) && !isNaN(value) ? value : null;

  const str = String(value || "").trim();
  if (!str) return null;

  let cleaned = str.replace(/\s+/g, "").replace(/[^\d,.-]/g, "");
  if (!cleaned) return null;

  if (cleaned.indexOf(",") !== -1 && cleaned.indexOf(".") !== -1) {
    cleaned = cleaned.lastIndexOf(",") > cleaned.lastIndexOf(".")
      ? cleaned.replace(/\./g, "").replace(",", ".")
      : cleaned.replace(/,/g, "");
  } else if (cleaned.indexOf(",") !== -1) {
    cleaned = /^-?\d{1,3}(,\d{3})+$/.test(cleaned) ? cleaned.replace(/,/g, "") : cleaned.replace(",", ".");
  } else if (cleaned.indexOf(".") !== -1 && /^-?\d{1,3}(\.\d{3})+$/.test(cleaned)) {
    cleaned = cleaned.replace(/\./g, "");
  }

  const parsed = Number(cleaned);
  return isFinite(parsed) && !isNaN(parsed) ? parsed : null;
}

function plCalculateCpm_(rateValue, avgValue) {
  const rate = plParseNumber_(rateValue);
  const avg = plParseNumber_(avgValue);
  if (rate == null || avg == null || avg <= 0) return null;
  return Math.round(((rate / avg) * 1000) * 100) / 100;
}

function plGetAverageViewsCached_(channelUrl) {
  if (!channelUrl || typeof channelUrl !== "string") return null;

  const apiKey = plGetYouTubeApiKey_();
  if (!apiKey) {
    Logger.log("❌ Missing script property YOUTUBE_API_KEY.");
    return null;
  }
  if (!plValidateYouTubeApiKey_(apiKey)) return null;

  const normalized = String(channelUrl).trim().replace(/\s+/g, "");
  const key = PL_CACHE_PREFIX_ + normalized;
  const cache = CacheService.getDocumentCache();
  const props = PropertiesService.getDocumentProperties();

  const shortCached = cache.get(key);
  if (shortCached != null) return Number(shortCached);

  const persisted = props.getProperty(key);
  if (persisted != null) {
    cache.put(key, persisted, PL_CACHE_TTL_SECONDS_);
    return Number(persisted);
  }

  const avg = plComputeAverageViewsForChannel_(normalized, apiKey);
  if (!plIsValidAverage_(avg)) return null;

  const stored = String(avg);
  props.setProperty(key, stored);
  cache.put(key, stored, PL_CACHE_TTL_SECONDS_);
  return avg;
}

function plGetYouTubeApiKey_() {
  return String(PropertiesService.getScriptProperties().getProperty(PL_KEY_PROP_) || "").trim();
}

function plValidateYouTubeApiKey_(apiKey) {
  if (!apiKey || PL_API_KEY_INVALID_) return false;

  const cache = CacheService.getDocumentCache();
  const cacheKey = PL_KEY_VALID_PREFIX_ + apiKey.slice(-8);
  const cached = cache.get(cacheKey);
  if (cached === "1") return true;
  if (cached === "0") return false;

  try {
    const testUrl =
      "https://www.googleapis.com/youtube/v3/search?part=id&type=channel&maxResults=1&q=youtube" +
      "&key=" + encodeURIComponent(apiKey);
    const resp = UrlFetchApp.fetch(testUrl, { muteHttpExceptions: true });
    const code = resp.getResponseCode();
    const text = String(resp.getContentText() || "");

    if (code >= 400) {
      if (plIsInvalidKeyResponse_(text)) plMarkApiKeyInvalid_();
      cache.put(cacheKey, "0", PL_KEY_INVALID_TTL_SECONDS_);
      return false;
    }

    cache.put(cacheKey, "1", PL_KEY_VALID_TTL_SECONDS_);
    return true;
  } catch (e) {
    Logger.log("❌ API key validation error: " + e);
    return false;
  }
}

function plComputeAverageViewsForChannel_(channelUrl, apiKey) {
  try {
    if (PL_API_KEY_INVALID_) return null;

    const channelId = plGetChannelIdFromUrl_(channelUrl, apiKey, 0);
    if (!channelId) return null;

    const channelData = plFetchJson_(
      `https://www.googleapis.com/youtube/v3/channels?part=contentDetails&id=${encodeURIComponent(channelId)}&key=${encodeURIComponent(apiKey)}`
    );
    const uploads = channelData.items && channelData.items[0] &&
      channelData.items[0].contentDetails &&
      channelData.items[0].contentDetails.relatedPlaylists &&
      channelData.items[0].contentDetails.relatedPlaylists.uploads;
    if (!uploads) return null;

    const playlistData = plFetchJson_(
      `https://www.googleapis.com/youtube/v3/playlistItems?part=snippet&maxResults=${PL_MAX_PLAYLIST_ITEMS_}&playlistId=${encodeURIComponent(uploads)}&key=${encodeURIComponent(apiKey)}`
    );
    if (!playlistData.items || playlistData.items.length === 0) return null;

    const cutoff = new Date();
    cutoff.setMonth(cutoff.getMonth() - PL_MONTHS_BACK_);

    const ids = [];
    for (const item of playlistData.items) {
      const sn = item.snippet;
      if (!sn || !sn.publishedAt) continue;
      const dt = new Date(sn.publishedAt);
      if (!dt || dt < cutoff) continue;
      const vid = sn.resourceId && sn.resourceId.videoId ? sn.resourceId.videoId : null;
      if (vid) ids.push(vid);
    }
    if (ids.length === 0) return null;

    const videosData = plFetchJson_(
      `https://www.googleapis.com/youtube/v3/videos?part=statistics,contentDetails&id=${encodeURIComponent(ids.slice(0, PL_MAX_PLAYLIST_ITEMS_).join(","))}&key=${encodeURIComponent(apiKey)}`
    );
    if (!videosData.items || videosData.items.length === 0) return null;

    const views = [];
    for (const video of videosData.items) {
      const duration = video.contentDetails && video.contentDetails.duration ? video.contentDetails.duration : "";
      if (plIso8601DurationToSeconds_(duration) <= PL_MIN_DURATION_SECONDS_) continue;
      const v = video.statistics && video.statistics.viewCount ? Number(video.statistics.viewCount) : NaN;
      if (!isFinite(v) || isNaN(v) || v <= 0) continue;
      views.push(v);
      if (views.length >= PL_MAX_VIDEOS_FOR_AVG_) break;
    }
    if (views.length === 0) return null;

    return Math.round(views.reduce((a, b) => a + b, 0) / views.length);
  } catch (e) {
    Logger.log("plComputeAverageViewsForChannel_ error: " + e);
    return null;
  }
}

function plGetChannelIdFromUrl_(channelUrl, apiKey, depth) {
  const level = Number(depth || 0);
  if (!channelUrl || level > 1) return null;

  const raw = String(channelUrl).trim();
  const direct = raw.match(/(UC[a-zA-Z0-9_-]{22})/);
  if (direct && direct[1]) return direct[1];

  const normalized = /^https?:\/\//i.test(raw) ? raw : ("https://" + raw);
  if (!/youtu\.?be|youtube\.com/i.test(normalized)) return null;

  const m = normalized.match(/^https?:\/\/(?:www\.)?(?:m\.)?(?:youtube\.com|youtu\.be)\/([^?#]*)/i);
  const parts = ((m && m[1]) ? m[1] : "").split("/").filter(Boolean).map(function (p) {
    try { return decodeURIComponent(p); } catch (e) { return p; }
  });
  if (parts.length === 0) return null;

  const first = parts[0];
  const second = parts[1] || "";

  if (first.charAt(0) === "@") {
    return plPickChannelIdFromHandle_(first, apiKey) ||
      plChannelIdFromOEmbed_(normalized, apiKey, level) ||
      plChannelIdFromPageHtml_(normalized);
  }
  if (first === "channel" && second) return second;
  if (first === "user" && second) return plResolveByNameOrFallback_(second, normalized, apiKey, level);
  if (first === "c" && second) return plResolveByNameOrFallback_(second, normalized, apiKey, level);
  if (first && PL_SKIP_PATHS_.indexOf(first) === -1) return plResolveByNameOrFallback_(first, normalized, apiKey, level);

  return plChannelIdFromOEmbed_(normalized, apiKey, level) || plChannelIdFromPageHtml_(normalized);
}

function plResolveByNameOrFallback_(name, channelUrl, apiKey, depth) {
  return plPickChannelIdFromUsername_(name, apiKey) ||
    plPickChannelIdFromSearch_(name, apiKey) ||
    plChannelIdFromOEmbed_(channelUrl, apiKey, depth) ||
    plChannelIdFromPageHtml_(channelUrl);
}

function plPickChannelIdFromSearch_(query, apiKey) {
  const q = String(query || "").trim();
  if (!q) return null;
  const data = plFetchJson_(
    "https://www.googleapis.com/youtube/v3/search?part=snippet&type=channel&maxResults=1" +
    "&q=" + encodeURIComponent(q) +
    "&key=" + encodeURIComponent(apiKey)
  );
  const item = data.items && data.items[0];
  if (!item) return null;
  return (item.id && item.id.channelId) ? item.id.channelId : ((item.snippet && item.snippet.channelId) || null);
}

function plPickChannelIdFromHandle_(handle, apiKey) {
  const clean = String(handle || "").replace(/^@+/, "").trim();
  if (!clean) return null;
  const candidates = [clean, clean.toLowerCase()];

  for (const candidate of candidates) {
    const data = plFetchJson_(
      "https://www.googleapis.com/youtube/v3/channels?part=id" +
      "&forHandle=" + encodeURIComponent(candidate) +
      "&key=" + encodeURIComponent(apiKey)
    );
    if (data.items && data.items.length > 0 && data.items[0].id) return data.items[0].id;
  }

  return plPickChannelIdFromSearch_("@" + clean, apiKey) || plPickChannelIdFromSearch_(clean, apiKey);
}

function plPickChannelIdFromUsername_(username, apiKey) {
  const clean = String(username || "").trim();
  if (!clean) return null;
  const data = plFetchJson_(
    "https://www.googleapis.com/youtube/v3/channels?part=id" +
    "&forUsername=" + encodeURIComponent(clean) +
    "&key=" + encodeURIComponent(apiKey)
  );
  if (data.items && data.items.length > 0 && data.items[0].id) return data.items[0].id;
  return null;
}

function plChannelIdFromOEmbed_(channelUrl, apiKey, depth) {
  const level = Number(depth || 0);
  if (!channelUrl || level > 1) return null;
  const data = plFetchJson_("https://www.youtube.com/oembed?url=" + encodeURIComponent(channelUrl) + "&format=json");
  const authorUrl = data && data.author_url ? String(data.author_url).trim() : "";
  if (!authorUrl || authorUrl === channelUrl) return null;
  return plGetChannelIdFromUrl_(authorUrl, apiKey, level + 1);
}

function plChannelIdFromPageHtml_(channelUrl) {
  if (!channelUrl) return null;
  const base = String(channelUrl).replace(/\/+$/, "");
  const urls = [base];
  if (!/\/about(\?|$)/i.test(base) && !/\/(watch|shorts|playlist|results|feed)(\/|$)/i.test(base)) {
    urls.push(base + "/about");
  }

  for (const u of urls) {
    try {
      const resp = UrlFetchApp.fetch(u, {
        muteHttpExceptions: true,
        followRedirects: true,
        headers: { "User-Agent": "Mozilla/5.0 (compatible; GoogleAppsScript)" }
      });
      if (resp.getResponseCode() >= 400) continue;
      const html = String(resp.getContentText() || "");
      if (!html) continue;

      for (const pattern of PL_CHANNEL_ID_PATTERNS_) {
        const m = html.match(pattern);
        if (m && m[1]) return m[1];
      }
    } catch (e) {}
  }
  return null;
}

function plIsInvalidKeyResponse_(text) {
  return String(text || "").indexOf("API_KEY_INVALID") !== -1 || String(text || "").indexOf("API key not valid") !== -1;
}

function plMarkApiKeyInvalid_() {
  PL_API_KEY_INVALID_ = true;
  if (!PL_API_KEY_INVALID_LOGGED_) {
    Logger.log("❌ YouTube API key invalid for Pitch List flow.");
    PL_API_KEY_INVALID_LOGGED_ = true;
  }
}

function plFetchJson_(url) {
  const strUrl = String(url || "");
  if (PL_API_KEY_INVALID_ && /googleapis\.com\/youtube\/v3/i.test(strUrl)) return {};

  const resp = UrlFetchApp.fetch(strUrl, { muteHttpExceptions: true });
  const code = resp.getResponseCode();
  const text = String(resp.getContentText() || "");
  if (code >= 400) {
    if (plIsInvalidKeyResponse_(text)) {
      plMarkApiKeyInvalid_();
      return {};
    }
    return {};
  }
  try {
    return JSON.parse(text);
  } catch (e) {
    return {};
  }
}

function plIso8601DurationToSeconds_(duration) {
  const m = String(duration || "").match(/PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+)S)?/);
  if (!m) return 0;
  return (Number(m[1] || 0) * 3600) + (Number(m[2] || 0) * 60) + Number(m[3] || 0);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🧰 Scripts")
    .addItem("Fill missing Average Views + CPM", "fillMissingAverageViews_PitchList_Safe")
    .addToUi();
}
