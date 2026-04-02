/*********************************************
 * INT Sheet Template – Code.gs
 *
 * Internal workflow scripts:
 * 1) Creator List → HubSpot (+ enrichment)
 * 2) Creator List → Pitching (Negotiation)
 * 3) INT Pitching → EXT Pitching
 * 4) Update INT Pitching from EXT
 * 5) Update INT Campaigns from EXT
 *********************************************/


// ============================================================
//  CONSTANTS
// ============================================================

const EXT_SPREADSHEET_ID_PROP_ = "EXT_SPREADSHEET_ID";

const HUBSPOT_SHARED_IMPORT_WEB_APP_URL_ =
  "https://script.google.com/a/macros/arch.agency/s/AKfycbzI6gAHnhSlRLzheWkS_wNYvYODvx1aztd27cf2DbbuBJTSqOYe-oqtKAnZqRc7jCE8/exec";
const HUBSPOT_SHARED_IMPORT_ACTION_ = "startImport";
const HUBSPOT_SHARED_LIBRARY_IDENTIFIER_ = "HubSpotSharedImporter";

const OPENAI_API_KEY_PROP_ = "OPENAI_API_KEY";
const OPENAI_MODEL_PROP_ = "OPENAI_MODEL";
const OPENAI_DEFAULT_MODEL_ = "gpt-nano-5";

// YouTube API
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

// INT ↔ EXT column name mappings
const INT_TO_EXT_PITCHING_MAP_ = { "EXT Rate": "Rate", "EXT CPM": "CPM" };
const EXT_TO_INT_PITCHING_MAP_ = { "Rate": "EXT Rate", "CPM": "EXT CPM" };
const EXT_TO_INT_CAMPAIGNS_MAP_ = { "Rate": "EXT Rate" };
const CREATOR_TO_PITCHING_NEGOTIATION_MAP_ = { "YouTube Video Median Views": "Median Views" };
const INT_ONLY_PITCHING_COLS_ = new Set(["INT Rate", "INT CPM"]);
const CREATOR_TO_PITCHING_NEGOTIATION_SKIP_COLS_ = new Set(["Status"]);
const HUBSPOT_IMPORT_EXCLUDED_COLS_ = new Set(["Channel Name", "HubSpot Record ID", "Channel URL", "Status"]);
const YOUTUBE_ENRICH_FIELDS_ = [
  "YouTube Handle",
  "YouTube URL",
  "YouTube Video Median Views",
  "YouTube Shorts Median Views",
  "YouTube Engagement Rate",
  "YouTube Followers"
];
const PROFILE_LLM_FIELDS_ = [
  "First Name",
  "Last Name",
  "Influencer Type",
  "Influencer Vertical",
  "Country/Region",
  "Language"
];
const PROFILE_LLM_DROPDOWN_FIELDS_ = new Set([
  "Influencer Type",
  "Influencer Vertical",
  "Country/Region",
  "Language"
]);
const PROFILE_LLM_CLASSIFICATION_FIELDS_ = [
  "Influencer Vertical",
  "Country/Region",
  "Language"
];
const PROFILE_LLM_CONTEXT_FIELDS_ = [
  "Channel Name",
  "Channel URL",
  "Campaign Name",
  "Email",
  "YouTube Handle",
  "YouTube URL",
  "Instagram Handle",
  "Instagram URL",
  "TikTok Handle",
  "TikTok URL",
  "Twitch Handle",
  "Twitch URL",
  "Kick Handle",
  "Kick URL",
  "X Handle",
  "X URL"
];

const MONTH_NAMES_ = [
  "January", "February", "March", "April", "May", "June",
  "July", "August", "September", "October", "November", "December"
];


// ============================================================
//  MENU
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🧰 Scripts")
    .addItem("Enrich & Import to HubSpot", "enrichAndImportToHubSpot")
    .addSeparator()
    .addItem("Responded → Negotiation", "pushRespondedToNegotiation")
    .addSeparator()
    .addItem("Push to EXT Pitching", "pushReadyForPitchingToExt")
    .addItem("Update Pitching from EXT", "updatePitchingFromExt")
    .addItem("Update Campaigns from EXT", "updateCampaignsFromExt")
    .addSeparator()
    .addItem("Fill Median Views", "fillMissingMedianViews_Pitching")
    .addToUi();
}


// ============================================================
//  (1) CREATOR LIST → HUBSPOT (+ ENRICHMENT)
// ============================================================

function enrichAndImportToHubSpot() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Creator List");
  if (!sheet) return ui.alert("❌ 'Creator List' sheet not found.");

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return ui.alert("ℹ️ Creator List has no data.");

  const header = data[0].map(v => String(v || "").trim());
  const numCols = header.length;
  const dropdownValuesByHeader = getDropdownValuesByHeader_(ss, "Dropdown Values");
  const defaultClientName = getDefaultClientName_(dropdownValuesByHeader);

  const contactingStart0 = findSectionRowByLabel_(data, "Contacting");
  const archivedStart0 = findSectionRowByLabel_(data, "Archived");
  if (contactingStart0 === -1) return ui.alert("❌ 'Contacting' section not found.");
  if (archivedStart0 === -1) return ui.alert("❌ 'Archived' section not found.");

  const emailCol = header.indexOf("Email");
  if (emailCol === -1) return ui.alert("❌ 'Email' column not found.");

  // Collect rows with email in Contacting section
  const rowItems = [];
  for (let r = contactingStart0 + 1; r < archivedStart0; r++) {
    const row = data[r];
    if (isBlankRow_(row) || isSectionLabelRow_(row)) continue;
    const email = String(row[emailCol] || "").trim();
    if (!email) continue;
    rowItems.push({ row1: r + 1, values: row.slice() });
  }

  if (rowItems.length === 0) return ui.alert("ℹ️ No rows with emails in Contacting section.");

  // Stage 1: Static enrichment (in memory)
  let enrichedCount = 0;
  for (const item of rowItems) {
    if (enrichCreatorRow_(item.values, header, defaultClientName)) enrichedCount++;
  }

  // Batch write enriched data back
  for (const item of rowItems) {
    sheet.getRange(item.row1, 1, 1, numCols).setValues([item.values]);
  }
  SpreadsheetApp.flush();
  Logger.log(`✅ Static enrichment done for ${enrichedCount} row(s).`);

  // Stage 2: YouTube API enrichment
  enrichYouTubeDataForRows_(sheet, rowItems, header);
  SpreadsheetApp.flush();

  // Stage 3: LLM enrichment (optional, requires OPENAI_API_KEY script property)
  enrichProfileFieldsViaLlm_(sheet, rowItems, header, dropdownValuesByHeader);
  SpreadsheetApp.flush();

  // Stage 4: Re-read enriched data and import to HubSpot
  const freshData = sheet.getDataRange().getValues();
  const freshRowItems = rowItems.map(item => ({
    row1: item.row1,
    values: freshData[item.row1 - 1]
  }));
  importCreatorListToHubSpot_(ss, sheet, header, freshRowItems, ui);
}


/**
 * Enriches a single Creator List row with derived fields.
 * Modifies the row array in place.
 */
function enrichCreatorRow_(row, header, defaultClientName) {
  const idx = name => header.indexOf(name);
  const get = name => { const i = idx(name); return i === -1 ? "" : String(row[i] || "").trim(); };
  let changed = false;
  const set = (name, value) => {
    const i = idx(name);
    const nextValue = value == null ? "" : value;
    if (i === -1) return;
    if (String(row[i] == null ? "" : row[i]) === String(nextValue)) return;
    row[i] = nextValue;
    changed = true;
  };
  const setIfEmpty = (name, value) => {
    const i = idx(name);
    const nextValue = value == null ? "" : value;
    if (i === -1 || nextValue === "") return;
    if (String(row[i] || "").trim()) return;
    row[i] = nextValue;
    changed = true;
  };

  const channelName = get("Channel Name");
  const campaignName = get("Campaign Name");
  if (!channelName && !campaignName) return false;

  // Contact Type: always Influencer
  set("Contact Type", "Influencer");

  const campaignParts = parseCampaignName_(campaignName);
  if (campaignParts) {
    setIfEmpty("Month", campaignParts.month);
    setIfEmpty("Year", campaignParts.year);
  }
  setIfEmpty("Client name", defaultClientName);

  // Deal name: Channel Name - Campaign Name
  if (channelName && campaignName) {
    setIfEmpty("Deal name", channelName + " - " + campaignName);
  }

  // Pipeline: always Sales Pipeline
  set("Pipeline", "Sales Pipeline");

  // Deal stage: always Scouted
  set("Deal stage", "Scouted");

  return changed;
}

function parseCampaignName_(campaignName) {
  const raw = String(campaignName || "").trim();
  if (!raw) return null;

  const match = raw.match(/^(.+?)\s+(\d{1,2})-(\d{4})$/);
  if (!match) return null;

  const monthNum = Number(match[2]);
  if (!isFinite(monthNum) || monthNum < 1 || monthNum > 12) return null;

  return {
    month: MONTH_NAMES_[monthNum - 1],
    year: match[3]
  };
}


/**
 * Enriches YouTube data (handle, URL, followers, median views, engagement rate)
 * using Channel URL first, then a strict Channel Name search fallback.
 */
function enrichYouTubeDataForRows_(sheet, rowItems, header) {
  const apiKey = getYouTubeApiKey_();
  if (!apiKey) {
    Logger.log("ℹ️ YouTube API key not set. Skipping YouTube enrichment.");
    return;
  }
  if (!validateYouTubeApiKey_(apiKey)) return;

  let filled = 0;
  for (const item of rowItems) {
    const needsYoutubeFields = YOUTUBE_ENRICH_FIELDS_.some(field => {
      const fieldIdx = header.indexOf(field);
      return fieldIdx !== -1 && !String(item.values[fieldIdx] || "").trim();
    });
    if (!needsYoutubeFields) continue;

    try {
      if (enrichSingleYouTubeRow_(item.values, header, apiKey)) {
        sheet.getRange(item.row1, 1, 1, header.length).setValues([item.values]);
        filled++;
      }
    } catch (e) {
      Logger.log(`⚠️ YouTube enrichment failed for row ${item.row1}: ${e}`);
    }
  }
  Logger.log(`✅ YouTube enrichment: filled ${filled} row(s).`);
}


function enrichSingleYouTubeRow_(row, header, apiKey) {
  const channelUrl = String(getValueByHeader_(row, header, "Channel URL") || "").trim();
  const channelName = String(getValueByHeader_(row, header, "Channel Name") || "").trim();
  const resolved = resolveYouTubeChannelForEnrichment_(channelUrl, channelName, apiKey);
  if (!resolved || !resolved.channelId) return false;

  let changed = false;

  // Channel statistics and snippet (subscribers, handle)
  const channelData = ytFetchJson_(
    "https://www.googleapis.com/youtube/v3/channels?part=statistics,snippet" +
    "&id=" + encodeURIComponent(resolved.channelId) +
    "&key=" + encodeURIComponent(apiKey)
  );

  if (channelData.items && channelData.items.length > 0) {
    const item = channelData.items[0];
    const stats = item.statistics || {};
    const snippet = item.snippet || {};
    const canonicalUrl = resolved.canonicalUrl || buildCanonicalYouTubeUrl_(resolved.channelId, snippet.customUrl);
    const handle = normalizeYouTubeHandle_(snippet.customUrl);
    if (setIfEmpty_(row, header, "YouTube Followers", stats.subscriberCount || "")) changed = true;
    if (handle && setIfEmpty_(row, header, "YouTube Handle", handle)) changed = true;
    if (canonicalUrl && setIfEmpty_(row, header, "YouTube URL", canonicalUrl)) changed = true;
  }

  // Video stats for long-form median views, shorts median views, and engagement rate
  const videoStats = computeYouTubeVideoStats_(resolved.channelId, apiKey);
  if (videoStats) {
    if (videoStats.medianVideoViews != null) {
      if (setIfEmpty_(row, header, "YouTube Video Median Views", videoStats.medianVideoViews)) changed = true;
    }
    if (videoStats.medianShortsViews != null) {
      if (setIfEmpty_(row, header, "YouTube Shorts Median Views", videoStats.medianShortsViews)) changed = true;
    }
    if (videoStats.medianVideoEngagementRate != null) {
      if (setIfEmpty_(row, header, "YouTube Engagement Rate", videoStats.medianVideoEngagementRate)) changed = true;
    }
  }

  return changed;
}

function resolveYouTubeChannelForEnrichment_(channelUrl, channelName, apiKey) {
  const direct = resolveYouTubeChannelIdFromInput_(channelUrl, apiKey);
  if (direct) return direct;

  if (channelName) {
    return findStrongYouTubeChannelMatchByName_(channelName, apiKey);
  }

  return null;
}

function resolveYouTubeChannelIdFromInput_(input, apiKey) {
  const raw = String(input || "").trim();
  if (!raw) return null;

  if (/^UC[a-zA-Z0-9_-]{22}$/.test(raw)) {
    return {
      channelId: raw,
      canonicalUrl: buildCanonicalYouTubeUrl_(raw, "")
    };
  }

  if (raw.charAt(0) === "@") {
    const handleChannelId = pickChannelIdFromHandle_(raw, apiKey);
    if (handleChannelId) {
      return {
        channelId: handleChannelId,
        canonicalUrl: "https://www.youtube.com/" + normalizeYouTubeHandle_(raw)
      };
    }
  }

  const normalized = /^https?:\/\//i.test(raw) ? raw : ("https://" + raw);
  if (!/youtu\.?be|youtube\.com/i.test(normalized)) return null;

  const channelId = getChannelIdFromUrl(normalized, apiKey);
  if (!channelId) return null;

  return {
    channelId: channelId,
    canonicalUrl: normalizeChannelUrl_(normalized)
  };
}

function findStrongYouTubeChannelMatchByName_(channelName, apiKey) {
  const query = String(channelName || "").trim();
  if (!query) return null;

  const data = ytFetchJson_(
    "https://www.googleapis.com/youtube/v3/search?part=snippet&type=channel&maxResults=5" +
    "&q=" + encodeURIComponent(query) +
    "&key=" + encodeURIComponent(apiKey)
  );
  if (!data.items || data.items.length === 0) return null;

  for (const item of data.items) {
    const candidateId = item.id && item.id.channelId
      ? item.id.channelId
      : ((item.snippet && item.snippet.channelId) || "");
    const candidateTitle = item.snippet && item.snippet.title
      ? String(item.snippet.title).trim()
      : "";

    if (!candidateId || !isStrongYouTubeChannelMatch_(query, candidateTitle)) continue;

    return {
      channelId: candidateId,
      canonicalUrl: buildCanonicalYouTubeUrl_(candidateId, "")
    };
  }

  return null;
}

function isStrongYouTubeChannelMatch_(query, title) {
  const queryNorm = normalizeMatchText_(query);
  const titleNorm = normalizeMatchText_(title);
  if (!queryNorm || !titleNorm) return false;

  if (queryNorm === titleNorm) return true;

  const queryTight = queryNorm.replace(/\s+/g, "");
  const titleTight = titleNorm.replace(/\s+/g, "");
  if (queryTight === titleTight) return true;

  if (queryNorm.length >= 6 && (titleNorm.indexOf(queryNorm) !== -1 || queryNorm.indexOf(titleNorm) !== -1)) {
    return true;
  }

  const queryTokens = queryNorm
    .split(" ")
    .filter(token => token.length > 2 && token !== "official");
  if (queryTokens.length < 2) return false;

  const titleTokens = new Set(
    titleNorm
      .split(" ")
      .filter(token => token.length > 2 && token !== "official")
  );

  let overlap = 0;
  queryTokens.forEach(token => {
    if (titleTokens.has(token)) overlap++;
  });

  return overlap === queryTokens.length;
}

function normalizeMatchText_(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9@]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function buildCanonicalYouTubeUrl_(channelId, customUrl) {
  const handle = normalizeYouTubeHandle_(customUrl);
  if (handle) return "https://www.youtube.com/" + handle;
  return channelId ? "https://www.youtube.com/channel/" + channelId : "";
}

function normalizeYouTubeHandle_(value) {
  const raw = String(value || "").trim().replace(/^https?:\/\/(?:www\.)?youtube\.com\//i, "");
  if (!raw) return "";
  return raw.charAt(0) === "@" ? raw : ("@" + raw.replace(/^\/+/, ""));
}


function computeYouTubeVideoStats_(channelId, apiKey) {
  const channelData = ytFetchJson_(
    "https://www.googleapis.com/youtube/v3/channels?part=contentDetails" +
    "&id=" + encodeURIComponent(channelId) +
    "&key=" + encodeURIComponent(apiKey)
  );
  if (!channelData.items || channelData.items.length === 0) return null;

  const uploadsPlaylistId = channelData.items[0].contentDetails.relatedPlaylists.uploads;
  if (!uploadsPlaylistId) return null;

  const playlistData = ytFetchJson_(
    "https://www.googleapis.com/youtube/v3/playlistItems?part=snippet" +
    "&maxResults=" + YT_MAX_PLAYLIST_ITEMS_ +
    "&playlistId=" + encodeURIComponent(uploadsPlaylistId) +
    "&key=" + encodeURIComponent(apiKey)
  );
  if (!playlistData.items || playlistData.items.length === 0) return null;

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
  if (candidateIds.length === 0) return null;

  const ids = candidateIds.slice(0, YT_MAX_PLAYLIST_ITEMS_);
  const videosData = ytFetchJson_(
    "https://www.googleapis.com/youtube/v3/videos?part=statistics,contentDetails" +
    "&id=" + encodeURIComponent(ids.join(",")) +
    "&key=" + encodeURIComponent(apiKey)
  );
  if (!videosData.items || videosData.items.length === 0) return null;

  const videoViewsList = [];
  const shortsViewsList = [];
  const erList = [];

  for (const video of videosData.items) {
    const duration = video.contentDetails && video.contentDetails.duration ? video.contentDetails.duration : null;
    const durationSeconds = iso8601DurationToSeconds_(duration);

    const views = video.statistics && video.statistics.viewCount ? Number(video.statistics.viewCount) : null;
    if (views == null || isNaN(views) || views === 0) continue;

    if (durationSeconds > YT_MIN_DURATION_SECONDS_) {
      if (videoViewsList.length < YT_MAX_VIDEOS_FOR_AVG_) {
        videoViewsList.push(views);

        const likes = Number(video.statistics.likeCount || 0);
        const comments = Number(video.statistics.commentCount || 0);
        erList.push((likes + comments) / views);
      }
    } else if (shortsViewsList.length < YT_MAX_VIDEOS_FOR_AVG_) {
      shortsViewsList.push(views);
    }

    if (
      videoViewsList.length >= YT_MAX_VIDEOS_FOR_AVG_ &&
      shortsViewsList.length >= YT_MAX_VIDEOS_FOR_AVG_
    ) {
      break;
    }
  }

  if (videoViewsList.length === 0 && shortsViewsList.length === 0) return null;

  const medianVideoViews = videoViewsList.length > 0
    ? computeMedian_(videoViewsList)
    : null;
  const medianShortsViews = shortsViewsList.length > 0
    ? computeMedian_(shortsViewsList)
    : null;
  const medianVideoEngagementRate = erList.length > 0
    ? Math.round(computeMedian_(erList) * 10000) / 10000
    : null;

  return {
    medianVideoViews: medianVideoViews,
    medianShortsViews: medianShortsViews,
    medianVideoEngagementRate: medianVideoEngagementRate
  };
}


function computeMedian_(arr) {
  const sorted = arr.slice().sort((a, b) => a - b);
  const mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 !== 0
    ? sorted[mid]
    : Math.round((sorted[mid - 1] + sorted[mid]) / 2);
}


// ---- LLM enrichment (optional) ----

function enrichProfileFieldsViaLlm_(sheet, rowItems, header, dropdownValuesByHeader) {
  const apiKey = getOpenAiApiKey_();
  if (!apiKey) {
    Logger.log("ℹ️ OpenAI API key not set. Skipping LLM enrichment.");
    return;
  }

  const model = getOpenAiModel_();

  let enriched = 0;
  for (const item of rowItems) {
    const channelName = String(item.values[header.indexOf("Channel Name")] || "").trim();
    const channelUrl = String(item.values[header.indexOf("Channel URL")] || "").trim();
    const campaignName = String(item.values[header.indexOf("Campaign Name")] || "").trim();
    if (!channelName && !channelUrl && !campaignName) continue;
    const contextText = buildCreatorProfileContext_(item.values, header);

    // Only attempt fields that are still empty
    const emptyFields = PROFILE_LLM_FIELDS_.filter(f => {
      const i = header.indexOf(f);
      return i !== -1 && !String(item.values[i] || "").trim();
    });
    if (emptyFields.length === 0) continue;

    try {
      const result = callLlmForCreatorProfileEnrichment_(
        apiKey,
        model,
        channelName,
        channelUrl,
        campaignName,
        emptyFields,
        dropdownValuesByHeader,
        contextText
      );

      let rowChanged = false;
      if (result) {
        const sanitized = sanitizeCreatorProfileEnrichment_(result, dropdownValuesByHeader);
        if (applySanitizedProfileFields_(item.values, header, emptyFields, sanitized)) {
          rowChanged = true;
        }
      }

      const remainingClassificationFields = PROFILE_LLM_CLASSIFICATION_FIELDS_.filter(field => {
        const i = header.indexOf(field);
        return emptyFields.indexOf(field) !== -1 && i !== -1 && !String(item.values[i] || "").trim();
      });

      if (remainingClassificationFields.length > 0) {
        const classificationResult = callLlmForCreatorDropdownClassification_(
          apiKey,
          model,
          remainingClassificationFields,
          dropdownValuesByHeader,
          contextText
        );
        if (classificationResult) {
          const sanitizedClassification = sanitizeCreatorProfileEnrichment_(
            classificationResult,
            dropdownValuesByHeader
          );
          if (applySanitizedProfileFields_(
            item.values,
            header,
            remainingClassificationFields,
            sanitizedClassification
          )) {
            rowChanged = true;
          }
        }
      }

      if (rowChanged) {
        sheet.getRange(item.row1, 1, 1, header.length).setValues([item.values]);
        enriched++;
      }
    } catch (e) {
      Logger.log(`⚠️ LLM enrichment failed for row ${item.row1}: ${e}`);
    }
  }
  Logger.log(`✅ LLM enrichment: filled ${enriched} row(s).`);
}


function callLlmForCreatorProfileEnrichment_(
  apiKey,
  model,
  channelName,
  channelUrl,
  campaignName,
  emptyFields,
  dropdownValuesByHeader,
  contextText
) {
  const prompt = [
    "Return a JSON object only.",
    "Fill only the requested fields that you are confident about.",
    "If a field is uncertain, omit it.",
    "Only include First Name and Last Name when the creator is clearly an individual person and the split is unambiguous.",
    "For Influencer Type, Influencer Vertical, Country/Region, and Language, use only exact values from the allowed lists below.",
    "Prefer filling Influencer Vertical, Country/Region, and Language with the closest exact allowed value instead of omitting them.",
    "",
    "Requested fields: " + emptyFields.join(", "),
    "Channel Name: " + (channelName || "(blank)"),
    "Channel URL: " + (channelUrl || "(blank)"),
    "Campaign Name: " + (campaignName || "(blank)"),
    "",
    "Creator context:",
    contextText || "(blank)",
    "",
    "Allowed Influencer Type values: " + JSON.stringify(dropdownValuesByHeader["Influencer Type"] || []),
    "Allowed Influencer Vertical values: " + JSON.stringify(dropdownValuesByHeader["Influencer Vertical"] || []),
    "Allowed Country/Region values: " + JSON.stringify(dropdownValuesByHeader["Country/Region"] || []),
    "Allowed Language values: " + JSON.stringify(dropdownValuesByHeader["Language"] || [])
  ].join("\n");

  return callOpenAiJsonChat_(
    apiKey,
    model,
    "You enrich creator CRM rows. Only provide values you are confident about. Respond with valid JSON only.",
    prompt
  );
}

function callLlmForCreatorDropdownClassification_(
  apiKey,
  model,
  requestedFields,
  dropdownValuesByHeader,
  contextText
) {
  const prompt = [
    "Return a JSON object only.",
    "Choose the single best exact allowed value for each requested field.",
    "Do not invent values outside the lists.",
    "These are required CRM classification fields, so prefer the best exact allowed match instead of omitting a field.",
    "Only omit a field when there is truly no reasonable signal in the creator context.",
    "",
    "Requested fields: " + requestedFields.join(", "),
    "",
    "Creator context:",
    contextText || "(blank)",
    "",
    "Allowed Influencer Vertical values: " + JSON.stringify(dropdownValuesByHeader["Influencer Vertical"] || []),
    "Allowed Country/Region values: " + JSON.stringify(dropdownValuesByHeader["Country/Region"] || []),
    "Allowed Language values: " + JSON.stringify(dropdownValuesByHeader["Language"] || [])
  ].join("\n");

  return callOpenAiJsonChat_(
    apiKey,
    model,
    "You classify creator CRM fields. Pick only exact values from the allowed lists and respond with valid JSON only.",
    prompt
  );
}

function callOpenAiJsonChat_(apiKey, model, systemPrompt, userPrompt) {
  const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
    method: "post",
    muteHttpExceptions: true,
    headers: {
      Authorization: "Bearer " + apiKey,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify({
      model: model,
      messages: [
        { role: "system", content: systemPrompt },
        { role: "user", content: userPrompt }
      ],
      response_format: { type: "json_object" }
    })
  });

  const code = response.getResponseCode();
  if (code >= 400) {
    Logger.log(`⚠️ OpenAI API error ${code}: ${response.getContentText().slice(0, 500)}`);
    return null;
  }

  try {
    const data = JSON.parse(response.getContentText());
    const content = data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content;
    if (!content) return null;
    return JSON.parse(content);
  } catch (e) {
    Logger.log("⚠️ LLM returned invalid JSON: " + e);
    return null;
  }
}

function sanitizeCreatorProfileEnrichment_(result, dropdownValuesByHeader) {
  const out = {};
  const firstName = normalizeLlmString_(result["First Name"]);
  const lastName = normalizeLlmString_(result["Last Name"]);

  if (isReasonablePersonName_(firstName) && isReasonablePersonName_(lastName)) {
    out["First Name"] = firstName;
    out["Last Name"] = lastName;
  }

  PROFILE_LLM_DROPDOWN_FIELDS_.forEach(field => {
    const value = coerceDropdownValue_(dropdownValuesByHeader[field] || [], result[field]);
    if (value) out[field] = value;
  });

  return out;
}

function applySanitizedProfileFields_(row, header, fields, sanitized) {
  let changed = false;

  fields.forEach(field => {
    const value = sanitized[field];
    if (value == null || String(value).trim() === "") return;

    const i = header.indexOf(field);
    if (i === -1 || String(row[i] || "").trim()) return;

    row[i] = value;
    changed = true;
  });

  return changed;
}

function buildCreatorProfileContext_(row, header) {
  return PROFILE_LLM_CONTEXT_FIELDS_
    .map(field => {
      const value = normalizeLlmString_(getValueByHeader_(row, header, field));
      return value ? (field + ": " + value) : "";
    })
    .filter(Boolean)
    .join("\n");
}

function normalizeLlmString_(value) {
  return String(value == null ? "" : value).trim();
}

function isAllowedDropdownValue_(allowedValues, value) {
  return !!value && Array.isArray(allowedValues) && allowedValues.indexOf(value) !== -1;
}

function coerceDropdownValue_(allowedValues, value) {
  const rawValue = normalizeLlmString_(value);
  if (!rawValue || !Array.isArray(allowedValues) || allowedValues.length === 0) return "";

  const exactMatch = allowedValues.find(option => normalizeLlmString_(option) === rawValue);
  if (exactMatch) return normalizeLlmString_(exactMatch);

  const lowerValue = rawValue.toLowerCase();
  const caseInsensitiveMatch = allowedValues.find(option => normalizeLlmString_(option).toLowerCase() === lowerValue);
  if (caseInsensitiveMatch) return normalizeLlmString_(caseInsensitiveMatch);

  const valueVariants = getDropdownMatchVariants_(rawValue);
  for (const variant of valueVariants) {
    const normalizedExactMatch = allowedValues.find(option => normalizeDropdownMatchText_(option) === variant);
    if (normalizedExactMatch) return normalizeLlmString_(normalizedExactMatch);
  }

  const sortedValueTokenVariants = valueVariants.map(sortNormalizedTokens_).filter(Boolean);
  for (const sortedVariant of sortedValueTokenVariants) {
    const tokenOrderInsensitiveMatch = allowedValues.find(
      option => sortNormalizedTokens_(normalizeDropdownMatchText_(option)) === sortedVariant
    );
    if (tokenOrderInsensitiveMatch) return normalizeLlmString_(tokenOrderInsensitiveMatch);
  }

  for (const variant of valueVariants) {
    const partialMatches = allowedValues.filter(option => {
      const normalizedOption = normalizeDropdownMatchText_(option);
      return normalizedOption && (normalizedOption.indexOf(variant) !== -1 || variant.indexOf(normalizedOption) !== -1);
    });
    if (partialMatches.length === 1) return normalizeLlmString_(partialMatches[0]);
  }

  return "";
}

function getDropdownMatchVariants_(value) {
  const variants = new Set();
  const normalized = normalizeDropdownMatchText_(value);
  if (normalized) variants.add(normalized);

  const aliasReplacements = [
    { from: /\bu\.?s\.?a?\b/g, to: "united states" },
    { from: /\bu\.?k\.?\b/g, to: "united kingdom" },
    { from: /\bu\.?a\.?e\.?\b/g, to: "united arab emirates" }
  ];

  aliasReplacements.forEach(rule => {
    const replaced = normalized.replace(rule.from, rule.to).replace(/\s+/g, " ").trim();
    if (replaced) variants.add(replaced);
  });

  return Array.from(variants);
}

function normalizeDropdownMatchText_(value) {
  return normalizeMatchText_(String(value || "").replace(/&/g, " and "));
}

function sortNormalizedTokens_(value) {
  return String(value || "")
    .split(" ")
    .filter(Boolean)
    .sort()
    .join(" ");
}

function isReasonablePersonName_(value) {
  if (!value) return false;
  if (/\d/.test(value)) return false;
  if (!/^[A-Za-zÀ-ÖØ-öø-ÿ'’.\- ]{2,60}$/.test(value)) return false;

  const words = value.split(/\s+/).filter(Boolean);
  return words.length >= 1 && words.length <= 3;
}

function getOpenAiApiKey_() {
  return String(
    PropertiesService.getScriptProperties().getProperty(OPENAI_API_KEY_PROP_) || ""
  ).trim();
}

function getOpenAiModel_() {
  const configured = String(
    PropertiesService.getScriptProperties().getProperty(OPENAI_MODEL_PROP_) || ""
  ).trim();
  return configured || OPENAI_DEFAULT_MODEL_;
}


// ---- HubSpot import ----

function importCreatorListToHubSpot_(ss, sheet, header, rowItems, ui) {
  const emailCol = header.indexOf("Email");

  // Validate emails
  const emailIssues = [];
  rowItems.forEach(item => {
    const email = String(item.values[emailCol] || "").trim();
    if (!email) return;
    const issue = getHubSpotImportEmailValidationIssue_(email);
    if (issue) {
      emailIssues.push(`Row ${item.row1}: "${email}" ${issue}`);
    }
  });

  if (emailIssues.length > 0) {
    ui.alert("❌ Email validation failed:\n\n" + emailIssues.slice(0, 15).join("\n"));
    return;
  }

  const activeColIndexes = [];
  header.forEach((h, idx) => {
    if (!h) return;
    if (HUBSPOT_IMPORT_EXCLUDED_COLS_.has(h)) return;
    activeColIndexes.push(idx);
  });

  const activeHeaders = activeColIndexes.map(i => header[i]);
  const activeRows = rowItems.map(item =>
    activeColIndexes.map(i => String(item.values[i] != null ? item.values[i] : ""))
  );

  const payload = {
    sheetName: "Creator List",
    spreadsheetName: ss.getName().replace(/[\\/:*?"<>|]+/g, "-"),
    spreadsheetLocale: ss.getSpreadsheetLocale(),
    headers: activeHeaders,
    rows: activeRows,
    sourceRowNumbers: rowItems.map(item => item.row1),
    rowCount: activeRows.length,
    columnCount: activeHeaders.length
  };

  // Try shared library
  if (hasHubSpotSharedImporterLibrary_()) {
    try {
      const result = HubSpotSharedImporter.startImport(payload);
      if (result && result.ok) {
        const savedRecordIds = saveImportedHubSpotRecordIds_(sheet, header, result.dealRecordIds);
        ui.alert(buildImportSuccessMessage_(payload.rowCount, result.importId, result.state, savedRecordIds));
        return;
      }
      throw new Error(result && result.error ? result.error : "Library import failed.");
    } catch (e) {
      ui.alert("❌ HubSpot import failed: " + e.message);
      return;
    }
  }

  // Try shared web app
  const webAppUrl = String(HUBSPOT_SHARED_IMPORT_WEB_APP_URL_ || "").trim();
  if (webAppUrl) {
    try {
      const response = UrlFetchApp.fetch(webAppUrl, {
        method: "post",
        muteHttpExceptions: true,
        followRedirects: false,
        contentType: "application/json",
        headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
        payload: JSON.stringify(Object.assign({ action: HUBSPOT_SHARED_IMPORT_ACTION_ }, payload))
      });

      const code = response.getResponseCode();
      const text = String(response.getContentText() || "");

      if (code === 302 || code === 401 || code === 403) {
        throw new Error("Shared HubSpot importer web app denied access.");
      }
      if (code >= 400) throw new Error("Web app error " + code + ": " + text.slice(0, 500));

      const parsed = JSON.parse(text);
      if (!parsed || parsed.ok !== true) {
        throw new Error(parsed && parsed.error ? parsed.error : "Web app import failed.");
      }

      const savedRecordIds = saveImportedHubSpotRecordIds_(sheet, header, parsed.dealRecordIds);
      ui.alert(buildImportSuccessMessage_(payload.rowCount, parsed.importId, parsed.state, savedRecordIds));
      return;
    } catch (e) {
      ui.alert("❌ HubSpot import failed: " + e.message);
      return;
    }
  }

  ui.alert(
    "❌ No HubSpot importer configured.\n\n" +
    "Recommended: add the shared Apps Script library with identifier " +
    HUBSPOT_SHARED_LIBRARY_IDENTIFIER_ +
    ", or set HUBSPOT_SHARED_IMPORT_WEB_APP_URL_ in this file."
  );
}


function saveImportedHubSpotRecordIds_(sheet, header, dealRecordIds) {
  const recordIdCol = header.indexOf("HubSpot Record ID");
  if (recordIdCol === -1 || !Array.isArray(dealRecordIds) || dealRecordIds.length === 0) {
    return 0;
  }

  let saved = 0;
  dealRecordIds.forEach(item => {
    const row1 = Number(item && item.sourceRowNumber);
    const recordId = String(item && item.recordId || "").trim();
    if (!isFinite(row1) || row1 < 2 || !recordId) return;

    const cell = sheet.getRange(row1, recordIdCol + 1);
    if (String(cell.getValue() || "").trim() === recordId) return;

    cell.setValue(recordId);
    saved++;
  });

  return saved;
}

function buildImportSuccessMessage_(rowCount, importId, state, savedRecordIds) {
  const importState = String(state || "STARTED").trim() || "STARTED";
  const suffix = importState === "DONE"
    ? "HubSpot finished the import and the returned deal IDs were saved."
    : "HubSpot continues processing the import in the background.";

  return (
    "✅ HubSpot import submitted.\n\n" +
    "Rows: " + rowCount + "\n" +
    "Import ID: " + (importId || "N/A") + "\n" +
    "State: " + importState + "\n" +
    "HubSpot Record IDs saved: " + Number(savedRecordIds || 0) + "\n\n" +
    suffix
  );
}


// ============================================================
//  (2) CREATOR LIST → PITCHING (NEGOTIATION)
// ============================================================

function pushRespondedToNegotiation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const creatorSheet = ss.getSheetByName("Creator List");
  const pitchingSheet = ss.getSheetByName("Pitching");
  if (!creatorSheet || !pitchingSheet) {
    return Logger.log("❌ Missing 'Creator List' or 'Pitching' sheet.");
  }

  const creatorData = creatorSheet.getDataRange().getValues();
  const pitchingData = pitchingSheet.getDataRange().getValues();
  if (creatorData.length < 2 || pitchingData.length < 1) {
    return Logger.log("ℹ️ One of the sheets has no data.");
  }

  const creatorHeader = creatorData[0].map(v => String(v || "").trim());
  const pitchingHeader = pitchingData[0].map(v => String(v || "").trim());

  const statusCol = creatorHeader.indexOf("Status");
  if (statusCol === -1) return Logger.log("❌ Creator List missing 'Status' column.");

  const contactingStart0 = findSectionRowByLabel_(creatorData, "Contacting");
  const creatorArchived0 = findSectionRowByLabel_(creatorData, "Archived");
  if (contactingStart0 === -1) return Logger.log("❌ 'Contacting' section not found in Creator List.");
  if (creatorArchived0 === -1) return Logger.log("❌ 'Archived' section not found in Creator List.");

  const negotiationStart0 = findSectionRowByLabel_(pitchingData, "Negotiation");
  if (negotiationStart0 === -1) return Logger.log("❌ 'Negotiation' section not found in Pitching.");

  // Find responded rows
  const respondedRows = [];
  for (let r = contactingStart0 + 1; r < creatorArchived0; r++) {
    const row = creatorData[r];
    if (isBlankRow_(row) || isSectionLabelRow_(row)) continue;
    if (String(row[statusCol] || "").trim() !== "Responded") continue;
    respondedRows.push({ row1: r + 1, values: row.slice() });
  }

  if (respondedRows.length === 0) {
    return Logger.log("ℹ️ No rows with Status 'Responded' in Contacting section.");
  }

  // Map Creator List columns → Pitching columns, excluding Status and
  // mapping YouTube Video Median Views → Median Views.
  const colMap = buildCrossSheetColumnMap_(
    creatorHeader,
    pitchingHeader,
    CREATOR_TO_PITCHING_NEGOTIATION_MAP_,
    CREATOR_TO_PITCHING_NEGOTIATION_SKIP_COLS_
  );
  const pitchingNumCols = pitchingHeader.length;
  const existingPitchingCompositeKeys = buildCompositeKeySet_(pitchingData, pitchingHeader);
  const existingPitchingSignatures = buildMappedRowSignatureSetFromTarget_(pitchingData, colMap);

  // Build Pitching rows
  const pitchingRows = [];
  let skippedDup = 0;
  respondedRows.forEach(item => {
    const compositeKey = buildCompositeKey_(item.values, creatorHeader);
    const signature = buildMappedRowSignatureFromSource_(item.values, colMap);
    if (
      (compositeKey && existingPitchingCompositeKeys.has(compositeKey)) ||
      (signature && existingPitchingSignatures.has(signature))
    ) {
      if (compositeKey) existingPitchingCompositeKeys.add(compositeKey);
      skippedDup++;
      return;
    }

    const out = new Array(pitchingNumCols).fill("");
    for (const m of colMap) out[m.targetIdx] = item.values[m.sourceIdx];
    pitchingRows.push(out);

    if (compositeKey) existingPitchingCompositeKeys.add(compositeKey);
    if (signature) existingPitchingSignatures.add(signature);
  });

  if (pitchingRows.length === 0) {
    return Logger.log(`ℹ️ No new rows copied to Pitching Negotiation. ${skippedDup} duplicate(s) skipped.`);
  }

  // Insert into Negotiation section
  let insertAt1 = findFirstEmptyRowInSection1_(pitchingData, negotiationStart0);
  if (insertAt1 === -1) insertAt1 = negotiationStart0 + 2;

  if (pitchingRows.length > 1) {
    pitchingSheet.insertRowsBefore(insertAt1, pitchingRows.length);
  } else {
    pitchingSheet.insertRowBefore(insertAt1);
  }

  const writeRange = pitchingSheet.getRange(insertAt1, 1, pitchingRows.length, pitchingNumCols);
  const templateRange = getSafeFormatRow_(pitchingSheet, insertAt1 + pitchingRows.length, pitchingNumCols);
  templateRange.copyTo(writeRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  templateRange.copyTo(writeRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  writeRange.setValues(pitchingRows);

  Logger.log(`✅ Copied ${pitchingRows.length} row(s) into Pitching Negotiation section. ${skippedDup} duplicate(s) skipped.`);
}


function moveCreatorRowsToArchived_(sheet, rowItems, header) {
  const numCols = header.length;

  // Delete from bottom up to preserve row numbers
  const sorted = rowItems.slice().sort((a, b) => b.row1 - a.row1);
  for (const item of sorted) {
    sheet.deleteRow(item.row1);
  }

  // Re-read to find Archived section after deletions
  const refreshed = sheet.getDataRange().getValues();
  const newArchived0 = findSectionRowByLabel_(refreshed, "Archived");
  if (newArchived0 === -1) {
    Logger.log("❌ 'Archived' section disappeared after deletion.");
    return;
  }

  // Insert rows into Archived section
  let insertAt1 = findFirstEmptyRowInSection1_(refreshed, newArchived0);
  if (insertAt1 === -1) insertAt1 = newArchived0 + 2;

  if (rowItems.length > 1) {
    sheet.insertRowsBefore(insertAt1, rowItems.length);
  } else {
    sheet.insertRowBefore(insertAt1);
  }

  const writeRange = sheet.getRange(insertAt1, 1, rowItems.length, numCols);
  const templateRange = getSafeFormatRow_(sheet, insertAt1 + rowItems.length, numCols);
  templateRange.copyTo(writeRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  templateRange.copyTo(writeRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  writeRange.setValues(rowItems.map(item => item.values));

  // Restore blank rows in Contacting section
  const refreshed2 = sheet.getDataRange().getValues();
  const contactingStart0 = findSectionRowByLabel_(refreshed2, "Contacting");
  if (contactingStart0 !== -1) {
    const restoreAt1 = contactingStart0 + 2;
    if (rowItems.length > 1) {
      sheet.insertRowsBefore(restoreAt1, rowItems.length);
    } else {
      sheet.insertRowBefore(restoreAt1);
    }
    const fmtSource = getSafeFormatRow_(sheet, restoreAt1 + rowItems.length, numCols);
    const restoreRange = sheet.getRange(restoreAt1, 1, rowItems.length, numCols);
    fmtSource.copyTo(restoreRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    fmtSource.copyTo(restoreRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
  }

  Logger.log(`✅ Moved ${rowItems.length} row(s) to Creator List Archived section.`);
}


// ============================================================
//  (3) INT PITCHING → EXT PITCHING
// ============================================================

function pushReadyForPitchingToExt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const intPitching = ss.getSheetByName("Pitching");
  if (!intPitching) return Logger.log("❌ 'Pitching' sheet not found.");

  const extSs = getExtSpreadsheet_();
  if (!extSs) return;
  const extPitching = extSs.getSheetByName("Pitching");
  if (!extPitching) return Logger.log("❌ 'Pitching' sheet not found in EXT spreadsheet.");

  const intData = intPitching.getDataRange().getValues();
  const extData = extPitching.getDataRange().getValues();
  if (intData.length < 2 || extData.length < 1) {
    return Logger.log("ℹ️ One of the sheets has no data.");
  }

  const intHeader = intData[0].map(v => String(v || "").trim());
  const extHeader = extData[0].map(v => String(v || "").trim());

  const intStatusCol = intHeader.indexOf("Status");
  if (intStatusCol === -1) return Logger.log("❌ INT Pitching missing 'Status' column.");

  // Find rows above Archived with Status = "Ready for pitching"
  const archivedStart0 = findSectionRowByLabel_(intData, "Archived");
  const endRow = archivedStart0 === -1 ? intData.length : archivedStart0;

  const candidates = [];
  for (let r = 1; r < endRow; r++) {
    const row = intData[r];
    if (isBlankRow_(row) || isSectionLabelRow_(row)) continue;
    if (String(row[intStatusCol] || "").trim() !== "Ready for pitching") continue;
    candidates.push({ row1: r + 1, values: row });
  }

  if (candidates.length === 0) {
    return Logger.log("ℹ️ No rows with Status 'Ready for pitching'.");
  }

  // Column mapping: INT → EXT (rename EXT Rate→Rate, EXT CPM→CPM; skip INT Rate, INT CPM)
  const colMap = buildCrossSheetColumnMap_(intHeader, extHeader, INT_TO_EXT_PITCHING_MAP_, INT_ONLY_PITCHING_COLS_);

  // Duplicate detection via composite key
  const extCompositeKeys = buildCompositeKeySet_(extData, extHeader);

  // EXT Active Pitches section
  const extActiveStart0 = findSectionRowByLabel_(extData, "Active Pitches");
  if (extActiveStart0 === -1) return Logger.log("❌ 'Active Pitches' section not found in EXT Pitching.");

  const extNumCols = extHeader.length;
  const extStatusCol = extHeader.indexOf("Status");
  let inserted = 0;
  let skippedDup = 0;

  for (const item of candidates) {
    const key = buildCompositeKey_(item.values, intHeader);
    if (key && extCompositeKeys.has(key)) { skippedDup++; continue; }

    const out = new Array(extNumCols).fill("");
    for (const m of colMap) out[m.targetIdx] = item.values[m.sourceIdx];

    // Override Status to "Pitched" in EXT
    if (extStatusCol !== -1) out[extStatusCol] = "Pitched";

    // Find insertion point (re-read each time because rows shift)
    const extLive = extPitching.getDataRange().getValues();
    const liveActiveStart0 = findSectionRowByLabel_(extLive, "Active Pitches");
    let insertAt1 = findFirstEmptyRowInSection1_(extLive, liveActiveStart0);
    if (insertAt1 === -1) {
      const extArchived0 = findSectionRowByLabel_(extLive, "Archived");
      insertAt1 = extArchived0 !== -1 ? extArchived0 + 1 : extLive.length + 1;
    }

    extPitching.insertRowBefore(insertAt1);
    const targetRange = extPitching.getRange(insertAt1, 1, 1, extNumCols);
    const templateRange = getSafeFormatRow_(extPitching, insertAt1 + 1, extNumCols);
    templateRange.copyTo(targetRange, { formatOnly: true });
    targetRange.setValues([out]);

    if (key) extCompositeKeys.add(key);
    inserted++;
  }

  // Update INT Status from "Ready for pitching" → "Pitched"
  if (inserted > 0) {
    for (const item of candidates) {
      if (String(item.values[intStatusCol] || "").trim() === "Ready for pitching") {
        intPitching.getRange(item.row1, intStatusCol + 1).setValue("Pitched");
      }
    }
  }

  Logger.log(`✅ Pushed to EXT Pitching: ${inserted} inserted, ${skippedDup} skipped (duplicates).`);
}


// ============================================================
//  (4) UPDATE INT PITCHING FROM EXT
// ============================================================

function updatePitchingFromExt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const intPitching = ss.getSheetByName("Pitching");
  if (!intPitching) return Logger.log("❌ 'Pitching' sheet not found.");

  const extSs = getExtSpreadsheet_();
  if (!extSs) return;
  const extPitching = extSs.getSheetByName("Pitching");
  if (!extPitching) return Logger.log("❌ 'Pitching' sheet not found in EXT spreadsheet.");

  const intData = intPitching.getDataRange().getValues();
  const extData = extPitching.getDataRange().getValues();
  if (intData.length < 2 || extData.length < 2) {
    return Logger.log("ℹ️ One of the sheets has insufficient data.");
  }

  const intHeader = intData[0].map(v => String(v || "").trim());
  const extHeader = extData[0].map(v => String(v || "").trim());

  // Build INT row lookup by composite key
  const intRowMap = new Map();
  for (let r = 1; r < intData.length; r++) {
    const row = intData[r];
    if (isBlankRow_(row) || isSectionLabelRow_(row)) continue;
    const key = buildCompositeKey_(row, intHeader);
    if (key && !intRowMap.has(key)) intRowMap.set(key, r);
  }

  // Column mapping: EXT → INT (Rate→EXT Rate, CPM→EXT CPM)
  const colMap = buildCrossSheetColumnMap_(extHeader, intHeader, EXT_TO_INT_PITCHING_MAP_, new Set());

  let updated = 0;
  for (let r = 1; r < extData.length; r++) {
    const extRow = extData[r];
    if (isBlankRow_(extRow) || isSectionLabelRow_(extRow)) continue;

    const key = buildCompositeKey_(extRow, extHeader);
    if (!key) continue;

    const intRowIdx = intRowMap.get(key);
    if (intRowIdx === undefined) continue;

    const intRow = intData[intRowIdx];
    let changed = false;

    for (const m of colMap) {
      const extVal = extRow[m.sourceIdx];
      if (String(extVal || "").trim() === "") continue;
      if (String(intRow[m.targetIdx] || "").trim() !== String(extVal || "").trim()) {
        intRow[m.targetIdx] = extVal;
        changed = true;
      }
    }

    if (changed) {
      intPitching.getRange(intRowIdx + 1, 1, 1, intHeader.length).setValues([intRow]);
      updated++;
    }
  }

  Logger.log(`✅ Updated ${updated} row(s) in INT Pitching from EXT.`);
}


// ============================================================
//  (5) UPDATE INT CAMPAIGNS FROM EXT
// ============================================================

function updateCampaignsFromExt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const intCampaigns = ss.getSheetByName("Campaigns");
  if (!intCampaigns) return Logger.log("❌ 'Campaigns' sheet not found.");

  const extSs = getExtSpreadsheet_();
  if (!extSs) return;
  const extCampaigns = extSs.getSheetByName("Campaigns");
  if (!extCampaigns) return Logger.log("❌ 'Campaigns' sheet not found in EXT spreadsheet.");

  const intData = intCampaigns.getDataRange().getValues();
  const extData = extCampaigns.getDataRange().getValues();
  if (intData.length < 2 || extData.length < 2) {
    return Logger.log("ℹ️ One of the sheets has insufficient data.");
  }

  const intHeader = intData[0].map(v => String(v || "").trim());
  const extHeader = extData[0].map(v => String(v || "").trim());

  // Build INT row lookup by composite key
  const intRowMap = new Map();
  for (let r = 1; r < intData.length; r++) {
    const row = intData[r];
    if (isBlankRow_(row) || isSectionLabelRow_(row)) continue;
    const key = buildCompositeKey_(row, intHeader);
    if (key && !intRowMap.has(key)) intRowMap.set(key, r);
  }

  // Column mapping: EXT → INT (Rate→EXT Rate)
  const colMap = buildCrossSheetColumnMap_(extHeader, intHeader, EXT_TO_INT_CAMPAIGNS_MAP_, new Set());

  let updated = 0;
  let inserted = 0;

  for (let r = 1; r < extData.length; r++) {
    const extRow = extData[r];
    if (isBlankRow_(extRow) || isSectionLabelRow_(extRow)) continue;

    const key = buildCompositeKey_(extRow, extHeader);
    if (!key) continue;

    const intRowIdx = intRowMap.get(key);
    if (intRowIdx !== undefined) {
      // Update existing row
      const intRow = intData[intRowIdx];
      let changed = false;

      for (const m of colMap) {
        const extVal = extRow[m.sourceIdx];
        if (String(extVal || "").trim() === "") continue;
        if (String(intRow[m.targetIdx] || "").trim() !== String(extVal || "").trim()) {
          intRow[m.targetIdx] = extVal;
          changed = true;
        }
      }

      if (changed) {
        intCampaigns.getRange(intRowIdx + 1, 1, 1, intHeader.length).setValues([intRow]);
        updated++;
      }
    } else {
      // Insert new row into the matching month section
      const monthSection = findRowSection_(extData, r);
      if (!monthSection) continue;

      const intLive = intCampaigns.getDataRange().getValues();
      const sectionStart0 = findSectionRowByLabel_(intLive, monthSection);
      if (sectionStart0 === -1) continue;

      let insertAt1 = findFirstEmptyRowInSection1_(intLive, sectionStart0);
      if (insertAt1 === -1) continue;

      const intNumCols = intHeader.length;
      const out = new Array(intNumCols).fill("");
      for (const m of colMap) out[m.targetIdx] = extRow[m.sourceIdx];

      intCampaigns.insertRowBefore(insertAt1);
      const targetRange = intCampaigns.getRange(insertAt1, 1, 1, intNumCols);
      const templateRange = getSafeFormatRow_(intCampaigns, insertAt1 + 1, intNumCols);
      templateRange.copyTo(targetRange, { formatOnly: true });
      targetRange.setValues([out]);

      inserted++;
    }
  }

  Logger.log(`✅ INT Campaigns: ${updated} updated, ${inserted} inserted from EXT.`);
}


// ============================================================
//  FILL MEDIAN VIEWS (Pitching)
// ============================================================

function fillMissingMedianViews_Pitching() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("Pitching");
  if (!sh) return Logger.log("❌ 'Pitching' sheet not found.");

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return Logger.log("ℹ️ Pitching has no data.");

  const values = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const header = (values[0] || []).map(h => String(h || "").trim());

  const urlCol = header.indexOf("Channel URL");
  const medianCol = header.indexOf("Median Views");
  if (urlCol === -1 || medianCol === -1) {
    return Logger.log("❌ Missing 'Channel URL' or 'Median Views' column in Pitching.");
  }

  // Stop at Archived section
  let stopAt = lastRow + 1;
  for (let r = 2; r <= lastRow; r++) {
    if (String(values[r - 1][0] || "").trim() === "Archived") { stopAt = r; break; }
  }

  const updates = [];
  for (let r = 2; r < stopAt; r++) {
    const row = values[r - 1];
    if (isBlankRow_(row) || isSectionLabelRow_(row)) continue;

    const url = String(row[urlCol] || "").trim();
    if (!url) continue;
    if (String(row[medianCol] || "").trim() !== "") continue;

    const median = computeMedianViewsForChannel_(url);
    if (median != null) updates.push({ row1: r, value: median });
  }

  if (updates.length === 0) return Logger.log("ℹ️ No missing Median Views to fill.");

  updates.forEach(u => sh.getRange(u.row1, medianCol + 1).setValue(u.value));
  Logger.log(`✅ Filled Median Views for ${updates.length} row(s).`);
}


// ============================================================
//  CROSS-SPREADSHEET HELPERS
// ============================================================

function getExtSpreadsheet_() {
  const props = PropertiesService.getScriptProperties();
  let extId = String(props.getProperty(EXT_SPREADSHEET_ID_PROP_) || "").trim();

  if (!extId) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt(
      "EXT Spreadsheet",
      "Paste the EXT spreadsheet ID or URL:",
      ui.ButtonSet.OK_CANCEL
    );
    if (response.getSelectedButton() !== ui.Button.OK) return null;

    extId = String(response.getResponseText() || "").trim();
    // Extract ID from URL if needed
    const urlMatch = extId.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
    if (urlMatch) extId = urlMatch[1];

    if (!extId) {
      ui.alert("EXT spreadsheet ID is required.");
      return null;
    }
    props.setProperty(EXT_SPREADSHEET_ID_PROP_, extId);
  }

  try {
    return SpreadsheetApp.openById(extId);
  } catch (e) {
    Logger.log("❌ Could not open EXT spreadsheet: " + e);
    SpreadsheetApp.getUi().alert("❌ Could not open EXT spreadsheet. Check the ID and permissions.");
    return null;
  }
}


function buildCrossSheetColumnMap_(sourceHeader, targetHeader, renameMap, skipCols) {
  const map = [];
  const targetMap = new Map();
  for (let i = 0; i < targetHeader.length; i++) {
    const name = targetHeader[i];
    if (name) targetMap.set(name, i);
  }

  for (let i = 0; i < sourceHeader.length; i++) {
    const sourceName = sourceHeader[i];
    if (!sourceName) continue;
    if (skipCols && skipCols.has(sourceName)) continue;

    const targetName = renameMap[sourceName] || sourceName;
    const targetIdx = targetMap.get(targetName);
    if (targetIdx !== undefined) {
      map.push({ sourceIdx: i, targetIdx: targetIdx, sourceName: sourceName, targetName: targetName });
    }
  }

  return map;
}


function buildCompositeKey_(row, header) {
  const keys = [
    getValueByHeader_(row, header, "HubSpot Record ID"),
    getValueByHeader_(row, header, "Deal Type"),
    getValueByHeader_(row, header, "Activation Type")
  ].map(v => String(v || "").trim());

  if (!keys[0]) return "";
  return keys.join("||");
}


function buildCompositeKeySet_(data, header) {
  const set = new Set();
  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (isBlankRow_(row) || isSectionLabelRow_(row)) continue;
    const key = buildCompositeKey_(row, header);
    if (key) set.add(key);
  }
  return set;
}


function findRowSection_(data, rowIdx) {
  for (let r = rowIdx - 1; r >= 0; r--) {
    if (isSectionLabelRow_(data[r])) return String(data[r][0] || "").trim();
  }
  return "";
}


function getValueByHeader_(row, header, columnName) {
  const idx = header.indexOf(columnName);
  return idx === -1 ? "" : row[idx];
}

function buildMappedRowSignatureFromSource_(row, colMap) {
  if (!row || !colMap || colMap.length === 0) return "";
  return colMap
    .map(m => m.targetIdx + "=" + String(row[m.sourceIdx] == null ? "" : row[m.sourceIdx]).trim())
    .join("|");
}

function buildMappedRowSignatureSetFromTarget_(data, colMap) {
  const set = new Set();
  if (!Array.isArray(data) || !colMap || colMap.length === 0) return set;

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (isBlankRow_(row) || isSectionLabelRow_(row)) continue;

    const signature = colMap
      .map(m => m.targetIdx + "=" + String(row[m.targetIdx] == null ? "" : row[m.targetIdx]).trim())
      .join("|");

    if (signature) set.add(signature);
  }

  return set;
}


// ============================================================
//  SHARED HELPERS
// ============================================================

function getSafeFormatRow_(sheet, preferredRow, numCols) {
  const lastRow = Math.max(sheet.getLastRow(), 1);
  const row = (preferredRow >= 1 && preferredRow <= lastRow) ? preferredRow : lastRow;
  return sheet.getRange(row, 1, 1, numCols);
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
    if (sourceMap.has(name)) common.push({ name: name, pIdx: sourceMap.get(name), cIdx: j });
  }
  return common;
}

function setIfEmpty_(row, header, columnName, value) {
  const idx = header.indexOf(columnName);
  if (idx === -1) return false;
  if (String(row[idx] || "").trim() !== "") return false;
  row[idx] = value;
  return true;
}

function getDefaultClientName_(dropdownValuesByHeader) {
  const clientValues = dropdownValuesByHeader["Client"] || dropdownValuesByHeader["Client name"] || [];
  return clientValues.length > 0 ? String(clientValues[0] || "").trim() : "";
}

function getDropdownValuesByHeader_(ss, sheetName) {
  const dropdownSheet = ss.getSheetByName(sheetName || "Dropdown Values");
  if (!dropdownSheet) return {};

  const data = dropdownSheet.getDataRange().getDisplayValues();
  if (data.length === 0) return {};

  const header = (data[0] || []).map(value => String(value || "").trim());
  const out = {};

  for (let c = 0; c < header.length; c++) {
    const columnName = header[c];
    if (!columnName) continue;

    const values = [];
    const seen = new Set();
    for (let r = 1; r < data.length; r++) {
      const value = String((data[r] && data[r][c]) || "").trim();
      if (!value || seen.has(value)) continue;
      seen.add(value);
      values.push(value);
    }

    out[columnName] = values;
  }

  return out;
}

function chunkArray_(arr, size) {
  const out = [];
  for (let i = 0; i < arr.length; i += size) {
    out.push(arr.slice(i, i + size));
  }
  return out;
}


// ============================================================
//  YOUTUBE API
// ============================================================

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
      if (isInvalidApiKeyResponse_(text)) markApiKeyInvalid_();
      else Logger.log("❌ YouTube API key validation failed: HTTP " + code);
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

function computeMedianViewsFromYouTube_(channelUrl, apiKey) {
  try {
    if (YT_API_KEY_INVALID_) return null;

    const channelId = getChannelIdFromUrl(channelUrl, apiKey);
    if (!channelId) {
      Logger.log("resolve failed for URL: " + channelUrl);
      return null;
    }

    const channelData = ytFetchJson_(
      "https://www.googleapis.com/youtube/v3/channels?part=contentDetails" +
      "&id=" + encodeURIComponent(channelId) +
      "&key=" + encodeURIComponent(apiKey)
    );
    if (!channelData.items || channelData.items.length === 0) return null;

    const uploadsPlaylistId = channelData.items[0].contentDetails.relatedPlaylists.uploads;
    if (!uploadsPlaylistId) return null;

    const playlistData = ytFetchJson_(
      "https://www.googleapis.com/youtube/v3/playlistItems?part=snippet" +
      "&maxResults=" + YT_MAX_PLAYLIST_ITEMS_ +
      "&playlistId=" + encodeURIComponent(uploadsPlaylistId) +
      "&key=" + encodeURIComponent(apiKey)
    );
    if (!playlistData.items || playlistData.items.length === 0) return null;

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

    if (candidateIds.length === 0) return null;

    const ids = candidateIds.slice(0, YT_MAX_PLAYLIST_ITEMS_);
    const videosData = ytFetchJson_(
      "https://www.googleapis.com/youtube/v3/videos?part=statistics,contentDetails" +
      "&id=" + encodeURIComponent(ids.join(",")) +
      "&key=" + encodeURIComponent(apiKey)
    );
    if (!videosData.items || videosData.items.length === 0) return null;

    const pickedViews = [];
    for (const video of videosData.items) {
      const duration = video.contentDetails && video.contentDetails.duration ? video.contentDetails.duration : null;
      if (iso8601DurationToSeconds_(duration) <= YT_MIN_DURATION_SECONDS_) continue;
      const rawViews = video.statistics && video.statistics.viewCount ? Number(video.statistics.viewCount) : null;
      if (rawViews == null || isNaN(rawViews)) continue;
      pickedViews.push(rawViews);
      if (pickedViews.length >= YT_MAX_VIDEOS_FOR_AVG_) break;
    }

    if (pickedViews.length === 0) return null;
    return computeMedian_(pickedViews);
  } catch (e) {
    Logger.log("computeMedianViewsFromYouTube_ error: " + (e && e.stack ? e.stack : e));
    return null;
  }
}


// ---- Channel ID resolution ----

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
  return (item.id && item.id.channelId)
    ? item.id.channelId
    : ((item.snippet && item.snippet.channelId) || null);
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
    } catch (e) { /* continue */ }
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
    if (isInvalidApiKeyResponse_(text)) { markApiKeyInvalid_(); return {}; }
    Logger.log("YouTube API error " + code + ": " + text.slice(0, 500));
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
  return (Number(match[1] || 0) * 3600) + (Number(match[2] || 0) * 60) + Number(match[3] || 0);
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


// ============================================================
//  EMAIL VALIDATION (for HubSpot import)
// ============================================================

function getHubSpotImportEmailValidationIssue_(value) {
  if (!value) return "is blank.";
  if (value.length > 254) return "is longer than 254 characters.";
  if (/[^\x00-\x7F]/.test(value)) return "contains non-ASCII characters.";
  if (/\s/.test(value)) return "contains whitespace.";

  const atCount = (value.match(/@/g) || []).length;
  if (atCount !== 1) return "must contain exactly one @ symbol.";

  const parts = value.split("@");
  const local = parts[0];
  const domain = parts[1];

  if (!local) return "is missing the part before @.";
  if (!domain) return "is missing the domain after @.";
  if (local.length > 64) return "has more than 64 characters before @.";
  if (local.startsWith(".") || local.endsWith(".")) return "has a dot at the start or end before @.";
  if (local.indexOf("..") !== -1) return "has consecutive dots before @.";
  if (!/^[A-Za-z0-9!#$%&'*+/=?^_`{|}~.-]+$/.test(local)) {
    return "contains unsupported characters before @.";
  }

  if (domain.length > 253) return "has a domain longer than 253 characters.";
  if (domain.indexOf(".") === -1) return "must include a full domain such as domain.com.";
  if (domain.startsWith(".") || domain.endsWith(".")) return "has a domain that starts or ends with a dot.";
  if (domain.indexOf("..") !== -1) return "has consecutive dots in the domain.";

  const labels = domain.split(".");
  if (labels.some(function (label) { return !label; })) return "has an empty section in the domain.";

  for (let i = 0; i < labels.length; i++) {
    const label = labels[i];
    if (label.length > 63) return "has a domain section longer than 63 characters.";
    if (!/^[A-Za-z0-9-]+$/.test(label)) return "contains unsupported characters in the domain.";
    if (label.startsWith("-") || label.endsWith("-")) {
      return "has a domain section that starts or ends with a hyphen.";
    }
  }

  const topLevelDomain = labels[labels.length - 1];
  if (!/^(xn--[A-Za-z0-9-]{2,59}|[A-Za-z]{2,63})$/.test(topLevelDomain)) {
    return "has an invalid top-level domain.";
  }

  return "";
}


// ============================================================
//  HUBSPOT SHARED LIBRARY CHECK
// ============================================================

function hasHubSpotSharedImporterLibrary_() {
  return (
    typeof HubSpotSharedImporter !== "undefined" &&
    HubSpotSharedImporter &&
    typeof HubSpotSharedImporter.startImport === "function"
  );
}
