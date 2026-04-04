/*********************************************
 * INT Sheet Template – Code.gs
 *
 * Internal workflow scripts:
 * 1) Creator List → HubSpot (+ enrichment)
 * 2) Creator List → Pitching (Negotiation)
 * 3) INT Pitching → EXT Pitching
 * 4) Update INT Pitching from EXT
 * 5) Update INT Campaigns from EXT
 * 6) Update INT Performance from EXT
 *********************************************/


// ============================================================
//  CONSTANTS
// ============================================================

const EXT_SPREADSHEET_ID_PROP_ = "EXT_SPREADSHEET_ID";
const EXT_SHEET_DROPDOWN_HEADER_ = "EXT Sheet";

const HUBSPOT_SHARED_IMPORT_WEB_APP_URL_ =
  "https://script.google.com/a/macros/arch.agency/s/AKfycbzI6gAHnhSlRLzheWkS_wNYvYODvx1aztd27cf2DbbuBJTSqOYe-oqtKAnZqRc7jCE8/exec";
const HUBSPOT_SHARED_IMPORT_ACTION_ = "startImport";
const HUBSPOT_SHARED_LIBRARY_IDENTIFIER_ = "HubSpotSharedImporter";

const OPENAI_API_KEY_PROP_ = "OPENAI_API_KEY";
const OPENAI_MODEL_PROP_ = "OPENAI_MODEL";
const OPENAI_DEFAULT_MODEL_ = "gpt-5-nano";
const OPENAI_CREATOR_PROFILE_BATCH_SIZE_ = 5;
const CREATOR_LIST_ENRICH_BATCH_SIZE_ = 20;
const CREATOR_LIST_ENRICH_CURSOR_PROP_ = "CREATOR_LIST_ENRICH_CURSOR";
const CREATOR_YOUTUBE_INSIGHT_CACHE_PREFIX_ = "CREATOR_YOUTUBE_INSIGHT_V2::";
const CREATOR_LONGFORM_SIGNAL_CACHE_PREFIX_ = "CREATOR_LONGFORM_SIGNAL_V2::";
const CREATOR_CHANNEL_PAGE_SIGNAL_CACHE_PREFIX_ = "CREATOR_CHANNEL_PAGE_SIGNAL_V1::";
const CREATOR_YOUTUBE_SIGNAL_TTL_SECONDS_ = 21600;
const CREATOR_YOUTUBE_SIGNAL_SAMPLE_SIZE_ = 12;
const CREATOR_YOUTUBE_SIGNAL_MIN_LONGFORM_SECONDS_ = 60;
const CREATOR_YOUTUBE_DESCRIPTION_SAMPLE_SIZE_ = 6;
const SHEETS_MULTISELECT_DELIMITER_ = ", ";
const HUBSPOT_MULTISELECT_DELIMITER_ = ";";
const HUBSPOT_MULTISELECT_FIELDS_ = new Set(["Influencer Vertical"]);

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
const YT_SHORTS_MAX_DURATION_SECONDS_ = 60;
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
var LAST_OPENAI_DIAGNOSTIC_ = "";
var LAST_OPENAI_RAW_RESPONSE_ = "";

// INT ↔ EXT column name mappings
const INT_TO_EXT_PITCHING_MAP_ = { "EXT Rate": "Rate" };
const EXT_TO_INT_PITCHING_MAP_ = {};
const EXT_TO_INT_CAMPAIGNS_MAP_ = { "Rate": "EXT Rate" };
const CREATOR_TO_PITCHING_NEGOTIATION_MAP_ = {
  "YouTube Average Views": "Median Views",
  "YouTube Video Median Views": "Median Views"
};
const INT_ONLY_PITCHING_COLS_ = new Set(["INT Rate", "INT CPM", "EXT CPM"]);
const EXT_ONLY_PITCHING_COLS_ = new Set(["Rate", "CPM"]);
const INT_PITCHING_FORMULA_COLS_ = new Set(["INT CPM", "EXT CPM"]);
const PITCHING_NEGOTIATION_BLOCK_WIDTH_ = 5;
const CREATOR_TO_PITCHING_NEGOTIATION_SKIP_COLS_ = new Set(["Status"]);
const HUBSPOT_IMPORT_EXCLUDED_COLS_ = new Set(["Channel Name", "HubSpot Record ID", "Channel URL", "Status"]);
const YOUTUBE_ENRICH_FIELDS_ = [
  "YouTube Handle",
  "YouTube URL",
  "YouTube Average Views",
  "YouTube Video Median Views",
  "YouTube Shorts Median Views",
  "YouTube Engagement Rate",
  "YouTube Followers"
];
const PROFILE_LLM_FIELDS_ = [
  "First Name",
  "Last Name",
  "Email",
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
  "YouTube Average Views",
  "YouTube Video Median Views",
  "YouTube Shorts Median Views",
  "YouTube Engagement Rate",
  "YouTube Followers",
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

function isKnownSectionLabel_(value) {
  const text = String(value || "").trim();
  if (!text) return false;

  const knownLabels = new Set(
    ["Contacting", "Archived", "Negotiation", "Active Pitches"].concat(MONTH_NAMES_)
  );
  return knownLabels.has(text);
}


// ============================================================
//  MENU
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🧰 Scripts")
    .addItem("Enrich Creator List", "enrichCreatorListRowsInBatches")
    .addItem("Enrich selected Creator List rows", "enrichSelectedCreatorListRows")
    .addItem("Diagnose Creator List enrichment", "diagnoseCreatorListEnrichment")
    .addItem("Diagnose selected Creator row", "diagnoseSelectedCreatorListRow")
    .addItem("Reset Creator List enrichment", "resetCreatorListEnrichmentProgress")
    .addSeparator()
    .addItem("Enrich & Import to HubSpot", "enrichAndImportToHubSpot")
    .addSeparator()
    .addItem("Responded → Negotiation", "pushRespondedToNegotiation")
    .addSeparator()
    .addItem("Push to EXT Pitching", "pushReadyForPitchingToExt")
    .addItem("Update Pitching from EXT", "updatePitchingFromExt")
    .addSeparator()
    .addItem("Update Campaigns from EXT", "updateCampaignsFromExt")
    .addSeparator()
    .addItem("Update Performance from EXT", "updatePerformanceFromExt")
    .addToUi();
}


// ============================================================
//  CREATOR LIST ENRICHMENT ONLY
// ============================================================

function enrichCreatorListRowsInBatches() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Creator List");
  if (!sheet) {
    showSpreadsheetToast_("Creator List sheet not found.");
    return Logger.log("❌ 'Creator List' sheet not found.");
  }

  const context = getCreatorListSheetContext_(sheet);
  if (!context) {
    showSpreadsheetToast_("Creator List has no data.");
    return Logger.log("ℹ️ Creator List has no data.");
  }

  const data = context.data;
  const header = context.header;
  const stopRow1 = context.archivedStart0 === -1 ? data.length : context.archivedStart0;
  const props = PropertiesService.getScriptProperties();
  const startRow1 = Math.max(
    Number(props.getProperty(CREATOR_LIST_ENRICH_CURSOR_PROP_) || (context.headerRow1 + 1)),
    context.headerRow1 + 1
  );

  const rowItems = [];
  let lastScannedRow1 = startRow1 - 1;

  for (let row1 = startRow1; row1 <= stopRow1; row1++) {
    lastScannedRow1 = row1;
    const row = data[row1 - 1];
    if (isBlankRow_(row) || isSectionLabelRow_(row)) continue;
    if (!canAttemptCreatorListEnrichment_(row, header)) continue;

    rowItems.push({ row1: row1, values: row.slice() });
    if (rowItems.length >= CREATOR_LIST_ENRICH_BATCH_SIZE_) break;
  }

  if (rowItems.length === 0) {
    props.deleteProperty(CREATOR_LIST_ENRICH_CURSOR_PROP_);
    showSpreadsheetToast_("Creator List enrichment complete. No remaining rows to process.");
    return Logger.log("✅ Creator List enrichment complete. No remaining rows to process.");
  }

  const result = runCreatorListEnrichment_(sheet, rowItems, header);
  if (lastScannedRow1 < stopRow1) {
    props.setProperty(CREATOR_LIST_ENRICH_CURSOR_PROP_, String(lastScannedRow1 + 1));
    showSpreadsheetToast_(
      `Batch done. Rows: ${rowItems.length}. Static: ${result.staticEnriched}. ` +
      `YouTube: ${result.youtubeEnriched}. AI: ${result.profileEnriched}.`
    );
    return Logger.log(
      `✅ Creator List enrichment batch complete. Rows scanned through ${lastScannedRow1}. ` +
      `Static: ${result.staticEnriched}, YouTube: ${result.youtubeEnriched}, AI: ${result.profileEnriched}.`
    );
  }

  props.deleteProperty(CREATOR_LIST_ENRICH_CURSOR_PROP_);
  showSpreadsheetToast_(
    `Enrichment complete. Rows: ${rowItems.length}. Static: ${result.staticEnriched}. ` +
    `YouTube: ${result.youtubeEnriched}. AI: ${result.profileEnriched}.`
  );
  Logger.log(
    `✅ Creator List enrichment complete. Static: ${result.staticEnriched}, ` +
    `YouTube: ${result.youtubeEnriched}, AI: ${result.profileEnriched}.`
  );
}

function enrichSelectedCreatorListRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Creator List");
  if (!sheet) {
    showSpreadsheetToast_("Creator List sheet not found.");
    return Logger.log("❌ 'Creator List' sheet not found.");
  }

  const activeSheet = ss.getActiveSheet();
  if (!activeSheet || activeSheet.getSheetId() !== sheet.getSheetId()) {
    showSpreadsheetToast_("Open Creator List and select rows first.");
    return Logger.log("ℹ️ Open the 'Creator List' sheet and select the rows you want to enrich.");
  }

  const selection = activeSheet.getActiveRange();
  if (!selection) {
    showSpreadsheetToast_("Select at least one row on Creator List.");
    return Logger.log("ℹ️ Select at least one row on 'Creator List'.");
  }

  const context = getCreatorListSheetContext_(sheet);
  if (!context) {
    showSpreadsheetToast_("Creator List has no data.");
    return Logger.log("ℹ️ Creator List has no data.");
  }

  const data = context.data;
  const header = context.header;
  const stopRow1 = context.archivedStart0 === -1 ? data.length : context.archivedStart0;
  const rowStart1 = Math.max(selection.getRow(), context.headerRow1 + 1);
  const rowEnd1 = Math.min(selection.getRow() + selection.getNumRows() - 1, stopRow1);

  const rowItems = [];
  for (let row1 = rowStart1; row1 <= rowEnd1; row1++) {
    const row = data[row1 - 1];
    if (isBlankRow_(row) || isSectionLabelRow_(row)) continue;
    if (!canAttemptCreatorListEnrichment_(row, header)) continue;
    rowItems.push({ row1: row1, values: row.slice() });
  }

  if (rowItems.length === 0) {
    showSpreadsheetToast_("No enrichable rows found in the selection.");
    return Logger.log("ℹ️ No enrichable Creator List rows found in the current selection.");
  }

  const result = runCreatorListEnrichment_(sheet, rowItems, header);
  showSpreadsheetToast_(
    `Selected rows done. Rows: ${rowItems.length}. Static: ${result.staticEnriched}. ` +
    `YouTube: ${result.youtubeEnriched}. AI: ${result.profileEnriched}.`
  );
  Logger.log(
    `✅ Selected Creator List rows enriched. Static: ${result.staticEnriched}, ` +
    `YouTube: ${result.youtubeEnriched}, AI: ${result.profileEnriched}.`
  );
}

function resetCreatorListEnrichmentProgress() {
  PropertiesService.getScriptProperties().deleteProperty(CREATOR_LIST_ENRICH_CURSOR_PROP_);
  showSpreadsheetToast_("Creator List enrichment progress reset.");
  Logger.log("✅ Creator List enrichment progress reset.");
}

function diagnoseCreatorListEnrichment() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Creator List");
  if (!sheet) {
    showSpreadsheetToast_("Creator List sheet not found.");
    return Logger.log("❌ 'Creator List' sheet not found.");
  }

  const context = getCreatorListSheetContext_(sheet);
  if (!context) {
    showSpreadsheetToast_("Creator List has no data.");
    return Logger.log("ℹ️ Creator List has no data.");
  }

  const data = context.data;
  const header = context.header;
  const stopRow1 = context.archivedStart0 === -1 ? data.length : context.archivedStart0;
  const youtubeKeyPresent = !!getYouTubeApiKey_();
  const openAiKeyPresent = !!getOpenAiApiKey_();
  const expectedHeaders = ["Channel Name", "Channel URL", "Campaign Name", "Email", "Status"];
  const missingHeaders = expectedHeaders.filter(name => findHeaderIndex_(header, name) === -1);

  let candidateRows = 0;
  let youtubeEligibleRows = 0;
  let profileEligibleRows = 0;
  const sampleNotes = [];

  for (let row1 = context.headerRow1 + 1; row1 <= stopRow1; row1++) {
    const row = data[row1 - 1];
    if (isBlankRow_(row) || isSectionLabelRow_(row)) continue;
    if (!canAttemptCreatorListEnrichment_(row, header)) continue;

    candidateRows++;

    const missingYoutube = YOUTUBE_ENRICH_FIELDS_.some(field => {
      const idx = header.indexOf(field);
      return idx !== -1 && !String(row[idx] || "").trim();
    });
    if (missingYoutube) youtubeEligibleRows++;

    const missingProfile = PROFILE_LLM_FIELDS_.some(field => {
      const idx = header.indexOf(field);
      return idx !== -1 && !String(row[idx] || "").trim();
    });
    if (missingProfile) profileEligibleRows++;

    if (sampleNotes.length < 8) {
      const channelName = String(getValueByHeader_(row, header, "Channel Name") || "").trim() || "(blank)";
      const reasons = [];
      if (missingYoutube) reasons.push("needs YouTube");
      if (missingProfile) reasons.push("needs AI");
      if (reasons.length === 0) reasons.push("nothing missing");
      sampleNotes.push(`Row ${row1} ${channelName}: ${reasons.join(", ")}`);
    }
  }

  const message = [
    `Header row: ${context.headerRow1}`,
    `YouTube key present: ${youtubeKeyPresent ? "yes" : "no"}`,
    `OpenAI key present: ${openAiKeyPresent ? "yes" : "no"}`,
    `Rows before Archived: ${Math.max(stopRow1 - 1, 0)}`,
    `Candidate rows: ${candidateRows}`,
    `Rows needing YouTube enrichment: ${youtubeEligibleRows}`,
    `Rows needing AI enrichment: ${profileEligibleRows}`
  ].join(" | ");

  showSpreadsheetToast_(message);
  Logger.log("Creator List enrichment diagnosis: " + message);
  if (missingHeaders.length > 0) {
    Logger.log("Missing or unrecognized key headers: " + missingHeaders.join(", "));
  }
  Logger.log("Detected Creator List headers: " + header.join(" | "));
  sampleNotes.forEach(note => Logger.log(note));
}

function diagnoseSelectedCreatorListRow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Creator List");
  if (!sheet) {
    showSpreadsheetToast_("Creator List sheet not found.");
    return Logger.log("❌ 'Creator List' sheet not found.");
  }

  const activeSheet = ss.getActiveSheet();
  if (!activeSheet || activeSheet.getSheetId() !== sheet.getSheetId()) {
    showSpreadsheetToast_("Open Creator List and select one row first.");
    return Logger.log("ℹ️ Open the 'Creator List' sheet and select one row first.");
  }

  const selection = activeSheet.getActiveRange();
  if (!selection) {
    showSpreadsheetToast_("Select one row first.");
    return Logger.log("ℹ️ Select one row first.");
  }

  const context = getCreatorListSheetContext_(sheet);
  if (!context) {
    showSpreadsheetToast_("Creator List has no data.");
    return Logger.log("ℹ️ Creator List has no data.");
  }

  const row1 = selection.getRow();
  if (row1 <= context.headerRow1) {
    showSpreadsheetToast_("Select a data row, not the header.");
    return Logger.log("ℹ️ Select a data row, not the header.");
  }

  const row = sheet.getRange(row1, 1, 1, context.header.length).getValues()[0];
  const header = context.header;
  const selectedInput = resolveCreatorYouTubeInputFromRow_(row, header);
  const rawChannelName = String(getValueByHeader_(row, header, "Channel Name") || "").trim();
  const rawChannelUrl = String(getValueByHeader_(row, header, "Channel URL") || "").trim();
  const rawYouTubeUrl = String(getValueByHeader_(row, header, "YouTube URL") || "").trim();
  const rawYouTubeHandle = String(getValueByHeader_(row, header, "YouTube Handle") || "").trim();
  const youtubeKeyPresent = !!getYouTubeApiKey_();
  const openAiKeyPresent = !!getOpenAiApiKey_();
  const youtubeKeyValid = youtubeKeyPresent ? validateYouTubeApiKey_(getYouTubeApiKey_()) : false;
  const youtubeTargetColumns = YOUTUBE_ENRICH_FIELDS_.filter(field => findHeaderIndex_(header, field) !== -1);
  const profileTargetColumns = PROFILE_LLM_FIELDS_.filter(field => findHeaderIndex_(header, field) !== -1);
  const emptyYoutubeTargets = youtubeTargetColumns.filter(field => {
    const idx = findHeaderIndex_(header, field);
    return idx !== -1 && !String(row[idx] || "").trim();
  });
  const emptyProfileTargets = profileTargetColumns.filter(field => {
    const idx = findHeaderIndex_(header, field);
    return idx !== -1 && !String(row[idx] || "").trim();
  });
  const channelId = selectedInput && youtubeKeyValid
    ? getChannelIdFromUrl(selectedInput, getYouTubeApiKey_())
    : "";
  const insight = selectedInput && youtubeKeyValid
    ? getCreatorYouTubeInsight_(
      selectedInput,
      rawChannelName,
      getYouTubeApiKey_(),
      { includeEmailSignals: emptyProfileTargets.indexOf("Email") !== -1 }
    )
    : null;
  const preferredEmailSignal = extractPreferredCreatorEmailFromContext_(row, header, insight);

  let aiDraft = null;
  let profileDropdownOptions = {};
  if (openAiKeyPresent && emptyProfileTargets.length > 0) {
    try {
      const dropdownValuesByHeader = getDropdownValuesByHeader_(ss, "Dropdown Values");
      profileDropdownOptions = getProfileDropdownOptionsByHeader_(sheet, header, dropdownValuesByHeader);
      const contextText = buildCreatorProfileContext_(row, header, insight);
      aiDraft = callLlmForCreatorProfileEnrichment_(
        getOpenAiApiKey_(),
        getOpenAiModel_(),
        rawChannelName,
        selectedInput,
        String(getValueByHeader_(row, header, "Campaign Name") || "").trim(),
        emptyProfileTargets,
        profileDropdownOptions,
        contextText
      );

      const remainingClassificationFields = PROFILE_LLM_CLASSIFICATION_FIELDS_.filter(field => {
        return emptyProfileTargets.indexOf(field) !== -1;
      });
      if (remainingClassificationFields.length > 0) {
        const fallbackDraft = callLlmForCreatorDropdownClassification_(
          getOpenAiApiKey_(),
          getOpenAiModel_(),
          remainingClassificationFields,
          profileDropdownOptions,
          contextText
        );
        aiDraft = Object.assign({}, aiDraft || {}, fallbackDraft || {});
      }
    } catch (e) {
      Logger.log("Selected Creator row AI dry-run error: " + (e && e.stack ? e.stack : e));
    }
  }

  const message = [
    `Row ${row1}`,
    `YT key: ${youtubeKeyPresent ? "yes" : "no"}`,
    `YT key valid: ${youtubeKeyValid ? "yes" : "no"}`,
    `OpenAI key: ${openAiKeyPresent ? "yes" : "no"}`,
    `Resolved input: ${selectedInput ? "yes" : "no"}`,
    `Channel ID: ${channelId || "none"}`,
    `Insight: ${insight ? "yes" : "no"}`,
    `YT targets: ${youtubeTargetColumns.length}`,
    `AI targets: ${profileTargetColumns.length}`
  ].join(" | ");

  const detailLines = [
    "Selected Creator row diagnosis",
    "This tool does not write values. It only shows what enrichment would return.",
    "",
    message,
    "Resolved YouTube input: " + (selectedInput || "(blank)"),
    "Detected AI target columns: " + (profileTargetColumns.join(", ") || "(none)"),
    "Empty AI target columns: " + (emptyProfileTargets.join(", ") || "(none)"),
    "Detected explicit email candidate: " + (
      preferredEmailSignal.email
        ? (preferredEmailSignal.email + " (" + preferredEmailSignal.source + ")")
        : "(none)"
    ),
    "AI dropdown option counts: " + JSON.stringify({
      "Influencer Type": (profileDropdownOptions["Influencer Type"] || []).length,
      "Influencer Vertical": (profileDropdownOptions["Influencer Vertical"] || []).length,
      "Country/Region": (profileDropdownOptions["Country/Region"] || []).length,
      "Language": (profileDropdownOptions["Language"] || []).length
    }),
    "AI dry-run result: " + truncateForUi_(aiDraft ? JSON.stringify(aiDraft) : "(none)", 800),
    "Last OpenAI diagnostic: " + truncateForUi_(LAST_OPENAI_DIAGNOSTIC_ || "(none)", 800)
  ];

  showSpreadsheetToast_(message);
  showSpreadsheetAlert_(detailLines.join("\n"));
  Logger.log("Selected Creator row diagnosis: " + message);
  Logger.log("Channel Name raw: " + (rawChannelName || "(blank)"));
  Logger.log("Channel URL raw: " + (rawChannelUrl || "(blank)"));
  Logger.log("YouTube URL raw: " + (rawYouTubeUrl || "(blank)"));
  Logger.log("YouTube Handle raw: " + (rawYouTubeHandle || "(blank)"));
  Logger.log("Resolved YouTube input: " + (selectedInput || "(blank)"));
  Logger.log("Detected YouTube target columns: " + (youtubeTargetColumns.join(", ") || "(none)"));
  Logger.log("Empty YouTube target columns: " + (emptyYoutubeTargets.join(", ") || "(none)"));
  Logger.log("Detected AI target columns: " + (profileTargetColumns.join(", ") || "(none)"));
  Logger.log("Empty AI target columns: " + (emptyProfileTargets.join(", ") || "(none)"));
  Logger.log(
    "Detected explicit email candidate: " + (
      preferredEmailSignal.email
        ? (preferredEmailSignal.email + " (" + preferredEmailSignal.source + ")")
        : "(none)"
    )
  );
  Logger.log("AI dropdown option counts: " + JSON.stringify({
    "Influencer Type": (profileDropdownOptions["Influencer Type"] || []).length,
    "Influencer Vertical": (profileDropdownOptions["Influencer Vertical"] || []).length,
    "Country/Region": (profileDropdownOptions["Country/Region"] || []).length,
    "Language": (profileDropdownOptions["Language"] || []).length
  }));
  if (insight) {
    Logger.log("Resolved channel name: " + (insight.channelName || "(blank)"));
    Logger.log("Resolved handle: " + (insight.handle || "(blank)"));
    Logger.log("Resolved canonical URL: " + (insight.canonicalUrl || "(blank)"));
    Logger.log("Median video views: " + String(insight.medianVideoViews || ""));
    Logger.log("Median shorts views: " + String(insight.medianShortsViews || ""));
    Logger.log("Median engagement rate: " + String(insight.medianVideoEngagementRate || ""));
  }
  if (aiDraft) {
    Logger.log("AI dry-run result: " + JSON.stringify(aiDraft));
  } else {
    Logger.log("AI dry-run result: (none)");
  }
  if (LAST_OPENAI_DIAGNOSTIC_) {
    Logger.log("Last OpenAI diagnostic: " + LAST_OPENAI_DIAGNOSTIC_);
  }
  if (LAST_OPENAI_RAW_RESPONSE_) {
    Logger.log("Last OpenAI raw response: " + LAST_OPENAI_RAW_RESPONSE_);
  }
}

function runCreatorListEnrichment_(sheet, rowItems, header) {
  const dropdownValuesByHeader = getDropdownValuesByHeader_(sheet.getParent(), "Dropdown Values");
  const defaultClientName = getDefaultClientName_(dropdownValuesByHeader);

  const staticUpdates = [];
  for (const item of rowItems) {
    const rowChanges = enrichCreatorRow_(item.values, header, defaultClientName);
    if (rowChanges.length === 0) continue;
    staticUpdates.push.apply(staticUpdates, trackedRowChangesToSheetUpdates_(item, rowChanges));
  }

  const staticWrite = writeSparseCellUpdates_(sheet, header, staticUpdates, "Static enrichment");

  const youtubeEnriched = enrichYouTubeDataForRows_(sheet, rowItems, header);
  const profileEnriched = enrichProfileFieldsViaLlm_(sheet, rowItems, header, dropdownValuesByHeader);

  return {
    staticEnriched: staticWrite.writtenRowCount,
    youtubeEnriched: youtubeEnriched,
    profileEnriched: profileEnriched
  };
}

function canAttemptCreatorListEnrichment_(row, header) {
  const candidateFields = [
    "Channel Name",
    "Channel URL",
    "Campaign Name",
    "YouTube URL",
    "YouTube Handle"
  ];

  return candidateFields.some(field => String(getValueByHeader_(row, header, field) || "").trim());
}


// ============================================================
//  (1) CREATOR LIST → HUBSPOT (+ ENRICHMENT)
// ============================================================

function enrichAndImportToHubSpot() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Creator List");
  if (!sheet) return ui.alert("❌ 'Creator List' sheet not found.");

  const context = getCreatorListSheetContext_(sheet);
  if (!context) return ui.alert("ℹ️ Creator List has no data.");

  const data = context.data;
  const header = context.header;
  const numCols = header.length;
  const dropdownValuesByHeader = getDropdownValuesByHeader_(ss, "Dropdown Values");
  const defaultClientName = getDefaultClientName_(dropdownValuesByHeader);

  const contactingStart0 = findSectionRowByLabel_(data, "Contacting");
  const archivedStart0 = findSectionRowByLabel_(data, "Archived");
  if (contactingStart0 === -1) return ui.alert("❌ 'Contacting' section not found.");
  if (archivedStart0 === -1) return ui.alert("❌ 'Archived' section not found.");

  const emailCol = findHeaderIndex_(header, "Email");
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
  const staticUpdates = [];
  for (const item of rowItems) {
    const rowChanges = enrichCreatorRow_(item.values, header, defaultClientName);
    if (rowChanges.length === 0) continue;
    staticUpdates.push.apply(staticUpdates, trackedRowChangesToSheetUpdates_(item, rowChanges));
  }

  // Batch write enriched data back
  const staticWrite = writeSparseCellUpdates_(sheet, header, staticUpdates, "Static enrichment");
  Logger.log(`✅ Static enrichment done for ${staticWrite.writtenRowCount} row(s).`);

  // Stage 2: YouTube API enrichment
  enrichYouTubeDataForRows_(sheet, rowItems, header);

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
  const idx = name => findHeaderIndex_(header, name);
  const get = name => { const i = idx(name); return i === -1 ? "" : String(row[i] || "").trim(); };
  const changeMap = {};
  const set = (name, value) => setTrackedValueByHeader_(row, header, name, value, changeMap);
  const setIfEmpty = (name, value) => setIfEmpty_(row, header, name, value, changeMap);

  const channelName = get("Channel Name");
  const campaignName = get("Campaign Name");
  if (!channelName && !campaignName) return [];

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

  return trackedChangeMapToRowChanges_(changeMap);
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
    return 0;
  }
  if (!validateYouTubeApiKey_(apiKey)) return 0;

  const updates = [];
  for (const item of rowItems) {
    const needsYoutubeFields = YOUTUBE_ENRICH_FIELDS_.some(field => {
      const fieldIdx = findHeaderIndex_(header, field);
      return fieldIdx !== -1 && !String(item.values[fieldIdx] || "").trim();
    });
    if (!needsYoutubeFields) continue;

    try {
      const result = enrichSingleYouTubeRow_(item, header, apiKey);
      if (result.updates.length === 0) continue;
      updates.push.apply(updates, trackedRowChangesToSheetUpdates_(item, result.updates));
    } catch (e) {
      Logger.log(`⚠️ YouTube enrichment failed for row ${item.row1}: ${e}`);
    }
  }
  const writeResult = writeSparseCellUpdates_(sheet, header, updates, "YouTube enrichment");
  Logger.log(`✅ YouTube enrichment: filled ${writeResult.writtenRowCount} row(s).`);
  return writeResult.writtenRowCount;
}


function enrichSingleYouTubeRow_(item, header, apiKey) {
  const row = item.values;
  const channelUrl = resolveCreatorYouTubeInputFromRow_(row, header);
  const channelName = String(getValueByHeader_(row, header, "Channel Name") || "").trim();
  const insight = item.youtubeInsight || getCreatorYouTubeInsight_(channelUrl, channelName, apiKey);
  if (!insight || !insight.channelId) {
    return {
      insight: null,
      updates: []
    };
  }

  item.youtubeInsight = insight;
  const changeMap = {};
  if (insight.subscribers != null) {
    setIfEmpty_(row, header, "YouTube Followers", insight.subscribers, changeMap);
  }
  if (insight.handle) setIfEmpty_(row, header, "YouTube Handle", insight.handle, changeMap);
  if (insight.canonicalUrl) setIfEmpty_(row, header, "YouTube URL", insight.canonicalUrl, changeMap);

  if (insight.medianVideoViews != null) {
    setIfEmpty_(row, header, "YouTube Video Median Views", insight.medianVideoViews, changeMap);
    setIfEmpty_(row, header, "YouTube Average Views", insight.medianVideoViews, changeMap);
  }
  if (insight.medianShortsViews != null) {
    setIfEmpty_(row, header, "YouTube Shorts Median Views", insight.medianShortsViews, changeMap);
  }
  if (insight.medianVideoEngagementRate != null) {
    setIfEmpty_(row, header, "YouTube Engagement Rate", insight.medianVideoEngagementRate, changeMap);
  }

  return {
    insight: insight,
    updates: trackedChangeMapToRowChanges_(changeMap)
  };
}

function getCreatorYouTubeInsight_(channelUrl, channelName, apiKey, options) {
  const includeEmailSignals = !!(options && options.includeEmailSignals);
  const resolved = resolveYouTubeChannelForEnrichment_(channelUrl, channelName, apiKey);
  if (!resolved || !resolved.channelId) return null;

  const cacheKey = CREATOR_YOUTUBE_INSIGHT_CACHE_PREFIX_ + resolved.channelId;
  const cached = getJsonFromScriptCache_(cacheKey);
  if (cached) return decorateCreatorInsightWithEmailSignals_(cached, includeEmailSignals);

  const channelData = ytFetchJson_(
    "https://www.googleapis.com/youtube/v3/channels?part=snippet,statistics,contentDetails,brandingSettings" +
    "&id=" + encodeURIComponent(resolved.channelId) +
    "&key=" + encodeURIComponent(apiKey)
  );
  const item = channelData.items && channelData.items[0];
  if (!item) return null;

  const snippet = item.snippet || {};
  const branding = item.brandingSettings && item.brandingSettings.channel
    ? item.brandingSettings.channel
    : {};
  const statistics = item.statistics || {};
  const uploadsPlaylistId =
    item.contentDetails &&
    item.contentDetails.relatedPlaylists &&
    item.contentDetails.relatedPlaylists.uploads;

  const videoStats = computeYouTubeVideoStats_(resolved.channelId, apiKey) || {};
  const longformSignals = uploadsPlaylistId
    ? getCreatorLongformSignals_(uploadsPlaylistId, apiKey)
    : null;

  let handle = normalizeYouTubeHandle_(snippet.customUrl);
  if (!handle && /\/@/i.test(String(resolved.canonicalUrl || ""))) {
    handle = normalizeYouTubeHandle_(resolved.canonicalUrl);
  }

  const insight = {
    channelId: resolved.channelId,
    channelName: String(snippet.title || channelName || "").trim(),
    handle: handle,
    canonicalUrl: resolved.canonicalUrl || buildCanonicalYouTubeUrl_(resolved.channelId, snippet.customUrl),
    description: String(snippet.description || branding.description || "").trim(),
    countryCode: String(snippet.country || branding.country || "").trim(),
    subscribers: Number(statistics.subscriberCount || 0),
    medianVideoViews: videoStats.medianVideoViews != null ? videoStats.medianVideoViews : null,
    medianShortsViews: videoStats.medianShortsViews != null ? videoStats.medianShortsViews : null,
    medianVideoEngagementRate: videoStats.medianVideoEngagementRate != null ? videoStats.medianVideoEngagementRate : null,
    dominantCategoryName: longformSignals && longformSignals.dominantCategoryName
      ? longformSignals.dominantCategoryName
      : "",
    sampleSize: longformSignals && longformSignals.sampleSize ? longformSignals.sampleSize : 0,
    sampledTitles: longformSignals && longformSignals.sampledTitles ? longformSignals.sampledTitles : [],
    sampledVideoDescriptions: longformSignals && longformSignals.sampledVideoDescriptions
      ? longformSignals.sampledVideoDescriptions
      : []
  };

  putJsonInScriptCache_(cacheKey, insight, CREATOR_YOUTUBE_SIGNAL_TTL_SECONDS_);
  return decorateCreatorInsightWithEmailSignals_(insight, includeEmailSignals);
}

function getCreatorLongformSignals_(uploadsPlaylistId, apiKey) {
  const cacheKey = CREATOR_LONGFORM_SIGNAL_CACHE_PREFIX_ + uploadsPlaylistId;
  const cached = getJsonFromScriptCache_(cacheKey);
  if (cached) return cached;

  const sampledTitles = [];
  const sampledVideoDescriptions = [];
  const sampledCategories = [];
  let pageToken = "";

  while (sampledTitles.length < CREATOR_YOUTUBE_SIGNAL_SAMPLE_SIZE_) {
    const playlistData = ytFetchJson_(
      "https://www.googleapis.com/youtube/v3/playlistItems?part=snippet,contentDetails" +
      "&maxResults=50" +
      "&playlistId=" + encodeURIComponent(uploadsPlaylistId) +
      "&key=" + encodeURIComponent(apiKey) +
      (pageToken ? "&pageToken=" + encodeURIComponent(pageToken) : "")
    );

    const items = Array.isArray(playlistData.items) ? playlistData.items : [];
    if (items.length === 0) break;

    const videoIds = items
      .map(item => item && item.contentDetails && item.contentDetails.videoId
        ? String(item.contentDetails.videoId || "").trim()
        : "")
      .filter(Boolean);

    const videos = getCreatorVideosByIds_(videoIds, apiKey);
    videos.forEach(video => {
      if (video.durationSeconds < CREATOR_YOUTUBE_SIGNAL_MIN_LONGFORM_SECONDS_) return;
      if (sampledTitles.length >= CREATOR_YOUTUBE_SIGNAL_SAMPLE_SIZE_) return;

      if (video.title) sampledTitles.push(video.title);
      if (
        video.description &&
        sampledVideoDescriptions.length < CREATOR_YOUTUBE_DESCRIPTION_SAMPLE_SIZE_ &&
        sampledVideoDescriptions.indexOf(video.description) === -1
      ) {
        sampledVideoDescriptions.push(video.description);
      }
      if (video.categoryId) sampledCategories.push(video.categoryId);
    });

    pageToken = String(playlistData.nextPageToken || "");
    if (!pageToken) break;
  }

  if (sampledTitles.length === 0) return null;

  const dominantCategoryId = modeStringArray_(sampledCategories);
  const result = {
    sampleSize: sampledTitles.length,
    sampledTitles: sampledTitles,
    sampledVideoDescriptions: sampledVideoDescriptions,
    dominantCategoryName: getEnglishYouTubeCategoryName_(dominantCategoryId)
  };

  putJsonInScriptCache_(cacheKey, result, CREATOR_YOUTUBE_SIGNAL_TTL_SECONDS_);
  return result;
}

function getCreatorVideosByIds_(ids, apiKey) {
  if (!ids || ids.length === 0) return [];

  const data = ytFetchJson_(
    "https://www.googleapis.com/youtube/v3/videos?part=snippet,contentDetails" +
    "&id=" + encodeURIComponent(ids.join(",")) +
    "&key=" + encodeURIComponent(apiKey)
  );
  const items = Array.isArray(data.items) ? data.items : [];

  return items.map(item => {
    const snippet = item.snippet || {};
    return {
      title: String(snippet.title || "").trim(),
      description: String(snippet.description || "").trim(),
      categoryId: String(snippet.categoryId || "").trim(),
      durationSeconds: iso8601DurationToSeconds_(item.contentDetails && item.contentDetails.duration)
    };
  });
}

function decorateCreatorInsightWithEmailSignals_(insight, includeEmailSignals) {
  if (!insight) return null;

  const out = Object.assign({}, insight);
  out.sampledTitles = uniqueNonEmptyStrings_(out.sampledTitles || []);
  out.sampledVideoDescriptions = uniqueNonEmptyStrings_(out.sampledVideoDescriptions || []);
  out.bioEmails = extractExplicitEmailsFromText_(out.description || "");
  out.videoDescriptionEmails = extractExplicitEmailsFromTextList_(out.sampledVideoDescriptions || []);

  let channelPageEmails = Array.isArray(out.channelPageEmails)
    ? uniqueNonEmptyStrings_(out.channelPageEmails)
    : [];
  let channelPageSnippet = String(out.channelPageSnippet || "");

  if (includeEmailSignals && out.canonicalUrl && channelPageEmails.length === 0 && !channelPageSnippet) {
    const channelPageSignal = getYouTubeChannelPageEmailSignal_(out.canonicalUrl);
    if (channelPageSignal) {
      channelPageEmails = Array.isArray(channelPageSignal.emails)
        ? uniqueNonEmptyStrings_(channelPageSignal.emails)
        : [];
      channelPageSnippet = String(channelPageSignal.snippet || "");
    }
  }

  out.channelPageEmails = channelPageEmails;
  out.channelPageSnippet = channelPageSnippet;

  const preferredEmailSignal = pickPreferredCreatorEmailSignal_(out);
  out.preferredEmail = preferredEmailSignal.email;
  out.preferredEmailSource = preferredEmailSignal.source;
  return out;
}

function pickPreferredCreatorEmailSignal_(insight) {
  const sources = [
    { source: "channel bio", emails: insight && insight.bioEmails },
    { source: "channel page", emails: insight && insight.channelPageEmails },
    { source: "video description", emails: insight && insight.videoDescriptionEmails }
  ];

  for (let i = 0; i < sources.length; i++) {
    const emails = Array.isArray(sources[i].emails) ? sources[i].emails : [];
    if (emails.length > 0) {
      return {
        email: emails[0],
        source: sources[i].source
      };
    }
  }

  return {
    email: "",
    source: ""
  };
}

function extractPreferredCreatorEmailFromContext_(row, header, youtubeInsight) {
  const prioritizedInsightSignal = pickPreferredCreatorEmailSignal_(youtubeInsight || {});
  if (prioritizedInsightSignal.email) return prioritizedInsightSignal;

  const fallbackParts = [];
  (header || []).forEach(function (field, index) {
    if (normalizeHeaderName_(field) === normalizeHeaderName_("Email")) return;
    const value = normalizeLlmString_(row && row[index]);
    if (!value) return;
    fallbackParts.push(String(field || "").trim() + ": " + value);
  });

  const fallbackEmails = extractExplicitEmailsFromText_(fallbackParts.join("\n"));
  return {
    email: fallbackEmails.length > 0 ? fallbackEmails[0] : "",
    source: fallbackEmails.length > 0 ? "row context" : ""
  };
}

function getYouTubeChannelPageEmailSignal_(canonicalUrl) {
  const normalizedUrl = String(canonicalUrl || "").trim().replace(/\/+$/, "");
  if (!normalizedUrl) {
    return { emails: [], snippet: "" };
  }

  const cacheKey = CREATOR_CHANNEL_PAGE_SIGNAL_CACHE_PREFIX_ + normalizedUrl;
  const cached = getJsonFromScriptCache_(cacheKey);
  if (cached) return cached;

  const fetchUrls = uniqueNonEmptyStrings_([
    buildYouTubeAboutPageUrl_(normalizedUrl),
    normalizedUrl
  ]);

  const emails = [];
  let snippet = "";

  for (let i = 0; i < fetchUrls.length; i++) {
    const url = fetchUrls[i];
    const text = fetchUrlTextForEmailScan_(url);
    if (!text) continue;

    const foundEmails = extractExplicitEmailsFromText_(text);
    foundEmails.forEach(function (email) {
      if (emails.indexOf(email) !== -1) return;
      emails.push(email);
    });

    if (!snippet && emails.length > 0) {
      snippet = extractEmailSnippetFromText_(text, emails[0]);
    }
    if (emails.length > 0) break;
  }

  const result = {
    emails: emails,
    snippet: snippet
  };
  putJsonInScriptCache_(cacheKey, result, CREATOR_YOUTUBE_SIGNAL_TTL_SECONDS_);
  return result;
}

function buildYouTubeAboutPageUrl_(canonicalUrl) {
  const base = String(canonicalUrl || "").trim().replace(/\/+$/, "");
  if (!base) return "";
  return /\/about$/i.test(base) ? base : (base + "/about");
}

function fetchUrlTextForEmailScan_(url) {
  const target = String(url || "").trim();
  if (!target) return "";

  try {
    const response = UrlFetchApp.fetch(target, {
      muteHttpExceptions: true,
      headers: {
        "Accept-Language": "en-US,en;q=0.9"
      }
    });
    if (response.getResponseCode() >= 400) return "";

    return String(response.getContentText() || "").slice(0, 250000);
  } catch (e) {
    Logger.log("⚠️ Channel page fetch failed for email extraction: " + e);
    return "";
  }
}

function extractEmailSnippetFromText_(text, email) {
  const raw = String(text || "");
  const normalizedEmail = String(email || "").toLowerCase();
  if (!raw || !normalizedEmail) return "";

  const startIndex = raw.toLowerCase().indexOf(normalizedEmail);
  if (startIndex === -1) return "";

  const snippet = raw.slice(Math.max(0, startIndex - 160), startIndex + normalizedEmail.length + 160);
  return truncateForUi_(snippet.replace(/\s+/g, " ").trim(), 400);
}

function extractExplicitEmailsFromTextList_(values) {
  const results = [];
  (values || []).forEach(function (value) {
    extractExplicitEmailsFromText_(value).forEach(function (email) {
      if (results.indexOf(email) !== -1) return;
      results.push(email);
    });
  });
  return results;
}

function extractExplicitEmailsFromText_(value) {
  const raw = String(value || "");
  if (!raw) return [];

  const results = [];
  const seen = {};

  function collectFromText(text) {
    const pattern = /(?:mailto:)?([A-Z0-9.!#$%&'*+/=?^_`{|}~-]+@[A-Z0-9.-]+\.[A-Z]{2,63})/ig;
    let match;
    while ((match = pattern.exec(text)) !== null) {
      const email = normalizeExtractedEmailCandidate_(match[1] || match[0]);
      if (!email) continue;
      if (getHubSpotImportEmailValidationIssue_(email)) continue;

      const key = email.toLowerCase();
      if (seen[key]) continue;
      seen[key] = true;
      results.push(email);
    }
  }

  const decodedHtml = raw
    .replace(/&commat;|&#64;|&#x40;/ig, "@")
    .replace(/&period;|&#46;|&#x2e;/ig, ".");
  collectFromText(decodedHtml);

  const deobfuscated = decodedHtml
    .replace(/\s*\[\s*at\s*\]\s*/ig, "@")
    .replace(/\s*\(\s*at\s*\)\s*/ig, "@")
    .replace(/\s+\bat\b\s+/ig, "@")
    .replace(/\s*\[\s*dot\s*\]\s*/ig, ".")
    .replace(/\s*\(\s*dot\s*\)\s*/ig, ".")
    .replace(/\s+\bdot\b\s+/ig, ".");
  if (deobfuscated !== decodedHtml) {
    collectFromText(deobfuscated);
  }

  return results;
}

function normalizeExtractedEmailCandidate_(value) {
  return String(value || "")
    .replace(/^mailto:/i, "")
    .replace(/^[<("'`\[]+/, "")
    .replace(/[>"')\],;:!?]+$/, "")
    .trim()
    .toLowerCase();
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
    } else if (durationSeconds <= YT_SHORTS_MAX_DURATION_SECONDS_ && shortsViewsList.length < YT_MAX_VIDEOS_FOR_AVG_) {
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
  const model = apiKey ? getOpenAiModel_() : "";
  if (!apiKey) {
    Logger.log("ℹ️ OpenAI API key not set. AI enrichment will be skipped, but explicit email extraction will still run.");
  }
  const profileDropdownOptions = getProfileDropdownOptionsByHeader_(sheet, header, dropdownValuesByHeader);
  const youtubeApiKey = getYouTubeApiKey_();
  const validatedYouTubeApiKey = (youtubeApiKey && validateYouTubeApiKey_(youtubeApiKey))
    ? youtubeApiKey
    : "";
  const aiRequests = [];
  const directUpdates = [];
  for (const item of rowItems) {
    const channelName = String(getValueByHeader_(item.values, header, "Channel Name") || "").trim();
    const channelUrl = String(getValueByHeader_(item.values, header, "Channel URL") || "").trim();
    const campaignName = String(getValueByHeader_(item.values, header, "Campaign Name") || "").trim();
    if (!channelName && !channelUrl && !campaignName) continue;

    // Only attempt fields that are still empty
    let emptyFields = PROFILE_LLM_FIELDS_.filter(f => {
      const i = findHeaderIndex_(header, f);
      return i !== -1 && !String(item.values[i] || "").trim();
    });
    if (emptyFields.length === 0) continue;

    try {
      const changeMap = {};
      const needsEmail = emptyFields.indexOf("Email") !== -1;
      let youtubeInsight = item.youtubeInsight || null;

      if (youtubeInsight) {
        youtubeInsight = decorateCreatorInsightWithEmailSignals_(youtubeInsight, needsEmail);
        item.youtubeInsight = youtubeInsight;
      } else if ((apiKey || needsEmail) && validatedYouTubeApiKey) {
        youtubeInsight = getCreatorProfileInsightForLlm_(
          item.values,
          header,
          validatedYouTubeApiKey,
          { includeEmailSignals: needsEmail }
        );
        if (youtubeInsight) item.youtubeInsight = youtubeInsight;
      }

      if (needsEmail) {
        const preferredEmailSignal = extractPreferredCreatorEmailFromContext_(item.values, header, youtubeInsight);
        if (preferredEmailSignal.email) {
          applySanitizedProfileFields_(
            item.values,
            header,
            ["Email"],
            { "Email": preferredEmailSignal.email },
            changeMap
          );
        }

        emptyFields = PROFILE_LLM_FIELDS_.filter(f => {
          const i = findHeaderIndex_(header, f);
          return i !== -1 && !String(item.values[i] || "").trim();
        });
      }

      if (emptyFields.length === 0) {
        directUpdates.push.apply(
          directUpdates,
          trackedRowChangesToSheetUpdates_(item, trackedChangeMapToRowChanges_(changeMap))
        );
        continue;
      }

      if (!apiKey) {
        directUpdates.push.apply(
          directUpdates,
          trackedRowChangesToSheetUpdates_(item, trackedChangeMapToRowChanges_(changeMap))
        );
        continue;
      }

      const contextText = buildCreatorProfileContext_(item.values, header, youtubeInsight);
      aiRequests.push({
        rowKey: String(item.row1),
        item: item,
        channelName: channelName,
        channelUrl: channelUrl,
        campaignName: campaignName,
        requestedFields: emptyFields.slice(),
        contextText: contextText,
        changeMap: changeMap
      });
    } catch (e) {
      Logger.log(`⚠️ LLM enrichment failed for row ${item.row1}: ${e}`);
    }
  }

  if (apiKey && aiRequests.length > 0) {
    applyCreatorProfileBatchResults_(apiKey, model, header, profileDropdownOptions, aiRequests);
  }

  const updates = directUpdates.slice();
  aiRequests.forEach(function (request) {
    updates.push.apply(
      updates,
      trackedRowChangesToSheetUpdates_(request.item, trackedChangeMapToRowChanges_(request.changeMap))
    );
  });

  const writeResult = writeSparseCellUpdates_(sheet, header, updates, "Profile enrichment");
  Logger.log(`✅ Profile enrichment: filled ${writeResult.writtenRowCount} row(s).`);
  return writeResult.writtenRowCount;
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
    "Return JSON only.",
    "Fill only the requested fields.",
    "Use an empty string when a field is uncertain.",
    "Only include First Name and Last Name when the creator is clearly an individual person and the split is unambiguous.",
    "Only include Email when an explicit email address appears in the provided evidence.",
    "If multiple explicit emails appear, prefer the one from the channel bio, then the channel page/about text, then video descriptions.",
    "Return Email as a plain email address only.",
    "For Influencer Type, Country/Region, and Language, use only exact values from the allowed lists below.",
    "For Influencer Vertical, return an array of exact allowed values.",
    "Use 1 Influencer Vertical whenever a single best fit exists.",
    "Only return 2 or 3 Influencer Vertical values when one value would clearly lose important context.",
    "Never return more than 3 Influencer Vertical values.",
    "Do not infer sensitive traits.",
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

  return callOpenAiStructuredCreatorProfile_(
    apiKey,
    model,
    "You enrich creator CRM rows using explicit creator evidence. Respond with valid JSON only.",
    prompt,
    dropdownValuesByHeader
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
    "Return JSON only.",
    "Choose the best exact allowed value for each requested field when the creator context supports it.",
    "For these CRM classification fields, prefer the closest exact allowed match over leaving a field blank.",
    "Use an empty string only when there is truly no reasonable signal.",
    "For Influencer Vertical, return an array of 1 to 3 exact allowed values.",
    "Return 1 Influencer Vertical whenever possible.",
    "Only return 2 or 3 Influencer Vertical values when one value would clearly lose important context.",
    "Never return more than 3 Influencer Vertical values.",
    "",
    "Requested fields: " + requestedFields.join(", "),
    "",
    "Creator context:",
    contextText || "(blank)",
    "",
    "Allowed Influencer Type values: " + JSON.stringify(dropdownValuesByHeader["Influencer Type"] || []),
    "Allowed Influencer Vertical values: " + JSON.stringify(dropdownValuesByHeader["Influencer Vertical"] || []),
    "Allowed Country/Region values: " + JSON.stringify(dropdownValuesByHeader["Country/Region"] || []),
    "Allowed Language values: " + JSON.stringify(dropdownValuesByHeader["Language"] || [])
  ].join("\n");

  return callOpenAiStructuredCreatorProfile_(
    apiKey,
    model,
    "You classify creator CRM fields from creator evidence. Use exact allowed values and respond with valid JSON only.",
    prompt,
    dropdownValuesByHeader
  );
}

function applyCreatorProfileBatchResults_(apiKey, model, header, dropdownValuesByHeader, requests) {
  const requestChunks = chunkArray_(requests || [], OPENAI_CREATOR_PROFILE_BATCH_SIZE_);

  requestChunks.forEach(function (requestChunk) {
    if (!requestChunk || requestChunk.length === 0) return;

    try {
      const profileResultsByRowKey = callLlmForCreatorProfileEnrichmentBatch_(
        apiKey,
        model,
        requestChunk,
        dropdownValuesByHeader
      );

      requestChunk.forEach(function (request) {
        const result = profileResultsByRowKey[request.rowKey];
        if (!result) return;

        const sanitized = sanitizeCreatorProfileEnrichment_(result, dropdownValuesByHeader);
        applySanitizedProfileFields_(
          request.item.values,
          header,
          request.requestedFields,
          sanitized,
          request.changeMap
        );
      });

      const classificationRequests = requestChunk
        .map(function (request) {
          const remainingClassificationFields = PROFILE_LLM_CLASSIFICATION_FIELDS_.filter(function (field) {
            const i = findHeaderIndex_(header, field);
            return request.requestedFields.indexOf(field) !== -1 && i !== -1 && !String(request.item.values[i] || "").trim();
          });

          return {
            rowKey: request.rowKey,
            item: request.item,
            contextText: request.contextText,
            requestedFields: remainingClassificationFields,
            changeMap: request.changeMap
          };
        })
        .filter(function (request) {
          return request.requestedFields.length > 0;
        });

      if (classificationRequests.length === 0) return;

      const classificationResultsByRowKey = callLlmForCreatorDropdownClassificationBatch_(
        apiKey,
        model,
        classificationRequests,
        dropdownValuesByHeader
      );

      classificationRequests.forEach(function (request) {
        const result = classificationResultsByRowKey[request.rowKey];
        if (!result) return;

        const sanitized = sanitizeCreatorProfileEnrichment_(result, dropdownValuesByHeader);
        applySanitizedProfileFields_(
          request.item.values,
          header,
          request.requestedFields,
          sanitized,
          request.changeMap
        );
      });
    } catch (e) {
      Logger.log("⚠️ Batched LLM enrichment failed: " + (e && e.stack ? e.stack : e));
    }
  });
}

function callLlmForCreatorProfileEnrichmentBatch_(
  apiKey,
  model,
  requests,
  dropdownValuesByHeader
) {
  const prompt = [
    "Return JSON only.",
    "Process each creator independently.",
    "Return exactly one result object for each provided row_key.",
    "Use the same row_key values provided in the input.",
    "Fill only the requested fields for each row.",
    "For any field not listed in requested_fields for that row, return an empty string or an empty array.",
    "Use an empty string when a requested scalar field is uncertain.",
    "Use an empty array when a requested Influencer Vertical value is uncertain.",
    "Only include First Name and Last Name when the creator is clearly an individual person and the split is unambiguous.",
    "Only include Email when an explicit email address appears in the provided evidence.",
    "If multiple explicit emails appear, prefer the one from the channel bio, then the channel page/about text, then video descriptions.",
    "Return Email as a plain email address only.",
    "For Influencer Type, Country/Region, and Language, use only exact values from the allowed lists below.",
    "For Influencer Vertical, return an array of exact allowed values.",
    "Use 1 Influencer Vertical whenever a single best fit exists.",
    "Only return 2 or 3 Influencer Vertical values when one value would clearly lose important context.",
    "Never return more than 3 Influencer Vertical values.",
    "Do not infer sensitive traits.",
    "",
    "Rows to process:",
    JSON.stringify(buildCreatorProfileBatchPromptRows_(requests), null, 2),
    "",
    "Allowed Influencer Type values: " + JSON.stringify(dropdownValuesByHeader["Influencer Type"] || []),
    "Allowed Influencer Vertical values: " + JSON.stringify(dropdownValuesByHeader["Influencer Vertical"] || []),
    "Allowed Country/Region values: " + JSON.stringify(dropdownValuesByHeader["Country/Region"] || []),
    "Allowed Language values: " + JSON.stringify(dropdownValuesByHeader["Language"] || [])
  ].join("\n");

  return callOpenAiStructuredCreatorProfileBatch_(
    apiKey,
    model,
    "You enrich creator CRM rows using explicit creator evidence. Respond with valid JSON only.",
    prompt,
    dropdownValuesByHeader,
    requests,
    "creator_profile_enrichment_batch"
  );
}

function callLlmForCreatorDropdownClassificationBatch_(
  apiKey,
  model,
  requests,
  dropdownValuesByHeader
) {
  const prompt = [
    "Return JSON only.",
    "Process each creator independently.",
    "Return exactly one result object for each provided row_key.",
    "Use the same row_key values provided in the input.",
    "Choose the best exact allowed value for each requested field when the creator context supports it.",
    "For these CRM classification fields, prefer the closest exact allowed match over leaving a field blank.",
    "Use an empty string only when there is truly no reasonable signal.",
    "For any field not listed in requested_fields for that row, return an empty string or an empty array.",
    "For Influencer Vertical, return an array of 1 to 3 exact allowed values.",
    "Return 1 Influencer Vertical whenever possible.",
    "Only return 2 or 3 Influencer Vertical values when one value would clearly lose important context.",
    "Never return more than 3 Influencer Vertical values.",
    "",
    "Rows to process:",
    JSON.stringify(buildCreatorProfileBatchPromptRows_(requests), null, 2),
    "",
    "Allowed Influencer Type values: " + JSON.stringify(dropdownValuesByHeader["Influencer Type"] || []),
    "Allowed Influencer Vertical values: " + JSON.stringify(dropdownValuesByHeader["Influencer Vertical"] || []),
    "Allowed Country/Region values: " + JSON.stringify(dropdownValuesByHeader["Country/Region"] || []),
    "Allowed Language values: " + JSON.stringify(dropdownValuesByHeader["Language"] || [])
  ].join("\n");

  return callOpenAiStructuredCreatorProfileBatch_(
    apiKey,
    model,
    "You classify creator CRM fields from creator evidence. Use exact allowed values and respond with valid JSON only.",
    prompt,
    dropdownValuesByHeader,
    requests,
    "creator_profile_classification_batch"
  );
}

function buildCreatorProfileBatchPromptRows_(requests) {
  return (requests || []).map(function (request) {
    return {
      row_key: String(request.rowKey || "").trim(),
      requested_fields: Array.isArray(request.requestedFields) ? request.requestedFields : [],
      channel_name: String(request.channelName || "").trim(),
      channel_url: String(request.channelUrl || "").trim(),
      campaign_name: String(request.campaignName || "").trim(),
      creator_context: String(request.contextText || "")
    };
  });
}

function callOpenAiStructuredCreatorProfileBatch_(
  apiKey,
  model,
  systemPrompt,
  userPrompt,
  dropdownValuesByHeader,
  requests,
  schemaName
) {
  const rowKeys = (requests || []).map(function (request) {
    return String(request.rowKey || "").trim();
  }).filter(Boolean);
  if (rowKeys.length === 0) return {};

  const schema = buildCreatorProfileBatchResponseSchema_(dropdownValuesByHeader, rowKeys.length);
  const data = callOpenAiStructuredJson_(
    apiKey,
    model,
    systemPrompt,
    userPrompt,
    schema,
    String(schemaName || "creator_profile_batch")
  );
  if (!data || !Array.isArray(data.rows)) return {};

  const out = {};
  data.rows.forEach(function (row) {
    const rowKey = String(row && row.row_key || "").trim();
    if (!rowKey || rowKeys.indexOf(rowKey) === -1 || out[rowKey]) return;
    out[rowKey] = row;
  });
  return out;
}

function callOpenAiStructuredCreatorProfile_(apiKey, model, systemPrompt, userPrompt, dropdownValuesByHeader) {
  return callOpenAiStructuredJson_(
    apiKey,
    model,
    systemPrompt,
    userPrompt,
    buildCreatorProfileResponseSchema_(dropdownValuesByHeader),
    "creator_profile_enrichment"
  );
}

function callOpenAiStructuredJson_(apiKey, model, systemPrompt, userPrompt, schema, schemaName) {
  const responsesAttempt = callOpenAiResponsesJson_(
    apiKey,
    model,
    systemPrompt,
    userPrompt,
    schema,
    schemaName
  );
  if (responsesAttempt.ok) return responsesAttempt.data;

  Logger.log("ℹ️ OpenAI Responses API fallback triggered: " + responsesAttempt.diagnostic);

  const chatAttempt = callOpenAiChatCompletionsJson_(
    apiKey,
    model,
    systemPrompt,
    userPrompt,
    schema,
    schemaName
  );
  if (chatAttempt.ok) return chatAttempt.data;

  setLastOpenAiDiagnostic_(
    "failed",
    "responses=" + responsesAttempt.diagnostic + " | chat=" + chatAttempt.diagnostic
  );
  if (chatAttempt.diagnostic) {
    Logger.log("⚠️ OpenAI structured enrichment failed: " + chatAttempt.diagnostic);
  }
  return null;
}

function callOpenAiResponsesJson_(apiKey, model, systemPrompt, userPrompt, schema, schemaName) {
  const response = UrlFetchApp.fetch("https://api.openai.com/v1/responses", {
    method: "post",
    muteHttpExceptions: true,
    headers: {
      Authorization: "Bearer " + apiKey,
      "Content-Type": "application/json"
    },
    payload: JSON.stringify({
      model: model,
      instructions: systemPrompt,
      input: [
        {
          role: "user",
          content: [
            { type: "input_text", text: userPrompt }
          ]
        }
      ],
      text: {
        format: {
          type: "json_schema",
          name: String(schemaName || "creator_profile_enrichment"),
          strict: true,
          schema: schema
        }
      }
    })
  });

  const code = response.getResponseCode();
  const text = String(response.getContentText() || "");
  LAST_OPENAI_RAW_RESPONSE_ = text.slice(0, 1000);

  if (code >= 400) {
    const diagnostic = `responses ${code}: ${text.slice(0, 500)}`;
    setLastOpenAiDiagnostic_("responses_error", diagnostic);
    Logger.log("⚠️ OpenAI API error: " + diagnostic);
    return { ok: false, diagnostic: diagnostic };
  }

  try {
    const data = JSON.parse(text);
    const content = stripJsonFences_(extractOpenAiTextFromResponses_(data));
    if (!content) {
      const diagnostic = "responses empty output";
      setLastOpenAiDiagnostic_("responses_empty", diagnostic);
      return { ok: false, diagnostic: diagnostic };
    }

    const parsed = JSON.parse(content);
    setLastOpenAiDiagnostic_("responses_ok", "responses ok");
    return { ok: true, data: parsed, diagnostic: "responses ok" };
  } catch (e) {
    const diagnostic = "responses parse error: " + e;
    setLastOpenAiDiagnostic_("responses_parse_error", diagnostic);
    Logger.log("⚠️ LLM returned invalid JSON: " + e);
    return { ok: false, diagnostic: diagnostic };
  }
}

function callOpenAiChatCompletionsJson_(apiKey, model, systemPrompt, userPrompt, schema, schemaName) {
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
      response_format: {
        type: "json_schema",
        json_schema: {
          name: String(schemaName || "creator_profile_enrichment"),
          strict: true,
          schema: schema
        }
      }
    })
  });

  const code = response.getResponseCode();
  const text = String(response.getContentText() || "");
  LAST_OPENAI_RAW_RESPONSE_ = text.slice(0, 1000);

  if (code >= 400) {
    const diagnostic = `chat ${code}: ${text.slice(0, 500)}`;
    Logger.log("⚠️ OpenAI chat completions error: " + diagnostic);
    return { ok: false, diagnostic: diagnostic };
  }

  try {
    const data = JSON.parse(text);
    const content = stripJsonFences_(
      data &&
      data.choices &&
      data.choices[0] &&
      data.choices[0].message &&
      data.choices[0].message.content
    );
    if (!content) {
      return { ok: false, diagnostic: "chat empty output" };
    }

    const parsed = JSON.parse(content);
    setLastOpenAiDiagnostic_("chat_ok", "chat ok");
    return { ok: true, data: parsed, diagnostic: "chat ok" };
  } catch (e) {
    const diagnostic = "chat parse error: " + e;
    Logger.log("⚠️ OpenAI chat completions invalid JSON: " + e);
    return { ok: false, diagnostic: diagnostic };
  }
}

function setLastOpenAiDiagnostic_(status, details) {
  LAST_OPENAI_DIAGNOSTIC_ = [String(status || "").trim(), String(details || "").trim()]
    .filter(Boolean)
    .join(": ");
}

function buildCreatorProfileResponseSchema_(dropdownValuesByHeader) {
  return {
    type: "object",
    additionalProperties: false,
    required: PROFILE_LLM_FIELDS_,
    properties: {
      "First Name": { type: "string" },
      "Last Name": { type: "string" },
      "Email": { type: "string" },
      "Influencer Type": buildCreatorProfileScalarEnumSchema_(dropdownValuesByHeader && dropdownValuesByHeader["Influencer Type"]),
      "Influencer Vertical": buildCreatorProfileVerticalSchema_(dropdownValuesByHeader && dropdownValuesByHeader["Influencer Vertical"]),
      "Country/Region": buildCreatorProfileScalarEnumSchema_(dropdownValuesByHeader && dropdownValuesByHeader["Country/Region"]),
      "Language": buildCreatorProfileScalarEnumSchema_(dropdownValuesByHeader && dropdownValuesByHeader["Language"])
    }
  };
}

function buildCreatorProfileBatchResponseSchema_(dropdownValuesByHeader, rowCount) {
  return {
    type: "object",
    additionalProperties: false,
    required: ["rows"],
    properties: {
      rows: {
        type: "array",
        minItems: Number(rowCount) || 0,
        maxItems: Number(rowCount) || 0,
        items: buildCreatorProfileBatchRowSchema_(dropdownValuesByHeader)
      }
    }
  };
}

function buildCreatorProfileBatchRowSchema_(dropdownValuesByHeader) {
  return {
    type: "object",
    additionalProperties: false,
    required: ["row_key"].concat(PROFILE_LLM_FIELDS_),
    properties: {
      row_key: { type: "string" },
      "First Name": { type: "string" },
      "Last Name": { type: "string" },
      "Email": { type: "string" },
      "Influencer Type": buildCreatorProfileScalarEnumSchema_(dropdownValuesByHeader && dropdownValuesByHeader["Influencer Type"]),
      "Influencer Vertical": buildCreatorProfileVerticalSchema_(dropdownValuesByHeader && dropdownValuesByHeader["Influencer Vertical"]),
      "Country/Region": buildCreatorProfileScalarEnumSchema_(dropdownValuesByHeader && dropdownValuesByHeader["Country/Region"]),
      "Language": buildCreatorProfileScalarEnumSchema_(dropdownValuesByHeader && dropdownValuesByHeader["Language"])
    }
  };
}

function buildCreatorProfileScalarEnumSchema_(options) {
  const values = uniqueNonEmptyStrings_((options || []).concat([""]));
  if (values.length <= 1) return { type: "string" };
  return {
    type: "string",
    enum: values
  };
}

function buildCreatorProfileVerticalSchema_(options) {
  const values = uniqueNonEmptyStrings_(options || []);
  if (values.length === 0) {
    return {
      type: "array",
      items: { type: "string" },
      minItems: 0,
      maxItems: 3
    };
  }

  return {
    type: "array",
    items: {
      type: "string",
      enum: values
    },
    minItems: 0,
    maxItems: 3
  };
}

function sanitizeCreatorProfileEnrichment_(result, dropdownValuesByHeader) {
  const out = {};
  const firstName = normalizeLlmString_(result["First Name"]);
  const lastName = normalizeLlmString_(result["Last Name"]);

  if (isReasonablePersonName_(firstName) && isReasonablePersonName_(lastName)) {
    out["First Name"] = firstName;
    out["Last Name"] = lastName;
  }

  const email = coerceExplicitEmailValue_(result["Email"]);
  if (email) out["Email"] = email;

  const influencerType = coerceDropdownValue_(dropdownValuesByHeader["Influencer Type"] || [], result["Influencer Type"]);
  if (influencerType) out["Influencer Type"] = influencerType;

  const influencerVertical = coerceMultiSelectDropdownValues_(
    dropdownValuesByHeader["Influencer Vertical"] || [],
    result["Influencer Vertical"]
  );
  if (influencerVertical) out["Influencer Vertical"] = influencerVertical;

  const countryRegion = coerceDropdownValue_(dropdownValuesByHeader["Country/Region"] || [], result["Country/Region"]);
  if (countryRegion) out["Country/Region"] = countryRegion;

  const language = coerceDropdownValue_(dropdownValuesByHeader["Language"] || [], result["Language"]);
  if (language) out["Language"] = language;

  return out;
}

function applySanitizedProfileFields_(row, header, fields, sanitized, changeMap) {
  fields.forEach(field => {
    const value = sanitized[field];
    if (value == null || String(value).trim() === "") return;

    const i = findHeaderIndex_(header, field);
    if (i === -1 || String(row[i] || "").trim()) return;

    setTrackedValue_(row, i, value, changeMap);
  });
}

function buildCreatorProfileContext_(row, header, youtubeInsight) {
  const lines = PROFILE_LLM_CONTEXT_FIELDS_
    .map(field => {
      const value = normalizeLlmString_(getValueByHeader_(row, header, field));
      return value ? (field + ": " + value) : "";
    })
    .filter(Boolean);

  if (youtubeInsight) {
    if (youtubeInsight.description) lines.push("Resolved YouTube Description: " + youtubeInsight.description);
    if (youtubeInsight.countryCode) lines.push("Resolved YouTube Country Code: " + youtubeInsight.countryCode);
    if (youtubeInsight.dominantCategoryName) lines.push("Resolved YouTube Category: " + youtubeInsight.dominantCategoryName);
    if (youtubeInsight.preferredEmail) lines.push("Preferred Explicit Email: " + youtubeInsight.preferredEmail);
    if (youtubeInsight.preferredEmailSource) lines.push("Preferred Explicit Email Source: " + youtubeInsight.preferredEmailSource);
    if (youtubeInsight.bioEmails && youtubeInsight.bioEmails.length > 0) {
      lines.push("Explicit Emails from YouTube Bio: " + youtubeInsight.bioEmails.join(" | "));
    }
    if (youtubeInsight.channelPageEmails && youtubeInsight.channelPageEmails.length > 0) {
      lines.push("Explicit Emails from Channel Page: " + youtubeInsight.channelPageEmails.join(" | "));
    }
    if (youtubeInsight.videoDescriptionEmails && youtubeInsight.videoDescriptionEmails.length > 0) {
      lines.push("Explicit Emails from Video Descriptions: " + youtubeInsight.videoDescriptionEmails.join(" | "));
    }
    if (youtubeInsight.channelPageSnippet) {
      lines.push("Channel Page Snippet: " + truncateForUi_(youtubeInsight.channelPageSnippet, 400));
    }
    if (youtubeInsight.sampledTitles && youtubeInsight.sampledTitles.length > 0) {
      lines.push("Sampled Video Titles: " + youtubeInsight.sampledTitles.slice(0, 10).join(" | "));
    }
    if (youtubeInsight.sampledVideoDescriptions && youtubeInsight.sampledVideoDescriptions.length > 0) {
      lines.push(
        "Sampled Video Descriptions: " +
        youtubeInsight.sampledVideoDescriptions
          .slice(0, 3)
          .map(function (description) { return truncateForUi_(description, 280); })
          .join(" || ")
      );
    }
  }

  return lines.join("\n");
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

function coerceMultiSelectDropdownValues_(allowedValues, value) {
  if (!Array.isArray(allowedValues) || allowedValues.length === 0) return "";

  let rawValues = [];
  if (Array.isArray(value)) {
    rawValues = value;
  } else {
    const rawText = normalizeLlmString_(value);
    if (!rawText) return "";
    rawValues = rawText.split(/[;,|]/);
  }

  const normalizedValues = [];
  rawValues.forEach(item => {
    const coerced = coerceDropdownValue_(allowedValues, item);
    if (!coerced) return;
    if (normalizedValues.indexOf(coerced) !== -1) return;
    if (normalizedValues.length >= 3) return;
    normalizedValues.push(coerced);
  });

  return normalizedValues.join(SHEETS_MULTISELECT_DELIMITER_);
}

function normalizeHubSpotImportCellValue_(headerName, value) {
  const raw = String(value == null ? "" : value).trim();
  if (!raw) return "";
  if (!HUBSPOT_MULTISELECT_FIELDS_.has(String(headerName || "").trim())) return raw;

  return raw
    .split(/\s*[;,|]\s*/)
    .filter(Boolean)
    .join(HUBSPOT_MULTISELECT_DELIMITER_);
}

function coerceExplicitEmailValue_(value) {
  const matches = extractExplicitEmailsFromText_(value);
  if (matches.length > 0) return matches[0];

  const rawValue = normalizeExtractedEmailCandidate_(value);
  if (!rawValue) return "";
  return getHubSpotImportEmailValidationIssue_(rawValue) ? "" : rawValue;
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

function getCreatorProfileInsightForLlm_(row, header, apiKey, options) {
  if (!apiKey) return null;
  const channelUrl = resolveCreatorYouTubeInputFromRow_(row, header);
  const channelName = String(getValueByHeader_(row, header, "Channel Name") || "").trim();
  if (!channelUrl && !channelName) return null;

  return getCreatorYouTubeInsight_(channelUrl, channelName, apiKey, options);
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
    activeColIndexes.map(i => normalizeHubSpotImportCellValue_(header[i], item.values[i]))
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
  // mapping YouTube Average Views / YouTube Video Median Views → Median Views.
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

  // Column mapping: INT → EXT (rename EXT Rate→Rate; skip INT-only fields and EXT CPM)
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

  // Column mapping: EXT → INT, excluding the negotiation rate block and CPM.
  const colMap = buildCrossSheetColumnMap_(extHeader, intHeader, EXT_TO_INT_PITCHING_MAP_, EXT_ONLY_PITCHING_COLS_);

  let updated = 0;
  for (let r = 1; r < extData.length; r++) {
    const extRow = extData[r];
    if (isBlankRow_(extRow) || isSectionLabelRow_(extRow)) continue;

    const key = buildCompositeKey_(extRow, extHeader);
    if (!key) continue;

    const intRowIdx = intRowMap.get(key);
    if (intRowIdx === undefined) continue;

    const intRow = intData[intRowIdx];
    const changedIndexes = new Set();

    syncPitchingNegotiationBlock_(extRow, extHeader, intRow, intHeader).forEach(idx => changedIndexes.add(idx));

    for (const m of colMap) {
      const extVal = extRow[m.sourceIdx];
      if (String(extVal || "").trim() === "") continue;
      if (String(intRow[m.targetIdx] || "").trim() !== String(extVal || "").trim()) {
        intRow[m.targetIdx] = extVal;
        changedIndexes.add(m.targetIdx);
      }
    }

    if (writeChangedRowCells_(intPitching, intRowIdx + 1, intHeader, intRow, changedIndexes, INT_PITCHING_FORMULA_COLS_)) {
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
//  (6) UPDATE INT PERFORMANCE FROM EXT
// ============================================================

function updatePerformanceFromExt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const intPerformance = ss.getSheetByName("Performance");
  if (!intPerformance) return Logger.log("❌ 'Performance' sheet not found.");

  const extSs = getExtSpreadsheet_();
  if (!extSs) return;
  const extPerformance = extSs.getSheetByName("Performance");
  if (!extPerformance) return Logger.log("❌ 'Performance' sheet not found in EXT spreadsheet.");

  const intData = intPerformance.getDataRange().getValues();
  const extData = extPerformance.getDataRange().getValues();
  if (intData.length < 1 || extData.length < 2) {
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

  // Column mapping: EXT → INT using matching column names only.
  const colMap = buildCrossSheetColumnMap_(extHeader, intHeader, {}, new Set());
  const existingPerformanceSignatures = buildMappedRowSignatureSetFromTarget_(intData, colMap);

  let updated = 0;
  let inserted = 0;

  for (let r = 1; r < extData.length; r++) {
    const extRow = extData[r];
    if (isBlankRow_(extRow) || isSectionLabelRow_(extRow)) continue;

    const key = buildCompositeKey_(extRow, extHeader);
    const signature = buildMappedRowSignatureFromSource_(extRow, colMap);

    const intRowIdx = key ? intRowMap.get(key) : undefined;
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
        intPerformance.getRange(intRowIdx + 1, 1, 1, intHeader.length).setValues([intRow]);
        updated++;
      }
    } else {
      if (signature && existingPerformanceSignatures.has(signature)) continue;

      const intNumCols = intHeader.length;
      const out = new Array(intNumCols).fill("");
      for (const m of colMap) out[m.targetIdx] = extRow[m.sourceIdx];

      const insertAt1 = findFirstFreeRowFrom1_(intPerformance, 3, intNumCols);
      const targetRange = intPerformance.getRange(insertAt1, 1, 1, intNumCols);
      targetRange.setValues([out]);

      if (key) intRowMap.set(key, insertAt1 - 1);
      if (signature) existingPerformanceSignatures.add(signature);
      inserted++;
    }
  }

  Logger.log(`✅ INT Performance: ${updated} updated, ${inserted} inserted from EXT.`);
}

// ============================================================
//  CROSS-SPREADSHEET HELPERS
// ============================================================

function getExtSpreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  const dropdownValuesByHeader = getDropdownValuesByHeader_(ss, "Dropdown Values");
  const dropdownExtSheetValue = Array.isArray(dropdownValuesByHeader[EXT_SHEET_DROPDOWN_HEADER_]) &&
    dropdownValuesByHeader[EXT_SHEET_DROPDOWN_HEADER_].length > 0
    ? String(dropdownValuesByHeader[EXT_SHEET_DROPDOWN_HEADER_][0] || "").trim()
    : "";

  let extId = extractSpreadsheetId_(dropdownExtSheetValue);
  if (!extId) {
    extId = extractSpreadsheetId_(props.getProperty(EXT_SPREADSHEET_ID_PROP_));
  }

  if (!extId) {
    SpreadsheetApp.getUi().alert(
      "❌ EXT spreadsheet not configured.\n\n" +
      "Add an 'EXT Sheet' column in 'Dropdown Values' and put the EXT spreadsheet ID or URL in row 2."
    );
    return null;
  }

  props.setProperty(EXT_SPREADSHEET_ID_PROP_, extId);

  try {
    return SpreadsheetApp.openById(extId);
  } catch (e) {
    Logger.log("❌ Could not open EXT spreadsheet: " + e);
    SpreadsheetApp.getUi().alert("❌ Could not open EXT spreadsheet. Check the ID and permissions.");
    return null;
  }
}

function extractSpreadsheetId_(value) {
  const raw = String(value || "").trim().replace(/^['"]|['"]$/g, "");
  if (!raw) return "";

  const urlMatch = raw.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/i);
  if (urlMatch && urlMatch[1]) return urlMatch[1];

  if (/^[a-zA-Z0-9_-]{20,}$/.test(raw)) return raw;
  return "";
}


function buildCrossSheetColumnMap_(sourceHeader, targetHeader, renameMap, skipCols) {
  const map = [];
  const targetMap = new Map();
  const usedTargetIndexes = new Set();
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
    if (targetIdx !== undefined && !usedTargetIndexes.has(targetIdx)) {
      map.push({ sourceIdx: i, targetIdx: targetIdx, sourceName: sourceName, targetName: targetName });
      usedTargetIndexes.add(targetIdx);
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

function syncPitchingNegotiationBlock_(extRow, extHeader, intRow, intHeader) {
  const extStartIdx = extHeader.indexOf("Rate");
  const intStartIdx = intHeader.indexOf("EXT Rate");
  if (extStartIdx === -1 || intStartIdx === -1) return [];

  const changedIndexes = [];
  const maxWidth = Math.min(
    PITCHING_NEGOTIATION_BLOCK_WIDTH_,
    extHeader.length - extStartIdx,
    intHeader.length - intStartIdx
  );

  for (let offset = 0; offset < maxWidth; offset++) {
    const extColumnName = String(extHeader[extStartIdx + offset] || "").trim();
    const intColumnName = String(intHeader[intStartIdx + offset] || "").trim();
    if (
      extColumnName === "CPM" ||
      intColumnName === "EXT CPM" ||
      intColumnName === "INT CPM"
    ) {
      continue;
    }

    const extVal = extRow[extStartIdx + offset];
    if (String(extVal || "").trim() === "") continue;

    if (String(intRow[intStartIdx + offset] || "").trim() !== String(extVal || "").trim()) {
      intRow[intStartIdx + offset] = extVal;
      changedIndexes.push(intStartIdx + offset);
    }
  }

  return changedIndexes;
}

function writeChangedRowCells_(sheet, row1, header, row, changedIndexes, skipHeaders) {
  if (!changedIndexes || changedIndexes.size === 0) return false;

  let wrote = false;
  Array.from(changedIndexes)
    .sort((a, b) => a - b)
    .forEach(idx => {
      const columnName = String(header[idx] || "").trim();
      if (skipHeaders && skipHeaders.has(columnName)) return;

      sheet.getRange(row1, idx + 1).setValue(row[idx]);
      wrote = true;
    });

  return wrote;
}


function findRowSection_(data, rowIdx) {
  for (let r = rowIdx - 1; r >= 0; r--) {
    if (isSectionLabelRow_(data[r])) return String(data[r][0] || "").trim();
  }
  return "";
}


function getValueByHeader_(row, header, columnName) {
  const idx = findHeaderIndex_(header, columnName);
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
  if (!isKnownSectionLabel_(first)) return false;
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
    if (sourceMap.has(name)) common.push({ name: name, pIdx: sourceMap.get(name), cIdx: j });
  }
  return common;
}

function setTrackedValue_(row, colIndex, value, changeMap) {
  if (!Array.isArray(row) || colIndex < 0) return false;

  const nextValue = value == null ? "" : value;
  const currentValue = row[colIndex];
  const currentText = String(currentValue == null ? "" : currentValue);
  const nextText = String(nextValue);
  if (currentText === nextText) return false;

  row[colIndex] = nextValue;

  if (!changeMap) return true;

  const key = String(colIndex);
  if (!changeMap[key]) {
    changeMap[key] = {
      colIndex: colIndex,
      oldValue: currentValue,
      newValue: nextValue
    };
  } else {
    changeMap[key].newValue = nextValue;
  }

  const originalText = String(changeMap[key].oldValue == null ? "" : changeMap[key].oldValue);
  if (originalText === nextText) delete changeMap[key];
  return true;
}

function setTrackedValueByHeader_(row, header, columnName, value, changeMap) {
  const idx = findHeaderIndex_(header, columnName);
  if (idx === -1) return false;
  return setTrackedValue_(row, idx, value, changeMap);
}

function setIfEmpty_(row, header, columnName, value, changeMap) {
  const idx = findHeaderIndex_(header, columnName);
  if (idx === -1) return false;

  const nextValue = value == null ? "" : value;
  if (String(nextValue).trim() === "") return false;
  if (String(row[idx] || "").trim() !== "") return false;
  return setTrackedValue_(row, idx, nextValue, changeMap);
}

function trackedChangeMapToRowChanges_(changeMap) {
  if (!changeMap) return [];

  return Object.keys(changeMap)
    .map(function (key) { return changeMap[key]; })
    .filter(Boolean)
    .sort(function (a, b) { return a.colIndex - b.colIndex; });
}

function trackedRowChangesToSheetUpdates_(item, rowChanges) {
  if (!item || !item.row1 || !Array.isArray(rowChanges) || rowChanges.length === 0) return [];

  return rowChanges.map(function (change) {
    return {
      rowItem: item,
      row1: item.row1,
      colIndex: change.colIndex,
      oldValue: change.oldValue,
      newValue: change.newValue
    };
  });
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

function writeSparseCellUpdates_(sheet, header, updates, stageLabel) {
  if (!sheet || !Array.isArray(updates) || updates.length === 0) {
    return {
      writtenRowCount: 0,
      writtenCellCount: 0,
      skippedCellCount: 0
    };
  }

  const deduped = {};
  updates.forEach(function (update) {
    if (!update || !update.row1 || update.colIndex == null) return;

    const key = String(update.row1) + ":" + String(update.colIndex);
    if (!deduped[key]) {
      deduped[key] = {
        rowItem: update.rowItem,
        row1: update.row1,
        colIndex: update.colIndex,
        oldValue: update.oldValue,
        newValue: update.newValue
      };
      return;
    }

    deduped[key].newValue = update.newValue;
    if (update.rowItem) deduped[key].rowItem = update.rowItem;
  });

  const groupedByColumn = {};
  Object.keys(deduped).forEach(function (key) {
    const update = deduped[key];
    const normalizedOldValue = String(update.oldValue == null ? "" : update.oldValue);
    const normalizedNewValue = String(update.newValue == null ? "" : update.newValue);
    if (normalizedOldValue === normalizedNewValue) return;

    const columnKey = String(update.colIndex);
    if (!groupedByColumn[columnKey]) groupedByColumn[columnKey] = [];
    groupedByColumn[columnKey].push(update);
  });

  const writtenRows = {};
  let writtenCellCount = 0;
  let skippedCellCount = 0;

  Object.keys(groupedByColumn)
    .sort(function (a, b) { return Number(a) - Number(b); })
    .forEach(function (columnKey) {
      const columnUpdates = groupedByColumn[columnKey].sort(function (a, b) { return a.row1 - b.row1; });
      let startIndex = 0;

      while (startIndex < columnUpdates.length) {
        const block = [columnUpdates[startIndex]];
        let expectedRow1 = columnUpdates[startIndex].row1 + 1;
        let cursor = startIndex + 1;

        while (cursor < columnUpdates.length && columnUpdates[cursor].row1 === expectedRow1) {
          block.push(columnUpdates[cursor]);
          expectedRow1++;
          cursor++;
        }

        const colIndex = block[0].colIndex;
        const columnName = Array.isArray(header) ? String(header[colIndex] || "") : "";

        try {
          sheet
            .getRange(block[0].row1, colIndex + 1, block.length, 1)
            .setValues(block.map(function (update) { return [update.newValue]; }));

          block.forEach(function (update) {
            writtenRows[update.row1] = true;
            writtenCellCount++;
          });
        } catch (e) {
          Logger.log(
            "⚠️ " + String(stageLabel || "Enrichment") +
            " block write failed for column " +
            buildSheetColumnLabel_(colIndex) +
            (columnName ? (" (" + columnName + ")") : "") +
            " rows " + block[0].row1 + "-" + block[block.length - 1].row1 +
            ". Retrying per cell. " + e
          );

          block.forEach(function (update) {
            try {
              sheet.getRange(update.row1, update.colIndex + 1).setValue(update.newValue);
              writtenRows[update.row1] = true;
              writtenCellCount++;
            } catch (cellError) {
              if (update.rowItem && Array.isArray(update.rowItem.values)) {
                update.rowItem.values[update.colIndex] = update.oldValue;
              }
              skippedCellCount++;
              Logger.log(
                "⚠️ " + String(stageLabel || "Enrichment") +
                " skipped " + buildSheetColumnLabel_(update.colIndex) + update.row1 +
                (columnName ? (" (" + columnName + ")") : "") +
                " value \"" + truncateForUi_(String(update.newValue == null ? "" : update.newValue), 120) +
                "\": " + cellError
              );
            }
          });
        }

        startIndex = cursor;
      }
    });

  return {
    writtenRowCount: Object.keys(writtenRows).length,
    writtenCellCount: writtenCellCount,
    skippedCellCount: skippedCellCount
  };
}

function buildSheetColumnLabel_(colIndex) {
  let number = Number(colIndex) + 1;
  if (!isFinite(number) || number < 1) return "?";

  let out = "";
  while (number > 0) {
    const remainder = (number - 1) % 26;
    out = String.fromCharCode(65 + remainder) + out;
    number = Math.floor((number - 1) / 26);
  }
  return out;
}

function getCreatorListSheetContext_(sheet) {
  if (!sheet) return null;

  const data = sheet.getDataRange().getValues();
  if (!data || data.length < 2) return null;

  const headerRow0 = detectHeaderRow0ForColumns_(
    data,
    [
      "Channel Name",
      "Channel URL",
      "Campaign Name",
      "Email",
      "Status",
      "First Name",
      "Last Name",
      "Influencer Type",
      "Influencer Vertical"
    ],
    10
  );

  if (headerRow0 === -1) return null;

  const header = (data[headerRow0] || []).map(v => String(v || "").trim());
  const archivedStart0 = findSectionRowByLabel_(data, "Archived");

  return {
    data: data,
    header: header,
    headerRow0: headerRow0,
    headerRow1: headerRow0 + 1,
    archivedStart0: archivedStart0
  };
}

function detectHeaderRow0ForColumns_(data, expectedColumns, maxScanRows) {
  const scanRows = Math.min(Number(maxScanRows) || 10, data.length);
  let bestRow0 = -1;
  let bestScore = -1;

  for (let row0 = 0; row0 < scanRows; row0++) {
    const header = (data[row0] || []).map(v => String(v || "").trim());
    let score = 0;

    (expectedColumns || []).forEach(columnName => {
      if (findHeaderIndex_(header, columnName) !== -1) score++;
    });

    if (score > bestScore) {
      bestScore = score;
      bestRow0 = row0;
    }
  }

  return bestScore > 0 ? bestRow0 : -1;
}

function resolveCreatorYouTubeInputFromRow_(row, header) {
  const candidateFields = ["Channel URL", "YouTube URL", "YouTube Handle", "Channel Name"];

  for (let i = 0; i < candidateFields.length; i++) {
    const value = String(getValueByHeader_(row, header, candidateFields[i]) || "").trim();
    if (!looksLikeYouTubeInput_(value)) continue;

    if (value.charAt(0) === "@") {
      return "https://www.youtube.com/" + value;
    }
    return value;
  }

  return "";
}

function looksLikeYouTubeInput_(value) {
  const text = String(value || "").trim().toLowerCase();
  if (!text) return false;
  if (text.charAt(0) === "@") return true;
  return /youtube\.com|youtu\.be/.test(text);
}

function findHeaderIndex_(header, columnName) {
  const target = String(columnName || "").trim();
  if (!Array.isArray(header) || !target) return -1;

  const exactIdx = header.indexOf(target);
  if (exactIdx !== -1) return exactIdx;

  const normalizedTarget = normalizeHeaderName_(target);
  for (let i = 0; i < header.length; i++) {
    if (normalizeHeaderName_(header[i]) === normalizedTarget) return i;
  }

  return -1;
}

function normalizeHeaderName_(value) {
  return String(value || "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "");
}

function showSpreadsheetToast_(message) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (ss) ss.toast(String(message || ""), "Scripts", 8);
  } catch (e) {
    // Ignore UI/toast errors when run outside the spreadsheet UI.
  }
}

function showSpreadsheetAlert_(message) {
  try {
    SpreadsheetApp.getUi().alert(String(message || ""));
  } catch (e) {
    // Ignore UI/alert errors when run outside the spreadsheet UI.
  }
}

function truncateForUi_(value, maxLength) {
  const text = String(value == null ? "" : value);
  const limit = Math.max(Number(maxLength) || 0, 0);
  if (!limit || text.length <= limit) return text;
  return text.slice(0, Math.max(limit - 3, 0)) + "...";
}

function getProfileDropdownOptionsByHeader_(sheet, header, dropdownValuesByHeader) {
  const out = {};
  PROFILE_LLM_DROPDOWN_FIELDS_.forEach(field => {
    const explicitValues = Array.isArray(dropdownValuesByHeader[field]) ? dropdownValuesByHeader[field] : [];
    if (explicitValues.length > 0) {
      out[field] = explicitValues;
      return;
    }

    out[field] = getDropdownOptionsByHeader_(sheet, header, field, 2);
  });
  return out;
}

function getDropdownOptionsByHeader_(sheet, header, columnName, templateRow1) {
  const col0 = findHeaderIndex_(header, String(columnName || "").trim());
  if (col0 === -1) return [];

  const preferredRows = [templateRow1, 2, 3, 4, 5]
    .filter((row1, index, arr) => row1 >= 1 && arr.indexOf(row1) === index);

  for (let i = 0; i < preferredRows.length; i++) {
    const options = extractValidationOptions_(sheet.getRange(preferredRows[i], col0 + 1).getDataValidation());
    if (options.length > 0) return options;
  }

  const scanLimit = Math.min(sheet.getLastRow(), 200);
  for (let row1 = 2; row1 <= scanLimit; row1++) {
    const options = extractValidationOptions_(sheet.getRange(row1, col0 + 1).getDataValidation());
    if (options.length > 0) return options;
  }

  return [];
}

function extractValidationOptions_(rule) {
  if (!rule) return [];

  const criteriaType = rule.getCriteriaType();
  const args = rule.getCriteriaValues() || [];

  if (criteriaType === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
    return uniqueNonEmptyStrings_(args[0] || []);
  }

  if (criteriaType === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
    const range = args[0];
    if (!range) return [];
    return uniqueNonEmptyStrings_(flatten2d_(range.getDisplayValues()));
  }

  return [];
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

function getEnglishYouTubeCategoryName_(id) {
  const map = {
    1: "Film & Animation",
    2: "Autos & Vehicles",
    10: "Music",
    15: "Pets & Animals",
    17: "Sports",
    19: "Travel & Events",
    20: "Gaming",
    22: "People & Blogs",
    23: "Comedy",
    24: "Entertainment",
    25: "News & Politics",
    26: "Howto & Style",
    27: "Education",
    28: "Science & Technology"
  };
  return map[id] || String(id || "");
}

function getJsonFromScriptCache_(key) {
  const raw = CacheService.getScriptCache().get(String(key || ""));
  if (!raw) return null;

  try {
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function putJsonInScriptCache_(key, value, ttlSeconds) {
  CacheService.getScriptCache().put(
    String(key || ""),
    JSON.stringify(value),
    Number(ttlSeconds) || CREATOR_YOUTUBE_SIGNAL_TTL_SECONDS_
  );
}

function modeStringArray_(values) {
  const counts = {};
  (values || []).forEach(value => {
    const key = String(value || "").trim();
    if (!key) return;
    counts[key] = (counts[key] || 0) + 1;
  });

  return Object.keys(counts).sort((a, b) => counts[b] - counts[a])[0] || "";
}

function flatten2d_(matrix) {
  const out = [];
  (matrix || []).forEach(row => {
    (row || []).forEach(value => {
      out.push(value);
    });
  });
  return out;
}

function uniqueNonEmptyStrings_(values) {
  const out = [];
  const seen = {};

  (values || []).forEach(value => {
    const text = String(value || "").trim();
    if (!text || seen[text]) return;
    seen[text] = true;
    out.push(text);
  });

  return out;
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

function extractOpenAiTextFromResponses_(data) {
  if (data && data.output_text) return String(data.output_text || "");

  if (data && Array.isArray(data.output)) {
    const parts = [];
    data.output.forEach(item => {
      const content = item && Array.isArray(item.content) ? item.content : [];
      content.forEach(chunk => {
        if (chunk && chunk.type === "output_text" && chunk.text) {
          parts.push(String(chunk.text));
        }
      });
    });
    if (parts.length > 0) return parts.join("\n");
  }

  return "";
}

function stripJsonFences_(text) {
  const raw = String(text || "").trim();
  if (!raw) return raw;

  return raw
    .replace(/^```json\s*/i, "")
    .replace(/^```\s*/i, "")
    .replace(/\s*```$/, "")
    .trim();
}

function hasHubSpotSharedImporterLibrary_() {
  return (
    typeof HubSpotSharedImporter !== "undefined" &&
    HubSpotSharedImporter &&
    typeof HubSpotSharedImporter.startImport === "function"
  );
}
