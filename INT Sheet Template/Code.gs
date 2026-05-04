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
const HUBSPOT_IMPORT_MAX_ROWS_PER_PAYLOAD_ = 500;
const HUBSPOT_IMPORT_DELAY_MS_ = 1500;
const HUBSPOT_IMPORT_RETRY_COUNT_ = 4;
const HUBSPOT_IMPORT_RETRY_BASE_MS_ = 5000;
const HUBSPOT_API_BASE_ = "https://api.hubapi.com";
const HUBSPOT_API_KEY_ = "key here";
const HUBSPOT_API_KEY_PROP_ = "HUBSPOT_API_KEY";
const HUBSPOT_DEALS_OBJECT_TYPE_ID_ = "0-3";
const HUBSPOT_DEALS_OBJECT_API_NAME_ = "deals";
const HUBSPOT_CONTACTS_OBJECT_API_NAME_ = "contacts";
const HUBSPOT_DEAL_STAGE_PROPERTY_LABEL_ = "Deal Stage";
const HUBSPOT_PITCHING_STATUS_PROPERTY_LABEL_ = "Pitching Status";
const HUBSPOT_BATCH_UPDATE_SIZE_ = 100;
const HUBSPOT_REQUEST_MIN_INTERVAL_MS_ = 250;
const HUBSPOT_REQUEST_RETRY_COUNT_ = 5;
const HUBSPOT_REQUEST_RETRY_BASE_MS_ = 2000;
const HUBSPOT_PROPERTY_CACHE_TTL_SECONDS_ = 21600;
const HUBSPOT_PROPERTY_CACHE_MAX_CHARS_ = 90000;
const HUBSPOT_LAST_REQUEST_MS_PROP_ = "HUBSPOT_LAST_REQUEST_MS";
const HUBSPOT_ACTIVATION_OBJECT_KEY_ = "activations";
const HUBSPOT_DROPDOWN_VALUES_COLUMNS_ = {
  clientName: "Client name",
  client: "Client",
  dealOwner: "Deal Owner",
  campaignName: "Campaign Name",
  dealType: "Deal Type",
  activationType: "Activation Type",
  pipeline: "Pipeline",
  dealStage: "Deal Stage",
  currency: "Currency",
  influencerType: "Influencer Type",
  influencerVertical: "Influencer Vertical",
  contactType: "Contact Type",
  countryRegion: "Country/Region",
  language: "Language"
};
const HUBSPOT_DROPDOWN_VALUES_SYNC_CONFIG_ = {
  sheetName: "Dropdown Values",
  columns: HUBSPOT_DROPDOWN_VALUES_COLUMNS_,
  client: {
    objectTypeId: "",
    aliases: ["Client", "Clients"],
    propertyName: "Client Name"
  },
  campaign: {
    objectTypeId: "",
    aliases: ["Client Campaign", "Client Campaigns"],
    propertyName: "Campaign Name",
    createdWithinMonths: 3
  },
  activation: {
    objectTypeId: "",
    aliases: ["Activation", "Activations", HUBSPOT_ACTIVATION_OBJECT_KEY_]
  },
  optionColumns: [
    {
      columnName: HUBSPOT_DROPDOWN_VALUES_COLUMNS_.dealType,
      objectTypeId: HUBSPOT_DEALS_OBJECT_API_NAME_,
      propertyName: "Deal Type"
    },
    {
      columnName: HUBSPOT_DROPDOWN_VALUES_COLUMNS_.activationType,
      objectConfigKey: "activation",
      propertyName: "Activation Type"
    },
    {
      columnName: HUBSPOT_DROPDOWN_VALUES_COLUMNS_.influencerType,
      objectTypeId: HUBSPOT_CONTACTS_OBJECT_API_NAME_,
      propertyName: "Influencer Type"
    },
    {
      columnName: HUBSPOT_DROPDOWN_VALUES_COLUMNS_.influencerVertical,
      objectTypeId: HUBSPOT_CONTACTS_OBJECT_API_NAME_,
      propertyName: "Influencer Vertical"
    },
    {
      columnName: HUBSPOT_DROPDOWN_VALUES_COLUMNS_.contactType,
      objectTypeId: HUBSPOT_CONTACTS_OBJECT_API_NAME_,
      propertyName: "Contact Type"
    },
    {
      columnName: HUBSPOT_DROPDOWN_VALUES_COLUMNS_.countryRegion,
      objectTypeId: HUBSPOT_CONTACTS_OBJECT_API_NAME_,
      propertyName: "Country/Region"
    },
    {
      columnName: HUBSPOT_DROPDOWN_VALUES_COLUMNS_.language,
      objectTypeId: HUBSPOT_CONTACTS_OBJECT_API_NAME_,
      propertyName: "Language"
    }
  ]
};
const HUBSPOT_CAMPAIGN_SYNC_STAGE_DELAY_MS_ = 30000;
const HUBSPOT_INT_CAMPAIGN_DEAL_SYNC_STAGES_ = [
  {
    stageLabel: "Contract",
    mode: "first_non_empty",
    fields: [
      { propertyLabel: "Contract", sourceColumn: "Contract URL" }
    ]
  },
  {
    stageLabel: "Script Approved",
    mode: "all_rows_have_value",
    fields: [
      {
        propertyLabel: "Script Approved",
        sourceColumn: "Script URL",
        completeValue: "Yes",
        incompleteValue: "No"
      }
    ]
  },
  {
    stageLabel: "Preview Approved",
    mode: "all_rows_have_value",
    fields: [
      {
        propertyLabel: "Preview Approved",
        sourceColumn: "Preview URL",
        completeValue: "Yes",
        incompleteValue: "No"
      }
    ]
  },
  {
    stageLabel: "Published",
    mode: "all_rows_have_value",
    fields: [
      {
        propertyLabel: "Published",
        sourceColumn: "Activation URL",
        completeValue: "Yes",
        incompleteValue: "No"
      }
    ]
  }
];
const HUBSPOT_INT_CAMPAIGN_ACTIVATION_SYNC_STAGES_ = [
  {
    stageLabel: "Activation Fields",
    fields: [
      { propertyLabel: "Publication Date", sourceColumn: "Publication Date" },
      { propertyLabel: "Script URL", sourceColumn: "Script URL" },
      { propertyLabel: "Preview URL", sourceColumn: "Preview URL" },
      { propertyLabel: "Activation URL", sourceColumn: "Activation URL" },
      { propertyLabel: "Amount", sourceColumn: "INT Rate" },
      { propertyLabel: "EXT Amount", sourceColumn: "EXT Rate" }
    ]
  }
];
const HUBSPOT_INT_PERFORMANCE_ACTIVATION_SYNC_STAGES_ = [
  {
    stageLabel: "Performance Fields",
    fields: [
      { propertyLabel: "CPA", sourceColumn: "CPA" },
      { propertyLabel: "ROAS D7", sourceColumn: "ROAS D7" },
      { propertyLabel: "ROAS D30", sourceColumn: "ROAS D30" }
    ]
  }
];

const OPENAI_API_KEY_ = "key here";
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

// YouTube API
const YOUTUBE_API_KEY_ = "key here";
const YOUTUBE_API_KEY_PROP_ = "YOUTUBE_API_KEY";
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
const INT_PERFORMANCE_FORMULA_COLS_ = new Set(["CPM"]);
const PITCHING_NEGOTIATION_BLOCK_WIDTH_ = 5;
const CREATOR_TO_PITCHING_NEGOTIATION_SKIP_COLS_ = new Set(["Status"]);
const CREATOR_LIST_TIMESTAMP_IMPORTED_HEADER_ = "Timestamp Imported";
const WOODPECKER_EXPORT_HEADERS_ = ["Email", "Channel Name", "First Name", "Last Name"];
const CREATOR_LIST_ARCHIVE_SHEET_NAME_ = "Archive";
const CREATOR_LIST_STALE_ARCHIVE_DAYS_ = 30;
const CREATOR_LIST_ARCHIVE_FIRST_DATA_ROW_ = 3;
const CREATOR_LIST_ACTIVE_STATUSES_ = new Set(["contacted", "responded"]);
const HUBSPOT_IMPORT_EXCLUDED_COLS_ = new Set([
  "Channel Name",
  "HubSpot Record ID",
  "Channel URL",
  "Status",
  CREATOR_LIST_TIMESTAMP_IMPORTED_HEADER_,
  "Activation Type",
  "Activation name",
  "Activation Name"
]);
const HUBSPOT_IMPORT_MULTISELECT_COLS_ = new Set(["influencervertical"]);
const HUBSPOT_MULTISELECT_IMPORT_DELIMITER_ = ";";
const CREATOR_LIST_HUBSPOT_STAGE_TO_STATUS_ = {
  contacted: "Contacted",
  responded: "Responded"
};
const YOUTUBE_ENRICH_FIELDS_ = [
  "YouTube Handle",
  "YouTube URL",
  "YouTube Average Views",
  "YouTube Video Median Views",
  "YouTube Shorts Median Views",
  "YouTube Engagement Rate",
  "YouTube Followers"
];
const CREATOR_PLATFORM_SPECS_ = [
  {
    key: "youtube",
    label: "YouTube",
    handleHeader: "YouTube Handle",
    urlHeader: "YouTube URL",
    handlePrefix: "@"
  },
  {
    key: "instagram",
    label: "Instagram",
    handleHeader: "Instagram Handle",
    urlHeader: "Instagram URL",
    handlePrefix: "@"
  },
  {
    key: "tiktok",
    label: "TikTok",
    handleHeader: "TikTok Handle",
    urlHeader: "TikTok URL",
    handlePrefix: "@"
  },
  {
    key: "twitch",
    label: "Twitch",
    handleHeader: "Twitch Handle",
    urlHeader: "Twitch URL",
    handlePrefix: "@"
  },
  {
    key: "kick",
    label: "Kick",
    handleHeader: "Kick Handle",
    urlHeader: "Kick URL",
    handlePrefix: "@"
  },
  {
    key: "x",
    label: "X",
    handleHeader: "X Handle",
    urlHeader: "X URL",
    handlePrefix: "@"
  }
];
const PROFILE_LLM_FIELDS_ = [
  "First Name",
  "Last Name",
  "Email",
  "Phone Number",
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
  "Phone Number",
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
    .createMenu("📜 Scripts")
    .addItem("Enrich Creator List", "enrichCreatorListRowsInBatches")
    .addItem("Creator List → HubSpot", "importCreatorListToHubSpotOnly")
    .addItem("Woodpecker Export", "downloadCreatorListWoodpeckerCsv")
    .addSeparator()
    .addItem("Push to EXT Pitching", "pushReadyForPitchingToExt")
    .addToUi();

  SpreadsheetApp.getUi()
    .createMenu("⏰ Triggers")
    .addItem("Sync Dropdown Values from HubSpot", "syncDropdownValuesFromHubSpot")
    .addItem("Update Creator statuses from HubSpot", "updateCreatorListStatusesFromHubSpot")
    .addItem("Responded → Negotiation", "pushRespondedToNegotiation")
    .addItem("Archive Creator List", "archiveCreatorList")
    .addSeparator()
    .addItem("Update Pitching from EXT", "updatePitchingFromExt")
    .addItem("Update Campaigns from EXT", "updateCampaignsFromExt")
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

  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(CREATOR_LIST_ENRICH_CURSOR_PROP_);
  const header = context.header;
  const stopRow1 = context.archivedStart0 === -1 ? context.data.length : context.archivedStart0;
  let nextStartRow1 = context.headerRow1 + 1;
  let totalRows = 0;
  let totalStatic = 0;
  let totalYoutube = 0;
  let totalProfile = 0;

  while (nextStartRow1 <= stopRow1) {
    const batch = collectCreatorListEnrichmentBatch_(context.data, header, nextStartRow1, stopRow1);
    const rowItems = batch.rowItems;
    if (rowItems.length === 0) break;

    props.setProperty(CREATOR_LIST_ENRICH_CURSOR_PROP_, String(batch.nextStartRow1));
    const result = runCreatorListEnrichment_(sheet, rowItems, header);
    totalRows += rowItems.length;
    totalStatic += result.staticEnriched;
    totalYoutube += result.youtubeEnriched;
    totalProfile += result.profileEnriched;
    nextStartRow1 = batch.nextStartRow1;
  }

  props.deleteProperty(CREATOR_LIST_ENRICH_CURSOR_PROP_);
  if (totalRows === 0) {
    showSpreadsheetToast_("Creator List enrichment complete. No remaining rows to process.");
    return Logger.log("✅ Creator List enrichment complete. No remaining rows to process.");
  }

  showSpreadsheetToast_(
    `Enrichment complete. Rows: ${totalRows}. Static: ${totalStatic}. ` +
    `YouTube: ${totalYoutube}. AI: ${totalProfile}.`
  );
  Logger.log(
    `✅ Creator List enrichment complete. Rows: ${totalRows}. Static: ${totalStatic}, ` +
    `YouTube: ${totalYoutube}, AI: ${totalProfile}.`
  );
}

function runCreatorListEnrichment_(sheet, rowItems, header) {
  const dropdownValuesByHeader = getDropdownValuesByHeader_(sheet.getParent(), "Dropdown Values");
  const clientNameByCampaignName = getClientNameByCampaignName_(sheet.getParent(), "Dropdown Values");

  const staticUpdates = [];
  for (const item of rowItems) {
    const rowChanges = enrichCreatorRow_(item.values, header, clientNameByCampaignName);
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
  const platformFields = [];
  CREATOR_PLATFORM_SPECS_.forEach(function (spec) {
    platformFields.push(spec.handleHeader);
    platformFields.push(spec.urlHeader);
  });

  const candidateFields = [
    "Channel Name",
    "Channel URL",
    "Campaign Name",
    "Deal name",
    "YouTube URL",
    "YouTube Handle"
  ].concat(platformFields);

  return candidateFields.some(field => String(getValueByHeader_(row, header, field) || "").trim());
}

function collectCreatorListEnrichmentBatch_(data, header, startRow1, stopRow1) {
  const rowItems = [];
  let lastScannedRow1 = Math.max(Number(startRow1) || 1, 1) - 1;

  for (let row1 = Math.max(Number(startRow1) || 1, 1); row1 <= stopRow1; row1++) {
    lastScannedRow1 = row1;
    const row = data[row1 - 1];
    if (isBlankRow_(row) || isSectionLabelRow_(row)) continue;
    if (!canAttemptCreatorListEnrichment_(row, header)) continue;

    rowItems.push({ row1: row1, values: row.slice() });
    if (rowItems.length >= CREATOR_LIST_ENRICH_BATCH_SIZE_) break;
  }

  return {
    rowItems: rowItems,
    nextStartRow1: lastScannedRow1 + 1
  };
}


// ============================================================
//  (1) CREATOR LIST → HUBSPOT
// ============================================================

function importCreatorListToHubSpotOnly() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Creator List");
  if (!sheet) return ui.alert("❌ 'Creator List' sheet not found.");

  const context = getCreatorListSheetContext_(sheet);
  if (!context) return ui.alert("ℹ️ Creator List has no data.");

  const data = context.data;
  const header = context.header;

  const missingColumns = getMissingCreatorListHubSpotImportColumns_(header);
  if (missingColumns.length > 0) {
    return ui.alert(
      "❌ Creator List is missing required column(s):\n\n" +
      missingColumns.join("\n")
    );
  }

  const contactingStart0 = findSectionRowByLabel_(data, "Contacting");
  if (contactingStart0 === -1) return ui.alert("❌ 'Contacting' section not found.");
  const contactingEnd0 = findNextSectionStart0_(data, contactingStart0);
  const endRow0 = contactingEnd0 === -1 ? data.length : contactingEnd0;

  // Collect importable rows in Contacting section.
  const candidateRowItems = [];
  for (let r = contactingStart0 + 1; r < endRow0; r++) {
    const row = data[r];
    if (isBlankRow_(row) || isSectionLabelRow_(row)) continue;
    candidateRowItems.push({ row1: r + 1, values: row.slice() });
  }

  if (candidateRowItems.length === 0) return ui.alert("ℹ️ No importable rows in Contacting section.");

  const preparedRows = prepareCreatorListHubSpotImportRows_(header, candidateRowItems);
  if (preparedRows.rowItems.length === 0) {
    return ui.alert(
      "ℹ️ No new unique deals to import.\n\n" +
      "Rows with a filled HubSpot Record ID and duplicate Deal Name values were skipped."
    );
  }

  importCreatorListToHubSpot_(ss, sheet, header, preparedRows.rowItems, ui, preparedRows);
}

function enrichAndImportToHubSpot() {
  importCreatorListToHubSpotOnly();
}

function downloadCreatorListWoodpeckerCsv() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Creator List");
  if (!sheet) return ui.alert("❌ 'Creator List' sheet not found.");

  const context = getCreatorListSheetContext_(sheet);
  if (!context) return ui.alert("ℹ️ Creator List has no data.");

  const header = context.header;
  const missingColumns = WOODPECKER_EXPORT_HEADERS_.filter(function (columnName) {
    return findHeaderIndex_(header, columnName) === -1;
  });
  if (missingColumns.length > 0) {
    return ui.alert(
      "❌ Creator List is missing required column(s):\n\n" +
      missingColumns.join("\n")
    );
  }

  const exportRows = buildCreatorListWoodpeckerExportRows_(context.data, header, context.headerRow1);
  if (exportRows.length === 0) {
    return ui.alert("ℹ️ No creator rows found to export.");
  }

  const filename = buildWoodpeckerExportFilename_(ss.getName());
  showCsvDownloadDialog_(
    filename,
    toCsvBytes_([WOODPECKER_EXPORT_HEADERS_].concat(exportRows)),
    "Woodpecker Import"
  );
}

function updateCreatorListStatusesFromHubSpot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Creator List");
  if (!sheet) return Logger.log("❌ 'Creator List' sheet not found.");

  const context = getCreatorListSheetContext_(sheet);
  if (!context) return Logger.log("ℹ️ Creator List has no data.");

  const data = context.data;
  const header = context.header;
  const statusCol = findHeaderIndex_(header, "Status");
  const recordIdCol = findHeaderIndex_(header, "HubSpot Record ID");
  if (statusCol === -1) return Logger.log("❌ Creator List missing 'Status' column.");
  if (recordIdCol === -1) return Logger.log("❌ Creator List missing 'HubSpot Record ID' column.");

  const token = getHubSpotApiToken_();
  if (!token) return Logger.log("ℹ️ HubSpot token not set in this project. Skipping Creator List status refresh.");

  const rowItems = [];
  for (let r = context.headerRow0 + 1; r < data.length; r++) {
    const row = data[r];
    if (isBlankRow_(row) || isSectionLabelRow_(row)) continue;

    const recordId = String(row[recordIdCol] || "").trim();
    if (!recordId) continue;

    rowItems.push({
      row1: r + 1,
      recordId: recordId,
      values: row.slice()
    });
  }

  if (rowItems.length === 0) {
    return Logger.log("ℹ️ No Creator List rows with HubSpot Record ID found.");
  }

  let stageLabelsByDealId;
  try {
    stageLabelsByDealId = fetchHubSpotDealStageLabelsByIds_(
      token,
      rowItems.map(function (item) { return item.recordId; })
    );
  } catch (e) {
    return Logger.log("⚠️ Creator List status refresh failed: " + (e && e.stack ? e.stack : e));
  }

  const updates = [];
  let skippedUnmatchedStage = 0;
  let skippedNoChange = 0;
  const unmatchedStages = {};
  rowItems.forEach(function (item) {
    const stageLabel = stageLabelsByDealId[item.recordId] || "";
    const nextStatus = getCreatorListStatusForHubSpotDealStage_(stageLabel);
    if (!nextStatus) {
      skippedUnmatchedStage++;
      const stageKey = stageLabel || "(blank)";
      unmatchedStages[stageKey] = (unmatchedStages[stageKey] || 0) + 1;
      return;
    }

    const currentStatus = item.values[statusCol];
    if (String(currentStatus || "").trim() === nextStatus) {
      skippedNoChange++;
      return;
    }

    updates.push({
      rowItem: item,
      row1: item.row1,
      colIndex: statusCol,
      oldValue: currentStatus,
      newValue: nextStatus
    });
  });

  const writeResult = writeSparseCellUpdates_(
    sheet,
    header,
    updates,
    "Creator List HubSpot status refresh"
  );

  Logger.log(
    `✅ Creator List status refresh done. Updated ${writeResult.writtenRowCount} row(s). ` +
    `Rows with HubSpot ID: ${rowItems.length}. No change: ${skippedNoChange}. ` +
    `Ignored non-Contacted/Responded stages: ${skippedUnmatchedStage}. ` +
    `Skipped cells: ${writeResult.skippedCellCount}.`
  );

  if (skippedUnmatchedStage > 0) {
    Logger.log(
      "ℹ️ Unmatched HubSpot stage values: " +
      Object.keys(unmatchedStages)
        .slice(0, 10)
        .map(function (stage) { return stage + " (" + unmatchedStages[stage] + ")"; })
        .join(", ")
    );
  }
}

function archiveCreatorList() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const creatorSheet = ss.getSheetByName("Creator List");
  const archiveSheet = ss.getSheetByName(CREATOR_LIST_ARCHIVE_SHEET_NAME_);
  if (!creatorSheet || !archiveSheet) {
    const missing = [];
    if (!creatorSheet) missing.push("'Creator List'");
    if (!archiveSheet) missing.push("'" + CREATOR_LIST_ARCHIVE_SHEET_NAME_ + "'");
    const message = "❌ Missing sheet(s): " + missing.join(", ") + ".";
    showSpreadsheetAlert_(message);
    return Logger.log(message);
  }

  const creatorContext = getCreatorListSheetContext_(creatorSheet);
  if (!creatorContext) {
    const message = "ℹ️ Creator List has no data.";
    showSpreadsheetToast_(message);
    return Logger.log(message);
  }

  const creatorData = creatorContext.data;
  const creatorHeader = creatorContext.header;
  const statusCol = findHeaderIndex_(creatorHeader, "Status");
  const timestampCol = findHeaderIndex_(creatorHeader, CREATOR_LIST_TIMESTAMP_IMPORTED_HEADER_);
  if (statusCol === -1) {
    const message = "❌ Creator List missing 'Status' column.";
    showSpreadsheetAlert_(message);
    return Logger.log(message);
  }
  if (timestampCol === -1) {
    const message = "❌ Creator List missing '" + CREATOR_LIST_TIMESTAMP_IMPORTED_HEADER_ + "' column.";
    showSpreadsheetAlert_(message);
    return Logger.log(message);
  }

  const cutoffMs = Date.now() - CREATOR_LIST_STALE_ARCHIVE_DAYS_ * 24 * 60 * 60 * 1000;
  const staleRows = [];
  let skippedActiveStatus = 0;
  let skippedMissingTimestamp = 0;
  let skippedFreshTimestamp = 0;

  for (let r = creatorContext.headerRow0 + 1; r < creatorData.length; r++) {
    const row = creatorData[r];
    if (isBlankRow_(row) || isSectionLabelRow_(row)) continue;

    const status = String(row[statusCol] || "").trim().toLowerCase();
    if (CREATOR_LIST_ACTIVE_STATUSES_.has(status)) {
      skippedActiveStatus++;
      continue;
    }

    const importedAt = parseSpreadsheetDateValue_(row[timestampCol]);
    if (!importedAt) {
      skippedMissingTimestamp++;
      continue;
    }

    if (importedAt.getTime() > cutoffMs) {
      skippedFreshTimestamp++;
      continue;
    }

    staleRows.push({ row1: r + 1, values: row.slice() });
  }

  if (staleRows.length === 0) {
    const message = "ℹ️ Archive Creator List: no stale rows to archive.";
    showSpreadsheetToast_(message);
    return Logger.log(
      message +
      ` Active status: ${skippedActiveStatus}. Missing timestamp: ${skippedMissingTimestamp}. ` +
      `Younger than ${CREATOR_LIST_STALE_ARCHIVE_DAYS_} days: ${skippedFreshTimestamp}.`
    );
  }

  const archiveContext = getArchiveSheetContext_(archiveSheet, creatorHeader);
  if (!archiveContext) {
    const message = `❌ Could not resolve a header in '${CREATOR_LIST_ARCHIVE_SHEET_NAME_}' sheet.`;
    showSpreadsheetAlert_(message);
    return Logger.log(message);
  }

  const archiveHeader = archiveContext.header;
  const colMap = buildCrossSheetColumnMap_(creatorHeader, archiveHeader, {}, null);
  if (colMap.length === 0) {
    const message = `❌ No overlapping columns between Creator List and ${CREATOR_LIST_ARCHIVE_SHEET_NAME_}.`;
    showSpreadsheetAlert_(message);
    return Logger.log(message);
  }

  const rowsToWrite = staleRows.map(function (item) {
    const out = new Array(archiveHeader.length).fill("");
    colMap.forEach(function (mapping) {
      out[mapping.targetIdx] = item.values[mapping.sourceIdx];
    });
    return out;
  });

  const archiveDataStartRow1 = Math.max(
    Number(CREATOR_LIST_ARCHIVE_FIRST_DATA_ROW_) || 3,
    archiveContext.headerRow1 + 2
  );
  const insertAt1 = findFirstFreeRowFrom1_(
    archiveSheet,
    archiveDataStartRow1,
    archiveHeader.length
  );
  archiveSheet
    .getRange(insertAt1, 1, rowsToWrite.length, archiveHeader.length)
    .setValues(rowsToWrite);

  staleRows
    .slice()
    .sort(function (a, b) { return b.row1 - a.row1; })
    .forEach(function (item) {
      creatorSheet.deleteRow(item.row1);
    });

  const message = `✅ Archived ${staleRows.length} row(s) from Creator List to '${CREATOR_LIST_ARCHIVE_SHEET_NAME_}'.`;
  showSpreadsheetToast_(message);
  Logger.log(
    message +
    ` Kept by status: ${skippedActiveStatus}. Missing timestamp: ${skippedMissingTimestamp}. ` +
    `Younger than ${CREATOR_LIST_STALE_ARCHIVE_DAYS_} days: ${skippedFreshTimestamp}.`
  );
}


/**
 * Enriches a single Creator List row with derived fields.
 * Modifies the row array in place.
 */
function enrichCreatorRow_(row, header, clientNameByCampaignName) {
  const idx = name => findHeaderIndex_(header, name);
  const get = name => { const i = idx(name); return i === -1 ? "" : String(row[i] || "").trim(); };
  const changeMap = {};
  const set = (name, value) => setTrackedValueByHeader_(row, header, name, value, changeMap);
  const setIfEmpty = (name, value) => setIfEmpty_(row, header, name, value, changeMap);

  const channelName = get("Channel Name");
  const campaignName = get("Campaign Name");
  const platformIdentity = applyCreatorPlatformIdentityFields_(row, header, changeMap);
  if (!channelName && !platformIdentity && !campaignName) return [];

  // Contact Type: always Influencer
  set("Contact Type", "Influencer");

  const campaignParts = parseCampaignName_(campaignName);
  if (campaignParts) {
    setIfEmpty("Month", campaignParts.month);
    setIfEmpty("Year", campaignParts.year);
  }

  const clientName = getClientNameForCampaign_(campaignName, clientNameByCampaignName);
  setIfEmpty(idx("Client name") !== -1 ? "Client name" : "Client", clientName);

  const creatorLabel = getPreferredCreatorLabelForRow_(row, header, platformIdentity);
  applyCreatorCampaignNames_(row, header, creatorLabel, campaignName, changeMap);

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

function applyCreatorPlatformIdentityFields_(row, header, changeMap) {
  const identities = collectCreatorPlatformIdentitiesFromRow_(row, header);
  const identityByKey = {};

  identities.forEach(function (identity) {
    if (!identity || !identity.key || identityByKey[identity.key]) return;
    identityByKey[identity.key] = identity;
  });

  CREATOR_PLATFORM_SPECS_.forEach(function (spec) {
    const identity = identityByKey[spec.key];
    if (!identity) return;

    if (identity.handle) {
      setIfEmpty_(row, header, spec.handleHeader, identity.handle, changeMap);
    }
    if (identity.canonicalUrl && (identity.handle || identity.key === "youtube")) {
      setIfEmpty_(row, header, spec.urlHeader, identity.canonicalUrl, changeMap);
    }
  });

  const preferredHandle = getPreferredCreatorHandleForRow_(row, header, identities.length > 0 ? identities[0] : null);
  if (preferredHandle) {
    setIfEmpty_(row, header, "Channel Name", preferredHandle, changeMap);
  }

  return identities.length > 0 ? identities[0] : null;
}

function collectCreatorPlatformIdentitiesFromRow_(row, header) {
  const identities = [];
  const channelUrlIdentity = parseCreatorPlatformUrl_(getValueByHeader_(row, header, "Channel URL"));
  if (channelUrlIdentity) identities.push(channelUrlIdentity);

  CREATOR_PLATFORM_SPECS_.forEach(function (spec) {
    const urlIdentity = parseCreatorPlatformUrl_(getValueByHeader_(row, header, spec.urlHeader));
    if (urlIdentity && urlIdentity.key === spec.key) {
      identities.push(urlIdentity);
    }

    const handle = normalizeCreatorPlatformHandle_(spec.key, getValueByHeader_(row, header, spec.handleHeader));
    if (handle) {
      identities.push({
        key: spec.key,
        spec: spec,
        handle: handle,
        canonicalUrl: buildCreatorPlatformCanonicalUrl_(spec.key, handle)
      });
    }
  });

  return identities;
}

function getPreferredCreatorHandleForRow_(row, header, preferredIdentity) {
  const primaryIdentity = preferredIdentity || parseCreatorPlatformUrl_(getValueByHeader_(row, header, "Channel URL"));
  if (primaryIdentity) {
    const primarySpec = getCreatorPlatformSpec_(primaryIdentity.key);
    const primaryHandle = primarySpec
      ? normalizeCreatorPlatformHandle_(primarySpec.key, getValueByHeader_(row, header, primarySpec.handleHeader))
      : "";
    if (primaryHandle) return primaryHandle;
    if (primaryIdentity.handle) return primaryIdentity.handle;
  }

  const identities = collectCreatorPlatformIdentitiesFromRow_(row, header);
  for (let i = 0; i < identities.length; i++) {
    if (identities[i].handle) return identities[i].handle;
  }

  return "";
}

function getPreferredCreatorLabelForRow_(row, header, preferredIdentity) {
  const preferredHandle = getPreferredCreatorHandleForRow_(row, header, preferredIdentity);
  if (preferredHandle) return preferredHandle;
  return String(getValueByHeader_(row, header, "Channel Name") || "").trim();
}

function applyCreatorCampaignNames_(row, header, creatorLabel, campaignName, changeMap) {
  const label = String(creatorLabel || "").trim();
  const campaign = String(campaignName || "").trim();
  if (!label || !campaign) return false;

  const name = label + " - " + campaign;
  let changed = false;
  changed = setTrackedValueByHeader_(row, header, "Deal name", name, changeMap) || changed;
  changed = setTrackedValueByHeader_(row, header, "Activation name", name, changeMap) || changed;
  return changed;
}

function parseCreatorPlatformUrl_(value) {
  const parsed = parseBasicUrl_(value);
  if (!parsed) return null;

  const host = parsed.host;
  const pathParts = parsed.pathParts;
  let key = "";
  let handle = "";

  if (isYouTubeHost_(host)) {
    key = "youtube";
    handle = parseYouTubeHandleFromPathParts_(pathParts);
  } else if (isInstagramHost_(host)) {
    key = "instagram";
    handle = parseInstagramHandleFromPathParts_(pathParts);
  } else if (isTikTokHost_(host)) {
    key = "tiktok";
    handle = parseTikTokHandleFromPathParts_(pathParts);
  } else if (isTwitchHost_(host)) {
    key = "twitch";
    handle = parseFirstPathHandle_(pathParts, [
      "activate", "bits", "broadcast", "collections", "directory", "downloads",
      "drops", "event", "friends", "jobs", "login", "logout", "moderator",
      "p", "popout", "prime", "products", "settings", "store", "subscriptions",
      "teams", "turbo", "videos", "wallet"
    ]);
  } else if (isKickHost_(host)) {
    key = "kick";
    handle = parseFirstPathHandle_(pathParts, [
      "about", "auth", "browse", "categories", "category", "community-guidelines",
      "dashboard", "dmca", "following", "jobs", "login", "privacy", "search",
      "signup", "terms", "video", "videos"
    ]);
  } else if (isXHost_(host)) {
    key = "x";
    handle = parseFirstPathHandle_(pathParts, [
      "about", "compose", "download", "explore", "hashtag", "home", "i",
      "intent", "jobs", "login", "messages", "notifications", "privacy",
      "search", "settings", "share", "tos"
    ]);
  }

  if (!key) return null;

  const spec = getCreatorPlatformSpec_(key);
  const normalizedHandle = normalizeCreatorPlatformHandle_(key, handle);
  return {
    key: key,
    spec: spec,
    handle: normalizedHandle,
    canonicalUrl: normalizedHandle ? buildCreatorPlatformCanonicalUrl_(key, normalizedHandle) : normalizeChannelUrl_(parsed.url)
  };
}

function parseBasicUrl_(value) {
  const raw = String(value || "").trim();
  if (!raw) return null;

  const normalized = /^https?:\/\//i.test(raw) ? raw : ("https://" + raw);
  const match = normalized.match(/^https?:\/\/([^\/?#]+)(\/[^?#]*)?/i);
  if (!match) return null;

  const host = String(match[1] || "").toLowerCase().replace(/^www\./, "");
  const path = String(match[2] || "");
  const pathParts = path
    .split("/")
    .filter(Boolean)
    .map(function (part) {
      try { return decodeURIComponent(part); } catch (e) { return part; }
    });

  return {
    url: normalized,
    host: host,
    pathParts: pathParts
  };
}

function parseYouTubeHandleFromPathParts_(pathParts) {
  if (!pathParts || pathParts.length === 0) return "";
  const first = String(pathParts[0] || "").trim();
  if (first.charAt(0) === "@") return first;
  return "";
}

function parseInstagramHandleFromPathParts_(pathParts) {
  if (!pathParts || pathParts.length === 0) return "";
  const first = String(pathParts[0] || "").trim();
  const second = String(pathParts[1] || "").trim();
  if (first.toLowerCase() === "stories" && second) return second;
  return parseFirstPathHandle_(pathParts, [
    "about", "accounts", "api", "challenge", "developer", "direct", "explore",
    "legal", "oauth", "p", "privacy", "reel", "reels", "stories", "terms", "tv"
  ]);
}

function parseTikTokHandleFromPathParts_(pathParts) {
  if (!pathParts || pathParts.length === 0) return "";

  for (let i = 0; i < pathParts.length; i++) {
    const part = String(pathParts[i] || "").trim();
    if (part.charAt(0) === "@") return part;
  }

  return "";
}

function parseFirstPathHandle_(pathParts, reservedWords) {
  if (!pathParts || pathParts.length === 0) return "";
  const first = String(pathParts[0] || "").trim();
  if (!first) return "";

  const normalizedFirst = first.toLowerCase();
  const reserved = reservedWords || [];
  if (reserved.indexOf(normalizedFirst) !== -1) return "";
  return first;
}

function normalizeCreatorPlatformHandle_(platformKey, value) {
  const spec = getCreatorPlatformSpec_(platformKey);
  let text = String(value || "").trim();
  if (!spec || !text) return "";

  text = text
    .replace(/^[<("'`\[]+/, "")
    .replace(/[>"')\],;:!?]+$/, "")
    .replace(/^https?:\/\//i, "")
    .trim();

  const parsed = parseCreatorPlatformUrl_(text);
  if (parsed && parsed.key === platformKey && parsed.handle) return parsed.handle;

  text = text.replace(/^@+/, "").replace(/^\/+|\/+$/g, "").trim();
  if (!text) return "";

  if (platformKey === "youtube") {
    const firstPart = text.split(/[/?#]/)[0];
    if (!/^[A-Za-z0-9._-]{2,100}$/.test(firstPart)) return "";
    return "@" + firstPart;
  }

  if (platformKey === "instagram") {
    const firstPart = text.split(/[/?#]/)[0];
    if (!/^[A-Za-z0-9._]{1,30}$/.test(firstPart)) return "";
    return "@" + firstPart;
  }

  if (platformKey === "tiktok") {
    const firstPart = text.split(/[/?#]/)[0];
    if (!/^[A-Za-z0-9._]{1,32}$/.test(firstPart)) return "";
    return "@" + firstPart;
  }

  if (platformKey === "twitch") {
    const firstPart = text.split(/[/?#]/)[0];
    if (!/^[A-Za-z0-9_]{1,25}$/.test(firstPart)) return "";
    return "@" + firstPart;
  }

  if (platformKey === "kick") {
    const firstPart = text.split(/[/?#]/)[0];
    if (!/^[A-Za-z0-9_.-]{1,40}$/.test(firstPart)) return "";
    return "@" + firstPart;
  }

  if (platformKey === "x") {
    const firstPart = text.split(/[/?#]/)[0];
    if (!/^[A-Za-z0-9_]{1,15}$/.test(firstPart)) return "";
    return "@" + firstPart;
  }

  return "";
}

function buildCreatorPlatformCanonicalUrl_(platformKey, handle) {
  const normalizedHandle = normalizeCreatorPlatformHandle_(platformKey, handle);
  if (!normalizedHandle) return "";

  const pathHandle = normalizedHandle.replace(/^@/, "");
  if (platformKey === "youtube") return "https://www.youtube.com/@" + pathHandle;
  if (platformKey === "instagram") return "https://www.instagram.com/" + pathHandle + "/";
  if (platformKey === "tiktok") return "https://www.tiktok.com/@" + pathHandle;
  if (platformKey === "twitch") return "https://www.twitch.tv/" + pathHandle;
  if (platformKey === "kick") return "https://kick.com/" + pathHandle;
  if (platformKey === "x") return "https://x.com/" + pathHandle;
  return "";
}

function getCreatorPlatformSpec_(platformKey) {
  const key = String(platformKey || "").trim();
  for (let i = 0; i < CREATOR_PLATFORM_SPECS_.length; i++) {
    if (CREATOR_PLATFORM_SPECS_[i].key === key) return CREATOR_PLATFORM_SPECS_[i];
  }
  return null;
}

function isYouTubeHost_(host) {
  return /(^|\.)youtube\.com$|(^|\.)youtu\.be$/i.test(String(host || ""));
}

function isInstagramHost_(host) {
  return /(^|\.)instagram\.com$/i.test(String(host || ""));
}

function isTikTokHost_(host) {
  return /(^|\.)tiktok\.com$/i.test(String(host || ""));
}

function isTwitchHost_(host) {
  return /(^|\.)twitch\.tv$/i.test(String(host || ""));
}

function isKickHost_(host) {
  return /(^|\.)kick\.com$/i.test(String(host || ""));
}

function isXHost_(host) {
  return /(^|\.)x\.com$|(^|\.)twitter\.com$/i.test(String(host || ""));
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
    if (!shouldRunYouTubeEnrichmentForRow_(item.values, header)) continue;

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

function shouldRunYouTubeEnrichmentForRow_(row, header) {
  const channelUrlIdentity = parseCreatorPlatformUrl_(getValueByHeader_(row, header, "Channel URL"));
  const hasExplicitYouTubeInput = hasExplicitYouTubeInputForRow_(row, header);
  const channelName = String(getValueByHeader_(row, header, "Channel Name") || "").trim();
  if (channelUrlIdentity && channelUrlIdentity.key !== "youtube" && !hasExplicitYouTubeInput) {
    return false;
  }
  if (!hasExplicitYouTubeInput && !channelName) return false;

  const needsYoutubeFields = YOUTUBE_ENRICH_FIELDS_.some(field => {
    const fieldIdx = findHeaderIndex_(header, field);
    return fieldIdx !== -1 && !String(row[fieldIdx] || "").trim();
  });
  if (needsYoutubeFields) return true;

  const youtubeHandle = String(getValueByHeader_(row, header, "YouTube Handle") || "").trim();
  if (youtubeHandle && youtubeHandle !== channelName) return true;

  const campaignName = String(getValueByHeader_(row, header, "Campaign Name") || "").trim();
  const dealName = String(getValueByHeader_(row, header, "Deal name") || "").trim();
  const creatorLabel = getPreferredCreatorLabelForRow_(row, header, channelUrlIdentity);
  if (creatorLabel && campaignName && dealName !== (creatorLabel + " - " + campaignName)) return true;

  return false;
}

function hasExplicitYouTubeInputForRow_(row, header) {
  const channelUrlIdentity = parseCreatorPlatformUrl_(getValueByHeader_(row, header, "Channel URL"));
  if (channelUrlIdentity && channelUrlIdentity.key === "youtube") return true;

  const youtubeUrlIdentity = parseCreatorPlatformUrl_(getValueByHeader_(row, header, "YouTube URL"));
  if (youtubeUrlIdentity && youtubeUrlIdentity.key === "youtube") return true;

  return !!normalizeCreatorPlatformHandle_("youtube", getValueByHeader_(row, header, "YouTube Handle"));
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
  const preferredHandle = getPreferredCreatorHandleForRow_(row, header, null) || String(insight.handle || "").trim();
  if (preferredHandle) {
    setIfEmpty_(row, header, "Channel Name", preferredHandle, changeMap);
  }

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

  const campaignName = String(getValueByHeader_(row, header, "Campaign Name") || "").trim();
  const creatorLabel = getPreferredCreatorLabelForRow_(row, header, null) || String(insight.handle || "").trim();
  applyCreatorCampaignNames_(row, header, creatorLabel, campaignName, changeMap);

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

function extractPreferredCreatorPhoneFromContext_(row, header, youtubeInsight) {
  const description = youtubeInsight && youtubeInsight.description
    ? String(youtubeInsight.description || "")
    : "";
  const phones = extractExplicitPhoneNumbersFromText_(description);

  return {
    phoneNumber: phones.length > 0 ? phones[0] : "",
    source: phones.length > 0 ? "channel bio" : ""
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
    const pattern = /(?:mailto:)?([A-Z0-9.!#$%&'*+=?^_`{|}~-]+@[A-Z0-9.-]+\.[A-Z]{2,63})/ig;
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
    Logger.log("ℹ️ OpenAI API key not set. AI enrichment will be skipped, but explicit email and phone extraction will still run.");
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
      const needsPhone = emptyFields.indexOf("Phone Number") !== -1;
      let youtubeInsight = item.youtubeInsight || null;

      if (youtubeInsight) {
        youtubeInsight = decorateCreatorInsightWithEmailSignals_(youtubeInsight, needsEmail);
        item.youtubeInsight = youtubeInsight;
      } else if ((apiKey || needsEmail || needsPhone) && validatedYouTubeApiKey) {
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

      if (needsPhone) {
        const preferredPhoneSignal = extractPreferredCreatorPhoneFromContext_(item.values, header, youtubeInsight);
        if (preferredPhoneSignal.phoneNumber) {
          applySanitizedProfileFields_(
            item.values,
            header,
            ["Phone Number"],
            { "Phone Number": preferredPhoneSignal.phoneNumber },
            changeMap
          );
        }

        emptyFields = PROFILE_LLM_FIELDS_.filter(f => {
          const i = findHeaderIndex_(header, f);
          return i !== -1 && !String(item.values[i] || "").trim();
        });
      }

      const llmFields = emptyFields.filter(f => f !== "Phone Number");
      if (llmFields.length === 0) {
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
        requestedFields: llmFields.slice(),
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
    "Only include Phone Number when an explicit phone number appears in the provided evidence.",
    "Return Phone Number as a plain phone number only. Preserve a leading + and normal separators when present.",
    "Do not use follower counts, view counts, dates, IDs, or social handles as phone numbers.",
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
    "Only include Phone Number when an explicit phone number appears in the provided evidence.",
    "Return Phone Number as a plain phone number only. Preserve a leading + and normal separators when present.",
    "Do not use follower counts, view counts, dates, IDs, or social handles as phone numbers.",
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
      "Phone Number": { type: "string" },
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
      "Phone Number": { type: "string" },
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

  const phoneNumber = coerceExplicitPhoneNumberValue_(result["Phone Number"]);
  if (phoneNumber) out["Phone Number"] = phoneNumber;

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
    const value = sanitizeProfileFieldValueForWrite_(field, sanitized[field]);
    if (value == null || String(value).trim() === "") return;

    const i = findHeaderIndex_(header, field);
    if (i === -1 || String(row[i] || "").trim()) return;

    setTrackedValue_(row, i, value, changeMap);
  });
}

function sanitizeProfileFieldValueForWrite_(field, value) {
  if (normalizeHeaderName_(field) === normalizeHeaderName_("Email")) {
    return coerceExplicitEmailValue_(value);
  }
  if (normalizeHeaderName_(field) === normalizeHeaderName_("Phone Number")) {
    return coerceExplicitPhoneNumberValue_(value);
  }
  return value;
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

function coerceExplicitEmailValue_(value) {
  const matches = extractExplicitEmailsFromText_(value);
  if (matches.length > 0) return matches[0];

  const rawValue = normalizeExtractedEmailCandidate_(value);
  if (!rawValue) return "";
  return getHubSpotImportEmailValidationIssue_(rawValue) ? "" : rawValue;
}

function coerceExplicitPhoneNumberValue_(value) {
  const matches = extractExplicitPhoneNumbersFromText_(value);
  return matches.length > 0 ? matches[0] : "";
}

function extractExplicitPhoneNumbersFromText_(value) {
  const raw = String(value || "");
  if (!raw) return [];

  const normalizedText = raw
    .replace(/&plus;|&#43;|&#x2b;/ig, "+")
    .replace(/\btel:\s*/ig, "");
  const results = [];
  const seen = {};
  const pattern = /(?:\+|00)?\d[\d\s().-]{5,}\d(?:\s*(?:ext\.?|extension|x)\s*\d{1,6})?/ig;
  let match;

  while ((match = pattern.exec(normalizedText)) !== null) {
    const phoneNumber = normalizeExtractedPhoneNumberCandidate_(match[0]);
    if (!phoneNumber) continue;

    const key = phoneNumber.replace(/\D/g, "");
    if (seen[key]) continue;
    seen[key] = true;
    results.push(phoneNumber);
  }

  return results;
}

function normalizeExtractedPhoneNumberCandidate_(value) {
  let text = String(value || "")
    .replace(/^[<("'`\[]+/, "")
    .replace(/[>"')\],;:!?]+$/, "")
    .replace(/\s+/g, " ")
    .trim();

  if (!text) return "";

  const extensionMatch = text.match(/\b(?:ext\.?|extension|x)\s*(\d{1,6})\b/i);
  const extension = extensionMatch ? extensionMatch[1] : "";
  if (extensionMatch) text = text.slice(0, extensionMatch.index).trim();

  text = text
    .replace(/[^\d+().\-\s]/g, "")
    .replace(/\s+/g, " ")
    .replace(/\(\s+/g, "(")
    .replace(/\s+\)/g, ")")
    .trim();

  if (!text) return "";
  if ((text.match(/\+/g) || []).length > 1) return "";
  if (text.indexOf("+") > 0) return "";

  const digitCount = (text.match(/\d/g) || []).length;
  if (digitCount < 7 || digitCount > 20) return "";
  if (!/^\+?[\d\s().-]+$/.test(text)) return "";
  if (/^\d{1,4}[-/.]\d{1,2}[-/.]\d{1,4}$/.test(text)) return "";

  return extension ? text + " ext " + extension : text;
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
  const configured = String(
    PropertiesService.getScriptProperties().getProperty(OPENAI_API_KEY_PROP_) || ""
  ).trim();
  return configured || String(OPENAI_API_KEY_ || "").trim();
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
  const channelUrlIdentity = parseCreatorPlatformUrl_(getValueByHeader_(row, header, "Channel URL"));
  if (!channelUrl && channelUrlIdentity && channelUrlIdentity.key !== "youtube") return null;

  const channelName = String(getValueByHeader_(row, header, "Channel Name") || "").trim();
  if (!channelUrl && !channelName) return null;

  return getCreatorYouTubeInsight_(channelUrl, channelName, apiKey, options);
}


// ---- HubSpot import ----

function importCreatorListToHubSpot_(ss, sheet, header, rowItems, ui, importSummary) {
  const emailCol = findHeaderIndex_(header, "Email");

  const validationIssues = collectCreatorListHubSpotImportValidationIssues_(header, rowItems, emailCol);
  if (validationIssues.length > 0) {
    ui.alert("❌ HubSpot import validation failed:\n\n" + validationIssues.join("\n"));
    return;
  }

  const payloads = buildHubSpotImportPayloads_(ss, header, rowItems, emailCol);
  if (payloads.length === 0) {
    ui.alert("❌ No importable HubSpot columns found.");
    return;
  }

  const payloadValidationError = validateHubSpotImportPayloads_(payloads);
  if (payloadValidationError) {
    ui.alert("❌ HubSpot import validation failed:\n\n" + payloadValidationError);
    return;
  }

  const saveMarkers = function (dealRecordIds) {
    return saveImportedCreatorListDealMarkers_(sheet, header, dealRecordIds);
  };

  // Try shared library
  if (hasHubSpotSharedImporterLibrary_()) {
    try {
      const result = runHubSpotSharedLibraryImports_(payloads, saveMarkers);
      ui.alert(buildImportSuccessMessage_(
        result.importResults,
        result.savedRecordIds,
        result.savedTimestamps,
        importSummary
      ));
      return;
    } catch (e) {
      ui.alert("❌ HubSpot import failed: " + e.message);
      return;
    }
  }

  // Try shared web app
  const webAppUrl = String(HUBSPOT_SHARED_IMPORT_WEB_APP_URL_ || "").trim();
  if (webAppUrl) {
    try {
      const result = runHubSpotSharedWebAppImports_(payloads, webAppUrl, saveMarkers);
      ui.alert(buildImportSuccessMessage_(
        result.importResults,
        result.savedRecordIds,
        result.savedTimestamps,
        importSummary
      ));
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

function getMissingCreatorListHubSpotImportColumns_(header) {
  return ["Deal Name", "HubSpot Record ID", CREATOR_LIST_TIMESTAMP_IMPORTED_HEADER_]
    .filter(function (columnName) {
      return findHeaderIndex_(header, columnName) === -1;
    });
}

function prepareCreatorListHubSpotImportRows_(header, candidateRowItems) {
  const recordIdCol = findHeaderIndex_(header, "HubSpot Record ID");
  const dealNameCol = findHeaderIndex_(header, "Deal Name");
  const importedDealNames = {};

  candidateRowItems.forEach(function (item) {
    const recordId = String(item.values[recordIdCol] || "").trim();
    const dealName = String(item.values[dealNameCol] || "").trim();
    if (recordId && dealName) importedDealNames[dealName] = true;
  });

  const seenDealNames = {};
  const rowItems = [];
  let skippedWithRecordId = 0;
  let skippedDuplicateDealName = 0;

  candidateRowItems.forEach(function (item) {
    const recordId = String(item.values[recordIdCol] || "").trim();
    const dealName = String(item.values[dealNameCol] || "").trim();

    if (recordId) {
      skippedWithRecordId++;
      return;
    }

    if (dealName && importedDealNames[dealName]) {
      skippedDuplicateDealName++;
      return;
    }

    if (dealName && seenDealNames[dealName]) {
      skippedDuplicateDealName++;
      return;
    }

    if (dealName) seenDealNames[dealName] = true;
    rowItems.push(item);
  });

  return {
    rowItems: rowItems,
    skippedWithRecordId: skippedWithRecordId,
    skippedDuplicateDealName: skippedDuplicateDealName
  };
}

function collectCreatorListHubSpotImportValidationIssues_(header, rowItems, emailCol) {
  const issues = [];
  const dealNameCol = findHeaderIndex_(header, "Deal Name");

  rowItems.forEach(item => {
    const dealName = String(item.values[dealNameCol] || "").trim();
    if (!dealName) {
      issues.push(`Row ${item.row1}, column "Deal Name": is blank.`);
    }

    if (emailCol === -1) return;
    const email = String(item.values[emailCol] || "").trim();
    if (!email) return;

    const emailIssue = getHubSpotImportEmailValidationIssue_(email);
    if (emailIssue) {
      issues.push(`Row ${item.row1}, column "Email": "${email}" ${emailIssue}`);
    }
  });

  if (issues.length > 15) {
    issues.splice(15, issues.length - 15, `${issues.length - 15} more validation issue(s) not shown.`);
  }

  return issues;
}

function validateHubSpotImportPayloads_(payloads) {
  if (
    typeof HubSpotSharedImporter === "undefined" ||
    !HubSpotSharedImporter ||
    typeof HubSpotSharedImporter.validateImport !== "function"
  ) {
    return "";
  }

  const errors = [];
  payloads.forEach(function (payload) {
    const result = HubSpotSharedImporter.validateImport(payload);
    if (!result || result.ok !== true) {
      const label = String(payload.importLabel || "").trim();
      errors.push(
        (label ? label + ": " : "") +
        (result && result.error ? result.error : "Shared HubSpot importer validation failed.")
      );
    }
  });

  return errors.join("\n\n");
}

function buildHubSpotImportPayloads_(ss, header, rowItems, emailCol) {
  if (emailCol === -1) {
    const payload = buildHubSpotImportPayload_(ss, header, rowItems, null, "");
    return payload ? [payload] : [];
  }

  const rowsWithEmail = [];
  const rowsWithoutEmail = [];
  rowItems.forEach(item => {
    const email = String(item.values[emailCol] || "").trim();
    (email ? rowsWithEmail : rowsWithoutEmail).push(item);
  });

  const payloads = [];
  const shouldLabelPayloads = rowsWithEmail.length > 0 && rowsWithoutEmail.length > 0;
  if (rowsWithEmail.length > 0) {
    const payload = buildHubSpotImportPayload_(
      ss,
      header,
      rowsWithEmail,
      null,
      shouldLabelPayloads ? "with email" : ""
    );
    if (payload) payloads.push(payload);
  }

  if (rowsWithoutEmail.length > 0) {
    const payload = buildHubSpotImportPayload_(
      ss,
      header,
      rowsWithoutEmail,
      new Set([emailCol]),
      shouldLabelPayloads ? "without email" : ""
    );
    if (payload) payloads.push(payload);
  }

  return splitHubSpotImportPayloadsByRows_(payloads);
}

function buildHubSpotImportPayload_(ss, header, rowItems, excludedColIndexes, importLabel) {
  const activeColIndexes = [];
  header.forEach((h, idx) => {
    if (!h) return;
    if (excludedColIndexes && excludedColIndexes.has(idx)) return;
    if (shouldExcludeCreatorListColumnFromHubSpotImport_(h)) return;
    const hasValue = rowItems.some(item =>
      String(formatHubSpotImportCellValue_(h, item.values[idx]) || "").trim() !== ""
    );
    if (!hasValue) return;
    activeColIndexes.push(idx);
  });

  if (activeColIndexes.length === 0 || rowItems.length === 0) return null;

  const activeHeaders = activeColIndexes.map(i => header[i]);
  const activeRows = rowItems.map(item =>
    activeColIndexes.map(i => formatHubSpotImportCellValue_(header[i], item.values[i]))
  );
  const safeSpreadsheetName = ss.getName().replace(/[\\/:*?"<>|]+/g, "-");
  const label = String(importLabel || "").trim();

  return {
    sheetName: "Creator List",
    spreadsheetName: label ? safeSpreadsheetName + " - " + label : safeSpreadsheetName,
    spreadsheetLocale: ss.getSpreadsheetLocale(),
    headers: activeHeaders,
    rows: activeRows,
    sourceRowNumbers: rowItems.map(item => item.row1),
    rowCount: activeRows.length,
    columnCount: activeHeaders.length,
    importLabel: label
  };
}

function splitHubSpotImportPayloadsByRows_(payloads) {
  const out = [];
  (payloads || []).forEach(function (payload) {
    splitHubSpotImportPayloadByRows_(payload).forEach(function (chunk) {
      out.push(chunk);
    });
  });
  return out;
}

function splitHubSpotImportPayloadByRows_(payload) {
  const maxRows = Math.max(1, Number(HUBSPOT_IMPORT_MAX_ROWS_PER_PAYLOAD_) || 500);
  const rows = Array.isArray(payload && payload.rows) ? payload.rows : [];
  if (!payload || rows.length <= maxRows) return payload ? [payload] : [];

  const sourceRowNumbers = Array.isArray(payload.sourceRowNumbers) ? payload.sourceRowNumbers : [];
  const chunks = [];
  const totalChunks = Math.ceil(rows.length / maxRows);

  for (let start = 0; start < rows.length; start += maxRows) {
    const chunkIndex = chunks.length + 1;
    const chunkLabel = buildHubSpotImportChunkLabel_(payload.importLabel, chunkIndex, totalChunks);
    chunks.push(Object.assign({}, payload, {
      spreadsheetName:
        String(payload.spreadsheetName || "HubSpot Import") +
        " - part " + chunkIndex + " of " + totalChunks,
      rows: rows.slice(start, start + maxRows),
      sourceRowNumbers: sourceRowNumbers.slice(start, start + maxRows),
      rowCount: Math.min(maxRows, rows.length - start),
      importLabel: chunkLabel
    }));
  }

  return chunks;
}

function buildHubSpotImportChunkLabel_(label, chunkIndex, totalChunks) {
  const base = String(label || "").trim();
  const suffix = "part " + chunkIndex + "/" + totalChunks;
  return base ? base + " " + suffix : suffix;
}

function runHubSpotSharedLibraryImports_(payloads, onPayloadImported) {
  const importResults = [];
  const dealRecordIds = [];
  let savedRecordIds = 0;
  let savedTimestamps = 0;

  payloads.forEach((payload, index) => {
    if (index > 0) Utilities.sleep(HUBSPOT_IMPORT_DELAY_MS_);

    const result = runHubSpotSharedLibraryImportWithRetry_(payload);
    if (result && result.ok) {
      importResults.push(normalizeHubSpotImportResult_(payload, result));
      appendHubSpotDealRecordIds_(dealRecordIds, result.dealRecordIds);
      if (typeof onPayloadImported === "function") {
        const markerResult = onPayloadImported(result.dealRecordIds);
        savedRecordIds += Number(markerResult && markerResult.recordIds || 0);
        savedTimestamps += Number(markerResult && markerResult.timestamps || 0);
      }
      return;
    }

    throw new Error(result && result.error ? result.error : "Library import failed.");
  });

  return {
    importResults: importResults,
    dealRecordIds: dealRecordIds,
    savedRecordIds: savedRecordIds,
    savedTimestamps: savedTimestamps
  };
}

function runHubSpotSharedLibraryImportWithRetry_(payload) {
  let lastError = null;
  const maxRetries = Math.max(0, Number(HUBSPOT_IMPORT_RETRY_COUNT_) || 0);

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    try {
      const result = HubSpotSharedImporter.startImport(payload);
      if (result && result.ok) return result;

      throw new Error(result && result.error ? result.error : "Library import failed.");
    } catch (e) {
      lastError = e;
      if (!isRetryableHubSpotImportError_(e) || attempt >= maxRetries) throw e;

      sleepBeforeHubSpotImportRetry_(attempt);
    }
  }

  throw lastError || new Error("Library import failed.");
}

function runHubSpotSharedWebAppImports_(payloads, webAppUrl, onPayloadImported) {
  const importResults = [];
  const dealRecordIds = [];
  let savedRecordIds = 0;
  let savedTimestamps = 0;

  payloads.forEach((payload, index) => {
    if (index > 0) Utilities.sleep(HUBSPOT_IMPORT_DELAY_MS_);

    const parsed = runHubSpotSharedWebAppImportWithRetry_(payload, webAppUrl);
    importResults.push(normalizeHubSpotImportResult_(payload, parsed));
    appendHubSpotDealRecordIds_(dealRecordIds, parsed.dealRecordIds);
    if (typeof onPayloadImported === "function") {
      const markerResult = onPayloadImported(parsed.dealRecordIds);
      savedRecordIds += Number(markerResult && markerResult.recordIds || 0);
      savedTimestamps += Number(markerResult && markerResult.timestamps || 0);
    }
  });

  return {
    importResults: importResults,
    dealRecordIds: dealRecordIds,
    savedRecordIds: savedRecordIds,
    savedTimestamps: savedTimestamps
  };
}

function runHubSpotSharedWebAppImportWithRetry_(payload, webAppUrl) {
  let lastError = null;
  const maxRetries = Math.max(0, Number(HUBSPOT_IMPORT_RETRY_COUNT_) || 0);

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
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

      return parsed;
    } catch (e) {
      lastError = e;
      if (!isRetryableHubSpotImportError_(e) || attempt >= maxRetries) throw e;

      sleepBeforeHubSpotImportRetry_(attempt);
    }
  }

  throw lastError || new Error("Web app import failed.");
}

function isRetryableHubSpotImportError_(error) {
  const message = String(error && error.message ? error.message : error || "");
  return /Bandwidth quota exceeded|Service invoked too many times|Try Utilities\.sleep|rate limit|RATE_LIMIT|TEN_SECONDLY_ROLLING|429|temporarily|timeout|timed out|Address unavailable|Socket|Connection/i.test(message);
}

function sleepBeforeHubSpotImportRetry_(attempt) {
  const base = Math.max(500, Number(HUBSPOT_IMPORT_RETRY_BASE_MS_) || 5000);
  const delay = Math.min(60000, base * Math.pow(2, Math.max(0, Number(attempt) || 0)));
  Utilities.sleep(delay);
}

function normalizeHubSpotImportResult_(payload, result) {
  return {
    label: String(payload.importLabel || "").trim(),
    rowCount: Number(payload.rowCount || 0),
    importId: result && result.importId,
    state: result && result.state
  };
}

function appendHubSpotDealRecordIds_(out, dealRecordIds) {
  if (!Array.isArray(out) || !Array.isArray(dealRecordIds)) return;
  dealRecordIds.forEach(item => out.push(item));
}

function shouldExcludeCreatorListColumnFromHubSpotImport_(headerName) {
  const text = String(headerName || "").trim();
  if (!text) return true;
  if (HUBSPOT_IMPORT_EXCLUDED_COLS_.has(text)) return true;

  const normalized = normalizeHeaderName_(text);
  return normalized === "activationtype" || normalized === "activationname";
}


function saveImportedCreatorListDealMarkers_(sheet, header, dealRecordIds) {
  const recordIdCol = findHeaderIndex_(header, "HubSpot Record ID");
  const timestampCol = findHeaderIndex_(header, CREATOR_LIST_TIMESTAMP_IMPORTED_HEADER_);
  if (
    recordIdCol === -1 ||
    timestampCol === -1 ||
    !Array.isArray(dealRecordIds) ||
    dealRecordIds.length === 0
  ) {
    return { recordIds: 0, timestamps: 0 };
  }

  const now = new Date();
  let savedRecordIds = 0;
  let savedTimestamps = 0;

  dealRecordIds.forEach(function (item) {
    const row1 = Number(item && item.sourceRowNumber);
    const recordId = String(item && item.recordId || "").trim();
    if (!isFinite(row1) || row1 < 2 || !recordId) return;

    const recordIdCell = sheet.getRange(row1, recordIdCol + 1);
    if (String(recordIdCell.getValue() || "").trim() !== recordId) {
      recordIdCell.setValue(recordId);
      savedRecordIds++;
    }

    const timestampCell = sheet.getRange(row1, timestampCol + 1);
    const timestampValue = timestampCell.getValue();
    if (
      !parseSpreadsheetDateValue_(timestampValue) &&
      String(timestampValue == null ? "" : timestampValue).trim() === ""
    ) {
      timestampCell.setValue(now);
      savedTimestamps++;
    }
  });

  return {
    recordIds: savedRecordIds,
    timestamps: savedTimestamps
  };
}

function formatHubSpotImportCellValue_(headerName, value) {
  const text = String(value == null ? "" : value).trim();
  if (!text) return "";

  if (!HUBSPOT_IMPORT_MULTISELECT_COLS_.has(normalizeHeaderName_(headerName))) {
    return text;
  }

  return normalizeHubSpotMultiselectImportValue_(text);
}

function normalizeHubSpotMultiselectImportValue_(value) {
  const seen = {};
  const values = String(value || "")
    .split(/[;,]/)
    .map(function (part) { return part.trim(); })
    .filter(function (part) {
      if (!part || seen[part]) return false;
      seen[part] = true;
      return true;
    });

  return values.join(HUBSPOT_MULTISELECT_IMPORT_DELIMITER_);
}

function buildImportSuccessMessage_(importResults, savedRecordIds, savedTimestamps, importSummary) {
  const results = Array.isArray(importResults) ? importResults : [];
  const rowCount = results.reduce((sum, result) => sum + Number(result.rowCount || 0), 0);
  const allDone = results.length > 0 && results.every(result => String(result.state || "").trim() === "DONE");
  const suffix = allDone
    ? "HubSpot finished the import and the returned deal IDs were saved."
    : "HubSpot continues processing the import in the background.";
  const importLines = results.map(result => {
    const label = String(result.label || "").trim();
    const prefix = label ? label + ": " : "";
    const importState = String(result.state || "STARTED").trim() || "STARTED";
    return (
      "- " + prefix +
      "Unique deals " + Number(result.rowCount || 0) +
      ", Import ID " + (result.importId || "N/A") +
      ", State " + importState
    );
  });
  const skippedWithRecordId = Number(importSummary && importSummary.skippedWithRecordId || 0);
  const skippedDuplicateDealName = Number(importSummary && importSummary.skippedDuplicateDealName || 0);

  return (
    "✅ HubSpot import submitted.\n\n" +
    "Unique deals imported: " + rowCount + "\n" +
    "Imports:\n" + importLines.join("\n") + "\n" +
    "HubSpot Record IDs saved: " + Number(savedRecordIds || 0) + "\n\n" +
    CREATOR_LIST_TIMESTAMP_IMPORTED_HEADER_ + " saved: " + Number(savedTimestamps || 0) + "\n\n" +
    "Skipped rows with HubSpot Record ID: " + skippedWithRecordId + "\n" +
    "Skipped duplicate Deal Name rows: " + skippedDuplicateDealName + "\n\n" +
    suffix
  );
}

function buildCreatorListWoodpeckerExportRows_(data, header, headerRow1) {
  const out = [];
  const startRow0 = Math.max(Number(headerRow1) || 1, 1);

  for (let r = startRow0; r < data.length; r++) {
    const row = data[r];
    if (isBlankRow_(row) || isSectionLabelRow_(row)) continue;

    const exportRow = WOODPECKER_EXPORT_HEADERS_.map(function (columnName) {
      const value = getValueByHeader_(row, header, columnName);
      return String(value == null ? "" : value);
    });
    if (exportRow.join("").trim() === "") continue;
    out.push(exportRow);
  }

  return out;
}

function buildWoodpeckerExportFilename_(spreadsheetName) {
  const safeSpreadsheetName = String(spreadsheetName || "Creator List")
    .replace(/[\\/:*?"<>|]+/g, "-")
    .trim();
  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd_HH-mm-ss"
  );
  return safeSpreadsheetName + " - Woodpecker Import - " + timestamp + ".csv";
}

function showCsvDownloadDialog_(filename, csvBytes, title) {
  const safeFilename = String(filename || "export.csv").replace(/\\/g, "\\\\").replace(/"/g, '\\"');
  const base64 = Utilities.base64Encode(csvBytes || []);
  const html = HtmlService.createHtmlOutput(
    '<!DOCTYPE html>' +
    '<html><head><base target="_top"><meta charset="utf-8"></head>' +
    '<body style="font-family:Arial,sans-serif;padding:16px;">' +
    '<p style="margin:0 0 12px;">Your CSV download should start automatically.</p>' +
    '<a id="download-link" download="' + safeFilename + '">Download CSV</a>' +
    '<script>' +
    '(function(){' +
    'var base64="' + base64 + '";' +
    'var binary=atob(base64);' +
    'var bytes=new Uint8Array(binary.length);' +
    'for(var i=0;i<binary.length;i++){bytes[i]=binary.charCodeAt(i);}' +
    'var blob=new Blob([bytes],{type:"text/csv;charset=utf-8"});' +
    'var url=URL.createObjectURL(blob);' +
    'var link=document.getElementById("download-link");' +
    'link.href=url;' +
    'setTimeout(function(){link.click();},0);' +
    'setTimeout(function(){URL.revokeObjectURL(url);if(typeof google!=="undefined"&&google.script&&google.script.host){google.script.host.close();}},1500);' +
    '})();' +
    '</script></body></html>'
  )
    .setWidth(320)
    .setHeight(120);

  SpreadsheetApp.getUi().showModalDialog(html, String(title || "CSV Download"));
}

function fetchHubSpotDealStageLabelsByIds_(token, recordIds) {
  const ids = uniqueNonEmptyStrings_(recordIds);
  if (ids.length === 0) return {};

  let propertyName = "dealstage";
  try {
    propertyName =
      resolveHubSpotDealPropertyName_(token, HUBSPOT_DEAL_STAGE_PROPERTY_LABEL_) ||
      propertyName;
  } catch (e) {
    Logger.log("⚠️ Could not resolve HubSpot Deal Stage property. Using 'dealstage'. " + e);
  }

  let optionLabelByValue = {};
  try {
    optionLabelByValue = fetchHubSpotPropertyOptionLabelMap_(
      HUBSPOT_DEALS_OBJECT_TYPE_ID_,
      propertyName,
      token
    );
  } catch (e) {
    Logger.log("⚠️ Could not load HubSpot Deal Stage options. Raw stage values will be used. " + e);
  }

  let pipelineStageLabelMaps = { byPipelineStage: {}, byStage: {} };
  try {
    pipelineStageLabelMaps = fetchHubSpotDealPipelineStageLabelMaps_(token);
  } catch (e) {
    Logger.log("⚠️ Could not load HubSpot deal pipeline stages. Raw stage values will be used. " + e);
  }

  const out = {};
  const chunks = chunkArray_(ids, HUBSPOT_BATCH_UPDATE_SIZE_);

  chunks.forEach(function (chunk) {
    try {
      const data = hubspotRequestJson_(
        HUBSPOT_API_BASE_ + "/crm/v3/objects/deals/batch/read",
        token,
        {
          method: "post",
          payload: JSON.stringify({
            properties: [propertyName, "pipeline"],
            inputs: chunk.map(function (id) { return { id: id }; })
          })
        }
      );
      const results = Array.isArray(data && data.results) ? data.results : [];
      results.forEach(function (item) {
        const id = String(item && item.id || "").trim();
        if (!id) return;
        out[id] = getHubSpotDealStageLabelFromObject_(
          item,
          propertyName,
          optionLabelByValue,
          pipelineStageLabelMaps
        );
      });
    } catch (batchError) {
      Logger.log("⚠️ HubSpot Deal Stage batch read failed. Retrying per deal. " + batchError);
      chunk.forEach(function (id) {
        try {
          out[id] = fetchHubSpotDealStageLabelById_(
            token,
            id,
            propertyName,
            optionLabelByValue,
            pipelineStageLabelMaps
          );
        } catch (singleError) {
          Logger.log(`⚠️ HubSpot Deal Stage read failed for deal ${id}: ${singleError}`);
        }
      });
    }
  });

  return out;
}

function fetchHubSpotDealStageLabelById_(
  token,
  recordId,
  propertyName,
  optionLabelByValue,
  pipelineStageLabelMaps
) {
  const data = hubspotRequestJson_(
    HUBSPOT_API_BASE_ +
      "/crm/v3/objects/deals/" +
      encodeURIComponent(recordId) +
      "?properties=" +
      encodeURIComponent(propertyName) +
      "&properties=pipeline",
    token
  );
  return getHubSpotDealStageLabelFromObject_(
    data,
    propertyName,
    optionLabelByValue,
    pipelineStageLabelMaps
  );
}

function getHubSpotDealStageLabelFromObject_(
  item,
  propertyName,
  optionLabelByValue,
  pipelineStageLabelMaps
) {
  const properties = item && item.properties ? item.properties : {};
  const rawStage = String(properties[propertyName] || "").trim();
  const pipelineId = String(properties.pipeline || "").trim();
  const byPipelineStage = pipelineStageLabelMaps && pipelineStageLabelMaps.byPipelineStage
    ? pipelineStageLabelMaps.byPipelineStage
    : {};
  const byStage = pipelineStageLabelMaps && pipelineStageLabelMaps.byStage
    ? pipelineStageLabelMaps.byStage
    : {};

  if (pipelineId && rawStage && byPipelineStage[pipelineId + "||" + rawStage]) {
    return byPipelineStage[pipelineId + "||" + rawStage];
  }
  if (rawStage && byStage[rawStage]) return byStage[rawStage];
  return (optionLabelByValue && optionLabelByValue[rawStage]) || rawStage;
}

function fetchHubSpotDealPipelineStageLabelMaps_(token) {
  const data = hubspotRequestJson_(
    HUBSPOT_API_BASE_ +
      "/crm/v3/pipelines/" +
      encodeURIComponent(HUBSPOT_DEALS_OBJECT_API_NAME_),
    token
  );
  const pipelines = Array.isArray(data && data.results) ? data.results : [];
  const out = {
    byPipelineStage: {},
    byStage: {}
  };

  pipelines.forEach(function (pipeline) {
    const pipelineId = String(pipeline && pipeline.id || "").trim();
    const stages = Array.isArray(pipeline && pipeline.stages) ? pipeline.stages : [];
    stages.forEach(function (stage) {
      const stageId = String(stage && stage.id || "").trim();
      const stageLabel = String(stage && stage.label || "").trim();
      if (!stageId || !stageLabel) return;

      if (pipelineId) out.byPipelineStage[pipelineId + "||" + stageId] = stageLabel;
      if (!out.byStage[stageId]) out.byStage[stageId] = stageLabel;
    });
  });

  return out;
}

function fetchHubSpotPropertyOptionLabelMap_(objectType, propertyName, token) {
  const data = hubspotRequestJson_(
    HUBSPOT_API_BASE_ +
      "/crm/v3/properties/" +
      encodeURIComponent(objectType) +
      "/" +
      encodeURIComponent(propertyName),
    token
  );
  const options = Array.isArray(data && data.options) ? data.options : [];
  const out = {};

  options.forEach(function (option) {
    const value = String(option && option.value || "").trim();
    if (!value) return;
    out[value] = String(option.label || value).trim();
  });

  return out;
}

function getCreatorListStatusForHubSpotDealStage_(stageValueOrLabel) {
  const normalized = normalizeHeaderName_(stageValueOrLabel);
  return CREATOR_LIST_HUBSPOT_STAGE_TO_STATUS_[normalized] || "";
}

function syncPitchingStatusesToHubSpot_(items) {
  const prepared = prepareHubSpotPitchingStatusUpdates_(items);
  if (prepared.updates.length === 0) {
    Logger.log(
      `ℹ️ HubSpot Pitching Status sync skipped. Missing record ID: ${prepared.missingRecordIdCount}, ` +
      `missing status: ${prepared.missingStatusCount}.`
    );
    return {
      attempted: 0,
      updated: 0,
      failed: 0,
      skipped: prepared.skippedCount
    };
  }

  const token = getHubSpotApiToken_();
  if (!token) {
    Logger.log("ℹ️ HubSpot token not set in this project. Skipping Pitching Status sync.");
    return {
      attempted: prepared.updates.length,
      updated: 0,
      failed: 0,
      skipped: prepared.skippedCount
    };
  }

  try {
    const propertyName = resolveHubSpotDealPropertyName_(token, HUBSPOT_PITCHING_STATUS_PROPERTY_LABEL_);
    if (!propertyName) {
      throw new Error(`Deal property not found: ${HUBSPOT_PITCHING_STATUS_PROPERTY_LABEL_}`);
    }

    let updated = 0;
    let failed = 0;
    const chunks = chunkArray_(prepared.updates, HUBSPOT_BATCH_UPDATE_SIZE_);

    chunks.forEach(chunk => {
      try {
        updateHubSpotDealPropertyBatch_(token, propertyName, chunk);
        updated += chunk.length;
      } catch (batchError) {
        Logger.log("⚠️ HubSpot Pitching Status batch update failed. Retrying per deal. " + batchError);
        chunk.forEach(item => {
          try {
            updateHubSpotDealPropertySingle_(token, propertyName, item);
            updated++;
          } catch (singleError) {
            failed++;
            Logger.log(
              `⚠️ HubSpot Pitching Status sync failed for deal ${item.recordId} ` +
              `(row ${item.row1}): ${singleError}`
            );
          }
        });
      }
    });

    Logger.log(
      `✅ HubSpot Pitching Status synced for ${updated} deal(s). ` +
      `Skipped ${prepared.skippedCount}. Failed ${failed}.`
    );
    return {
      attempted: prepared.updates.length,
      updated: updated,
      failed: failed,
      skipped: prepared.skippedCount
    };
  } catch (e) {
    Logger.log("⚠️ HubSpot Pitching Status sync failed: " + (e && e.stack ? e.stack : e));
    return {
      attempted: prepared.updates.length,
      updated: 0,
      failed: prepared.updates.length,
      skipped: prepared.skippedCount
    };
  }
}

function prepareHubSpotPitchingStatusUpdates_(items) {
  const deduped = {};
  let missingRecordIdCount = 0;
  let missingStatusCount = 0;

  (items || []).forEach(item => {
    const recordId = String(item && item.recordId || "").trim();
    const status = String(item && item.status || "").trim();
    if (!recordId) {
      missingRecordIdCount++;
      return;
    }
    if (!status) {
      missingStatusCount++;
      return;
    }

    deduped[recordId] = {
      row1: Number(item && item.row1) || 0,
      recordId: recordId,
      status: status
    };
  });

  return {
    updates: Object.keys(deduped).map(key => deduped[key]),
    missingRecordIdCount: missingRecordIdCount,
    missingStatusCount: missingStatusCount,
    skippedCount: missingRecordIdCount + missingStatusCount
  };
}

function getHubSpotApiToken_() {
  const configured = String(
    PropertiesService.getScriptProperties().getProperty(HUBSPOT_API_KEY_PROP_) || ""
  ).trim();
  return configured || String(HUBSPOT_API_KEY_ || "").trim();
}

function syncDropdownValuesFromHubSpot() {
  try {
    const result = syncHubSpotDropdownValues_(SpreadsheetApp.getActiveSpreadsheet());
    const message = buildHubSpotDropdownSyncMessage_(result);

    Logger.log(message);
    showSpreadsheetToast_(
      "Dropdown Values synced from HubSpot. " +
      result.updatedColumnCount +
      " columns updated."
    );
    return result;
  } catch (e) {
    const message = "HubSpot Dropdown Values sync failed: " + (e && e.message ? e.message : e);
    Logger.log(message);
    showSpreadsheetAlert_(message);
    throw e;
  }
}

function syncHubSpotDropdownValues_(ss) {
  const config = HUBSPOT_DROPDOWN_VALUES_SYNC_CONFIG_;
  const spreadsheet = ss || SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) throw new Error("No active spreadsheet found.");

  const sheet = spreadsheet.getSheetByName(config.sheetName);
  if (!sheet) throw new Error("Sheet not found: " + config.sheetName);

  const token = getHubSpotApiToken_();
  if (!token) {
    throw new Error(
      "HubSpot Private App token is not configured. Set script property " +
      HUBSPOT_API_KEY_PROP_ + "."
    );
  }

  const campaignClientRows = fetchHubSpotDropdownCampaignClientRows_(token);
  const valuesByColumnName = {};
  valuesByColumnName[config.columns.clientName] = getDistinctHubSpotDropdownCampaignClientNames_(campaignClientRows);
  valuesByColumnName[config.columns.client] = campaignClientRows.map(function (row) {
    return row.clientName;
  });
  valuesByColumnName[config.columns.dealOwner] = fetchHubSpotDropdownDealOwnerNames_(token);
  valuesByColumnName[config.columns.campaignName] = campaignClientRows.map(function (row) {
    return row.campaignName;
  });
  valuesByColumnName[config.columns.pipeline] = fetchHubSpotDropdownDealPipelineNames_(token);
  valuesByColumnName[config.columns.dealStage] = fetchHubSpotDropdownDealStageNames_(token);
  valuesByColumnName[config.columns.currency] = fetchHubSpotDropdownCurrencyCodes_(token);
  (config.optionColumns || []).forEach(function (columnConfig) {
    valuesByColumnName[columnConfig.columnName] = fetchHubSpotDropdownPropertyOptions_(token, columnConfig);
  });

  writeHubSpotDropdownValues_(sheet, valuesByColumnName);

  return buildHubSpotDropdownSyncResult_(valuesByColumnName);
}

function fetchHubSpotDropdownDealOwnerNames_(token) {
  const names = [];
  let after = "";

  do {
    const query = buildHubSpotQueryString_({
      limit: 500,
      archived: false,
      after: after
    });
    const data = hubspotRequestJson_(HUBSPOT_API_BASE_ + "/crm/v3/owners/?" + query, token);
    const owners = Array.isArray(data && data.results) ? data.results : [];

    owners.forEach(function (owner) {
      if (!owner || owner.archived === true) return;

      const userId = String(owner.userId || "").trim();
      if (!userId) return;

      const firstName = String(owner.firstName || "").trim();
      const lastName = String(owner.lastName || "").trim();
      const fullName = [firstName, lastName].filter(Boolean).join(" ").trim();
      if (fullName) names.push(fullName);
    });

    after = String(data && data.paging && data.paging.next && data.paging.next.after || "").trim();
  } while (after);

  return normalizeHubSpotDropdownValues_(names);
}

function fetchHubSpotDropdownCampaignClientRows_(token) {
  const config = HUBSPOT_DROPDOWN_VALUES_SYNC_CONFIG_;
  const campaignConfig = config.campaign;
  const clientConfig = config.client;
  const createdAfter = getHubSpotDropdownCampaignCreatedAfter_(campaignConfig);
  const campaignObjectTypeId = resolveHubSpotDropdownObjectTypeId_(token, campaignConfig);
  const clientObjectTypeId = resolveHubSpotDropdownObjectTypeId_(token, clientConfig);
  const campaignPropertyName = resolveHubSpotDropdownPropertyName_(
    campaignObjectTypeId,
    campaignConfig && campaignConfig.propertyName,
    token
  );
  const clientPropertyName = resolveHubSpotDropdownPropertyName_(
    clientObjectTypeId,
    clientConfig && clientConfig.propertyName,
    token
  );
  const campaignRecords = fetchHubSpotCrmObjectRecords_(
    campaignObjectTypeId,
    [campaignPropertyName],
    token
  ).filter(function (record) {
    return isHubSpotRecordCreatedAtOrAfter_(record, createdAfter);
  });

  const campaignIds = campaignRecords.map(function (record) {
    return String(record && record.id || "").trim();
  });
  const clientIdsByCampaignId = fetchHubSpotAssociatedObjectIdsByFromId_(
    campaignObjectTypeId,
    clientObjectTypeId,
    campaignIds,
    token
  );
  const clientNameById = fetchHubSpotObjectPropertyValuesByIds_(
    clientObjectTypeId,
    flattenHubSpotDropdownIdMapValues_(clientIdsByCampaignId),
    clientPropertyName,
    token
  );
  const rows = [];

  campaignRecords.forEach(function (record) {
    const campaignId = String(record && record.id || "").trim();
    const properties = record && record.properties ? record.properties : {};
    const campaignName = String(properties[campaignPropertyName] || "").trim();
    if (!campaignName) return;

    const clientNames = uniqueHubSpotDropdownValues_(
      (clientIdsByCampaignId[campaignId] || []).map(function (clientId) {
        return clientNameById[clientId] || "";
      })
    );

    rows.push({
      campaignName: campaignName,
      clientName: clientNames.join("; "),
      clientNames: clientNames
    });
  });

  rows.sort(function (a, b) {
    const campaignCompare = normalizeDropdownLookupValue_(a.campaignName)
      .localeCompare(normalizeDropdownLookupValue_(b.campaignName));
    if (campaignCompare !== 0) return campaignCompare;
    return normalizeDropdownLookupValue_(a.clientName)
      .localeCompare(normalizeDropdownLookupValue_(b.clientName));
  });

  return rows;
}

function getDistinctHubSpotDropdownCampaignClientNames_(campaignClientRows) {
  const names = [];
  (campaignClientRows || []).forEach(function (row) {
    (row && row.clientNames || []).forEach(function (clientName) {
      names.push(clientName);
    });
  });
  return normalizeHubSpotDropdownValues_(names);
}

function fetchHubSpotDropdownCurrencyCodes_(token) {
  const codes = [];
  const errors = [];

  try {
    let after = "";
    do {
      const query = buildHubSpotQueryString_({
        limit: 100,
        after: after
      });
      const data = hubspotRequestJson_(
        HUBSPOT_API_BASE_ + "/settings/v3/currencies/exchange-rates/current?" + query,
        token
      );
      const rates = Array.isArray(data && data.results)
        ? data.results
        : (Array.isArray(data) ? data : []);

      rates.forEach(function (rate) {
        if (rate && rate.visibleInUI === false) return;
        appendHubSpotCurrencyCode_(codes, rate && rate.fromCurrencyCode);
        appendHubSpotCurrencyCode_(codes, rate && rate.toCurrencyCode);
      });

      after = String(data && data.paging && data.paging.next && data.paging.next.after || "").trim();
    } while (after);
  } catch (e) {
    errors.push(e);
  }

  try {
    const companyCurrency = hubspotRequestJson_(
      HUBSPOT_API_BASE_ + "/settings/v3/currencies/company-currency",
      token
    );
    appendHubSpotCurrencyCode_(codes, companyCurrency && companyCurrency.currencyCode);
  } catch (e) {
    errors.push(e);
  }

  const out = normalizeHubSpotCurrencyCodes_(codes);
  if (out.length > 0) return out;

  const fallback = fetchHubSpotDropdownPropertyOptions_(
    token,
    {
      objectTypeId: HUBSPOT_DEALS_OBJECT_API_NAME_,
      propertyName: "deal_currency_code"
    }
  );
  if (fallback.length > 0) return fallback;

  throw new Error(
    "HubSpot currency sync returned no currencies. " +
    "Make sure the Private App token has the multi-currency-read scope. " +
    errors.map(function (e) { return String(e && e.message ? e.message : e); }).join(" ")
  );
}

function appendHubSpotCurrencyCode_(out, value) {
  const code = String(value || "").trim().toUpperCase();
  if (code) out.push(code);
}

function normalizeHubSpotCurrencyCodes_(codes) {
  return uniqueHubSpotDropdownValues_(codes).sort();
}

function fetchHubSpotDropdownPropertyOptions_(token, propertyConfig) {
  const objectConfig = getHubSpotDropdownPropertyObjectConfig_(propertyConfig);
  const objectType = resolveHubSpotDropdownObjectTypeId_(token, objectConfig);
  const propertyName = resolveHubSpotDropdownPropertyName_(
    objectType,
    propertyConfig && propertyConfig.propertyName,
    token
  );
  const property = fetchHubSpotPropertyDefinition_(objectType, propertyName, token);
  const options = Array.isArray(property && property.options) ? property.options : [];

  return uniqueHubSpotDropdownValues_(
    options.map(function (option) {
      return String(option && (option.label || option.value) || "").trim();
    })
  );
}

function getHubSpotDropdownPropertyObjectConfig_(propertyConfig) {
  const configKey = String(propertyConfig && propertyConfig.objectConfigKey || "").trim();
  const sharedConfig = configKey
    ? HUBSPOT_DROPDOWN_VALUES_SYNC_CONFIG_[configKey]
    : null;

  return Object.assign({}, sharedConfig || {}, propertyConfig || {});
}

function fetchHubSpotDropdownDealPipelineNames_(token) {
  return uniqueHubSpotDropdownValues_(
    fetchHubSpotDealPipelines_(token).map(function (pipeline) {
      return String(pipeline && pipeline.label || "").trim();
    })
  );
}

function fetchHubSpotDropdownDealStageNames_(token) {
  const stageNames = [];
  fetchHubSpotDealPipelines_(token).forEach(function (pipeline) {
    const stages = Array.isArray(pipeline && pipeline.stages) ? pipeline.stages : [];
    stages.forEach(function (stage) {
      if (stage && stage.archived === true) return;
      stageNames.push(String(stage && stage.label || "").trim());
    });
  });
  return uniqueHubSpotDropdownValues_(stageNames);
}

function fetchHubSpotDealPipelines_(token) {
  const data = hubspotRequestJson_(
    HUBSPOT_API_BASE_ +
      "/crm/v3/pipelines/" +
      encodeURIComponent(HUBSPOT_DEALS_OBJECT_API_NAME_),
    token
  );
  const pipelines = Array.isArray(data && data.results) ? data.results : [];
  return pipelines.filter(function (pipeline) {
    return !(pipeline && pipeline.archived === true);
  });
}

function fetchHubSpotDropdownCustomObjectPropertyValues_(token, objectConfig, options) {
  const objectTypeId = resolveHubSpotDropdownObjectTypeId_(token, objectConfig);
  const propertyName = resolveHubSpotDropdownPropertyName_(
    objectTypeId,
    objectConfig && objectConfig.propertyName,
    token
  );
  const records = fetchHubSpotCrmObjectRecords_(objectTypeId, [propertyName], token);
  const values = [];

  records.forEach(function (record) {
    if (options && options.createdAfter && !isHubSpotRecordCreatedAtOrAfter_(record, options.createdAfter)) {
      return;
    }

    const properties = record && record.properties ? record.properties : {};
    const value = String(properties[propertyName] || "").trim();
    if (value) values.push(value);
  });

  return normalizeHubSpotDropdownValues_(values);
}

function resolveHubSpotDropdownObjectTypeId_(token, objectConfig) {
  const configuredObjectTypeId = String(objectConfig && objectConfig.objectTypeId || "").trim();
  if (configuredObjectTypeId) return configuredObjectTypeId;

  const aliases = uniqueNonEmptyStrings_(objectConfig && objectConfig.aliases || []);
  const info = loadHubSpotCustomObjectInfo_(
    {
      key: aliases[0] || "",
      aliases: aliases
    },
    token
  );
  if (!info || !info.objectTypeId) {
    throw new Error("Could not resolve HubSpot object type ID for " + aliases.join(", ") + ".");
  }
  return info.objectTypeId;
}

function resolveHubSpotDropdownPropertyName_(objectTypeId, propertyLabelOrName, token) {
  const targets = uniqueNonEmptyStrings_(
    Array.isArray(propertyLabelOrName) ? propertyLabelOrName : [propertyLabelOrName]
  ).map(function (value) {
    return normalizeHeaderName_(value);
  }).filter(Boolean);
  if (targets.length === 0) throw new Error("HubSpot dropdown property name is not configured.");

  const properties = fetchHubSpotPropertiesForObject_(objectTypeId, token);
  for (let i = 0; i < properties.length; i++) {
    const property = properties[i] || {};
    if (targets.indexOf(normalizeHeaderName_(property.name)) !== -1) return String(property.name || "").trim();
    if (targets.indexOf(normalizeHeaderName_(property.label)) !== -1) return String(property.name || "").trim();
  }

  throw new Error(
    "HubSpot property not found on object " +
    objectTypeId +
    ": " +
    getHubSpotDropdownPropertyConfigLabel_(propertyLabelOrName)
  );
}

function fetchHubSpotPropertyDefinition_(objectType, propertyName, token) {
  return hubspotRequestJson_(
    HUBSPOT_API_BASE_ +
      "/crm/v3/properties/" +
      encodeURIComponent(objectType) +
      "/" +
      encodeURIComponent(propertyName),
    token
  );
}

function fetchHubSpotCrmObjectRecords_(objectTypeId, propertyNames, token) {
  const records = [];
  const normalizedObjectTypeId = String(objectTypeId || "").trim();
  if (!normalizedObjectTypeId) return records;

  const properties = uniqueNonEmptyStrings_(propertyNames || []);
  let after = "";

  do {
    const query = buildHubSpotQueryString_({
      limit: 100,
      archived: false,
      properties: properties.join(","),
      after: after
    });
    const data = hubspotRequestJson_(
      HUBSPOT_API_BASE_ + "/crm/v3/objects/" + encodeURIComponent(normalizedObjectTypeId) + "?" + query,
      token
    );
    const pageRecords = Array.isArray(data && data.results) ? data.results : [];
    pageRecords.forEach(function (record) {
      if (!record || record.archived === true) return;
      records.push(record);
    });

    after = String(data && data.paging && data.paging.next && data.paging.next.after || "").trim();
  } while (after);

  return records;
}

function fetchHubSpotAssociatedObjectIdsByFromId_(fromObjectType, toObjectType, fromIds, token) {
  const ids = uniqueNonEmptyStrings_(fromIds || []);
  const out = {};
  ids.forEach(function (id) {
    out[id] = [];
  });
  if (ids.length === 0) return out;

  chunkArray_(ids, HUBSPOT_BATCH_UPDATE_SIZE_).forEach(function (chunk) {
    const data = hubspotRequestJson_(
      HUBSPOT_API_BASE_ +
        "/crm/v4/associations/" +
        encodeURIComponent(fromObjectType) +
        "/" +
        encodeURIComponent(toObjectType) +
        "/batch/read",
      token,
      {
        method: "post",
        payload: JSON.stringify({
          inputs: chunk.map(function (id) { return { id: id }; })
        })
      }
    );
    const results = Array.isArray(data && data.results) ? data.results : [];
    results.forEach(function (result) {
      const fromId = String(
        (result && result.from && result.from.id) ||
        (result && result.fromObjectId) ||
        ""
      ).trim();
      if (!fromId) return;

      const associatedIds = Array.isArray(result && result.to)
        ? result.to.map(function (item) {
          return String(
            (item && item.toObjectId) ||
            (item && item.id) ||
            (item && item.to && item.to.id) ||
            ""
          ).trim();
        })
        : [];
      out[fromId] = uniqueNonEmptyStrings_((out[fromId] || []).concat(associatedIds));
    });
  });

  return out;
}

function fetchHubSpotObjectPropertyValuesByIds_(objectTypeId, ids, propertyName, token) {
  const objectIds = uniqueNonEmptyStrings_(ids || []);
  const property = String(propertyName || "").trim();
  const out = {};
  if (!objectTypeId || objectIds.length === 0 || !property) return out;

  chunkArray_(objectIds, HUBSPOT_BATCH_UPDATE_SIZE_).forEach(function (chunk) {
    const data = hubspotRequestJson_(
      HUBSPOT_API_BASE_ +
        "/crm/v3/objects/" +
        encodeURIComponent(objectTypeId) +
        "/batch/read",
      token,
      {
        method: "post",
        payload: JSON.stringify({
          properties: [property],
          inputs: chunk.map(function (id) { return { id: id }; })
        })
      }
    );
    const results = Array.isArray(data && data.results) ? data.results : [];
    results.forEach(function (record) {
      if (!record || record.archived === true) return;

      const id = String(record.id || "").trim();
      const properties = record.properties || {};
      const value = String(properties[property] || "").trim();
      if (id && value) out[id] = value;
    });
  });

  return out;
}

function flattenHubSpotDropdownIdMapValues_(idMap) {
  const out = [];
  Object.keys(idMap || {}).forEach(function (key) {
    (idMap[key] || []).forEach(function (id) {
      out.push(id);
    });
  });
  return uniqueNonEmptyStrings_(out);
}

function writeHubSpotDropdownValues_(sheet, valuesByColumnName) {
  const columnNames = Object.keys(valuesByColumnName || {});
  if (columnNames.length === 0) return;

  const header = sheet
    .getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1))
    .getDisplayValues()[0]
    .map(function (value) { return String(value || "").trim(); });

  const columnIndexesByName = {};
  const missingColumns = [];
  columnNames.forEach(function (columnName) {
    const idx = findHeaderIndex_(header, columnName);
    if (idx === -1) {
      missingColumns.push(columnName);
      return;
    }
    columnIndexesByName[columnName] = idx;
  });

  if (missingColumns.length) {
    throw new Error(
      "Dropdown Values sheet is missing required column(s): " + missingColumns.join(", ")
    );
  }

  const maxValueCount = columnNames.reduce(function (max, columnName) {
    const values = valuesByColumnName[columnName] || [];
    return Math.max(max, values.length);
  }, 0);
  ensureHubSpotDropdownRowCapacity_(sheet, maxValueCount + 1);

  const clearRowCount = Math.max(sheet.getMaxRows() - 1, 0);
  columnNames.forEach(function (columnName) {
    const col1 = columnIndexesByName[columnName] + 1;
    const values = valuesByColumnName[columnName] || [];

    if (clearRowCount > 0) {
      sheet.getRange(2, col1, clearRowCount, 1).clearContent();
    }

    if (values.length > 0) {
      sheet.getRange(2, col1, values.length, 1).setValues(
        values.map(function (value) { return [value]; })
      );
    }
  });
}

function ensureHubSpotDropdownRowCapacity_(sheet, requiredRows) {
  const rowCount = Math.max(2, Number(requiredRows) || 2);
  const maxRows = sheet.getMaxRows();
  if (maxRows >= rowCount) return;

  sheet.insertRowsAfter(maxRows, rowCount - maxRows);
}

function getHubSpotDropdownCampaignCreatedAfter_(campaignConfig) {
  const months = Math.max(1, Number(campaignConfig && campaignConfig.createdWithinMonths) || 3);
  const cutoff = new Date();
  cutoff.setMonth(cutoff.getMonth() - months);
  return cutoff;
}

function isHubSpotRecordCreatedAtOrAfter_(record, cutoff) {
  const cutoffDate = Object.prototype.toString.call(cutoff) === "[object Date]" ? cutoff : null;
  if (!cutoffDate || isNaN(cutoffDate.getTime())) return true;

  const properties = record && record.properties ? record.properties : {};
  const createdValue =
    record && (record.createdAt || record.createdate) ||
    properties.createdate ||
    properties.hs_createdate ||
    "";
  const createdAt = parseHubSpotTimestamp_(createdValue);
  return createdAt && createdAt.getTime() >= cutoffDate.getTime();
}

function parseHubSpotTimestamp_(value) {
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return isNaN(value.getTime()) ? null : value;
  }

  const text = String(value || "").trim();
  if (!text) return null;

  const timestamp = /^\d+$/.test(text) ? Number(text) : NaN;
  const date = isNaN(timestamp) ? new Date(text) : new Date(timestamp);
  return isNaN(date.getTime()) ? null : date;
}

function normalizeHubSpotDropdownValues_(values) {
  const out = uniqueHubSpotDropdownValues_(values);
  out.sort(function (a, b) {
    return normalizeDropdownLookupValue_(a).localeCompare(normalizeDropdownLookupValue_(b));
  });
  return out;
}

function uniqueHubSpotDropdownValues_(values) {
  const out = [];
  const seen = {};

  (values || []).forEach(function (value) {
    const text = String(value || "").trim().replace(/\s+/g, " ");
    const key = normalizeDropdownLookupValue_(text);
    if (!key || seen[key]) return;

    seen[key] = true;
    out.push(text);
  });

  return out;
}

function buildHubSpotDropdownSyncResult_(valuesByColumnName) {
  const countsByColumnName = {};
  Object.keys(valuesByColumnName || {}).forEach(function (columnName) {
    countsByColumnName[columnName] = (valuesByColumnName[columnName] || []).length;
  });

  return {
    updatedColumnCount: Object.keys(countsByColumnName).length,
    countsByColumnName: countsByColumnName
  };
}

function buildHubSpotDropdownSyncMessage_(result) {
  const counts = result && result.countsByColumnName ? result.countsByColumnName : {};
  const parts = Object.keys(counts).map(function (columnName) {
    return columnName + ": " + counts[columnName];
  });

  return "Dropdown Values synced from HubSpot. " + parts.join(", ") + ".";
}

function getHubSpotDropdownPropertyConfigLabel_(propertyLabelOrName) {
  if (Array.isArray(propertyLabelOrName)) {
    return propertyLabelOrName.join(", ");
  }
  return String(propertyLabelOrName || "");
}

function buildHubSpotQueryString_(params) {
  return Object.keys(params || {})
    .filter(function (key) {
      const value = params[key];
      return value !== null && value !== undefined && String(value) !== "";
    })
    .map(function (key) {
      return encodeURIComponent(key) + "=" + encodeURIComponent(String(params[key]));
    })
    .join("&");
}

function resolveHubSpotDealPropertyName_(token, propertyLabelOrName) {
  const target = normalizeHeaderName_(propertyLabelOrName);
  if (!target) return "";

  const properties = fetchHubSpotPropertiesForObject_(HUBSPOT_DEALS_OBJECT_TYPE_ID_, token);

  for (let i = 0; i < properties.length; i++) {
    const property = properties[i] || {};
    if (normalizeHeaderName_(property.name) === target) return String(property.name || "").trim();
    if (normalizeHeaderName_(property.label) === target) return String(property.name || "").trim();
  }

  return "";
}

function fetchHubSpotPropertiesForObject_(objectType, token) {
  const normalizedObjectType = String(objectType || "").trim();
  if (!normalizedObjectType) return [];

  const cacheKey = getHubSpotPropertiesCacheKey_(token, normalizedObjectType);
  if (cacheKey) {
    try {
      const cached = CacheService.getScriptCache().get(cacheKey);
      if (cached) {
        const parsed = JSON.parse(cached);
        if (Array.isArray(parsed)) return parsed;
      }
    } catch (e) {
      // Property cache is best-effort; fall back to HubSpot if it is unavailable.
    }
  }

  const data = hubspotRequestJson_(
    HUBSPOT_API_BASE_ + "/crm/v3/properties/" + encodeURIComponent(normalizedObjectType),
    token
  );
  const properties = Array.isArray(data && data.results)
    ? data.results.map(compactHubSpotPropertyInfo_)
    : [];

  if (cacheKey) {
    try {
      const serialized = JSON.stringify(properties);
      if (serialized.length <= HUBSPOT_PROPERTY_CACHE_MAX_CHARS_) {
        CacheService.getScriptCache().put(
          cacheKey,
          serialized,
          HUBSPOT_PROPERTY_CACHE_TTL_SECONDS_
        );
      }
    } catch (e) {
      // Cache writes are best-effort; imports/syncs should continue.
    }
  }

  return properties;
}

function compactHubSpotPropertyInfo_(property) {
  return {
    name: String(property && property.name || "").trim(),
    label: String(property && property.label || property && property.name || "").trim(),
    type: String(property && property.type || "").trim(),
    fieldType: String(property && property.fieldType || "").trim()
  };
}

function getHubSpotPropertiesCacheKey_(token, objectType) {
  const type = String(objectType || "").trim();
  if (!type) return "";
  return "HS_PROPERTIES::" + getHubSpotTokenFingerprint_(token) + "::" + type;
}

function getHubSpotTokenFingerprint_(token) {
  try {
    const digest = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      String(token || ""),
      Utilities.Charset.UTF_8
    );
    return digest
      .map(function (byte) {
        const value = (byte + 256) % 256;
        return ("0" + value.toString(16)).slice(-2);
      })
      .join("")
      .slice(0, 12);
  } catch (e) {
    return "default";
  }
}

function updateHubSpotDealPropertyBatch_(token, propertyName, items) {
  return hubspotRequestJson_(
    HUBSPOT_API_BASE_ + "/crm/v3/objects/deals/batch/update",
    token,
    {
      method: "post",
      payload: JSON.stringify({
        inputs: (items || []).map(item => ({
          id: item.recordId,
          properties: { [propertyName]: item.status }
        }))
      })
    }
  );
}

function updateHubSpotDealPropertySingle_(token, propertyName, item) {
  return hubspotRequestJson_(
    HUBSPOT_API_BASE_ + "/crm/v3/objects/deals/" + encodeURIComponent(item.recordId),
    token,
    {
      method: "patch",
      payload: JSON.stringify({
        properties: { [propertyName]: item.status }
      })
    }
  );
}

function syncCampaignRowsToHubSpot_(sheet) {
  if (!sheet) {
    Logger.log("ℹ️ HubSpot Campaign sync skipped. Campaigns sheet not available.");
    return {
      stageResults: [],
      updated: 0,
      failed: 0
    };
  }

  const token = getHubSpotApiToken_();
  if (!token) {
    Logger.log("ℹ️ HubSpot token not set in this project. Skipping Campaigns sync.");
    return {
      stageResults: [],
      updated: 0,
      failed: 0
    };
  }

  const data = sheet.getDataRange().getValues();
  if (!data || data.length < 2) {
    Logger.log("ℹ️ HubSpot Campaign sync skipped. Campaigns sheet has no data rows.");
    return {
      stageResults: [],
      updated: 0,
      failed: 0
    };
  }

  const header = (data[0] || []).map(function (value) {
    return String(value || "").trim();
  });
  const spreadsheetTimeZone =
    (sheet.getParent() && sheet.getParent().getSpreadsheetTimeZone()) ||
    Session.getScriptTimeZone() ||
    "UTC";

  const dealIdCol = findHeaderIndex_(header, "HubSpot Record ID");
  const activationIdCol = findHeaderIndex_(header, "HubSpot Activation ID");
  if (dealIdCol === -1 && activationIdCol === -1) {
    Logger.log(
      "ℹ️ HubSpot Campaign sync skipped. Missing 'HubSpot Record ID' and 'HubSpot Activation ID' columns."
    );
    return {
      stageResults: [],
      updated: 0,
      failed: 0
    };
  }

  const dealPropertyLabels = uniqueNonEmptyStrings_(
    flatten2d_(
      HUBSPOT_INT_CAMPAIGN_DEAL_SYNC_STAGES_.map(function (stage) {
        return (stage.fields || []).map(function (field) { return field.propertyLabel; });
      })
    )
  );
  const dealPropertyInfoByLabel = dealPropertyLabels.length > 0
    ? resolveHubSpotPropertyInfosByLabel_(HUBSPOT_DEALS_OBJECT_API_NAME_, dealPropertyLabels, token)
    : {};
  const unresolvedDealProperties = dealPropertyLabels.filter(function (label) {
    return !dealPropertyInfoByLabel[label];
  });
  if (unresolvedDealProperties.length > 0) {
    Logger.log(
      "⚠️ HubSpot Campaign sync could not resolve deal properties: " +
      unresolvedDealProperties.join(", ")
    );
  }
  if (dealIdCol === -1 || activationIdCol === -1) {
    Logger.log(
      "ℹ️ HubSpot Campaign deal sync requires both 'HubSpot Record ID' and " +
      "'HubSpot Activation ID' columns."
    );
  }

  const groupedDealStagePlan = buildHubSpotGroupedDealStageUpdatesFromCampaignRows_(
    data,
    header,
    HUBSPOT_INT_CAMPAIGN_DEAL_SYNC_STAGES_,
    dealPropertyInfoByLabel,
    spreadsheetTimeZone
  );

  let activationStagePlan = {
    stages: [],
    rowsConsidered: 0,
    missingIdCount: 0,
    rowsWithNoStageValues: 0
  };

  if (activationIdCol !== -1) {
    try {
      const activationsInfo = loadHubSpotCustomObjectInfo_(
        {
          key: HUBSPOT_ACTIVATION_OBJECT_KEY_,
          aliases: ["Activation", "Activations", HUBSPOT_ACTIVATION_OBJECT_KEY_]
        },
        token
      );

      const activationPropertyLabels = uniqueNonEmptyStrings_(
        flatten2d_(
          HUBSPOT_INT_CAMPAIGN_ACTIVATION_SYNC_STAGES_.map(function (stage) {
            return (stage.fields || []).map(function (field) { return field.propertyLabel; });
          })
        )
      );
      const activationPropertyInfoByLabel = activationPropertyLabels.length > 0
        ? resolveHubSpotPropertyInfosByLabel_(activationsInfo.objectTypeId, activationPropertyLabels, token)
        : {};
      const unresolvedActivationProperties = activationPropertyLabels.filter(function (label) {
        return !activationPropertyInfoByLabel[label];
      });
      if (unresolvedActivationProperties.length > 0) {
        Logger.log(
          "⚠️ HubSpot Campaign sync could not resolve activation properties: " +
          unresolvedActivationProperties.join(", ")
        );
      }

      activationStagePlan = buildHubSpotStageUpdatesFromCampaignRows_(
        data,
        header,
        HUBSPOT_INT_CAMPAIGN_ACTIVATION_SYNC_STAGES_,
        activationPropertyInfoByLabel,
        "HubSpot Activation ID",
        activationsInfo.objectTypeId,
        spreadsheetTimeZone
      );
    } catch (e) {
      Logger.log("⚠️ HubSpot Campaign sync could not resolve Activations object type: " + e);
    }
  } else {
    Logger.log("ℹ️ HubSpot Campaign activation sync skipped. Missing 'HubSpot Activation ID' column.");
  }

  const stageResults = [];
  let totalUpdated = 0;
  let totalFailed = 0;

  activationStagePlan.stages.forEach(function (stage) {
    const stageResult = stage && Array.isArray(stage.updates) && stage.updates.length > 0
      ? updateHubSpotObjectPropertiesBatch_(stage.objectTypeId, stage.updates, token)
      : { updated: 0, failed: 0 };

    stageResults.push({
      stageLabel: stage.stageLabel,
      updated: stageResult.updated,
      failed: stageResult.failed
    });
    totalUpdated += stageResult.updated;
    totalFailed += stageResult.failed;
  });

  groupedDealStagePlan.stages.forEach(function (stage, index) {
    const stageResult = stage && Array.isArray(stage.updates) && stage.updates.length > 0
      ? updateHubSpotObjectPropertiesBatch_(stage.objectTypeId, stage.updates, token)
      : { updated: 0, failed: 0 };

    stageResults.push({
      stageLabel: stage.stageLabel,
      updated: stageResult.updated,
      failed: stageResult.failed
    });
    totalUpdated += stageResult.updated;
    totalFailed += stageResult.failed;

    const hasLaterStageWithUpdates = groupedDealStagePlan.stages.slice(index + 1).some(function (nextStage) {
      return nextStage && Array.isArray(nextStage.updates) && nextStage.updates.length > 0;
    });
    if (stage && Array.isArray(stage.updates) && stage.updates.length > 0 && hasLaterStageWithUpdates) {
      Logger.log(
        `⏳ Waiting ${Math.round(HUBSPOT_CAMPAIGN_SYNC_STAGE_DELAY_MS_ / 1000)} seconds before ` +
        `the next Campaigns HubSpot sync stage.`
      );
      Utilities.sleep(HUBSPOT_CAMPAIGN_SYNC_STAGE_DELAY_MS_);
    }
  });

  Logger.log(
    `✅ HubSpot Campaign sync done. Total updated: ${totalUpdated}. Failed: ${totalFailed}. ` +
    `Deal rows considered: ${groupedDealStagePlan.rowsConsidered}. ` +
    `Deals evaluated: ${groupedDealStagePlan.dealsEvaluated}. ` +
    `Missing deal ID: ${groupedDealStagePlan.missingDealIdCount}. ` +
    `Missing associated activation ID: ${groupedDealStagePlan.missingAssociatedActivationIdCount}. ` +
    `Deals with no values: ${groupedDealStagePlan.dealsWithNoStageValues}. ` +
    `Activation rows considered: ${activationStagePlan.rowsConsidered}. ` +
    `Missing activation ID: ${activationStagePlan.missingIdCount}. ` +
    `Activation rows with no values: ${activationStagePlan.rowsWithNoStageValues}. ` +
    `Stages: ${stageResults.map(function (stage) {
      return stage.stageLabel + "=" + stage.updated + "/" + stage.failed;
    }).join(", ")}.`
  );

  return {
    stageResults: stageResults,
    updated: totalUpdated,
    failed: totalFailed
  };
}

function syncPerformanceRowsToHubSpot_(sheet) {
  if (!sheet) {
    Logger.log("ℹ️ HubSpot Performance sync skipped. Performance sheet not available.");
    return {
      stageResults: [],
      updated: 0,
      failed: 0
    };
  }

  const token = getHubSpotApiToken_();
  if (!token) {
    Logger.log("ℹ️ HubSpot token not set in this project. Skipping Performance sync.");
    return {
      stageResults: [],
      updated: 0,
      failed: 0
    };
  }

  const data = sheet.getDataRange().getValues();
  if (!data || data.length < 2) {
    Logger.log("ℹ️ HubSpot Performance sync skipped. Performance sheet has no data rows.");
    return {
      stageResults: [],
      updated: 0,
      failed: 0
    };
  }

  const header = (data[0] || []).map(function (value) {
    return String(value || "").trim();
  });
  const activationIdCol = findHeaderIndex_(header, "HubSpot Activation ID");
  if (activationIdCol === -1) {
    Logger.log("ℹ️ HubSpot Performance sync skipped. Missing 'HubSpot Activation ID' column.");
    return {
      stageResults: [],
      updated: 0,
      failed: 0
    };
  }

  const spreadsheetTimeZone =
    (sheet.getParent() && sheet.getParent().getSpreadsheetTimeZone()) ||
    Session.getScriptTimeZone() ||
    "UTC";

  try {
    const activationsInfo = loadHubSpotCustomObjectInfo_(
      {
        key: HUBSPOT_ACTIVATION_OBJECT_KEY_,
        aliases: ["Activation", "Activations", HUBSPOT_ACTIVATION_OBJECT_KEY_]
      },
      token
    );

    const activationPropertyLabels = uniqueNonEmptyStrings_(
      flatten2d_(
        HUBSPOT_INT_PERFORMANCE_ACTIVATION_SYNC_STAGES_.map(function (stage) {
          return (stage.fields || []).map(function (field) { return field.propertyLabel; });
        })
      )
    );
    const activationPropertyInfoByLabel = activationPropertyLabels.length > 0
      ? resolveHubSpotPropertyInfosByLabel_(activationsInfo.objectTypeId, activationPropertyLabels, token)
      : {};
    const unresolvedActivationProperties = activationPropertyLabels.filter(function (label) {
      return !activationPropertyInfoByLabel[label];
    });
    if (unresolvedActivationProperties.length > 0) {
      Logger.log(
        "⚠️ HubSpot Performance sync could not resolve activation properties: " +
        unresolvedActivationProperties.join(", ")
      );
    }

    const stagePlan = buildHubSpotStageUpdatesFromCampaignRows_(
      data,
      header,
      HUBSPOT_INT_PERFORMANCE_ACTIVATION_SYNC_STAGES_,
      activationPropertyInfoByLabel,
      "HubSpot Activation ID",
      activationsInfo.objectTypeId,
      spreadsheetTimeZone
    );

    const stageResults = [];
    let totalUpdated = 0;
    let totalFailed = 0;

    stagePlan.stages.forEach(function (stage) {
      const stageResult = stage && Array.isArray(stage.updates) && stage.updates.length > 0
        ? updateHubSpotObjectPropertiesBatch_(stage.objectTypeId, stage.updates, token)
        : { updated: 0, failed: 0 };

      stageResults.push({
        stageLabel: stage.stageLabel,
        updated: stageResult.updated,
        failed: stageResult.failed
      });
      totalUpdated += stageResult.updated;
      totalFailed += stageResult.failed;
    });

    Logger.log(
      `✅ HubSpot Performance sync done. Total updated: ${totalUpdated}. Failed: ${totalFailed}. ` +
      `Rows considered: ${stagePlan.rowsConsidered}. ` +
      `Missing activation ID: ${stagePlan.missingIdCount}. ` +
      `Rows with no values: ${stagePlan.rowsWithNoStageValues}. ` +
      `Stages: ${stageResults.map(function (stage) {
        return stage.stageLabel + "=" + stage.updated + "/" + stage.failed;
      }).join(", ")}.`
    );

    return {
      stageResults: stageResults,
      updated: totalUpdated,
      failed: totalFailed
    };
  } catch (e) {
    Logger.log("⚠️ HubSpot Performance sync failed: " + (e && e.stack ? e.stack : e));
    return {
      stageResults: [],
      updated: 0,
      failed: 0
    };
  }
}

function buildHubSpotGroupedDealStageUpdatesFromCampaignRows_(
  data,
  header,
  stageDefs,
  propertyInfoByLabel,
  spreadsheetTimeZone
) {
  const normalizedStages = (stageDefs || []).map(function (stage) {
    return {
      stageLabel: String(stage && stage.stageLabel || "").trim(),
      mode: String(stage && stage.mode || "").trim(),
      fields: Array.isArray(stage && stage.fields) ? stage.fields : []
    };
  });
  const dealIdCol = findHeaderIndex_(header, "HubSpot Record ID");
  const activationIdCol = findHeaderIndex_(header, "HubSpot Activation ID");
  const stageMaps = normalizedStages.map(function () { return {}; });
  const rowsByDealId = {};
  let rowsConsidered = 0;
  let missingDealIdCount = 0;
  let missingAssociatedActivationIdCount = 0;
  let dealsWithNoStageValues = 0;

  if (normalizedStages.length === 0 || dealIdCol === -1 || activationIdCol === -1) {
    return {
      stages: normalizedStages.map(function (stage) {
        return {
          stageLabel: stage.stageLabel,
          objectTypeId: HUBSPOT_DEALS_OBJECT_API_NAME_,
          updates: []
        };
      }),
      rowsConsidered: 0,
      dealsEvaluated: 0,
      missingDealIdCount: 0,
      missingAssociatedActivationIdCount: 0,
      dealsWithNoStageValues: 0
    };
  }

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (isBlankRow_(row) || isSectionLabelRow_(row)) continue;
    rowsConsidered++;

    const dealId = String(row[dealIdCol] || "").trim();
    if (!dealId) {
      missingDealIdCount++;
      continue;
    }

    const activationId = String(row[activationIdCol] || "").trim();
    if (!activationId) {
      missingAssociatedActivationIdCount++;
      continue;
    }

    if (!rowsByDealId[dealId]) {
      rowsByDealId[dealId] = {
        row1: r + 1,
        rows: []
      };
    }
    rowsByDealId[dealId].rows.push(row);
  }

  Object.keys(rowsByDealId).forEach(function (dealId) {
    const entry = rowsByDealId[dealId];
    if (!entry || !Array.isArray(entry.rows) || entry.rows.length === 0) return;

    let dealHadStageValues = false;

    normalizedStages.forEach(function (stage, index) {
      const properties = buildHubSpotGroupedDealStageProperties_(
        entry.rows,
        header,
        stage,
        propertyInfoByLabel,
        spreadsheetTimeZone
      );
      if (Object.keys(properties).length === 0) return;

      dealHadStageValues = true;
      stageMaps[index][dealId] = {
        id: dealId,
        properties: properties,
        row1: entry.row1
      };
    });

    if (!dealHadStageValues) dealsWithNoStageValues++;
  });

  return {
    stages: normalizedStages.map(function (stage, index) {
      return {
        stageLabel: stage.stageLabel,
        objectTypeId: HUBSPOT_DEALS_OBJECT_API_NAME_,
        updates: Object.keys(stageMaps[index]).map(function (dealId) {
          return stageMaps[index][dealId];
        })
      };
    }),
    rowsConsidered: rowsConsidered,
    dealsEvaluated: Object.keys(rowsByDealId).length,
    missingDealIdCount: missingDealIdCount,
    missingAssociatedActivationIdCount: missingAssociatedActivationIdCount,
    dealsWithNoStageValues: dealsWithNoStageValues
  };
}

function buildHubSpotGroupedDealStageProperties_(
  rows,
  header,
  stage,
  propertyInfoByLabel,
  spreadsheetTimeZone
) {
  const properties = {};
  if (!stage || !Array.isArray(stage.fields) || stage.fields.length === 0) return properties;

  (stage.fields || []).forEach(function (field) {
    if (!field || !field.propertyLabel || !field.sourceColumn) return;

    const propertyInfo = propertyInfoByLabel[field.propertyLabel];
    if (!propertyInfo) return;

    const colIdx = findHeaderIndex_(header, field.sourceColumn);
    if (colIdx === -1) return;

    let rawValue = "";
    if (stage.mode === "all_rows_have_value") {
      rawValue = rows.every(function (row) {
        return !isEmptyHubSpotSyncValue_(row[colIdx]);
      })
        ? field.completeValue
        : field.incompleteValue;
    } else {
      for (let i = 0; i < rows.length; i++) {
        if (!isEmptyHubSpotSyncValue_(rows[i][colIdx])) {
          rawValue = rows[i][colIdx];
          break;
        }
      }
    }

    if (isEmptyHubSpotSyncValue_(rawValue)) return;

    const serializedValue = serializeHubSpotPropertyValue_(rawValue, propertyInfo, spreadsheetTimeZone);
    if (serializedValue === "") return;
    properties[propertyInfo.name] = serializedValue;
  });

  return properties;
}

function loadHubSpotCustomObjectInfo_(spec, token) {
  const schemasResponse = hubspotRequestJson_(
    HUBSPOT_API_BASE_ + "/crm-object-schemas/v3/schemas",
    token
  );
  const schemas = Array.isArray(schemasResponse && schemasResponse.results)
    ? schemasResponse.results
    : [];
  const wanted = uniqueNonEmptyStrings_(
    (spec && spec.aliases ? spec.aliases : [spec && spec.key]).map(function (value) {
      return normalizeHeaderName_(value);
    })
  );

  for (let i = 0; i < schemas.length; i++) {
    const schema = schemas[i] || {};
    const candidates = uniqueNonEmptyStrings_([
      normalizeHeaderName_(schema.name),
      normalizeHeaderName_(schema.labels && schema.labels.singular),
      normalizeHeaderName_(schema.labels && schema.labels.plural)
    ]);
    const matched = wanted.some(function (alias) {
      return candidates.indexOf(alias) !== -1;
    });
    if (!matched) continue;

    return {
      objectTypeId: String(schema.objectTypeId || "").trim(),
      label: String(
        (schema.labels && (schema.labels.plural || schema.labels.singular)) ||
        (spec && spec.key) ||
        ""
      ).trim(),
      primaryDisplayProperty: String(schema.primaryDisplayProperty || "").trim()
    };
  }

  throw new Error(`Could not find HubSpot custom object schema for "${spec && spec.key}".`);
}

function resolveHubSpotPropertyInfosByLabel_(objectType, propertyLabels, token) {
  const properties = fetchHubSpotPropertiesForObject_(objectType, token);

  const lookup = {};
  properties.forEach(function (property) {
    if (!property) return;

    [
      normalizeHeaderName_(property.label),
      normalizeHeaderName_(property.name)
    ].forEach(function (variant) {
      if (!variant || lookup[variant]) return;
      lookup[variant] = {
        name: String(property.name || "").trim(),
        label: String(property.label || property.name || "").trim(),
        type: String(property.type || "").trim(),
        fieldType: String(property.fieldType || "").trim()
      };
    });
  });

  const out = {};
  (propertyLabels || []).forEach(function (label) {
    const key = normalizeHeaderName_(label);
    if (!key || !lookup[key]) return;
    out[String(label)] = lookup[key];
  });
  return out;
}

function updateHubSpotObjectPropertiesBatch_(objectType, items, token) {
  const updates = Array.isArray(items) ? items.filter(function (item) {
    return item && item.id && item.properties && Object.keys(item.properties).length > 0;
  }) : [];
  if (updates.length === 0) {
    return {
      updated: 0,
      failed: 0
    };
  }

  let updated = 0;
  let failed = 0;
  const chunks = chunkArray_(updates, HUBSPOT_BATCH_UPDATE_SIZE_);

  chunks.forEach(function (chunk) {
    try {
      hubspotRequestJson_(
        HUBSPOT_API_BASE_ + "/crm/v3/objects/" + encodeURIComponent(objectType) + "/batch/update",
        token,
        {
          method: "post",
          payload: JSON.stringify({
            inputs: chunk.map(function (item) {
              return {
                id: item.id,
                properties: item.properties
              };
            })
          })
        }
      );
      updated += chunk.length;
    } catch (batchError) {
      Logger.log("⚠️ HubSpot Campaign batch update failed. Retrying per object. " + batchError);
      chunk.forEach(function (item) {
        try {
          updateHubSpotObjectPropertiesSingle_(objectType, item, token);
          updated++;
        } catch (singleError) {
          failed++;
          Logger.log(
            `⚠️ HubSpot Campaign sync failed for object ${item.id} ` +
            `(row ${item.row1}): ${singleError}`
          );
        }
      });
    }
  });

  return {
    updated: updated,
    failed: failed
  };
}

function updateHubSpotObjectPropertiesSingle_(objectType, item, token) {
  return hubspotRequestJson_(
    HUBSPOT_API_BASE_ + "/crm/v3/objects/" + encodeURIComponent(objectType) + "/" + encodeURIComponent(item.id),
    token,
    {
      method: "patch",
      payload: JSON.stringify({
        properties: item.properties
      })
    }
  );
}

function buildHubSpotStageUpdatesFromCampaignRows_(
  data,
  header,
  stageDefs,
  propertyInfoByLabel,
  idColumnName,
  objectTypeId,
  spreadsheetTimeZone
) {
  const normalizedStages = (stageDefs || []).map(function (stage) {
    return {
      stageLabel: String(stage && stage.stageLabel || "").trim(),
      fields: Array.isArray(stage && stage.fields) ? stage.fields : []
    };
  });
  const idCol = findHeaderIndex_(header, idColumnName);
  const stageMaps = normalizedStages.map(function () { return {}; });
  let rowsConsidered = 0;
  let missingIdCount = 0;
  let rowsWithNoStageValues = 0;

  if (normalizedStages.length === 0 || idCol === -1) {
    return {
      stages: normalizedStages.map(function (stage) {
        return {
          stageLabel: stage.stageLabel,
          objectTypeId: objectTypeId,
          updates: []
        };
      }),
      rowsConsidered: 0,
      missingIdCount: 0,
      rowsWithNoStageValues: 0
    };
  }

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (isBlankRow_(row) || isSectionLabelRow_(row)) continue;
    rowsConsidered++;

    const objectId = String(row[idCol] || "").trim();
    if (!objectId) {
      missingIdCount++;
      continue;
    }

    let rowHadStageValues = false;
    normalizedStages.forEach(function (stage, index) {
      const properties = buildHubSpotPropertiesFromRowForFields_(
        row,
        header,
        stage.fields,
        propertyInfoByLabel,
        spreadsheetTimeZone
      );
      if (Object.keys(properties).length === 0) return;

      rowHadStageValues = true;
      if (!stageMaps[index][objectId]) {
        stageMaps[index][objectId] = {
          id: objectId,
          properties: {},
          row1: r + 1
        };
      }
      Object.keys(properties).forEach(function (propertyName) {
        stageMaps[index][objectId].properties[propertyName] = properties[propertyName];
      });
      stageMaps[index][objectId].row1 = r + 1;
    });

    if (!rowHadStageValues) rowsWithNoStageValues++;
  }

  return {
    stages: normalizedStages.map(function (stage, index) {
      return {
        stageLabel: stage.stageLabel,
        objectTypeId: objectTypeId,
        updates: Object.keys(stageMaps[index]).map(function (key) {
          return stageMaps[index][key];
        })
      };
    }),
    rowsConsidered: rowsConsidered,
    missingIdCount: missingIdCount,
    rowsWithNoStageValues: rowsWithNoStageValues
  };
}

function buildHubSpotPropertiesFromRowForFields_(
  row,
  header,
  fields,
  propertyInfoByLabel,
  spreadsheetTimeZone
) {
  const properties = {};

  (fields || []).forEach(function (field) {
    if (!field || !field.propertyLabel || !field.sourceColumn) return;

    const propertyInfo = propertyInfoByLabel[field.propertyLabel];
    if (!propertyInfo) return;

    const colIdx = findHeaderIndex_(header, field.sourceColumn);
    if (colIdx === -1) return;

    const rawValue = row[colIdx];
    if (isEmptyHubSpotSyncValue_(rawValue)) return;

    const serializedValue = serializeHubSpotPropertyValue_(rawValue, propertyInfo, spreadsheetTimeZone);
    if (serializedValue === "") return;
    properties[propertyInfo.name] = serializedValue;
  });

  return properties;
}

function isEmptyHubSpotSyncValue_(value) {
  if (value == null) return true;
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return isNaN(value.getTime());
  }
  return String(value).trim() === "";
}

function serializeHubSpotPropertyValue_(value, propertyInfo, spreadsheetTimeZone) {
  const type = String(propertyInfo && propertyInfo.type || "").trim().toLowerCase();
  const fieldType = String(propertyInfo && propertyInfo.fieldType || "").trim().toLowerCase();

  if (Object.prototype.toString.call(value) === "[object Date]") {
    if (isNaN(value.getTime())) return "";

    if (type === "date" || fieldType === "date") {
      const dateText = Utilities.formatDate(
        value,
        spreadsheetTimeZone || Session.getScriptTimeZone() || "UTC",
        "yyyy-MM-dd"
      );
      const parts = dateText.split("-").map(function (part) { return Number(part); });
      if (parts.length !== 3 || parts.some(function (part) { return isNaN(part); })) return "";
      return String(Date.UTC(parts[0], parts[1] - 1, parts[2]));
    }
    if (type === "datetime") {
      return String(value.getTime());
    }

    return Utilities.formatDate(
      value,
      spreadsheetTimeZone || Session.getScriptTimeZone() || "UTC",
      "yyyy-MM-dd"
    );
  }

  const text = String(value == null ? "" : value).trim();
  if (!text) return "";

  if (type === "date" || fieldType === "date") {
    const parsedDate = new Date(text);
    if (isNaN(parsedDate.getTime())) return text;
    return String(Date.UTC(parsedDate.getFullYear(), parsedDate.getMonth(), parsedDate.getDate()));
  }

  if (type === "datetime") {
    const parsedDateTime = new Date(text);
    return isNaN(parsedDateTime.getTime()) ? text : String(parsedDateTime.getTime());
  }

  if (type === "bool" || fieldType === "booleancheckbox") {
    const normalized = text.toLowerCase();
    if (normalized === "true" || normalized === "yes" || normalized === "1") return "true";
    if (normalized === "false" || normalized === "no" || normalized === "0") return "false";
  }

  return text;
}

function hubspotRequestJson_(url, token, options) {
  const requestOptions = Object.assign(
    {
      method: "get",
      muteHttpExceptions: true,
      headers: {
        Authorization: "Bearer " + String(token || ""),
        "Content-Type": "application/json"
      }
    },
    options || {}
  );
  const response = hubspotFetchWithRetry_(String(url || ""), requestOptions);

  const code = response.getResponseCode();
  const text = String(response.getContentText() || "");
  if (code >= 400) {
    throw new Error("HubSpot API error " + code + ": " + text.slice(0, 1000));
  }
  if (!text) return {};

  try {
    return JSON.parse(text);
  } catch (e) {
    throw new Error("HubSpot API parse error: " + e);
  }
}

function hubspotFetchWithRetry_(url, options) {
  const requestOptions = options || {};
  const maxRetries = Math.max(0, Number(HUBSPOT_REQUEST_RETRY_COUNT_) || 0);
  let lastError = null;

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    waitForHubSpotRequestSlot_();

    try {
      const response = UrlFetchApp.fetch(url, requestOptions);
      const code = response.getResponseCode();
      const text = String(response.getContentText() || "");
      if (!shouldRetryHubSpotResponse_(code, text) || attempt >= maxRetries) {
        return response;
      }

      Utilities.sleep(getHubSpotRetryDelayMs_(response, attempt));
    } catch (e) {
      lastError = e;
      if (!isRetryableHubSpotFetchException_(e) || attempt >= maxRetries) throw e;

      Utilities.sleep(getHubSpotRetryDelayMs_(null, attempt));
    }
  }

  throw lastError || new Error("HubSpot request failed.");
}

function shouldRetryHubSpotResponse_(code, text) {
  if (code === 429 || code >= 500) return true;
  return /TEN_SECONDLY_ROLLING|RATE_LIMIT|temporarily unavailable/i.test(String(text || ""));
}

function isRetryableHubSpotFetchException_(error) {
  const message = String(error && error.message ? error.message : error || "");
  return /Bandwidth quota exceeded|Service invoked too many times|Address unavailable|timed out|timeout|DNS|Socket|Connection/i.test(message);
}

function getHubSpotRetryDelayMs_(response, attempt) {
  const retryAfterMs = getHubSpotRetryAfterMs_(response);
  if (retryAfterMs > 0) return retryAfterMs;

  const base = Math.max(500, Number(HUBSPOT_REQUEST_RETRY_BASE_MS_) || 2000);
  const exponential = base * Math.pow(2, Math.max(0, Number(attempt) || 0));
  return Math.min(60000, exponential);
}

function getHubSpotRetryAfterMs_(response) {
  if (!response || typeof response.getAllHeaders !== "function") return 0;

  try {
    const headers = response.getAllHeaders() || {};
    const retryAfter = headers["Retry-After"] || headers["retry-after"];
    if (!retryAfter) return 0;

    const seconds = Number(retryAfter);
    if (isFinite(seconds) && seconds > 0) return Math.min(60000, seconds * 1000);
  } catch (e) {
    return 0;
  }

  return 0;
}

function waitForHubSpotRequestSlot_() {
  const lock = LockService.getScriptLock();
  let hasLock = false;

  try {
    lock.waitLock(10000);
    hasLock = true;

    const props = PropertiesService.getScriptProperties();
    const lastAt = Number(props.getProperty(HUBSPOT_LAST_REQUEST_MS_PROP_) || 0);
    let now = Date.now();
    const waitMs = lastAt + HUBSPOT_REQUEST_MIN_INTERVAL_MS_ - now;
    if (waitMs > 0) {
      Utilities.sleep(waitMs);
      now = Date.now();
    }

    props.setProperty(HUBSPOT_LAST_REQUEST_MS_PROP_, String(now));
  } catch (e) {
    if (!hasLock) return;
  } finally {
    if (hasLock) {
      try {
        lock.releaseLock();
      } catch (ignore) {}
    }
  }
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
  if (contactingStart0 === -1) return Logger.log("❌ 'Contacting' section not found in Creator List.");
  const contactingEnd0 = findNextSectionStart0_(creatorData, contactingStart0);
  const creatorEnd0 = contactingEnd0 === -1 ? creatorData.length : contactingEnd0;

  const negotiationStart0 = findSectionRowByLabel_(pitchingData, "Negotiation");
  if (negotiationStart0 === -1) return Logger.log("❌ 'Negotiation' section not found in Pitching.");

  // Find responded rows
  const respondedRows = [];
  for (let r = contactingStart0 + 1; r < creatorEnd0; r++) {
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
    const key = buildCampaignSyncCompositeKey_(row, intHeader);
    if (key && !intRowMap.has(key)) intRowMap.set(key, r);
  }

  // Column mapping: EXT → INT, excluding the negotiation rate block and CPM.
  const colMap = buildCrossSheetColumnMap_(extHeader, intHeader, EXT_TO_INT_PITCHING_MAP_, EXT_ONLY_PITCHING_COLS_);
  const hubSpotRecordIdIdx = findHeaderIndex_(intHeader, "HubSpot Record ID");
  const statusIdx = findHeaderIndex_(intHeader, "Status");
  const hubSpotPitchingStatusItems = [];

  let updated = 0;
  for (let r = 1; r < extData.length; r++) {
    const extRow = extData[r];
    if (isBlankRow_(extRow) || isSectionLabelRow_(extRow)) continue;

    const key = buildCampaignSyncCompositeKey_(extRow, extHeader);
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

    if (hubSpotRecordIdIdx !== -1 && statusIdx !== -1) {
      hubSpotPitchingStatusItems.push({
        row1: intRowIdx + 1,
        recordId: String(intRow[hubSpotRecordIdIdx] || "").trim(),
        status: String(intRow[statusIdx] || "").trim()
      });
    }

    if (writeChangedRowCells_(intPitching, intRowIdx + 1, intHeader, intRow, changedIndexes, INT_PITCHING_FORMULA_COLS_)) {
      updated++;
    }
  }

  Logger.log(`✅ Updated ${updated} row(s) in INT Pitching from EXT.`);
  if (hubSpotRecordIdIdx === -1 || statusIdx === -1) {
    return Logger.log("ℹ️ HubSpot Pitching Status sync skipped. Missing 'HubSpot Record ID' or 'Status' column in Pitching.");
  }
  syncPitchingStatusesToHubSpot_(hubSpotPitchingStatusItems);
}


// ============================================================
//  (5) UPDATE INT CAMPAIGNS FROM EXT
// ============================================================

function updateCampaignsFromExt() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const intCampaigns = ss.getSheetByName("Campaigns");
  const intPitching = ss.getSheetByName("Pitching");
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
  const pitchingIntRateLookup = buildCampaignIntRateLookupFromPitching_(intPitching);

  // Build INT row lookup by composite key
  const intRowMap = new Map();
  for (let r = 1; r < intData.length; r++) {
    const row = intData[r];
    if (isBlankRow_(row) || isSectionLabelRow_(row)) continue;
    const key = buildCampaignSyncCompositeKey_(row, intHeader);
    if (key && !intRowMap.has(key)) intRowMap.set(key, r);
  }

  // Column mapping: EXT → INT (Rate→EXT Rate)
  const colMap = buildCrossSheetColumnMap_(extHeader, intHeader, EXT_TO_INT_CAMPAIGNS_MAP_, new Set());

  let updated = 0;
  let inserted = 0;

  for (let r = 1; r < extData.length; r++) {
    const extRow = extData[r];
    if (isBlankRow_(extRow) || isSectionLabelRow_(extRow)) continue;

    const key = buildCampaignSyncCompositeKey_(extRow, extHeader);
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
      if (applyCampaignIntRateFromPitchingLookup_(intRow, intHeader, pitchingIntRateLookup)) {
        changed = true;
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
      applyCampaignIntRateFromPitchingLookup_(out, intHeader, pitchingIntRateLookup);

      intCampaigns.insertRowBefore(insertAt1);
      const targetRange = intCampaigns.getRange(insertAt1, 1, 1, intNumCols);
      const templateRange = getSafeFormatRow_(intCampaigns, insertAt1 + 1, intNumCols);
      templateRange.copyTo(targetRange, { formatOnly: true });
      targetRange.setValues([out]);

      inserted++;
    }
  }

  Logger.log(`✅ INT Campaigns: ${updated} updated, ${inserted} inserted from EXT.`);
  try {
    syncCampaignRowsToHubSpot_(intCampaigns);
  } catch (e) {
    Logger.log("⚠️ INT Campaigns HubSpot sync failed: " + (e && e.stack ? e.stack : e));
  }
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

  // Column mapping: EXT → INT using matching column names, excluding formula columns.
  const colMap = buildCrossSheetColumnMap_(extHeader, intHeader, {}, INT_PERFORMANCE_FORMULA_COLS_);
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
  try {
    syncPerformanceRowsToHubSpot_(intPerformance);
  } catch (e) {
    Logger.log("⚠️ INT Performance HubSpot sync failed: " + (e && e.stack ? e.stack : e));
  }
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

function buildCampaignIntRateLookupFromPitching_(sheet) {
  const lookup = new Map();
  if (!sheet) return lookup;

  const data = sheet.getDataRange().getValues();
  if (!data || data.length < 2) return lookup;

  const header = (data[0] || []).map(function (value) {
    return String(value || "").trim();
  });
  const dealIdCol = findHeaderIndex_(header, "HubSpot Record ID");
  if (dealIdCol === -1) return lookup;

  for (let r = 1; r < data.length; r++) {
    const row = data[r];
    if (isBlankRow_(row) || isSectionLabelRow_(row)) continue;

    const dealId = String(row[dealIdCol] || "").trim();
    const extRate = getNegotiatedPitchingValueByHeader_(row, header, "EXT Rate");
    const intRate = getNegotiatedPitchingValueByHeader_(row, header, "INT Rate");
    if (!dealId || extRate === "" || intRate === "") continue;

    lookup.set(buildCampaignDealRateLookupKey_(dealId, extRate), intRate);
  }

  return lookup;
}

function applyCampaignIntRateFromPitchingLookup_(campaignRow, campaignHeader, lookup) {
  if (!campaignRow || !Array.isArray(campaignHeader) || !lookup || lookup.size === 0) return false;

  const intRateIdx = findHeaderIndex_(campaignHeader, "INT Rate");
  if (intRateIdx === -1) return false;

  const intRate = getCampaignIntRateFromPitchingLookup_(campaignRow, campaignHeader, lookup);
  if (intRate === "") return false;

  const currentValue = String(campaignRow[intRateIdx] == null ? "" : campaignRow[intRateIdx]).trim();
  const nextValue = String(intRate == null ? "" : intRate).trim();
  if (currentValue === nextValue) return false;

  campaignRow[intRateIdx] = intRate;
  return true;
}

function getCampaignIntRateFromPitchingLookup_(campaignRow, campaignHeader, lookup) {
  const dealId = String(getValueByHeader_(campaignRow, campaignHeader, "HubSpot Record ID") || "").trim();
  const extRate = getValueByHeader_(campaignRow, campaignHeader, "EXT Rate");
  if (!dealId || extRate === "") return "";

  const key = buildCampaignDealRateLookupKey_(dealId, extRate);
  return lookup.has(key) ? lookup.get(key) : "";
}

function buildCampaignDealRateLookupKey_(dealId, rate) {
  const normalizedDealId = String(dealId || "").trim();
  const normalizedRate = normalizeCampaignRateLookupValue_(rate);
  return normalizedDealId && normalizedRate ? (normalizedDealId + "||" + normalizedRate) : "";
}

function normalizeCampaignRateLookupValue_(value) {
  if (value == null) return "";
  if (typeof value === "number") {
    return isFinite(value) ? String(value) : "";
  }

  const text = String(value || "").trim();
  if (!text) return "";

  const numericText = text.replace(/,/g, "");
  if (/^-?\d+(?:\.\d+)?$/.test(numericText)) {
    return String(Number(numericText));
  }

  return text;
}

function getNegotiatedPitchingValueByHeader_(row, header, startHeaderName) {
  const startIdx = findHeaderIndex_(header, startHeaderName);
  if (startIdx === -1) return "";

  let endIdx = header.length - 1;
  for (let c = startIdx + 1; c < header.length; c++) {
    if (String(header[c] || "").trim() !== "") {
      endIdx = c - 1;
      break;
    }
  }

  for (let c = endIdx; c >= startIdx; c--) {
    const value = row[c];
    if (String(value == null ? "" : value).trim() !== "") return value;
  }

  return "";
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

function buildCampaignSyncCompositeKey_(row, header) {
  const activationId = String(getValueByHeader_(row, header, "HubSpot Activation ID") || "").trim();
  if (activationId) {
    return [
      activationId,
      String(getValueByHeader_(row, header, "Deal Type") || "").trim(),
      String(getValueByHeader_(row, header, "Activation Type") || "").trim()
    ].join("||");
  }

  return buildCompositeKey_(row, header);
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

function getClientNameByCampaignName_(ss, sheetName) {
  const dropdownSheet = ss.getSheetByName(sheetName || "Dropdown Values");
  if (!dropdownSheet) return {};

  const data = dropdownSheet.getDataRange().getDisplayValues();
  if (data.length < 2) return {};

  const header = (data[0] || []).map(value => String(value || "").trim());
  const campaignCol = findHeaderIndex_(header, "Campaign Name");
  let clientCol = findHeaderIndex_(header, "Client");
  if (clientCol === -1) clientCol = findHeaderIndex_(header, "Client name");
  if (campaignCol === -1 || clientCol === -1) return {};

  const out = {};
  for (let r = 1; r < data.length; r++) {
    const campaignName = String((data[r] && data[r][campaignCol]) || "").trim();
    const clientName = String((data[r] && data[r][clientCol]) || "").trim();
    const key = normalizeDropdownLookupValue_(campaignName);
    if (!key || !clientName || Object.prototype.hasOwnProperty.call(out, key)) continue;
    out[key] = clientName;
  }

  return out;
}

function getClientNameForCampaign_(campaignName, clientNameByCampaignName) {
  const key = normalizeDropdownLookupValue_(campaignName);
  return key && clientNameByCampaignName ? String(clientNameByCampaignName[key] || "").trim() : "";
}

function normalizeDropdownLookupValue_(value) {
  return String(value || "")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
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

function getArchiveSheetContext_(sheet, fallbackHeader) {
  if (!sheet) return null;

  let data = sheet.getDataRange().getValues();
  let headerRow0 = detectHeaderRow0ForColumns_(
    data,
    ["Status", CREATOR_LIST_TIMESTAMP_IMPORTED_HEADER_],
    10
  );

  if (headerRow0 === -1) {
    const header = Array.isArray(fallbackHeader) && fallbackHeader.length > 0
      ? fallbackHeader.slice()
      : ["Status", CREATOR_LIST_TIMESTAMP_IMPORTED_HEADER_];

    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    data = sheet.getDataRange().getValues();
    headerRow0 = 0;
  }

  const resolvedHeader = (data[headerRow0] || []).map(function (value) {
    return String(value || "").trim();
  });

  return {
    data: data,
    header: resolvedHeader,
    headerRow0: headerRow0,
    headerRow1: headerRow0 + 1
  };
}

function parseSpreadsheetDateValue_(value) {
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return isNaN(value.getTime()) ? null : value;
  }

  if (typeof value === "number" && isFinite(value)) {
    const millis = Math.round((value - 25569) * 24 * 60 * 60 * 1000);
    const numericDate = new Date(millis);
    return isNaN(numericDate.getTime()) ? null : numericDate;
  }

  const text = String(value == null ? "" : value).trim();
  if (!text) return null;

  const parsed = new Date(text);
  return isNaN(parsed.getTime()) ? null : parsed;
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
  const channelUrl = String(getValueByHeader_(row, header, "Channel URL") || "").trim();
  if (looksLikeYouTubeUrl_(channelUrl)) return channelUrl;

  const youtubeUrl = String(getValueByHeader_(row, header, "YouTube URL") || "").trim();
  if (looksLikeYouTubeInput_(youtubeUrl)) {
    if (youtubeUrl.charAt(0) === "@") return "https://www.youtube.com/" + youtubeUrl;
    return youtubeUrl;
  }

  const youtubeHandle = normalizeCreatorPlatformHandle_("youtube", getValueByHeader_(row, header, "YouTube Handle"));
  if (youtubeHandle) {
    return "https://www.youtube.com/" + youtubeHandle;
  }

  return "";
}

function looksLikeYouTubeInput_(value) {
  const text = String(value || "").trim().toLowerCase();
  if (!text) return false;
  if (text.charAt(0) === "@") return true;
  return looksLikeYouTubeUrl_(text);
}

function looksLikeYouTubeUrl_(value) {
  const text = String(value || "").trim().toLowerCase();
  if (!text) return false;
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
    Logger.log("❌ Missing YouTube API key. Set YOUTUBE_API_KEY_");
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
  const configured = String(
    PropertiesService.getScriptProperties().getProperty(YOUTUBE_API_KEY_PROP_) || ""
  ).trim();
  return configured || String(YOUTUBE_API_KEY_ || "").trim();
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
    Logger.log("❌ YouTube API key invalid. Update YOUTUBE_API_KEY_ and rerun.");
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

/**
 * Returns UTF-8 CSV bytes with BOM.
 */
function toCsvBytes_(grid) {
  const csv = (grid || [])
    .map(function (row) {
      return (row || [])
        .map(function (cell) {
          if (cell === null || cell === undefined) return "";
          let text = String(cell);
          text = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
          const needsQuotes = /[",\n]/.test(text);
          text = text.replace(/"/g, '""');
          return needsQuotes ? ('"' + text + '"') : text;
        })
        .join(",");
    })
    .join("\r\n");

  return Utilities.newBlob("\uFEFF" + csv, "text/csv;charset=utf-8").getBytes();
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
  if (!/^[A-Za-z0-9!#$%&'*+=?^_`{|}~.-]+$/.test(local)) {
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
