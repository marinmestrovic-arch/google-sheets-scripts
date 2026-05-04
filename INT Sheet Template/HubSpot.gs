/*********************************************
 * INT Sheet Template - HubSpot menu actions
 *********************************************/

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🧰 Scripts")
    .addItem("Enrich Creator List", "enrichCreatorListRowsInBatches")
    .addItem("Import to HubSpot", "importCreatorListToHubSpotOnly")
    .addItem("Woodpecker Export", "downloadCreatorListWoodpeckerCsv")
    .addToUi();

  SpreadsheetApp.getUi()
    .createMenu("⏰ Triggers")
    .addItem("Sync Dropdown Values from HubSpot", "syncDropdownValuesFromHubSpot")
    .addToUi();
}

function enrichCreatorListRowsInBatches() {
  return INT_HUBSPOT_MENU_.enrichCreatorListRowsInBatches();
}

function importCreatorListToHubSpotOnly() {
  return INT_HUBSPOT_MENU_.importCreatorListToHubSpotOnly();
}

function downloadCreatorListWoodpeckerCsv() {
  return INT_HUBSPOT_MENU_.downloadCreatorListWoodpeckerCsv();
}

function syncDropdownValuesFromHubSpot() {
  return INT_HUBSPOT_MENU_.syncDropdownValuesFromHubSpot();
}

const INT_HUBSPOT_MENU_ = (function () {
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
  const HUBSPOT_DEALS_OBJECT_API_NAME_ = "deals";
  const HUBSPOT_CONTACTS_OBJECT_API_NAME_ = "contacts";
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

  const OPENAI_API_KEY_ = "api key here";
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
  const YOUTUBE_API_KEY_ = "api key here";
  const YT_KEY_VALID_CACHE_PREFIX_ = "YT_API_KEY_VALID::";
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

  const CREATOR_LIST_TIMESTAMP_IMPORTED_HEADER_ = "Timestamp Imported";
  const WOODPECKER_EXPORT_HEADERS_ = ["Email", "Channel Name", "First Name", "Last Name"];
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
    let nextStartRow1 = Math.max(context.headerRow1 + 1, 3);
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

    const archivedStart0 = findSectionRowByLabel_(data, "Archived");
    const startRow1 = Math.max(context.headerRow1 + 1, 3);
    const stopRow0 = archivedStart0 === -1 ? data.length : archivedStart0;

    const candidateRowItems = [];
    for (let r = startRow1 - 1; r < stopRow0; r++) {
      const row = data[r];
      if (isBlankRow_(row) || isSectionLabelRow_(row)) continue;
      candidateRowItems.push({ row1: r + 1, values: row.slice() });
    }

    if (candidateRowItems.length === 0) {
      return ui.alert("ℹ️ No importable rows from row " + startRow1 + " onward.");
    }

    const preparedRows = prepareCreatorListHubSpotImportRows_(header, candidateRowItems);
    if (preparedRows.rowItems.length === 0) {
      return ui.alert(
        "ℹ️ No new unique deals to import.\n\n" +
        "Rows with a filled HubSpot Record ID and duplicate Deal Name values were skipped."
      );
    }

    importCreatorListToHubSpot_(ss, sheet, header, preparedRows.rowItems, ui, preparedRows);
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

  function isKnownSectionLabel_(value) {
    const text = String(value || "").trim();
    if (!text) return false;

    const knownLabels = new Set(
      ["Contacting", "Archived", "Negotiation", "Active Pitches"].concat(MONTH_NAMES_)
    );
    return knownLabels.has(text);
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
      "Channel URL",
      "Campaign Name",
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
    if (youtubeHandle && channelName && youtubeHandle !== channelName) return true;

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
    return String(OPENAI_API_KEY_ || "").trim();
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

  function getValueByHeader_(row, header, columnName) {
    const idx = findHeaderIndex_(header, columnName);
    return idx === -1 ? "" : row[idx];
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

  function getHubSpotApiToken_() {
    const configured = String(
      PropertiesService.getScriptProperties().getProperty(HUBSPOT_API_KEY_PROP_) || ""
    ).trim();
    return configured || String(HUBSPOT_API_KEY_ || "").trim();
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

  function normalizeChannelUrl_(url) {
    return String(url).trim().replace(/\s+/g, "");
  }

  function getYouTubeApiKey_() {
    return String(YOUTUBE_API_KEY_ || "").trim();
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

  return {
    enrichCreatorListRowsInBatches: enrichCreatorListRowsInBatches,
    importCreatorListToHubSpotOnly: importCreatorListToHubSpotOnly,
    downloadCreatorListWoodpeckerCsv: downloadCreatorListWoodpeckerCsv,
    syncDropdownValuesFromHubSpot: syncDropdownValuesFromHubSpot
  };
})();
