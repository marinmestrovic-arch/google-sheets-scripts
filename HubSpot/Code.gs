function fillCrmImport() {
  const SKIP_HEADER = "Deal name";
  const EXPORT_SHEET = "HypeAuditor Export";
  const CRM_SHEET = "HubSpot Import";
  const MAPPING_SHEET = "Mapping";

  // --- Language value mapping ---
  // Create a sheet tab with 2 columns:
  // A: code (e.g., fr)
  // B: language (e.g., French)
  const LANGUAGE_MAPPING_SHEET = "Language Mapping";
  const LANGUAGE_CRM_HEADER = "Language"; // HubSpot Import column header that should receive full name

  const CONTACT_TYPE_HEADER = "Contact Type";
  const CONTACT_TYPE_VALUE = "Influencer";

  const ss = SpreadsheetApp.getActive();
  const exportSh = ss.getSheetByName(EXPORT_SHEET);
  const crmSh = ss.getSheetByName(CRM_SHEET);
  const mapSh = ss.getSheetByName(MAPPING_SHEET);
  const langSh = ss.getSheetByName(LANGUAGE_MAPPING_SHEET);

  if (!exportSh) throw new Error(`Sheet not found: "${EXPORT_SHEET}"`);
  if (!crmSh) throw new Error(`Sheet not found: "${CRM_SHEET}"`);
  if (!mapSh) throw new Error(`Sheet not found: "${MAPPING_SHEET}"`);
  if (!langSh) throw new Error(`Sheet not found: "${LANGUAGE_MAPPING_SHEET}"`);

  const exportData = exportSh.getDataRange().getValues();
  if (exportData.length < 2) return;

  const exportHeaders = exportData[0];
  const exportRows = exportData.slice(1);

  const crmHeaders = crmSh.getRange(1, 1, 1, crmSh.getLastColumn()).getValues()[0];

  // Pull validations sized to exportRows (so row-indexed validation checks work)
  const crmValidations = crmSh
    .getRange(2, 1, Math.max(exportRows.length, 1), crmHeaders.length)
    .getDataValidations();

  // Mapping (guard for empty mapping sheet)
  const mapLastRow = mapSh.getLastRow();
  const mappingData = mapLastRow > 1 ? mapSh.getRange(2, 1, mapLastRow - 1, 2).getValues() : [];

  // Export header → index
  const exportIndex = {};
  exportHeaders.forEach((h, i) => (exportIndex[h] = i));

  // CRM → Export mapping
  const mapping = {};
  mappingData.forEach(([crm, exp]) => {
    if (crm && exp) mapping[crm] = exp;
  });

  // ---- Load language map (code -> full name) ----
  const langLastRow = langSh.getLastRow();
  const langData = langLastRow > 1 ? langSh.getRange(2, 1, langLastRow - 1, 2).getValues() : [];

  const langMap = {};
  langData.forEach(([code, name]) => {
    if (!code || !name) return;
    langMap[String(code).trim().toLowerCase()] = String(name).trim();
  });

  // Normalize incoming language values (supports: "fr", "fr-FR", "fr_FR", " fr ")
  function normalizeLangCode_(v) {
    if (v === null || v === undefined) return "";
    const s = String(v).trim().toLowerCase();
    if (!s) return "";
    return s.split(/[-_]/)[0]; // "fr-FR" -> "fr"
  }

  function normalizeKey_(v) {
    return String(v || "")
      .toLowerCase()
      .replace(/[^a-z0-9]+/g, " ")
      .trim();
  }

  const SOCIAL_ER_HEADER_PAIRS_ = new Set([
    "youtube engagement rate|youtube er",
    "instagram engagement rate|instagram er",
    "tiktok engagement rate|tiktok er",
    "twitter engagement rate|twitter er"
  ]);

  function isSocialErField_(crmHeader, exportHeader) {
    const pairKey = `${normalizeKey_(crmHeader)}|${normalizeKey_(exportHeader)}`;
    return SOCIAL_ER_HEADER_PAIRS_.has(pairKey);
  }

  function parseMaybePercentNumber_(raw) {
    if (raw === null || raw === undefined || raw === "") return null;
    if (typeof raw === "number") return isFinite(raw) ? raw : null;

    const cleaned = String(raw)
      .trim()
      .replace("%", "")
      .replace(",", ".");

    if (!cleaned) return null;
    const parsed = Number(cleaned);
    return isFinite(parsed) ? parsed : null;
  }

  // ---- Validation helpers (prevents setValues() from throwing) ----

  // Build a Set of allowed dropdown values (list OR range). Returns null if not a dropdown.
  function getAllowedSet(validation) {
    if (!validation) return null;

    const criteria = validation.getCriteriaType();
    const critVals = validation.getCriteriaValues();

    if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      const list = critVals?.[0] || [];
      return new Set(list.map(v => String(v)));
    }

    if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
      const range = critVals?.[0];
      if (!range) return null;
      const vals = range.getValues().flat().filter(v => v !== "" && v !== null);
      return new Set(vals.map(v => String(v)));
    }

    return null;
  }

  // Returns true if it's safe to write `value` into a cell with `validation`.
  // For dropdowns: enforce membership.
  // For other validation types we can't easily validate: skip writing to avoid runtime errors.
  function canWriteValue(validation, value) {
    if (value === "" || value === null) return true; // blanks are safe
    if (!validation) return true;

    const criteria = validation.getCriteriaType();

    if (
      criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST ||
      criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE
    ) {
      const allowedSet = getAllowedSet(validation);
      if (!allowedSet) return false;
      return allowedSet.has(String(value));
    }

    // Unknown/other validation types (NUMBER_*, DATE_*, CUSTOM_FORMULA, etc.)
    // To guarantee the script never fails, don't write into these cells.
    return false;
  }

  // ---- Build desired output (invalid dropdown values become "") ----
  const output = exportRows.map((row, rIdx) =>
    crmHeaders.map((header, cIdx) => {
      let value = "";

      if (header === CONTACT_TYPE_HEADER) {
        value = row.some(v => v !== "" && v !== null) ? CONTACT_TYPE_VALUE : "";
      } else {
        const mappedHeader = mapping[header];
        if (!mappedHeader) return "";

        const idx = exportIndex[mappedHeader];
        if (idx === undefined) return "";

        value = row[idx];

        // HypeAuditor exports social ER as percent points (e.g. 3.47),
        // while HubSpot percentage fields expect fractions (0.0347 for 3.47%).
        if (isSocialErField_(header, mappedHeader)) {
          const er = parseMaybePercentNumber_(value);
          value = er === null ? "" : er / 100;
        }
      }

      // Hide locked values
      if (value === "🔒") value = "";

      // --- Special case: map Language codes -> full language name ---
      if (header === LANGUAGE_CRM_HEADER) {
        const code = normalizeLangCode_(value);
        value = langMap[code] || ""; // unknown -> blank (safe for validation)
      }

      // Validation safety (dropdown list/range, and conservative skip for other validation types)
      const validation = crmValidations[rIdx]?.[cIdx];
      if (!canWriteValue(validation, value)) return "";

      return value;
    })
  );

  if (!output.length) return;

  // ---- Read existing values + formulas so we can preserve them ----
  const targetRange = crmSh.getRange(2, 1, output.length, crmHeaders.length);
  const existingValues = targetRange.getValues();
  const existingFormulas = targetRange.getFormulas(); // "" if no formula in that cell

  // Identify columns that have ANY formulas in the target rows (we won't write these columns at all)
  const colHasFormula = crmHeaders.map((_, c) => existingFormulas.some(r => r[c] !== ""));

  // Merge:
  // - don't overwrite existing cells with ""
  // - don't write values that would violate validation (keeps existing instead)
  const nextValues = output.map((row, r) =>
    row.map((val, c) => {
      const validation = crmValidations[r]?.[c];

      if (val === "") return existingValues[r][c];
      if (!canWriteValue(validation, val)) return existingValues[r][c];

      return val;
    })
  );

  const skipCol = crmHeaders.indexOf(SKIP_HEADER); // 0-based
  if (skipCol === -1) throw new Error(`Header not found: "${SKIP_HEADER}"`);

  // Write contiguous blocks, skipping:
  // - the SKIP_HEADER column
  // - any column that contains formulas (so formulas never get overwritten)
  function writeNonFormulaBlocks(colStartInclusive, colEndInclusive) {
    let c = colStartInclusive;

    while (c <= colEndInclusive) {
      // advance to next writable column
      while (c <= colEndInclusive && (c === skipCol || colHasFormula[c])) c++;
      if (c > colEndInclusive) break;

      const blockStart = c;

      // extend block while columns are writable
      while (c <= colEndInclusive && c !== skipCol && !colHasFormula[c]) c++;

      const blockEnd = c - 1;
      const width = blockEnd - blockStart + 1;

      const blockValues = nextValues.map(r => r.slice(blockStart, blockEnd + 1));
      crmSh.getRange(2, blockStart + 1, output.length, width).setValues(blockValues);
    }
  }

  // Write left side (before "Deal name")
  if (skipCol > 0) writeNonFormulaBlocks(0, skipCol - 1);

  // Write right side (after "Deal name")
  if (skipCol < crmHeaders.length - 1) writeNonFormulaBlocks(skipCol + 1, crmHeaders.length - 1);
}

const HUBSPOT_API_BASE_ = "https://api.hubapi.com";
const HUBSPOT_IMPORT_SHEET_ = "HubSpot Import";
const HUBSPOT_IMPORT_EMAIL_HEADER_ = "Email";
// HUBSPOT_API_KEY is accepted as a legacy property name, but the value must
// still be a private app or service key token for HubSpot's current APIs.
const HUBSPOT_IMPORT_TOKEN_PROPS_ = [
  "HUBSPOT_ACCESS_TOKEN",
  "HUBSPOT_PRIVATE_APP_TOKEN",
  "HUBSPOT_API_KEY"
];
const HUBSPOT_SHARED_LIBRARY_IDENTIFIER_ = "HubSpotSharedImporter";
// Set this once in the template file after deploying the shared importer web app.
// All future copies created from the template will inherit the URL automatically.
const HUBSPOT_SHARED_IMPORT_WEB_APP_URL_ = "https://script.google.com/a/macros/arch.agency/s/AKfycbzI6gAHnhSlRLzheWkS_wNYvYODvx1aztd27cf2DbbuBJTSqOYe-oqtKAnZqRc7jCE8/exec";
const HUBSPOT_SHARED_IMPORT_ACTION_ = "startImport";

const HUBSPOT_IMPORT_OBJECT_SPECS_ = [
  {
    key: "contacts",
    label: "Contacts",
    objectTypeId: "0-1",
    aliases: ["Contact", "Contacts"]
  },
  {
    key: "deals",
    label: "Deals",
    objectTypeId: "0-3",
    aliases: ["Deal", "Deals"]
  },
  {
    key: "clientCampaigns",
    label: "Client Campaigns",
    aliases: ["Client Campaign", "Client Campaigns"]
  },
  {
    key: "clients",
    label: "Clients",
    aliases: ["Client", "Clients"]
  },
  {
    key: "activations",
    label: "Activations",
    aliases: ["Activation", "Activations"]
  }
];

const HUBSPOT_IMPORT_HEADER_OVERRIDES_ = {
  email: {
    objectKey: "contacts",
    propertyName: "email",
    columnType: "HUBSPOT_ALTERNATE_ID"
  },
  "first name": {
    objectKey: "contacts",
    propertyName: "firstname"
  },
  "last name": {
    objectKey: "contacts",
    propertyName: "lastname"
  },
  "deal name": {
    objectKey: "deals",
    propertyName: "dealname"
  }
};

const HUBSPOT_IMPORT_HEADER_OBJECT_HINTS_ = {
  "contact type": "contacts",
  "campaign name": "clientCampaigns",
  month: "clientCampaigns",
  year: "clientCampaigns",
  "client name": "clients",
  "deal owner": "deals",
  "deal name": "deals",
  pipeline: "deals",
  "deal stage": "deals",
  amount: "deals",
  currency: "deals",
  "deal type": "deals",
  "activation type": "deals",
  "first name": "contacts",
  "last name": "contacts",
  email: "contacts",
  "phone number": "contacts",
  "influencer type": "contacts",
  "influencer vertical": "contacts",
  "country region": "contacts",
  language: "contacts",
  "youtube handle": "contacts",
  "youtube url": "contacts",
  "youtube average views": "contacts",
  "youtube engagement rate": "contacts",
  "youtube followers": "contacts",
  "instagram handle": "contacts",
  "instagram url": "contacts",
  "instagram post average views": "contacts",
  "instagram reel average views": "contacts",
  "instagram story 7 day average views": "contacts",
  "instagram story 30 day average views": "contacts",
  "instagram engagement rate": "contacts",
  "instagram followers": "contacts",
  "tiktok handle": "contacts",
  "tiktok url": "contacts",
  "tiktok average views": "contacts",
  "tiktok engagement rate": "contacts",
  "tiktok followers": "contacts",
  "twitch handle": "contacts",
  "twitch url": "contacts",
  "twitch average views": "contacts",
  "twitch engagement rate": "contacts",
  "twitch followers": "contacts",
  "kick handle": "contacts",
  "kick url": "contacts",
  "kick average views": "contacts",
  "kick engagement rate": "contacts",
  "kick followers": "contacts",
  "x handle": "contacts",
  "x url": "contacts",
  "x average views": "contacts",
  "x engagement rate": "contacts",
  "x followers": "contacts"
};

function exportToCsv() {
  return importToHubSpot();
}

function importToHubSpot() {
  const ui = SpreadsheetApp.getUi();

  try {
    const ss = SpreadsheetApp.getActive();
    const sheetPayload = prepareHubSpotImportSheetPayload_(ss);
    if (sheetPayload.emptyReason) {
      ui.alert(sheetPayload.emptyReason);
      return;
    }

    if (hasHubSpotSharedImporterLibrary_()) {
      const startedImport = startHubSpotImportViaLibrary_(sheetPayload);
      ui.alert(
        buildHubSpotImportSubmittedMessage_({
          rowCount: startedImport.rowCount || sheetPayload.rowCount,
          columnCount: startedImport.columnCount || sheetPayload.columnCount,
          objectLabels: startedImport.objectLabels || [],
          importId: startedImport.importId,
          importState: startedImport.state
        })
      );
      return;
    }

    const sharedWebAppUrl = getHubSpotSharedImportWebAppUrl_();
    if (sharedWebAppUrl) {
      const startedImport = startHubSpotImportViaWebApp_(sharedWebAppUrl, sheetPayload);
      ui.alert(
        buildHubSpotImportSubmittedMessage_({
          rowCount: startedImport.rowCount || sheetPayload.rowCount,
          columnCount: startedImport.columnCount || sheetPayload.columnCount,
          objectLabels: startedImport.objectLabels || [],
          importId: startedImport.importId,
          importState: startedImport.state
        })
      );
      return;
    }

    const token = getHubSpotImportToken_();
    if (!token) {
      ui.alert(
        "Missing HubSpot importer configuration.\n\n" +
        "Recommended: add the shared Apps Script library with identifier " +
        HUBSPOT_SHARED_LIBRARY_IDENTIFIER_ +
        ", or set HUBSPOT_SHARED_IMPORT_WEB_APP_URL_ in this file.\n\n" +
        "Fallback: set one of these Script Properties in this copy: " +
        HUBSPOT_IMPORT_TOKEN_PROPS_.join(", ") +
        "."
      );
      return;
    }

    const prepared = prepareHubSpotImport_(sheetPayload, token);
    const startedImport = startHubSpotImport_(token, prepared.fileBlob, prepared.importRequest);

    ui.alert(
      buildHubSpotImportSubmittedMessage_({
        rowCount: prepared.rowCount,
        columnCount: prepared.columnCount,
        objectLabels: prepared.objectLabels,
        importId: startedImport && startedImport.id != null ? String(startedImport.id) : "not returned",
        importState: startedImport && startedImport.state ? startedImport.state : "STARTED"
      })
    );
  } catch (error) {
    const message = error && error.message ? error.message : String(error);
    Logger.log("HubSpot import failed: " + message);
    ui.alert("HubSpot import failed.\n\n" + message);
  }
}

function prepareHubSpotImportSheetPayload_(ss) {
  const sh = ss.getSheetByName(HUBSPOT_IMPORT_SHEET_);
  if (!sh) throw new Error(`Sheet not found: "${HUBSPOT_IMPORT_SHEET_}"`);

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) {
    return { emptyReason: "No data to import." };
  }

  const values = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();
  const headers = values[0].map(v => String(v || "").trim());
  const rows = values.slice(1);
  const importRows = filterHubSpotImportRows_(rows, headers);

  if (importRows.length === 0) {
    const emailIdx = headers.indexOf(HUBSPOT_IMPORT_EMAIL_HEADER_);
    return {
      emptyReason: emailIdx === -1
        ? "No populated rows to import."
        : `No rows with a filled ${HUBSPOT_IMPORT_EMAIL_HEADER_} column to import.`
    };
  }

  const activeColIndexes = getActiveHubSpotImportColumnIndexes_(headers, importRows);
  if (activeColIndexes.length === 0) {
    return { emptyReason: "No populated columns to import." };
  }

  const activeHeaders = activeColIndexes.map(idx => headers[idx]);
  const duplicateHeaders = findDuplicateHubSpotHeaders_(activeHeaders);
  if (duplicateHeaders.length) {
    throw new Error(
      "Duplicate HubSpot Import headers found: " + duplicateHeaders.join(", ")
    );
  }

  const activeRows = importRows.map(row => activeColIndexes.map(idx => row[idx]));
  const spreadsheetName = ss.getName().replace(/[\\/:*?"<>|]+/g, "-");

  return {
    spreadsheetName,
    spreadsheetLocale: ss.getSpreadsheetLocale(),
    headers: activeHeaders,
    rows: activeRows,
    rowCount: activeRows.length,
    columnCount: activeHeaders.length
  };
}

function prepareHubSpotImport_(sheetPayload, token) {
  const activeHeaders = sheetPayload.headers;
  const activeRows = sheetPayload.rows;
  const objectCatalog = loadHubSpotImportObjectCatalog_(token);
  const resolvedColumns = resolveHubSpotImportColumns_(activeHeaders, objectCatalog);

  if (resolvedColumns.errors.length) {
    throw new Error(
      "Some HubSpot Import columns could not be mapped automatically.\n\n" +
      resolvedColumns.errors.join("\n") +
      "\n\nUse the HubSpot property label or internal name in the sheet header. " +
      "For ambiguous columns, prefix the header with the object name, for example: " +
      '"Contacts Email", "Deals Record ID", or "Client Campaigns Name".'
    );
  }

  const fileName = `HubSpot Import - ${sheetPayload.spreadsheetName}.csv`;
  const fileBlob = Utilities.newBlob(toCsvBytes_([activeHeaders].concat(activeRows)), "text/csv", fileName);
  const importRequest = buildHubSpotImportRequest_(
    fileName,
    resolvedColumns.mappings,
    sheetPayload.spreadsheetLocale
  );

  return {
    fileBlob,
    importRequest,
    rowCount: sheetPayload.rowCount,
    columnCount: sheetPayload.columnCount,
    objectLabels: resolvedColumns.objectLabels
  };
}

function buildHubSpotImportSubmittedMessage_(details) {
  return (
    "HubSpot import submitted.\n\n" +
    `Rows: ${details.rowCount}\n` +
    `Columns: ${details.columnCount}\n` +
    `Objects: ${(details.objectLabels || []).join(", ")}\n` +
    `Import ID: ${details.importId || "not returned"}\n` +
    `State: ${details.importState || "STARTED"}\n\n` +
    "HubSpot continues processing the import in the background."
  );
}

/**
 * Returns UTF-8 CSV BYTES with BOM (matches Google Sheets download behavior more closely).
 */
function toCsvBytes_(grid) {
  const csv = grid
    .map(row =>
      row
        .map(cell => {
          if (cell === null || cell === undefined) return "";
          let s = String(cell);

          // Normalize line breaks like Sheets CSV export
          s = s.replace(/\r\n/g, "\n").replace(/\r/g, "\n");

          const needsQuotes = /[",\n]/.test(s);
          s = s.replace(/"/g, '""');
          return needsQuotes ? `"${s}"` : s;
        })
        .join(",")
    )
    .join("\r\n");

  // UTF-8 BOM so Excel + others keep special chars correctly
  const withBom = "\uFEFF" + csv;

  // Convert to bytes explicitly (UTF-8)
  return Utilities.newBlob(withBom, "text/csv;charset=utf-8").getBytes();
}

function filterHubSpotImportRows_(rows, headers) {
  const emailIdx = headers.indexOf(HUBSPOT_IMPORT_EMAIL_HEADER_);

  return rows.filter(row => {
    const hasData = row.some(cell => String(cell || "").trim() !== "");
    if (!hasData) return false;

    if (emailIdx === -1) return true;

    const email = row[emailIdx];
    return email !== null && email !== undefined && String(email).trim() !== "";
  });
}

function getActiveHubSpotImportColumnIndexes_(headers, rows) {
  const activeIndexes = [];
  const blankHeadersWithData = [];

  headers.forEach((header, idx) => {
    const hasData = rows.some(row => String(row[idx] || "").trim() !== "");
    if (!hasData) return;

    if (!String(header || "").trim()) {
      blankHeadersWithData.push(columnNumberToLetter_(idx + 1));
      return;
    }

    activeIndexes.push(idx);
  });

  if (blankHeadersWithData.length) {
    throw new Error(
      "These populated columns are missing a header: " + blankHeadersWithData.join(", ")
    );
  }

  return activeIndexes;
}

function buildHubSpotImportRequest_(fileName, columnMappings, spreadsheetLocale) {
  return {
    name: `Google Sheets Import - ${fileName.replace(/\.csv$/i, "")}`,
    dateFormat: getHubSpotImportDateFormat_(spreadsheetLocale),
    importOperations: buildHubSpotImportOperations_(columnMappings),
    files: [
      {
        fileName,
        fileFormat: "CSV",
        fileImportPage: {
          hasHeader: true,
          columnMappings
        }
      }
    ]
  };
}

function buildHubSpotImportOperations_(columnMappings) {
  const operations = {};

  columnMappings.forEach(mapping => {
    if (!mapping || mapping.columnType === "FLEXIBLE_ASSOCIATION_LABEL") return;

    const objectTypeId = mapping.columnObjectTypeId;
    if (!objectTypeId || operations[objectTypeId] === "UPSERT") return;

    operations[objectTypeId] =
      mapping.columnType === "HUBSPOT_OBJECT_ID" ||
      mapping.columnType === "HUBSPOT_ALTERNATE_ID"
        ? "UPSERT"
        : "CREATE";
  });

  return operations;
}

function loadHubSpotImportObjectCatalog_(token) {
  const schemasResponse = hubspotFetchJson_(
    `${HUBSPOT_API_BASE_}/crm-object-schemas/v3/schemas`,
    token
  );
  const schemas = Array.isArray(schemasResponse && schemasResponse.results)
    ? schemasResponse.results
    : [];

  const list = HUBSPOT_IMPORT_OBJECT_SPECS_.map(spec => {
    const schema = resolveHubSpotObjectSchema_(spec, schemas);
    const propertiesResponse = hubspotFetchJson_(
      `${HUBSPOT_API_BASE_}/crm/v3/properties/${encodeURIComponent(schema.objectTypeId)}`,
      token
    );
    const properties = Array.isArray(propertiesResponse && propertiesResponse.results)
      ? propertiesResponse.results
      : [];

    if (!properties.length) {
      throw new Error(`No HubSpot properties returned for ${spec.label}.`);
    }

    const propertyLookup = { label: {}, name: {} };
    properties.forEach(property => {
      addPropertyLookupEntry_(propertyLookup.label, property.label, property);
      addPropertyLookupEntry_(propertyLookup.name, property.name, property);
    });

    const aliasSet = {};
    spec.aliases.forEach(alias => {
      const key = normalizeHubSpotLookupKey_(alias);
      if (key) aliasSet[key] = true;
    });

    const singular = schema.labels && schema.labels.singular ? schema.labels.singular : "";
    const plural = schema.labels && schema.labels.plural ? schema.labels.plural : "";
    const schemaName = schema.name || "";

    [singular, plural, schemaName, spec.label].forEach(alias => {
      const key = normalizeHubSpotLookupKey_(alias);
      if (key) aliasSet[key] = true;
    });

    const primaryDisplayProperty = properties.find(
      property => property.name === schema.primaryDisplayProperty
    ) || null;

    return {
      key: spec.key,
      label: plural || singular || spec.label,
      objectTypeId: schema.objectTypeId,
      properties,
      propertyLookup,
      aliases: Object.keys(aliasSet),
      primaryDisplayProperty
    };
  });

  const byKey = {};
  list.forEach(objectInfo => {
    byKey[objectInfo.key] = objectInfo;
  });

  return { list, byKey };
}

function resolveHubSpotObjectSchema_(spec, schemas) {
  if (spec.objectTypeId) {
    return {
      objectTypeId: spec.objectTypeId,
      name: spec.key,
      labels: {
        singular: spec.aliases[0],
        plural: spec.label
      },
      primaryDisplayProperty: spec.key === "deals" ? "dealname" : ""
    };
  }

  const wanted = spec.aliases.map(alias => normalizeHubSpotLookupKey_(alias));
  const matches = schemas.filter(schema => {
    const candidates = [
      schema && schema.name,
      schema && schema.labels && schema.labels.singular,
      schema && schema.labels && schema.labels.plural
    ]
      .map(value => normalizeHubSpotLookupKey_(value))
      .filter(Boolean);

    return wanted.some(alias => candidates.indexOf(alias) !== -1);
  });

  if (matches.length === 0) {
    throw new Error(`Could not find the HubSpot custom object schema for ${spec.label}.`);
  }

  if (matches.length > 1) {
    throw new Error(`Multiple HubSpot schemas matched ${spec.label}.`);
  }

  return matches[0];
}

function resolveHubSpotImportColumns_(headers, objectCatalog) {
  const mappings = [];
  const objectLabels = [];
  const seenObjectTypeIds = {};
  const errors = [];

  headers.forEach(header => {
    const associationMapping = resolveHubSpotAssociationColumn_(header, objectCatalog.list);
    if (associationMapping) {
      mappings.push(associationMapping);
      return;
    }

    const overrideMapping = resolveHubSpotHeaderOverride_(header, objectCatalog);
    if (overrideMapping) {
      mappings.push(overrideMapping.mapping);
      if (!seenObjectTypeIds[overrideMapping.objectInfo.objectTypeId]) {
        seenObjectTypeIds[overrideMapping.objectInfo.objectTypeId] = true;
        objectLabels.push(overrideMapping.objectInfo.label);
      }
      return;
    }

    const resolvedProperty = resolveHubSpotPropertyColumn_(header, objectCatalog.list);
    if (resolvedProperty.error) {
      errors.push(`- ${header}: ${resolvedProperty.error}`);
      return;
    }

    mappings.push(resolvedProperty.mapping);

    if (!seenObjectTypeIds[resolvedProperty.objectInfo.objectTypeId]) {
      seenObjectTypeIds[resolvedProperty.objectInfo.objectTypeId] = true;
      objectLabels.push(resolvedProperty.objectInfo.label);
    }
  });

  return { mappings, objectLabels, errors };
}

function resolveHubSpotHeaderOverride_(header, objectCatalog) {
  const key = normalizeHubSpotLookupKey_(header);
  const override = HUBSPOT_IMPORT_HEADER_OVERRIDES_[key];
  if (!override) return null;

  const objectInfo = objectCatalog.byKey[override.objectKey];
  if (!objectInfo) return null;

  const property = findHubSpotPropertyByName_(objectInfo, override.propertyName);
  if (!property) return null;

  return {
    objectInfo,
    mapping: buildHubSpotPropertyMapping_(header, objectInfo, property, override.columnType)
  };
}

function resolveHubSpotPropertyColumn_(header, objectInfos) {
  const objectHint = extractHubSpotObjectHint_(header, objectInfos);
  const defaultObjectInfo = objectHint ? null : getHubSpotDefaultObjectForHeader_(header, objectInfos);
  const keyVariants = buildHubSpotHeaderKeyVariants_(header, objectHint);
  let matches = collectHubSpotPropertyMatches_(
    keyVariants,
    objectHint ? [objectHint.objectInfo] : (defaultObjectInfo ? [defaultObjectInfo] : objectInfos),
    objectHint ? 10 : (defaultObjectInfo ? 5 : 0)
  );

  if (!matches.length && defaultObjectInfo) {
    matches = collectHubSpotPropertyMatches_(keyVariants, objectInfos, 0);
  }

  const dedupedMatches = dedupeHubSpotPropertyMatches_(matches);
  if (!dedupedMatches.length) {
    const preferredObjectInfo = objectHint ? objectHint.objectInfo : defaultObjectInfo;
    const fallbackRemainderKey = objectHint
      ? objectHint.remainderKey
      : normalizeHubSpotLookupKey_(header);

    if (preferredObjectInfo && shouldUsePrimaryDisplayProperty_(fallbackRemainderKey)) {
      const primaryProperty = preferredObjectInfo.primaryDisplayProperty;
      if (primaryProperty) {
        return {
          objectInfo: preferredObjectInfo,
          mapping: buildHubSpotPropertyMapping_(header, preferredObjectInfo, primaryProperty)
        };
      }
    }

    return { error: "no matching property found in Contacts, Deals, Client Campaigns, Clients, or Activations" };
  }

  dedupedMatches.sort((a, b) => b.score - a.score);

  const topMatch = dedupedMatches[0];
  const secondMatch = dedupedMatches[1];
  if (
    secondMatch &&
    secondMatch.score === topMatch.score &&
    (
      secondMatch.objectInfo.objectTypeId !== topMatch.objectInfo.objectTypeId ||
      secondMatch.property.name !== topMatch.property.name
    )
  ) {
    return {
      error:
        "matched multiple HubSpot properties equally well; prefix the header with the object name"
    };
  }

  return {
    objectInfo: topMatch.objectInfo,
    mapping: buildHubSpotPropertyMapping_(header, topMatch.objectInfo, topMatch.property)
  };
}

function resolveHubSpotAssociationColumn_(header, objectInfos) {
  const key = normalizeHubSpotLookupKey_(header);
  if (key.indexOf("association label") === -1) return null;

  const aliasMatches = getHubSpotAliasMatches_(key, objectInfos)
    .sort((a, b) => {
      if (a.start !== b.start) return a.start - b.start;
      return b.alias.length - a.alias.length;
    });

  const chosen = [];
  aliasMatches.forEach(match => {
    const overlaps = chosen.some(
      existing => match.start < existing.end && match.end > existing.start
    );
    if (!overlaps && !chosen.some(existing => existing.objectInfo.key === match.objectInfo.key)) {
      chosen.push(match);
    }
  });

  if (chosen.length < 2) {
    throw new Error(
      `Association label column "${header}" must mention both objects, for example "Contacts to Deals Association Label".`
    );
  }

  chosen.sort((a, b) => a.start - b.start);

  return {
    columnName: header,
    columnType: "FLEXIBLE_ASSOCIATION_LABEL",
    columnObjectTypeId: chosen[0].objectInfo.objectTypeId,
    toColumnObjectTypeId: chosen[1].objectInfo.objectTypeId
  };
}

function collectHubSpotPropertyMatches_(keyVariants, objectInfos, objectBonus) {
  const matches = [];

  objectInfos.forEach(objectInfo => {
    keyVariants.forEach(variant => {
      const labelMatches = objectInfo.propertyLookup.label[variant.key] || [];
      labelMatches.forEach(property => {
        matches.push({
          objectInfo,
          property,
          score: variant.labelScore + objectBonus
        });
      });

      const nameMatches = objectInfo.propertyLookup.name[variant.key] || [];
      nameMatches.forEach(property => {
        matches.push({
          objectInfo,
          property,
          score: variant.nameScore + objectBonus
        });
      });
    });
  });

  return matches;
}

function buildHubSpotPropertyMapping_(columnName, objectInfo, property, forcedColumnType) {
  const mapping = {
    columnName,
    columnObjectTypeId: objectInfo.objectTypeId,
    propertyName: property.name
  };

  const columnType = forcedColumnType || getHubSpotPropertyColumnType_(objectInfo, property);
  if (columnType) mapping.columnType = columnType;

  return mapping;
}

function getHubSpotPropertyColumnType_(objectInfo, property) {
  if (!property) return "";
  if (property.name === "hs_object_id") return "HUBSPOT_OBJECT_ID";

  if (property.hasUniqueValue || (objectInfo.key === "contacts" && property.name === "email")) {
    return "HUBSPOT_ALTERNATE_ID";
  }

  return "";
}

function buildHubSpotHeaderKeyVariants_(header, objectHint) {
  const variants = [];

  addHubSpotHeaderKeyVariants_(variants, header, 120, 110);

  if (objectHint && objectHint.remainderOriginal) {
    addHubSpotHeaderKeyVariants_(variants, objectHint.remainderOriginal, 100, 90);
  }

  if (objectHint && !objectHint.remainderOriginal) {
    addHubSpotHeaderKeyVariants_(variants, "name", 75, 65);
  }

  return variants;
}

function addHubSpotHeaderKeyVariants_(variants, rawValue, labelScore, nameScore) {
  getHubSpotLookupKeys_(rawValue).forEach(key => {
    if (!variants.some(variant => variant.key === key)) {
      variants.push({ key, labelScore, nameScore });
    }
  });
}

function dedupeHubSpotPropertyMatches_(matches) {
  const bestByKey = {};

  matches.forEach(match => {
    const dedupeKey = `${match.objectInfo.objectTypeId}::${match.property.name}`;
    if (!bestByKey[dedupeKey] || bestByKey[dedupeKey].score < match.score) {
      bestByKey[dedupeKey] = match;
    }
  });

  return Object.keys(bestByKey).map(key => bestByKey[key]);
}

function extractHubSpotObjectHint_(header, objectInfos) {
  const key = normalizeHubSpotLookupKey_(header);
  const aliasMatches = getHubSpotAliasMatches_(key, objectInfos);

  const startMatch = aliasMatches
    .filter(match => match.start === 0)
    .sort((a, b) => b.alias.length - a.alias.length)[0];

  if (startMatch) {
    return {
      objectInfo: startMatch.objectInfo,
      remainderOriginal: header.slice(startMatch.rawLength).replace(/^[\s:|/-]+/, "").trim(),
      remainderKey: key.slice(startMatch.alias.length).replace(/^[\s]+/, "").trim()
    };
  }

  const endMatch = aliasMatches
    .filter(match => match.end === key.length)
    .sort((a, b) => b.alias.length - a.alias.length)[0];

  if (endMatch) {
    const rawPrefix = header.slice(0, Math.max(0, header.length - endMatch.rawLength));
    return {
      objectInfo: endMatch.objectInfo,
      remainderOriginal: rawPrefix.replace(/[\s:|/-]+$/, "").trim(),
      remainderKey: key.slice(0, Math.max(0, key.length - endMatch.alias.length)).replace(/[\s]+$/, "").trim()
    };
  }

  return null;
}

function getHubSpotAliasMatches_(key, objectInfos) {
  const matches = [];

  objectInfos.forEach(objectInfo => {
    objectInfo.aliases.forEach(alias => {
      const pattern = new RegExp(`(^|\\s)${escapeRegExp_(alias).replace(/\\ /g, "\\s+")}(?=\\s|$)`);
      const match = pattern.exec(key);
      if (!match) return;

      const start = match.index + (match[1] ? match[1].length : 0);
      matches.push({
        objectInfo,
        alias,
        start,
        end: start + alias.length,
        rawLength: alias.length
      });
    });
  });

  return matches;
}

function shouldUsePrimaryDisplayProperty_(remainderKey) {
  return !remainderKey || remainderKey === "name" || /(^| )name$/.test(remainderKey);
}

function findHubSpotPropertyByName_(objectInfo, propertyName) {
  const lookupKeys = getHubSpotLookupKeys_(propertyName);

  for (let i = 0; i < lookupKeys.length; i++) {
    const matches = objectInfo.propertyLookup.name[lookupKeys[i]] || [];
    if (matches.length) return matches[0];
  }

  return null;
}

function addPropertyLookupEntry_(lookup, rawKey, property) {
  getHubSpotLookupKeys_(rawKey).forEach(key => {
    if (!lookup[key]) lookup[key] = [];
    lookup[key].push(property);
  });
}

function getHubSpotLookupKeys_(rawValue) {
  const normalized = normalizeHubSpotLookupKey_(rawValue);
  if (!normalized) return [];

  const compact = normalized.replace(/\s+/g, "");
  return compact && compact !== normalized ? [normalized, compact] : [normalized];
}

function getHubSpotDefaultObjectForHeader_(header, objectInfos) {
  const objectKey = HUBSPOT_IMPORT_HEADER_OBJECT_HINTS_[normalizeHubSpotLookupKey_(header)];
  if (!objectKey) return null;

  for (let i = 0; i < objectInfos.length; i++) {
    if (objectInfos[i].key === objectKey) return objectInfos[i];
  }

  return null;
}

function normalizeHubSpotLookupKey_(rawValue) {
  return String(rawValue || "")
    .replace(/([a-z])([A-Z])/g, "$1 $2")
    .toLowerCase()
    .replace(/&/g, " and ")
    .replace(/[_/]+/g, " ")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function findDuplicateHubSpotHeaders_(headers) {
  const seen = {};
  const duplicates = {};

  headers.forEach(header => {
    const key = normalizeHubSpotLookupKey_(header);
    if (!key) return;

    if (seen[key]) {
      duplicates[key] = true;
      return;
    }

    seen[key] = header;
  });

  return Object.keys(duplicates).map(key => seen[key]);
}

function getHubSpotSharedImportWebAppUrl_() {
  return String(HUBSPOT_SHARED_IMPORT_WEB_APP_URL_ || "").trim();
}

function hasHubSpotSharedImporterLibrary_() {
  return (
    typeof HubSpotSharedImporter !== "undefined" &&
    HubSpotSharedImporter &&
    typeof HubSpotSharedImporter.startImport === "function"
  );
}

function getHubSpotImportToken_() {
  const props = PropertiesService.getScriptProperties();

  for (let i = 0; i < HUBSPOT_IMPORT_TOKEN_PROPS_.length; i++) {
    const value = String(props.getProperty(HUBSPOT_IMPORT_TOKEN_PROPS_[i]) || "").trim();
    if (value) return value;
  }

  return "";
}

function startHubSpotImportViaLibrary_(sheetPayload) {
  const result = HubSpotSharedImporter.startImport({
    spreadsheetName: sheetPayload.spreadsheetName,
    spreadsheetLocale: sheetPayload.spreadsheetLocale,
    headers: sheetPayload.headers,
    rows: sheetPayload.rows
  });

  if (!result || result.ok !== true) {
    throw new Error(
      result && result.error
        ? result.error
        : "Shared HubSpot importer library failed."
    );
  }

  return result;
}

function startHubSpotImportViaWebApp_(webAppUrl, sheetPayload) {
  const response = UrlFetchApp.fetch(webAppUrl, {
    method: "post",
    muteHttpExceptions: true,
    followRedirects: false,
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()
    },
    payload: JSON.stringify({
      action: HUBSPOT_SHARED_IMPORT_ACTION_,
      spreadsheetName: sheetPayload.spreadsheetName,
      spreadsheetLocale: sheetPayload.spreadsheetLocale,
      headers: sheetPayload.headers,
      rows: sheetPayload.rows
    })
  });

  const code = response.getResponseCode();
  const text = String(response.getContentText() || "");

  if (code === 302 || code === 401 || code === 403) {
    throw new Error(
      "Shared HubSpot importer web app denied access. " +
      "Confirm the deployment URL is correct and the web app is deployed to users who should run this template."
    );
  }

  if (code >= 400) {
    throw new Error(`Shared HubSpot importer error ${code}: ${text.slice(0, 1000)}`);
  }

  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch (error) {
    throw new Error(
      "Shared HubSpot importer returned a non-JSON response. " +
      "This usually means the deployment URL is wrong or the web app access settings are too restrictive."
    );
  }

  if (!parsed || parsed.ok !== true) {
    throw new Error(parsed && parsed.error ? parsed.error : "Shared HubSpot importer failed.");
  }

  return parsed;
}

function startHubSpotImport_(token, fileBlob, importRequest) {
  const response = UrlFetchApp.fetch(`${HUBSPOT_API_BASE_}/crm/v3/imports`, {
    method: "post",
    muteHttpExceptions: true,
    headers: {
      Authorization: "Bearer " + token
    },
    payload: {
      importRequest: JSON.stringify(importRequest),
      files: fileBlob.setName(importRequest.files[0].fileName)
    }
  });

  const code = response.getResponseCode();
  const text = String(response.getContentText() || "");

  if (code >= 400) {
    throw new Error(`HubSpot API error ${code}: ${text.slice(0, 1000)}`);
  }

  try {
    return JSON.parse(text);
  } catch (error) {
    throw new Error("HubSpot import response could not be parsed: " + error);
  }
}

function hubspotFetchJson_(url, token, options) {
  const response = UrlFetchApp.fetch(
    url,
    Object.assign(
      {
        method: "get",
        muteHttpExceptions: true,
        headers: {
          Authorization: "Bearer " + token,
          "Content-Type": "application/json"
        }
      },
      options || {}
    )
  );

  const code = response.getResponseCode();
  const text = String(response.getContentText() || "");

  if (code >= 400) {
    throw new Error(`HubSpot API error ${code}: ${text.slice(0, 1000)}`);
  }

  try {
    return JSON.parse(text);
  } catch (error) {
    throw new Error("HubSpot API parse error: " + error);
  }
}

function getHubSpotImportDateFormat_(spreadsheetLocale) {
  const locale = String(spreadsheetLocale || "").replace("-", "_").toLowerCase();
  if (locale === "en_us") return "MONTH_DAY_YEAR";
  if (/^(ja|ko|zh)(_|$)/.test(locale)) return "YEAR_MONTH_DAY";
  return "DAY_MONTH_YEAR";
}

function columnNumberToLetter_(columnNumber) {
  let current = Number(columnNumber || 0);
  let out = "";

  while (current > 0) {
    const remainder = (current - 1) % 26;
    out = String.fromCharCode(65 + remainder) + out;
    current = Math.floor((current - 1) / 26);
  }

  return out || "?";
}

function escapeRegExp_(value) {
  return String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("HypeAuditor → HubSpot")
    .addItem("Fill CRM Import", "fillCrmImport")
    .addSeparator()
    .addItem("Import to HubSpot", "importToHubSpot")
    .addSeparator()
    .addItem("🔵 CM input", "showLegend")
    .addItem("🟡 Scouter input", "showLegend")
    .addItem("🔴 Filled automatically", "showLegend")
    .addItem("⚪ Optional", "showLegend")
    .addToUi();
}

function showLegend() {
  SpreadsheetApp.getUi().alert(
    "Legend:\n\n" +
    "🔵 CM input\n" +
    "🟡 Scouter input\n" +
    "🔴 Filled automatically\n" +
    "⚪ Optional"
  );
}
