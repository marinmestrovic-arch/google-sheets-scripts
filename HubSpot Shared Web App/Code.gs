var HUBSPOT_API_BASE_ = "https://api.hubapi.com";
var HUBSPOT_IMPORT_TOKEN_PROPS_ = [
  "HUBSPOT_ACCESS_TOKEN",
  "HUBSPOT_PRIVATE_APP_TOKEN",
  "HUBSPOT_API_KEY"
];
var HUBSPOT_WEB_APP_ACTION_START_IMPORT_ = "startImport";
var HUBSPOT_WEB_APP_ALLOWED_DOMAINS_PROP_ = "HUBSPOT_IMPORT_ALLOWED_DOMAINS";
var HUBSPOT_WEB_APP_ALLOWED_EMAILS_PROP_ = "HUBSPOT_IMPORT_ALLOWED_EMAILS";

var HUBSPOT_IMPORT_OBJECT_SPECS_ = [
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

var HUBSPOT_IMPORT_HEADER_OVERRIDES_ = {
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

var HUBSPOT_IMPORT_HEADER_OBJECT_HINTS_ = {
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

function doGet() {
  return jsonOutput_({
    ok: true,
    message: "HubSpot shared importer is deployed."
  });
}

function doPost(e) {
  try {
    enforceHubSpotWebAppAccess_();

    var payload = parseHubSpotImportWebAppRequest_(e);
    if (payload.action !== HUBSPOT_WEB_APP_ACTION_START_IMPORT_) {
      throw new Error("Unsupported action: " + payload.action);
    }

    var token = getHubSpotImportToken_();
    if (!token) {
      throw new Error(
        "Missing HubSpot token in the web app project. " +
        "Set one of these Script Properties: " + HUBSPOT_IMPORT_TOKEN_PROPS_.join(", ")
      );
    }

    var prepared = prepareHubSpotImportFromPayload_(payload, token);
    var startedImport = startHubSpotImport_(token, prepared.fileBlob, prepared.importRequest);

    return jsonOutput_({
      ok: true,
      rowCount: prepared.rowCount,
      columnCount: prepared.columnCount,
      objectLabels: prepared.objectLabels,
      importId: startedImport && startedImport.id != null ? String(startedImport.id) : "not returned",
      state: startedImport && startedImport.state ? startedImport.state : "STARTED"
    });
  } catch (error) {
    var message = error && error.message ? error.message : String(error);
    Logger.log("HubSpot shared importer failed: " + message);

    return jsonOutput_({
      ok: false,
      error: message
    });
  }
}

function enforceHubSpotWebAppAccess_() {
  var props = PropertiesService.getScriptProperties();
  var allowedDomains = parseCsvProperty_(props.getProperty(HUBSPOT_WEB_APP_ALLOWED_DOMAINS_PROP_));
  var allowedEmails = parseCsvProperty_(props.getProperty(HUBSPOT_WEB_APP_ALLOWED_EMAILS_PROP_));

  if (!allowedDomains.length && !allowedEmails.length) return;

  var email = String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();
  if (!email) {
    throw new Error(
      "Could not determine the active user email for access control. " +
      "Deploy the web app for your Workspace domain and make sure the caller is signed in."
    );
  }

  var emailAllowed = allowedEmails.indexOf(email) !== -1;
  var domain = email.split("@")[1] || "";
  var domainAllowed = allowedDomains.indexOf(domain) !== -1;

  if (!emailAllowed && !domainAllowed) {
    throw new Error("Access denied for " + email + ".");
  }
}

function parseHubSpotImportWebAppRequest_(e) {
  var raw = e && e.postData && e.postData.contents ? e.postData.contents : "";
  if (!raw) throw new Error("Request body is empty.");

  var parsed;
  try {
    parsed = JSON.parse(raw);
  } catch (error) {
    throw new Error("Request body is not valid JSON.");
  }

  var headers = Array.isArray(parsed.headers)
    ? parsed.headers.map(value => String(value || "").trim())
    : null;
  var rows = Array.isArray(parsed.rows)
    ? parsed.rows.map(row => Array.isArray(row) ? row.map(value => value == null ? "" : String(value)) : [])
    : null;

  if (!headers || !rows) {
    throw new Error("Request must include headers[] and rows[].");
  }

  return {
    action: String(parsed.action || "").trim(),
    spreadsheetName: sanitizeFileName_(parsed.spreadsheetName || "HubSpot Import"),
    spreadsheetLocale: String(parsed.spreadsheetLocale || "").trim(),
    headers,
    rows,
    rowCount: rows.length,
    columnCount: headers.length
  };
}

function prepareHubSpotImportFromPayload_(payload, token) {
  if (!payload.headers.length) {
    throw new Error("No columns provided for import.");
  }

  if (!payload.rows.length) {
    throw new Error("No rows provided for import.");
  }

  var objectCatalog = loadHubSpotImportObjectCatalog_(token);
  var resolvedColumns = resolveHubSpotImportColumns_(payload.headers, objectCatalog);

  if (resolvedColumns.errors.length) {
    throw new Error(
      "Some HubSpot Import columns could not be mapped automatically.\n\n" +
      resolvedColumns.errors.join("\n") +
      "\n\nUse the HubSpot property label or internal name in the sheet header. " +
      "For ambiguous columns, prefix the header with the object name, for example: " +
      '"Contacts Email", "Deals Record ID", or "Client Campaigns Name".'
    );
  }

  var fileName = "HubSpot Import - " + payload.spreadsheetName + ".csv";
  var fileBlob = Utilities.newBlob(
    toCsvBytes_([payload.headers].concat(payload.rows)),
    "text/csv",
    fileName
  );
  var importRequest = buildHubSpotImportRequest_(
    fileName,
    resolvedColumns.mappings,
    payload.spreadsheetLocale
  );

  return {
    fileBlob,
    importRequest,
    rowCount: payload.rowCount,
    columnCount: payload.columnCount,
    objectLabels: resolvedColumns.objectLabels
  };
}

function toCsvBytes_(grid) {
  var csv = grid
    .map(row =>
      row
        .map(cell => {
          if (cell === null || cell === undefined) return "";
          var s = String(cell);
          s = s.replace(/\r\n/g, "\n").replace(/\r/g, "\n");

          var needsQuotes = /[",\n]/.test(s);
          s = s.replace(/"/g, '""');
          return needsQuotes ? '"' + s + '"' : s;
        })
        .join(",")
    )
    .join("\r\n");

  return Utilities.newBlob("\uFEFF" + csv, "text/csv;charset=utf-8").getBytes();
}

function buildHubSpotImportRequest_(fileName, columnMappings, spreadsheetLocale) {
  return {
    name: "Google Sheets Import - " + fileName.replace(/\.csv$/i, ""),
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
  var operations = {};

  columnMappings.forEach(mapping => {
    if (!mapping || mapping.columnType === "FLEXIBLE_ASSOCIATION_LABEL") return;

    var objectTypeId = mapping.columnObjectTypeId;
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
  var schemasResponse = hubspotFetchJson_(
    HUBSPOT_API_BASE_ + "/crm-object-schemas/v3/schemas",
    token
  );
  var schemas = Array.isArray(schemasResponse && schemasResponse.results)
    ? schemasResponse.results
    : [];

  var list = HUBSPOT_IMPORT_OBJECT_SPECS_.map(spec => {
    var schema = resolveHubSpotObjectSchema_(spec, schemas);
    var propertiesResponse = hubspotFetchJson_(
      HUBSPOT_API_BASE_ + "/crm/v3/properties/" + encodeURIComponent(schema.objectTypeId),
      token
    );
    var properties = Array.isArray(propertiesResponse && propertiesResponse.results)
      ? propertiesResponse.results
      : [];

    if (!properties.length) {
      throw new Error("No HubSpot properties returned for " + spec.label + ".");
    }

    var propertyLookup = { label: {}, name: {} };
    properties.forEach(property => {
      addPropertyLookupEntry_(propertyLookup.label, property.label, property);
      addPropertyLookupEntry_(propertyLookup.name, property.name, property);
    });

    var aliasSet = {};
    spec.aliases.forEach(alias => {
      var key = normalizeHubSpotLookupKey_(alias);
      if (key) aliasSet[key] = true;
    });

    var singular = schema.labels && schema.labels.singular ? schema.labels.singular : "";
    var plural = schema.labels && schema.labels.plural ? schema.labels.plural : "";
    var schemaName = schema.name || "";

    [singular, plural, schemaName, spec.label].forEach(alias => {
      var key = normalizeHubSpotLookupKey_(alias);
      if (key) aliasSet[key] = true;
    });

    var primaryDisplayProperty = properties.find(
      property => property.name === schema.primaryDisplayProperty
    ) || null;

    return {
      key: spec.key,
      label: plural || singular || spec.label,
      objectTypeId: schema.objectTypeId,
      propertyLookup,
      aliases: Object.keys(aliasSet),
      primaryDisplayProperty
    };
  });

  var byKey = {};
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

  var wanted = spec.aliases.map(alias => normalizeHubSpotLookupKey_(alias));
  var matches = schemas.filter(schema => {
    var candidates = [
      schema && schema.name,
      schema && schema.labels && schema.labels.singular,
      schema && schema.labels && schema.labels.plural
    ]
      .map(value => normalizeHubSpotLookupKey_(value))
      .filter(Boolean);

    return wanted.some(alias => candidates.indexOf(alias) !== -1);
  });

  if (matches.length === 0) {
    throw new Error("Could not find the HubSpot custom object schema for " + spec.label + ".");
  }

  if (matches.length > 1) {
    throw new Error("Multiple HubSpot schemas matched " + spec.label + ".");
  }

  return matches[0];
}

function resolveHubSpotImportColumns_(headers, objectCatalog) {
  var mappings = [];
  var objectLabels = [];
  var seenObjectTypeIds = {};
  var errors = [];

  headers.forEach(header => {
    var associationMapping = resolveHubSpotAssociationColumn_(header, objectCatalog.list);
    if (associationMapping) {
      mappings.push(associationMapping);
      return;
    }

    var overrideMapping = resolveHubSpotHeaderOverride_(header, objectCatalog);
    if (overrideMapping) {
      mappings.push(overrideMapping.mapping);
      if (!seenObjectTypeIds[overrideMapping.objectInfo.objectTypeId]) {
        seenObjectTypeIds[overrideMapping.objectInfo.objectTypeId] = true;
        objectLabels.push(overrideMapping.objectInfo.label);
      }
      return;
    }

    var resolvedProperty = resolveHubSpotPropertyColumn_(header, objectCatalog.list);
    if (resolvedProperty.error) {
      errors.push("- " + header + ": " + resolvedProperty.error);
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
  var key = normalizeHubSpotLookupKey_(header);
  var override = HUBSPOT_IMPORT_HEADER_OVERRIDES_[key];
  if (!override) return null;

  var objectInfo = objectCatalog.byKey[override.objectKey];
  if (!objectInfo) return null;

  var property = findHubSpotPropertyByName_(objectInfo, override.propertyName);
  if (!property) return null;

  return {
    objectInfo,
    mapping: buildHubSpotPropertyMapping_(header, objectInfo, property, override.columnType)
  };
}

function resolveHubSpotPropertyColumn_(header, objectInfos) {
  var objectHint = extractHubSpotObjectHint_(header, objectInfos);
  var defaultObjectInfo = objectHint ? null : getHubSpotDefaultObjectForHeader_(header, objectInfos);
  var keyVariants = buildHubSpotHeaderKeyVariants_(header, objectHint);
  var matches = collectHubSpotPropertyMatches_(
    keyVariants,
    objectHint ? [objectHint.objectInfo] : (defaultObjectInfo ? [defaultObjectInfo] : objectInfos),
    objectHint ? 10 : (defaultObjectInfo ? 5 : 0)
  );

  if (!matches.length && defaultObjectInfo) {
    matches = collectHubSpotPropertyMatches_(keyVariants, objectInfos, 0);
  }

  var dedupedMatches = dedupeHubSpotPropertyMatches_(matches);
  if (!dedupedMatches.length) {
    var preferredObjectInfo = objectHint ? objectHint.objectInfo : defaultObjectInfo;
    var fallbackRemainderKey = objectHint
      ? objectHint.remainderKey
      : normalizeHubSpotLookupKey_(header);

    if (preferredObjectInfo && shouldUsePrimaryDisplayProperty_(fallbackRemainderKey)) {
      var primaryProperty = preferredObjectInfo.primaryDisplayProperty;
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

  var topMatch = dedupedMatches[0];
  var secondMatch = dedupedMatches[1];
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
  var key = normalizeHubSpotLookupKey_(header);
  if (key.indexOf("association label") === -1) return null;

  var aliasMatches = getHubSpotAliasMatches_(key, objectInfos)
    .sort((a, b) => {
      if (a.start !== b.start) return a.start - b.start;
      return b.alias.length - a.alias.length;
    });

  var chosen = [];
  aliasMatches.forEach(match => {
    var overlaps = chosen.some(
      existing => match.start < existing.end && match.end > existing.start
    );
    if (!overlaps && !chosen.some(existing => existing.objectInfo.key === match.objectInfo.key)) {
      chosen.push(match);
    }
  });

  if (chosen.length < 2) {
    throw new Error(
      'Association label column "' + header + '" must mention both objects, for example "Contacts to Deals Association Label".'
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
  var matches = [];

  objectInfos.forEach(objectInfo => {
    keyVariants.forEach(variant => {
      var labelMatches = objectInfo.propertyLookup.label[variant.key] || [];
      labelMatches.forEach(property => {
        matches.push({
          objectInfo,
          property,
          score: variant.labelScore + objectBonus
        });
      });

      var nameMatches = objectInfo.propertyLookup.name[variant.key] || [];
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
  var mapping = {
    columnName,
    columnObjectTypeId: objectInfo.objectTypeId,
    propertyName: property.name
  };

  var columnType = forcedColumnType || getHubSpotPropertyColumnType_(objectInfo, property);
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
  var variants = [];

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
  var bestByKey = {};

  matches.forEach(match => {
    var dedupeKey = match.objectInfo.objectTypeId + "::" + match.property.name;
    if (!bestByKey[dedupeKey] || bestByKey[dedupeKey].score < match.score) {
      bestByKey[dedupeKey] = match;
    }
  });

  return Object.keys(bestByKey).map(key => bestByKey[key]);
}

function extractHubSpotObjectHint_(header, objectInfos) {
  var key = normalizeHubSpotLookupKey_(header);
  var aliasMatches = getHubSpotAliasMatches_(key, objectInfos);

  var startMatch = aliasMatches
    .filter(match => match.start === 0)
    .sort((a, b) => b.alias.length - a.alias.length)[0];

  if (startMatch) {
    return {
      objectInfo: startMatch.objectInfo,
      remainderOriginal: header.slice(startMatch.rawLength).replace(/^[\s:|/-]+/, "").trim(),
      remainderKey: key.slice(startMatch.alias.length).replace(/^[\s]+/, "").trim()
    };
  }

  var endMatch = aliasMatches
    .filter(match => match.end === key.length)
    .sort((a, b) => b.alias.length - a.alias.length)[0];

  if (endMatch) {
    var rawPrefix = header.slice(0, Math.max(0, header.length - endMatch.rawLength));
    return {
      objectInfo: endMatch.objectInfo,
      remainderOriginal: rawPrefix.replace(/[\s:|/-]+$/, "").trim(),
      remainderKey: key.slice(0, Math.max(0, key.length - endMatch.alias.length)).replace(/[\s]+$/, "").trim()
    };
  }

  return null;
}

function getHubSpotAliasMatches_(key, objectInfos) {
  var matches = [];

  objectInfos.forEach(objectInfo => {
    objectInfo.aliases.forEach(alias => {
      var pattern = new RegExp("(^|\\s)" + escapeRegExp_(alias).replace(/\\ /g, "\\s+") + "(?=\\s|$)");
      var match = pattern.exec(key);
      if (!match) return;

      var start = match.index + (match[1] ? match[1].length : 0);
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
  var lookupKeys = getHubSpotLookupKeys_(propertyName);

  for (var i = 0; i < lookupKeys.length; i++) {
    var matches = objectInfo.propertyLookup.name[lookupKeys[i]] || [];
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
  var normalized = normalizeHubSpotLookupKey_(rawValue);
  if (!normalized) return [];

  var compact = normalized.replace(/\s+/g, "");
  return compact && compact !== normalized ? [normalized, compact] : [normalized];
}

function getHubSpotDefaultObjectForHeader_(header, objectInfos) {
  var objectKey = HUBSPOT_IMPORT_HEADER_OBJECT_HINTS_[normalizeHubSpotLookupKey_(header)];
  if (!objectKey) return null;

  for (var i = 0; i < objectInfos.length; i++) {
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

function getHubSpotImportToken_() {
  var props = PropertiesService.getScriptProperties();

  for (var i = 0; i < HUBSPOT_IMPORT_TOKEN_PROPS_.length; i++) {
    var value = String(props.getProperty(HUBSPOT_IMPORT_TOKEN_PROPS_[i]) || "").trim();
    if (value) return value;
  }

  return "";
}

function startHubSpotImport_(token, fileBlob, importRequest) {
  var response = UrlFetchApp.fetch(HUBSPOT_API_BASE_ + "/crm/v3/imports", {
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

  var code = response.getResponseCode();
  var text = String(response.getContentText() || "");

  if (code >= 400) {
    throw new Error("HubSpot API error " + code + ": " + text.slice(0, 1000));
  }

  try {
    return JSON.parse(text);
  } catch (error) {
    throw new Error("HubSpot import response could not be parsed: " + error);
  }
}

function hubspotFetchJson_(url, token, options) {
  var response = UrlFetchApp.fetch(
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

  var code = response.getResponseCode();
  var text = String(response.getContentText() || "");

  if (code >= 400) {
    throw new Error("HubSpot API error " + code + ": " + text.slice(0, 1000));
  }

  try {
    return JSON.parse(text);
  } catch (error) {
    throw new Error("HubSpot API parse error: " + error);
  }
}

function getHubSpotImportDateFormat_(spreadsheetLocale) {
  var locale = String(spreadsheetLocale || "").replace("-", "_").toLowerCase();
  if (locale === "en_us") return "MONTH_DAY_YEAR";
  if (/^(ja|ko|zh)(_|$)/.test(locale)) return "YEAR_MONTH_DAY";
  return "DAY_MONTH_YEAR";
}

function sanitizeFileName_(value) {
  return String(value || "HubSpot Import")
    .replace(/[\\/:*?"<>|]+/g, "-")
    .trim() || "HubSpot Import";
}

function parseCsvProperty_(value) {
  return String(value || "")
    .split(",")
    .map(item => item.trim().toLowerCase())
    .filter(Boolean);
}

function jsonOutput_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function escapeRegExp_(value) {
  return String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
