var HUBSPOT_API_BASE_ = "https://api.hubapi.com";
var HUBSPOT_IMPORT_TOKEN_PROPS_ = [
  "HUBSPOT_ACCESS_TOKEN",
  "HUBSPOT_PRIVATE_APP_TOKEN",
  "HUBSPOT_API_KEY"
];

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

function startImport(payload) {
  try {
    var token = getHubSpotImportToken_();
    if (!token) {
      throw new Error(
        "Missing HubSpot token in the shared library project. " +
        "Set one of these Script Properties: " + HUBSPOT_IMPORT_TOKEN_PROPS_.join(", ")
      );
    }

    var normalizedPayload = parseHubSpotImportPayload_(payload);
    var prepared = prepareHubSpotImportFromPayload_(normalizedPayload, token);
    var startedImport = startHubSpotImport_(token, prepared.fileBlob, prepared.importRequest);

    return {
      ok: true,
      rowCount: prepared.rowCount,
      columnCount: prepared.columnCount,
      objectLabels: prepared.objectLabels,
      importId: startedImport && startedImport.id != null ? String(startedImport.id) : "not returned",
      state: startedImport && startedImport.state ? startedImport.state : "STARTED"
    };
  } catch (error) {
    return {
      ok: false,
      error: error && error.message ? error.message : String(error)
    };
  }
}

function validateImport(payload) {
  try {
    var token = getHubSpotImportToken_();
    if (!token) {
      throw new Error(
        "Missing HubSpot token in the shared library project. " +
        "Set one of these Script Properties: " + HUBSPOT_IMPORT_TOKEN_PROPS_.join(", ")
      );
    }

    var normalizedPayload = parseHubSpotImportPayload_(payload);
    var prepared = prepareHubSpotImportFromPayload_(normalizedPayload, token, {
      skipFileBuild: true
    });

    return {
      ok: true,
      rowCount: prepared.rowCount,
      columnCount: prepared.columnCount,
      objectLabels: prepared.objectLabels
    };
  } catch (error) {
    return {
      ok: false,
      error: error && error.message ? error.message : String(error)
    };
  }
}

function parseHubSpotImportPayload_(payload) {
  if (!payload || typeof payload !== "object") {
    throw new Error("Import payload is missing.");
  }

  var headers = Array.isArray(payload.headers)
    ? payload.headers.map(function(value) { return String(value || "").trim(); })
    : null;
  var rows = Array.isArray(payload.rows)
    ? payload.rows.map(function(row) {
        if (!Array.isArray(row)) return [];
        return row.map(function(value) {
          return value == null ? "" : String(value);
        });
      })
    : null;

  if (!headers || !rows) {
    throw new Error("Import payload must include headers[] and rows[].");
  }

  return {
    sheetName: String(payload.sheetName || "HubSpot Import").trim() || "HubSpot Import",
    spreadsheetName: sanitizeFileName_(payload.spreadsheetName || "HubSpot Import"),
    spreadsheetLocale: String(payload.spreadsheetLocale || "").trim(),
    headers: headers,
    rows: rows,
    sourceRowNumbers: normalizeHubSpotImportSourceRowNumbers_(payload.sourceRowNumbers, rows.length),
    rowCount: rows.length,
    columnCount: headers.length
  };
}

function prepareHubSpotImportFromPayload_(payload, token, options) {
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

  validateHubSpotImportEmails_(
    payload.headers,
    payload.rows,
    resolvedColumns.mappings,
    objectCatalog.byKey.contacts,
    payload.sourceRowNumbers,
    payload.sheetName
  );

  if (options && options.skipFileBuild) {
    return {
      fileBlob: null,
      importRequest: null,
      rowCount: payload.rowCount,
      columnCount: payload.columnCount,
      objectLabels: resolvedColumns.objectLabels
    };
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
    fileBlob: fileBlob,
    importRequest: importRequest,
    rowCount: payload.rowCount,
    columnCount: payload.columnCount,
    objectLabels: resolvedColumns.objectLabels
  };
}

function normalizeHubSpotImportSourceRowNumbers_(rawRowNumbers, rowCount) {
  var normalized = [];

  for (var i = 0; i < rowCount; i++) {
    var raw = Array.isArray(rawRowNumbers) ? Number(rawRowNumbers[i]) : NaN;
    normalized.push(isFinite(raw) && raw >= 2 ? Math.floor(raw) : i + 2);
  }

  return normalized;
}

function validateHubSpotImportEmails_(
  headers,
  rows,
  columnMappings,
  contactsObjectInfo,
  sourceRowNumbers,
  sheetName
) {
  var issues = collectHubSpotImportEmailIssues_(
    headers,
    rows,
    columnMappings,
    contactsObjectInfo,
    sourceRowNumbers
  );

  if (!issues.length) return;

  var issueLimit = 15;
  var lines = issues.slice(0, issueLimit).map(function(issue) {
    return (
      '- Row ' + issue.rowNumber +
      ', column "' + issue.columnName +
      '": ' + formatHubSpotImportValueForDisplay_(issue.value) +
      ' ' + issue.reason
    );
  });

  if (issues.length > issueLimit) {
    lines.push("- " + (issues.length - issueLimit) + " more invalid email value(s) not shown.");
  }

  throw new Error(
    "HubSpot import stopped because some contact email values will be rejected by HubSpot.\n\n" +
    'Fix these cells in the "' + (sheetName || "HubSpot Import") + '" sheet and run the import again:\n' +
    lines.join("\n") +
    "\n\nHubSpot only accepts typical email addresses such as name@domain.com."
  );
}

function collectHubSpotImportEmailIssues_(
  headers,
  rows,
  columnMappings,
  contactsObjectInfo,
  sourceRowNumbers
) {
  var emailColumns = getHubSpotImportEmailColumns_(headers, columnMappings, contactsObjectInfo);
  if (!emailColumns.length) return [];

  var issues = [];
  rows.forEach(function(row, rowIdx) {
    emailColumns.forEach(function(column) {
      var rawValue = row && row[column.index];
      var value = String(rawValue == null ? "" : rawValue).trim();
      if (!value) return;

      var reason = getHubSpotImportEmailValidationIssue_(value);
      if (!reason) return;

      issues.push({
        rowNumber: getHubSpotImportSourceRowNumber_(sourceRowNumbers, rowIdx),
        columnName: column.name,
        value: value,
        reason: reason
      });
    });
  });

  return issues;
}

function getHubSpotImportEmailColumns_(headers, columnMappings, contactsObjectInfo) {
  var contactObjectTypeId = contactsObjectInfo && contactsObjectInfo.objectTypeId
    ? contactsObjectInfo.objectTypeId
    : "0-1";

  return columnMappings.reduce(function(columns, mapping, idx) {
    if (
      mapping &&
      mapping.columnObjectTypeId === contactObjectTypeId &&
      mapping.propertyName === "email"
    ) {
      columns.push({
        index: idx,
        name: headers[idx] || "Email"
      });
    }

    return columns;
  }, []);
}

function getHubSpotImportEmailValidationIssue_(value) {
  if (!value) return "is blank.";
  if (value.length > 254) return "is longer than 254 characters.";
  if (/[^\x00-\x7F]/.test(value)) return "contains non-ASCII characters. Replace it with a plain email address.";
  if (/\s/.test(value)) return "contains whitespace.";

  var atCount = (value.match(/@/g) || []).length;
  if (atCount !== 1) return "must contain exactly one @ symbol.";

  var parts = value.split("@");
  var local = parts[0];
  var domain = parts[1];

  if (!local) return "is missing the part before @.";
  if (!domain) return "is missing the domain after @.";
  if (local.length > 64) return "has more than 64 characters before @.";
  if (local.charAt(0) === "." || local.charAt(local.length - 1) === ".") {
    return "has a dot at the start or end before @.";
  }
  if (local.indexOf("..") !== -1) return "has consecutive dots before @.";
  if (!/^[A-Za-z0-9!#$%&'*+/=?^_`{|}~.-]+$/.test(local)) {
    return "contains unsupported characters before @.";
  }

  if (domain.length > 253) return "has a domain longer than 253 characters.";
  if (domain.indexOf(".") === -1) return "must include a full domain such as domain.com.";
  if (domain.charAt(0) === "." || domain.charAt(domain.length - 1) === ".") {
    return "has a domain that starts or ends with a dot.";
  }
  if (domain.indexOf("..") !== -1) return "has consecutive dots in the domain.";

  var labels = domain.split(".");
  if (labels.some(function(label) { return !label; })) return "has an empty section in the domain.";

  for (var i = 0; i < labels.length; i++) {
    var label = labels[i];
    if (label.length > 63) return "has a domain section longer than 63 characters.";
    if (!/^[A-Za-z0-9-]+$/.test(label)) return "contains unsupported characters in the domain.";
    if (label.charAt(0) === "-" || label.charAt(label.length - 1) === "-") {
      return "has a domain section that starts or ends with a hyphen.";
    }
  }

  var topLevelDomain = labels[labels.length - 1];
  if (!/^(xn--[A-Za-z0-9-]{2,59}|[A-Za-z]{2,63})$/.test(topLevelDomain)) {
    return "has an invalid top-level domain.";
  }

  return "";
}

function getHubSpotImportSourceRowNumber_(sourceRowNumbers, rowIdx) {
  if (Array.isArray(sourceRowNumbers)) {
    var rowNumber = Number(sourceRowNumbers[rowIdx]);
    if (isFinite(rowNumber) && rowNumber >= 2) {
      return Math.floor(rowNumber);
    }
  }

  return rowIdx + 2;
}

function formatHubSpotImportValueForDisplay_(value) {
  var singleLine = String(value == null ? "" : value).replace(/\s+/g, " ").trim();
  var shortened = singleLine.length > 120 ? singleLine.slice(0, 117) + "..." : singleLine;
  return JSON.stringify(shortened);
}

function toCsvBytes_(grid) {
  var csv = grid
    .map(function(row) {
      return row
        .map(function(cell) {
          if (cell === null || cell === undefined) return "";
          var s = String(cell);
          s = s.replace(/\r\n/g, "\n").replace(/\r/g, "\n");

          var needsQuotes = /[",\n]/.test(s);
          s = s.replace(/"/g, '""');
          return needsQuotes ? '"' + s + '"' : s;
        })
        .join(",");
    })
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
        fileName: fileName,
        fileFormat: "CSV",
        fileImportPage: {
          hasHeader: true,
          columnMappings: columnMappings
        }
      }
    ]
  };
}

function buildHubSpotImportOperations_(columnMappings) {
  var operations = {};

  columnMappings.forEach(function(mapping) {
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

  var list = HUBSPOT_IMPORT_OBJECT_SPECS_.map(function(spec) {
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
    properties.forEach(function(property) {
      addPropertyLookupEntry_(propertyLookup.label, property.label, property);
      addPropertyLookupEntry_(propertyLookup.name, property.name, property);
    });

    var aliasSet = {};
    spec.aliases.forEach(function(alias) {
      var key = normalizeHubSpotLookupKey_(alias);
      if (key) aliasSet[key] = true;
    });

    var singular = schema.labels && schema.labels.singular ? schema.labels.singular : "";
    var plural = schema.labels && schema.labels.plural ? schema.labels.plural : "";
    var schemaName = schema.name || "";

    [singular, plural, schemaName, spec.label].forEach(function(alias) {
      var key = normalizeHubSpotLookupKey_(alias);
      if (key) aliasSet[key] = true;
    });

    var primaryDisplayProperty = null;
    for (var i = 0; i < properties.length; i++) {
      if (properties[i].name === schema.primaryDisplayProperty) {
        primaryDisplayProperty = properties[i];
        break;
      }
    }

    return {
      key: spec.key,
      label: plural || singular || spec.label,
      objectTypeId: schema.objectTypeId,
      propertyLookup: propertyLookup,
      aliases: Object.keys(aliasSet),
      primaryDisplayProperty: primaryDisplayProperty
    };
  });

  var byKey = {};
  list.forEach(function(objectInfo) {
    byKey[objectInfo.key] = objectInfo;
  });

  return { list: list, byKey: byKey };
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

  var wanted = spec.aliases.map(function(alias) {
    return normalizeHubSpotLookupKey_(alias);
  });
  var matches = schemas.filter(function(schema) {
    var candidates = [
      schema && schema.name,
      schema && schema.labels && schema.labels.singular,
      schema && schema.labels && schema.labels.plural
    ]
      .map(function(value) { return normalizeHubSpotLookupKey_(value); })
      .filter(Boolean);

    return wanted.some(function(alias) { return candidates.indexOf(alias) !== -1; });
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

  headers.forEach(function(header) {
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

  return { mappings: mappings, objectLabels: objectLabels, errors: errors };
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
    objectInfo: objectInfo,
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

  dedupedMatches.sort(function(a, b) { return b.score - a.score; });

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
    .sort(function(a, b) {
      if (a.start !== b.start) return a.start - b.start;
      return b.alias.length - a.alias.length;
    });

  var chosen = [];
  aliasMatches.forEach(function(match) {
    var overlaps = chosen.some(function(existing) {
      return match.start < existing.end && match.end > existing.start;
    });
    if (!overlaps && !chosen.some(function(existing) { return existing.objectInfo.key === match.objectInfo.key; })) {
      chosen.push(match);
    }
  });

  if (chosen.length < 2) {
    throw new Error(
      'Association label column "' + header + '" must mention both objects, for example "Contacts to Deals Association Label".'
    );
  }

  chosen.sort(function(a, b) { return a.start - b.start; });

  return {
    columnName: header,
    columnType: "FLEXIBLE_ASSOCIATION_LABEL",
    columnObjectTypeId: chosen[0].objectInfo.objectTypeId,
    toColumnObjectTypeId: chosen[1].objectInfo.objectTypeId
  };
}

function collectHubSpotPropertyMatches_(keyVariants, objectInfos, objectBonus) {
  var matches = [];

  objectInfos.forEach(function(objectInfo) {
    keyVariants.forEach(function(variant) {
      var labelMatches = objectInfo.propertyLookup.label[variant.key] || [];
      labelMatches.forEach(function(property) {
        matches.push({
          objectInfo: objectInfo,
          property: property,
          score: variant.labelScore + objectBonus
        });
      });

      var nameMatches = objectInfo.propertyLookup.name[variant.key] || [];
      nameMatches.forEach(function(property) {
        matches.push({
          objectInfo: objectInfo,
          property: property,
          score: variant.nameScore + objectBonus
        });
      });
    });
  });

  return matches;
}

function buildHubSpotPropertyMapping_(columnName, objectInfo, property, forcedColumnType) {
  var mapping = {
    columnName: columnName,
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
  getHubSpotLookupKeys_(rawValue).forEach(function(key) {
    if (!variants.some(function(variant) { return variant.key === key; })) {
      variants.push({ key: key, labelScore: labelScore, nameScore: nameScore });
    }
  });
}

function dedupeHubSpotPropertyMatches_(matches) {
  var bestByKey = {};

  matches.forEach(function(match) {
    var dedupeKey = match.objectInfo.objectTypeId + "::" + match.property.name;
    if (!bestByKey[dedupeKey] || bestByKey[dedupeKey].score < match.score) {
      bestByKey[dedupeKey] = match;
    }
  });

  return Object.keys(bestByKey).map(function(key) {
    return bestByKey[key];
  });
}

function extractHubSpotObjectHint_(header, objectInfos) {
  var key = normalizeHubSpotLookupKey_(header);
  var aliasMatches = getHubSpotAliasMatches_(key, objectInfos);

  var startMatch = aliasMatches
    .filter(function(match) { return match.start === 0; })
    .sort(function(a, b) { return b.alias.length - a.alias.length; })[0];

  if (startMatch) {
    return {
      objectInfo: startMatch.objectInfo,
      remainderOriginal: header.slice(startMatch.rawLength).replace(/^[\s:|/-]+/, "").trim(),
      remainderKey: key.slice(startMatch.alias.length).replace(/^[\s]+/, "").trim()
    };
  }

  var endMatch = aliasMatches
    .filter(function(match) { return match.end === key.length; })
    .sort(function(a, b) { return b.alias.length - a.alias.length; })[0];

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

  objectInfos.forEach(function(objectInfo) {
    objectInfo.aliases.forEach(function(alias) {
      var pattern = new RegExp("(^|\\s)" + escapeRegExp_(alias).replace(/\\ /g, "\\s+") + "(?=\\s|$)");
      var match = pattern.exec(key);
      if (!match) return;

      var start = match.index + (match[1] ? match[1].length : 0);
      matches.push({
        objectInfo: objectInfo,
        alias: alias,
        start: start,
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
  getHubSpotLookupKeys_(rawKey).forEach(function(key) {
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

function escapeRegExp_(value) {
  return String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
