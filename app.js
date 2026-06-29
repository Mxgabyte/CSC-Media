function normalizeHex(value, fallback) {
  let hex = String(value || "").trim();

  if (!hex) return fallback;

  if (!hex.startsWith("#")) hex = "#" + hex;

  return /^#[0-9a-fA-F]{6}$/.test(hex) ? hex.toUpperCase() : fallback;
}

function getSafeErrorMessage(err, fallback = "Unknown error") {
  if (err instanceof Error && err.message) return err.message;

  if (typeof err === "string" && err.trim()) return err.trim();

  try {
    const text = JSON.stringify(err);
    return text && text !== "undefined" ? text : fallback;
  } catch {
    return fallback;
  }
}

function isFileProtocol() {
  return window.location && window.location.protocol === "file:";
}

async function fetchOpenSheetDataOnly(url, label) {
  const response = await fetch(url, {
    cache: "no-store"
  });

  if (!response.ok) {
    throw new Error(`${label} failed to load through OpenSheet. Status: ${response.status}`);
  }

  const data = await response.json();

  if (!Array.isArray(data)) {
    throw new Error(`${label} did not return rows.`);
  }

  return data;
}

function showBlankMatchupsMessage() {
  const gallery = document.getElementById("matchGallery");

  gallery.innerHTML = `
    <div class="blank-sheet-card">
      <div class="blank-sheet-title">Document Currently Blank</div>
      <div class="blank-sheet-text">Add matchups to the sheet, then hit Update Graphics.</div>
    </div>
  `;
}
async function loadPlayerProfileNames() {
  try {
    const tabsToTry = [...new Set([PLAYER_LIST_TAB, ...PLAYER_LIST_FALLBACK_TABS])];

    let bestEntries = [];
    let loadedTab = "";
    let loadedRows = 0;

    for (const tabName of tabsToTry) {
      try {
        const url = openSheetURL(PLAYER_LIST_SHEET_ID, tabName);
        const rows = await fetchSheetData(url, `Player Profile Names (${tabName})`);

        if (!Array.isArray(rows) || !rows.length) continue;

        const entries = rows
          .flatMap(row => getPlayerProfileEntriesFromRow_(row))
          .filter(entry => entry && entry.name);

        console.log(`PLAYER PROFILE TAB CHECK ${tabName}:`, {
          rows: rows.length,
          entries: entries.length,
          sample: entries.slice(0, 5).map(entry => entry.name)
        });

        // Do not stop just because the tab has rows. Stop when it has usable profile names.
        if (entries.length > bestEntries.length) {
          bestEntries = entries;
          loadedTab = tabName;
          loadedRows = rows.length;
        }

        if (entries.length) break;
      } catch (tabErr) {
        console.warn(`Player names tab failed: ${tabName}`, tabErr);
      }
    }

    if (!bestEntries.length) {
      throw new Error("No usable player names loaded from the hyperlink/player list sheet.");
    }

    const entryMap = {};

    bestEntries.forEach(entry => {
      addPlayerProfileEntry_(entryMap, entry.name, entry.url);

      // If the sheet accidentally includes the marker dash for normal names, support both.
      // Do not strip names like -Cram- because the dash is part of the actual name.
      if (entry.name.endsWith("-") && entry.name.indexOf("-") === entry.name.length - 1) {
        addPlayerProfileEntry_(entryMap, entry.name.slice(0, -1), entry.url);
      }
    });

    playerProfileEntries = Object.values(entryMap)
      .sort((a, b) => b.name.length - a.name.length);

    playerProfileNames = playerProfileEntries.map(entry => entry.name);
    playerProfileUrlMap = {};
    playerProfileEntries.forEach(entry => {
      playerProfileUrlMap[normalizeProfileNameKey_(entry.name)] = entry.url || profileUrlForName_(entry.name);
    });

    console.log(`PLAYER PROFILE NAMES LOADED from ${loadedTab}:`, {
      rows: loadedRows,
      names: playerProfileNames.length,
      sample: playerProfileNames.slice(0, 12)
    });
  } catch (err) {
    console.warn("Player profile names failed to load. Article links will not be created.", err);
    playerProfileNames = [];
    playerProfileEntries = [];
    playerProfileUrlMap = {};
  }
}

function normalizeProfileNameKey_(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, " ");
}

function cleanProfileNameCandidate_(value) {
  let text = String(value || "")
    .replace(/\s+/g, " ")
    .trim();

  if (!text) return "";

  // Ignore obvious headers, records, team names, or URLs.
  const normalized = normalizeHeaderName(text);
  const blocked = [
    "name", "player", "players", "username", "handle", "alias", "aliases",
    "url", "link", "profile", "hyperlink", "record", "team", "teams", "franchise"
  ];

  if (blocked.includes(normalized)) return "";
  if (/^https?:\/\//i.test(text)) return "";
  if (/^www\./i.test(text)) return "";
  if (/^\d+[-–]\d+$/.test(text)) return "";
  if (text.length > 64) return "";

  return text;
}

function cleanProfileUrlCandidate_(value) {
  const text = String(value || "").trim();

  if (!text) return "";
  if (/^https?:\/\//i.test(text)) return text;

  return "";
}

function extractProfileNameFromUrl_(url) {
  const text = String(url || "").trim();
  if (!text) return "";

  const match = text.match(/(?:stats\/profile\/|profile\/)([^/?#]+)/i);
  if (!match) return "";

  try {
    return decodeURIComponent(match[1]).trim();
  } catch {
    return match[1].trim();
  }
}

function profileUrlForName_(profileName) {
  return PROFILE_BASE_URL + encodeURIComponent(profileName);
}

function addPlayerProfileEntry_(entryMap, name, url = "") {
  const cleanName = cleanProfileNameCandidate_(name);
  if (!cleanName) return;

  const key = normalizeProfileNameKey_(cleanName);
  if (!key) return;

  const cleanUrl = cleanProfileUrlCandidate_(url) || profileUrlForName_(cleanName);

  if (!entryMap[key]) {
    entryMap[key] = {
      name: cleanName,
      url: cleanUrl
    };
    return;
  }

  // Prefer an actual sheet URL over the generated profile URL when available.
  if (cleanUrl && cleanUrl !== profileUrlForName_(cleanName)) {
    entryMap[key].url = cleanUrl;
  }
}

function getLooseRowValue_(row, possibleHeaders) {
  if (!row) return "";

  const actualKey = Object.keys(row).find(key =>
    possibleHeaders.some(header => normalizeHeaderName(key) === normalizeHeaderName(header))
  );

  return actualKey ? row[actualKey] : "";
}

function getPlayerProfileEntriesFromRow_(row) {
  if (!row) return [];

  const values = Object.values(row).map(value => String(value || "").trim());
  const entries = [];

  const explicitName = cleanProfileNameCandidate_(getLooseRowValue_(row, PLAYER_LIST_NAME_HEADERS));
  const explicitUrl = cleanProfileUrlCandidate_(getLooseRowValue_(row, PLAYER_LIST_URL_HEADERS));

  if (explicitName) {
    entries.push({ name: explicitName, url: explicitUrl });
  }

  // Column B fallback, because the original hyperlink sheet uses B for names.
  const columnBName = cleanProfileNameCandidate_(values[PLAYER_LIST_COLUMN_INDEX]);
  if (columnBName) {
    entries.push({ name: columnBName, url: explicitUrl });
  }

  // If a row has a CSC profile URL but no clean name column, use the URL slug.
  const urls = values.map(cleanProfileUrlCandidate_).filter(Boolean);
  urls.forEach(url => {
    const urlName = cleanProfileNameCandidate_(extractProfileNameFromUrl_(url));
    if (urlName) entries.push({ name: urlName, url });
  });

  // Last-resort fallback: collect short non-url cells as possible player names.
  // This handles OpenSheet when blank columns collapse and column B becomes the first value.
  values.forEach(value => {
    const possibleName = cleanProfileNameCandidate_(value);
    if (possibleName) entries.push({ name: possibleName, url: explicitUrl });
  });

  const map = {};
  entries.forEach(entry => addPlayerProfileEntry_(map, entry.name, entry.url));
  return Object.values(map);
}

function getProfileUrlForName(profileName) {
  const key = normalizeProfileNameKey_(profileName);
  return playerProfileUrlMap[key] || profileUrlForName_(profileName);
}

function escapeRegex(text) {
  return String(text).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function findProfileNameMatches(text) {
  const raw = String(text || "");
  const matches = [];
  const usedRanges = [];

  if (!raw || !playerProfileNames.length) return matches;

  playerProfileNames.forEach(profileName => {
    const safeName = escapeRegex(profileName);

    // Match exactly: Name- where the final dash is the marker.
    // This avoids linking normal words and still supports names with dashes, e.g. -Cram--.
    const regex = new RegExp(
      `(^|[^A-Za-z0-9_])(${safeName})-(?=$|[^A-Za-z0-9_])`,
      "gi"
    );

    let match;

    while ((match = regex.exec(raw)) !== null) {
      const prefixLength = match[1].length;

      const start = match.index + prefixLength;
      const end = start + match[2].length;
      const markerEnd = end + 1;

      const overlaps = usedRanges.some(range => start < range.end && markerEnd > range.start);
      if (overlaps) continue;

      matches.push({
        start,
        end,
        markerEnd,
        profileName
      });

      usedRanges.push({ start, end: markerEnd });
    }
  });

  return matches.sort((a, b) => a.start - b.start);
}

function profileLinkedHTML(text) {
  const raw = String(text || "");

  if (!raw || !playerProfileNames.length) {
    return escapeHTML(raw);
  }

  const matches = findProfileNameMatches(raw);

  if (!matches.length) {
    return escapeHTML(raw);
  }

  let html = "";
  let cursor = 0;

  matches.forEach(match => {
    html += escapeHTML(raw.slice(cursor, match.start));

    const visibleText = raw.slice(match.start, match.end);
    const url = getProfileUrlForName(match.profileName);

    html += `<a href="${escapeHTML(url)}" target="_blank" rel="noopener noreferrer">${escapeHTML(visibleText)}</a>`;

    // Skip the trailing marker dash so it does not show.
    cursor = match.markerEnd;
  });

  html += escapeHTML(raw.slice(cursor));

  return html;
}

function profileCleanedPlainText(text) {
  const raw = String(text || "");

  if (!raw || !playerProfileNames.length) {
    return raw;
  }

  const matches = findProfileNameMatches(raw);

  if (!matches.length) {
    return raw;
  }

  let cleaned = "";
  let cursor = 0;

  matches.forEach(match => {
    cleaned += raw.slice(cursor, match.end);
    cursor = match.markerEnd;
  });

  cleaned += raw.slice(cursor);

  return cleaned;
}

function profileCleanedTextBlocks(blocks) {
  return (blocks || []).map(block => profileCleanedPlainText(block));
}
function parseTeamColors(row) {
  const raw = String(
    row.Color ||
    row.color ||
    row.Colors ||
    row.colors ||
    ""
  ).trim();

  let colors = raw
    .split(",")
    .map(c => c.trim())
    .filter(Boolean)
    .map(c => normalizeHex(c, ""))
    .filter(Boolean);

  const primary = colors[0] || defaultStyle.primary;
  const secondary = colors[1] || primary;
  const accent = colors[2] || secondary;
  const accent2 = colors[3] || accent;
  const detail = colors[4] || accent2;
  const dark = "#020617";

  return { primary, secondary, accent, accent2, detail, dark };
}
function extractGoogleSheetId(value) {
  const text = String(value || "").trim();

  const match = text.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);

  if (match) return match[1];

  if (/^[a-zA-Z0-9-_]{20,}$/.test(text)) return text;

  return "";
}

function getRowValue(row, possibleHeaders) {
  const actualKey = Object.keys(row).find(key =>
    possibleHeaders.some(header =>
      String(key || "").trim().toLowerCase() === String(header || "").trim().toLowerCase()
    )
  );

  return actualKey ? row[actualKey] : "";
}

function getRowValueLoose(row, possibleHeaders, fallbackIndex = -1) {
  if (!row) return "";

  const wanted = possibleHeaders.map(normalizeHeaderName);
  const actualKey = Object.keys(row).find(key => wanted.includes(normalizeHeaderName(key)));

  if (actualKey && String(row[actualKey] || "").trim()) {
    return row[actualKey];
  }

  if (fallbackIndex >= 0) {
    const values = Object.values(row);
    return values[fallbackIndex] || "";
  }

  return "";
}

function normalizeTeamKey(value) {
  return String(value || "").trim().toUpperCase();
}

function normalizeFranchiseLookupKey(value) {
  return String(value || "")
    .trim()
    .replace(/\s+/g, " ")
    .toUpperCase();
}

function getFranchiseAliasMap() {
  if (typeof FRANCHISE_CODE_ALIASES === "undefined" || !FRANCHISE_CODE_ALIASES) {
    return {};
  }

  return FRANCHISE_CODE_ALIASES;
}

function resolveFranchiseCode(franchiseName, explicitCode = "") {
  const directCode = normalizeTeamKey(explicitCode);

  if (directCode) return directCode;

  const rawKey = normalizeFranchiseLookupKey(franchiseName);
  if (!rawKey) return "";

  const aliasMap = getFranchiseAliasMap();
  const matchedAliasKey = Object.keys(aliasMap).find(aliasName =>
    normalizeFranchiseLookupKey(aliasName) === rawKey
  );

  if (matchedAliasKey) {
    return normalizeTeamKey(aliasMap[matchedAliasKey]);
  }

  // If Column A is still an abbreviation, keep it as-is.
  if (/^[A-Z0-9]{2,6}$/.test(rawKey)) {
    return rawKey;
  }

  // Full names should not be treated as real logo/ranking codes. Returning
  // blank here prevents row-order fallback from attaching the wrong logo to
  // the wrong standings row.
  return "";
}

function strippedFranchiseText(value) {
  return normalizeFranchiseLookupKey(value).replace(/[^A-Z0-9]/g, "");
}

function franchiseInitialVariants(value) {
  const words = normalizeFranchiseLookupKey(value)
    .split(/[^A-Z0-9]+/)
    .filter(Boolean);

  if (!words.length) return [];

  const variants = new Set();
  variants.add(words.map(word => word[0]).join(""));
  variants.add(words.map(word => word.slice(0, 1)).join(""));

  if (words.length > 1) {
    variants.add(words[0] + words.slice(1).map(word => word[0]).join(""));
    variants.add(words[0].slice(0, 2) + words.slice(1).map(word => word[0]).join(""));
    variants.add(words.map(word => word.slice(0, 2)).join(""));
  }

  variants.add(words.join(""));

  return Array.from(variants).filter(Boolean);
}

function codeIsSubsequence(code, text) {
  let cursor = 0;

  for (const char of text) {
    if (char === code[cursor]) cursor++;
    if (cursor >= code.length) return true;
  }

  return false;
}

function guessFranchiseCode(franchiseName, knownCodes = []) {
  const cleanKnownCodes = [...new Set((knownCodes || []).map(normalizeTeamKey).filter(Boolean))];
  const stripped = strippedFranchiseText(franchiseName);
  const variants = franchiseInitialVariants(franchiseName).map(strippedFranchiseText);

  if (!stripped || !cleanKnownCodes.length) return "";

  const matches = cleanKnownCodes.filter(code => {
    const cleanCode = strippedFranchiseText(code);
    if (!cleanCode) return false;

    return (
      variants.includes(cleanCode) ||
      stripped === cleanCode ||
      stripped.startsWith(cleanCode) ||
      (cleanCode.length >= 3 && codeIsSubsequence(cleanCode, stripped))
    );
  });

  return matches.length === 1 ? matches[0] : "";
}

function getTierRecordTabs(tierName) {
  const configuredFallbacks = (TEAM_RECORDS_TAB_FALLBACKS && TEAM_RECORDS_TAB_FALLBACKS[tierName]) || [];
  return [...new Set([tierName, ...configuredFallbacks])].filter(Boolean);
}

function getTeamRecordFromRow(row) {
  const franchiseDisplay = String(getRowValueLoose(row, [
    "Franchise",
    "Franchise Name",
    "Team",
    "Teams"
  ], 0) || "").trim();

  const explicitCode = String(getRowValueLoose(row, [
    "Franchise Code",
    "Franchise Abbreviation",
    "Franchise Abbreviations",
    "Abbreviation",
    "Abbreviations",
    "Code",
    "Franchise ID",
    "ID"
  ], 3) || "").trim();

  const franchise = resolveFranchiseCode(franchiseDisplay, explicitCode);
  const franchiseRawKey = normalizeFranchiseLookupKey(franchiseDisplay);

  const teamName = String(getRowValueLoose(row, [
    "Team Name",
    "TeamName",
    "Name",
    "Roster Name"
  ], 1) || "").trim();

  const record = String(getRowValueLoose(row, [
    "Record",
    "Current Record",
    "W-L",
    "W/L",
    "Wins-Losses"
  ], 2) || "").trim();

  const lookupKeys = [...new Set([
    franchise,
    normalizeTeamKey(explicitCode),
    franchiseRawKey
  ].filter(Boolean))];

  return {
    franchise,
    franchiseDisplay,
    franchiseRawKey,
    teamName,
    record,
    lookupKeys
  };
}

function buildTeamRecordMap(teamRecordData = [], rankData = []) {
  const map = {};
  const knownCodes = [...new Set((rankData || [])
    .map(row => String(row.Teams || row.Team || "").trim().toUpperCase())
    .filter(Boolean))];

  (teamRecordData || []).forEach(row => {
    const recordModel = getTeamRecordFromRow(row);
    if (!recordModel.franchise && !recordModel.franchiseRawKey) return;

    const guessedCode = recordModel.franchise
      || guessFranchiseCode(recordModel.franchiseDisplay, knownCodes);

    const lookupKeys = [...new Set([
      guessedCode,
      ...recordModel.lookupKeys
    ].map(normalizeTeamKey).filter(Boolean))];

    lookupKeys.forEach(key => {
      // Never let a full-name/raw key overwrite a real abbreviation key.
      if (!map[key] || key === guessedCode) {
        map[key] = recordModel;
      }
    });
  });

  return map;
}

async function fetchTeamRecordData(tierName) {
  const recordsId = extractGoogleSheetId(TEAM_RECORDS_SHEET_URL);

  if (!recordsId) {
    console.warn("Team records sheet URL is invalid. Record boxes will be blank.");
    return [];
  }

  const tabsToTry = getTierRecordTabs(tierName);

  for (const tabName of tabsToTry) {
    try {
      const rows = await fetchSheetData(
        openSheetURL(recordsId, tabName),
        `${tierName} Team Records (${tabName})`
      );

      if (Array.isArray(rows) && rows.length) {
        console.log(`TEAM RECORDS LOADED from ${tabName}:`, rows.length);
        return rows;
      }
    } catch (err) {
      console.warn(`Team records tab failed: ${tabName}`, err);
    }
  }

  console.warn(`No team records loaded for ${tierName}.`);
  return [];
}

async function loadTierConfig() {
  const configId = extractGoogleSheetId(CONFIG_SHEET_URL);

  if (!configId) {
    throw new Error("Config sheet URL is invalid.");
  }

  const configURL = openSheetURL(configId, CONFIG_TAB);
  const rows = await fetchSheetData(configURL, "Tier Config");

  const nextConfig = {};

  rows.forEach(row => {
    const rawTier = String(getRowValue(row, [
      "Tier",
      "Tier Name"
    ])).trim();

    const tier = TIER_ORDER.find(t =>
      t.toLowerCase() === rawTier.toLowerCase()
    ) || rawTier;

    const rankingsId = extractGoogleSheetId(getRowValue(row, [
      "Power Rankings Link",
      "Power Ranking Link",
      "Power Rankings",
      "Power Ranking",
      "Rankings",
      "Ranking",
      "PR"
    ]));

    const picksId = extractGoogleSheetId(getRowValue(row, [
      "Predictions Link",
      "Prediction Link",
      "Predictions",
      "Prediction",
      "Picks",
      "Picks Link"
    ]));

    if (!tier || !rankingsId || !picksId) return;

    nextConfig[tier] = {
      rankingsId,
      picksId
    };
  });

  if (!Object.keys(nextConfig).length) {
    throw new Error("No valid tier links found in the config sheet.");
  }

  TIER_CONFIG = nextConfig;

  console.log("LOADED TIER CONFIG:", TIER_CONFIG);
}
function openSheetURL(sheetId, tabName) {
  return `https://opensheet.elk.sh/${sheetId}/${encodeURIComponent(tabName)}`;
}

function localLogoPath(team) {
  const key = String(team || "").trim().toUpperCase();
  return `${LOGO_FOLDER}/${key}.${LOGO_EXTENSION}`;
}

function tierCacheKey(tierName) {
  return `csc_article_generator_saved_data_v11_${tierName}`;
}

function saveTierData(tierName, rankData, matchData, finalRows = [], yapRows = [], teamRecordData = []) {
  const payload = {
    savedAt: new Date().toISOString(),
    rankData,
    matchData,
    finalRows,
    yapRows,
    teamRecordData
  };

  localStorage.setItem(tierCacheKey(tierName), JSON.stringify(payload));
}

function getSavedTierData(tierName) {
  const raw = localStorage.getItem(tierCacheKey(tierName));
  if (!raw) return null;

  try {
    const payload = JSON.parse(raw);

    if (!Array.isArray(payload.rankData) || !Array.isArray(payload.matchData)) {
      return null;
    }

    if (!Array.isArray(payload.finalRows)) {
      payload.finalRows = [];
    }

    if (!Array.isArray(payload.yapRows)) {
      payload.yapRows = [];
    }

    if (!Array.isArray(payload.teamRecordData)) {
      payload.teamRecordData = [];
    }

    return payload;
  } catch {
    return null;
  }
}

function formatSavedTime(iso) {
  if (!iso) return "unknown";

  return new Date(iso).toLocaleString([], {
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit"
  });
}

function showPage(page, button) {
  document.getElementById("power-page").classList.toggle("active", page === "power");
  document.getElementById("matches-page").classList.toggle("active", page === "matches");
  document.getElementById("article-page").classList.toggle("active", page === "article");
  document.getElementById("predictionArticle-page").classList.toggle("active", page === "predictionArticle");

  document.querySelectorAll(".menu button").forEach(btn => btn.classList.remove("active"));
  if (button) button.classList.add("active");
}

function safeName(text) {
  return String(text || "")
    .replace(/[^a-z0-9]/gi, "_")
    .replace(/_+/g, "_")
    .toLowerCase();
}

function logo(team, map, fallbackClass = "power-logo-fallback") {
  const key = String(team || "").trim().toUpperCase();
  const src = localLogoPath(key);

  return `
    <img
      src="${src}"
      alt="${key} logo"
      onerror="this.outerHTML='<div class=&quot;${fallbackClass}&quot;>${key}</div>'"
    >
  `;
}
function getTeamStyle(team) {
  const key = String(team || "").trim().toUpperCase();
  return teamStyles[key] || defaultStyle;
}

function movement(lastWeek, current) {
  const last = Number(lastWeek);
  const now = Number(current);

  if (!Number.isFinite(last) || !Number.isFinite(now)) {
    return `<span class="move-same">—</span>`;
  }

  const diff = last - now;

  if (diff > 0) return `<span class="move-up">▲${diff}</span>`;
  if (diff < 0) return `<span class="move-down">▼${Math.abs(diff)}</span>`;
  return `<span class="move-same">—</span>`;
}

function rankDisplay(value) {
  const cleaned = String(value || "").trim().replace("#", "");
  return cleaned ? `#${cleaned}` : "—";
}

function isRankOnlyValue(value) {
  const cleaned = String(value || "")
    .trim()
    .replace("#", "");

  return /^\d+(\.\d+)?$/.test(cleaned);
}

function cleanRankValue(value) {
  return String(value || "")
    .trim()
    .replace("#", "");
}

function extractOpenSheetInfo(url) {
  const match = String(url || "").match(/opensheet\.elk\.sh\/([^\/?#]+)\/([^?#]+)/);

  if (!match) return null;

  return {
    sheetId: decodeURIComponent(match[1]),
    tabName: decodeURIComponent(match[2])
  };
}

function makeUniqueHeaders(headers) {
  const seen = {};

  return headers.map((header, index) => {
    let clean = String(header || "").trim();

    if (!clean) clean = `Column ${index + 1}`;

    const base = clean;
    const count = seen[base] || 0;
    seen[base] = count + 1;

    return count ? `${base}__${count + 1}` : base;
  });
}

function csvRowsToObjects(rows) {
  if (!Array.isArray(rows) || rows.length < 2) return [];

  const headers = makeUniqueHeaders(rows[0]);

  return rows.slice(1).map(row => {
    const obj = {};

    headers.forEach((header, index) => {
      obj[header] = row[index] || "";
    });

    return obj;
  });
}

async function fetchCsvRowsOnly(sheetId, tabName, label) {
  const response = await fetch(csvSheetURL(sheetId, tabName), {
    cache: "no-store"
  });

  if (!response.ok) {
    throw new Error(`${label} CSV failed to load. Status: ${response.status}`);
  }

  const text = await response.text();
  const rows = parseCSV(text);

  if (!rows.length) {
    throw new Error(`${label} CSV returned no usable rows.`);
  }

  return rows;
}

async function fetchSheetData(url, label) {
  try {
    const response = await fetch(url, {
      cache: "no-store"
    });

    if (!response.ok) {
      throw new Error(`${label} failed to load through OpenSheet. Status: ${response.status}`);
    }

    const data = await response.json();

    if (!Array.isArray(data)) {
      throw new Error(`${label} did not return rows.`);
    }

    return data;
  } catch (openSheetErr) {
    const sheetInfo = extractOpenSheetInfo(url);

    if (!sheetInfo) {
      throw openSheetErr;
    }

    if (isFileProtocol()) {
      console.warn(`${label} OpenSheet pull failed. Google CSV fallback skipped because the app is running from file://. Use localhost/Netlify for CSV fallback.`, openSheetErr);
      throw openSheetErr;
    }

    console.warn(`${label} OpenSheet pull failed. Trying Google CSV fallback.`, openSheetErr);

    const csvRows = await fetchCsvRowsOnly(
      sheetInfo.sheetId,
      sheetInfo.tabName,
      label
    );

    return csvRowsToObjects(csvRows);
  }
}

function csvSheetURL(sheetId, tabName) {
  return `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(tabName)}`;
}

function parseCSV(text) {
  const cleanText = String(text || "").replace(/^\uFEFF/, "");
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;

  for (let i = 0; i < cleanText.length; i++) {
    const char = cleanText[i];
    const next = cleanText[i + 1];

    if (char === '"') {
      if (inQuotes && next === '"') {
        cell += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (char === "," && !inQuotes) {
      row.push(cell);
      cell = "";
      continue;
    }

    if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") i++;
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
      continue;
    }

    cell += char;
  }

  row.push(cell);
  rows.push(row);

  return rows.filter(r => r.some(c => String(c || "").trim()));
}

function gvizSheetURL(sheetId, tabName) {
  const handlerName = `__cscGviz_${Date.now()}_${Math.random().toString(36).slice(2)}`;
  const url = `https://docs.google.com/spreadsheets/d/${sheetId}/gviz/tq?` +
    `tqx=out:json;responseHandler:${handlerName}&` +
    `headers=0&sheet=${encodeURIComponent(tabName)}`;

  return { url, handlerName };
}

function gvizValueToString(cell) {
  if (!cell) return "";
  if (cell.f !== undefined && cell.f !== null) return String(cell.f);
  if (cell.v !== undefined && cell.v !== null) return String(cell.v);
  return "";
}

function gvizTableToRows(table) {
  const rows = table && Array.isArray(table.rows) ? table.rows : [];
  const columnCount = Math.max(
    table && Array.isArray(table.cols) ? table.cols.length : 0,
    ...rows.map(row => Array.isArray(row.c) ? row.c.length : 0),
    14
  );

  return rows
    .map(row => {
      const cells = Array.isArray(row.c) ? row.c : [];
      const out = [];

      for (let i = 0; i < columnCount; i++) {
        out.push(gvizValueToString(cells[i]));
      }

      return out;
    })
    .filter(row => row.some(cell => String(cell || "").trim()));
}

function fetchGvizRowsNoCors(sheetId, tabName, label) {
  return new Promise((resolve, reject) => {
    const { url, handlerName } = gvizSheetURL(sheetId, tabName);
    const script = document.createElement("script");
    let settled = false;

    const cleanup = () => {
      try { delete window[handlerName]; } catch { window[handlerName] = undefined; }
      if (script.parentNode) script.parentNode.removeChild(script);
    };

    const fail = message => {
      if (settled) return;
      settled = true;
      cleanup();
      reject(new Error(message));
    };

    window[handlerName] = response => {
      if (settled) return;
      settled = true;
      cleanup();

      if (!response || response.status === "error") {
        reject(new Error(`${label} Google visualization fallback failed.`));
        return;
      }

      resolve(gvizTableToRows(response.table));
    };

    script.onerror = () => fail(`${label} Google visualization script failed to load.`);
    script.src = url;
    document.head.appendChild(script);

    setTimeout(() => fail(`${label} Google visualization fallback timed out.`), 12000);
  });
}

function openSheetRowsPreservingHeader(jsonRows) {
  if (!Array.isArray(jsonRows) || !jsonRows.length) return [];

  const headers = Object.keys(jsonRows[0]);
  const valueRows = jsonRows.map(row => headers.map(header => row[header] || ""));

  // OpenSheet uses the first sheet row as headers. For Final tabs, row 1 can be
  // a real article cell, including N1. Add the header row back so N1 is not lost.
  return [headers, ...valueRows].filter(row =>
    row.some(cell => String(cell || "").trim())
  );
}

async function fetchRawSheetRows(sheetId, tabName, label) {
  try {
    return await fetchCsvRowsOnly(sheetId, tabName, label);
  } catch (csvErr) {
    console.warn(`${label} CSV pull failed. Trying Google visualization fallback.`, csvErr);

    try {
      return await fetchGvizRowsNoCors(sheetId, tabName, label);
    } catch (gvizErr) {
      console.warn(`${label} Google visualization fallback failed. Trying OpenSheet fallback.`, gvizErr);

      const jsonRows = await fetchSheetData(openSheetURL(sheetId, tabName), label);
      return openSheetRowsPreservingHeader(jsonRows);
    }
  }
}

function escapeHTML(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function getFirstNonEmptyCell(row) {
  return (row || []).find(cell => String(cell || "").trim()) || "";
}

function cleanArticleLine(value) {
  return String(value || "")
    .replace(/\s+/g, " ")
    .trim();
}

function parseWriterBlurb(line) {
  const writerRankMatch = line.match(/^(.+?)\s*#\s*(\d+(?:\.\d+)?)\s*:\s*(.+)$/);

  if (writerRankMatch) {
    return {
      writer: writerRankMatch[1].trim(),
      rank: writerRankMatch[2].trim(),
      text: writerRankMatch[3].trim(),
      raw: `${writerRankMatch[1].trim()} #${writerRankMatch[2].trim()}: ${writerRankMatch[3].trim()}`
    };
  }

  const colonIndex = line.indexOf(":");

  if (colonIndex > -1) {
    return {
      writer: line.slice(0, colonIndex).trim(),
      rank: "",
      text: line.slice(colonIndex + 1).trim(),
      raw: line
    };
  }

  return {
    writer: "",
    rank: "",
    text: line,
    raw: line
  };
}

function parseFinalArticleRows(finalRows) {
  const sections = {};
  let current = null;

  finalRows.forEach(row => {
    const firstCell = getFirstNonEmptyCell(row);
    if (!firstCell) return;

    const lines = String(firstCell)
      .split(/\r?\n/)
      .map(cleanArticleLine)
      .filter(Boolean);

    lines.forEach(line => {
      const rankMatch = line.match(/^Rank\s*#?\s*(\d+)\s*:\s*(.+)$/i);

      if (rankMatch) {
        const rank = rankMatch[1].trim();
        const team = rankMatch[2].trim().toUpperCase();

        if (sections[team]) {
          current = null;
          return;
        }

        current = {
          rank,
          team,
          blurbs: [],
          rawLines: [`Rank ${rank}: ${team}`]
        };

        sections[team] = current;
        return;
      }

      if (!current) return;

      const blurb = parseWriterBlurb(line);
      if (!blurb.text) return;

      current.blurbs.push(blurb);
      current.rawLines.push(blurb.raw);
    });
  });

  return sections;
}

function getSortedPowerData(data) {
  return [...data].sort((a, b) => {
    const rankA = getFinalRankNumber(a);
    const rankB = getFinalRankNumber(b);

    if (rankA !== rankB) return rankA - rankB;;

    const teamA = String(a.Teams || a.Team || "");
    const teamB = String(b.Teams || b.Team || "");

    return teamA.localeCompare(teamB);
  });
}

function extractTeamsFromData(rankData, matchData) {
  const teams = new Set();

  rankData.forEach(row => {
    const team = String(row.Teams || row.Team || "").trim().toUpperCase();
    if (team) teams.add(team);
  });

  matchData.forEach(row => {
    const matchup = String(row.Matchup || "").trim();

    if (matchup) {
      const parts = matchup.split(/vs/i).map(t => t.trim().toUpperCase());
      parts.forEach(team => {
        if (team) teams.add(team);
      });
    }

    Object.keys(row).forEach(key => {
      if (key === "Matchup") return;

      const pick = String(row[key] || "").trim().toUpperCase();
      if (pick) teams.add(pick);
    });
  });

  return Array.from(teams);
}
function loadImage(src) {
  return new Promise((resolve, reject) => {
    const img = new Image();

    img.crossOrigin = "anonymous";

    img.onload = () => resolve(img);
    img.onerror = () => reject(new Error("Logo not found: " + src));

    img.src = src;
  });
}

function rgbToHex(r, g, b) {
  return "#" + [r, g, b]
    .map(v => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, "0"))
    .join("")
    .toUpperCase();
}

function hexToRgb(hex) {
  const clean = String(hex || "").replace("#", "");

  if (!/^[0-9a-fA-F]{6}$/.test(clean)) {
    return { r: 51, g: 65, b: 85 };
  }

  return {
    r: parseInt(clean.slice(0, 2), 16),
    g: parseInt(clean.slice(2, 4), 16),
    b: parseInt(clean.slice(4, 6), 16)
  };
}

function rgbToHsl(r, g, b) {
  r /= 255;
  g /= 255;
  b /= 255;

  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);

  let h = 0;
  let s = 0;
  const l = (max + min) / 2;

  if (max !== min) {
    const d = max - min;

    s = l > .5
      ? d / (2 - max - min)
      : d / (max + min);

    switch (max) {
      case r:
        h = (g - b) / d + (g < b ? 6 : 0);
        break;
      case g:
        h = (b - r) / d + 2;
        break;
      case b:
        h = (r - g) / d + 4;
        break;
    }

    h /= 6;
  }

  return { h, s, l };
}

function colorDistance(a, b) {
  const ca = hexToRgb(a);
  const cb = hexToRgb(b);

  return Math.sqrt(
    Math.pow(ca.r - cb.r, 2) +
    Math.pow(ca.g - cb.g, 2) +
    Math.pow(ca.b - cb.b, 2)
  );
}

function darkenHex(hex, amount = .45) {
  const rgb = hexToRgb(hex);

  return rgbToHex(
    rgb.r * amount,
    rgb.g * amount,
    rgb.b * amount
  );
}

function softenHex(hex, amount = .75) {
  const rgb = hexToRgb(hex);

  return rgbToHex(
    rgb.r * amount + 20,
    rgb.g * amount + 20,
    rgb.b * amount + 20
  );
}

async function extractColorsFromLogo(team) {
  const key = String(team || "").trim().toUpperCase();

  if (teamColorCache[key]) {
    return teamColorCache[key];
  }

  const src = localLogoPath(key);

  try {
    const img = await loadImage(src);

    const canvas = document.createElement("canvas");
    const size = 72;

    canvas.width = size;
    canvas.height = size;

    const ctx = canvas.getContext("2d", { willReadFrequently: true });
    ctx.clearRect(0, 0, size, size);
    ctx.drawImage(img, 0, 0, size, size);

    const pixels = ctx.getImageData(0, 0, size, size).data;
    const buckets = new Map();

    for (let i = 0; i < pixels.length; i += 4) {
      const r = pixels[i];
      const g = pixels[i + 1];
      const b = pixels[i + 2];
      const a = pixels[i + 3];

      if (a < 90) continue;

      const hsl = rgbToHsl(r, g, b);

      if (hsl.l < .06) continue;
      if (hsl.l > .94 && hsl.s < .25) continue;

      const qr = Math.round(r / 24) * 24;
      const qg = Math.round(g / 24) * 24;
      const qb = Math.round(b / 24) * 24;

      const bucketKey = `${qr},${qg},${qb}`;
      const weight = (a / 255) * (.45 + hsl.s * 1.35) * (.75 + Math.abs(.52 - hsl.l));

      const current = buckets.get(bucketKey) || {
        r: 0,
        g: 0,
        b: 0,
        weight: 0
      };

      current.r += r * weight;
      current.g += g * weight;
      current.b += b * weight;
      current.weight += weight;

      buckets.set(bucketKey, current);
    }

    const colors = Array.from(buckets.values())
      .filter(bucket => bucket.weight > 0)
      .map(bucket => {
        const r = bucket.r / bucket.weight;
        const g = bucket.g / bucket.weight;
        const b = bucket.b / bucket.weight;

        return {
          hex: rgbToHex(r, g, b),
          weight: bucket.weight
        };
      })
      .sort((a, b) => b.weight - a.weight);

    if (!colors.length) {
      teamColorCache[key] = defaultStyle;
      return defaultStyle;
    }

    const primary = colors[0].hex;

    let secondary = colors.find(c => colorDistance(c.hex, primary) > 95)?.hex;
    let accent = colors.find(c =>
      colorDistance(c.hex, primary) > 65 &&
      (!secondary || colorDistance(c.hex, secondary) > 65)
    )?.hex;

    if (!secondary) secondary = darkenHex(primary, .38);
    if (!accent) accent = softenHex(primary, .82);

    const style = {
      primary,
      secondary,
      accent,
      accent2: darkenHex(secondary, .72),
      detail: accent,
      dark: "#020617"
    };

    teamColorCache[key] = style;
    return style;
  } catch (err) {
    console.warn(getSafeErrorMessage(err));

    teamColorCache[key] = defaultStyle;
    return defaultStyle;
  }
}

async function prepareTeamAssets(rankData, matchData) {
  const teams = extractTeamsFromData(rankData, matchData);

  teams.forEach(team => {
    teamLogos[team] = localLogoPath(team);
  });

  try {
    const logoData = await fetchSheetData(logosURL, "Logo Colors");

    const colorMap = {};

    logoData.forEach(row => {
      const team = String(
        row["Team Name"] ||
        row.Team ||
        row.Teams ||
        ""
      ).trim().toUpperCase();

      if (!team) return;

      colorMap[team] = parseTeamColors(row);
    });

    teams.forEach(team => {
      teamStyles[team] = colorMap[team] || defaultStyle;
    });

    console.log("LOCAL LOGOS READY:", teamLogos);
    console.log("SHEET COLORS READY:", teamStyles);
  } catch (err) {
    console.warn("Color sheet failed. Using default colors.", err);

    teams.forEach(team => {
      teamStyles[team] = defaultStyle;
    });
  }
}

function resizeTeamNames() {
  document.querySelectorAll(".power-team-name").forEach(name => {
    const container = name.closest(".power-content");
    if (!container) return;

    let size = 116;
    name.style.fontSize = size + "px";

    while (name.scrollWidth > container.clientWidth && size > 62) {
      size -= 4;
      name.style.fontSize = size + "px";
    }
  });

  document.querySelectorAll(".power-team-subtitle").forEach(name => {
    const container = name.closest(".power-content");
    if (!container) return;

    let size = 54;
    name.style.fontSize = size + "px";

    while (name.scrollWidth > container.clientWidth && size > 30) {
      size -= 3;
      name.style.fontSize = size + "px";
    }
  });
}

function resizeAnalysts() {
  document.querySelectorAll(".power-analysts").forEach(row => {
    const votes = row.querySelectorAll(".analyst-vote");
    const count = votes.length;

    let gap = 22;
    let boxWidth = 260;
    let boxHeight = 90;
    let nameSize = 34;
    let rankSize = 42;

    if (count > 6) {
      gap = 16;
      boxWidth = 235;
      boxHeight = 84;
      nameSize = 30;
      rankSize = 38;
    }

    if (count > 10) {
      gap = 10;
      boxWidth = 205;
      boxHeight = 76;
      nameSize = 25;
      rankSize = 34;
    }

    row.style.gap = gap + "px";
    row.style.flexWrap = "wrap";

    votes.forEach(vote => {
      vote.style.minWidth = boxWidth + "px";
      vote.style.height = boxHeight + "px";

      const name = vote.querySelector(".vote-name");
      const rank = vote.querySelector(".vote-rank");

      if (name) name.style.fontSize = nameSize + "px";
      if (rank) rank.style.fontSize = rankSize + "px";
    });
  });
}

function waitForImages(element) {
  const imgs = Array.from(element.querySelectorAll("img"));

  return Promise.all(imgs.map(img => {
    if (img.complete && img.naturalWidth > 0) return Promise.resolve();

    return new Promise(resolve => {
      img.onload = resolve;
      img.onerror = resolve;
    });
  }));
}

const TRANSPARENT_PIXEL = "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==";

function blobToDataURL(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

function resolveImageURL(src) {
  try {
    return new URL(src, window.location.href).href;
  } catch {
    return src;
  }
}

function fallbackLogoDataURL(img) {
  const label = String(img.alt || "LOGO")
    .replace(/\s+logo$/i, "")
    .trim()
    .slice(0, 12) || "LOGO";

  const safeLabel = escapeHTML(label);

  const svg = `
    <svg xmlns="http://www.w3.org/2000/svg" width="240" height="240" viewBox="0 0 240 240">
      <rect width="240" height="240" rx="28" fill="#0f172a"/>
      <rect x="8" y="8" width="224" height="224" rx="24" fill="none" stroke="rgba(255,255,255,.25)" stroke-width="8"/>
      <text x="120" y="126" dominant-baseline="middle" text-anchor="middle" font-family="Arial, sans-serif" font-size="42" font-weight="900" fill="#ffffff">${safeLabel}</text>
    </svg>
  `;

  return "data:image/svg+xml;charset=utf-8," + encodeURIComponent(svg);
}

function shouldUseCanvasSafeFallback(src) {
  const cleanSrc = String(src || "").trim();

  if (!cleanSrc) return true;
  if (/^data:/i.test(cleanSrc)) return false;
  if (/^blob:/i.test(cleanSrc)) return false;

  // When the app is opened directly as file://, browsers block fetch/canvas
  // export for local images. The on-screen card can still show the real logo,
  // but the rich-copy canvas needs a same-document fallback.
  if (isFileProtocol()) return true;

  const resolved = resolveImageURL(cleanSrc);
  if (/^file:/i.test(resolved)) return true;

  return false;
}

function makeClonedImagesCanvasSafe(clonedElement) {
  if (!clonedElement) return;

  const imgs = Array.from(clonedElement.querySelectorAll("img"));

  imgs.forEach(img => {
    const src = img.getAttribute("src") || "";

    if (!shouldUseCanvasSafeFallback(src)) return;

    img.removeAttribute("srcset");
    img.removeAttribute("crossorigin");
    img.removeAttribute("onerror");
    img.setAttribute("src", fallbackLogoDataURL(img));
  });
}

async function renderToCanvas(card) {
  if (!card) throw new Error("Card not found for canvas render.");

  const oldTransform = card.style.transform;
  const oldMargin = card.style.marginBottom;

  const buttons = card.querySelectorAll("button");

  try {
    buttons.forEach(btn => btn.style.display = "none");
    await waitForImages(card);

    card.style.transform = "none";
    card.style.marginBottom = "0";

    return await html2canvas(card, {
      backgroundColor: null,
      scale: 1,
      useCORS: true,
      allowTaint: false,
      imageTimeout: 15000,
      onclone: (clonedDocument, clonedElement) => {
        if (!clonedElement) return;

        clonedElement.style.transform = "none";
        clonedElement.style.marginBottom = "0";

        clonedElement.querySelectorAll("button").forEach(btn => {
          btn.style.display = "none";
        });

        makeClonedImagesCanvasSafe(clonedElement);
      }
    });
  } finally {
    card.style.transform = oldTransform;
    card.style.marginBottom = oldMargin;

    buttons.forEach(btn => btn.style.display = "");
  }
}

function downloadCanvasPNG(canvas, filename) {
  const link = document.createElement("a");
  link.download = filename;
  link.href = canvas.toDataURL("image/png");
  document.body.appendChild(link);
  link.click();
  link.remove();
}

async function copyCard(cardId, button) {
  const oldText = button ? button.innerText : "";

  try {
    const card = document.getElementById(cardId);
    if (!card) throw new Error("Card not found: " + cardId);

    const canvas = await renderToCanvas(card);

    const blob = await new Promise(resolve => {
      canvas.toBlob(resolve, "image/png");
    });

    if (!blob) throw new Error("Could not create PNG blob.");

    try {
      await navigator.clipboard.write([
        new ClipboardItem({ "image/png": blob })
      ]);

      if (button) {
        button.innerText = "Copied!";
        button.classList.add("copied");

        setTimeout(() => {
          button.innerText = oldText;
          button.classList.remove("copied");
        }, 1200);
      }
    } catch (clipboardErr) {
      console.warn("Clipboard image write failed. Downloading PNG fallback.", clipboardErr);
      downloadCanvasPNG(canvas, `${cardId}.png`);

      if (button) {
        button.innerText = "Downloaded";
        setTimeout(() => button.innerText = oldText, 1600);
      }

      alert("Clipboard blocked the image copy, so I downloaded the PNG instead.");
    }
  } catch (err) {
    console.error("COPY ERROR:", err);
    if (button) button.innerText = oldText;
    alert("Copy failed: " + getSafeErrorMessage(err));
  }
}
function normalizeHeaderName(value) {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
}

function getPowerColumn(row, possibleNames) {
  const wanted = possibleNames.map(normalizeHeaderName);

  const actualKey = Object.keys(row).find(key =>
    wanted.includes(normalizeHeaderName(key))
  );

  return actualKey ? row[actualKey] : "";
}

function formatMovementValue(value, lastWeek, currentRank) {
  const raw = String(value || "").trim();
  const lastRaw = String(lastWeek || "").trim().replace("#", "");
  const currentRaw = String(currentRank || "").trim().replace("#", "");

  const last = Number(lastRaw);
  const current = Number(currentRaw);

  // First choice: use the Movement column if it actually has something.
  if (raw && raw !== "-" && raw !== "—") {
    const clean = raw.replace("#", "").trim();
    const lowered = clean.toLowerCase();

    if (["0", "same", "-", "—", "--", "none"].includes(lowered)) {
      return `<span class="move-same">-</span>`;
    }

    const numberMatch = clean.match(/-?\d+(\.\d+)?/);

    if (
      lowered.includes("up") ||
      clean.includes("▲") ||
      clean.includes("↑") ||
      clean.startsWith("+")
    ) {
      const num = numberMatch ? Math.abs(Number(numberMatch[0])) : "";
      return num ? `<span class="move-up">▲${num}</span>` : `<span class="move-same">-</span>`;
    }

    if (
      lowered.includes("down") ||
      clean.includes("▼") ||
      clean.includes("↓") ||
      clean.startsWith("-")
    ) {
      const num = numberMatch ? Math.abs(Number(numberMatch[0])) : "";
      return num ? `<span class="move-down">▼${num}</span>` : `<span class="move-same">-</span>`;
    }

    if (numberMatch) {
      const num = Number(numberMatch[0]);

      if (num > 0) return `<span class="move-up">▲${num}</span>`;
      if (num < 0) return `<span class="move-down">▼${Math.abs(num)}</span>`;

      return `<span class="move-same">-</span>`;
    }
  }

  // Fallback: calculate movement from Last Week and Current Rank.
  if (!Number.isFinite(last) || !Number.isFinite(current)) {
    return `<span class="move-same">-</span>`;
  }

  const diff = last - current;

  if (diff > 0) return `<span class="move-up">▲${diff}</span>`;
  if (diff < 0) return `<span class="move-down">▼${Math.abs(diff)}</span>`;

  return `<span class="move-same">-</span>`;
}

const POWER_COLUMN_HEADERS = {
  currentRank: ["Final Rank", "Current Rank", "Rank", "Ranking"],
  lastWeek: ["Last Week", "Last Week Rank", "Previous Week", "Previous Week Rank", "Previous Rank", "Prev Rank", "Prev", "Last Rank"],
  movement: ["Movement", "Move", "Rank Movement"]
};

const IGNORED_POWER_COLUMNS = [
  "Teams", "Team", "Avg Rank", "Final Rank", "Current Rank", "Rank", "Ranking",
  "Last Week", "Last Week Rank", "Previous Week", "Previous Week Rank", "Previous Rank", "Prev Rank", "Prev", "Last Rank",
  "Movement", "Move", "Rank Movement", "Writeup", "Write Up", "Write-Up", "Blurb", "Notes", "Comment", "Comments",
  "Description", "Reason", "Reasoning", "Article", "Text"
];

function getPowerCardModel(row, index, idPrefix = "power", teamRecordMap = {}) {
  const team = String(row.Teams || row.Team || "").trim().toUpperCase();
  const style = getTeamStyle(team);
  const currentRank = getPowerColumn(row, POWER_COLUMN_HEADERS.currentRank);
  const lastWeek = getPowerColumn(row, POWER_COLUMN_HEADERS.lastWeek);
  const movementValue = getPowerColumn(row, POWER_COLUMN_HEADERS.movement);
  const teamRecord = teamRecordMap[normalizeTeamKey(team)] || {};
  const ignoredPowerColumnsNormalized = IGNORED_POWER_COLUMNS.map(normalizeHeaderName);

  const analysts = Object.keys(row)
    .filter(key => !ignoredPowerColumnsNormalized.includes(normalizeHeaderName(key)))
    .filter(key => isRankOnlyValue(row[key]))
    .map(name => ({
      name,
      rank: cleanRankValue(row[name])
    }));

  return {
    team,
    teamName: teamRecord.teamName || "",
    franchiseName: teamRecord.franchiseDisplay || "",
    record: teamRecord.record || "",
    style,
    currentRank,
    lastWeek,
    movementValue,
    analysts,
    cardId: `${idPrefix}_${safeName(team)}_${index}`
  };
}

function buildPowerCardHTML(row, index, options = {}) {
  const model = getPowerCardModel(
    row,
    index,
    options.idPrefix || "power",
    options.teamRecordMap || {}
  );
  if (!model.team) return "";

  const analystHTML = model.analysts.map(analyst => `
    <div class="analyst-vote">
      <div class="vote-name">${escapeHTML(analyst.name)}</div>
      <div class="vote-rank">#${escapeHTML(analyst.rank)}</div>
    </div>
  `).join("");

  return `
    <div
      id="${model.cardId}"
      class="power-card ${model.record ? "has-record" : ""}"
      style="
        --team-primary:${model.style.primary};
        --team-secondary:${model.style.secondary};
        --team-accent:${model.style.accent};
        --team-accent2:${model.style.accent2};
        --team-detail:${model.style.detail};
        --team-dark:${model.style.dark};
      "
    >
      <button class="card-copy-overlay" onclick="copyCard('${model.cardId}', this)">Copy</button>

      <div class="power-accent-slab"></div>

      <div class="power-logo-big">
        ${logo(model.team, teamLogos)}
      </div>

      <div class="power-content">
        <div class="power-title-block">
          <div class="power-team-name">${escapeHTML(model.teamName || model.team)}</div>
          ${model.franchiseName ? `<div class="power-team-subtitle">${escapeHTML(model.franchiseName)}</div>` : ""}
        </div>

        <div class="power-stats-line">
          <div class="stat-block">
            <span class="stat-label">RANKING:</span>
            <span class="stat-value">${rankDisplay(model.currentRank)}</span>
          </div>

          <div class="stat-block">
            <span class="stat-label">LAST WEEK:</span>
            <span class="stat-value">${rankDisplay(model.lastWeek)}</span>
          </div>

          <div class="stat-block">
            <span class="stat-label">MOVEMENT:</span>
            <span class="stat-value">${formatMovementValue(model.movementValue, model.lastWeek, model.currentRank)}</span>
          </div>
        </div>

        <div class="power-analysts">
          ${analystHTML}
        </div>
      </div>

      ${model.record ? `
        <div class="power-record-box">
          <div class="record-label">CURRENT RECORD</div>
          <div class="record-value">${escapeHTML(model.record)}</div>
        </div>
      ` : ""}
    </div>
  `;
}

function buildPowerCards(data, teamRecordData = []) {
  const gallery = document.getElementById("powerGallery");
  gallery.innerHTML = "";

  const teamRecordMap = buildTeamRecordMap(teamRecordData, data);

  getSortedPowerData(data).forEach((row, index) => {
    const team = String(row.Teams || row.Team || "").trim().toUpperCase();
    if (!team) return;

    const wrap = document.createElement("div");
    wrap.className = "graphic-wrap";
    wrap.innerHTML = `
      <div class="card-actions"></div>
      ${buildPowerCardHTML(row, index, { idPrefix: "power", teamRecordMap })}
    `;

    gallery.appendChild(wrap);
  });

  setTimeout(() => {
    resizeTeamNames();
    resizeAnalysts();
  }, 150);
}

function removeRankTitleLine(text) {
  return String(text || "")
    .split(/\r?\n/)
    .map(line => line.trim())
    .filter(Boolean)
    .filter((line, index) => {
      if (index !== 0) return true;
      return !/^Rank\s*#?\s*\d+\s*:/i.test(line);
    })
    .join("\n");
}

function buildPowerArticleGraphicHTML(row, index, teamRecordMap = {}) {
  return buildPowerCardHTML(row, index, { idPrefix: "article_power", teamRecordMap });
}

function articleTextForSection(team, section, fallbackRank) {
  if (!section) {
    return `Rank ${fallbackRank || ""}: ${team}\nNo writeups found in the Final tab.`.trim();
  }

  return section.rawLines.join("\n").trim();
}

function buildArticleView(rankData, finalRows, teamRecordData = []) {
  const gallery = document.getElementById("articleGallery");
  gallery.innerHTML = "";
  currentArticleTextBlocks = [];

  const sortedData = getSortedPowerData(rankData).reverse();
  const teamRecordMap = buildTeamRecordMap(teamRecordData);

  const toolbar = document.createElement("div");
  toolbar.className = "article-toolbar";
  toolbar.innerHTML = `
    <div>
      <div class="article-toolbar-title">Power Rankings Article View</div>
      <div class="article-toolbar-note">Graphic first, then Final column N writeups underneath by row order.</div>
    </div>
    <button onclick="copyFullArticleText(this)">Copy Full Article</button>
  `;
  gallery.appendChild(toolbar);

  sortedData.forEach((row, index) => {
    const team = String(row.Teams || row.Team || "").trim().toUpperCase();
    if (!team) return;

    const finalRank = cleanRankValue(getPowerColumn(row, [
      "Final Rank",
      "Current Rank",
      "Rank",
      "Ranking"
    ]));

    const finalText = getPowerFinalTextForTeam(finalRows || [], team, finalRank, index, sortedData.length);
	const cleanFinalText = removeRankTitleLine(finalText);
    const fallbackText = `Rank ${finalRank || ""}: ${team}\nNo writeups found for this team after checking Final column N.`.trim();
    const copyText = cleanFinalText || fallbackText;
    const copyIndex = currentArticleTextBlocks.length;
    currentArticleTextBlocks.push(copyText);

    const titleLine = firstNonEmptyLine(finalText) || `Rank ${finalRank}: ${team}`;
    const blurbsHTML = finalTextToArticleHTML(
	   cleanFinalText,
      `No writeups found for ${team} after checking Final column N.`
    );

    const item = document.createElement("div");
    item.className = "article-item";
    item.innerHTML = `
      ${buildPowerArticleGraphicHTML(row, index, teamRecordMap)}

      <div class="article-text-box">
        <div class="article-rank-title">
          <span>${escapeHTML(titleLine)}</span>
          <button class="article-copy-btn" onclick="copyTeamArticleText(${copyIndex}, this)">Copy Text</button>
        </div>

        ${blurbsHTML}
      </div>
    `;

    gallery.appendChild(item);
  });

  setTimeout(() => {
    resizeTeamNames();
    resizeAnalysts();
  }, 150);
}

function articlePlainTextToHTML(text) {
  const lines = String(text || "")
    .split(/\r?\n/)
    .map(line => line.trim())
    .filter(Boolean);

  if (!lines.length) return "";

  return lines.map((line, index) => {
    if (index === 0 && /^Rank\s+/i.test(line)) {
      return `<h3 style="margin:18px 0 10px;font-family:Arial,sans-serif;font-size:20px;line-height:1.25;color:#111827;">${profileLinkedHTML(line)}</h3>`;
    }

    const colonIndex = line.indexOf(":");

    if (colonIndex > -1) {
      const label = line.slice(0, colonIndex + 1);
      const body = line.slice(colonIndex + 1).trim();

      return `<p style="margin:10px 0;font-family:Arial,sans-serif;font-size:15px;line-height:1.45;color:#111827;"><strong>${profileLinkedHTML(label)}</strong> ${profileLinkedHTML(body)}</p>`;
    }

    return `<p style="margin:10px 0;font-family:Arial,sans-serif;font-size:15px;line-height:1.45;color:#111827;">${profileLinkedHTML(line)}</p>`;
  }).join("");
}

async function powerCardToDataURL(card) {
  if (!card) throw new Error("Article graphic card not found.");

  const canvas = await renderToCanvas(card);
  return canvas.toDataURL("image/png");
}

async function buildFullArticleClipboardHTML(button) {
  const articleItems = Array.from(document.querySelectorAll("#articleGallery .article-item"));

  if (!articleItems.length) {
    throw new Error("No article sections found.");
  }

  const htmlBlocks = [];

  for (let i = 0; i < articleItems.length; i++) {
    if (button) button.innerText = `Copying ${i + 1}/${articleItems.length}...`;

    const item = articleItems[i];
    const card = item.querySelector(".power-card");
    const dataURL = await powerCardToDataURL(card);
    const text = currentArticleTextBlocks[i] || "";

    htmlBlocks.push(`
      <div style="margin:0 0 34px 0;page-break-inside:avoid;break-inside:avoid;">
        <img src="${dataURL}" alt="Power ranking graphic" style="display:block;width:900px;max-width:100%;height:auto;margin:0 0 14px 0;border:0;">
        <div style="font-family:Arial,sans-serif;color:#111827;">
          ${articlePlainTextToHTML(text)}
        </div>
      </div>
    `);
  }

  return `
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="UTF-8">
      </head>
      <body>
        ${htmlBlocks.join("")} 
      </body>
    </html>
  `;
}

async function copyFullArticleText(button) {
  const originalText = button ? button.innerText : "";

  try {
    const plainText = profileCleanedTextBlocks(currentArticleTextBlocks).join("\n\n").trim();

    if (!plainText) throw new Error("No article text found.");

    const html = await buildFullArticleClipboardHTML(button);

    await navigator.clipboard.write([
      new ClipboardItem({
        "text/html": new Blob([html], { type: "text/html" }),
        "text/plain": new Blob([plainText], { type: "text/plain" })
      })
    ]);

    if (button) {
      button.innerText = "Copied!";
      setTimeout(() => button.innerText = originalText, 1200);
    }
  } catch (err) {
    console.error("ARTICLE COPY ERROR:", err);

    try {
      const plainText = profileCleanedTextBlocks(currentArticleTextBlocks).join("\n\n").trim();
      if (plainText) await navigator.clipboard.writeText(plainText);

      if (button) {
        button.innerText = "Copied Text Only";
        setTimeout(() => button.innerText = originalText, 1600);
      }

      alert("Full article rich copy failed, so I copied the text only. Error: " + getSafeErrorMessage(err));
    } catch (fallbackErr) {
      if (button) button.innerText = originalText;
      alert("Article copy failed: " + getSafeErrorMessage(err));
    }
  }
}

async function copyTeamArticleText(index, button) {
  try {
    const text = currentArticleTextBlocks[index] || "";

    if (!text.trim()) throw new Error("No article text found for this team.");

    const html = `
      <!DOCTYPE html>
      <html>
        <head><meta charset="UTF-8"></head>
        <body>
          <div style="font-family:Arial,sans-serif;color:#111827;">
            ${articlePlainTextToHTML(text)}
          </div>
        </body>
      </html>
    `;

    await navigator.clipboard.write([
      new ClipboardItem({
        "text/html": new Blob([html], { type: "text/html" }),
        "text/plain": new Blob([profileCleanedPlainText(text)], { type: "text/plain" })
      })
    ]);

    if (button) {
      const oldText = button.innerText;
      button.innerText = "Copied!";
      setTimeout(() => button.innerText = oldText, 1200);
    }
  } catch (err) {
    console.error("TEAM ARTICLE COPY ERROR:", err);

    try {
      const text = currentArticleTextBlocks[index] || "";
      if (text.trim()) {
        await navigator.clipboard.writeText(profileCleanedPlainText(text));

        if (button) {
          const oldText = button.innerText;
          button.innerText = "Copied Text Only";
          setTimeout(() => button.innerText = oldText, 1400);
        }

        return;
      }
    } catch (fallbackErr) {
      console.error("TEAM ARTICLE TEXT FALLBACK ERROR:", fallbackErr);
    }

    alert("Article copy failed: " + getSafeErrorMessage(err));
  }
}

function splitIntoBalancedRows(items) {
  const count = items.length;

  if (count <= 3) return [items];
  if (count === 4) return [items.slice(0, 2), items.slice(2)];
  if (count === 5) return [items.slice(0, 3), items.slice(3)];
  if (count === 6) return [items.slice(0, 3), items.slice(3)];
  if (count === 7) return [items.slice(0, 4), items.slice(4)];
  if (count === 8) return [items.slice(0, 4), items.slice(4)];

  const midpoint = Math.ceil(count / 2);

  return [
    items.slice(0, midpoint),
    items.slice(midpoint)
  ];
}

function buildMatchCards(data) {
  const gallery = document.getElementById("matchGallery");
  gallery.innerHTML = "";

  if (!Array.isArray(data) || data.length === 0) {
    showBlankMatchupsMessage();
    return;
  }

  let renderedCount = 0;

  data.forEach((row, index) => {
    const matchup = getMatchupValue(row);
    if (!matchup) return;

    const parts = matchup.split(/vs/i).map(t => t.trim().toUpperCase());
    const a = parts[0];
    const b = parts[1];
    if (!a || !b) return;

    renderedCount++;

    const cardId = `match_${safeName(a)}_${safeName(b)}_${index}`;

    const analysts = Object.keys(row).filter(k =>
      k !== "Matchup" && String(row[k] || "").trim()
    );

    const analystCount = analysts.length;
    const analystFont = analystCount > 8 ? "20px" : analystCount > 5 ? "23px" : "26px";
    const pickLogoSize = analystCount > 8 ? "72px" : analystCount > 5 ? "82px" : "90px";

    const pickCards = analysts.map(name => {
      const pick = String(row[name] || "").trim().toUpperCase();

      return `
        <div class="prediction-pick-card">
          <div class="prediction-name">${name}</div>
          ${logo(pick, teamLogos, "match-logo-fallback")}
        </div>
      `;
    });

    const rows = splitIntoBalancedRows(pickCards);

    const pickRowsHTML = rows.map(rowItems => `
      <div class="prediction-pick-row">
        ${rowItems.join("")}
      </div>
    `).join("");

    gallery.innerHTML += `
      <div class="graphic-wrap">
        <div
          id="${cardId}"
          class="match-card"
          style="
            --analyst-count:${analystCount};
            --analyst-font:${analystFont};
            --pick-logo-size:${pickLogoSize};
          "
        >
          <button class="card-copy-overlay" onclick="copyCard('${cardId}', this)">Copy</button>

          <div class="prediction-grid">
            <div class="matchup-cell">
              <div class="mini-team">
                ${logo(a, teamLogos, "match-logo-fallback")}
                <div>${a}</div>
              </div>

              <div class="mini-vs">VS</div>

              <div class="mini-team">
                ${logo(b, teamLogos, "match-logo-fallback")}
                <div>${b}</div>
              </div>
            </div>

            <div class="prediction-picks-wrap">
              ${pickRowsHTML}
            </div>
          </div>
        </div>
      </div>
    `;
  });

  if (renderedCount === 0) {
    showBlankMatchupsMessage();
  }
}

function buildPredictionArticleGraphicHTML(row, index) {
  const matchup = getMatchupValue(row);
  const parts = matchup.split(/vs/i).map(t => t.trim().toUpperCase());
  const a = parts[0];
  const b = parts[1];

  const cardId = `article_match_${safeName(a)}_${safeName(b)}_${index}`;

  const analysts = Object.keys(row).filter(k =>
    normalizeHeaderName(k) !== "matchup" &&
    normalizeHeaderName(k) !== "matchups" &&
    normalizeHeaderName(k) !== "matchup" &&
    String(row[k] || "").trim()
  );

  const analystCount = analysts.length;
  const analystFont = analystCount > 8 ? "20px" : analystCount > 5 ? "23px" : "26px";
  const pickLogoSize = analystCount > 8 ? "72px" : analystCount > 5 ? "82px" : "90px";

  const pickCards = analysts.map(name => {
    const pick = String(row[name] || "").trim().toUpperCase();

    return `
      <div class="prediction-pick-card">
        <div class="prediction-name">${escapeHTML(name)}</div>
        ${logo(pick, teamLogos, "match-logo-fallback")}
      </div>
    `;
  });

  const rows = splitIntoBalancedRows(pickCards);

  const pickRowsHTML = rows.map(rowItems => `
    <div class="prediction-pick-row">
      ${rowItems.join("")}
    </div>
  `).join("");

  return `
    <div
      id="${cardId}"
      class="match-card"
      style="
        --analyst-count:${analystCount};
        --analyst-font:${analystFont};
        --pick-logo-size:${pickLogoSize};
      "
    >
      <button class="card-copy-overlay" onclick="copyCard('${cardId}', this)">Copy</button>

      <div class="prediction-grid">
        <div class="matchup-cell">
          <div class="mini-team">
            ${logo(a, teamLogos, "match-logo-fallback")}
            <div>${escapeHTML(a)}</div>
          </div>

          <div class="mini-vs">VS</div>

          <div class="mini-team">
            ${logo(b, teamLogos, "match-logo-fallback")}
            <div>${escapeHTML(b)}</div>
          </div>
        </div>

        <div class="prediction-picks-wrap">
          ${pickRowsHTML}
        </div>
      </div>
    </div>
  `;
}

function normalizeMatchupKey(value) {
  return String(value || "")
    .trim()
    .toUpperCase()
    .replace(/\s+/g, " ")
    .replace(/\bVERSUS\b/g, "VS")
    .replace(/\s+V\.?\s+/g, " VS ")
    .replace(/\s*VS\s*/g, " VS ");
}

function findMatchupInRawRow(row) {
  const cells = row || [];

  for (const cell of cells) {
    const value = String(cell || "").trim();
    if (/\bvs\b/i.test(value)) return value;
  }

  return "";
}

function getRawCell(row, index) {
  return String((row || [])[index] || "").trim();
}

function looksLikeArticleHeaderOnly(value) {
  const clean = String(value || "").trim().toLowerCase();
  return ["article", "writeup", "write up", "write-up", "final", "column n"].includes(clean);
}

function getLikelyFinalArticleCell(row) {
  const cells = Array.isArray(row)
    ? row.map(cell => String(cell || "").trim())
    : Object.values(row || {}).map(cell => String(cell || "").trim());

  if (!cells.length) return "";

  // Best path: when CSV works, preserve the true Column N position.
  const columnNValue = String(cells[13] || "").trim();
  if (columnNValue && !looksLikeArticleHeaderOnly(columnNValue)) {
    return columnNValue;
  }

  // GitHub Pages/file:// can force the Google CSV request to fail CORS, which means
  // the OpenSheet fallback may return only populated cells and lose the true Column N index.
  // In that case, pick the cell that looks most like the article/writeup text.
  const usableCells = cells.filter(cell => cell && !looksLikeArticleHeaderOnly(cell));
  if (!usableCells.length) return "";

  const rankSection = usableCells.find(cell =>
    /^Rank\s*#?\s*\d+\s*:/i.test(firstNonEmptyLine(cell)) ||
    /^Rank\s*#?\s*\d+\s*:/i.test(cell)
  );
  if (rankSection) return rankSection;

  const matchupSection = usableCells.find(cell =>
    /\bvs\b/i.test(firstNonEmptyLine(cell)) && cell.length > 12
  );
  if (matchupSection) return matchupSection;

  const multilineWriteup = usableCells
    .filter(cell => /\r?\n/.test(cell) && /:/.test(cell))
    .sort((a, b) => b.length - a.length)[0];
  if (multilineWriteup) return multilineWriteup;

  const writerStyle = usableCells
    .filter(cell => /#\s*\d+\s*:/.test(cell) || /[A-Za-z0-9_ .-]+\s*:/.test(cell))
    .sort((a, b) => b.length - a.length)[0];
  if (writerStyle) return writerStyle;

  return usableCells.sort((a, b) => b.length - a.length)[0] || "";
}

function getFinalColumnNSections(finalRows) {
  return (finalRows || [])
    .map(row => getLikelyFinalArticleCell(row))
    .map(value => String(value || "").trim())
    .filter(value => value && !looksLikeArticleHeaderOnly(value));
}

function getFinalColumnNText(finalRows, index) {
  return getFinalColumnNSections(finalRows)[index] || "";
}

function textLooksLikeTeamSection(text, team, finalRank) {
  const value = String(text || "").trim();
  if (!value) return false;

  const firstLine = firstNonEmptyLine(value).toUpperCase();
  const teamKey = String(team || "").trim().toUpperCase();
  const rankKey = String(finalRank || "").trim().replace("#", "");

  if (!teamKey) return false;

  const hasTeam = firstLine.includes(teamKey) || value.toUpperCase().includes(`: ${teamKey}`);
  if (!hasTeam) return false;

  if (!rankKey) return true;

  const rankRegex = new RegExp(`\\bRANK\\s*#?\\s*${escapeRegex(rankKey)}\\b`, "i");
  const hashRegex = new RegExp(`(^|[^0-9])#${escapeRegex(rankKey)}([^0-9]|$)`, "i");

  return rankRegex.test(value) || hashRegex.test(firstLine) || hasTeam;
}

function getPowerFinalTextForTeam(finalRows, team, finalRank, displayIndex, totalTeams) {
  const sections = getFinalColumnNSections(finalRows);
  if (!sections.length) return "";

  const exactMatch = sections.find(text => textLooksLikeTeamSection(text, team, finalRank));
  if (exactMatch) return exactMatch;

  // Most Final tabs are built in display order, so use the rendered row order first.
  if (Number.isInteger(displayIndex) && sections[displayIndex]) {
    return sections[displayIndex];
  }

  const rankNumber = Number(String(finalRank || "").replace("#", ""));

  // Fallback for Final tabs ordered #1, #2, #3...
  if (Number.isFinite(rankNumber) && sections[rankNumber - 1]) {
    return sections[rankNumber - 1];
  }

  // Fallback for Final tabs ordered worst-to-best.
  const reverseRankIndex = Number.isFinite(rankNumber) && Number.isFinite(totalTeams)
    ? totalTeams - rankNumber
    : -1;

  if (reverseRankIndex >= 0 && sections[reverseRankIndex]) {
    return sections[reverseRankIndex];
  }

  return "";
}

function firstNonEmptyLine(text) {
  return String(text || "")
    .split(/\r?\n/)
    .map(line => line.trim())
    .find(Boolean) || "";
}

function finalTextToArticleHTML(text, emptyMessage = "No writeup found in the Final tab.") {
  const paragraphs = String(text || "")
    .split(/\r?\n+/)
    .map(line => line.trim())
    .filter(Boolean);

  if (!paragraphs.length) {
    return `<div class="article-empty">${escapeHTML(emptyMessage)}</div>`;
  }

  return paragraphs.map(paragraph => {
    const parsed = parseWriterBlurb(paragraph);

    if (parsed.writer && parsed.text) {
      const writerLabel = `${profileLinkedHTML(parsed.writer)}${parsed.rank ? ` #${escapeHTML(parsed.rank)}` : ""}`;

      return `
        <div class="article-blurb">
          <span class="article-writer">${writerLabel}:</span>
          ${profileLinkedHTML(parsed.text)}
        </div>
      `;
    }

    return `
      <div class="article-blurb">
        ${profileLinkedHTML(paragraph)}
      </div>
    `;
  }).join("");
}

function getPredictionYapTextFromRow(row) {
  const preferredIndexes = [13, 14, 15, 12, 11, 10, 9, 8];

  for (const index of preferredIndexes) {
    const value = getRawCell(row, index);
    if (value && !/^\s*#?N\/A\s*$/i.test(value) && !/^\bvs\b$/i.test(value)) {
      return value;
    }
  }

  const matchup = normalizeMatchupKey(findMatchupInRawRow(row));

  const candidates = (row || [])
    .map(cell => String(cell || "").trim())
    .filter(Boolean)
    .filter(cell => normalizeMatchupKey(cell) !== matchup)
    .filter(cell => !/^\d+(\.\d+)?$/.test(cell));

  return candidates[candidates.length - 1] || "";
}

function getKnownPredictionMatchups(matchData = []) {
  return (matchData || [])
    .map(row => getMatchupValue(row))
    .map(matchup => String(matchup || "").trim())
    .filter(Boolean)
    .map(matchup => ({
      raw: matchup,
      key: normalizeMatchupKey(matchup)
    }));
}

function findKnownPredictionMatchupInText(text, knownMatchups = []) {
  const value = String(text || "").trim();
  if (!value || !knownMatchups.length) return "";

  const normalizedText = normalizeMatchupKey(value);
  const firstLine = normalizeMatchupKey(firstNonEmptyLine(value));

  // First choice: the writeup starts with the matchup title, like "ATL VS HR".
  const firstLineMatch = knownMatchups.find(item =>
    firstLine === item.key ||
    firstLine.startsWith(item.key + " ") ||
    firstLine.startsWith(item.key + ":") ||
    firstLine.startsWith(item.key + " -")
  );
  if (firstLineMatch) return firstLineMatch.key;

  // Second choice: the matchup appears anywhere in the article cell.
  const anywhereMatch = knownMatchups.find(item =>
    normalizedText === item.key ||
    normalizedText.includes(item.key + " ") ||
    normalizedText.includes(item.key + ":") ||
    normalizedText.includes(item.key + " -") ||
    normalizedText.includes("\n" + item.key)
  );
  if (anywhereMatch) return anywhereMatch.key;

  return "";
}

function parsePredictionYapRows(finalRows, matchData = []) {
  const byMatchup = {};
  const knownMatchups = getKnownPredictionMatchups(matchData);
  const finalColumnNSections = getFinalColumnNSections(finalRows || []);

  const unmatchedTexts = [];

  finalColumnNSections.forEach(text => {
    const matchupKey = findKnownPredictionMatchupInText(text, knownMatchups);

    if (matchupKey) {
      byMatchup[matchupKey] = text;
    } else if (String(text || "").trim()) {
      unmatchedTexts.push(text);
    }
  });

  // Fallback only for cells that did NOT announce a matchup. This prevents
  // "ATL VS HR" yapping from landing under "BOO VS NAN" just because it was
  // the first row in Column N.
  let fallbackIndex = 0;
  knownMatchups.forEach(item => {
    if (byMatchup[item.key]) return;

    const fallbackText = unmatchedTexts[fallbackIndex] || "";
    if (fallbackText) {
      byMatchup[item.key] = fallbackText;
      fallbackIndex++;
    }
  });

  return byMatchup;
}

function predictionArticleTextToHTML(text) {
  return finalTextToArticleHTML(text, "No writeup found in Final column N for this matchup.");
}

function buildPredictionsArticleView(matchData, yapRows) {
  const gallery = document.getElementById("predictionArticleGallery");
  gallery.innerHTML = "";
  currentPredictionArticleTextBlocks = [];

  const toolbar = document.createElement("div");
  toolbar.className = "article-toolbar";
  toolbar.innerHTML = `
    <div>
      <div class="article-toolbar-title">Predictions Article View</div>
      <div class="article-toolbar-note">Graphic first, then the predictions Final tab writeup underneath.</div>
    </div>
    <button onclick="copyFullPredictionArticleText(this)">Copy Full Article</button>
  `;
  gallery.appendChild(toolbar);

  if (!Array.isArray(matchData) || !matchData.length) {
    const blank = document.createElement("div");
    blank.className = "blank-sheet-card";
    blank.innerHTML = `
      <div class="blank-sheet-title">Document Currently Blank</div>
      <div class="blank-sheet-text">Add matchups to the predictions sheet, then hit Update Graphics.</div>
    `;
    gallery.appendChild(blank);
    return;
  }

  const writeups = parsePredictionYapRows(yapRows || [], matchData);

  matchData.forEach((row, index) => {
    const matchup = getMatchupValue(row);
    if (!matchup) return;

    const parts = matchup.split(/vs/i).map(t => t.trim().toUpperCase());
    const a = parts[0];
    const b = parts[1];
    if (!a || !b) return;

    const key = normalizeMatchupKey(matchup);
    const writeup = writeups[key] || "No writeup found in the Final tab.";

    const copyText = `${a} VS ${b}\n${writeup}`.trim();
    const copyIndex = currentPredictionArticleTextBlocks.length;
    currentPredictionArticleTextBlocks.push(copyText);

    const item = document.createElement("div");
    item.className = "article-item";
    item.innerHTML = `
      ${buildPredictionArticleGraphicHTML(row, index)}

      <div class="article-text-box">
        <div class="article-rank-title">
          <span>${escapeHTML(a)} VS ${escapeHTML(b)}</span>
          <button class="article-copy-btn" onclick="copyPredictionArticleText(${copyIndex}, this)">Copy Text</button>
        </div>

        ${predictionArticleTextToHTML(writeup)}
      </div>
    `;

    gallery.appendChild(item);
  });
}

async function matchCardToDataURL(card) {
  if (!card) throw new Error("Prediction graphic card not found.");

  const canvas = await renderToCanvas(card);
  return canvas.toDataURL("image/png");
}

async function buildFullPredictionArticleClipboardHTML(button) {
  const articleItems = Array.from(document.querySelectorAll("#predictionArticleGallery .article-item"));

  if (!articleItems.length) {
    throw new Error("No prediction article sections found.");
  }

  const htmlBlocks = [];

  for (let i = 0; i < articleItems.length; i++) {
    if (button) button.innerText = `Copying ${i + 1}/${articleItems.length}...`;

    const item = articleItems[i];
    const card = item.querySelector(".match-card");
    const dataURL = await matchCardToDataURL(card);
    const text = currentPredictionArticleTextBlocks[i] || "";

    htmlBlocks.push(`
      <div style="margin:0 0 34px 0;page-break-inside:avoid;break-inside:avoid;">
        <img src="${dataURL}" alt="Prediction graphic" style="display:block;width:900px;max-width:100%;height:auto;margin:0 0 14px 0;border:0;">
        <div style="font-family:Arial,sans-serif;color:#111827;">
          ${articlePlainTextToHTML(text)}
        </div>
      </div>
    `);
  }

  return `
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="UTF-8">
      </head>
      <body>
        ${htmlBlocks.join("")}
      </body>
    </html>
  `;
}

async function copyFullPredictionArticleText(button) {
  const originalText = button ? button.innerText : "";

  try {
    const plainText = profileCleanedTextBlocks(currentPredictionArticleTextBlocks).join("\n\n").trim();

    if (!plainText) throw new Error("No prediction article text found.");

    const html = await buildFullPredictionArticleClipboardHTML(button);

    await navigator.clipboard.write([
      new ClipboardItem({
        "text/html": new Blob([html], { type: "text/html" }),
        "text/plain": new Blob([plainText], { type: "text/plain" })
      })
    ]);

    if (button) {
      button.innerText = "Copied!";
      setTimeout(() => button.innerText = originalText, 1200);
    }
  } catch (err) {
    console.error("PREDICTION ARTICLE COPY ERROR:", err);

    try {
      const plainText = profileCleanedTextBlocks(currentPredictionArticleTextBlocks).join("\n\n").trim();
      if (plainText) await navigator.clipboard.writeText(plainText);

      if (button) {
        button.innerText = "Copied Text Only";
        setTimeout(() => button.innerText = originalText, 1600);
      }

      alert("Full prediction article rich copy failed, so I copied the text only. Error: " + getSafeErrorMessage(err));
    } catch (fallbackErr) {
      if (button) button.innerText = originalText;
      alert("Prediction article copy failed: " + getSafeErrorMessage(err));
    }
  }
}

async function copyPredictionArticleText(index, button) {
  try {
    const text = currentPredictionArticleTextBlocks[index] || "";

    if (!text.trim()) throw new Error("No prediction article text found for this matchup.");

    await navigator.clipboard.writeText(profileCleanedPlainText(text));

    if (button) {
      const oldText = button.innerText;
      button.innerText = "Copied!";
      setTimeout(() => button.innerText = oldText, 1200);
    }
  } catch (err) {
    console.error("PREDICTION TEAM ARTICLE COPY ERROR:", err);
    alert("Prediction article copy failed: " + getSafeErrorMessage(err));
  }
}

function getMatchupValue(row) {
  const keys = Object.keys(row);

  const preferredHeaders = [
    "Matchup",
    "Matchups",
    "Match Up",
    "Match-Up",
    "Match",
    "Game",
    "Column 1"
  ];

  for (const wantedHeader of preferredHeaders) {
    const actualKey = keys.find(k =>
      String(k || "").trim().toLowerCase() === wantedHeader.toLowerCase()
    );

    if (actualKey) {
      const value = String(row[actualKey] || "").trim();

      if (value && /\bvs\b/i.test(value)) {
        return value;
      }
    }
  }

  const firstKey = keys[0];
  if (firstKey) {
    const firstValue = String(row[firstKey] || "").trim();

    if (firstValue && /\bvs\b/i.test(firstValue)) {
      return firstValue;
    }
  }

  const anyVsCell = Object.values(row).find(value =>
    /\bvs\b/i.test(String(value || ""))
  );

  return anyVsCell ? String(anyVsCell).trim() : "";
}
function setupTierDropdown() {
  const select = document.getElementById("tierSelect");

  const orderedTiers = TIER_ORDER.filter(tier => TIER_CONFIG[tier]);
  const extraTiers = Object.keys(TIER_CONFIG).filter(tier => !orderedTiers.includes(tier));
  const tiersToShow = [...orderedTiers, ...extraTiers];

  if (!TIER_CONFIG[currentTier]) {
    currentTier = tiersToShow[0] || currentTier;
  }

  select.innerHTML = tiersToShow
    .map(tier => `<option value="${tier}">${tier}</option>`)
    .join("");

  select.value = currentTier;

  select.onchange = () => {
    loadTier(select.value, true);
  };
}

function hasAnyColumn(rows, possibleNames) {
  if (!Array.isArray(rows) || !rows.length) return false;

  const keys = Object.keys(rows[0] || {}).map(normalizeHeaderName);
  return possibleNames.some(name => keys.includes(normalizeHeaderName(name)));
}

function countRowsWithValue(rows, getter) {
  return (rows || []).filter(row => String(getter(row) || "").trim()).length;
}

function validateRankData(rankData) {
  const warnings = [];

  if (!Array.isArray(rankData) || !rankData.length) {
    return { ok: false, warnings: ["No power rankings rows loaded."] };
  }

  if (!hasAnyColumn(rankData, ["Teams", "Team"])) {
    warnings.push("Power Rankings is missing a Teams/Team column.");
  }

  if (!hasAnyColumn(rankData, POWER_COLUMN_HEADERS.currentRank)) {
    warnings.push("Power Rankings is missing a Final Rank/Rank column.");
  }

  const teamCount = countRowsWithValue(rankData, row => row.Teams || row.Team);
  const rankedCount = countRowsWithValue(rankData, row => getPowerColumn(row, POWER_COLUMN_HEADERS.currentRank));

  if (!teamCount) warnings.push("Power Rankings has no usable team names.");
  if (!rankedCount) warnings.push("Power Rankings has no usable rank values.");

  return { ok: warnings.length === 0, warnings, teamCount, rankedCount };
}

function validateMatchData(matchData) {
  const warnings = [];

  if (!Array.isArray(matchData) || !matchData.length) {
    return { ok: true, warnings: ["Predictions sheet is blank or unavailable."], matchupCount: 0 };
  }

  const matchupCount = countRowsWithValue(matchData, getMatchupValue);
  if (!matchupCount) warnings.push("Predictions loaded, but no matchup values like Team A vs Team B were found.");

  return { ok: warnings.length === 0, warnings, matchupCount };
}

function validateTeamRecordData(teamRecordData) {
  const warnings = [];

  if (!Array.isArray(teamRecordData) || !teamRecordData.length) {
    return {
      ok: true,
      warnings: ["Team records sheet is blank or unavailable. Power cards will not show current records."],
      recordRows: 0,
      recordsWithFranchise: 0,
      recordsWithRecord: 0
    };
  }

  const recordsWithFranchise = countRowsWithValue(teamRecordData, row => getTeamRecordFromRow(row).franchise);
  const recordsWithRecord = countRowsWithValue(teamRecordData, row => getTeamRecordFromRow(row).record);

  if (!recordsWithFranchise) warnings.push("Team records loaded, but no franchise values were found in column A.");
  if (!recordsWithRecord) warnings.push("Team records loaded, but no record values were found in column C.");

  return {
    ok: warnings.length === 0,
    warnings,
    recordRows: teamRecordData.length,
    recordsWithFranchise,
    recordsWithRecord
  };
}

function countFinalColumnNRows(rows) {
  return getFinalColumnNSections(rows || []).length;
}

function buildDebugReport(tierName, rankData, matchData, finalRows, yapRows, teamRecordData = [], savedAt) {
  const rankValidation = validateRankData(rankData);
  const matchValidation = validateMatchData(matchData);
  const teamRecordValidation = validateTeamRecordData(teamRecordData);
  const warnings = [...rankValidation.warnings, ...matchValidation.warnings, ...teamRecordValidation.warnings];

  if (!playerProfileNames.length) {
    warnings.push("No player names loaded from the hyperlink sheet, so dash-marker hyperlinks will not render.");
  }

  return {
    tierName,
    source: savedAt ? `Saved cache from ${formatSavedTime(savedAt)}` : "Fresh sheet pull",
    rankRows: Array.isArray(rankData) ? rankData.length : 0,
    rankedTeams: rankValidation.rankedCount || 0,
    matchRows: Array.isArray(matchData) ? matchData.length : 0,
    matchups: matchValidation.matchupCount || 0,
    recordRows: teamRecordValidation.recordRows || 0,
    recordsFound: teamRecordValidation.recordsWithRecord || 0,
    powerArticleRows: countFinalColumnNRows(finalRows),
    predictionArticleRows: countFinalColumnNRows(yapRows),
    profileNames: playerProfileNames.length || 0,
    warnings
  };
}

function updateDebugPanel(report) {
  const panel = document.getElementById("debugPanel");
  if (!panel || !report) return;

  const warningHTML = report.warnings.length
    ? `<div class="debug-warnings">${report.warnings.map(w => `<div>⚠ ${escapeHTML(w)}</div>`).join("")}</div>`
    : `<div class="debug-ok">No data warnings.</div>`;

  panel.innerHTML = `
    <div class="debug-title">Load Debug</div>
    <div class="debug-grid">
      <div><span>Tier</span><strong>${escapeHTML(report.tierName)}</strong></div>
      <div><span>Source</span><strong>${escapeHTML(report.source)}</strong></div>
      <div><span>Rank rows</span><strong>${report.rankRows}</strong></div>
      <div><span>Ranked teams</span><strong>${report.rankedTeams}</strong></div>
      <div><span>Prediction rows</span><strong>${report.matchRows}</strong></div>
      <div><span>Matchups</span><strong>${report.matchups}</strong></div>
      <div><span>Record rows</span><strong>${report.recordRows}</strong></div>
      <div><span>Records found</span><strong>${report.recordsFound}</strong></div>
      <div><span>Power article rows</span><strong>${report.powerArticleRows}</strong></div>
      <div><span>Prediction article rows</span><strong>${report.predictionArticleRows}</strong></div>
      <div><span>Profile names</span><strong>${report.profileNames || 0}</strong></div>
    </div>
    ${warningHTML}
  `;
}

function renderTierData(tierName, rankData, matchData, finalRows = [], yapRows = [], teamRecordData = [], savedAt) {
  const debugReport = buildDebugReport(tierName, rankData, matchData, finalRows, yapRows, teamRecordData, savedAt);
  updateDebugPanel(debugReport);

  buildPowerCards(rankData, teamRecordData);
  buildMatchCards(matchData);
  buildArticleView(rankData, finalRows, teamRecordData);
  buildPredictionsArticleView(matchData, yapRows);

  const message = savedAt
    ? `${tierName} graphics loaded from saved data. Last updated: ${formatSavedTime(savedAt)}.`
    : `${tierName} graphics loaded.`;

  const warningNote = debugReport.warnings.length
    ? ` ${debugReport.warnings.length} warning${debugReport.warnings.length === 1 ? "" : "s"} shown in Load Debug.`
    : "";

  document.getElementById("status").innerText = message + warningNote;
}
function getFinalRankNumber(row) {
  const raw = String(row["Final Rank"] || "")
    .trim()
    .replace("#", "");

  const num = Number(raw);

  return Number.isFinite(num) ? num : 9999;
}
async function loadTier(tierName, forceFresh = false) {
  try {
    currentTier = tierName;

    const select = document.getElementById("tierSelect");
    if (select) select.value = tierName;

    document.getElementById("powerGallery").innerHTML = "";
    document.getElementById("matchGallery").innerHTML = "";
    document.getElementById("articleGallery").innerHTML = "";
    document.getElementById("predictionArticleGallery").innerHTML = "";

    if (!forceFresh) {
      const saved = getSavedTierData(tierName);

      if (saved) {
        document.getElementById("status").innerText = `Loading saved ${tierName} graphics...`;
await prepareTeamAssets(saved.rankData, saved.matchData);
await loadPlayerProfileNames();

renderTierData(tierName, saved.rankData, saved.matchData, saved.finalRows || [], saved.yapRows || [], saved.teamRecordData || [], saved.savedAt);
return;
      }
    }

    const tier = TIER_CONFIG[tierName];

    if (!tier) {
      throw new Error(`Tier not found: ${tierName}`);
    }

    document.getElementById("status").innerText = `Pulling fresh ${tierName} graphics from sheets...`;

    const rankingsURL = openSheetURL(tier.rankingsId, POWER_TAB);
const picksURL = openSheetURL(tier.picksId, PICKS_TAB);

const rankData = await fetchSheetData(
  rankingsURL,
  `${tierName} Power Rankings`
);

let matchData = [];
let finalRows = [];
let yapRows = [];
let teamRecordData = [];

try {
  matchData = await fetchSheetData(
    picksURL,
    `${tierName} Predictions`
  );
} catch (err) {
  console.warn(`${tierName} predictions failed to load. Showing blank matchups page.`, err);
  matchData = [];
}

try {
  finalRows = await fetchRawSheetRows(
    tier.rankingsId,
    FINAL_TAB,
    `${tierName} Final Article Text`
  );
} catch (err) {
  console.warn(`${tierName} Final tab failed to load. Article text will be blank.`, err);
  finalRows = [];
}

try {
  yapRows = await fetchRawSheetRows(
    tier.picksId,
    PREDICTION_FINAL_TAB,
    `${tierName} Prediction Final Text`
  );
} catch (err) {
  console.warn(`${tierName} Prediction Final tab failed to load. Prediction article text will be blank.`, err);
  yapRows = [];
}

try {
  teamRecordData = await fetchTeamRecordData(tierName);
} catch (err) {
  console.warn(`${tierName} team records failed to load. Record boxes will be blank.`, err);
  teamRecordData = [];
}

    document.getElementById("status").innerText = `Loading ${tierName} logos...`;

await prepareTeamAssets(rankData, matchData);
await loadPlayerProfileNames();

saveTierData(tierName, rankData, matchData, finalRows, yapRows, teamRecordData);
renderTierData(tierName, rankData, matchData, finalRows, yapRows, teamRecordData, new Date().toISOString());
  } catch (err) {
    console.error("LOAD ERROR:", err);
    document.getElementById("status").innerText = `Failed to load ${tierName}: ${getSafeErrorMessage(err)}`;
  }
}

async function updateGraphicsFromSheets() {
  try {
    document.getElementById("status").innerText = "Refreshing tier config...";

    await loadTierConfig();

    setupTierDropdown();

    await loadTier(currentTier, true);
  } catch (err) {
    console.error("UPDATE ERROR:", err);
    document.getElementById("status").innerText = `Update failed: ${getSafeErrorMessage(err)}`;
  }
}

async function init() {
  try {
    document.getElementById("status").innerText = "Loading tier config...";

    await loadTierConfig();

    setupTierDropdown();

    await loadTier(currentTier, false);
  } catch (err) {
    console.error("CONFIG LOAD ERROR:", err);
    document.getElementById("status").innerText = `Failed to load tier config: ${getSafeErrorMessage(err)}`;
  }
}


/* =========================
   CSC MEDIA HUB V2 OVERRIDES
========================= */

let currentView = "article";
let currentDebugReport = null;
let commandIndex = 0;

function normalizeWriterName(value) {
  return String(value || "")
    .replace(/\s*#\s*\d+(?:\.\d+)?\s*:?\s*$/i, "")
    .replace(/:\s*$/, "")
    .trim();
}

function getLogoCandidates(team) {
  const raw = String(team || "").trim();
  const upper = raw.toUpperCase();
  const lower = raw.toLowerCase();
  const folderCandidates = [LOGO_FOLDER, "Logos", "logos"];
  const keyCandidates = [lower, upper, raw].filter(Boolean);
  const extCandidates = [LOGO_EXTENSION, String(LOGO_EXTENSION).toUpperCase(), "png", "PNG"];
  const seen = new Set();
  const candidates = [];

  folderCandidates.forEach(folder => {
    keyCandidates.forEach(key => {
      extCandidates.forEach(ext => {
        const path = `${folder}/${key}.${ext}`;
        if (!seen.has(path)) {
          seen.add(path);
          candidates.push(path);
        }
      });
    });
  });

  return candidates;
}

function localLogoPath(team) {
  return getLogoCandidates(team)[0];
}

function handleLogoError(img) {
  if (!img) return;

  const key = img.getAttribute("data-logo-key") || img.getAttribute("alt") || "";
  const fallbackClass = img.getAttribute("data-fallback-class") || "power-logo-fallback";
  const candidates = getLogoCandidates(key);
  const currentAttempt = Number(img.getAttribute("data-logo-attempt") || "0");
  const nextAttempt = currentAttempt + 1;

  if (nextAttempt < candidates.length) {
    img.setAttribute("data-logo-attempt", String(nextAttempt));
    img.src = candidates[nextAttempt];
    return;
  }

  const cleanKey = escapeHTML(String(key || "LOGO").toUpperCase());
  img.outerHTML = `<div class="${fallbackClass}">${cleanKey}</div>`;
}

function logo(team, map, fallbackClass = "power-logo-fallback") {
  const key = String(team || "").trim().toUpperCase();
  const src = localLogoPath(key);

  return `
    <img
      src="${src}"
      data-logo-key="${escapeHTML(key)}"
      data-logo-attempt="0"
      data-fallback-class="${escapeHTML(fallbackClass)}"
      alt="${escapeHTML(key)} logo"
      onerror="handleLogoError(this)"
    >
  `;
}

function finalTextToArticleHTML(text, emptyMessage = "No writeup found in the Final tab.") {
  const paragraphs = String(text || "")
    .split(/\r?\n+/)
    .map(line => line.trim())
    .filter(Boolean);

  if (!paragraphs.length) {
    return `<div class="article-empty">${escapeHTML(emptyMessage)}</div>`;
  }

  return paragraphs.map(paragraph => {
    const parsed = parseWriterBlurb(paragraph);

    if (parsed.writer && parsed.text) {
      const writer = normalizeWriterName(profileCleanedPlainText(parsed.writer));
      const writerLabel = `${profileLinkedHTML(parsed.writer)}${parsed.rank ? ` #${escapeHTML(parsed.rank)}` : ""}`;

      return `
        <div class="article-blurb" data-writer="${escapeHTML(writer)}">
          <span class="article-writer">${writerLabel}:</span>
          ${profileLinkedHTML(parsed.text)}
        </div>
      `;
    }

    return `
      <div class="article-blurb" data-writer="">
        ${profileLinkedHTML(paragraph)}
      </div>
    `;
  }).join("");
}

function setView(view) {
  currentView = view === "predictionArticle" ? "predictionArticle" : "article";

  const articlePage = document.getElementById("article-page");
  const predictionPage = document.getElementById("predictionArticle-page");
  const powerNav = document.getElementById("navPower");
  const predictionsNav = document.getElementById("navPredictions");
  const publishPowerNav = document.getElementById("publishNavPower");
  const publishPredictionsNav = document.getElementById("publishNavPredictions");
  const viewTitle = document.getElementById("viewTitle");

  if (articlePage) articlePage.classList.toggle("active", currentView === "article");
  if (predictionPage) predictionPage.classList.toggle("active", currentView === "predictionArticle");
  if (powerNav) powerNav.classList.toggle("active", currentView === "article");
  if (predictionsNav) predictionsNav.classList.toggle("active", currentView === "predictionArticle");
  if (publishPowerNav) publishPowerNav.classList.toggle("active", currentView === "article");
  if (publishPredictionsNav) publishPredictionsNav.classList.toggle("active", currentView === "predictionArticle");
  if (viewTitle) viewTitle.innerText = currentView === "article" ? "Power Rankings" : "Predictions";

  saveUiPrefs({ currentView });
  populateWriterFilter();
  applyArticleFilters();
  closeAllSectionMenus();
  window.setTimeout(resizeResponsiveGraphics, 80);
}

function getCurrentGalleryId() {
  return currentView === "article" ? "articleGallery" : "predictionArticleGallery";
}

function getCurrentGallery() {
  return document.getElementById(getCurrentGalleryId());
}

function getCurrentTextBlocks() {
  return currentView === "article" ? currentArticleTextBlocks : currentPredictionArticleTextBlocks;
}

function createSectionHeader({ kicker, title, meta, copyIndex }) {
  return `
    <div class="article-section-header" onclick="toggleArticleItem(this)">
      <div>
        <div class="section-kicker">${escapeHTML(kicker)}</div>
        <div class="section-title">${escapeHTML(title)}</div>
        ${meta ? `<div class="section-meta">${escapeHTML(meta)}</div>` : ""}
      </div>
      <div class="section-header-actions" onclick="event.stopPropagation()">
        <span class="match-count-pill">${escapeHTML(currentTier)}</span>
        <button class="article-section-toggle" onclick="toggleArticleItem(this.closest('.article-section-header'))">⌄</button>
        <button class="section-menu-trigger" onclick="toggleSectionMenu(this)">⋮</button>
        <div class="section-menu">
          <button onclick="copySectionGraphic(this)">Copy Graphic</button>
          <button onclick="downloadSectionGraphic(this)">Download PNG</button>
          <button onclick="copySectionText(${copyIndex}, this)">Copy Text</button>
          <button onclick="expandThisSection(this)">Expand Only This</button>
        </div>
      </div>
    </div>
  `;
}

function buildArticleView(rankData, finalRows, teamRecordData = []) {
  const gallery = document.getElementById("articleGallery");
  gallery.innerHTML = "";
  currentArticleTextBlocks = [];

  const sortedData = getSortedPowerData(rankData).reverse();
  const teamRecordMap = buildTeamRecordMap(teamRecordData, rankData);

  sortedData.forEach((row, index) => {
    const team = String(row.Teams || row.Team || "").trim().toUpperCase();
    if (!team) return;

    const finalRank = cleanRankValue(getPowerColumn(row, POWER_COLUMN_HEADERS.currentRank));
    const finalText = getPowerFinalTextForTeam(finalRows || [], team, finalRank, index, sortedData.length);
    const cleanFinalText = removeRankTitleLine(finalText);
    const fallbackText = `Rank ${finalRank || ""}: ${team}\nNo writeups found for this team after checking Final column N.`.trim();
    const copyText = cleanFinalText || fallbackText;
    const copyIndex = currentArticleTextBlocks.length;
    currentArticleTextBlocks.push(copyText);

    const sectionTitle = finalRank ? `#${finalRank} ${team}` : team;
    const titleLine = firstNonEmptyLine(finalText) || `Rank ${finalRank}: ${team}`;
    const blurbsHTML = finalTextToArticleHTML(cleanFinalText, `No writeups found for ${team} after checking Final column N.`);

    const item = document.createElement("div");
    item.className = "article-item";
    item.dataset.search = `${team} ${titleLine} ${copyText}`.toLowerCase();
    item.dataset.copyIndex = String(copyIndex);
    item.innerHTML = `
      ${createSectionHeader({
        kicker: "Power Ranking",
        title: sectionTitle,
        meta: titleLine,
        copyIndex
      })}

      <div class="article-section-body">
        ${buildPowerArticleGraphicHTML(row, index, teamRecordMap)}

        <div class="article-text-box">
          ${blurbsHTML}
        </div>
      </div>
    `;

    gallery.appendChild(item);
  });

  afterArticleRender();
}

function buildPredictionsArticleView(matchData, yapRows) {
  const gallery = document.getElementById("predictionArticleGallery");
  gallery.innerHTML = "";
  currentPredictionArticleTextBlocks = [];

  if (!Array.isArray(matchData) || !matchData.length) {
    const blank = document.createElement("div");
    blank.className = "blank-sheet-card";
    blank.innerHTML = `
      <div class="blank-sheet-title">Document Currently Blank</div>
      <div class="blank-sheet-text">Add matchups to the predictions sheet, then refresh data.</div>
    `;
    gallery.appendChild(blank);
    afterArticleRender();
    return;
  }

  const writeups = parsePredictionYapRows(yapRows || [], matchData);

  matchData.forEach((row, index) => {
    const matchup = getMatchupValue(row);
    if (!matchup) return;

    const parts = matchup.split(/vs/i).map(t => t.trim().toUpperCase());
    const a = parts[0];
    const b = parts[1];
    if (!a || !b) return;

    const key = normalizeMatchupKey(matchup);
    const writeup = writeups[key] || "No writeup found in the Final tab.";
    const copyText = `${a} VS ${b}\n${writeup}`.trim();
    const copyIndex = currentPredictionArticleTextBlocks.length;
    currentPredictionArticleTextBlocks.push(copyText);

    const item = document.createElement("div");
    item.className = "article-item";
    item.dataset.search = `${a} ${b} ${a} vs ${b} ${writeup}`.toLowerCase();
    item.dataset.copyIndex = String(copyIndex);
    item.innerHTML = `
      ${createSectionHeader({
        kicker: "Prediction",
        title: `${a} VS ${b}`,
        meta: "Matchup article section",
        copyIndex
      })}

      <div class="article-section-body">
        ${buildPredictionArticleGraphicHTML(row, index)}

        <div class="article-text-box">
          ${predictionArticleTextToHTML(writeup)}
        </div>
      </div>
    `;

    gallery.appendChild(item);
  });

  afterArticleRender();
}

function afterArticleRender() {
  setTimeout(() => {
    resizeTeamNames();
    resizeAnalysts();
    populateWriterFilter();
    applyArticleFilters();
  }, 150);
}

function toggleArticleItem(header) {
  const item = header?.closest?.(".article-item");
  if (!item) return;

  item.classList.toggle("collapsed");
}

function expandThisSection(button) {
  const item = button?.closest?.(".article-item");
  const gallery = item?.closest?.(".article-gallery");
  if (!item || !gallery) return;

  gallery.querySelectorAll(".article-item").forEach(other => {
    other.classList.toggle("collapsed", other !== item);
  });

  closeAllSectionMenus();
  item.scrollIntoView({ behavior: "smooth", block: "start" });
}

function expandVisibleSections() {
  const gallery = getCurrentGallery();
  gallery?.querySelectorAll(".article-item:not(.filtered-out)").forEach(item => item.classList.remove("collapsed"));
}

function collapseVisibleSections() {
  const gallery = getCurrentGallery();
  gallery?.querySelectorAll(".article-item:not(.filtered-out)").forEach(item => item.classList.add("collapsed"));
}

function toggleSectionMenu(button) {
  const menu = button?.parentElement?.querySelector?.(".section-menu");
  if (!menu) return;
  const shouldOpen = !menu.classList.contains("open");
  closeAllSectionMenus();
  menu.classList.toggle("open", shouldOpen);
}

function closeAllSectionMenus() {
  document.querySelectorAll(".section-menu.open").forEach(menu => menu.classList.remove("open"));
}

function getSectionCardFromButton(button) {
  const item = button?.closest?.(".article-item");
  if (!item) throw new Error("Article section not found.");

  const card = item.querySelector(".power-card, .match-card");
  if (!card) throw new Error("Graphic card not found in this section.");

  return { item, card };
}

async function withItemExpanded(item, callback) {
  const wasCollapsed = item.classList.contains("collapsed");
  if (wasCollapsed) item.classList.remove("collapsed");

  try {
    return await callback();
  } finally {
    if (wasCollapsed) item.classList.add("collapsed");
  }
}

async function copySectionGraphic(button) {
  const oldText = button ? button.innerText : "";

  try {
    const { item, card } = getSectionCardFromButton(button);
    closeAllSectionMenus();

    await withItemExpanded(item, async () => {
      await copyCard(card.id, button);
    });
  } catch (err) {
    if (button) button.innerText = oldText;
    alert("Graphic copy failed: " + getSafeErrorMessage(err));
  }
}

async function downloadSectionGraphic(button) {
  const oldText = button ? button.innerText : "";

  try {
    const { item, card } = getSectionCardFromButton(button);
    closeAllSectionMenus();
    if (button) button.innerText = "Rendering...";

    await withItemExpanded(item, async () => {
      const canvas = await renderToCanvas(card);
      downloadCanvasPNG(canvas, `${card.id}.png`);
    });

    if (button) {
      button.innerText = "Downloaded";
      setTimeout(() => button.innerText = oldText, 1300);
    }
  } catch (err) {
    if (button) button.innerText = oldText;
    alert("PNG download failed: " + getSafeErrorMessage(err));
  }
}

async function copySectionText(index, button) {
  closeAllSectionMenus();
  return currentView === "article"
    ? copyTeamArticleText(index, button)
    : copyPredictionArticleText(index, button);
}

function collectWritersFromCurrentView() {
  const gallery = getCurrentGallery();
  const writers = new Set();

  gallery?.querySelectorAll(".article-blurb[data-writer]").forEach(blurb => {
    const writer = normalizeWriterName(blurb.dataset.writer || "");
    if (writer) writers.add(writer);
  });

  return Array.from(writers).sort((a, b) => a.localeCompare(b));
}

function populateWriterFilter() {
  const select = document.getElementById("writerFilter");
  if (!select) return;

  const oldValue = select.value;
  const writers = collectWritersFromCurrentView();

  select.innerHTML = `<option value="">All writers</option>` + writers
    .map(writer => `<option value="${escapeHTML(writer)}">${escapeHTML(writer)}</option>`)
    .join("");

  if (writers.includes(oldValue)) select.value = oldValue;
}

function applyArticleFilters() {
  const gallery = getCurrentGallery();
  const search = String(document.getElementById("articleSearch")?.value || "").trim().toLowerCase();
  const writer = String(document.getElementById("writerFilter")?.value || "").trim();
  let visibleCount = 0;

  gallery?.querySelectorAll(".article-item").forEach(item => {
    const searchBlob = String(item.dataset.search || item.innerText || "").toLowerCase();
    const searchMatch = !search || searchBlob.includes(search);
    item.classList.toggle("filtered-out", !searchMatch);
    if (searchMatch) visibleCount++;

    item.querySelectorAll(".article-blurb[data-writer]").forEach(blurb => {
      const blurbWriter = normalizeWriterName(blurb.dataset.writer || "");
      const writerMatch = !writer || blurbWriter === writer;
      blurb.dataset.writerHidden = writerMatch ? "false" : "true";
      blurb.classList.toggle("writer-focused", Boolean(writer && writerMatch));
    });
  });

  updateFilterStatus(visibleCount, search, writer);
}

function updateFilterStatus(visibleCount, search, writer) {
  const status = document.getElementById("status");
  if (!status || !currentDebugReport) return;

  const base = currentDebugReport.savedAt
    ? `${currentTier} loaded from saved data.`
    : `${currentTier} loaded.`;

  const pieces = [];
  if (search) pieces.push(`search: ${search}`);
  if (writer) pieces.push(`writer: ${writer}`);

  if (pieces.length) {
    status.innerText = `${base} Showing ${visibleCount} section${visibleCount === 1 ? "" : "s"} for ${pieces.join(" + ")}.`;
  }
}

function updateOverviewCards(report) {
  // V2.4: top overview stat cards were removed to keep the article view clean.
}

function renderTierData(tierName, rankData, matchData, finalRows = [], yapRows = [], teamRecordData = [], savedAt) {
  const debugReport = buildDebugReport(tierName, rankData, matchData, finalRows, yapRows, teamRecordData, savedAt);
  debugReport.savedAt = savedAt || "";
  currentDebugReport = debugReport;

  updateDebugPanel(debugReport);
  updateOverviewCards(debugReport);

  buildArticleView(rankData, finalRows, teamRecordData);
  buildPredictionsArticleView(matchData, yapRows);

  const message = savedAt
    ? `${tierName} loaded from saved data. Last updated: ${formatSavedTime(savedAt)}.`
    : `${tierName} loaded.`;

  const warningNote = debugReport.warnings.length
    ? ` ${debugReport.warnings.length} warning${debugReport.warnings.length === 1 ? "" : "s"} in Load Debug.`
    : "";

  document.getElementById("status").innerText = message + warningNote;
  setView(currentView);
}

function setupTierDropdown() {
  const select = document.getElementById("tierSelect");
  if (!select) return;

  const orderedTiers = TIER_ORDER.filter(tier => TIER_CONFIG[tier]);
  const extraTiers = Object.keys(TIER_CONFIG).filter(tier => !orderedTiers.includes(tier));
  const tiersToShow = [...orderedTiers, ...extraTiers];

  if (!TIER_CONFIG[currentTier]) {
    currentTier = tiersToShow[0] || currentTier;
  }

  select.innerHTML = tiersToShow
    .map(tier => `<option value="${escapeHTML(tier)}">${escapeHTML(tier)}</option>`)
    .join("");

  select.value = currentTier;

  select.onchange = () => {
    currentTier = select.value;
    loadTier(select.value, true);
  };
}

async function withExpandedGallery(galleryId, callback) {
  const gallery = document.getElementById(galleryId);
  const collapsed = gallery ? Array.from(gallery.querySelectorAll(".article-item.collapsed")) : [];
  collapsed.forEach(item => item.classList.remove("collapsed"));

  try {
    return await callback();
  } finally {
    collapsed.forEach(item => item.classList.add("collapsed"));
  }
}

async function copyFullArticleText(button) {
  const originalText = button ? button.innerText : "";

  try {
    const plainText = profileCleanedTextBlocks(currentArticleTextBlocks).join("\n\n").trim();
    if (!plainText) throw new Error("No article text found.");

    const html = await withExpandedGallery("articleGallery", () => buildFullArticleClipboardHTML(button));

    await navigator.clipboard.write([
      new ClipboardItem({
        "text/html": new Blob([html], { type: "text/html" }),
        "text/plain": new Blob([plainText], { type: "text/plain" })
      })
    ]);

    if (button) {
      button.innerText = "Copied!";
      setTimeout(() => button.innerText = originalText, 1200);
    }
  } catch (err) {
    console.error("ARTICLE COPY ERROR:", err);

    try {
      const plainText = profileCleanedTextBlocks(currentArticleTextBlocks).join("\n\n").trim();
      if (plainText) await navigator.clipboard.writeText(plainText);

      if (button) {
        button.innerText = "Copied Text Only";
        setTimeout(() => button.innerText = originalText, 1600);
      }

      alert("Full article rich copy failed, so I copied the text only. Error: " + getSafeErrorMessage(err));
    } catch (fallbackErr) {
      if (button) button.innerText = originalText;
      alert("Article copy failed: " + getSafeErrorMessage(err));
    }
  }
}

async function copyFullPredictionArticleText(button) {
  const originalText = button ? button.innerText : "";

  try {
    const plainText = profileCleanedTextBlocks(currentPredictionArticleTextBlocks).join("\n\n").trim();
    if (!plainText) throw new Error("No prediction article text found.");

    const html = await withExpandedGallery("predictionArticleGallery", () => buildFullPredictionArticleClipboardHTML(button));

    await navigator.clipboard.write([
      new ClipboardItem({
        "text/html": new Blob([html], { type: "text/html" }),
        "text/plain": new Blob([plainText], { type: "text/plain" })
      })
    ]);

    if (button) {
      button.innerText = "Copied!";
      setTimeout(() => button.innerText = originalText, 1200);
    }
  } catch (err) {
    console.error("PREDICTION ARTICLE COPY ERROR:", err);

    try {
      const plainText = profileCleanedTextBlocks(currentPredictionArticleTextBlocks).join("\n\n").trim();
      if (plainText) await navigator.clipboard.writeText(plainText);

      if (button) {
        button.innerText = "Copied Text Only";
        setTimeout(() => button.innerText = originalText, 1600);
      }

      alert("Full prediction article rich copy failed, so I copied the text only. Error: " + getSafeErrorMessage(err));
    } catch (fallbackErr) {
      if (button) button.innerText = originalText;
      alert("Prediction article copy failed: " + getSafeErrorMessage(err));
    }
  }
}

async function runSelectedArticleAction(button) {
  const select = document.getElementById("copyActionSelect");
  const action = select ? select.value : "rich";

  if (action === "text") return copyCurrentArticleTextOnly(button);
  if (action === "discord") return copyCurrentDiscordArticle(button);
  if (action === "html") return exportCurrentArticleHTML(button);

  return copyCurrentArticleRich(button);
}

async function copyCurrentArticleRich(button) {
  return currentView === "article"
    ? copyFullArticleText(button)
    : copyFullPredictionArticleText(button);
}

async function copyCurrentArticleTextOnly(button) {
  const oldText = button ? button.innerText : "";

  try {
    const text = profileCleanedTextBlocks(getCurrentTextBlocks()).join("\n\n").trim();
    if (!text) throw new Error("No article text found.");

    await navigator.clipboard.writeText(profileCleanedPlainText(text));

    if (button) {
      button.innerText = "Copied!";
      setTimeout(() => button.innerText = oldText, 1200);
    }
  } catch (err) {
    if (button) button.innerText = oldText;
    alert("Text copy failed: " + getSafeErrorMessage(err));
  }
}

function buildDiscordPlainText() {
  const title = currentView === "article" ? `${currentTier} Power Rankings` : `${currentTier} Predictions`;
  const divider = "==============================";
  return `${title}\n${divider}\n\n${profileCleanedTextBlocks(getCurrentTextBlocks()).join("\n\n")}`.trim();
}

async function copyCurrentDiscordArticle(button) {
  const oldText = button ? button.innerText : "";

  try {
    const text = buildDiscordPlainText();
    if (!text) throw new Error("No article text found.");

    await navigator.clipboard.writeText(profileCleanedPlainText(text));

    if (button) {
      button.innerText = "Copied!";
      setTimeout(() => button.innerText = oldText, 1200);
    }
  } catch (err) {
    if (button) button.innerText = oldText;
    alert("Discord copy failed: " + getSafeErrorMessage(err));
  }
}

function downloadTextFile(filename, text, mime = "text/plain") {
  const blob = new Blob([text], { type: mime });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

async function exportCurrentArticleHTML(button) {
  const oldText = button ? button.innerText : "";

  try {
    if (button) button.innerText = "Exporting...";

    const html = currentView === "article"
      ? await withExpandedGallery("articleGallery", () => buildFullArticleClipboardHTML(button))
      : await withExpandedGallery("predictionArticleGallery", () => buildFullPredictionArticleClipboardHTML(button));

    const fileName = `${safeName(currentTier)}_${currentView === "article" ? "power_rankings" : "predictions"}_article.html`;
    downloadTextFile(fileName, html, "text/html");

    if (button) {
      button.innerText = "Exported";
      setTimeout(() => button.innerText = oldText, 1400);
    }
  } catch (err) {
    console.error("EXPORT HTML ERROR:", err);
    const fallback = `<!DOCTYPE html><html><head><meta charset="UTF-8"></head><body><pre>${escapeHTML(getCurrentTextBlocks().join("\n\n"))}</pre></body></html>`;
    downloadTextFile(`${safeName(currentTier)}_${currentView}_article_text_only.html`, fallback, "text/html");

    if (button) {
      button.innerText = "Exported Text Only";
      setTimeout(() => button.innerText = oldText, 1600);
    }
  }
}

const commandDefinitions = [
  { label: "Refresh Data", hint: "Pull the current tier from sheets", action: () => updateGraphicsFromSheets() },
  { label: "Go to Power Rankings", hint: "Switch to Power Rankings article", action: () => setView("article") },
  { label: "Go to Predictions", hint: "Switch to Prediction article", action: () => setView("predictionArticle") },
  { label: "Copy Rich Article", hint: "Copy graphics and article text", action: btn => copyCurrentArticleRich(btn) },
  { label: "Copy Text Only", hint: "Copy all article text", action: btn => copyCurrentArticleTextOnly(btn) },
  { label: "Copy Discord Article", hint: "Copy Discord-friendly plain text", action: btn => copyCurrentDiscordArticle(btn) },
  { label: "Export HTML", hint: "Download the article as an HTML file", action: btn => exportCurrentArticleHTML(btn) },
  { label: "Expand Visible Sections", hint: "Open every visible section", action: () => expandVisibleSections() },
  { label: "Collapse Visible Sections", hint: "Close every visible section", action: () => collapseVisibleSections() }
];

function openCommandPalette() {
  const palette = document.getElementById("commandPalette");
  const input = document.getElementById("commandInput");
  if (!palette || !input) return;

  palette.classList.add("open");
  palette.setAttribute("aria-hidden", "false");
  input.value = "";
  commandIndex = 0;
  renderCommandPalette();
  setTimeout(() => input.focus(), 20);
}

function closeCommandPalette() {
  const palette = document.getElementById("commandPalette");
  if (!palette) return;
  palette.classList.remove("open");
  palette.setAttribute("aria-hidden", "true");
}

function getFilteredCommands() {
  const query = String(document.getElementById("commandInput")?.value || "").trim().toLowerCase();

  if (query.startsWith("search ")) {
    return [{
      label: `Search for “${query.slice(7).trim()}”`,
      hint: "Apply this to the article search box",
      action: () => {
        const input = document.getElementById("articleSearch");
        if (input) input.value = query.slice(7).trim();
        applyArticleFilters();
      }
    }];
  }

  return commandDefinitions.filter(command => {
    const blob = `${command.label} ${command.hint}`.toLowerCase();
    return !query || blob.includes(query);
  });
}

function renderCommandPalette() {
  const list = document.getElementById("commandList");
  if (!list) return;

  const commands = getFilteredCommands();
  commandIndex = Math.max(0, Math.min(commandIndex, commands.length - 1));

  list.innerHTML = commands.length
    ? commands.map((command, index) => `
      <button class="command-item ${index === commandIndex ? "active" : ""}" onclick="runCommand(${index}, this)">
        ${escapeHTML(command.label)}
        <span>${escapeHTML(command.hint)}</span>
      </button>
    `).join("")
    : `<div class="article-empty" style="padding:16px;">No commands found.</div>`;
}

async function runCommand(index, button) {
  const commands = getFilteredCommands();
  const command = commands[index];
  if (!command) return;

  closeCommandPalette();
  await command.action(button);
}

function handleCommandKeydown(event) {
  const palette = document.getElementById("commandPalette");
  const isOpen = palette?.classList.contains("open");

  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === "k") {
    event.preventDefault();
    openCommandPalette();
    return;
  }

  if (!isOpen) return;

  const commands = getFilteredCommands();

  if (event.key === "Escape") {
    event.preventDefault();
    closeCommandPalette();
  } else if (event.key === "ArrowDown") {
    event.preventDefault();
    commandIndex = Math.min(commandIndex + 1, commands.length - 1);
    renderCommandPalette();
  } else if (event.key === "ArrowUp") {
    event.preventDefault();
    commandIndex = Math.max(commandIndex - 1, 0);
    renderCommandPalette();
  } else if (event.key === "Enter") {
    event.preventDefault();
    runCommand(commandIndex);
  }
}

document.addEventListener("click", event => {
  if (!event.target.closest(".section-header-actions")) {
    closeAllSectionMenus();
  }

  if (event.target.id === "commandPalette") {
    closeCommandPalette();
  }
});

// V2.4: command palette keyboard shortcut disabled for a simpler UI.


function getVisibleArticleItems(galleryId) {
  return Array.from(document.querySelectorAll(`#${galleryId} .article-item:not(.filtered-out)`));
}

async function buildFullArticleClipboardHTML(button) {
  const articleItems = getVisibleArticleItems("articleGallery");

  if (!articleItems.length) {
    throw new Error("No visible power article sections found.");
  }

  const htmlBlocks = [];

  for (let i = 0; i < articleItems.length; i++) {
    if (button) button.innerText = `Copying ${i + 1}/${articleItems.length}...`;

    const item = articleItems[i];
    const card = item.querySelector(".power-card");
    const dataURL = await powerCardToDataURL(card);
    const textIndex = Number(item.dataset.copyIndex || i);
    const text = currentArticleTextBlocks[textIndex] || "";

    htmlBlocks.push(`
      <div style="margin:0 0 34px 0;page-break-inside:avoid;break-inside:avoid;">
        <img src="${dataURL}" alt="Power ranking graphic" style="display:block;width:900px;max-width:100%;height:auto;margin:0 0 14px 0;border:0;">
        <div style="font-family:Arial,sans-serif;color:#111827;">
          ${articlePlainTextToHTML(text)}
        </div>
      </div>
    `);
  }

  return `
    <!DOCTYPE html>
    <html>
      <head><meta charset="UTF-8"></head>
      <body>${htmlBlocks.join("")}</body>
    </html>
  `;
}

async function buildFullPredictionArticleClipboardHTML(button) {
  const articleItems = getVisibleArticleItems("predictionArticleGallery");

  if (!articleItems.length) {
    throw new Error("No visible prediction article sections found.");
  }

  const htmlBlocks = [];

  for (let i = 0; i < articleItems.length; i++) {
    if (button) button.innerText = `Copying ${i + 1}/${articleItems.length}...`;

    const item = articleItems[i];
    const card = item.querySelector(".match-card");
    const dataURL = await matchCardToDataURL(card);
    const textIndex = Number(item.dataset.copyIndex || i);
    const text = currentPredictionArticleTextBlocks[textIndex] || "";

    htmlBlocks.push(`
      <div style="margin:0 0 34px 0;page-break-inside:avoid;break-inside:avoid;">
        <img src="${dataURL}" alt="Prediction graphic" style="display:block;width:900px;max-width:100%;height:auto;margin:0 0 14px 0;border:0;">
        <div style="font-family:Arial,sans-serif;color:#111827;">
          ${articlePlainTextToHTML(text)}
        </div>
      </div>
    `);
  }

  return `
    <!DOCTYPE html>
    <html>
      <head><meta charset="UTF-8"></head>
      <body>${htmlBlocks.join("")}</body>
    </html>
  `;
}


/* =========================
   V2.14 RESPONSIVE UI OVERRIDES
========================= */
const UI_PREFS_KEY = "csc_media_hub_ui_prefs_v214";

function getUiPrefs() {
  try {
    return JSON.parse(localStorage.getItem(UI_PREFS_KEY) || "{}");
  } catch {
    return {};
  }
}

function saveUiPrefs(nextPrefs = {}) {
  const prefs = { ...getUiPrefs(), ...nextPrefs };
  localStorage.setItem(UI_PREFS_KEY, JSON.stringify(prefs));
}

function updateSidebarButtons() {
  const collapsed = document.body.classList.contains("sidebar-collapsed");
  const hideButton = document.getElementById("sidebarHideButton");
  const openButton = document.getElementById("sidebarOpenButton");

  if (hideButton) hideButton.innerText = collapsed ? "Show Controls" : "Hide Controls";
  if (openButton) openButton.innerText = collapsed ? "☰ Controls" : "Controls Open";
}

function toggleSidebar(forceCollapsed) {
  const nextCollapsed = typeof forceCollapsed === "boolean"
    ? forceCollapsed
    : !document.body.classList.contains("sidebar-collapsed");

  document.body.classList.toggle("sidebar-collapsed", nextCollapsed);
  updateSidebarButtons();
  saveUiPrefs({ sidebarCollapsed: nextCollapsed });

  window.setTimeout(resizeResponsiveGraphics, 80);
}

function updateDebugButton() {
  const enabled = document.body.classList.contains("debug-enabled");
  const button = document.getElementById("debugToggleButton");
  const details = document.querySelector(".debug-details");

  if (button) button.innerText = enabled ? "Disable Debug" : "Enable Debug";
  if (details && enabled) details.open = true;
}

function toggleDebugPanel(forceEnabled) {
  const nextEnabled = typeof forceEnabled === "boolean"
    ? forceEnabled
    : !document.body.classList.contains("debug-enabled");

  document.body.classList.toggle("debug-enabled", nextEnabled);
  updateDebugButton();
  saveUiPrefs({ debugEnabled: nextEnabled });
}

function loadUiPrefs() {
  const prefs = getUiPrefs();

  const shouldCollapseSidebar = typeof prefs.sidebarCollapsed === "boolean"
    ? prefs.sidebarCollapsed
    : false;

  document.body.classList.toggle("sidebar-collapsed", shouldCollapseSidebar);
  document.body.classList.toggle("debug-enabled", Boolean(prefs.debugEnabled));
  updateSidebarButtons();
  updateDebugButton();
}

function getAvailableGraphicWidth(card) {
  const item = card.closest(".article-item");
  const body = card.closest(".article-section-body");
  const base = item?.clientWidth || body?.clientWidth || window.innerWidth;
  return Math.max(260, base - 28);
}

function resizeResponsiveGraphics() {
  document.querySelectorAll(".article-section-body .power-card").forEach(card => {
    const available = getAvailableGraphicWidth(card);
    const scale = Math.min(.36, Math.max(.10, available / 3000));
    card.style.setProperty("--power-display-scale", scale.toFixed(4));
  });

  document.querySelectorAll(".article-section-body .match-card").forEach(card => {
    const available = getAvailableGraphicWidth(card);
    const scale = Math.min(.68, Math.max(.18, available / 1500));
    card.style.setProperty("--match-display-scale", scale.toFixed(4));
  });
}

function afterArticleRender() {
  setTimeout(() => {
    resizeTeamNames();
    resizeAnalysts();
    resizeResponsiveGraphics();
    populateWriterFilter();
    applyArticleFilters();
  }, 150);
}

window.addEventListener("resize", () => {
  window.clearTimeout(window.__cscResizeTimer);
  window.__cscResizeTimer = window.setTimeout(resizeResponsiveGraphics, 90);
});


/* =========================
   V2.15 PUBLISH / DRAFT FLOW
========================= */
const PUBLISH_CACHE_VERSION = "v215";

function tierDraftCacheKey(tierName) {
  return `csc_media_hub_draft_${PUBLISH_CACHE_VERSION}_${tierName}`;
}

function tierPublishedCacheKey(tierName) {
  return `csc_media_hub_published_${PUBLISH_CACHE_VERSION}_${tierName}`;
}

function normalizeTierPayload(payload) {
  if (!payload || !Array.isArray(payload.rankData) || !Array.isArray(payload.matchData)) {
    return null;
  }

  return {
    savedAt: payload.savedAt || payload.publishedAt || payload.draftAt || new Date().toISOString(),
    publishedAt: payload.publishedAt || "",
    draftAt: payload.draftAt || "",
    rankData: Array.isArray(payload.rankData) ? payload.rankData : [],
    matchData: Array.isArray(payload.matchData) ? payload.matchData : [],
    finalRows: Array.isArray(payload.finalRows) ? payload.finalRows : [],
    yapRows: Array.isArray(payload.yapRows) ? payload.yapRows : [],
    teamRecordData: Array.isArray(payload.teamRecordData) ? payload.teamRecordData : []
  };
}

function makeTierPayload(rankData, matchData, finalRows = [], yapRows = [], teamRecordData = [], extra = {}) {
  const now = new Date().toISOString();
  return normalizeTierPayload({
    savedAt: extra.savedAt || now,
    draftAt: extra.draftAt || "",
    publishedAt: extra.publishedAt || "",
    rankData,
    matchData,
    finalRows,
    yapRows,
    teamRecordData
  });
}

function saveDraftTierData(tierName, payload) {
  const normalized = normalizeTierPayload({ ...payload, draftAt: new Date().toISOString() });
  if (!normalized) return null;
  localStorage.setItem(tierDraftCacheKey(tierName), JSON.stringify(normalized));
  return normalized;
}

function getDraftTierData(tierName) {
  try {
    return normalizeTierPayload(JSON.parse(localStorage.getItem(tierDraftCacheKey(tierName)) || "null"));
  } catch {
    return null;
  }
}

function savePublishedTierData(tierName, payload) {
  const normalized = normalizeTierPayload({
    ...payload,
    publishedAt: new Date().toISOString(),
    savedAt: payload?.savedAt || new Date().toISOString()
  });
  if (!normalized) return null;
  localStorage.setItem(tierPublishedCacheKey(tierName), JSON.stringify(normalized));

  // Keep the old saved-data cache populated too, so older versions and fallbacks still have data.
  saveTierData(
    tierName,
    normalized.rankData,
    normalized.matchData,
    normalized.finalRows,
    normalized.yapRows,
    normalized.teamRecordData
  );

  return normalized;
}

function getPublishedTierData(tierName) {
  try {
    return normalizeTierPayload(JSON.parse(localStorage.getItem(tierPublishedCacheKey(tierName)) || "null"));
  } catch {
    return null;
  }
}

function isDraftNewerThanPublished(draft, published) {
  if (!draft) return false;
  if (!published) return true;

  const draftTime = Date.parse(draft.draftAt || draft.savedAt || "");
  const publishedTime = Date.parse(published.publishedAt || published.savedAt || "");

  return Number.isFinite(draftTime) && (!Number.isFinite(publishedTime) || draftTime > publishedTime + 1000);
}

function updatePublishStatus(message = "") {
  const status = document.getElementById("publishStatus");
  if (!status) return;

  if (message) {
    status.innerText = message;
    return;
  }

  const published = getPublishedTierData(currentTier);
  const draft = getDraftTierData(currentTier);
  const pieces = [];

  pieces.push(published
    ? `Live: ${formatSavedTime(published.publishedAt || published.savedAt)}`
    : "Live: not published yet");

  if (draft) {
    pieces.push(`Draft: ${formatSavedTime(draft.draftAt || draft.savedAt)}`);
  }

  if (isDraftNewerThanPublished(draft, published)) {
    pieces.push("Draft is waiting to publish");
  }

  status.innerText = pieces.join(" • ");
}

function setPublishButtonsBusy(isBusy, label = "") {
  const refreshButton = document.getElementById("refreshDraftButton");
  const publishButton = document.getElementById("publishDraftButton");

  if (refreshButton) {
    refreshButton.disabled = isBusy;
    refreshButton.innerText = isBusy && label ? label : "Refresh Draft";
  }

  if (publishButton) {
    publishButton.disabled = isBusy;
  }
}

async function fetchFreshTierPayload(tierName) {
  if (!Object.keys(TIER_CONFIG || {}).length) {
    await loadTierConfig();
  }

  const tier = TIER_CONFIG[tierName];

  if (!tier) {
    throw new Error(`Tier not found: ${tierName}`);
  }

  const rankingsURL = openSheetURL(tier.rankingsId, POWER_TAB);
  const picksURL = openSheetURL(tier.picksId, PICKS_TAB);

  const rankData = await fetchSheetData(rankingsURL, `${tierName} Power Rankings`);

  let matchData = [];
  let finalRows = [];
  let yapRows = [];
  let teamRecordData = [];

  try {
    matchData = await fetchSheetData(picksURL, `${tierName} Predictions`);
  } catch (err) {
    console.warn(`${tierName} predictions failed to load.`, err);
    matchData = [];
  }

  try {
    finalRows = await fetchRawSheetRows(tier.rankingsId, FINAL_TAB, `${tierName} Final Article Text`);
  } catch (err) {
    console.warn(`${tierName} Final tab failed to load.`, err);
    finalRows = [];
  }

  try {
    yapRows = await fetchRawSheetRows(tier.picksId, PREDICTION_FINAL_TAB, `${tierName} Prediction Final Text`);
  } catch (err) {
    console.warn(`${tierName} Prediction Final tab failed to load.`, err);
    yapRows = [];
  }

  try {
    teamRecordData = await fetchTeamRecordData(tierName);
  } catch (err) {
    console.warn(`${tierName} team records failed to load.`, err);
    teamRecordData = [];
  }

  return makeTierPayload(rankData, matchData, finalRows, yapRows, teamRecordData, {
    savedAt: new Date().toISOString()
  });
}

async function renderTierPayload(tierName, payload, modeLabel = "published") {
  const normalized = normalizeTierPayload(payload);
  if (!normalized) throw new Error(`No usable ${modeLabel} data for ${tierName}.`);

  currentTier = tierName;

  const select = document.getElementById("tierSelect");
  if (select) select.value = tierName;

  document.getElementById("powerGallery").innerHTML = "";
  document.getElementById("matchGallery").innerHTML = "";
  document.getElementById("articleGallery").innerHTML = "";
  document.getElementById("predictionArticleGallery").innerHTML = "";

  document.getElementById("status").innerText = `Loading ${tierName} ${modeLabel} data...`;

  await prepareTeamAssets(normalized.rankData, normalized.matchData);
  await loadPlayerProfileNames();

  renderTierData(
    tierName,
    normalized.rankData,
    normalized.matchData,
    normalized.finalRows,
    normalized.yapRows,
    normalized.teamRecordData,
    normalized.publishedAt || normalized.savedAt
  );

  const liveTime = formatSavedTime(normalized.publishedAt || normalized.savedAt);
  document.getElementById("status").innerText = `${tierName} showing ${modeLabel} data from ${liveTime}.`;
  updatePublishStatus();
}

async function loadPublishedTier(tierName) {
  try {
    currentTier = tierName;

    const select = document.getElementById("tierSelect");
    if (select) select.value = tierName;

    const published = getPublishedTierData(tierName);

    if (published) {
      await renderTierPayload(tierName, published, "published");
      return;
    }

    // First-run safety: if the user already had a local saved cache from older versions, treat it as live.
    const oldSaved = getSavedTierData(tierName);
    if (oldSaved) {
      const firstPublished = savePublishedTierData(tierName, oldSaved);
      await renderTierPayload(tierName, firstPublished, "published");
      updatePublishStatus(`No V2.15 publish snapshot existed, so your saved ${tierName} data was set as live.`);
      return;
    }

    document.getElementById("status").innerText = `No published ${tierName} data found. Pulling first live snapshot...`;
    const freshPayload = await fetchFreshTierPayload(tierName);
    const publishedFirstRun = savePublishedTierData(tierName, freshPayload);
    saveDraftTierData(tierName, freshPayload);
    await renderTierPayload(tierName, publishedFirstRun, "published");
    updatePublishStatus(`First ${tierName} snapshot published. Future sheet edits will stay hidden until you publish a draft.`);
  } catch (err) {
    console.error("PUBLISHED LOAD ERROR:", err);
    document.getElementById("status").innerText = `Failed to load ${tierName}: ${getSafeErrorMessage(err)}`;
    updatePublishStatus(`Failed to load ${tierName}: ${getSafeErrorMessage(err)}`);
  }
}

async function refreshDraftFromSheets(button, options = {}) {
  const oldText = button?.innerText || "";

  try {
    setPublishButtonsBusy(true, "Refreshing...");
    document.getElementById("status").innerText = `Refreshing ${currentTier} draft from sheets...`;
    updatePublishStatus(`Refreshing ${currentTier} draft from sheets...`);

    await loadTierConfig();
    setupTierDropdown();

    const payload = await fetchFreshTierPayload(currentTier);
    const draft = saveDraftTierData(currentTier, payload);

    // Keep latest raw data around for compatibility, but do not render it.
    saveTierData(
      currentTier,
      draft.rankData,
      draft.matchData,
      draft.finalRows,
      draft.yapRows,
      draft.teamRecordData
    );

    if (options.autoPublishIfMissing && !getPublishedTierData(currentTier)) {
      const published = savePublishedTierData(currentTier, draft);
      await renderTierPayload(currentTier, published, "published");
      updatePublishStatus(`${currentTier} had no live data, so the refreshed draft was published.`);
      return draft;
    }

    const draftReport = buildDebugReport(
      currentTier,
      draft.rankData,
      draft.matchData,
      draft.finalRows,
      draft.yapRows,
      draft.teamRecordData,
      draft.draftAt || draft.savedAt
    );
    draftReport.savedAt = draft.draftAt || draft.savedAt;
    updateDebugPanel(draftReport);

    document.getElementById("status").innerText = `${currentTier} draft refreshed. Current page is still showing the last published data.`;
    updatePublishStatus(`${currentTier} draft ready from ${formatSavedTime(draft.draftAt || draft.savedAt)}. Hit Publish Draft when you want it live.`);

    return draft;
  } catch (err) {
    console.error("DRAFT REFRESH ERROR:", err);
    document.getElementById("status").innerText = `Draft refresh failed: ${getSafeErrorMessage(err)}`;
    updatePublishStatus(`Draft refresh failed: ${getSafeErrorMessage(err)}`);
    return null;
  } finally {
    setPublishButtonsBusy(false);
    if (button && oldText) button.innerText = oldText;
  }
}

async function publishDraftData(button) {
  const oldText = button?.innerText || "";

  try {
    const draft = getDraftTierData(currentTier);

    if (!draft) {
      updatePublishStatus(`No ${currentTier} draft is loaded yet. Hit Refresh Draft first.`);
      document.getElementById("status").innerText = `No ${currentTier} draft is ready to publish.`;
      return;
    }

    if (button) button.innerText = "Publishing...";
    setPublishButtonsBusy(true, "Publishing...");

    const published = savePublishedTierData(currentTier, draft);
    await renderTierPayload(currentTier, published, "published");

    document.getElementById("status").innerText = `${currentTier} draft published.`;
    updatePublishStatus(`${currentTier} live data updated at ${formatSavedTime(published.publishedAt || published.savedAt)}.`);
  } catch (err) {
    console.error("PUBLISH ERROR:", err);
    document.getElementById("status").innerText = `Publish failed: ${getSafeErrorMessage(err)}`;
    updatePublishStatus(`Publish failed: ${getSafeErrorMessage(err)}`);
  } finally {
    setPublishButtonsBusy(false);
    if (button && oldText) button.innerText = oldText;
  }
}


// Keep the old function name working, but make it respect the new publish flow.
async function updateGraphicsFromSheets() {
  return refreshDraftFromSheets(document.getElementById("refreshDraftButton"));
}

function getConfiguredTiersToShow() {
  const orderedTiers = TIER_ORDER.filter(tier => TIER_CONFIG[tier]);
  const extraTiers = Object.keys(TIER_CONFIG).filter(tier => !orderedTiers.includes(tier));
  const tiersToShow = [...orderedTiers, ...extraTiers];

  if (!TIER_CONFIG[currentTier]) {
    currentTier = tiersToShow[0] || currentTier;
  }

  return tiersToShow;
}

function syncTierSelectValues(tierName) {
  ["tierSelect", "publishTierSelect"].forEach(id => {
    const select = document.getElementById(id);
    if (select && select.value !== tierName) select.value = tierName;
  });
}

function setupTierDropdown() {
  const tiersToShow = getConfiguredTiersToShow();
  const optionsHTML = tiersToShow
    .map(tier => `<option value="${escapeHTML(tier)}">${escapeHTML(tier)}</option>`)
    .join("");

  ["tierSelect", "publishTierSelect"].forEach(id => {
    const select = document.getElementById(id);
    if (!select) return;

    select.innerHTML = optionsHTML;
    select.value = currentTier;
    select.onchange = () => {
      currentTier = select.value;
      syncTierSelectValues(currentTier);
      loadPublishedTier(currentTier);
    };
  });
}

function getAdminUrlValue() {
  const params = new URLSearchParams(window.location.search || "");
  return params.get(typeof ADMIN_UNLOCK_QUERY_PARAM !== "undefined" ? ADMIN_UNLOCK_QUERY_PARAM : "admin");
}

function isAdminUnlocked() {
  if (typeof PUBLIC_SITE_DEFAULTS_TO_PUBLISH_MODE === "undefined" || !PUBLIC_SITE_DEFAULTS_TO_PUBLISH_MODE) {
    return true;
  }

  const urlValue = getAdminUrlValue();
  const requiredKey = typeof ADMIN_UNLOCK_KEY !== "undefined" ? String(ADMIN_UNLOCK_KEY || "") : "";

  if (urlValue !== null) {
    const ok = requiredKey ? urlValue === requiredKey : Boolean(urlValue);
    if (ok) {
      try {
        localStorage.setItem(ADMIN_UNLOCK_STORAGE_KEY || "csc_media_hub_admin_unlocked", "1");
      } catch {}
      return true;
    }
  }

  try {
    return localStorage.getItem(ADMIN_UNLOCK_STORAGE_KEY || "csc_media_hub_admin_unlocked") === "1";
  } catch {
    return false;
  }
}

function isViewerOnlyMode() {
  return Boolean(typeof PUBLIC_SITE_DEFAULTS_TO_PUBLISH_MODE !== "undefined" && PUBLIC_SITE_DEFAULTS_TO_PUBLISH_MODE && !isAdminUnlocked());
}

function applyAccessMode() {
  const viewerOnly = isViewerOnlyMode();
  document.body.classList.toggle("viewer-only", viewerOnly);

  if (viewerOnly) {
    document.body.classList.add("publish-mode");
    saveUiPrefs({ publishMode: true });
  }

  updatePublishModeButtons();
}

async function init() {
  try {
    document.getElementById("status").innerText = "Loading tier config...";

    await loadTierConfig();
    setupTierDropdown();

    const prefs = getUiPrefs();
    if (prefs.currentView) setView(prefs.currentView);

    await loadPublishedTier(currentTier);
    applyAccessMode();
  } catch (err) {
    console.error("CONFIG LOAD ERROR:", err);
    document.getElementById("status").innerText = `Failed to load tier config: ${getSafeErrorMessage(err)}`;
  }
}

function updatePublishModeButtons() {
  const enabled = document.body.classList.contains("publish-mode");
  const viewerOnly = document.body.classList.contains("viewer-only");
  const enterButton = document.getElementById("publishModeButton");
  const exitButton = document.getElementById("publishModeExitButton");

  if (enterButton) enterButton.innerText = enabled ? "Exit Publish Mode" : "Enter Publish Mode";
  if (exitButton) exitButton.innerText = viewerOnly ? "" : "Exit Publish Mode";
}

function togglePublishMode(forceEnabled) {
  if (isViewerOnlyMode()) {
    document.body.classList.add("publish-mode", "viewer-only");
    updatePublishModeButtons();
    return;
  }

  const nextEnabled = typeof forceEnabled === "boolean"
    ? forceEnabled
    : !document.body.classList.contains("publish-mode");

  document.body.classList.toggle("publish-mode", nextEnabled);
  saveUiPrefs({ publishMode: nextEnabled });
  updatePublishModeButtons();
  closeAllSectionMenus();

  if (nextEnabled) {
    expandVisibleSections();
  }

  window.setTimeout(resizeResponsiveGraphics, 100);
}

function loadPublishPrefs() {
  const prefs = getUiPrefs();
  const viewerOnly = isViewerOnlyMode();
  document.body.classList.toggle("viewer-only", viewerOnly);
  document.body.classList.toggle("publish-mode", viewerOnly ? true : Boolean(prefs.publishMode));

  if (prefs.currentView) currentView = prefs.currentView === "predictionArticle" ? "predictionArticle" : "article";

  updatePublishModeButtons();
  updatePublishStatus();
}


loadUiPrefs();
loadPublishPrefs();
init();
