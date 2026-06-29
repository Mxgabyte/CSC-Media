const POWER_TAB = "Master Data";
const PICKS_TAB = "Picks/Records";
const FINAL_TAB = "Final";
const PREDICTION_FINAL_TAB = "Final";

const LOGO_FOLDER = "Logos";
const LOGO_EXTENSION = "png";

const logosURL = "https://opensheet.elk.sh/1vx2l-UI_eUlrz8b6f3vca24tUAfqO5w8_kmZsHxVSmw/logos";

const PROFILE_BASE_URL = "https://playcsc.com/stats/profile/";

// Sheet that has your player list.
// Column B = player names, so index 1 because JS arrays start at 0.
const PLAYER_LIST_SHEET_ID = "1TF6C-wP2ZFErV7o7bMF1sH0DFjsBmzMlWdeFYKd9-zo";

// Main tab to try first. Fallbacks below are tried automatically if this fails.
const PLAYER_LIST_TAB = "Players";
const PLAYER_LIST_FALLBACK_TABS = [
  "Players",
  "External",
  "Sheet1",
  "Player List",
  "PlayerList",
  "Names",
  "Hyperlinks",
  "Hyperlink",
  "Player Hyperlinks",
  "Player Links",
  "Links",
  "Profiles",
  "CSC Profiles"
];

// Column B fallback if the header lookup cannot find a name column.
const PLAYER_LIST_COLUMN_INDEX = 1;

// These are checked first, so B1 can be Name, Player, Username, etc.
const PLAYER_LIST_NAME_HEADERS = [
  "Name",
  "Player",
  "Players",
  "Player Name",
  "Player Names",
  "Username",
  "Handle",
  "CSC Name",
  "Profile Name",
  "IGN",
  "Alias",
  "Aliases",
  "Display Name",
  "Full Name",
  "CSC Profile Name"
];

const PLAYER_LIST_URL_HEADERS = [
  "URL",
  "Link",
  "Links",
  "Profile",
  "Profile URL",
  "CSC Profile",
  "CSC Profile URL",
  "Hyperlink",
  "Player Link"
];

let playerProfileNames = [];
let playerProfileEntries = [];
let playerProfileUrlMap = {};

const TIER_ORDER = [
  "Recruit",
  "Prospect",
  "Contender",
  "Challenger",
  "Elite",
  "Premier"
];

const CONFIG_SHEET_URL = "https://docs.google.com/spreadsheets/d/1dCGPRYudr0X5ektYEA55WWwAxRax5zSdqzjoddAM0QE/edit?usp=sharing";
const CONFIG_TAB = "Config";

// Sheet with one worksheet per tier.
// Column A = Franchise, Column B = Team Name, Column C = Record.
const TEAM_RECORDS_SHEET_URL = "https://docs.google.com/spreadsheets/d/1nioy7RQUH5OIO4u_jBh8FTV3GtJzmUae2AAj1eAE1C4/edit?usp=sharing";
const TEAM_RECORDS_TAB_FALLBACKS = {
  Premier: ["Premier", "Master"],
  Master: ["Master", "Premier"]
};

// Optional map for full franchise names in the standings sheet.
// Left side = the exact/full franchise name after normalizing spaces/case.
// Right side = the abbreviation used by the Power Rankings and logo files.
// Add more here if Column A uses full names instead of abbreviations.
const FRANCHISE_CODE_ALIASES = {
  "NA NADES": "NAN",
  "ALL GOOD": "AG",
  "GONE FISHIN'": "GF",
  "GONE FISHIN": "GF",
  "GONE FISHING": "GF",
  "KINGSNAKES": "BOA",
  "AUTOMATA": "ATO",
  "THE BEACH": "BCH",
  "BEACH": "BCH",
  "THE 19TH HOLE": "TEE",
  "19TH HOLE": "TEE",
  "FINAL GIRL": "FNL",
  "THE TOAD-EM POLE": "FRG",
  "TOAD-EM POLE": "FRG",
  "THE TOADEM POLE": "FRG",
  "TOADEM POLE": "FRG",
  "WHAT DO YOU BEEF": "COW"
};

let TIER_CONFIG = {};

let currentTier = "Premier";
let teamStyles = {};
let teamLogos = {};
let teamColorCache = {};
let currentArticleTextBlocks = [];
let currentPredictionArticleTextBlocks = [];


// Public viewer/admin mode.
// Normal visitors are forced into Publish Mode.
// Open the site with ?admin=1 to unlock the full editor UI.
// For a light obscurity gate, set ADMIN_UNLOCK_KEY to a custom value and use ?admin=yourKey.
// This is a UI gate only, not real security, because GitHub Pages serves all code publicly.
const PUBLIC_SITE_DEFAULTS_TO_PUBLISH_MODE = true;
const ADMIN_UNLOCK_QUERY_PARAM = "admin";
const ADMIN_UNLOCK_KEY = "";
const ADMIN_UNLOCK_STORAGE_KEY = "csc_media_hub_admin_unlocked";

const defaultStyle = {
  primary: "#334155",
  secondary: "#0f172a",
  accent: "#64748b",
  accent2: "#475569",
  detail: "#94a3b8",
  dark: "#020617"
};
