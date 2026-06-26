const POWER_TAB = "Master Data";
const PICKS_TAB = "Picks/Records";
const FINAL_TAB = "Final";
const PREDICTION_FINAL_TAB = "Final";

const LOGO_FOLDER = "logos";
const LOGO_EXTENSION = "png";

const logosURL = "https://opensheet.elk.sh/1vx2l-UI_eUlrz8b6f3vca24tUAfqO5w8_kmZsHxVSmw/logos";

const PROFILE_BASE_URL = "https://playcsc.com/stats/profile/";

// Sheet that has your player list.
// Column B = player names, so index 1 because JS arrays start at 0.
const PLAYER_LIST_SHEET_ID = "1Judoh6GL4Xev9Xx4biKSVNP7HVJSnXNepz1VjY9cy5k";

// Main tab to try first. Fallbacks below are tried automatically if this fails.
const PLAYER_LIST_TAB = "Players";
const PLAYER_LIST_FALLBACK_TABS = ["Players", "Sheet1", "Player List", "PlayerList", "Names"];

// Column B fallback if the header lookup cannot find a name column.
const PLAYER_LIST_COLUMN_INDEX = 1;

// These are checked first, so B1 can be Name, Player, Username, etc.
const PLAYER_LIST_NAME_HEADERS = ["Name", "Player", "Players", "Username", "Handle"];

let playerProfileNames = [];

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

let TIER_CONFIG = {};

let currentTier = "Premier";
let teamStyles = {};
let teamLogos = {};
let teamColorCache = {};
let currentArticleTextBlocks = [];
let currentPredictionArticleTextBlocks = [];

const defaultStyle = {
  primary: "#334155",
  secondary: "#0f172a",
  accent: "#64748b",
  accent2: "#475569",
  detail: "#94a3b8",
  dark: "#020617"
};
