// ====================================================================================
// GLOBAL CONSTANTS AND SETUP
// ====================================================================================

const SS = SpreadsheetApp.getActiveSpreadsheet();

// Sheet Names
const SH_CONFIG = "Config";
const SH_PEOPLE = "People";
const SH_SUGGESTIONS = "Suggestions";
const SH_VOTES = "Votes";
const SH_REQUESTS = "Requests";
const SH_ADMIN_AUDIT = "AdminAudit";
const SH_WINNERS = "Winners";

// Column Indices (0-indexed)
const COL_PEOPLE_NAME = 0;
const COL_PEOPLE_TOKEN = 1;
const COL_PEOPLE_RELATION = 2;
const COL_PEOPLE_ROLE = 3; // New: 'Admin', 'Parent', 'Voter'

const COL_SUGGESTIONS_NAME = 0;
const COL_SUGGESTIONS_GENDER = 1;
const COL_SUGGESTIONS_SUGGESTER = 2;
const COL_SUGGESTIONS_GUESS = 3;
const COL_SUGGESTIONS_RELATION = 4;
const COL_SUGGESTIONS_TIMESTAMP = 5;
const COL_SUGGESTIONS_MEANING = 6; // New Column

const COL_VOTES_NAME = 0;
const COL_VOTES_GENDER = 1;
const COL_VOTES_VOTER = 2;
const COL_VOTES_SCORE = 3;
const COL_VOTES_TIMESTAMP = 4;

// ====================================================================================
// UTILITY HELPERS
// ====================================================================================

/**
 * Normalizes a person's identifier to a consistent, case-insensitive value.
 * @param {string} name
 * @returns {string}
 */
function _normalizePersonName(name) {
  return String(name || '').trim().toLowerCase();
}

// ====================================================================================
// CORE SERVICE FUNCTIONS (doGet/doPost)
// ====================================================================================

/**
 * Wrapper function for google.script.run calls from the frontend.
 * This is the preferred method for communication.
 * @param {object} data The payload from the frontend.
 * @returns {object} The result object.
 */
function handleFrontendRequest(data) {
  try {
    const mode = data.mode;
    if (!mode) {
      return { success: false, error: "Missing 'mode' parameter in request." };
    }
    return _handleApiRequest(mode, data);
  } catch (error) {
    Logger.log("handleFrontendRequest Error: " + error.toString());
    return { success: false, error: "Server error in request: " + error.message };
  }
}

/**
 * Handles GET requests to the web app. Used for serving the HTML and API calls.
 * @param {GoogleAppsScript.Events.DoGet} e The event parameter.
 * @returns {GoogleAppsScript.Content.TextOutput | GoogleAppsScript.HTML.HtmlOutput}
 */
function doGet(e) {
  try {
    // Serve the main HTML page
    return HtmlService.createHtmlOutputFromFile('Frontend')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);

  } catch (error) {
    Logger.log("doGet Error: " + error.toString());
    return HtmlService.createHtmlOutput(`<h1>Error</h1><p>Server error in GET request: ${error.message}</p>`);
  }
}

/**
 * The doPost function is no longer used for API calls, as google.script.run is now used.
 * It is kept here to prevent unexpected behavior if a POST request is still made.
 * @param {GoogleAppsScript.Events.DoPost} e The event parameter.
 * @returns {GoogleAppsScript.Content.TextOutput}
 */
function doPost(e) {
  return _json({ success: false, error: "API calls should now use google.script.run." });
}

// ====================================================================================
// API ROUTING AND ACCESS CONTROL
// ====================================================================================

/**
 * Centralized router for all API requests (GET and POST).
 * @param {string} mode The requested API action.
 * @param {object} data The request parameters/body.
 * @returns {object} The result object.
 */
function _handleApiRequest(mode, data) {
  const { person, token } = data;
  const access = _getAccessLevel(person, token);

  // Define API endpoints and required roles
  const routes = {
    // Public Read/Write
    'status': { handler: _getStatus, role: 'Voter' },
    'suggest': { handler: _handleSuggest, role: 'Voter' },
    'vote': { handler: _handleVote, role: 'Voter' },
    'request': { handler: _handleRequest, role: 'Voter' },
    'names': { handler: _getNames, role: 'Voter' },
    'voterstate': { handler: _getVoterState, role: 'Voter' },
    'quota': { handler: _getQuota, role: 'Voter' },
    'my_suggestions': { handler: _getMySuggestions, role: 'Voter' },
    'update_suggestion': { handler: _updateSuggestion, role: 'Voter' },
    'delete_suggestion': { handler: _deleteSuggestion, role: 'Voter' },

    // Admin Only
    'admin': { handler: _handleAdmin, role: 'Admin' },
    'config': { handler: _getConfig, role: 'Admin' },
    'stats': { handler: _getStats, role: 'Admin' },
    'reveal': { handler: _getReveal, role: 'Voter' },
    'winners': { handler: _getWinners, role: 'Admin' },
    'export_data': { handler: _exportData, role: 'Admin' }, // New Admin Feature
  };

  const route = routes[mode];

  if (!route) {
    return { success: false, error: `Unknown API mode: ${mode}` };
  }

  // Role-Based Access Control (RBAC)
  if (route.role === 'Admin' && access.role !== 'Admin') {
    return { success: false, error: "Permission denied. Admin access required." };
  }
  if (route.role === 'Voter' && access.role === 'None') {
    return { success: false, error: "Authentication failed. Please check your name and token." };
  }

  // Execute handler
  return route.handler(data, access);
}

/**
 * Admin API: return the full configuration.
 */
function _getConfig(data, access) {
  // Only admins should reach this because of the routes.role check
  const config = _readConfig();
  return {
    success: true,
    config: config,
  };
}


/**
 * Admin API: basic funfair statistics.
 */
function _getStats(data, access) {
  const config = _readConfig();

  const suggestionsSheet = _getSheet(SH_SUGGESTIONS);
  const votesSheet       = _getSheet(SH_VOTES);
  const peopleSheet      = _getSheet(SH_PEOPLE);

  const suggestions = suggestionsSheet.getDataRange().getValues();
  const votes       = votesSheet.getDataRange().getValues();
  const people      = peopleSheet.getDataRange().getValues();

  let totalSuggestions = 0;
  let uniqueNames = 0;
  let girlNames = 0;
  let boyNames = 0;

  const distinctNames = new Set();

  if (suggestions.length > 1) {
    for (let i = 1; i < suggestions.length; i++) {
      const name   = suggestions[i][COL_SUGGESTIONS_NAME];
      const gender = suggestions[i][COL_SUGGESTIONS_GENDER];

      if (!name || !gender) continue;

      totalSuggestions++;

      const key = name.toLowerCase() + '|' + gender.toLowerCase();
      if (!distinctNames.has(key)) {
        distinctNames.add(key);
        const g = gender.toLowerCase();
        if (g === 'girl') girlNames++;
        if (g === 'boy')  boyNames++;
      }
    }
  }

  uniqueNames = distinctNames.size;

  const totalVotes  = votes.length  > 1 ? votes.length  - 1 : 0;
  const totalPeople = people.length > 1 ? people.length - 1 : 0;

  return {
    success: true,
    phase: config.PHASE,
    totals: {
      people: totalPeople,
      suggestions: totalSuggestions,
      uniqueNames: uniqueNames,
      girlNames: girlNames,
      boyNames: boyNames,
      votes: totalVotes,
    }
  };
}


/**
 * Admin API: return the list of winners (names only).
 */
function _getWinners(data, access) {
  const winnersSheet = _getSheet(SH_WINNERS);

  const winners = (winnersSheet.getLastRow() > 1)
    ? winnersSheet
        .getRange(2, 1, winnersSheet.getLastRow() - 1, 1)
        .getValues()
        .flat()
    : [];

  return {
    success: true,
    winners: winners,
  };
}


/**
 * Checks the user's name and token against the People sheet to determine their role.
 * @param {string} name The user's name.
 * @param {string} token The user's device token.
 * @returns {{role: string, relation: string}} The user's role ('Admin', 'Parent', 'Voter', 'None') and relation.
 */
/**
 * Checks the user's name against the People sheet.
 * Token is now ignored.
 */
function _getAccessLevel(name, token) {
  const nameLower = _normalizePersonName(name);
  if (!nameLower) return { role: 'None', relation: '' };
  const peopleSheet = _getSheet(SH_PEOPLE);
  const peopleData = peopleSheet.getDataRange().getValues();

  // 1. Check if name exists in the People sheet (for Admins or Parents)
  if (peopleData.length > 1) {
    for (let i = 1; i < peopleData.length; i++) {
      const row = peopleData[i];
      const rowName = _normalizePersonName(row[COL_PEOPLE_NAME]);

      // Match Name Only (Ignore Token)
      if (rowName === nameLower) {
        const role = row[COL_PEOPLE_ROLE] || 'Voter';
        const relation = row[COL_PEOPLE_RELATION] || 'Family';
        return { role: role, relation: relation };
      }
    }
  }

  // 2. Auto-Allow: If name is not in the sheet, let them in as a Guest Voter
  // This saves you from typing every single family member into the sheet.
  return { role: 'Voter', relation: 'Family' };
}

// ====================================================================================
// CONFIGURATION CACHING
// ====================================================================================

/**
 * Reads the configuration from the Config sheet, using the CacheService for performance.
 * @returns {object} The configuration object.
 */
function _readConfig() {
  const cache = CacheService.getScriptCache();
  const cachedConfig = cache.get('config');

  if (cachedConfig) {
    return JSON.parse(cachedConfig);
  }

  const configSheet = _getSheet(SH_CONFIG);
  const configData = configSheet.getDataRange().getValues();
  const config = {};

  // CRITICAL FIX 3: Ensure Config sheet has data beyond the header row
  if (configData.length <= 1) {
    throw new Error("Configuration sheet is empty. Please run setupSheets() or populate the Config sheet.");
  }

  for (let i = 0; i < configData.length; i++) {
    const key = configData[i][0];
    const value = configData[i][1];
    if (key) {
      // Attempt to parse numbers and booleans
      if (!isNaN(value) && value !== "") {
        config[key] = Number(value);
      } else if (value === "TRUE" || value === "FALSE") {
        config[key] = value === "TRUE";
      } else {
        config[key] = value;
      }
    }
  }

  // Parse star budgets (e.g., "5:3,4:5,3:7")
  const budgets = {};
  if (config.STAR_BUDGETS) {
    config.STAR_BUDGETS.split(',').forEach(item => {
      const parts = item.split(':');
      if (parts.length === 2) {
        budgets[Number(parts[0])] = Number(parts[1]);
      }
    });
  }
  config.budgets = budgets;

  // Cache the result for 5 minutes (300 seconds)
  cache.put('config', JSON.stringify(config), 300);

  return config;
}

/**
 * Updates the configuration and clears the cache.
 * @param {object} newConfig The configuration object to write.
 */
function _setConfig(newConfig) {
  const configSheet = _getSheet(SH_CONFIG);
  const configData = configSheet.getDataRange().getValues();
  const configMap = {};

  // Create a map of existing keys for quick lookup
  for (let i = 0; i < configData.length; i++) {
    configMap[configData[i][0]] = i;
  }

  // Update values
  for (const key in newConfig) {
    if (configMap.hasOwnProperty(key)) {
      const row = configMap[key];
      const value = newConfig[key];
      configSheet.getRange(row + 1, 2).setValue(value);
    }
  }

  // Clear the cache immediately
  CacheService.getScriptCache().remove('config');
}

// ====================================================================================
// API HANDLERS
// ====================================================================================

/**
 * Gets the current status of the funfair (phase, deadlines, counters).
 */
function _getStatus(data, access) {
  const config = _readConfig();
  const suggestionsSheet = _getSheet(SH_SUGGESTIONS);
  const suggestions = suggestionsSheet.getDataRange().getValues();

  let girlCount = 0;
  let boyCount = 0;
  const distinctNames = new Set();

  // CRITICAL FIX 4: Handle empty Suggestions sheet (only header row)
  if (suggestions.length > 1) {
    for (let i = 1; i < suggestions.length; i++) {
      const name = suggestions[i][COL_SUGGESTIONS_NAME];
      const gender = suggestions[i][COL_SUGGESTIONS_GENDER];

      // Robustness check for empty rows
      if (!name || !gender) continue;

      const nameLower = name.toLowerCase();
      const genderLower = gender.toLowerCase();

      if (!distinctNames.has(nameLower + "|" + genderLower)) {
        distinctNames.add(nameLower + "|" + genderLower);
        if (genderLower === 'girl') girlCount++;
        if (genderLower === 'boy') boyCount++;
      }
    }
  }

  return {
    success: true,
    phase: config.PHASE,
    deadlines: {
      nominations: config.DEADLINE_NOMINATIONS,
      voting: config.DEADLINE_VOTING,
      reveal: config.DEADLINE_REVEAL,
    },
    counters: {
      girl: girlCount,
      boy: boyCount,
      total: distinctNames.size,
    },
    config: {
      maxSuggestions: config.MAX_SUGGESTIONS_PER_PERSON || 10,
      budgets: config.budgets,
      parentWeight: config.PARENT_WEIGHT || 1,
      normalization: config.NORMALIZATION || false,
    },
    access: access,
  };
}

/**
 * Handles a new name suggestion.
 */
function _handleSuggest(data, access) {
  const config = _readConfig();
  if (config.PHASE !== 'Nominations') return { success: false, error: "Nominations phase is closed." };

  const { name, gender, guess, relation, meaning } = data; // Added meaning
  const suggester = _normalizePersonName(data.person);

  if (!name || !gender || !suggester || !guess || !relation) return { success: false, error: "Missing required fields." };

  const suggestionsSheet = _getSheet(SH_SUGGESTIONS);
  const suggestionsData = suggestionsSheet.getDataRange().getValues();

  // (Check for duplicates code - same as before)
  const nameLower = String(name).trim().toLowerCase();
  const genderLower = String(gender).trim().toLowerCase();
  const newKey = nameLower + "|" + genderLower;
  const fallbackMax = Number(config.MAX_SUGGESTIONS_PER_PERSON || 3);
  const maxGirls = Number(config.MAX_GIRL_SUGGESTIONS || fallbackMax);
  const maxBoys  = Number(config.MAX_BOY_SUGGESTIONS || fallbackMax);
  let userGirlCount = 0;
  let userBoyCount = 0;

  if (suggestionsData.length > 1) {
    for (let i = 1; i < suggestionsData.length; i++) {
      const rowSuggester = _normalizePersonName(suggestionsData[i][COL_SUGGESTIONS_SUGGESTER]);
      const eName = String(suggestionsData[i][COL_SUGGESTIONS_NAME] || '').toLowerCase();
      const eGender = String(suggestionsData[i][COL_SUGGESTIONS_GENDER] || '').toLowerCase();
      
      if (rowSuggester === suggester) {
         if (eName + "|" + eGender === newKey) return { success: false, error: "You already suggested this." };
         if (eGender === 'girl') userGirlCount++;
         if (eGender === 'boy') userBoyCount++;
      }
    }
  }

  if (genderLower === 'girl' && userGirlCount >= maxGirls) return { success: false, error: `Limit reached for girls.` };
  if (genderLower === 'boy' && userBoyCount >= maxBoys) return { success: false, error: `Limit reached for boys.` };

  // SAVE MEANING IN COLUMN 7
  suggestionsSheet.appendRow([
    name,
    gender,
    suggester,
    guess,
    relation,
    new Date(),
    meaning || '' // Save the meaning
  ]);

  SpreadsheetApp.flush();
  return { success: true, message: "Suggestion recorded." };
}


/**
 * Handles a vote submission.
 */
function _handleVote(data, access) {
  const config = _readConfig();
  if (config.PHASE !== 'Voting') {
    return { success: false, error: "Voting phase is closed." };
  }

  const { name, gender, score } = data;
  const voter = _normalizePersonName(data.person);
  const budgets = config.budgets;

  if (!name || !gender || !voter || !score) {
    return { success: false, error: "Missing required fields." };
  }

  const votesSheet = _getSheet(SH_VOTES);
  const votesData = votesSheet.getDataRange().getValues();

  // 1. Check current votes and budget
  const currentVotes = {};
  let existingRowIndex = -1;

  if (votesData.length > 1) {
    for (let i = 1; i < votesData.length; i++) {
      const row = votesData[i];
      const rowVoter = _normalizePersonName(row[COL_VOTES_VOTER]);
      const rowScore = row[COL_VOTES_SCORE];
      const rowName = row[COL_VOTES_NAME];
      const rowGender = row[COL_VOTES_GENDER];

      if (rowVoter && rowVoter === voter) {
        currentVotes[rowScore] = (currentVotes[rowScore] || 0) + 1;

        if (rowName && rowGender && rowName.toLowerCase() === name.toLowerCase() && rowGender.toLowerCase() === gender.toLowerCase()) {
          existingRowIndex = i + 1; // 1-indexed row number
          // If updating an existing vote, remove the old score from the count
          currentVotes[rowScore]--;
        }
      }
    }
  }

  // 2. Check if the new score exceeds the budget
  const newCount = (currentVotes[score] || 0) + 1;
  if (budgets[score] && newCount > budgets[score]) {
    return { success: false, error: `You have exceeded your budget for ${score}-star votes.`, budgetExceeded: true };
  }

  // 3. Record or Update the vote
  if (existingRowIndex !== -1) {
    // Update existing vote
    votesSheet.getRange(existingRowIndex, COL_VOTES_SCORE + 1).setValue(score);
    votesSheet.getRange(existingRowIndex, COL_VOTES_TIMESTAMP + 1).setValue(new Date());
  } else {
    // Append new vote
    votesSheet.appendRow([
      name,
      gender,
      voter,
      score,
      new Date(),
    ]);
  }

  return { success: true, message: "Vote recorded." };
}

/**
 * Gets the list of all distinct names for the voting phase.
 */
function _getNames(data, access) {
  const suggestionsSheet = _getSheet(SH_SUGGESTIONS);
  const suggestionsData = suggestionsSheet.getDataRange().getValues();
  const distinctNames = {};

  if (suggestionsData.length > 1) {
    for (let i = 1; i < suggestionsData.length; i++) {
      const row = suggestionsData[i];
      const name = row[COL_SUGGESTIONS_NAME];
      const gender = row[COL_SUGGESTIONS_GENDER];
      const meaning = row[COL_SUGGESTIONS_MEANING]; // Read meaning
      const suggesterName = _normalizePersonName(row[COL_SUGGESTIONS_SUGGESTER]);

      if (!name || !gender) continue;
      const key = name.toLowerCase() + "|" + gender.toLowerCase();

      if (!distinctNames[key]) {
        distinctNames[key] = {
          name: name,
          gender: gender,
          suggesters: new Set(),
          meanings: new Set() // Store unique meanings
        };
      }
      if (suggesterName) distinctNames[key].suggesters.add(suggesterName);
      if (meaning) distinctNames[key].meanings.add(meaning);
    }
  }

  const namesArray = Object.values(distinctNames).map(n => {
    n.suggesters = Array.from(n.suggesters);
    n.meanings = Array.from(n.meanings); // Convert to array
    return n;
  });

  return { success: true, names: namesArray };
}

/**
 * Gets the current voting state (all votes) for a specific voter.
 */
function _getVoterState(data, access) {
  const voter = _normalizePersonName(data.person);
  const votesSheet = _getSheet(SH_VOTES);
  const votesData = votesSheet.getDataRange().getValues();
  const voterVotes = {};

  if (votesData.length > 1) {
    for (let i = 1; i < votesData.length; i++) {
      const row = votesData[i];
      const rowVoter = _normalizePersonName(row[COL_VOTES_VOTER]);
      if (rowVoter && rowVoter === voter) {
        const name = row[COL_VOTES_NAME];
        const gender = row[COL_VOTES_GENDER];
        const score = row[COL_VOTES_SCORE];
        if (!name || !gender) continue;
        
        const key = name + "|" + gender;
        voterVotes[key] = score;
      }
    }
  }

  return { success: true, votes: voterVotes };
}

/**
 * Gets the current vote quota usage for a specific voter.
 */
function _getQuota(data, access) {
  const config = _readConfig();
  const voter = _normalizePersonName(data.person);
  const votesSheet = _getSheet(SH_VOTES);
  const votesData = votesSheet.getDataRange().getValues();
  const budgets = config.budgets;
  const usage = {};

  // Initialize usage based on budgets
  for (const star in budgets) {
    usage[star] = 0;
  }

  if (votesData.length > 1) {
    for (let i = 1; i < votesData.length; i++) {
      const row = votesData[i];
      const rowVoter = _normalizePersonName(row[COL_VOTES_VOTER]);
      if (rowVoter && rowVoter === voter) {
        const score = row[COL_VOTES_SCORE];
        if (usage.hasOwnProperty(score)) {
          usage[score]++;
        }
      }
    }
  }

  return { success: true, budgets: budgets, usage: usage };
}

/**
 * Handles a request for grace period or extra slots.
 */
function _handleRequest(data, access) {
  const { type, details } = data;
  const requester = _normalizePersonName(data.person);

  if (!type || !details || !requester) {
    return { success: false, error: "Missing required fields." };
  }

  const requestsSheet = _getSheet(SH_REQUESTS);
  requestsSheet.appendRow([
    requester,
    type,
    details,
    "Pending", // Status
    new Date(),
  ]);

  return { success: true, message: "Request submitted successfully. An admin will review it shortly." };
}

// ====================================================================================
// ADMIN HANDLERS
// ====================================================================================

/**
 * Main handler for all admin actions.
 */
function _handleAdmin(data, access) {
  const { action, payload } = data;

  switch (action) {
    case 'set_phase':
      return _adminSetPhase(payload);
    case 'set_config':
      return _adminSetConfig(payload);
    case 'approve_request':
      return _adminApproveRequest(payload);
    case 'draw_winners':
      return _adminDrawWinners(payload);
    default:
      return { success: false, error: `Unknown admin action: ${action}` };
  }
}

/**
 * Admin action: Set the current funfair phase.
 */
function _adminSetPhase(payload) {
  const { newPhase } = payload;
  if (!['Nominations', 'Voting', 'Reveal', 'Closed'].includes(newPhase)) {
    return { success: false, error: "Invalid phase name." };
  }
  _setConfig({ PHASE: newPhase });
  _logAdminAction(`Set phase to ${newPhase}`);
  return { success: true, message: `Phase successfully set to ${newPhase}.` };
}

/**
 * Admin action: Set multiple configuration values.
 */
function _adminSetConfig(payload) {
  const configKeys = ['DEADLINE_NOMINATIONS', 'DEADLINE_VOTING', 'DEADLINE_REVEAL', 'NORMALIZATION', 'PARENT_WEIGHT', 'STAR_BUDGETS', 'MAX_SUGGESTIONS_PER_PERSON', 'ACTUAL_GENDER'];
  const updates = {};

  for (const key of configKeys) {
    if (payload.hasOwnProperty(key)) {
      updates[key] = payload[key];
    }
  }

  if (Object.keys(updates).length === 0) {
    return { success: false, error: "No valid configuration keys provided for update." };
  }

  _setConfig(updates);
  _logAdminAction(`Updated config: ${JSON.stringify(updates)}`);
  return { success: true, message: "Configuration updated successfully." };
}

/**
 * Admin action: Approve a request (e.g., extra slots, grace period).
 * NOTE: Actual implementation of granting extra slots would require updating the People sheet or a new sheet.
 * For simplicity, this only updates the Requests sheet status.
 */
function _adminApproveRequest(payload) {
  const { requestId, status } = payload;
  const requestsSheet = _getSheet(SH_REQUESTS);
  const requestsData = requestsSheet.getDataRange().getValues();

  // Find the request row (assuming requestId is the 1-indexed row number)
  const rowIndex = Number(requestId);
  if (rowIndex < 2 || rowIndex > requestsData.length) {
    return { success: false, error: "Invalid request ID." };
  }

  requestsSheet.getRange(rowIndex, 4).setValue(status); // Column D is Status
  _logAdminAction(`Request ID ${requestId} set to status: ${status}`);

  return { success: true, message: `Request ID ${requestId} status updated to ${status}.` };
}

/**
 * Admin action: Draw winners from correct gender guessers.
 */
function _adminDrawWinners(payload) {
  const config = _readConfig();
  const actualGender = config.ACTUAL_GENDER;
  const { count } = payload;

  if (!actualGender || actualGender === 'Unknown') {
    return { success: false, error: "Actual gender must be set in config before drawing winners." };
  }

  const suggestionsSheet = _getSheet(SH_SUGGESTIONS);
  const suggestionsData = suggestionsSheet.getDataRange().getValues();
  const correctGuessers = new Set();

  if (suggestionsData.length > 1) {
    // Find all unique correct guessers
    for (let i = 1; i < suggestionsData.length; i++) {
      const row = suggestionsData[i];
      const guess = row[COL_SUGGESTIONS_GUESS];
      const suggester = row[COL_SUGGESTIONS_SUGGESTER];

      if (guess && guess.toLowerCase() === actualGender.toLowerCase()) {
        correctGuessers.add(suggester);
      }
    }
  }

  const guesserArray = Array.from(correctGuessers);
  if (guesserArray.length === 0) {
    return { success: false, message: "No correct guessers found." };
  }

  // Shuffle and select winners
  _shuffleArray(guesserArray);
  const winners = guesserArray.slice(0, count);

  // Log winners to the Winners sheet
  const winnersSheet = _getSheet(SH_WINNERS);
  winners.forEach(winner => {
    winnersSheet.appendRow([winner, actualGender, new Date()]);
  });

  _logAdminAction(`Drew ${winners.length} winners from correct guessers.`);
  return { success: true, winners: winners, message: `${winners.length} winners drawn and recorded.` };
}

/**
 * Admin action: Export all data sheets to a single JSON object. (New Feature)
 */
function _exportData(data, access) {
  const sheetsToExport = [SH_CONFIG, SH_PEOPLE, SH_SUGGESTIONS, SH_VOTES, SH_REQUESTS, SH_ADMIN_AUDIT, SH_WINNERS];
  const exportData = {};

  sheetsToExport.forEach(sheetName => {
    const sheet = _getSheet(sheetName);
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();

    if (values.length > 0) {
      const headers = values[0];
      const rows = values.slice(1).map(row => {
        const rowObject = {};
        headers.forEach((header, index) => {
          rowObject[header] = row[index];
        });
        return rowObject;
      });
      exportData[sheetName] = rows;
    } else {
      exportData[sheetName] = [];
    }
  });

  return { success: true, data: exportData, message: "Data exported successfully." };
}

// ====================================================================================
// REVEAL HANDLERS
// ====================================================================================

/**
 * Calculates and returns the final reveal data (rankings, charts data).
 */
function _getReveal(data, access) {
  const config = _readConfig();
  const suggestionsSheet = _getSheet(SH_SUGGESTIONS);
  const votesSheet = _getSheet(SH_VOTES);
  const peopleSheet = _getSheet(SH_PEOPLE);

  // Robustly read data, handling empty sheets
  const suggestions = (suggestionsSheet.getLastRow() > 1) ? suggestionsSheet.getRange(2, 1, suggestionsSheet.getLastRow() - 1, suggestionsSheet.getLastColumn()).getValues() : [];
  const votes = (votesSheet.getLastRow() > 1) ? votesSheet.getRange(2, 1, votesSheet.getLastRow() - 1, votesSheet.getLastColumn()).getValues() : [];
  const people = (peopleSheet.getLastRow() > 1) ? peopleSheet.getRange(2, 1, peopleSheet.getLastRow() - 1, peopleSheet.getLastColumn()).getValues() : [];

  // 1. Pre-process Data
  const allNames = _processSuggestions(suggestions);
  const voterWeights = _getVoterWeights(people, config.PARENT_WEIGHT);
  const voteData = _processVotes(votes, voterWeights);

  // 2. Calculate Scores
  const rankedNames = _calculateScores(allNames, voteData, config.NORMALIZATION);

  // 3. Generate Chart Data
  const chartData = _generateChartData(suggestions, rankedNames, config.ACTUAL_GENDER);

  // 4. Get Winners List
  const winnersSheet = _getSheet(SH_WINNERS);
  const winners = (winnersSheet.getLastRow() > 1) ? winnersSheet.getRange(2, 1, winnersSheet.getLastRow() - 1, 1).getValues().flat() : [];

  return {
    success: true,
    actualGender: config.ACTUAL_GENDER,
    rankedNames: rankedNames,
    chartData: chartData,
    winners: winners,
  };
}

/**
 * Processes suggestions to create a map of all distinct names and their suggesters/guesses.
 */
function _processSuggestions(suggestions) {
  const allNames = {};
  suggestions.forEach(row => {
    const name = row[COL_SUGGESTIONS_NAME];
    const gender = row[COL_SUGGESTIONS_GENDER];
    const suggester = row[COL_SUGGESTIONS_SUGGESTER];
    const guess = row[COL_SUGGESTIONS_GUESS];
    
    if (!name || !gender) return;

    const key = name.toLowerCase() + "|" + gender.toLowerCase();

    if (!allNames[key]) {
      allNames[key] = {
        name: name,
        gender: gender,
        suggesters: new Set(),
        guesses: new Set(),
        totalScore: 0,
        voteCount: 0,
        starCounts: { 5: 0, 4: 0, 3: 0, 2: 0, 1: 0 },
      };
    }
    allNames[key].suggesters.add(suggester);
    allNames[key].guesses.add(guess);
  });

  // Convert Sets to Arrays for final output
  Object.values(allNames).forEach(n => {
    n.suggesters = Array.from(n.suggesters);
    n.guesses = Array.from(n.guesses);
  });

  return allNames;
}

/**
 * Determines the weight for each voter based on the Parent role.
 */
function _getVoterWeights(people, parentWeight) {
  const weights = {};
  people.forEach(row => {
    const name = row[COL_PEOPLE_NAME];
    const role = row[COL_PEOPLE_ROLE] || 'Voter';
    if (name) {
      weights[name.toLowerCase()] = (role === 'Parent') ? (parentWeight || 1) : 1;
    }
  });
  return weights;
}

/**
 * Processes raw votes, applying voter weights.
 */
function _processVotes(votes, voterWeights) {
  const voteData = {};
  votes.forEach(row => {
    const name = row[COL_VOTES_NAME];
    const gender = row[COL_VOTES_GENDER];
    const voter = row[COL_VOTES_VOTER];
    const score = row[COL_VOTES_SCORE];
    
    if (!name || !gender || !voter) return;

    const key = name.toLowerCase() + "|" + gender.toLowerCase();
    const weight = voterWeights[voter.toLowerCase()] || 1;

    if (!voteData[key]) {
      voteData[key] = [];
    }
    voteData[key].push({ score: score, weight: weight });
  });
  return voteData;
}

/**
 * Calculates the final score for each name using a weighted average.
 * Implements tie-breaking logic.
 */
function _calculateScores(allNames, voteData, normalization) {
  const rankedNames = Object.values(allNames);

  rankedNames.forEach(nameObj => {
    const key = nameObj.name.toLowerCase() + "|" + nameObj.gender.toLowerCase();
    const votes = voteData[key] || [];

    let totalWeightedScore = 0;
    let totalWeight = 0;

    nameObj.starCounts = { 5: 0, 4: 0, 3: 0, 2: 0, 1: 0 };

    votes.forEach(vote => {
      totalWeightedScore += vote.score * vote.weight;
      totalWeight += vote.weight;
      if (vote.score >= 1 && vote.score <= 5) {
        nameObj.starCounts[vote.score]++;
      }
    });

    nameObj.totalWeight = totalWeight;
    nameObj.averageScore = totalWeight > 0 ? totalWeightedScore / totalWeight : 0;

    // Bayesian Smoothing (Simple implementation: add a prior)
    const priorScore = 3;
    const priorWeight = 5;

    nameObj.bayesianScore = (totalWeightedScore + priorScore * priorWeight) / (totalWeight + priorWeight);
  });

  // Sort: Highest Bayesian Score first.
  // Tie-Breaking:
  // 1. Primary: Bayesian Score (desc)
  // 2. Secondary: Total 5-star votes (desc)
  // 3. Tertiary: Total Weight (desc)
  rankedNames.sort((a, b) => {
    if (b.bayesianScore !== a.bayesianScore) {
      return b.bayesianScore - a.bayesianScore;
    }
    // Tie-breaker 1: 5-star votes
    if (b.starCounts[5] !== a.starCounts[5]) {
      return b.starCounts[5] - a.starCounts[5];
    }
    // Tie-breaker 2: Total Weight
    return b.totalWeight - a.totalWeight;
  });

  return rankedNames;
}

/**
 * Generates data structures for the charts (top names, gender guess pie, relation breakdown).
 */
function _generateChartData(suggestions, rankedNames, actualGender) {
  const chartData = {};

  // 1. Top Names Chart Data
  const topGirls = rankedNames.filter(n => n.gender.toLowerCase() === 'girl').slice(0, 10);
  const topBoys = rankedNames.filter(n => n.gender.toLowerCase() === 'boy').slice(0, 10);

  chartData.topNames = {
    girls: [['Name', 'Bayesian Score']].concat(topGirls.map(n => [n.name, n.bayesianScore])),
    boys: [['Name', 'Bayesian Score']].concat(topBoys.map(n => [n.name, n.bayesianScore])),
  };

  // 2. Gender Guess Pie Chart Data
  const guessCounts = { girl: 0, boy: 0, total: 0 };
  const correctGuessers = new Set();

  suggestions.forEach(row => {
    const guess = row[COL_SUGGESTIONS_GUESS];
    const suggester = row[COL_SUGGESTIONS_SUGGESTER];

    if (!guess) return;

    const guessLower = guess.toLowerCase();

    if (guessLower === 'girl') guessCounts.girl++;
    if (guessLower === 'boy') guessCounts.boy++;
    guessCounts.total++;

    if (actualGender && guessLower === actualGender.toLowerCase()) {
      correctGuessers.add(suggester);
    }
  });

  chartData.guessPie = [['Guess', 'Count'], ['Girl', guessCounts.girl], ['Boy', guessCounts.boy]];
  
  // --- CHANGE: Send the list of names, not just the count ---
  chartData.correctGuessersList = Array.from(correctGuessers); 
  // --------------------------------------------------------

  // 3. Relation Breakdown Chart Data
  const relationCounts = {};
  suggestions.forEach(row => {
    const relation = row[COL_SUGGESTIONS_RELATION];
    if (relation) {
      relationCounts[relation] = (relationCounts[relation] || 0) + 1;
    }
  });

  chartData.relationBreakdown = [['Relation', 'Suggestions']].concat(
    Object.entries(relationCounts).sort((a, b) => b[1] - a[1])
  );

  return chartData;
}

// ====================================================================================
// UTILITY FUNCTIONS
// ====================================================================================

/**
 * Helper to get a sheet by name.
 * @param {string} name The name of the sheet.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The sheet object.
 */
function _getSheet(name) {
  const sheet = SS.getSheetByName(name);
  if (!sheet) {
    throw new Error(`Sheet not found: ${name}`);
  }
  return sheet;
}

/**
 * Helper to return JSON content.
 * @param {object} obj The object to serialize.
 * @returns {GoogleAppsScript.Content.TextOutput}
 */
function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Logs an admin action to the AdminAudit sheet.
 * @param {string} action The action description.
 */
function _logAdminAction(action) {
  const auditSheet = _getSheet(SH_ADMIN_AUDIT);
  auditSheet.appendRow([new Date(), action, Session.getActiveUser().getEmail()]);
}

/**
 * Shuffles an array in place (Fisher-Yates algorithm).
 * @param {Array} array The array to shuffle.
 */
function _shuffleArray(array) {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    [array[i], array[j]] = [array[j], array[i]];
  }
}

// ====================================================================================
// INITIAL SETUP (Run once manually after deployment)
// ====================================================================================

/**
 * Creates the necessary sheets if they don't exist and sets up headers.
 */
function setupSheets() {
  const sheetNames = [
    SH_CONFIG, SH_PEOPLE, SH_SUGGESTIONS, SH_VOTES, SH_REQUESTS, SH_ADMIN_AUDIT, SH_WINNERS
  ];
  const headers = {
    [SH_CONFIG]: ["Key", "Value", "Description"],
    [SH_PEOPLE]: ["Name", "Device Token", "Relation", "Role"], // Added Role
    [SH_SUGGESTIONS]: ["Name", "Gender", "Suggester", "Gender Guess", "Relation", "Timestamp", "Meaning"],
    [SH_VOTES]: ["Name", "Gender", "Voter", "Score", "Timestamp"],
    [SH_REQUESTS]: ["Requester", "Type", "Details", "Status", "Timestamp"],
    [SH_ADMIN_AUDIT]: ["Timestamp", "Action", "Admin Email"],
    [SH_WINNERS]: ["Name", "Correct Guess", "Timestamp"],
  };

  sheetNames.forEach(name => {
    let sheet = SS.getSheetByName(name);
    if (!sheet) {
      sheet = SS.insertSheet(name);
      sheet.appendRow(headers[name]);
      sheet.setFrozenRows(1);
    }
  });

  // Initialize Config sheet with default values
  const configSheet = _getSheet(SH_CONFIG);
  if (configSheet.getLastRow() === 1) { // Only if empty (only header exists)
    configSheet.getRange("A2:C10").setValues([
      ["PHASE", "Nominations", "Current phase: Nominations, Voting, or Reveal"],
      ["DEADLINE_NOMINATIONS", "2025-12-01T23:59:59", "ISO Date/Time for Nominations end"],
      ["DEADLINE_VOTING", "2025-12-15T23:59:59", "ISO Date/Time for Voting end"],
      ["DEADLINE_REVEAL", "2025-12-20T10:00:00", "ISO Date/Time for Reveal"],
      ["MAX_SUGGESTIONS_PER_PERSON", 10, "Max names a person can suggest"],
      ["STAR_BUDGETS", "5:3,4:5,3:7", "Comma-separated list of Star:Count budgets"],
      ["PARENT_WEIGHT", 2, "Weight multiplier for 'Parent' role votes"],
      ["NORMALIZATION", "TRUE", "Apply score normalization (Bayesian smoothing)"],
      ["ACTUAL_GENDER", "Unknown", "Set to 'Girl' or 'Boy' before Reveal"],
    ]);
  }

  // Initialize People sheet with example data
  const peopleSheet = _getSheet(SH_PEOPLE);
  if (peopleSheet.getLastRow() === 1) {
    peopleSheet.getRange("A2:D4").setValues([
      ["Admin User", "ADMIN_TOKEN_123", "Parent", "Admin"],
      ["Parent 1", "PARENT_TOKEN_456", "Parent", "Parent"],
      ["Voter 1", "VOTER_TOKEN_789", "Khala", "Voter"],
    ]);
  }
}


// ====================================================================================
// PER-USER SUGGESTION MANAGEMENT
// ====================================================================================

/**
 * Returns the current user's own suggestions for display/edit/delete.
 */
function _getMySuggestions(data, access) {
  const personRaw = data.person || '';
  const person = _normalizePersonName(personRaw);
  const sheet = _getSheet(SH_SUGGESTIONS);
  const values = sheet.getDataRange().getValues();
  const items = [];

  if (!person) {
    return { success: false, error: "Missing person identifier." };
  }

  if (values.length > 1) {
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const suggester = _normalizePersonName(row[COL_SUGGESTIONS_SUGGESTER]);
      
      if (suggester === person) {
        // Safe Date Handling: Convert to string or use empty string if invalid
        let ts = row[COL_SUGGESTIONS_TIMESTAMP];
        let dateStr = (ts instanceof Date) ? ts.toISOString() : String(ts);

        items.push({
          id: i + 1, 
          name: row[COL_SUGGESTIONS_NAME],
          gender: row[COL_SUGGESTIONS_GENDER],
          guess: row[COL_SUGGESTIONS_GUESS],
          relation: row[COL_SUGGESTIONS_RELATION],
          timestamp: dateStr, // <--- FIX: Send string, not Date object
          meaning: row[COL_SUGGESTIONS_MEANING] || '' // Add this line
        });
      }
    }
  }

  return { success: true, items: items };
}

/**
 * Updates a single suggestion row belonging to the current user.
 */
function _updateSuggestion(data, access) {
  const sheet = _getSheet(SH_SUGGESTIONS);
  const rowId = Number(data.id);
  const personRaw = data.person || '';
  const person = _normalizePersonName(personRaw);
  const config = _readConfig();

  if (config.PHASE !== 'Nominations') {
    return { success: false, error: "Suggestions can only be edited during the nominations phase." };
  }

  if (!rowId || rowId < 2 || rowId > sheet.getLastRow()) {
    return { success: false, error: "Invalid suggestion id." };
  }

  // Read 7 columns now (including Meaning)
  const row = sheet.getRange(rowId, 1, 1, 7).getValues()[0];
  const suggester = _normalizePersonName(row[COL_SUGGESTIONS_SUGGESTER]);
  
  if (!person || suggester !== person) {
    return { success: false, error: "You can only edit your own suggestions." };
  }

  const newName = data.name || row[COL_SUGGESTIONS_NAME];
  const newGender = data.gender || row[COL_SUGGESTIONS_GENDER];
  const newGuess = data.guess || row[COL_SUGGESTIONS_GUESS];
  const newRelation = data.relation || row[COL_SUGGESTIONS_RELATION];
  const newMeaning = data.meaning !== undefined ? data.meaning : (row[COL_SUGGESTIONS_MEANING] || '');
  const now = new Date();

  if (!newName || !newGender) {
    return { success: false, error: "Name and gender are required." };
  }

  // Save 7 columns
  sheet.getRange(rowId, 1, 1, 7).setValues([[
    newName,
    newGender,
    row[COL_SUGGESTIONS_SUGGESTER],
    newGuess,
    newRelation,
    now,
    newMeaning
  ]]);

  return { success: true };
}

/**
 * Deletes a suggestion row belonging to the current user.
 */
function _deleteSuggestion(data, access) {
  const sheet = _getSheet(SH_SUGGESTIONS);
  const rowId = Number(data.id);
  const personRaw = data.person || '';
  const person = _normalizePersonName(personRaw);

  if (!rowId || rowId < 2 || rowId > sheet.getLastRow()) {
    return { success: false, error: "Invalid suggestion id." };
  }

  const row = sheet.getRange(rowId, 1, 1, 6).getValues()[0];
  const suggester = _normalizePersonName(row[COL_SUGGESTIONS_SUGGESTER]);
  if (!person || suggester !== person) {
    return { success: false, error: "You can only delete your own suggestions." };
  }

  sheet.deleteRow(rowId);
  return { success: true };
}

