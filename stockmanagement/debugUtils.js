// debugUtils.gs
// Version 1.0
// Last updated: 2025-03-02

const DebugUtils = {
  debugMode: false,
  debugLevel: 'INFO',
  
  setDebugMode(enabled) {
    this.debugMode = enabled;
    PropertiesService.getScriptProperties().setProperty('debugMode', enabled.toString());
    LoggerManager.info(`Debug mode ${enabled ? 'enabled' : 'disabled'}`);
  },

  setDebugLevel(level) {
    const validLevels = ['DEBUG', 'INFO', 'WARN', 'ERROR'];
    if (!validLevels.includes(level)) {
      throw new Error(`Invalid debug level. Valid levels are: ${validLevels.join(', ')}`);
    }
    
    this.debugLevel = level;
    PropertiesService.getScriptProperties().setProperty('debugLevel', level);
    LoggerManager.info(`Debug level set to ${level}`);
  },

  async inspectVariable(variable, options = {}) {
    const {
      depth = 3,
      showFunctions = false,
      maxArrayLength = 100
    } = options;

    try {
      const inspection = this.inspect(variable, depth, showFunctions, maxArrayLength);
      LoggerManager.debug('Variable inspection', inspection);
      return inspection;
    } catch (error) {
      LoggerManager.error('Error during variable inspection', error);
      throw error;
    }
  },

  inspect(value, depth, showFunctions, maxArrayLength) {
    if (depth <= 0) return '...';
    
    if (value === null) return 'null';
    if (value === undefined) return 'undefined';
    
    if (Array.isArray(value)) {
      const items = value.slice(0, maxArrayLength).map(item => 
        this.inspect(item, depth - 1, showFunctions, maxArrayLength)
      );
      return `Array(${value.length}) [${items.join(', ')}${value.length > maxArrayLength ? ', ...' : ''}]`;
    }
    
    if (typeof value === 'object') {
      const entries = Object.entries(value).map(([key, val]) => {
        if (typeof val === 'function' && !showFunctions) return null;
        return `${key}: ${this.inspect(val, depth - 1, showFunctions, maxArrayLength)}`;
      }).filter(Boolean);
      return `{${entries.join(', ')}}`;
    }
    
    if (typeof value === 'function' && !showFunctions) return 'function';
    
    return String(value);
  },

  async captureState() {
    const state = {
      timestamp: new Date().toISOString(),
      scriptProperties: {},
      spreadsheetState: {},
      executionContext: {
        user: Session.getActiveUser().getEmail(),
        timeZone: Session.getScriptTimeZone(),
        quotaRemaining: {
          email: MailApp.getRemainingDailyQuota(),
          triggers: ScriptApp.getProjectTriggers().length
        }
      }
    };

    try {
      // Script Properties
      const props = PropertiesService.getScriptProperties().getProperties();
      state.scriptProperties = props;

      // Spreadsheet State
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      state.spreadsheetState = {
        name: ss.getName(),
        url: ss.getUrl(),
        sheets: ss.getSheets().map(sheet => ({
          name: sheet.getName(),
          lastRow: sheet.getLastRow(),
          lastColumn: sheet.getLastColumn()
        }))
      };

      await this.saveState(state);
      return state;
    } catch (error) {
      LoggerManager.error('Error capturing state', error);
      throw error;
    }
  },

  async saveState(state) {
    const sheet = this.getDebugSheet();
    sheet.appendRow([
      state.timestamp,
      JSON.stringify(state.scriptProperties),
      JSON.stringify(state.spreadsheetState),
      JSON.stringify(state.executionContext)
    ]);
  },

  getDebugSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('DebugLog');
    
    if (!sheet) {
      sheet = ss.insertSheet('DebugLog');
      sheet.getRange('A1:D1').setValues([[
        'Timestamp',
        'Script Properties',
        'Spreadsheet State',
        'Execution Context'
      ]]);
      sheet.setFrozenRows(1);
    }
    
    return sheet;
  },

  async compareStates(state1, state2) {
    const differences = {
      scriptProperties: this.getDifferences(state1.scriptProperties, state2.scriptProperties),
      spreadsheetState: this.getDifferences(state1.spreadsheetState, state2.spreadsheetState),
      executionContext: this.getDifferences(state1.executionContext, state2.executionContext)
    };

    LoggerManager.info('State comparison results', differences);
    return differences;
  },

  getDifferences(obj1, obj2, path = '') {
    const differences = {};
    
    const keys = new Set([...Object.keys(obj1), ...Object.keys(obj2)]);
    
    for (const key of keys) {
      const currentPath = path ? `${path}.${key}` : key;
      
      if (!(key in obj1)) {
        differences[currentPath] = {
          type: 'added',
          value: obj2[key]
        };
        continue;
      }
      
      if (!(key in obj2)) {
        differences[currentPath] = {
          type: 'removed',
          value: obj1[key]
        };
        continue;
      }
      
      if (typeof obj1[key] === 'object' && typeof obj2[key] === 'object') {
        const nestedDiffs = this.getDifferences(obj1[key], obj2[key], currentPath);
        if (Object.keys(nestedDiffs).length > 0) {
          Object.assign(differences, nestedDiffs);
        }
      } else if (obj1[key] !== obj2[key]) {
        differences[currentPath] = {
          type: 'changed',
          oldValue: obj1[key],
          newValue: obj2[key]
        };
      }
    }
    
    return differences;
  }
};

// Global functions for usage
function setDebugMode(enabled) {
  return DebugUtils.setDebugMode(enabled);
}

function setDebugLevel(level) {
  return DebugUtils.setDebugLevel(level);
}

function captureDebugState() {
  return DebugUtils.captureState();
}

function compareDebugStates(state1, state2) {
  return DebugUtils.compareStates(state1, state2);
}