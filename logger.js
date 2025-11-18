// @ts-nocheck
import fs from 'fs/promises';
import path from 'path';
import os from 'os';
import dotenv from 'dotenv';
dotenv.config();

class GSpreadLogger {
  constructor(config = {}) {
    this.levels = ['error', 'warn', 'info', 'debug'];
    this.config = {
      level: config.level || 'info',
      daily: config.daily !== false,
      maxFiles: config.maxFiles || '30d',
      filePrefix: config.filePrefix || 'gspread-appender',
      console: config.console !== false,
      file: config.file !== false,
      logDir: config.logDir || './logs',
      ...config
    };
    
    this.operations = new Map();
    this.ensureLogDir();
  }

  async ensureLogDir() {
    try {
      await fs.mkdir(this.config.logDir, { recursive: true });
    } catch (error) {
      console.error('Cannot create log directory:', error);
    }
  }

  // Format log entry
  formatLog(level, message, metadata = {}) {
    const timestamp = new Date().toISOString();
    const logEntry = {
      timestamp,
      level: level.toUpperCase(),
      message,
      ...metadata,
      pid: process.pid,
      hostname: os.hostname()
    };
    
    return JSON.stringify(logEntry);
  }

  // Get current log file name
  getLogFileName() {
    const date = new Date().toISOString().split('T')[0];
    return `${this.config.filePrefix}-${date}.log`;
  }

  // Write log to file (daily rotation)
  async writeToFile(logEntry) {
    if (!this.config.file) return;
    
    try {
      const filename = this.getLogFileName();
      const filePath = path.join(this.config.logDir, filename);
      await fs.appendFile(filePath, logEntry + '\n', 'utf8');
    } catch (error) {
      console.error('Log file write error:', error);
    }
  }

  // Console output with colors
  consoleOutput(level, message, metadata) {
    if (!this.config.console) return;
    
    const colors = {
      error: '\x1b[31m', // red
      warn: '\x1b[33m',  // yellow
      info: '\x1b[36m',  // cyan
      debug: '\x1b[90m', // gray
      reset: '\x1b[0m'
    };
    
    const timestamp = new Date().toLocaleString('id-ID');
    const levelColor = colors[level] || colors.info;
    
    const metaString = Object.keys(metadata).length > 0 
      ? ' ' + JSON.stringify(metadata)
      : '';
    
    console.log(
      `${levelColor}[${timestamp}] ${level.toUpperCase()}: ${message}${metaString}${colors.reset}`
    );
  }

  // Main log method
  log(level, message, metadata = {}) {
    if (!this.shouldLog(level)) return;
    
    const logEntry = this.formatLog(level, message, metadata);
    this.consoleOutput(level, message, metadata);
    this.writeToFile(logEntry);
  }

  // Check if should log based on level
  shouldLog(level) {
    const currentLevelIndex = this.levels.indexOf(this.config.level);
    const messageLevelIndex = this.levels.indexOf(level);
    return messageLevelIndex <= currentLevelIndex;
  }

  // ========== CONVENIENCE METHODS ==========
  error(message, metadata = {}) {
    this.log('error', message, metadata);
  }

  warn(message, metadata = {}) {
    this.log('warn', message, metadata);
  }

  info(message, metadata = {}) {
    this.log('info', message, metadata);
  }

  debug(message, metadata = {}) {
    this.log('debug', message, metadata);
  }

  // ========== SPECIALIZED METHODS FOR GSPREAD ==========
  spreadsheetOperation(operation, spreadsheetId, details = {}) {
    this.info(`Spreadsheet ${operation}`, {
      operation,
      spreadsheetId,
      ...details
    });
  }

  batchOperation(batchSize, successCount, errorCount, duration) {
    const successRate = batchSize > 0 ? ((successCount / batchSize) * 100).toFixed(1) : 0;
    this.info('Batch operation completed', {
      batchSize,
      successCount,
      errorCount,
      duration: `${duration}ms`,
      successRate: `${successRate}%`
    });
  }

  dataValidation(field, value, status) {
    this.debug(`Data validation: ${field}`, {
      field,
      value: String(value).substring(0, 100),
      status,
      validation: 'required_field_check'
    });
  }

  fileProcessing(filePath, action, details = {}) {
    this.info(`File ${action}`, {
      file: path.basename(filePath),
      action,
      ...details
    });
  }

  vmOperation(vmData, parentData, action) {
    this.info(`VM ${action}`, {
      vmHostname: vmData['Hostname'],
      vmSerial: vmData['Serial Number'],
      parentHostname: parentData?.['Hostname'],
      parentSerial: parentData?.['Serial Number'],
      action
    });
  }

  // ========== OPERATION TRACKING ==========
  startOperation(name, metadata = {}) {
    const operationId = `${name}-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`;
    const startTime = Date.now();
    
    this.operations.set(operationId, {
      name,
      startTime,
      metadata
    });
    
    this.info(`Operation started: ${name}`, {
      operationId,
      ...metadata
    });
    
    return operationId;
  }

  endOperation(operationId, result = {}) {
    const operation = this.operations.get(operationId);
    if (!operation) {
      this.warn('Unknown operation ended', { operationId });
      return;
    }
    
    const duration = Date.now() - operation.startTime;
    this.operations.delete(operationId);
    
    this.info(`Operation completed: ${operation.name}`, {
      operationId,
      duration: `${duration}ms`,
      ...operation.metadata,
      ...result
    });
    
    return duration;
  }

  errorOperation(operationId, error, context = {}) {
    const operation = this.operations.get(operationId);
    if (operation) {
      const duration = Date.now() - operation.startTime;
      this.operations.delete(operationId);
      
      this.error(`Operation failed: ${operation.name}`, {
        operationId,
        duration: `${duration}ms`,
        error: error.message,
        stack: error.stack,
        ...operation.metadata,
        ...context
      });
    } else {
      this.error('Operation error without start', {
        operationId,
        error: error.message
      });
    }
  }

  // ========== PERFORMANCE LOGGING ==========
  startTimer(operation) {
    const startTime = Date.now();
    return {
      end: (metadata = {}) => {
        const duration = Date.now() - startTime;
        this.debug(`Operation ${operation} completed`, {
          operation,
          duration: `${duration}ms`,
          ...metadata
        });
        return duration;
      }
    };
  }

  // ========== CLEANUP OLD LOG FILES ==========
  async cleanOldLogs() {
    if (!this.config.maxFiles) return;
    
    try {
      const files = await fs.readdir(this.config.logDir);
      const now = Date.now();
      let maxAge = 30 * 24 * 60 * 60 * 1000; // Default 30 days
      
      if (typeof this.config.maxFiles === 'string') {
        const match = this.config.maxFiles.match(/^(\d+)([dmy])$/);
        if (match) {
          const num = parseInt(match[1]);
          const unit = match[2];
          const multipliers = { d: 1, m: 30, y: 365 };
          maxAge = num * multipliers[unit] * 24 * 60 * 60 * 1000;
        }
      }
      
      let deletedCount = 0;
      
      for (const file of files) {
        if (file.startsWith(this.config.filePrefix) && file.endsWith('.log')) {
          const filePath = path.join(this.config.logDir, file);
          const stats = await fs.stat(filePath);
          
          if (now - stats.mtime.getTime() > maxAge) {
            await fs.unlink(filePath);
            deletedCount++;
            this.debug('Deleted old log file', { file });
          }
        }
      }
      
      if (deletedCount > 0) {
        this.info('Log cleanup completed', { deletedCount });
      }
      
    } catch (error) {
      this.error('Log cleanup failed', { error: error.message });
    }
  }

  // ========== GET OPERATION STATISTICS ==========
  getStats() {
    const stats = {
      activeOperations: this.operations.size,
      operations: Array.from(this.operations.entries()).map(([id, op]) => ({
        id,
        name: op.name,
        runningFor: `${Date.now() - op.startTime}ms`
      }))
    };
    
    this.debug('Operation statistics', stats);
    return stats;
  }
}

// ========== DEFAULT CONFIGURATION ==========
const DEFAULT_CONFIG = {
  level: process.env.LOG_LEVEL || 'info',
  daily: process.env.LOG_DAILY !== 'false',
  maxFiles: process.env.LOG_MAX_FILES || '30d',
  filePrefix: process.env.LOG_FILE_PREFIX || 'gspread-appender',
  console: process.env.LOG_CONSOLE !== 'false',
  file: process.env.LOG_FILE !== 'false',
  logDir: process.env.LOG_DIR || './logs'
};

// Create and export singleton instance
const logger = new GSpreadLogger(DEFAULT_CONFIG);

export default logger;