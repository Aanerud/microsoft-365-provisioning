/**
 * Logging System
 *
 * Provides comprehensive logging to both console and file.
 * Tracks all operations, errors, and warnings during provisioning.
 */

import fs from 'fs/promises';
import path from 'path';

export type LogLevel = 'info' | 'warn' | 'error' | 'success' | 'debug';

export interface LogEntry {
  timestamp: string;
  level: LogLevel;
  message: string;
  context?: any;
}

export class Logger {
  private logFilePath: string;
  private logBuffer: LogEntry[] = [];
  private consoleEnabled: boolean = true;

  constructor(logDir: string = 'logs') {
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, -5);
    this.logFilePath = path.join(logDir, `provision-${timestamp}.log`);
  }

  /**
   * Initialize logger (create log directory and file)
   */
  async initialize(): Promise<void> {
    const logDir = path.dirname(this.logFilePath);
    await fs.mkdir(logDir, { recursive: true });

    // Write header
    const header = [
      '='.repeat(80),
      `M365 Agent Provisioning Log`,
      `Started: ${new Date().toISOString()}`,
      '='.repeat(80),
      '',
    ].join('\n');

    await fs.writeFile(this.logFilePath, header, 'utf-8');
    this.info(`Logging to: ${this.logFilePath}`);
  }

  /**
   * Log an info message
   */
  info(message: string, context?: any): void {
    this.log('info', message, context);
  }

  /**
   * Log a warning message
   */
  warn(message: string, context?: any): void {
    this.log('warn', message, context);
  }

  /**
   * Log an error message
   */
  error(message: string, context?: any): void {
    this.log('error', message, context);
  }

  /**
   * Log a success message
   */
  success(message: string, context?: any): void {
    this.log('success', message, context);
  }

  /**
   * Log a debug message
   */
  debug(message: string, context?: any): void {
    this.log('debug', message, context);
  }

  /**
   * Core logging method
   */
  private log(level: LogLevel, message: string, context?: any): void {
    const entry: LogEntry = {
      timestamp: new Date().toISOString(),
      level,
      message,
      context,
    };

    this.logBuffer.push(entry);

    // Console output with colors
    if (this.consoleEnabled) {
      const prefix = this.getPrefix(level);
      console.log(`${prefix} ${message}`);

      if (context) {
        console.log('  Context:', JSON.stringify(context, null, 2));
      }
    }

    // Write to file (async, non-blocking)
    this.writeToFile(entry).catch(err => {
      console.error('Failed to write to log file:', err);
    });
  }

  /**
   * Get console prefix for log level
   */
  private getPrefix(level: LogLevel): string {
    switch (level) {
      case 'info':
        return 'ℹ️ ';
      case 'warn':
        return '⚠️ ';
      case 'error':
        return '❌';
      case 'success':
        return '✅';
      case 'debug':
        return '🔍';
      default:
        return '';
    }
  }

  /**
   * Write log entry to file
   */
  private async writeToFile(entry: LogEntry): Promise<void> {
    const line = this.formatLogEntry(entry);
    await fs.appendFile(this.logFilePath, line + '\n', 'utf-8');
  }

  /**
   * Format log entry for file output
   */
  private formatLogEntry(entry: LogEntry): string {
    const level = entry.level.toUpperCase().padEnd(7);
    const timestamp = entry.timestamp;
    const message = entry.message;

    let line = `[${timestamp}] [${level}] ${message}`;

    if (entry.context) {
      line += '\n' + JSON.stringify(entry.context, null, 2);
    }

    return line;
  }

  /**
   * Enable/disable console output
   */
  setConsoleEnabled(enabled: boolean): void {
    this.consoleEnabled = enabled;
  }

  /**
   * Get log file path
   */
  getLogFilePath(): string {
    return this.logFilePath;
  }

  /**
   * Get all log entries
   */
  getLogBuffer(): LogEntry[] {
    return [...this.logBuffer];
  }

  /**
   * Write summary to log file
   */
  async writeSummary(summary: {
    created: number;
    updated: number;
    deleted: number;
    errors: number;
    warnings: number;
  }): Promise<void> {
    const summaryText = [
      '',
      '='.repeat(80),
      'Provisioning Summary',
      '='.repeat(80),
      `Created:  ${summary.created}`,
      `Updated:  ${summary.updated}`,
      `Deleted:  ${summary.deleted}`,
      `Errors:   ${summary.errors}`,
      `Warnings: ${summary.warnings}`,
      `Completed: ${new Date().toISOString()}`,
      '='.repeat(80),
      '',
    ].join('\n');

    await fs.appendFile(this.logFilePath, summaryText, 'utf-8');
  }

  /**
   * Close logger
   */
  async close(): Promise<void> {
    const errorCount = this.logBuffer.filter(e => e.level === 'error').length;
    const warnCount = this.logBuffer.filter(e => e.level === 'warn').length;

    await this.writeSummary({
      created: 0,
      updated: 0,
      deleted: 0,
      errors: errorCount,
      warnings: warnCount,
    });
  }
}

// Global logger instance
let globalLogger: Logger | null = null;

/**
 * Get global logger instance
 */
export function getLogger(): Logger {
  if (!globalLogger) {
    globalLogger = new Logger();
  }
  return globalLogger;
}

/**
 * Initialize global logger
 */
export async function initializeLogger(logDir?: string): Promise<Logger> {
  globalLogger = new Logger(logDir);
  await globalLogger.initialize();
  return globalLogger;
}
