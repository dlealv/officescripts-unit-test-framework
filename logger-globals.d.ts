interface LogEvent {
  readonly type: LOG_EVENT
  readonly message: string
  readonly timestamp: Date
}

interface Layout {
  format(event: LogEvent): string
  toString(): string
}

interface Appender {
  log(event: LogEvent): void
  log(msg: string, event: LOG_EVENT): void
  getLastLogEvent(): LogEvent | null
  toString(): string
}

interface Logger {
  error(msg: string): void
  warn(msg: string): void
  info(msg: string): void
  trace(msg: string): void
  getCriticalEvents(): LogEvent[]
  getErrCnt(): number
  getWarnCnt(): number
  getAction(): number
  getLevel(): number
  getAppenders(): Appender[]
  setAppenders(appenders: Appender[]): void
  addAppender(appender: Appender): void
  removeAppender(appender: Appender): void
  hasErrors(): boolean
  hasWarnings(): boolean
  hasMessages(): boolean
  clear(): void
  exportState(): {
    level: string
    action: string
    errorCount: number
    warningCount: number
    criticalEvents: LogEvent[]
  }
  toString(): string
}

// The LOG_EVENT enum (not a var/const!)
declare enum LOG_EVENT {
  ERROR = 1, // Always starts with 1
  WARN,
  INFO,
  TRACE,
  // Add other log levels as needed
}
export {}; // To allow global augmentation if you ever need it